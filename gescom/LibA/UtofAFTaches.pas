{***********UNITE*************************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 14/12/2000
ModIfié le ... : 28/02/2001
Description .. : Source TOF de la TABLE : AFTACHES ()
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : TOF;AFTACHES
*****************************************************************}
Unit UtofAFTaches ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls
     ,HCtrls, HEnt1, HMsgBox, UTOF, DicoAF, UTob, UtilTaches
     ,SaisUtil,AGLInit,HTB97,HPanel,Vierge,Spin,AGLInitGC,LookUp
     ,windows, messages, AFUtilArticle
     ,CalcOleGenericAff
     ,AffaireUtil, UtilArticle, AFPLanningCst, paramsoc, UtilPlanning
     ,UTOFAFTRADUCCHAMPLIBRE, UtofAFPlanningViewer, UtofAFPlanningRes
     ,UTOFAFPLANCHARGE, AFPlanning, AFPlanningGene, menus
     ,UTofAFPLANNINGGENERER, UtilRessource
     ,UtofAfTacheModele,GereTobInterne,EntGC
     ,FactUtil,FactComm,Facture,FactTOB

{$IfDEF EAGLCLIENT}
    ,eFiche, MaineAGL
{$ELSE}
    ,Fiche,db,dbTables, FE_Main
{$EndIf}
      ;


Type

  TOF_AFTACHES = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure InitNewRessource(vTobRes : Tob);

    private

      fCurAffaire  : AffaireCourante;   // tache en cours
      fUpdateRes : TTabRes; // ensemble des ressources a reaffecter

      fStLesCols : String; // colonnes des taches

      fTobTaches : TOB;   // ensemble des tâches de l'affaire
      fTobDet    : TOB;   // tache courante de l'affaire
      fTobSup    : TOB;   // tâches supprimées
      fTobSupRes : TOB;   // ressources supprimées

      fTobCompet : TOB;   // Tob des competences

      fGSTaches : THGrid;       // grille de la liste des tâches
      fGSRes    : THGrid;       // grille des ressources
      fPaDet    : THPanel;      // panel contenant les onglets de détail de la tâche
      fAction   : TActionFiche; // action sur la fiche
      fStatut   : TActionFiche; // statut interne de la fiche (pour enregistrement)
      fArticle  : THEdit;
      fStOpenNumTache : String; // numero de la tache a l'ouverture

      fBoAFPLANDECHARGE : Boolean; // paramsoc indiquant si on est en plan de charge
      fBOAFGESTIONRAP   : Boolean; // paramsoc indiquant si on gere ou non le RAPlanifier


      fRB_MODEGENE1       : TRadioButton;
      fRB_MODEGENE2       : TRadioButton;
      fRB_QUOTIDIENNE     : TRadioButton;
      fRB_NBINTERVENTION  : TRadioButton;
      fRB_HEBDOMADAIRE    : TRadioButton;
      fRB_ANNUELLE        : TRadioButton;
      fRB_MENSUELLE       : TRadioButton;
      fPC_FREQUENCE       : TPageControl;
      fTS_QUOTIDIENNE     : TTAbSheet;
      fTS_NBINTERVENTION  : TTAbSheet;
      fTS_HEBDOMADAIRE    : TTAbSheet;
      fTS_ANNUELLE        : TTAbSheet;
      fTS_MENSUELLE       : TTAbSheet;

      fRB_MOISMETHODE1  : TRadioButton;
      fRB_MOISMETHODE2  : TRadioButton;
      fRB_MOISMETHODE3  : TRadioButton;
      fED_AFFAIRE       : THEdit;
      fBoClose          : Boolean; // fermeture de la fenetre

      fDtDebutAffaire   : TDateTime;
      fDtFinAffaire     : TDateTime;

      fStAfficheAffaire : String;

      fTobModele        : TOB;
      fStLignePiece     : String;

      procedure TraitementArgument(pStArgument : String);
      procedure AffichageEntete;
      procedure LoadAffaire(pInIndiceTob : Integer);
      procedure LoadCompetences;
      procedure GestionEcran;
      function ControleEcran : boolean;
      function Valider(pBoForce, pBoSupControl : Boolean) : Boolean;

      // Gestion des grids
      // Grid des tâches ...
      procedure InitGrid;
      procedure FormatColGridTaches;

      // ressources
      procedure InitGridRes;
      procedure AddUpdateRessList(aCol, aRow : Integer);
      procedure RefreshGridRes;
      procedure UpdateGridRes(pBoModifNbRes : Boolean; pInNum : Integer);

      // gestion de l'onglet règles
      procedure fRB_MODEGENE1OnClick(Sender: TObject);
      procedure fRB_MODEGENE2OnClick(Sender: TObject);
      procedure fRB_QUOTIDIENNEOnClick(Sender: TObject);
      procedure fRB_NBINTERVENTIONOnClick(Sender: TObject);
      procedure fRB_HEBDOMADAIREOnClick(Sender: TObject);
      procedure fRB_ANNUELLEOnClick(Sender: TObject);
      procedure fRB_MENSUELLEOnClick(Sender: TObject);
      procedure BPLANNINGOnClick(Sender: TObject);
      procedure BRESSOURCEOnClick(Sender: TObject);

      procedure BCALCULMNTRAPOnClick(Sender: TObject);
      function TestCoherence(pBoRAP : Boolean)  : Boolean;
      procedure BCALCULQTERAPOnClick(Sender: TObject);
      procedure BCALCULRAPRESOnClick(Sender: TObject);
      procedure PAGETACHEChange(Sender: TObject);

      // Gestion du grid
      procedure fGSTachesRowEnter(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure fGSTachesRowExit(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);

      procedure fGSResCellEnter(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
      procedure fGSResCellExit(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
      procedure fGSResKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure fGSResElipsisClick(SEnder: TObject);

      procedure FormKeyDown(SEnder: TObject; var Key: Word; ShIft: TShIftState);

      // Gestion des boutons
      procedure bInsertOnClick(SEnder: TObject);
      procedure bInsertResOnClick(SEnder: TObject);

      procedure bDeleteOnClick(SEnder: TObject);
      procedure bDeleteResOnClick(SEnder: TObject);

//      procedure bDuplicationOnClick(SEnder: TObject);
      procedure bFermerOnClick(SEnder: TObject);
      procedure bMemoOnClick(SEnder: TObject);
      procedure mnGeneAffaire_OnClick(SEnder: TObject);
      procedure mnGeneTache_OnClick(SEnder: TObject);
      procedure RefreshLastDateGene;

      procedure BPLANNINGVIEWEROnClick(Sender: TObject);
      procedure BPLANNINGRESOnClick(Sender: TObject);
      procedure BPLANDECHARGEOnClick(Sender: TObject);

      // Gestion de la tâche en cours
      //procedure ATA_PERIODICITEOnClick(SEnder: TObject);
      procedure fRB_MOISMETHODE1OnClick(SEnder: TObject);
      procedure fRB_MOISMETHODE2OnClick(SEnder: TObject);
      procedure fRB_MOISMETHODE3OnClick(SEnder: TObject);
      procedure EcranOnCloseQuery(Sender: TObject; var CanClose: Boolean);
      procedure ATA_MODESAISIEPDCOnChange(SEnder: TObject);
      procedure ATA_QTEINITIALEOnChange(SEnder: TObject);
      procedure ATA_QTEINITIALEOnExit(SEnder: TObject);
      procedure ATA_UNITETEMPSOnChange(SEnder: TObject);
      procedure ATA_PUPROnChange(Sender: TObject);
      procedure ATA_PUVENTEHTOnChange(Sender: TObject);
      procedure ATA_CODEARTICLEOnEnter(SEnder: TObject);
      procedure ATA_CODEARTICLEOnExit(SEnder: TObject);
      procedure ArticleOnChange(Sender: TObject);


      // tob... sur tâche en cours
      procedure InitNewTache (TobTache : Tob);
      function ControleZonesOblig : boolean;
      function FamilleArticleUnique : Boolean;
      function ArticleUnique : Boolean;
      function ControleRessource : Integer;
      function ScreenToTob : Boolean;
      procedure GetRegles;
      procedure PutRegles;
      function TobToScreen(Row : Integer) : Integer;
      procedure RefreshModeSaisie;
      procedure TobResToScreen;
      procedure ScreenToTobRes;
      procedure RefreshGrid;
      function IndiceTobTacheCourante : Integer;
      procedure InsertTache;
      procedure ValiderLesTaches;
      procedure InitNewRegles;

      procedure SaisiePlanCharge;
      procedure SaisieRAF;

      Function PlusTypeArticleTache : string;
      function ExistPlanning(pStAffaire, pStNumeroTache : String) : boolean;

      procedure CalculMontantRecurrent;
      procedure CalculQuantiteRecurrent;
      procedure CalculMontantPDC;
      procedure CalculQuantitePDC;
      procedure loadChampsCalc;

      // Tâche créée à partir de modèles tâches ou de ligne pièce affaire
      procedure TOBCopieModeleTache(FromTOB, ToTOB : TOB);
      procedure bInsertModeleOnClick(SEnder: TObject);
      procedure TraiteLignePiece;
      procedure RefreshLignePiece;
      procedure BTLIGNEOnClick(SEnder: TObject);

      Function RessourceSaisie : Boolean;
      function ClientFerme : Boolean;
  End;

  Function AFLanceFicheAFTaches(Argument : String) : String;

const 
	TexteMsgTache: array[1..32] of string 	= (
          {1}        'Saisie par affaire sur code affaire non valide.',
          {2}        'La mise à jour des tâches ne s''est pas effectuée correctement.',
          {3}        'Confirmez-vous la duplication de la tâche ?',
          {4}        'Une ressource ne peut être associée qu''une seule fois à une tâche.',
          {5}        'Le libellé de tâche est obligatoire.',
          {6}        'La famille de tâche est obligatoire.',
          {7}        'La grille des tâches n''est pas trouvée dans la fiche tâche.',
          {8}        'Erreur lors du changement de ressource. Opération annulée.',
          {9}        'Saisir le code de la ressource !',
          {10}       'La quantité initiale a été modifiée!' + #13#10 + 'Voulez vous recalculer la quantité Restant à Faire (si non, la quantité Restant à Faire précédente est conservée !)?',
          {11}       'Confirmez vous la suppression de la tâche ?',
          {12}       'Confirmez vous la suppression de la ressource et du planning associé' + #13#10 + '(si de l''activité à été générée pour ce planning, elle ne sera pas supprimée) ?',
          {13}       'La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.',
          {14}       'Aucune ressource n''est sélectionnée.',
          {15}       'La fonction est obligatoire.',
          {16}       'La date de début est supérieure à la date de fin.',
          {17}       'La ressource n''existe pas.',
          {18}       'Le type Article est obligatoire.',
          {19}       'L''article est obligatoire',
          {20}       'Un planning existe pour cette tâche. ' + #13#10 + 'Confirmez-vous la suppression de cette tâche et du planning correspondant ?',
          {21}       'En cours de développement !',
          {22}       'La durée d''une intervention doit être positive.',
          {23}       'Attention, la valeur initiale de la tâche n''est pas cohérente avec les valeurs initiales des ressources !' + #13#10 + ' Voulez vous quand même calculer les quantités ?',
          {24}       'Les quantités initiales des ressources ne sont pas cohérentes avec la quantité initiale de la tache.' + #13#10 + ' Le planning ne peut donc pas être généré!',
          {25}       'Aucune ressource n''est sélectionnée.',
          {26}       'Il y a un planning existant sur cette tâche. On ne peut pas modifier l''article.',
          {27}       'Attention, cette tâche a un lien avec une ligne de la pièce affaire.'+ #13#10 +'Confirmez-vous la suppression de cette tâche ?',
          {28}       'On ne peut pas avoir deux fois la même famille d''article dans une affaire.',
          {29}       'On ne peut pas avoir deux fois le même article dans une affaire.',
          {30}       'Veuillez affecter des ressources à la tâche pour pouvoir générer le planning.',
          {31}       'Attention, ce client est fermé. Vous ne devriez pas créer de tâches !',
          {32}       'Attention, ce client est fermé. Vous ne devriez pas générer de planning !'
                     );
 
Implementation

Function AFLanceFicheAFTaches(Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFTACHES','', '',Argument);
end;

//************************** Evènements de la TOF ******************************
procedure TOF_AFTACHES.OnNew;
Begin
  Inherited;
End;

procedure TOF_AFTACHES.OnDelete;
Begin
  Inherited;
End;

procedure TOF_AFTACHES.OnUpdate;
Begin
  inherited;
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then setFocusControl('ATA_FONCTION');
  if Valider(false, false) then RefreshGrid;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : validation des modifications
               : pBoForce : forcer la maj
               : pBoSupControl : forcer la maj lors de suppressions
Mots clefs ... :
*****************************************************************}
function TOF_AFTACHES.Valider(pBoForce, pBoSupControl : Boolean) : Boolean;
var
  io        : TIoErr;

Begin

  result := true;
  SetControlText('ATA_DATEFINPERIOD', GetControlText('ATA_DATEDEBPERIOD_'));

  // validation des données saisies dans l'écran si modification
  if pBoSupControl or ControleEcran then
    begin
      if pBoForce or ScreenToTob then
        begin
            io:=Transactions(ValiderLesTaches,1) ;
            if io<>oeOk then
              begin
                PGIBoxAF (TexteMsgTache[2],'');
                result := false;
              end
            else
              begin
                fStatut := taConsult;
                SetControlEnabled('ATA_MODESAISIEPDC', false);
              end;
        end;
    end
  else
    result := false;

  fBoClose := result;
end;

procedure TOF_AFTACHES.OnLoad;
Begin
  Inherited ;
End;

procedure TOF_AFTACHES.OnArgument (StArgument : String ) ;
Begin

  Inherited ;
  fBoClose := true;
  fUpdateRes := nil;
  fStatut := taConsult;
  fBoAFPLANDECHARGE := getparamsoc('SO_AFPLANDECHARGE');
  fBOAFGESTIONRAP   := getParamsoc('SO_AFREALPLAN');

  fCurAffaire.StAffaire := '';
  fCurAffaire.StTiers := '';

  TraitementArgument(stArgument);

  // chargement d'une affaire
  If fCurAffaire.StAffaire <> '' then
    Begin
      CodeAffaireDecoupe(fCurAffaire.StAffaire,fCurAffaire.StAff0,fCurAffaire.StAff1,fCurAffaire.StAff2,
                         fCurAffaire.StAff3,fCurAffaire.StAvenant, taModif, false);

      GestionEcran;
      LoadCompetences;
      LoadAffaire(0);
    end;


  // Cacher le lien sur le planning pour algoe
  setControlVisible('BPLANNINGVIEWER', fBoAFPLANDECHARGE);
  setControlVisible('BPLANNINGRES', fBoAFPLANDECHARGE);
  setControlVisible('BPLANNING', not fBoAFPLANDECHARGE);
  setControlVisible('BQTERESSOURCE', not fBoAFPLANDECHARGE);
  setControlVisible('BGENERER', not fBoAFPLANDECHARGE);
  setControlVisible('BCALCULMNTRAP', not fBoAFPLANDECHARGE);
  setControlVisible('BCALCULQTERAP', not fBoAFPLANDECHARGE);

  // si gestion du RAF et gestion en plan de charge
  setControlVisible('BPLANDECHARGE', fBoAFPLANDECHARGE);
  setControlVisible('BRAFQUANTITE', GetParamSoc('So_AFGESTIONRAF') and (not fBoAFPLANDECHARGE));
  setControlVisible('BRAFMONTANT', GetParamSoc('So_AFGESTIONRAF') and (not fBoAFPLANDECHARGE));
                                                         
  setControlEnabled('ATA_NBRESSOURCE', False);
  if fBoAFPLANDECHARGE then setControlEnabled('ATA_UNITETEMPS', False);
  setcontrolEnabled('RB_MODEGENE1', getParamSoc('SO_AFREALPLAN') and (StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0));
  setcontrolEnabled('RB_MODEGENE2', getParamSoc('SO_AFREALPLAN') and (StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0));

  fArticle := THedit(GetControl('ATA_CODEARTICLE'));
  fArticle.OnChange := ArticleOnChange;

  {$IFDEF CCS3}
  if (getcontrol('GBVALEUR') <> Nil) then SetControlVisible ('GBVALEUR', False);
  if (getcontrol('GBTEXTES') <> Nil)  then SetControlVisible ('GBTEXTES', False);
  if (getcontrol('GBDECISION') <> Nil)  then SetControlVisible ('GBDECISION', False);
  if (getcontrol('GBDATES__') <> Nil) then SetControlVisible ('GBDATES__', False);
  {$ENDIF}

    // C.B 03/11/2003 temporaire 5.00
  THPANEL(GetControl('PA_MONTANT1')).Top := 15;
  THPANEL(GetControl('PA_MONTANT1')).Left := 258;
  THPANEL(GetControl('PA_MONTANT1')).Width := 281;
  THEDIT(GetControl('ATA_INITPTPR')).Left := 166;
  THEDIT(GetControl('ATA_INITPTPR')).Width := 93;
  THPANEL(GetControl('PA_MONTANT2')).Top := 43;

End;

procedure TOF_AFTACHES.ArticleOnChange(Sender: TObject);
var s : STring;
begin
  S := RechDom ('AFTFAMARTICLE', GetControltext('ATA_CODEARTICLE'), false);
  if (s <> '') and (s <> 'Error') then
    SetControlText('FAMILLEARTICLE', RechDom ('GCFAMILLENIV1', s, false))
  else
    SetControlText('FAMILLEARTICLE','');
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... : 18/03/2002
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.OnClose;
Begin
  Inherited;
  Valider(false, false);
End;

//************************ Gestion de l'Entête *********************************

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.AffichageEntete;
var
  vSt         : String;
  vQr         : TQuery;
  vStAffFormat: String;
  vStaffiche  : String;

Begin
  // Par Affaire
  If fCurAffaire.StAffaire = '' then
    Begin
      PGIBoxAF (TexteMsgTache[1],'');
      close;
    End
  Else
    Begin
      SetControlText('ATA_AFFAIRE', fCurAffaire.StAffaire);
{      SetControlText('ATA_AFFAIRE0',fCurAffaire.StAff0);
      SetControlText('ATA_AFFAIRE1',fCurAffaire.StAff1);
      SetControlText('ATA_AFFAIRE2',fCurAffaire.StAff2);
      SetControlText('ATA_AFFAIRE3',fCurAffaire.StAff3);

      ChargeCleAffaire( THEdit(GetControl('ATA_AFFAIRE0')),
                        THEdit(GetControl('ATA_AFFAIRE1')),
                        THEdit(GetControl('ATA_AFFAIRE2')),
                        THEdit(GetControl('ATA_AFFAIRE3')),
                        THEdit(GetControl('ATA_AVENANT')),
                        Nil, taConsult, fCurAffaire.StAffaire, False);
}
      vSt := 'SELECT AFF_LIBELLE, T_TIERS, T_LIBELLE, AFF_DEVISE, ';
      vSt := vSt + 'AFF_DATEDEBUT, AFF_DATEFIN FROM TIERS, AFFAIRE ';
      vSt := vSt + 'WHERE AFF_AFFAIRE = "' + fCurAffaire.StAffaire + '"';
      vSt := vSt + 'AND AFF_TIERS = T_TIERS';

      vQr := nil;
      Try
        vQR := OpenSql(vSt, True);
        if Not vQR.Eof then
          begin
            //SetControlText('AFF_LIBELLE', vQR.FindField('AFF_LIBELLE').AsString);
            //SetControlText('T_LIBELLE', vQR.FindField('T_LIBELLE').AsString);
            fCurAffaire.StLibAff := vQR.FindField('AFF_LIBELLE').AsString;
            fCurAffaire.StLibTiers := vQR.FindField('T_LIBELLE').AsString;

            SetControlText('ATA_TIERS', vQR.FindField('T_TIERS').AsString);
            SetControlText('ATA_DEVISE', vQR.FindField('AFF_DEVISE').AsString);
            fCurAffaire.StTiers := vQR.FindField('T_TIERS').AsString;
            fCurAffaire.StDevise := vQR.FindField('AFF_DEVISE').AsString;
            fDtDebutAffaire := vQR.FindField('AFF_DATEDEBUT').AsDateTime;
            fDtFinAffaire := vQR.FindField('AFF_DATEFIN').AsDateTime;
          end;

         // affichage du tiers et de l'affaire
        vStAffFormat := CodeAffaireAffiche(fCurAffaire.StAffaire,' ');
        vStaffiche := format('%s   %s / %s   %s',[fCurAffaire.StTiers, fCurAffaire.StLibTiers, VStAffFormat, fCurAffaire.StLibAff]);
        SetcontrolText('TLIBCLIAFF', vStaffiche);
        // remis avait été supprimé. pourquoi ?
        fStAfficheAffaire := vStAffFormat + '  ' + fCurAffaire.StLibAff;
      finally
        if vQR <> nil then ferme(vQR);
      end;
    End;
End;

// ************************ Gestion des grids **********************************
// ** fGSTaches : Liste des tâches **
procedure TOF_AFTACHES.InitGrid;
Begin
  fGSTaches := THGRID(GetControl('GSTaches'));
  if assigned(fGSTaches) then
    begin 
      fGSTaches.OnRowEnter := fGSTachesRowEnter;
      fGSTaches.OnRowExit := fGSTachesRowExit;
      fGSTaches.ListeParam := 'AFLISTETACHES';
      fGSTaches.ColWidths[0] := 18;
      FormatColGridTaches;
      TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSTaches);
      AffecteGrid(fGSTaches, taConsult);
    end
  else
    PGIBoxAF (TexteMsgTache[7],'');
End;

procedure TOF_AFTACHES.FormatColGridTaches;
Var
  vStNomCol  : String;
  vStCols    : String;
  i          : Integer;
Begin
  fStLesCols := fGSTaches.Titres[0];
  vStCols :=fStLesCols;
  i := 0;
  Repeat
    vStNomCol := AnsiUppercase(Trim(ReadTokenSt(vStCols))) ;
    If (vStNomCol <> '') then
    Begin
      If (vStNomCol = 'ATA_DATEDEBPERIOD') or (vStNomCol = 'ATA_DATEFINPERIOD') then
      Begin
        fGSTaches.ColTypes[i]:='D';
        fGSTaches.ColFormats[i]:=ShortDateFormat;
      End
    End ;
    i := i + 1;
  Until ((vStCols='') or (vStNomCol=''));
End;

// *** fGSRes : Grid des ressources de la tâches ********************************
procedure TOF_AFTACHES.InitGridRes;
Begin

  fGSRes := THGRID(GetControl('GSRes'));

  fGSRes.OnCellEnter     := fGSResCellEnter;
  fGSRes.OnCellExit      := fGSResCellExit;
  fGSRes.OnKeyDown       := fGSResKeyDown;

  fGSRes.OnElipsisClick  := fGSResElipsisClick;
  fGSRes.OnDblClick      := fGSResElipsisClick;

  TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSRes);

  // Gestion des options de la grille
  AffecteGrid(fGSRes,fAction);

  fGSRes.colcount := 9;
  fGSRes.rowcount := 2;
  fGSRes.row      := 1;

  fGSRes.CellValues[0,0] := '';
  fGSRes.CellValues[cInColRes,0]    := 'Ressource';
  fGSRes.CellValues[cInColResLib,0] := 'Libellé';

  if (not fBoAFPLANDECHARGE) and (fBOAFGESTIONRAP) then
  begin
    fGSRes.CellValues[cInColQte,0]    := 'Qté Init';
    fGSRes.CellValues[cInColQteRAP,0] := 'Qté RAP';
  end;

  fGSRes.CellValues[cInColStatut,0] := 'Statut';
  fGSRes.CellValues[cInColComp1,0]  := 'Compétence 1';
  fGSRes.CellValues[cInColComp2,0]  := 'Compétence 2';
  fGSRes.CellValues[cInColComp3,0]  := 'Compétence 3';

  fGSRes.ColFormats[cInColQte] := '#,#00.00';
  fGSRes.ColTypes[cInColQte] := 'R';
  fGSRes.ColAligns[cInColQte] := taRightJustify;

  if fBoAFPLANDECHARGE or (not fBOAFGESTIONRAP) then
  begin
    fGSRes.ColWidths[0]             := 15;
    fGSRes.ColWidths[cInColRes]     := 80;
    fGSRes.ColWidths[cInColResLib]  := 131;
    fGSRes.ColWidths[cInColQte]     := -1;
    fGSRes.ColWidths[cInColQteRAP]  := -1;
    fGSRes.ColWidths[cInColStatut]  := 70;
    fGSRes.ColWidths[cInColComp1]   := 100;
    fGSRes.ColWidths[cInColComp2]   := 100;
    fGSRes.ColWidths[cInColComp3]   := 100;
  end
  else
  begin
    fGSRes.ColWidths[0]             := 15;
    fGSRes.ColWidths[cInColRes]     := 70;
    fGSRes.ColWidths[cInColResLib]  := 115;
    fGSRes.ColWidths[cInColQte]     := 48;
    fGSRes.ColWidths[cInColQteRAP]  := 48;
    fGSRes.ColWidths[cInColStatut]  := 60;
    fGSRes.ColWidths[cInColComp1]   := 80;
    fGSRes.ColWidths[cInColComp2]   := 80;
    fGSRes.ColWidths[cInColComp3]   := 80;
  end;

  // lignes
  fGSRes.RowHeights[0] := 22;
  fGSRes.ColEditables[cInColResLib]  := False;
  fGSRes.ColEditables[cInColStatut]  := False;
  fGSRes.ColEditables[cInColQteRAP]  := False;
End;                      
                               
//************************** Gestion du Grid Tâches *********************************//
procedure TOF_AFTACHES.fGSTachesRowEnter(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
Begin
  TobToScreen(Ou-1);
End;

// controle des données obligatoires saisies
procedure TOF_AFTACHES.fGSTachesRowExit(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
Begin
  fGSTaches.SynEnabled := false;
  cancel := not Valider(false, false);
  if not cancel then RefreshGrid;
  fGSTaches.SynEnabled := true;
End;

//*********************** Gestion du Grid Ressources / tâches ******************//
procedure TOF_AFTACHES.fGSResCellEnter(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
  // on affiche ... pour les colonnes ressources et fonctions
  If fAction <> taConsult then
    begin
      fGSRes.ElipsisButton := (fGSRes.Col = cInColRes) or
                              (fGSRes.Col = cInColComp1) or
                              (fGSRes.Col = cInColComp2) or
                              (fGSRes.Col = cInColComp3);
      // C.B
      //
      fTobDet.SetAllModifie(true);
    end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Reaffectation du planning d'une ressource
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.fGSResCellExit(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
  If fStatut <> taModif then Exit;

  //forcer l'ecriture en majuscule
  fGSRes.Cells[ACol, ARow] := AnsiUppercase(fGSRes.Cells[ACol, ARow]);

  // controle de l'existence de la ressource
  if (ACol = cInColRes) and not ExisteRessource (fGSRes.Cells[ACol, ARow]) then
    begin
      PGIBoxAF (TexteMsgTache[17],'');
      cancel := true;
    end
  else if (ACol = cInColRes) then
    fGSRes.Cells[cInColResLib, ARow] := RessourceLib (fGSRes.Cells[ACol, ARow]);

  //icone
  if fGSRes.Cells[cInColStatut, ARow] = '' then
    fGSRes.Cells[cInColStatut, ARow] := '#ICO#' + RechDom('AFSTATUTRESSOURCE', 'ACT', false);

  // initialisation de la quantite
  if fGSRes.Cells[cInColQte, ARow] = '' then
    fGSRes.Cells[cInColQte, ARow] := '0';

  // modification d'une ressource
  if (not cancel) and (aRow <> 0) and (ACol = cInColRes) and
    (fTobDet.Detail.Count >= aRow) and
    (fGSRes.Cells[ACol, ARow] <> fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE')) then
      begin
        if fGSRes.Cells[ACol, ARow] = '' then
        begin
          cancel := true;
          fGSRes.Cells[ACol, ARow] := fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE');
          fGSRes.Cells[cInColResLib, ARow] := RessourceLib (fGSRes.Cells[ACol, ARow]);
        end
        else if PGIAskAF(format('Confirmer-vous le remplacement de la ressource %s par la ressource %s .' + #13#10  + '(l''intégralité du planning sera transféré sur la nouvelle ressource même si de l''activité a été générée pour ce planning) ?', [ fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE'),fGSRes.Cells[ACol, ARow]]), '') = MrYes then
          AddUpdateRessList(aCol, aRow)
        else
        begin
          cancel := true;             
          fGSRes.Cells[ACol, ARow] := fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE');
          fGSRes.Cells[cInColResLib, ARow] := RessourceLib (fGSRes.Cells[ACol, ARow]);
        end;
      end;
  fGSRes.ElipsisButton := false;
End;
 
procedure TOF_AFTACHES.fGSResKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key <> VK_TAB) then fStatut := taModif;
  case key of
    VK_F5     : fGSResElipsisClick(SEnder);
    VK_DOWN   : if (fGSRes.row = (fGSRes.rowcount - 1)) then bInsertResOnClick(self);
    VK_RETURN : key := VK_TAB;
    VK_UP     : if (fGSRes.row = fGSRes.rowcount -1) and (fGSRes.Cells[cInColRes, fGSRes.row] = '') then
                begin
                  bDeleteResOnClick(self);
                  fGSRes.row := fGSRes.rowcount -2;
                  key := vk_down;
                end;
  end;                                                 
end;

procedure TOF_AFTACHES.fGSResElipsisClick(SEnder: TObject);
var
  vStRes : String;
Begin
  If fGSRes.Col = cInColRes then
    begin
      // fonction remplie
      if THValComboBox(GetControl('ATA_FONCTION')).Text <> '' then
        begin
          vStRes := DispatchRecherche(Nil,3,
                    'ARS_RESSOURCE='+fGSRes.Cells[fGSRes.Col,fGSRes.Row] +
                    ';ARS_FONCTION1=' + THValComboBox(GetControl('ATA_FONCTION')).value,'','');
          if vStRes <> '' then
            begin
              fGSRes.Cells[fGSRes.Col,fGSRes.Row] := vStRes;
              fStatut := taModif;
            end;
        end
      else
        begin
          vStRes := DispatchRecherche(Nil,3,
                    'ARS_RESSOURCE='+fGSRes.Cells[fGSRes.Col,fGSRes.Row],'','');
          if vStRes <> '' then
            begin
              fGSRes.Cells[fGSRes.Col,fGSRes.Row] := vStRes;
              fStatut := taModif;
            end;
        end;
    end
  else if (fGSRes.Col = cInColComp1) or
          (fGSRes.Col = cInColComp2) or
          (fGSRes.Col = cInColComp3) then
    LookUpList(fGSRes, TraduitGA('Compétences'), 'COMPETENCE','ACO_LIBELLE','','','ACO_LIBELLE',False,0);
End;

procedure TOF_AFTACHES.FormKeyDown(SEnder: TObject; var Key: Word; ShIft: TShIftState);
var
  vBoClose : Boolean;
Begin
  Case Key of
    VK_ESCAPE : begin vBoClose := true; EcranOnCloseQuery(self, vBoClose) end;
    VK_F10    : if Valider(false, false) then RefreshGrid;
  End;
End;

//************************** Gestion des boutons *******************************
procedure TOF_AFTACHES.bInsertOnClick(SEnder: TObject);
Begin
  if Valider(false, false) then
  begin
    RefreshGrid;
    InsertTache;
    SetActiveTabSheet('TAB_TACHE');
    SetFocusControl('ATA_LIBELLETACHE1');
  end;
End;
 
function TOF_AFTACHES.ClientFerme : Boolean;
var
  vSt : String;
  vQr : TQuery;
begin
  Result := False;
  vSt := 'SELECT T_FERME FROM TIERS, AFFAIRE ';
  vSt := vSt + ' WHERE AFF_TIERS = T_TIERS ';
  vSt := vSt + ' AND T_TIERS = "'+ fCurAffaire.StTiers + '"';
  vQr := nil;
  Try
    vQR := OpenSql(vSt, True);
    if Not vQR.Eof then
      Result := vQr.FindField('T_FERME').AsString = 'X';
  Finally
    Ferme(vQR);
  End;
end;
      
{*********************************************************************
Auteur  ...... : AB
Créé le ...... : 25/03/2002
Description .. : Création de tâches à partir de modèles de tâches
**********************************************************************}
procedure TOF_AFTACHES.bInsertModeleOnClick(SEnder: TObject);
var vTOBModeles : TOB;
    i :integer;
begin
  if (fStatut <> taCreat) and not Valider(false, false) then exit;
  RefreshGrid;
  vTOBModeles := MaTobInterne ('Liste ModeleTaches');
  AGLLanceFiche('AFF','AFTACHEMODELE_MUL','','','MULTI');
  for i:=0 to vTOBModeles.detail.count-1 do
  begin
    if (fStatut = taCreat) and (i = 0) then
    begin
      TOBCopieModeleTache (vTOBModeles.detail[i], fTobDet);
      fTobDet.PutEcran (Ecran,fPaDet);
      PutRegles;
    end else
    begin
      fTobModele := vTOBModeles.detail[i];
      InsertTache;
    end;
    if not Valider(false, false) then exit;
  end;
  if (VH_GC.AFTobInterne <> nil) then
    DetruitMaTobInterne ('Liste ModeleTaches');
  RefreshGrid;
  SetActiveTabSheet('TAB_TACHE');
  SetFocusControl('ATA_LIBELLETACHE1');
end;

procedure TOF_AFTACHES.bInsertResOnClick(SEnder: TObject);
Begin
  if (fGSRes.row = 1) or ((fGSRes.row <> 1) and (fGSRes.cells[cInColRes, fGSRes.row] <> '')) then
    UpdateGridRes(True, 1);
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : suppression des taches et des ressources
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.bDeleteOnClick(SEnder: TObject);
var
  vStMsg  : String;
Begin

  // suppression d'une tache
  if fGSTaches.Row <> 0 then
    begin
      fTobDet := TOB(fGSTaches.Objects[0,fGSTaches.Row]);
      // on verifie si des éléments planifiés existent
      if ExistPlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')) then
        vStMsg := TexteMsgTache[20]
      else if fTobDet.getvalue ('ATA_IDENTLIGNE') <> '' then
        vStMsg := TexteMsgTache[27] // AB-LigneAffaire-
      else
        vStMsg := TexteMsgTache[11];

      If (PGIAskAF(vStMsg,'')= mrYes) then
        Begin
          // entete compte
          if fGSTaches.RowCount > 2 then
            begin
              if assigned(fTobDet) then
                begin
                  fTobDet.ChangeParent(fTobSup,-1);
                  fGSTaches.DeleteRow(fGSTaches.Row);
                  if Valider(true, true) then
                    begin
                      RefreshGrid;
                      // repositionnement sur la precedente
                      TobToScreen(-1);
                    end
                end
              else
                begin
                fGSTaches.DeleteRow(fGSTaches.Row);
                RefreshGrid;
                TobToScreen(-1);
                end;
            end
          else
            begin
              fTobDet := fTobTaches.Detail[0];
              if assigned(fTobDet) then
                begin
                  fTobDet.ChangeParent(fTobSup,-1);
                  if Valider(true, true) then
                    begin
                      fGSTaches.DeleteRow(1);
                      InsertTache;
                    end;
                end;
            end;
        end;
    end;
End;

procedure TOF_AFTACHES.bDeleteResOnClick(SEnder: TObject);
var
  vTob        : Tob;
  vInRow      : integer;
  
begin
  // suppression d'une ressource
  if Screen.ActiveControl = fGSRes then
  begin

    if (fTobDet.detail.count = 0) and
       (GetControlText('ATA_NBRESSOURCE') > '1') then
    begin
      if fGSRes.RowCount > 2 then
      begin
        fGSRes.DeleteRow(fGSRes.Row);
        fGSRes.Row := 1;
        UpdateGridRes(True, -1);
        SetFocusControl('GSRES');
      end;
    end

    else if (fTobDet.detail.count = 0) and
       (GetControlText('ATA_NBRESSOURCE') = '1') then
      UpdateGridRes(False, 0)

    else if ((fGSRes.Cells[fGSRes.Col, fGSRes.Row] = '') and
             (fTobDet.Detail.count = (fGSRes.RowCount -2))) or
            (PGIAskAF(TexteMsgTache[12],'')= mrYes) then    
    Begin

      if (fTobDet.Detail.count >= (fGSRes.RowCount -1)) then
      begin
        //ajout a la liste de suppression
        SetLength(fUpdateRes, length(fUpdateRes) + 1);
        fUpdateRes[length(fUpdateRes) - 1].StOldRes := fTobDet.Detail[fGSRes.Row - 1].GetValue('ATR_RESSOURCE');
        fUpdateRes[length(fUpdateRes) - 1].StNewRes := fGSRes.Cells[fGSRes.Col, fGSRes.Row];
        fUpdateRes[length(fUpdateRes) - 1].rdAffaire := fCurAffaire;
        fUpdateRes[length(fUpdateRes) - 1].StNumTache := fTobDet.Detail[fGSRes.Row - 1].GetValue('ATR_NUMEROTACHE');
        fUpdateRes[length(fUpdateRes) - 1].InStatut := cInDelete;

        fTobDet := fTobTaches.Detail[fGSTaches.Row-1];
        vTob := fTobDet.Detail[fGSRes.Row-1];
        vTob.ChangeParent(fTobSupRes,-1);
      end;
 
      // pour gerer le pb de raffraichissement
      if fGSRes.RowCount > 2 then
      begin
        fGSRes.ElipsisButton := false;
        if fGSRes.Row = fGSRes.RowCount -1 then
          vInRow := fGSRes.Row -2
        else
          vInRow := fGSRes.Row -1;
        if vInRow = 0 then vInRow := 1;
         
        fGSRes.DeleteRow(fGSRes.Row);
        UpdateGridRes(True, -1);
        // mettre dans la tob les ressources sans sauvegarder
        ScreenToTob;
        SetFocusControl('GSRES');
        fGSRes.ElipsisButton := true;
        fGSRes.Row := vInRow;
      end

      else
      begin
        fGSRes.DeleteRow(fGSRes.Row);
        UpdateGridRes(False, 0);
        fGSRes.Row := 1;
        fGSRes.FixedRows := 1;
        // mettre dans la tob les ressources sans sauvegarder
        ScreenToTob;
        SetFocusControl('GSRES');
        fGSRes.ElipsisButton := false;
      end;
    End;
  end
  else
    PGIBoxAF (TexteMsgTache[14],'');
end;

{procedure TOF_AFTACHES.bDuplicationOnClick(SEnder: TObject);
Begin
 If (PGIAskAF(TexteMsgTache[3],'')<> mrYes) then Exit;
End;
}

procedure TOF_AFTACHES.bFermerOnClick(SEnder: TObject);
begin
  inherited;
  fBoClose := True;
end;

procedure TOF_AFTACHES.bMemoOnClick(SEnder: TObject);
var
  vStAction : String;
begin
  if fAction = taConsult then
    vStAction := 'ACTION=CONSULTATION'
  else
    vStAction := 'ACTION=MODIFICATION';

  AglLanceFiche('YY','YYLIENSOLE','ATA;'+ GetControlText('ATA_AFFAIRE') + '/' + GetControlText('ATA_NUMEROTACHE'),'', vStAction);
 
end;

//**************************Gestion de la Tâche en cours ***********************
// ********Evènemtents sur Bt,... de la tâche en cours *********

{procedure TOF_AFTACHES.ATA_PERIODICITEOnClick(SEnder: TObject);
Var
  stPeriodicite : string;

Begin
  stPeriodicite := GetControlText('ATA_PERIODICITE');
  SetControlEnabled('GB_INTERVAL',(stPeriodicite='J'));
  SetControlEnabled('GB_SEMAINE',(stPeriodicite='S'));
  SetControlEnabled('GB_ANNEE',(stPeriodicite='A'));
  SetControlEnabled('GB_MOIS',(stPeriodicite='M'));
  SetControlEnabled('GB_JOURINTERV',(stPeriodicite='S') or (stPeriodicite='M'));
End;
}

procedure TOF_AFTACHES.EcranOnCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  canClose := fBoClose;
  if canclose then
    begin
      if fTobTaches <> nil then begin fTobTaches.Free; fTobTaches := nil; end;
      if fTobSup <> nil then begin fTobSup.Free; fTobSup := nil; end;
      if fTobSupRes <> nil then begin fTobSupRes.Free; fTobSupRes := nil; end;
      if fTobCompet <> nil then begin fTobCompet.Free; fTobCompet := nil; end;
    end;                                                
end;

procedure TOF_AFTACHES.ATA_MODESAISIEPDCOnChange(SEnder: TObject);
begin
  RefreshModeSaisie;
end;

procedure TOF_AFTACHES.ATA_QTEINITIALEOnChange(SEnder: TObject);
begin
  if GetControlText('ATA_QTEINITIALE') <> '' then
    begin
      SetControlText('ATA_QTEINITUREF',
                     floatToStr(ConversionUnite(GetControlText('ATA_UNITETEMPS'),
                                                getparamsoc('SO_AFMESUREACTIVITE'),
                                                StrToFloat(GetControlText('ATA_QTEINITIALE')))));
      SetControlEnabled('RB_MODEGENE2', (GetControltext('ATA_QTEINITIALE')) <> '0');
      SetControlEnabled('RB_MODEGENE1', (GetControltext('ATA_QTEINITIALE')) <> '0');
    end
  else
    begin
      SetControlEnabled('RB_MODEGENE2', false);
      SetControlEnabled('RB_MODEGENE1', false);
    end;
end;

procedure TOF_AFTACHES.ATA_QTEINITIALEOnExit(SEnder: TObject);
begin
  if GetControlText('ATA_QTEINITIALE') = '' then
    SetControlText('ATA_QTEINITIALE', '0');
end;

procedure TOF_AFTACHES.ATA_PUPROnChange(SEnder: TObject);
var
  vDevise : RDevise;

begin
  // C.B 17/03/03 pas utilisé en prix de revient
  if THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA' then
  begin
    vDevise.code := fTobDet.GetValue('ATA_DEVISE');
    if (vDevise.code = '') or (vDevise.code = #0) then vDevise.Code:= V_PGI.DevisePivot;
    GetInfosDevise(vDevise);

    SetControlText('ATA_INITPTPR_', floatToStr(valeur(GetControlText('ATA_PUPR')) * valeur(GetControlText('ATA_QTEINITIALE'))));
//    SetControlText('ATA_INITPTPRDEV', floatToSTr(CalculMntDev(vDevise, valeur(GetControlText('ATA_PUPR')) * valeur(GetControlText('ATA_QTEINITIALE')))));
//    SetControlText('ATA_RAPPTPR', floatToStr(valeur(GetControlText('ATA_PUPR')) * valeur(GetControlText('ATA_QTEAPLANIFIER'))));
//    SetControlText('ATA_RAPPTPRDEV', floatToStr(CalculMntDev(vDevise, valeur(GetControlText('ATA_PUPR')) * valeur(GetControlText('ATA_QTEAPLANIFIER')))));
    SetControlText('ATA_RAFPTPRCALC', floatToStr(valeur(GetControlText('ATA_PUPR')) * valeur(GetControlText('ATA_QTEAPLANIFIERCALC'))));
//    SetControlText('ATA_RAFPTPRDEV', floatToStr(CalculMntDev(vDevise, valeur(GetControlText('ATA_PUPR')) * valeur(GetControlText('ATA_QTERAFAIRE')))));
  end;                   
end;

procedure TOF_AFTACHES.ATA_PUVENTEHTOnChange(SEnder: TObject);
var
  vDevise : RDevise;

begin
  vDevise.code := fTobDet.GetValue('ATA_DEVISE');
  if (vDevise.code = '') or (vDevise.code = #0) then vDevise.Code:= V_PGI.DevisePivot;
  GetInfosDevise(vDevise);

  SetControlText('ATA_INITPTVENTEHT', floatToStr(valeur(GetControlText('ATA_PUVENTEHT')) * valeur(GetControlText('ATA_QTEINITIALE'))));
  SetControlText('ATA_RAFPTVENTEHTCALC', floatToStr(valeur(GetControlText('ATA_PUVENTEHT')) * valeur(GetControlText('ATA_QTEAPLANIFIERCALC'))));
 
//  SetControlText('ATA_INITPTVTDEVHT', floatToStr(CalculMntDev(vDevise, valeur(GetControlText('ATA_PUVENTEHT')) * valeur(GetControlText('ATA_QTEINITIALE')))));
//  SetControlText('ATA_RAPPTVENTEHT', floatToStr(valeur(GetControlText('ATA_PUVENTEHT')) * valeur(GetControlText('ATA_QTEAPLANIFIERCALC'))));
//  SetControlText('ATA_RAPPTVTDEVHT', floatToStr(CalculMntDev(vDevise, valeur(GetControlText('ATA_PUVENTEHT')) * valeur(GetControlText('ATA_QTEAPLANIFIER')))));
//  SetControlText('ATA_RAFPTVENTEHT', floatToStr(valeur(GetControlText('ATA_PUVENTEHT')) * valeur(GetControlText('ATA_QTERAFAIRE'))));
//  SetControlText('ATA_RAFPTVTDEVHT', floatToStr(CalculMntDev(vDevise, valeur(GetControlText('ATA_PUVENTEHT')) * valeur(GetControlText('ATA_QTERAFAIRE')))));
end;

procedure TOF_AFTACHES.ATA_CODEARTICLEOnEnter(SEnder: TObject);

  Function ExistePlanning : Boolean;
  var
    vSt : String;
  Begin
    vSt := 'SELECT APL_NUMEROLIGNE FROM AFPLANNING WHERE APL_AFFAIRE = "';
    vSt := vSt + fCurAffaire.StAffaire;
    vSt := vSt + '" AND APL_NUMEROTACHE = ' + GetControlText('ATA_NUMEROTACHE');
    if ExisteSql(vSt) then
      result := true
    else
      result := false;
  End;

begin
  // on n'est pas en création
  if (GetControlText('ATA_NUMEROTACHE') <> '0') and ExistePLanning then
  begin
    PGIBoxAF (TexteMsgTache[26],'');
    SetFocusControl('ATA_LIBELLETACHE1');
  end;
end;

procedure TOF_AFTACHES.ATA_UNITETEMPSOnChange(SEnder: TObject);
begin 
  SetControlCaption('ATA_UNITETEMPS2', lowerCase(RechDom('GCQUALUNITTEMPS', getControlText('ATA_UNITETEMPS'), FALSE)));
end;

//*********** Gestion de la tâche TOB .... *********************
function TOF_AFTACHES.ControleZonesOblig : boolean;
begin

  result := true;
  if getControlText('ATA_LIBELLETACHE1') = '' then
    begin
      PGIBoxAF (TexteMsgTache[5],'');
      SetActiveTabSheet('TAB_TACHE');
      SetFocusControl('ATA_LIBELLETACHE1');
      result := false;
    end
  else if getControlText('ATA_FAMILLETACHE') = '' then
    begin
      PGIBoxAF (TexteMsgTache[6],'');
      SetActiveTabSheet('TAB_TACHE');
      SetFocusControl('ATA_FAMILLETACHE');
      result := false;
    end

  else if getControlText('ATA_TYPEARTICLE') = '' then
    begin
      PGIBoxAF (TexteMsgTache[18],'');
      SetActiveTabSheet('TAB_TACHE');
      SetFocusControl('ATA_TYPEARTICLE');
      result := false;
    end

  else if getControlText('ATA_CODEARTICLE') = '' then
    begin
      PGIBoxAF (TexteMsgTache[19],'');
      SetActiveTabSheet('TAB_TACHE');
      SetFocusControl('ATA_CODEARTICLE');
      result := false;
    end

  // la fonction n'est obligatoire que pour algoe
  else if fBoAFPLANDECHARGE and
          (getControlText('ATA_FONCTION') = '') and
          (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) and
          (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA') then
    begin
      PGIBoxAF (TexteMsgTache[15],'');
      SetActiveTabSheet('TAB_RESSOURCE');
      SetFocusControl('ATA_FONCTION');
      result := false;
    end
  else if strToDate(getControlText('ATA_DATEDEBPERIOD')) >  strToDate(getControlText('ATA_DATEFINPERIOD')) then
    begin
      PGIBoxAF (TexteMsgTache[16],'');
      SetActiveTabSheet('TAB_TACHE');
      SetFocusControl('ATA_DATEDEBPERIOD_');
      result := false;
    end
  else if StrToDate(GetControlText('ATA_DATEFINPERIOD')) > fDtFinAffaire then
    begin
      PGIBoxAF (TexteMsgTache[13],'');
      result := false;
    end
  else if (StrToFloat(GetControlText('ATA_QTEINTERVENT')) <= 0) and (not fBoAFPLANDECHARGE) then
    begin
      PGIBoxAF (TexteMsgTache[22],'');
      SetActiveTabSheet('TAB_REGLES');
      SetFocusControl('ATA_QTEINTERVENT');
      result := false;
    end
  else if (GetParamsoc('SO_AFLIENPLANACT') = 'FAM') and (not FamilleArticleUnique) then
    begin
      PGIBoxAF (TexteMsgTache[28],'');
      SetFocusControl('ATA_CODEARTICLE');
      result := false;
    end
  else if (GetParamsoc('SO_AFLIENPLANACT') = 'ART') and (not ArticleUnique) then
    begin
      PGIBoxAF (TexteMsgTache[29],'');
      SetFocusControl('ATA_CODEARTICLE');
      result := false;
    end
  else if result then
    loadChampsCalc;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Test si ont n'a pas cette famille d'article dans une
                 autre tache de la meme affaire
Mots clefs ... :
*****************************************************************}
function TOF_AFTACHES.FamilleArticleUnique : Boolean;
var
  i       : Integer;
  vStFam1 : String;
  vStFam2 : String;

begin
  result := True;
  vStFam2 := RechDom ('AFTFAMARTICLE', GetControlText('ATA_CODEARTICLE'), false);
  for i := 0 to fTobTaches.detail.count - 1 do
    if fTobTaches.detail[i] <> fTobDet then
    begin
      vStFam1 := RechDom ('AFTFAMARTICLE', fTobTaches.detail[i].Getvalue('ATA_CODEARTICLE'), false);
      if vStFam1 = vStFam2 then
      begin
        result := False;
        break;
       end;
    end;
end;

function TOF_AFTACHES.ArticleUnique : Boolean;
var
  i : integer;

begin
  result := True;
  for i := 0 to fTobTaches.detail.count - 1 do
    if fTobTaches.detail[i] <> fTobDet then
    begin
      if fTobTaches.detail[i].Getvalue('ATA_CODEARTICLE') = GetControlText('ATA_CODEARTICLE') then
      begin
        result := False;
        break;
       end;
    end;
end;
                  
procedure TOF_AFTACHES.loadChampsCalc;
begin
  // controle si la quantité prévue a été modifiée
  // on recalcul systematique si la quantité initiale etait à 0
  // seulemetn pour le plan de charge
  if fBoAFPLANDECHARGE and (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA') then
    CalculQuantitePDC

  else If fBoAFPLANDECHARGE and (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'MPR') then
    CalculMontantPDC

  // calcul à la demande bouton 
  else if ((not fBoAFPLANDECHARGE) and (fBOAFGESTIONRAP)) and (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA') then
    SetControlText('ATA_QTEAPLANIFIERCALC', '')

  else If ((not fBoAFPLANDECHARGE) and (fBOAFGESTIONRAP)) and (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'MPR') then
    SetControlText('ATA_RAPPTPRCALC', '');

  // calcul à la demande
{  else if ((not fBoAFPLANDECHARGE) and (fBOAFGESTIONRAP)) and (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA') then
    CalculQuantiteRecurrent

  else If ((not fBoAFPLANDECHARGE) and (fBOAFGESTIONRAP)) and (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'MPR') then
    CalculmontantRecurrent;
}
 // sinon, on ne recalcul pas
 // car on ne peut pas gerer le derapage

end;

procedure TOF_AFTACHES.CalculQuantitePDC;
var
  vRdQte      : Double;
  vRdRealise  : Double;
  vRdPlanifie : Double;
begin
  if ((valeur(GetControlText('ATA_QTEINITIALE')) <> 0) and
     (fTobDet.GetValue('ATA_QTEINITIALE') = 0)) or

     // question pour savoir si on ecrase le RAF avec la quantité initiale
     ((valeur(GetControlText('ATA_QTEINITIALE')) <>  valeur(fTobDet.GetValue('ATA_QTEINITIALE'))) and
     (valeur(GetControlText('ATA_QTEINITIALE')) <> 0) and
      fBoAFPLANDECHARGE and

      (PGIAskAF(TexteMsgTache[10],'')= mrYes)) then
        Begin
          vRdRealise := CalculRealise(fCurAffaire.StAffaire, getControlText('ATA_FONCTION'), 'ACT_QTE', '', fTobDet.GetValue('ATA_ARTICLE'));

          if getparamsoc('SO_AFRAFPLANNING') then
            vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE'), 'APL_QTEPLANIFIEE', '', false)
          else
            vRdPlanifie := 0;

          vRdQte := valeur(GetControlText('ATA_QTEINITIALE')) - vRdPlanifie
                    - vRdRealise + valeur(fTobDet.GetValue('ATA_QTEAPLANIFIER'));

          SetControlText('ATA_QTEAPLANIFIERCALC', FloatToStr(vRdQte));
          SetControlText('ATA_RAFPTPRCALC', FloatToStr(CalculMntPR(fTobDet)));
          //pas encore calculé
          //SetControlText('ATA_RAFPTVENTEHTCALC', FloatToStr(CalculMntPR(fTobDet)));
          SetControlText('ATA_RAFPTVENTEHTCALC', '0');
        end;
end;

procedure TOF_AFTACHES.CalculMontantPDC;
var
  vRdQte      : Double;
  vRdRealise  : Double;
  vRdPlanifie : Double;
begin
  if ((valeur(GetControlText('ATA_INITPTPR')) <> 0) and
  (fTobDet.GetValue('ATA_INITPTPR') = 0)) or

  // question pour savoir si on ecrase le RAF avec la quantité initiale
  ((valeur(GetControlText('ATA_INITPTPR')) <>  valeur(fTobDet.GetValue('ATA_INITPTPR'))) and
  (valeur(GetControlText('ATA_INITPTPR')) <> 0) and
  (PGIAskAF(TexteMsgTache[10],'')= mrYes)) then
     Begin
       vRdRealise := CalculRealise(fCurAffaire.StAffaire, '', 'ACT_TOTPR','', fTobDet.GetValue('ATA_ARTICLE'));
       if getparamsoc('SO_AFRAFPLANNING') then
         vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE'), 'APL_INITPTPR', '', false)
       else
         vRdPlanifie := 0;

       vRdQte := valeur(GetControlText('ATA_INITPTPR')) - vRdPlanifie
                 - vRdRealise + valeur(fTobDet.GetValue('ATA_RAPPTPR'));

       SetControlText('ATA_RAPPTPRCALC', FloatToStr(vRdQte));
       SetControlText('ATA_RAFPTPRCALC', FloatToStr(vRdQte));
       //pas encore calculé
      //SetControlText('ATA_RAFPTVENTEHTCALC', FloatToStr(CalculMntPR(fTobDet)));
      SetControlText('ATA_RAFPTVENTEHTCALC', '0');
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/12/2002
Modifié le ... :   /  /
Description .. : RAF avec prise en compte du planning n'est pas prévu pour l'instant
                 Forcé à 0 si négatif
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.CalculQuantiteRecurrent;
var
 vRdQte       : Double;
 vRdPlanifie  : Double;

begin
  vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE'), 'APL_QTEPLANIFIEE', '', true);

  vRdQte := valeur(GetControlText('ATA_QTEINITIALE')) - vRdPlanifie;
  if vRdQte < 0 then vRdQte := 0;
  SetControlText('ATA_QTEAPLANIFIERCALC', FloatToStr(vRdQte));

  if getparamsoc('SO_AFGESTIONRAF') then
    SetControlText('ATA_RAFPTPRCALC', FloatToStr(CalculMntPR(fTobDet)))
  else
    SetControlText('ATA_RAFPTPRCALC', '0');

  //pas encore calculé
  //SetControlText('ATA_RAFPTVENTEHTCALC', FloatToStr(CalculMntPR(fTobDet)));
  SetControlText('ATA_RAFPTVENTEHTCALC', '0')
end;

procedure TOF_AFTACHES.CalculMontantRecurrent;
var
 vRdQte     : Double;
 vRdPlanifie : Double;
begin
  vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE'), 'APL_INITPTPR', '', false);
  vRdQte := valeur(GetControlText('ATA_INITPTPR'))
            - vRdPlanifie + valeur(fTobDet.GetValue('ATA_RAPPTPR'));

  if vRdQte < 0 then vRdQte := 0;
  SetControlText('ATA_RAPPTPRCALC', FloatToStr(vRdQte));
  SetControlText('ATA_RAFPTPRCALC', FloatToStr(vRdQte));

  //pas encore calculé
  //SetControlText('ATA_RAFPTVENTEHTCALC', FloatToStr(CalculMntPR(fTobDet)));
  SetControlText('ATA_RAFPTVENTEHTCALC', '0')
end;


{***********A.G.L.***********************************************
Auteur  ...... : CB
xCréé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : controle de l'unicite de la ressource pour la prestation
                 et l'affaire. controle de la saisie de la ressource
Mots clefs ... :
*****************************************************************}
function TOF_AFTACHES.ControleRessource : Integer;
var
  i,j     : integer;
  vStRes  : String;

begin
  result := 0;
  fGSRes.BeginUpdate;

  // cas particulier
  // 0 ressource, on ne fait pas le controle
  if GetControlText('ATA_NBRESSOURCE') <> '0' then
    for i := 1 to fGSRes.RowCount - 1 do
      begin
        vStRes := fGSRes.CellValues[cInColRes,i];
        If vStRes = '' then
          result := -2
        else
          for j := i + 1  to fGSRes.RowCount - 1 do
            begin
              if vStRes = fGSRes.CellValues[cInColRes,j] then
                begin
                  result := -1;
                  break;
                end;
            end;
        if (result = -1) or (result = -2) then
          break;
      end;

  fGSRes.EndUpdate;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_AFTACHES.ScreenToTob : Boolean;
Begin
  If fTobDet <> nil then
    Begin

      SetControlText('ATA_DATEFINPERIOD', GetControlText('ATA_DATEDEBPERIOD_'));
      fTobDet.GetEcran (Ecran,fPaDet);
      GetRegles;

      // si des ressources saisies ou modifiées
      if fTobDet.modifie then ScreenToTobRes;
    end;

  // si des ressources saisies ou modifiées
  if fTobDet.modifie or (Length(fUpdateRes) <> 0) or (fStatut <> taConsult) then
    begin
      result := true;
      fStatut := taModif;
    end
  else
    result := false;
End;

procedure TOF_AFTACHES.GetRegles;
begin
  if fRB_MODEGENE1.Checked then
    fTobDet.PutValue('ATA_MODEGENE', 1)
  else
    fTobDet.PutValue('ATA_MODEGENE', 2);

  ftobDet.putvalue('ATA_QTEINTERVENT', GetControlText('ATA_QTEINTERVENT'));
  if fRB_QUOTIDIENNE.Checked then fTobDet.PutValue('ATA_PERIODICITE', 'Q')
  else if fRB_NBINTERVENTION.Checked then fTobDet.PutValue('ATA_PERIODICITE', 'NBI')
  else if fRB_HEBDOMADAIRE.Checked then
    begin
      fTobDet.PutValue('ATA_PERIODICITE', 'S');
      if TCheckBox(GetControl('ATA_JOUR1H')).Checked then fTobDet.PutValue('ATA_JOUR1', 'X') else fTobDet.PutValue('ATA_JOUR1', '-');
      if TCheckBox(GetControl('ATA_JOUR2H')).Checked then fTobDet.PutValue('ATA_JOUR2', 'X') else fTobDet.PutValue('ATA_JOUR2', '-');
      if TCheckBox(GetControl('ATA_JOUR3H')).Checked then fTobDet.PutValue('ATA_JOUR3', 'X') else fTobDet.PutValue('ATA_JOUR3', '-');
      if TCheckBox(GetControl('ATA_JOUR4H')).Checked then fTobDet.PutValue('ATA_JOUR4', 'X') else fTobDet.PutValue('ATA_JOUR4', '-');
      if TCheckBox(GetControl('ATA_JOUR5H')).Checked then fTobDet.PutValue('ATA_JOUR5', 'X') else fTobDet.PutValue('ATA_JOUR5', '-');
      if TCheckBox(GetControl('ATA_JOUR6H')).Checked then fTobDet.PutValue('ATA_JOUR6', 'X') else fTobDet.PutValue('ATA_JOUR6', '-');
      if TCheckBox(GetControl('ATA_JOUR7H')).Checked then fTobDet.PutValue('ATA_JOUR7', 'X') else fTobDet.PutValue('ATA_JOUR7', '-');
    end
  else if fRB_ANNUELLE.Checked then fTobDet.PutValue('ATA_PERIODICITE', 'A')
  else if fRB_MENSUELLE.Checked then
    begin
      fTobDet.PutValue('ATA_PERIODICITE', 'M');
      if fRB_MOISMETHODE1.Checked then
        begin
          fTobDet.PutValue('ATA_MOISMETHODE', '1');
          fTobDet.PutValue('ATA_MOISFIXE', GetControlText('ATA_MOISFIXE0'));
        end
      else if fRB_MOISMETHODE2.Checked then
        begin
          fTobDet.PutValue('ATA_MOISMETHODE', '2');
          fTobDet.PutValue('ATA_MOISSEMAINE', THValComboBox(GetControl('ATA_MOISSEMAINE1')).value);
          fTobDet.PutValue('ATA_MOISFIXE', GetControlText('ATA_MOISFIXE1'));
          fTobDet.PutValue('ATA_MOISJOURLIB', THValComboBox(GetControl('ATA_MOISJOURLIB')).Value); 
        end
      else if fRB_MOISMETHODE3.Checked then
        begin
          fTobDet.PutValue('ATA_MOISMETHODE', '3');
          fTobDet.PutValue('ATA_MOISSEMAINE', THValComboBox(GetControl('ATA_MOISSEMAINE2')).value);
          fTobDet.PutValue('ATA_MOISFIXE', GetControlText('ATA_MOISFIXE2'));

          if TCheckBox(GetControl('ATA_JOUR1M')).Checked then fTobDet.PutValue('ATA_JOUR1', 'X') else fTobDet.PutValue('ATA_JOUR1', '-');
          if TCheckBox(GetControl('ATA_JOUR2M')).Checked then fTobDet.PutValue('ATA_JOUR2', 'X') else fTobDet.PutValue('ATA_JOUR2', '-');
          if TCheckBox(GetControl('ATA_JOUR3M')).Checked then fTobDet.PutValue('ATA_JOUR3', 'X') else fTobDet.PutValue('ATA_JOUR3', '-');
          if TCheckBox(GetControl('ATA_JOUR4M')).Checked then fTobDet.PutValue('ATA_JOUR4', 'X') else fTobDet.PutValue('ATA_JOUR4', '-');
          if TCheckBox(GetControl('ATA_JOUR5M')).Checked then fTobDet.PutValue('ATA_JOUR5', 'X') else fTobDet.PutValue('ATA_JOUR5', '-');
          if TCheckBox(GetControl('ATA_JOUR6M')).Checked then fTobDet.PutValue('ATA_JOUR6', 'X') else fTobDet.PutValue('ATA_JOUR6', '-');
          if TCheckBox(GetControl('ATA_JOUR7M')).Checked then fTobDet.PutValue('ATA_JOUR7', 'X') else fTobDet.PutValue('ATA_JOUR7', '-');
        end;
    end;
end;

procedure TOF_AFTACHES.PutRegles;
begin
  // raz de l'écran
  InitNewRegles;

  if fTobDet.GetValue('ATA_MODEGENE') = 1 then
    fRB_MODEGENE1.Checked := True
  else
    fRB_MODEGENE2.Checked := True;

  SetControlText('ATA_QTEINTERVENT', ftobDet.getvalue('ATA_QTEINTERVENT'));
  SetControlText('ATA_MOISJOURFIXE', fTobDet.GetValue('ATA_MOISJOURFIXE'));

  if fTobDet.GetValue('ATA_PERIODICITE') = 'Q' then fRB_QUOTIDIENNE.Checked := true
  else if fTobDet.GetValue('ATA_PERIODICITE') = 'NBI' then fRB_NBINTERVENTION.Checked := true
  else if fTobDet.GetValue('ATA_PERIODICITE') = 'S' then
    begin                                           
      fRB_HEBDOMADAIRE.Checked := true;
      TCheckBox(GetControl('ATA_JOUR1H')).checked := fTobDet.GetValue('ATA_JOUR1') = 'X';
      TCheckBox(GetControl('ATA_JOUR2H')).checked := fTobDet.GetValue('ATA_JOUR2') = 'X';
      TCheckBox(GetControl('ATA_JOUR3H')).checked := fTobDet.GetValue('ATA_JOUR3') = 'X';
      TCheckBox(GetControl('ATA_JOUR4H')).checked := fTobDet.GetValue('ATA_JOUR4') = 'X';
      TCheckBox(GetControl('ATA_JOUR5H')).checked := fTobDet.GetValue('ATA_JOUR5') = 'X';
      TCheckBox(GetControl('ATA_JOUR6H')).checked := fTobDet.GetValue('ATA_JOUR6') = 'X';
      TCheckBox(GetControl('ATA_JOUR7H')).checked := fTobDet.GetValue('ATA_JOUR7') = 'X';
    end
  else if fTobDet.GetValue('ATA_PERIODICITE') = 'A' then fRB_ANNUELLE.Checked := true
  else if fTobDet.GetValue('ATA_PERIODICITE') = 'M' then
    begin
      fRB_MENSUELLE.Checked := true;
      if fTobDet.GetValue('ATA_MOISMETHODE') =  '1' then
        begin
          fRB_MOISMETHODE1.Checked := true;
          SetControlText('ATA_MOISFIXE0', fTobDet.GetValue('ATA_MOISFIXE'));
        end
      else if fTobDet.GetValue('ATA_MOISMETHODE') = '2' then
        begin
          fRB_MOISMETHODE2.Checked := true;
          THValComboBox(GetControl('ATA_MOISSEMAINE1')).value := fTobDet.GetValue('ATA_MOISSEMAINE');
          SetControlText('ATA_MOISFIXE1', fTobDet.GetValue('ATA_MOISFIXE'));
          THValComboBox(GetControl('ATA_MOISJOURLIB')).Value := fTobDet.GetValue('ATA_MOISJOURLIB');
        end                           
      else if fTobDet.GetValue('ATA_MOISMETHODE') = '3' then
        begin
          fRB_MOISMETHODE3.Checked := true;
          THValComboBox(GetControl('ATA_MOISSEMAINE2')).value := fTobDet.GetValue('ATA_MOISSEMAINE');
          SetControlText('ATA_MOISFIXE2', fTobDet.GetValue('ATA_MOISFIXE'));
          TCheckBox(GetControl('ATA_JOUR1M')).checked := fTobDet.GetValue('ATA_JOUR1') = 'X';
          TCheckBox(GetControl('ATA_JOUR2M')).checked := fTobDet.GetValue('ATA_JOUR2') = 'X';
          TCheckBox(GetControl('ATA_JOUR3M')).checked := fTobDet.GetValue('ATA_JOUR3') = 'X';
          TCheckBox(GetControl('ATA_JOUR4M')).checked := fTobDet.GetValue('ATA_JOUR4') = 'X';
          TCheckBox(GetControl('ATA_JOUR5M')).checked := fTobDet.GetValue('ATA_JOUR5') = 'X';
          TCheckBox(GetControl('ATA_JOUR6M')).checked := fTobDet.GetValue('ATA_JOUR6') = 'X';
          TCheckBox(GetControl('ATA_JOUR7M')).checked := fTobDet.GetValue('ATA_JOUR7') = 'X';
        end;                                                                                 
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.InitNewTache (TobTache : TOB);
var
  vStTypeArticle  : String;
  vStArticle      : String;
  vStCodeArticle  : String;
  vStFacturable   : String;
Begin

  TobTache.PutValue('ATA_TIERS', fCurAffaire.StTiers);
  TobTache.PutValue('ATA_AFFAIRE',fCurAffaire.StAffaire);
  TobTache.PutValue('ATA_AFFAIRE0',fCurAffaire.StAff0);
  TobTache.PutValue('ATA_AFFAIRE1',fCurAffaire.StAff1);
  TobTache.PutValue('ATA_AFFAIRE2',fCurAffaire.StAff2);
  TobTache.PutValue('ATA_AFFAIRE3',fCurAffaire.StAff3);
  TobTache.PutValue('ATA_AVENANT', fCurAffaire.StAvenant);
  TobTache.PutValue('ATA_DEVISE', fCurAffaire.StDevise);
  TobTache.PutValue('ATA_PERIODICITE', 'M');
  TobTache.PutValue('ATA_DATEDEBPERIOD', fDtDebutAffaire);
  TobTache.PutValue('ATA_DATEFINPERIOD', fDtFinAffaire);
  TobTache.PutValue('ATA_MOISSEMAINE', '1');
  TobTache.PutValue('ATA_MODEGENE', '1');
  TobTache.PutValue('ATA_QTEINTERVENT', '1');
  // ajout des parametres de decalage 
  TobTache.PutValue('ATA_NBJOURSDECAL', GetParamSoc('SO_AFJOURSDECAL'));
  TobTache.PutValue('ATA_METHODEDECAL', GetParamSoc('SO_AFJOURSPLANIF'));

  // article
  TobTache.PutValue('ATA_CODEARTICLE', getparamsoc('SO_AFPRESTDEFAUT'));
  vStCodeArticle := getparamsoc('SO_AFPRESTDEFAUT');
  if controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable) then
    begin
      TobTache.PutValue('ATA_TYPEARTICLE', vStTypeArticle);
      TobTache.PutValue('ATA_ARTICLE', vStArticle);
      TobTache.PutValue('ATA_ACTIVITEREPRIS', vStFacturable); // AB-030325-Facturable
    end
  else
    begin
      TobTache.PutValue('ATA_CODEARTICLE','');
      TobTache.PutValue('ATA_TYPEARTICLE', '');
      TobTache.PutValue('ATA_ARTICLE', '');
    end;

  // pour l'instant, on initialise avec jour comme unité de saisie
  TobTache.PutValue('ATA_UNITETEMPS', 'J');
  TobTache.PutValue('ATA_FAMILLETACHE', getparamsoc('SO_AFFAMILLEDEF'));

  // maj du nb de res
  UpdateGridRes(False, 0);

End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : affiche le contenu de la tob a l'ecran
Description .. : et retourne l'indice dans la grille
Mots clefs ... :
*****************************************************************}
function TOF_AFTACHES.TobToScreen(Row : Integer) : integer;
Begin

  if Row <> -1 then
    begin
      fTobDet := fTobTaches.Detail[Row];
      result := Row + 1;
    end
  else
    begin
      fTobDet := TOB(fGSTaches.Objects[0,fGSTaches.Row]);
      result := fGSTaches.Row;
    end;
 
  fTobDet.modifie := false;
  fTobDet.PutEcran (Ecran,fPaDet);
  ArticleOnChange(self);
  SetControlText('ATA_DATEDEBPERIOD_', GetControlText('ATA_DATEFINPERIOD'));
  PutRegles;

  // les taches ressources
  TobResToScreen;

  if fTobDet.Detail.Count <> 0 then
    UpdateGridRes(False, fTobDet.Detail.Count)
  else
    UpdateGridRes(False, 0);

  // ajout de la gestion du type de tache
  RefreshModeSaisie;
  RefreshLignePiece; //AB-LigneAffaire-Champs non modifiables

End;

{procedure TOF_AFTACHES.RefreshModeSaisie;
begin

  if THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA' then
    begin
      if fBOAFGESTIONRAP or fBoAFPLANDECHARGE then
      begin
        SetControlVisible('PA_QUANTITE2', true);
        SetControlVisible('PA_QUANTITE1', true);
      end
      else
      begin
        SetControlVisible('PA_QUANTITE2', false);
        SetControlVisible('PA_QUANTITE1', false);
      end;

      SetControlVisible('PA_MONTANT1', false);
      SetControlVisible('PA_MONTANT2', false);
      SetControlEnabled('TAB_VALORISATION', true);
    end
  else
    begin
      SetControlVisible('PA_QUANTITE1', false);
      SetControlVisible('PA_QUANTITE2', false);

      if fBOAFGESTIONRAP or fBoAFPLANDECHARGE then
      begin
        SetControlVisible('PA_MONTANT1', true);
        SetControlVisible('PA_MONTANT2', true);
      end
      else
      begin
        SetControlVisible('PA_MONTANT1', true);
        SetControlVisible('PA_MONTANT2', false);
      end;

      SetControlEnabled('TAB_VALORISATION', false);
    end;
end;
}

procedure TOF_AFTACHES.RefreshModeSaisie;
begin

  if THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA' then
    begin
      if fBOAFGESTIONRAP or fBoAFPLANDECHARGE then
      begin
        SetControlVisible('PA_QUANTITE2', true);
        SetControlVisible('PA_QUANTITE1', true);
      end
      else
      begin
        SetControlVisible('PA_QUANTITE2', false);
        SetControlVisible('PA_QUANTITE1', false);
      end;

      SetControlVisible('PA_MONTANT1', false);
      SetControlVisible('PA_MONTANT2', false);
      SetControlEnabled('TAB_VALORISATION', true);
    end
  else
    begin
      if fBOAFGESTIONRAP or fBoAFPLANDECHARGE then
      begin
        SetControlVisible('PA_MONTANT1', true);
        SetControlVisible('PA_MONTANT2', true);
      end
      else
      begin
        SetControlVisible('PA_MONTANT1', false);
        SetControlVisible('PA_MONTANT2', false);
      end;
 
      SetControlVisible('PA_QUANTITE1', false);
      SetControlVisible('PA_QUANTITE2', false);
      SetControlEnabled('TAB_VALORISATION', false);
    end;
end;

procedure TOF_AFTACHES.TobResToScreen;
var
  i           : Integer;

begin

  RefreshGridRes;
  for i := 0 to fTobDet.Detail.Count - 1 do
    begin
      fGSRes.Cells[cInColRes, i + 1] := fTobDet.Detail[i].GetValue('ATR_RESSOURCE');
      fGSRes.Cells[cInColQte, i + 1] := fTobDet.Detail[i].GetValue('ATR_QTEINITIALE');
      //fGSRes.Cells[cInColStatut, i + 1] := RechDom('AFSTATUTRESSOURCE', fTobDet.Detail[i].GetValue('ATR_STATUTRES'), false);
      fGSRes.Cells[cInColStatut, i + 1] := '#ICO#' + RechDom('AFSTATUTRESSOURCE', fTobDet.Detail[i].GetValue('ATR_STATUTRES'), false);

      if (fTobDet.Detail[i].GetValue('ARS_LIBELLE') = '') or
         (fTobDet.Detail[i].GetValue('ARS_LIBELLE') = #0) then
        fGSRes.Cells[cInColResLib, i + 1] := RessourceLib (fGSRes.Cells[cInColRes, i + 1])
      else
        fGSRes.Cells[cInColResLib, i + 1] := fTobDet.Detail[i].GetValue('ARS_LIBELLE');
      if fTobDet.Detail[i].GetValue('ATR_COMPETENCE1') <> '' then
        fGSRes.Cells[cInColComp1, i + 1] := (fTobCompet.FindFirst(['ACO_COMPETENCE'], [fTobDet.Detail[i].GetValue('ATR_COMPETENCE1')], true)).getvalue('ACO_LIBELLE');
      if fTobDet.Detail[i].GetValue('ATR_COMPETENCE2') <> '' then
        fGSRes.Cells[cInColComp2, i + 1] := (fTobCompet.FindFirst(['ACO_COMPETENCE'], [fTobDet.Detail[i].GetValue('ATR_COMPETENCE2')], true)).getvalue('ACO_LIBELLE');
      if fTobDet.Detail[i].GetValue('ATR_COMPETENCE3') <> '' then
        fGSRes.Cells[cInColComp3, i + 1] := (fTobCompet.FindFirst(['ACO_COMPETENCE'], [fTobDet.Detail[i].GetValue('ATR_COMPETENCE3')], true)).getvalue('ACO_LIBELLE');

      // calcul des quantités RAP
      // calcul à la demande
{      if fBOAFGESTIONRAP then
      begin
        vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire,
                                      fTobDet.Detail[i].GetValue('ATR_NUMEROTACHE'),
                                      'APL_QTEPLANIFIEE', '', true);
        vRdQte := StrToFloat(fTobDet.Detail[i].GetValue('ATR_QTEINITIALE')) - vRdPlanifie;
        fGSRes.Cells[cInColQteRAP, i + 1] := floatToStr(vRdQte);
      end;
}
    end;

  if fTobDet.Detail.Count <> 0 then
    begin
      fGSRes.Rowcount := fTobDet.Detail.Count + 1;
      fGSRes.enabled := True;
    end
  else
    begin
      fGSRes.Rowcount := 2;
      fGSRes.enabled := False;
    end;
end;

procedure TOF_AFTACHES.ScreenToTobRes;
var
  i       : Integer;
  vTobRes : Tob;

begin

  for i := 1 to fGSRes.rowcount - 1 do
    begin

      vTobRes := fTobDet.FindFirst(['ATR_RESSOURCE'], [fGSRes.Cells[cInColRes,i]], true);

      if fGSRes.Cells[cInColRes,i] <> '' then
        begin
          if (vTobRes = nil) then
            begin
              vTobRes := TOB.Create('TACHERESSOURCE', fTobDet, -1);
              initNewRessource(vTobRes);
            end;

          vTobRes.putValue('ATR_RESSOURCE', fGSRes.Cells[cInColRes,i]);
          vTobRes.putValue('ARS_LIBELLE', fGSRes.Cells[cInColResLib,i]);

          //
          vTobRes.putValue('ATR_QTEINITIALE', valeur(fGSRes.Cells[cInColQte,i]));
          vTobRes.putValue('ATR_QTEINITUREF', floatToStr(conversionUnite(GetControlText('ATA_UNITETEMPS'),
                                                                         getparamsoc('SO_AFMESUREACTIVITE'),
                                                                         valeur(fGSRes.Cells[cInColQte,i]))));

          if fGSRes.Cells[cInColComp1,i] <> '' then
            vTobRes.putValue('ATR_COMPETENCE1', (fTobCompet.FindFirst(['ACO_LIBELLE'], [fGSRes.Cells[cInColComp1,i]], true).GetValue('ACO_COMPETENCE')))
          else
            vTobRes.putValue('ATR_COMPETENCE1', '');

          if fGSRes.Cells[cInColComp2,i] <> '' then
            vTobRes.putValue('ATR_COMPETENCE2', (fTobCompet.FindFirst(['ACO_LIBELLE'], [fGSRes.Cells[cInColComp2,i]], true).GetValue('ACO_COMPETENCE')))
          else
            vTobRes.putValue('ATR_COMPETENCE2', '');

          if fGSRes.Cells[cInColComp3,i] <> '' then
            vTobRes.putValue('ATR_COMPETENCE3', (fTobCompet.FindFirst(['ACO_LIBELLE'], [fGSRes.Cells[cInColComp3,i]], true).GetValue('ACO_COMPETENCE')))
          else
            vTobRes.putValue('ATR_COMPETENCE3', '');
        end;
    end;
end;

procedure TOF_AFTACHES.InitNewRessource(vTobRes : Tob);
begin
  vTobRes.AddChampSup('ARS_LIBELLE', True);
  vTobRes.putValue('ATR_TIERS', fCurAffaire.StTiers);
  vTobRes.putValue('ATR_AFFAIRE', fCurAffaire.StAffaire);
  vTobRes.putValue('ATR_AFFAIRE0', fCurAffaire.StAff0);
  vTobRes.putValue('ATR_AFFAIRE1', fCurAffaire.StAff1);
  vTobRes.putValue('ATR_AFFAIRE2', fCurAffaire.StAff2);
  vTobRes.putValue('ATR_AFFAIRE3', fCurAffaire.StAff3);
  vTobRes.putValue('ATR_AVENANT', fCurAffaire.StAvenant);
  // ajout de la fonction de la tache juste avant l'enregistrement : ATR_FONCTION
  vTobRes.putValue('ATR_QTEAPLANIFIER', 0);
  vTobRes.putValue('ATR_POURCENTAGE', 0);

  vTobRes.putValue('ATR_DUREELUNDI', 0);
  vTobRes.putValue('ATR_DUREEMARDI', 0);
  vTobRes.putValue('ATR_DUREEMERCREDI', 0);
  vTobRes.putValue('ATR_DUREEJEUDI', 0);
  vTobRes.putValue('ATR_DUREEVENDREDI', 0);
  vTobRes.putValue('ATR_DUREESAMEDI', 0);
  vTobRes.putValue('ATR_DUREEDIMANCHE', 0);

  vTobRes.putValue('ATR_HDEBLUNDI', 0);
  vTobRes.putValue('ATR_HDEBMARDI', 0);
  vTobRes.putValue('ATR_HDEBMERCREDI', 0);
  vTobRes.putValue('ATR_HDEBJEUDI', 0);
  vTobRes.putValue('ATR_HDEBVENDREDI', 0);
  vTobRes.putValue('ATR_HDEBSAMEDI', 0);
  vTobRes.putValue('ATR_HDEBDIMANCHE', 0);

  vTobRes.putValue('ATR_HFINLUNDI', 0);
  vTobRes.putValue('ATR_HFINMARDI', 0);
  vTobRes.putValue('ATR_HFINMERCREDI', 0);
  vTobRes.putValue('ATR_HFINJEUDI', 0);
  vTobRes.putValue('ATR_HFINVENDREDI', 0);
  vTobRes.putValue('ATR_HFINSAMEDI', 0);
  vTobRes.putValue('ATR_HFINDIMANCHE', 0);
  vTobRes.putValue('ATR_REGLESRES', '-');
  vTobRes.putValue('ATR_MODEGENE', 0);
  vTobRes.putValue('ATR_QTEINTERVENT', 1);

//  vTobRes.putValue('ATR_PERIODICITE', 'Q');
//  vTobRes.putValue('ATR_DATEANNUELLE', '');
  vTobRes.putValue('ATR_JOURINTERVAL', 1);
  vTobRes.putValue('ATR_SEMAINEINTERV', 1);
  vTobRes.putValue('ATR_MOISMETHODE', cInMoisMethode1);
  vTobRes.putValue('ATR_MOISJOURFIXE', 1);
  vTobRes.putValue('ATR_MOISFIXE', 1);
  vTobRes.putValue('ATR_MOISSEMAINE', '');
  vTobRes.putValue('ATR_MOISJOURLIB', '');

  // les prix sont alimentés lors de la valorisation

  vTobRes.putValue('ATR_JOUR1', '-');
  vTobRes.putValue('ATR_JOUR2', '-');
  vTobRes.putValue('ATR_JOUR3', '-');
  vTobRes.putValue('ATR_JOUR4', '-');
  vTobRes.putValue('ATR_JOUR5', '-');
  vTobRes.putValue('ATR_JOUR6', '-');
  vTobRes.putValue('ATR_JOUR7', '-');
  vTobRes.putValue('ATR_JOURFERIE', '-');
  vTobRes.putValue('ATR_GESTIONDATEFIN', '-');
  vTobRes.putValue('ATR_STATUTRES', 'ACT');
                                         
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : recherche l'indice dans la tob de la tache courante
                 dans la grille
Mots clefs ... :
*****************************************************************}
function TOF_AFTACHES.IndiceTobTacheCourante : Integer;
var
    i : Integer;

Begin

  result := 1;
  if (fTobTaches <> Nil) and (fTobTaches.Detail.count > 0) then
    Begin
      For i:=0 to fTobTaches.Detail.count-1 do
        Begin
          if fTobDet = nil then
            begin
              fTobDet := fTobTaches.Detail[i];
              fTobDet.modifie := false;
            end
          else
            fTobDet := fTobTaches.Detail[i];

          // il n'y a pas de tache selectionné -> on se positionne sur la premiere
          if (fTobDet <> nil) and (fStOpenNumTache = '') then
            begin
              fStOpenNumTache := varAsType(fTobDet.GetValue('ATA_NUMEROTACHE'), varString);
              result := i;
              exit;
            end
          else if (fTobDet <> nil) and (varAsType(fTobDet.GetValue('ATA_NUMEROTACHE'),varString) = fStOpenNumTache) then
            Begin
              result := i;
              exit;
            End;
        End;
    end;
end;

procedure TOF_AFTACHES.fRB_MODEGENE1OnClick(Sender: TObject);
begin
  fRB_MODEGENE2.checked := not fRB_MODEGENE1.Checked;
  //SetControlEnabled('ATA_GESTIONDATEFIN', not fRB_MODEGENE2.checked);
end;

procedure TOF_AFTACHES.fRB_MODEGENE2OnClick(Sender: TObject);
begin
  fRB_MODEGENE1.Checked := not fRB_MODEGENE2.Checked;
  //SetControlEnabled('ATA_GESTIONDATEFIN', fRB_MODEGENE1.checked);
end;

procedure TOF_AFTACHES.fRB_HEBDOMADAIREOnClick(Sender: TObject);
begin
  fRB_QUOTIDIENNE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_MENSUELLE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_ANNUELLE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_NBINTERVENTION.checked := not fRB_HEBDOMADAIRE.checked;
  fPC_FREQUENCE.ActivePage := fTS_HEBDOMADAIRE;
  //SetControlEnabled('ATA_GESTIONDATEFIN', true);
  if GetControltext('ATA_QTEINITIALE') <> '' then
  begin
    SetControlEnabled('RB_MODEGENE2', StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0);
    SetControlEnabled('RB_MODEGENE1', StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0);
  end
  else
  begin
    SetControlEnabled('RB_MODEGENE2', false);
    SetControlEnabled('RB_MODEGENE1', false);
  end;
end;

procedure TOF_AFTACHES.fRB_QUOTIDIENNEOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_MENSUELLE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_ANNUELLE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_NBINTERVENTION.checked := not fRB_QUOTIDIENNE.checked;
  fPC_FREQUENCE.ActivePage := fTS_QUOTIDIENNE;
  //SetControlEnabled('ATA_GESTIONDATEFIN', true);
  if GetControltext('ATA_QTEINITIALE') <> '' then
  begin
    SetControlEnabled('RB_MODEGENE2', StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0);
    SetControlEnabled('RB_MODEGENE1', StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0);
  end
  else
  begin
    SetControlEnabled('RB_MODEGENE2', false);
    SetControlEnabled('RB_MODEGENE1', false);
  end;
end;

procedure TOF_AFTACHES.fRB_MENSUELLEOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_MENSUELLE.checked;
  fRB_QUOTIDIENNE.checked := not fRB_MENSUELLE.checked;
  fRB_ANNUELLE.checked := not fRB_MENSUELLE.checked;
  fRB_NBINTERVENTION.checked := not fRB_MENSUELLE.checked;
  fPC_FREQUENCE.ActivePage := fTS_MENSUELLE;
  //SetControlEnabled('ATA_GESTIONDATEFIN', true);
  if GetControltext('ATA_QTEINITIALE') <> '' then
  begin
    SetControlEnabled('RB_MODEGENE2', StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0);
    SetControlEnabled('RB_MODEGENE1', StrToFloat(GetControltext('ATA_QTEINITIALE')) <> 0);
  end
  else
  begin
    SetControlEnabled('RB_MODEGENE2', false);
    SetControlEnabled('RB_MODEGENE1', false);
  end;
end;

procedure TOF_AFTACHES.fRB_ANNUELLEOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_ANNUELLE.checked;
  fRB_QUOTIDIENNE.checked := not fRB_ANNUELLE.checked;
  fRB_MENSUELLE.checked := not fRB_ANNUELLE.checked;
  fRB_NBINTERVENTION.checked := not fRB_ANNUELLE.checked;
  fPC_FREQUENCE.ActivePage := fTS_ANNUELLE;
  //SetControlEnabled('ATA_GESTIONDATEFIN', false);
  SetControlEnabled('RB_MODEGENE2', false);
  SetControlEnabled('RB_MODEGENE1', false);
end;

procedure TOF_AFTACHES.fRB_NBINTERVENTIONOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_NBINTERVENTION.checked;
  fRB_QUOTIDIENNE.checked := not fRB_NBINTERVENTION.checked;
  fRB_MENSUELLE.checked := not fRB_NBINTERVENTION.checked;
  fRB_ANNUELLE.checked := not fRB_NBINTERVENTION.checked;
  fPC_FREQUENCE.ActivePage := fTS_NBINTERVENTION;
  //SetControlEnabled('ATA_GESTIONDATEFIN', false);
  SetControlEnabled('RB_MODEGENE2', false);
  SetControlEnabled('RB_MODEGENE1', false);
end;

procedure TOF_AFTACHES.fRB_MOISMETHODE1OnClick(SEnder: TObject);
begin
  fRB_MOISMETHODE2.checked := not fRB_MOISMETHODE1.checked;
  fRB_MOISMETHODE3.checked := not fRB_MOISMETHODE1.checked;

  SetControlEnabled('ATA_MOISJOURFIXE', True);
  SetControlEnabled('ATA_MOISFIXE0', True);

  SetControlEnabled('ATA_MOISSEMAINE1', false);
  SetControlEnabled('ATA_MOISJOURLIB', false);
  SetControlEnabled('ATA_MOISFIXE1', false);

  SetControlEnabled('ATA_MOISSEMAINE2', false);
  SetControlEnabled('ATA_MOISFIXE2', false);
  SetControlEnabled('ATA_JOUR1M', false);
  SetControlEnabled('ATA_JOUR2M', false);
  SetControlEnabled('ATA_JOUR3M', false);
  SetControlEnabled('ATA_JOUR4M', false);
  SetControlEnabled('ATA_JOUR5M', false);
  SetControlEnabled('ATA_JOUR6M', false);
  SetControlEnabled('ATA_JOUR7M', false);
end;

procedure TOF_AFTACHES.fRB_MOISMETHODE2OnClick(SEnder: TObject);
begin
  fRB_MOISMETHODE1.checked := not fRB_MOISMETHODE2.checked;
  fRB_MOISMETHODE3.checked := not fRB_MOISMETHODE2.checked;

  SetControlEnabled('ATA_MOISJOURFIXE', false);
  SetControlEnabled('ATA_MOISFIXE0', false);

  SetControlEnabled('ATA_MOISSEMAINE1', true);
  SetControlEnabled('ATA_MOISJOURLIB', true);
  SetControlEnabled('ATA_MOISFIXE1', true);

  SetControlEnabled('ATA_MOISSEMAINE2', false);
  SetControlEnabled('ATA_MOISFIXE2', false);
  SetControlEnabled('ATA_JOUR1M', false);
  SetControlEnabled('ATA_JOUR2M', false);
  SetControlEnabled('ATA_JOUR3M', false);
  SetControlEnabled('ATA_JOUR4M', false);
  SetControlEnabled('ATA_JOUR5M', false);
  SetControlEnabled('ATA_JOUR6M', false);
  SetControlEnabled('ATA_JOUR7M', false);
end;

procedure TOF_AFTACHES.fRB_MOISMETHODE3OnClick(SEnder: TObject);
begin
  fRB_MOISMETHODE1.checked := not fRB_MOISMETHODE3.checked;
  fRB_MOISMETHODE2.checked := not fRB_MOISMETHODE3.checked;

  SetControlEnabled('ATA_MOISJOURFIXE', false);
  SetControlEnabled('ATA_MOISFIXE0', false);

  SetControlEnabled('ATA_MOISSEMAINE1', false);
  SetControlEnabled('ATA_MOISJOURLIB', false);
  SetControlEnabled('ATA_MOISFIXE1', false);

  SetControlEnabled('ATA_MOISSEMAINE2', true);
  SetControlEnabled('ATA_MOISFIXE2', true);
  SetControlEnabled('ATA_JOUR1M', true);
  SetControlEnabled('ATA_JOUR2M', true);
  SetControlEnabled('ATA_JOUR3M', true);
  SetControlEnabled('ATA_JOUR4M', true);
  SetControlEnabled('ATA_JOUR5M', true);
  SetControlEnabled('ATA_JOUR6M', true);
  SetControlEnabled('ATA_JOUR7M', true);
end;

procedure TOF_AFTACHES.TraitementArgument(pStArgument : String);
var
  Tmp     : String;
  champ   : String;
  valeur  : String;

begin
  // traitement des arguments
  Tmp:=(Trim(ReadTokenSt(pStArgument)));
  While (Tmp <>'') do
    Begin
      If Tmp<>'' then
        Begin
          DecodeArgument(Tmp, Champ, valeur);

          If Champ='ATA_AFFAIRE'      then fCurAffaire.StAffaire := valeur else
          If Champ='ATA_NUMEROTACHE'  then fStOpenNumTache := valeur else
          If Champ='ATA_IDENTLIGNE'  then fStLignePiece := findEtreplace (valeur,'|',';',true) else
          If Champ='ACTION'           then
            Begin
              If valeur='MODIFICATION' then  fAction := taModIf else
              If valeur='CONSULTATION' then  fAction := taConsult ;
            End;
        End;
        Tmp:=(Trim(ReadTokenSt(pStArgument)));
    End;
end;

procedure TOF_AFTACHES.LoadAffaire(pInIndiceTob : Integer);
begin

  // Gestion des Grid
  InitGrid;
  initGridRes;

  // chargement du tiers et formatage des libelles affaire et tiers
  AffichageEntete;

  if assigned(fTobTaches) then fTobTaches.Free;
  if assigned(fTobSup) then fTobSup.Free;
  if assigned(fTobSupRes) then fTobSupRes.Free;

  fTobTaches := nil;
  fTobSup := nil;
  fTobSupRes := nil;

  // Chargement des taches
  fTobTaches := TOB.Create ('Liste des taches',Nil,-1);
  fTobSup := TOB.Create ('Liste taches sup',Nil,-1);
  fTobSupRes := TOB.Create ('Liste des ressources modifiées', Nil, -1);

  // des tâches existent ...
  If ChargeLesTaches(fTobTaches, fCurAffaire.StAffaire, '') then
  begin
    fTobTaches.PutGridDetailOnListe (fGSTaches,'AFLISTETACHES');
    fGSTaches.Fixedcols := 1;
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSTaches);
    //fTobTaches.PutGridDetail(fGSTaches, True, True, fStLesCols, True) //AB-LigneAffaire-
  end
  else
    // si aucune tache, on entre en creation
    fStatut := taCreat;
  TraiteLignePiece;  // AB-LigneAffaire-
  //SetControlVisible('BTLIGNE',False);
  if (fStatut = taCreat) then
    begin
      InsertTache;
//    fGSTaches.RowCount := 2;
    end
  else
    // affiche la ligne sélectionnée positionnement dans la grille
    If fTobTaches.Detail.Count > 0 then
      begin
        if pInIndiceTob = 0 then
          fGSTaches.Row := TobToScreen(IndiceTobTacheCourante)
        else
          fGSTaches.Row := TobToScreen(pInIndiceTob);
      end;
  SetActiveTabSheet('TAB_TACHE');
  SetFocusControl('ATA_LIBELLETACHE1');

  loadChampsCalc;

  if fStLignePiece <> '' then // AB-LigneAffaire-
  begin
    if fTobModele <> nil then fTobModele.free;
    fStLignePiece := '';
    fTobModele := nil;
    SetControlVisible('BTLIGNE',False);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 24/04/2002
Modifié le ... :
Description .. :
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.GestionEcran;
var
  Combo             : THValComboBox;
  Bt                : TToolBarButton97;
  vCBTypeArticle    : THValComboBox;
  vMenu             : TMenuItem;
  Ed                : THEdit;
  pcPage            : TPageControl;
begin

  // gestion des touches
  TFVierge(Ecran).OnKeyDown:=FormKeyDown ;

  // Gestion des Boutons
  Bt := TToolBarButton97(GetControl('BInsert'));
  If (Bt<>Nil) then Bt.Onclick:=BInsertOnClick;

  Bt := TToolBarButton97(GetControl('BInsertModele')); // AB-Modeletache-Création à partir d'un modele
  If (Bt<>Nil) then Bt.Onclick:=bInsertModeleOnClick;

  Bt := TToolBarButton97(GetControl('BInsertRes'));
  If (Bt<>Nil) then Bt.Onclick:=BInsertResOnClick;

  Bt := TToolBarButton97(GetControl('BDelete'));
  If (Bt<>Nil) then Bt.Onclick:=BDeleteOnClick;

  Bt := TToolBarButton97(GetControl('BDeleteRes'));
  If (Bt<>Nil) then Bt.Onclick:=BDeleteResOnClick;

  Bt := TToolbarButton97(GetControl('BFERME'));
  If (Bt<>Nil) then Bt.Onclick:=bFermerOnClick;

  Bt := TToolbarButton97(GetControl('BMEMO'));
  If (Bt<>Nil) then Bt.Onclick:=bMemoOnClick;

  Bt := TToolbarButton97(GetControl('BPLANNINGVIEWER'));
  if bt<>nil then Bt.Onclick := BPLANNINGVIEWEROnClick;

  Bt := TToolbarButton97(GetControl('BPLANNINGRES'));
  if bt<>nil then Bt.Onclick := BPLANNINGRESOnClick;

  Bt := TToolbarButton97(GetControl('BPLANDECHARGE'));
  if bt<>nil then Bt.Onclick := BPLANDECHARGEOnClick;

  Bt := TToolbarButton97(GetControl('BPLANNING'));
  if bt<>nil then Bt.Onclick := BPLANNINGOnClick;

  Bt := TToolbarButton97(GetControl('BRESSOURCE'));
  if bt<>nil then Bt.Onclick := BRESSOURCEOnClick;

  Bt := TToolbarButton97(GetControl('BCALCULMNTRAP'));
  if bt<>nil then Bt.Onclick := BCALCULMNTRAPOnClick;

  Bt := TToolbarButton97(GetControl('BCALCULQTERAP'));
  if bt<>nil then Bt.Onclick := BCALCULQTERAPOnClick;

  Bt := TToolbarButton97(GetControl('BCALCULRAPRES'));
  if bt<>nil then Bt.Onclick := BCALCULRAPRESOnClick;

  Bt := TToolbarButton97(GetControl('BTLIGNE')); //AB-LigneAffaire-
  if bt<>nil then Bt.Onclick := BTLIGNEOnClick;
  
  vMenu := TMenuItem(GetControl('mnGENEAFFAIRE'));
  if vMenu <> nil then vMenu.OnClick := mnGeneAffaire_OnClick;

  vMenu := TMenuItem(GetControl('mnGENETACHE'));
  if vMenu <> nil then vMenu.OnClick := mnGeneTache_OnClick;

  pcPage := TPageControl(getControl('PAGETACHE'));
  if pcPage <> nil then pcpage.OnChange := PAGETACHEChange;

  // Evènements de la tâche en cours
  //Combo := THValComboBox(GetControl('ATA_PERIODICITE'));
  //If Combo <> Nil then Combo.Onclick:=ATA_PERIODICITEOnClick;

  fED_AFFAIRE := THEdit(GetControl('ATA_AFFAIRE'));

  fPaDet := THPanel(GetControl('PCLIENT'));

  ecran.OnCloseQuery := EcranOnCloseQuery;

  // gestion de l'onglet règles
  fRB_MODEGENE1 := TRadioButton(GetControl('RB_MODEGENE1'));
  fRB_MODEGENE2 := TRadioButton(GetControl('RB_MODEGENE2'));
  fRB_MODEGENE1.OnClick := fRB_MODEGENE1OnClick;
  fRB_MODEGENE2.OnClick := fRB_MODEGENE2OnClick;

  fRB_QUOTIDIENNE     := TRadioButton(GetControl('RB_QUOTIDIENNE'));
  fRB_NBINTERVENTION  := TRadioButton(GetControl('RB_NBINTERVENTION'));
  fRB_HEBDOMADAIRE    := TRadioButton(GetControl('RB_HEBDOMADAIRE'));
  fRB_ANNUELLE        := TRadioButton(GetControl('RB_ANNUELLE'));
  fRB_MENSUELLE       := TRadioButton(GetControl('RB_MENSUELLE'));

  fRB_QUOTIDIENNE.OnClick     := fRB_QUOTIDIENNEOnClick;
  fRB_NBINTERVENTION.OnClick  := fRB_NBINTERVENTIONOnClick;
  fRB_HEBDOMADAIRE.OnClick    := fRB_HEBDOMADAIREOnClick;
  fRB_ANNUELLE.OnClick        := fRB_ANNUELLEOnClick;
  fRB_MENSUELLE.OnClick       := fRB_MENSUELLEOnClick;

  fPC_FREQUENCE       := TPageControl(GetControl('PC_FREQUENCE'));
  fTS_QUOTIDIENNE     := TTabSheet(GetControl('TS_QUOTIDIENNE'));
  fTS_NBINTERVENTION  := TTabSheet(GetControl('TS_NBINTERVENTION'));
  fTS_HEBDOMADAIRE    := TTabSheet(GetControl('TS_HEBDOMADAIRE'));
  fTS_ANNUELLE        := TTabSheet(GetControl('TS_ANNUELLE'));
  fTS_MENSUELLE       := TTabSheet(GetControl('TS_MENSUELLE'));

  fTS_QUOTIDIENNE.TabVisible := false;
  fTS_HEBDOMADAIRE.TabVisible := false;
  fTS_NBINTERVENTION.TabVisible := false;
  fTS_ANNUELLE.TabVisible := false;
  fTS_MENSUELLE.TabVisible := false;

  //TTabSheet(GetControl('TB_HEBDOMADAIRE')).TabVisible := False;

  {  fGB_NBINTERV := THPanel(GetControl('GB_NBINTERV'));
  fGB_PERIODE  := THPanel(GetControl('GB_PERIODE'));
 }

  fRB_MOISMETHODE1 := TRadioButton(GetControl('RB_MOISMETHODE1'));
  fRB_MOISMETHODE2 := TRadioButton(GetControl('RB_MOISMETHODE2'));
  fRB_MOISMETHODE3 := TRadioButton(GetControl('RB_MOISMETHODE3'));
  fRB_MOISMETHODE1.OnClick := fRB_MOISMETHODE1OnClick;
  fRB_MOISMETHODE2.OnClick := fRB_MOISMETHODE2OnClick;
  fRB_MOISMETHODE3.OnClick := fRB_MOISMETHODE3OnClick;


  combo := THValComboBox(GetControl('ATA_MODESAISIEPDC'));
  if combo <> nil then combo.OnChange := ATA_MODESAISIEPDCOnChange;

  combo := THValComboBox(GetControl('ATA_UNITETEMPS'));
  if combo <> nil then combo.OnChange := ATA_UNITETEMPSOnChange;

  Ed := THEdit(GetControl('ATA_QTEINITIALE'));
  if Ed <> nil then
  begin
    ed.OnChange := ATA_QTEINITIALEOnChange;
    ed.OnExit   := ATA_QTEINITIALEOnExit;
  end;

  Ed := THEdit(GetControl('ATA_PUPR'));
  if Ed <> nil then ed.OnChange := ATA_PUPROnChange;

  Ed := THEdit(GetControl('ATA_PUVENTEHT'));
  if Ed <> nil then ed.OnChange := ATA_PUVENTEHTOnChange;

  Ed := THEdit(GetControl('ATA_CODEARTICLE'));
  if Ed <> nil then
  begin
    ed.OnEnter := ATA_CODEARTICLEOnEnter;
    ed.OnExit := ATA_CODEARTICLEOnExit;
  end;
  combo := THValComboBox(GetControl('ATA_TYPEARTICLE'));
  if combo <> nil then combo.OnEnter := ATA_CODEARTICLEOnEnter;
                           

  // gestion du type d'article
  vCBTypeArticle := THValComboBox(GetControl('ATA_TYPEARTICLE'));
  vCBTypeArticle.plus := PlusTypeArticleTache;

  // gestion du planning récurrent
  setcontrolvisible('TAB_REGLES', not fBoAFPLANDECHARGE);

  // plan de charge cache la colonne affectée qui n'est pas calculée
  // de la même façon dans le standard et dans le plan de charge !
{  SetControlVisible('ATA_QTEAFFECTE', vBoPlanRecurrent);
  SetControlVisible('TATA_QTEAFFECTE', vBoPlanRecurrent);
  SetControlVisible('ATA_QTEREALISE', vBoPlanRecurrent);
  SetControlVisible('TATA_QTEREALISE', vBoPlanRecurrent);
  SetControlEnabled('ATA_QTEAPLANIFIERCALC', vBoPlanRecurrent);
  SetControlVisible('ATA_AFFPTPR', vBoPlanRecurrent);
  SetControlVisible('TATA_AFFPTPR', vBoPlanRecurrent);
  SetControlVisible('ATA_REALPTPR', vBoPlanRecurrent);
  SetControlVisible('TATA_REALPTPR', vBoPlanRecurrent);
  SetControlEnabled('ATA_RAPPTPRCALC', vBoPlanRecurrent);
  SetControlVisible('GB_PRIXAFF', vBoPlanRecurrent);
  SetControlVisible('GB_PRIXREAL', vBoPlanRecurrent);
}

  SetControlEnabled('ATA_MODESAISIEPDC', false);

  SetControlVisible('TATA_FONCTION', GetParamSoc('SO_AFFONCTION'));
  SetControlVisible('ATA_FONCTION', GetParamSoc('SO_AFFONCTION'));

  // apres la suppression des qtes dans les taches
//  SetControlVisible('ATA_RAPPTPRCALC', not vBoPlanRecurrent);
//  SetControlVisible('ATA_QTEAPLANIFIERCALC', not vBoPlanRecurrent);

  if fBoAFPLANDECHARGE then
    SetControlText('TATA_QTEAPLANIFIERCALC', 'Quantité Restant à faire')
  else
    SetControlText('TATA_QTEAPLANIFIERCALC', 'Quantité Restant à planifier');

  SetControlVisible('ATA_MODESAISIEPDC', fBoAFPLANDECHARGE);
  SetControlVisible('TATA_MODESAISIEPDC', fBoAFPLANDECHARGE);

  SetControlVisible('BDUPLICATION', false);
  SetControlVisible('BCALCULRAPRES', false);

  // si on est en mode de valorisation ressource
  // les prix unitaires de valorisation n'ont pas de sens
  if GetParamSoc('SO_AFVALOPR') = 'RES' then
  begin
    SetControlVisible('ATA_PUPR', false);
    SetControlVisible('TATA_PUPR', false);
  end;

  if GetParamSoc('SO_AFVALOPV') = 'RES' then
  begin                  
    SetControlVisible('ATA_PUVENTEHT', false);
    SetControlVisible('TATA_PUVENTEHT', false);
  end;

  if (GetParamSoc('SO_AFVALOPR') = 'RES') and
     (GetParamSoc('SO_AFVALOPV') = 'RES') then
    SetControlVisible('GB_PRIXUNITAIRE', false);

  if fAction=taConsult then
    SetControlEnabled('BInsertModele', false); // AB-Modeletache
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/04/2002
Modifié le ... :
Description .. : controle la saisie de la tcahe courante
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TOF_AFTACHES.ControleEcran : boolean;
var
  vStCodeArticle  : String;
  vStTypeArticle  : String;
  vStArticle      : String;
  vStFacturable   : String;
  vInRes          : Integer;

begin
  result := False;
  if ControleZonesOblig then
    Begin
      vInRes := ControleRessource;
      if vInRes = 0 then
        Begin
          //article
          vStCodeArticle  := THEdit(GetControl('ATA_CODEARTICLE')).Text;
          controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable);
          SetControlText('ATA_CODEARTICLE', vStCodeArticle);
          SetControlText('ATA_TYPEARTICLE', vStTypeArticle);
          SetControlText('ATA_ARTICLE', vStArticle);
          if (GetControlText('ATA_ACTIVITEREPRIS') = '') then
            SetControlText('ATA_ACTIVITEREPRIS', vStFacturable); // AB-030326-Facturable
          result := True;
        end
      else if vInRes = -1 then
        begin
          PGIBoxAF (TexteMsgTache[4],'');
          // vider la zone
          fGSRes.Cells[fGSRes.Col, fGSRes.Row] := '';
        end
      else if vInRes = -2 then
        PGIBoxAF (TexteMsgTache[9],'');
    End;
end;

procedure TOF_AFTACHES.AddUpdateRessList(aCol, aRow : Integer);
begin
  SetLength(fUpdateRes, length(fUpdateRes) + 1);
  fUpdateRes[length(fUpdateRes) - 1].StOldRes := fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE');
  fUpdateRes[length(fUpdateRes) - 1].StNewRes := fGSRes.Cells[ACol, ARow];
  fUpdateRes[length(fUpdateRes) - 1].rdAffaire := fCurAffaire;
  fUpdateRes[length(fUpdateRes) - 1].StNumTache := fTobDet.Detail[aRow - 1].GetValue('ATR_NUMEROTACHE');
  fUpdateRes[length(fUpdateRes) - 1].InStatut := cInUpdate;
end;

procedure TOF_AFTACHES.RefreshGrid;
var
  vInRow : Integer;
begin
  if fGSTaches.row > 0 then
    begin
      vInRow := fGSTaches.row;
//    fTobTaches.PutGridDetail(fGSTaches, True, True, fStLesCols, True);
      fTobTaches.PutGridDetailOnListe (fGSTaches,'AFLISTETACHES');
      fGSTaches.Fixedcols := 1;
      TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSTaches);
      fGSTaches.Row := vInRow;
      // lors de la creation, le numero de tache n'est pas rafraichit !
      if assigned(fTobDet) and (GetControlText('ATA_NUMEROTACHE') = '0') then
        begin
          fTobDet := TOB(fGSTaches.Objects[0,fGSTaches.Row]);
          SetControlText('ATA_NUMEROTACHE', fTobDet.GetValue('ATA_NUMEROTACHE'));
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.InsertTache;
var ARow :integer;
Begin
  if ClientFerme then PGIBoxAF (TexteMsgTache[31],'');
  fTobDet := TOB.Create('TACHE',fTobTaches,-1);
  fTobDet.AddChampSupValeur('ATA_QTEAPLANIFIERCALC', 0, false);
  fTobDet.AddChampSupValeur('ATA_RAPPTPRCALC', 0, false);
  fTobDet.AddChampSupValeur('ATA_RAPPTPRCAL', 0, false);
  fStatut := taCreat;

  if assigned(fGSTaches) then
    begin
//    fGSTaches.RowCount := fGSTaches.RowCount + 1;
//    fGSTaches.Row := fGSTaches.Rowcount - 1;
      Arow := fTobTaches.detail.count;
      fGSTaches.InsertRow (ARow);
      fGSTaches.Row := ARow;
      fGSTaches.RowCount := fTobTaches.detail.count + 1;
    end;

  RefreshGridRes;
  InitNewTache (fTobDet);
  InitNewRegles;
  if fTobModele <> nil then TOBCopieModeleTache (fTobModele, fTobDet);  // AB-Modeletache-copie la Tob
  fTobDet.PutEcran (Ecran,fPaDet);
  if (fTobModele <> nil) and (fTobModele.NomTable = 'AFMODELETACHE') then
    PutRegles; // AB-Modeletache-Reprend les règles du modèle
  SetControlText('ATA_DATEDEBPERIOD_', dateToStr(fDtFinAffaire));
  fGSTaches.Objects[0,fGSTaches.Row] := fTobDet;
  //fTobDet.PutLigneGrid(fGSTaches, fGSTaches.Row, false, false, fStLesCols); // AB-LigneAffaire
  // raffraichir les periodicites
  //ATA_PERIODICITEOnClick(self);

  //type tache
  SetControlEnabled('ATA_MODESAISIEPDC', true);
  THValComboBox(GetControl('ATA_MODESAISIEPDC')).value := 'QUA';
  RefreshModeSaisie;
  RefreshLignePiece; // AB-LigneAffaire-Champs non modifiables
End;

procedure TOF_AFTACHES.InitNewRegles;
begin
  // gestion des règles
  fRB_MODEGENE1.checked := true;
  fRB_QUOTIDIENNE.checked := true;
  fRB_MOISMETHODE1.checked := true;
  //TCheckBox(GetControl('ATA_GESTIONDATEFIN')).Checked := true;
  THValComboBox(GetControl('ATA_MOISSEMAINE1')).Value := '1';
  THValComboBox(GetControl('ATA_MOISJOURLIB')).Value := 'J1';
  THValComboBox(GetControl('ATA_MOISSEMAINE2')).Value := '1';
  SetControlText('ATA_MOISJOURFIXE', '1');
  SetControlText('ATA_MOISFIXE0', '1');
  SetControlText('ATA_MOISFIXE1', '1');
  SetControlText('ATA_MOISFIXE2', '1');
  TCheckBox(GetControl('ATA_JOUR1H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR2H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR3H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR4H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR5H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR6H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR7H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR1M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR2M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR3M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR4M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR5M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR6M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR7M')).checked := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : On passe l'element a inserer ou a supprimer
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHES.ValiderLesTaches;
begin

  // suppression des taches une par une
  if assigned(fTobSup) and (fTobSup.Detail.count = 1) then
    begin
      if ValideLesTaches(nil, fTobSup.detail[0], nil, fCurAffaire.StAffaire, fUpdateRes) then
        begin
          fTobSup.ClearDetail;
          fTobDet := nil;
        end;
    end
  // enregistrement de toutes les taches et de leurs ressources
  else
    if (fStatut <> taConsult) and ValideLesTaches(fTobDet, nil, fTobSupRes, fCurAffaire.StAffaire, fUpdateRes) then
      fTobDet.modifie := false;

  fUpdateRes := nil;           
  fTobSupRes.ClearDetail;  
end;                          

procedure TOF_AFTACHES.BPLANNINGVIEWEROnClick(Sender: TObject);
begin
//  SendMessage(ecran.activecontrol.Handle, WM_KEYDOWN, VK_RIGHT,  0);
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then setFocusControl('ATA_FONCTION');
  if Valider(false, false) then
    begin
      RefreshGrid;
      //AGLLanceFiche('AFF','AFPLANNINGVIEWER','', GetControlText('ATA_AFFAIRE'),'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';MONOFICHE')
      AFLanceFicheAFPlanningViewer(GetControlText('ATA_AFFAIRE'),'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';MONOFICHE');
    end;
end;

procedure TOF_AFTACHES.BPLANNINGRESOnClick(Sender: TObject);
begin
//  SendMessage(ecran.activecontrol.Handle, WM_KEYDOWN, VK_RIGHT,  0);
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then setFocusControl('ATA_FONCTION');
  if Valider(false, false) then
    begin
      RefreshGrid;
      //AGLLanceFiche('AFF','AFPLANNINGRES','', GetControlText('ATA_AFFAIRE'),'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';MONOFICHE')
      AFLanceFicheAFPlanningRes(GetControlText('ATA_AFFAIRE'),'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';MONOFICHE');
    end;
end;

procedure TOF_AFTACHES.BPLANDECHARGEOnClick(Sender: TObject);
begin
  if fBoAFPLANDECHARGE then SaisiePlanCharge else SaisieRAF;
end;

procedure TOF_AFTACHES.SaisiePlanCharge;
var
  vInCurIndiceTob : Integer;
begin
  vInCurIndiceTob := fGSTaches.row -1;
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then setFocusControl('ATA_FONCTION');
  if Valider(false, false) then
  begin
    RefreshGrid;
    if fTobDet.getvalue('ATA_MODESAISIEPDC') = 'QUA' then
    begin
      if AFLanceFicheAFPlanCharge(GetControlText('ATA_AFFAIRE'), 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') +
                                                                 ';LIBAFFAIRE:' + fStAfficheAffaire +
                                                                 ';TIERS:' + GetControlText('ATA_TIERS') +
                                                                 ';PLANCHARGEQTE;MONOFICHE') = 'TRUE' then
      //if AGLLanceFiche('AFF','AFPLANCHARGE','',GetControlText('ATA_AFFAIRE'),'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';PLANCHARGEQTE;MONOFICHE') = 'TRUE' then
        begin
          // recharger toutes les taches de l'affaire
          fStatut := taConsult;
          loadAffaire(vInCurIndiceTob);
        end;
    end
    else
      if AFLanceFicheAFPlanCharge(GetControlText('ATA_AFFAIRE'), 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') +
                                                                 ';LIBAFFAIRE:' + fStAfficheAffaire +
                                                                 ';TIERS:' + GetControlText('ATA_TIERS') +
                                                                 ';PLANCHARGEMNT;MONOFICHE') = 'TRUE' then
      //if AGLLanceFiche('AFF','AFPLANCHARGE','', GetControlText('ATA_AFFAIRE'),'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';PLANCHARGEMNT;MONOFICHE') = 'TRUE' then
      begin
        // recharger toutes les taches de l'affaire
        fStatut := taConsult;
        loadAffaire(vInCurIndiceTob);
      end;
    // ajouter les ressources supplémentaires 
    TobResToScreen;
  end;                
end;

procedure TOF_AFTACHES.SaisieRAF;
//var
//  vInCurIndiceTob : Integer;
begin

PGIBoxAF (TexteMsgTache[21],'');
{  vInCurIndiceTob := fGSTaches.row -1;
  if Valider(false, false) then
    begin
      RefreshGrid;
      if fTobDet.getvalue('ATA_MODESAISIEPDC') = 'QUA' then
        begin
          if AFLanceFicheAFPlanCharge(GetControlText('ATA_AFFAIRE'), 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';PLANCHARGEQTERAF;MONOFICHE') = 'TRUE' then
            begin
              // recharger toutes les taches de l'affaire
              fStatut := taConsult;
              loadAffaire(vInCurIndiceTob);
            end;
        end
      else
        if AFLanceFicheAFPlanCharge(GetControlText('ATA_AFFAIRE'), 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';PLANCHARGEMNTRAF;MONOFICHE') = 'TRUE' then
          begin
            // recharger toutes les taches de l'affaire
            fStatut := taConsult;
            loadAffaire(vInCurIndiceTob);
          end;
    end;
}
end;
  
procedure TOF_AFTACHES.LoadCompetences;
var
  vQR : TQuery;           
  vSt : String;

begin
  vSt := 'SELECT ACO_COMPETENCE, ACO_LIBELLE FROM COMPETENCE ';
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then
    begin
      fTobCompet := Tob.create('COMPETENCE', nil, -1);
      fTobCompet.LoadDetailDB('COMPETENCE','','',vQR,False,True);
    end 
  finally
    ferme(vQr);
  end;
end;

procedure TOF_AFTACHES.RefreshGridRes;
begin
  fGSRes.VidePile(False);
  initGridRes;
end;

procedure TOF_AFTACHES.UpdateGridRes(pBoModifNbRes : Boolean; pInNum : Integer);
Var NbRes     : integer;
Begin

  // si modif du nb de ressource
  if pBoModifNbRes then
    begin
      NbRes := strToint(GetControlText('ATA_NBRESSOURCE')) + pInNum;
      SetControlText('ATA_NBRESSOURCE', inttoStr(NbRes));
      fGSRes.RowCount := NbRes + 1;
    end
  else
    // fixe le nombre de ressource
    begin
      if pInNum = 0 then
         begin
           RefreshGridRes;
           SetControlText('ATA_NBRESSOURCE', intToStr(pInNum));
           fGSRes.RowCount := 2;
         end
      else
        begin
          SetControlText('ATA_NBRESSOURCE', intToStr(pInNum));
          fGSRes.RowCount := pInNum + 1;
        end;
    end;

  // positionnement dans la grille
  fGSRes.row := fGSRes.rowcount - 1;
  fGSRes.Enabled := (GetControlText('ATA_NBRESSOURCE') <> '0');
  fGSRes.ElipsisButton := fGSRes.Enabled;
  IF fGSRes.Enabled Then SetFocusControl('GSRES');
End;
                         
procedure TOF_AFTACHES.BPLANNINGOnClick(Sender: TObject);
var
  vInNum : Integer;
  vStAffaire : String;
begin
  if Valider(false, false) then
    begin
      RefreshGrid;
      vInNum := strtoint(copy(GetParamSoc('SO_AFPLANNINGDEF'), 2, 1));
      vStAffaire := ' APL_AFFAIRE = "' + fED_AFFAIRE.text + '"';
      ExecPlanning( inttostr(vInNum + 153200), dateToStr(now), vStAffaire, '', '-');
    end;                                                                      
end;
 
procedure TOF_AFTACHES.BRESSOURCEOnClick(Sender: TObject);
begin
  if fGSRes.Row <= fTobDet.Detail.Count then
    AglLanceFiche ('AFF','RESSOURCE','', fTobDet.Detail[fGSRes.Row - 1].GetValue('ATR_RESSOURCE'),'')
  else
    PGIBoxAF (TexteMsgTache[25],'');
end;
   
procedure TOF_AFTACHES.BCALCULMNTRAPOnClick(Sender: TObject);
begin
  if TestCoherence(true) then CalculMontantRecurrent;
end;

function TOF_AFTACHES.TestCoherence(pBoRAP : Boolean) : boolean;
var
  i         : Integer;
  vRdTotal  : Double;

begin
  vRdTotal := 0;
  result := true;
  if (fGSRes.RowCount > 2) or (fGSRes.CellValues[cInColRes,1] <> '') then
    for i := 1 to fGSRes.RowCount -1 do
    begin
      if fGSRes.CellValues[cInColQte, i] = '' then
        fGSRes.CellValues[cInColQte, i] := '0';
      vRdTotal := vRdTotal + strtofloat(fGSRes.CellValues[cInColQte, i]);
    end;                                        

  if (THValComboBox(GetControl('ATA_MODESAISIEPDC')).value = 'QUA') then
  begin
    if ((GetControlText('ATA_QTEINITIALE') = '') and (vRdTotal <> 0)) or
       (vRdTotal <> strtofloat(GetControlText('ATA_QTEINITIALE'))) then
    begin                                
      if pBoRAP then
      begin
        if PGIAskAF (TexteMsgTache[23],'') = MrYes then
          result := true
        else
          result := false;
      end
      // C.B 11/08/2003
      // message d'erreur que si gestion en plan de charge
      else if fBoAFPLANDECHARGE then
      begin
        PGIBoxAF (TexteMsgTache[24],'');
        result := false;
      end
      else
        result := True;
    end                 
  end
  else
  begin
    if vRdTotal <> strtofloat(GetControlText('ATA_INITPTPR')) then
    begin
      if pBoRAP then
      begin
        if PGIAskAF (TexteMsgTache[23],'') = MrYes then
          result := true
        else
          result := false;
      end
      // C.B 11/08/2003
      // message d'erreur que si gestion en plan de charge
      else if fBoAFPLANDECHARGE then
      begin
        PGIBoxAF (TexteMsgTache[24],'');
        result := false;
      end
      else
        result := True;
    end;
  end;
end;
 
procedure TOF_AFTACHES.BCALCULQTERAPOnClick(Sender: TObject);
begin
  if TestCoherence(true) then CalculQuantiteRecurrent;
end;

procedure TOF_AFTACHES.BCALCULRAPRESOnClick(Sender: TObject);
var
  i           : integer;
  vRdPlanifie : Double;
  vRdQte      : Double;
 
begin
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then setFocusControl('ATA_FONCTION');
  if valider(false, false) then
  begin
    RefreshGrid;
    if TestCoherence(true) then
    begin
      for i := 0 to fTobDet.Detail.Count - 1 do
        begin
            // C.B 04/11/2003
            vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire,
                                          fTobDet.Detail[i].GetValue('ATR_NUMEROTACHE'),
                                          'APL_QTEPLANIFIEE',  fTobDet.Detail[i].GetValue('ATR_RESSOURCE'), True);
            vRdQte := StrToFloat(fTobDet.Detail[i].GetValue('ATR_QTEINITIALE')) - vRdPlanifie;
            if vRdQte < 0 then vRdQte := 0;
            fGSRes.Cells[cInColQteRAP, i + 1] := floatToStr(vRdQte);
        end;
    end;                                                        
  end;
end;
 
procedure TOF_AFTACHES.PAGETACHEChange(Sender: TObject);
begin
  if (TPageControl(GetControl('PAGETACHE')).ActivePage = TTAbSheet(GetControl('TAB_RESSOURCE'))) and
     getParamSoc('SO_AFREALPLAN') and (not fBoAFPLANDECHARGE) then
    SetControlVisible('BCALCULRAPRES', true)
  else
    SetControlVisible('BCALCULRAPRES', false);
end;

Function TOF_AFTACHES.PlusTypeArticleTache : string;
begin
  result := 'AND (CO_CODE="PRE" OR CO_CODE="MAR" OR CO_CODE="FRA"';
  if (ctxTempo in V_PGI.PGIContexte) or (ctxGCAff in V_PGI.PGIContexte) then result:=result + '  OR CO_CODE="CTR"';
  result:=result + ')';
end;

function TOF_AFTACHES.ExistPlanning(pStAffaire, pStNumeroTache : String) : boolean;
begin
  result := ExisteSQL('SELECT APL_NUMEROTACHE from AFPLANNING where APL_NUMEROTACHE =' + pStNumeroTache + ' AND APL_AFFAIRE ="' + pStAffaire + '"');
end;

procedure TOF_AFTACHES.mnGeneAffaire_OnClick;
begin
  if RessourceSaisie then
  begin
    if Valider(false, false) then
    begin
      RefreshGrid;
      if ClientFerme then PGIBoxAF (TexteMsgTache[32],'');
      AFLanceFicheAFPlanningGenerer('', 'AFFAIRE:' + fED_AFFAIRE.text);
      RefreshLastDateGene;
    end;
  end
  else
  begin
    PGIBoxAF (TexteMsgTache[30], ecran.caption);
    TPAGECONTROL(GetControl('PAGETACHE')).ActivePage := TTabSheet(GetControl('TAB_RESSOURCE'));
  end;
end;

Function TOF_AFTACHES.RessourceSaisie : Boolean;
begin
  result := GetControlText('ATA_NBRESSOURCE') <> '0';
end;

procedure TOF_AFTACHES.mnGeneTache_OnClick;
begin
  if RessourceSaisie then
  begin
    if Valider(false, false) then
    begin
      RefreshGrid;                                         
      if ClientFerme then PGIBoxAF (TexteMsgTache[32],'');
      // controle de la coherence des quantites initiale
      if (GetControlText('ATA_QTEINITIALE') = '') or
         (StrToFloat(GetControlText('ATA_QTEINITIALE')) = 0) then
        AFLanceFicheAFPlanningGenerer('', 'AFFAIRE:' + fED_AFFAIRE.text + ';NUMEROTACHE:' + floatToStr(fTobDet.GetValue('ATA_NUMEROTACHE')) + ';LIBTACHE:' + fTobDet.GetValue('ATA_LIBELLETACHE1') + ';ARTICLE:' + getControlText('ATA_ARTICLE')  + ';TIERS:' + GetControlText('ATA_TIERS') + ';FCTTACHE:' + getControlText('ATA_FONCTION') + ';UNITE:' + GetControlText('ATA_UNITETEMPS') + ';LASTDATEGENE:' + GetControlText('ATA_LASTDATEGENE') + ';COMPTEUR:0' + ';ACTIVITEREPRIS:' + GetControlText('ATA_ACTIVITEREPRIS'))
      else if (TestCoherence(false) and (StrToFloat(GetControlText('ATA_QTEINITIALE')) <> 0))  then
        AFLanceFicheAFPlanningGenerer('', 'AFFAIRE:' + fED_AFFAIRE.text + ';NUMEROTACHE:' + floatToStr(fTobDet.GetValue('ATA_NUMEROTACHE')) + ';LIBTACHE:' + fTobDet.GetValue('ATA_LIBELLETACHE1') + ';ARTICLE:' + getControlText('ATA_ARTICLE')  + ';TIERS:' + GetControlText('ATA_TIERS') + ';FCTTACHE:' + getControlText('ATA_FONCTION') + ';UNITE:' + GetControlText('ATA_UNITETEMPS') + ';LASTDATEGENE:' + GetControlText('ATA_LASTDATEGENE') + ';COMPTEUR:' + GetControlText('ATA_QTEINITIALE') + ';ACTIVITEREPRIS:' + GetControlText('ATA_ACTIVITEREPRIS'));
      RefreshLastDateGene;
    end;
  end
  else
  begin
    PGIBoxAF (TexteMsgTache[30], ecran.caption);
    TPAGECONTROL(GetControl('PAGETACHE')).ActivePage := TTabSheet(GetControl('TAB_RESSOURCE'));
  end;
end;
 
procedure TOF_AFTACHES.RefreshLastDateGene;
var
  vSt : String;
  vQr : TQuery;

begin
  // raffraichir la date de derniere generation
  vSt := 'SELECT ATA_LASTDATEGENE FROM TACHE WHERE ATA_AFFAIRE = "' + fED_AFFAIRE.text + '" AND ATA_NUMEROTACHE = ' +  floatToStr(fTobDet.GetValue('ATA_NUMEROTACHE'));
  vQr := OpenSQL(vSt, TRUE);
  Try
    if Not vQr.EOF then
      SetControlText('ATA_LASTDATEGENE', vQr.FindField('ATA_LASTDATEGENE').AsString);
  Finally
    ferme(vQr);
  End;
End;

procedure TOF_AFTACHES.ATA_CODEARTICLEOnExit(SEnder: TObject);
var vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable  : String;
begin
  vStCodeArticle := getControlText('ATA_CODEARTICLE');
  if (vStCodeArticle <> fTobDet.GetValue('ATA_CODEARTICLE')) and
  controleCodeArticle (vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable) then
  begin
    SetControlText('ATA_TYPEARTICLE', vStTypeArticle);
    SetControlText('ATA_ARTICLE', vStArticle);
    SetControlText('ATA_ACTIVITEREPRIS', vStFacturable);
  end;
end;

{***********A.G.L.***************************************************
Auteur  ...... : AB
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : Tache créée à partir de modèle tache ou Ligne Affaire
               : copie de FromTOB AFMODELETACHE ou LIGNE à ToTOB TACHE
Mots clefs ... :
*********************************************************************}

procedure TOF_AFTACHES.TOBCopieModeleTache(FromTOB, ToTOB : TOB);
var i_pos,i_ind1: integer;
    FieldNameTo,FieldNameFrom,St:string;
    PrefixeTo,PrefixeFrom : string;
    vStCodeArticle,vStTypeArticle,vStArticle,vStFacturable : string;
begin
  PrefixeFrom := TableToPrefixe (FromTOB.NomTable);
  PrefixeTo := TableToPrefixe (ToTOB.NomTable);
  if PrefixeFrom = 'GL' then
  begin
    ToTOB.PutValue ('ATA_LIBELLETACHE1',FromTOB.GetValue('GL_LIBELLE'));
    ToTOB.PutValue ('ATA_CODEARTICLE'  ,FromTOB.GetValue('GL_CODEARTICLE'));
    ToTOB.PutValue ('ATA_QTEINITIALE'  ,FromTOB.GetValue('GL_QTEFACT'));
    ToTOB.PutValue ('ATA_IDENTLIGNE'   ,fStLignePiece);
    ToTOB.AddChampSupValeur ('IDENTLIGNE','X');
  end else
  for i_ind1 := 1 to FromTOB.NbChamps do
  begin
    FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
    St := FieldNameFrom ;
    i_pos := Pos ('_', St) ;
    System.Delete (St, 1, i_pos-1) ;
    FieldNameTo := PrefixeTo + St ;
    if ToTOB.FieldExists(FieldNameTo) then
      ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom))
  end;

  ToTOB.PutValue('ATA_TIERS', fCurAffaire.StTiers);
  ToTOB.PutValue('ATA_AFFAIRE',fCurAffaire.StAffaire);
  ToTOB.PutValue('ATA_AFFAIRE0',fCurAffaire.StAff0);
  ToTOB.PutValue('ATA_AFFAIRE1',fCurAffaire.StAff1);
  ToTOB.PutValue('ATA_AFFAIRE2',fCurAffaire.StAff2);
  ToTOB.PutValue('ATA_AFFAIRE3',fCurAffaire.StAff3);
  ToTOB.PutValue('ATA_AVENANT', fCurAffaire.StAvenant);
  ToTOB.PutValue('ATA_DEVISE', fCurAffaire.StDevise);

  if (ToTOB.getvalue('ATA_DATEDEBPERIOD') < fDtDebutAffaire)
  or (ToTOB.getvalue('ATA_DATEDEBPERIOD') > fDtFinAffaire)
  then
    ToTOB.putvalue('ATA_DATEDEBPERIOD', fDtDebutAffaire);
  if (ToTOB.getvalue('ATA_DATEFINPERIOD') > fDtFinAffaire)
  or (ToTOB.getvalue('ATA_DATEFINPERIOD') < fDtDebutAffaire)
  then
    ToTOB.putvalue('ATA_DATEFINPERIOD', fDtFinAffaire);

  vStCodeArticle := ToTOB.GetValue('ATA_CODEARTICLE');
  if controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable) then
  begin
    ToTOB.PutValue('ATA_TYPEARTICLE', vStTypeArticle);
    ToTOB.PutValue('ATA_ARTICLE', vStArticle);
    ToTOB.PutValue('ATA_ACTIVITEREPRIS', vStFacturable);
  end
end;
   
// champ non-modifiable pour tache créée à partir d'une ligne Affaire
procedure TOF_AFTACHES.RefreshLignePiece ;
begin
  if fTobDet.GetValue ('ATA_IDENTLIGNE') <> '' then
  begin
    SetControlEnabled('ATA_TYPEARTICLE', false);
    SetControlEnabled('ATA_CODEARTICLE', false);
    SetControlEnabled('ATA_QTEINITIALE', false);
    SetControlEnabled('ATA_UNITETEMPS', false);
    SetControlEnabled('BTLIGNE',(fAction <> taConsult));
  end
  else
  begin
    SetControlEnabled('BTLIGNE', False);
    SetControlEnabled('ATA_TYPEARTICLE', True);
    SetControlEnabled('ATA_CODEARTICLE', True);
    SetControlEnabled('ATA_QTEINITIALE', True);
    SetControlEnabled('ATA_UNITETEMPS', True);
  end;      
end;
    
// Initialisation tache créée à partir d'une ligne Affaire
procedure TOF_AFTACHES.TraiteLignePiece;
var CleLigne : R_CleDoc ;
    Q :TQuery;
    vTobDet : TOB;
begin
  if fStLignePiece = '' then exit;
  vTobDet := fTobTaches.FindFirst(['ATA_IDENTLIGNE'], [fStLignePiece], true);
  if (vTobDet <> nil) then
  begin
    fStOpenNumTache := vTobDet.GetValue('ATA_NUMEROTACHE');
  end else
  if fAction <> taConsult then
  begin
    fTobModele := Tob.create ('LIGNE', nil, -1);
    DecodeRefPiece(fStLignePiece,CleLigne);
    Q := OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleLigne,ttdLigne,true,true)+' ORDER BY GL_NUMLIGNE',True) ;
    fTobModele.SelectDB('',Q);
    Ferme(Q);
    fStatut := tacreat;
  end;
end;

// Consultation des lignes Affaire
procedure TOF_AFTACHES.BTLIGNEOnClick(SEnder: TObject);
var vStLigne: string;
    CleDoc : R_CleDoc;
begin
  vStLigne := fTobDet.Getvalue('ATA_IDENTLIGNE');
  if vStLigne <> '' then
  begin
    DecodeRefPiece (vStLigne,CleDoc);
    SaisiePiece(CleDoc,taConsult);
  end;
end;

Initialization
  registerclasses ( [TOF_AFTACHES] ) ;
End.
