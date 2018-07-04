
{***********UNITE*************************************************
Auteur  ...... : BELLARD Frédéric
Créé le ...... : 24/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCECRANA ()
Mots clefs ... : TOF;CPCECRANA
*****************************************************************}
Unit CPCECRANA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     Forms,
     UTof,
     UTob,
     Sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Menus,           // TMenuItem, PopUpMenu
     LookUp,          // LookUpList
     HQry,            // HQuery
     Htb97,           // TToolBarButton97, TToolWindow97
     HRichOle,        // THRichEditOle
     HSysMenu,
     Filtre,          // LoadFiltre
     Windows,         // VK_
     Grids,           // TGridDrawState
     PARAMSOC,        // GetParamSoc
     AGLInit,         // TheData
     Ent1,            // VH^, TFichierBase
{$IFDEF EAGLCLIENT}
     MainEAgl,        // AGLLanceFiche
{$ELSE}
     Fe_Main,         // AGLLanceFiche
     Db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
     uTofViergeMul,   // Fiche Ancetre Vierge MUL
     UtilSais,        // QuelZoomTableTNat
   {$IFDEF MODENT1}
   CPTypeCons,
   {$ENDIF MODENT1}
     uLibEcriture;    // TListJournal

Type
  TOF_CPCECRANA_VIE = Class (TOF_ViergeMul)

      Y_Section       : THEdit;
      Y_General       : THEdit;
      Y_Exercice      : THValComboBox;
      ComboAcces      : THValComboBox;
      Y_Etablissement : THMultiValComboBox;
      Y_DateComptable : THEdit;
      Y_DateComptable_: THEdit;
      ComboInfoSup    : THValComboBox;
      PageInfoSup     : TPageControl;
      BSUP            : TToolBarButton97;
      BSDOWN          : TToolBarButton97;
      BGUP            : TToolBarButton97;
      BGDOWN          : TToolBarButton97;
      LESOLDE         : THEdit;
      Y_NumeroPiece   : THEdit;
      Y_NumeroPiece_  : THEdit;
      Y_RefInterne    : THEdit;
      CBAvecCollectif : TCheckBox;
      Y_Valide        : TCheckBox;
      Y_Journal       : THMultiValComboBox;
      Y_NaturePiece   : THMultiValComboBox;
      Y_Devise        : THValComboBox;
      Y_QualifPiece    : THMultiValComboBox;
      ComboDefilCpt   : THValComboBox;
      ComboSoldePro   : THValComboBox;

      procedure OnClose; override ;
      procedure AfterShow ; override;
      procedure OnArgument (S : String ); override ;
      procedure OnLoad ; override ;
      procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState); override ;
      procedure OnClickBBlocNote ( Sender : TObject ); override;
      procedure OnRowEnterFListe (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean); override;
      procedure OnNew; override ;
      procedure OnDelete; override ;
      procedure OnUpdate; override ;
      procedure ApresChangementFiltre; override;
      procedure OnClickBValider( Sender : TObject ); override;
      procedure OnKeyDownFListe  ( Sender : TObject ; var Key: Word; Shift: TShiftState); override;
      procedure OnPopUpPopF11(Sender : TObject); override;

      procedure OnDblClickFListe (Sender: TObject);
      procedure OnChangeComboY_Exercice ( Sender : TObject );
      procedure OnClickBSUP( Sender : TObject );
      procedure OnClickBSDOWN( Sender : TObject );
      procedure OnClickBGUP( Sender : TObject );
      procedure OnClickBGDOWN( Sender : TObject );
      procedure OnExitY_DateComptable( Sender : TObject );
      procedure OnChangeComboInfoSup    ( Sender : TObject );
      procedure OnChangeComboAcces      ( Sender : TObject );

      procedure OnColumnWidthsChanged( Sender : TObject );

    private
      gLoad                : Boolean; // FQ 19513 
      FDroitEcritures      : Boolean;
      FOkCreateModif       : Boolean;
      FFromSaisie          : Boolean;
      FBoRenvoiDetailEcr   : Boolean; // Doit on renvoyer une TOb avec le contenu des écritures
      FStNomBase           : string;  // Nom de la base sur laquelle on fait la requête ( Bureau PGI )
      FOkLastRow           : Boolean; // Positionnement sur le dernier mouvement
      FTobAna              : TOB;           // Tob qui contient les infos du  ANALYTIQUE
      FTobGen              : TOB;           // Tob qui contient les infos du  GENERAL
      FBoConfidentielAna   : Boolean; // Gestion Confidentiel Analytique pour blocage des fonctions
      FBoConfidentielGen   : Boolean; // Gestion Confidentiel General pour blocage des fonctions
      FBoOkControleDate    : Boolean;
      SLibelle             : THLabel;
      GLibelle             : THLabel;
      FBoFaireRequete      : Boolean; //
      FTobEcrSelect        : TOB;           // Tob qui contient les infos des ECRITURES ANALYTIQUES
      // Variables contenant le numéro des colonnes imortantes de la grille
      FColSection, FColGeneral, FColSoldePro : Integer;
      FZListJournal        : TZListJournal; //

      SAxe : THValComboBox;
      FStArgumentTOF       : string;  // Paramètres passés à la TOF dans le OnArgument

      PopUpTraitement : TPopUpMenu;
      PopUpUtilitaire : TPopUpMenu;
      PopUpFctComp    : TPopUpMenu;
      PopUpEdition    : TPopUpMenu;

      FColDebitE         : integer;
      FColCreditE        : integer;

      procedure AfficheInfoSup;
      procedure VerifieLesDates;
      function ChargeAnalytique(Force : boolean): Boolean;
      procedure OnExitG_General( Sender : TObject );
      function ChargeGeneral(Force : boolean): Boolean;
      function FaireStWhere( vBoForceDetailAno : boolean = false ) : string ; // Balance Général par Général
      function RecupEcrSelect( TheRow : integer = 0 ) : Boolean;
      procedure IndiceColFListe;
      procedure AnaSuivant(vBoSuiv : Boolean; vBoREsteSurZone : Boolean);
      procedure GenSuivant (vBoSuiv : Boolean; vBoREsteSurZone : Boolean);
      function RecupWhereAGL : string ;
      function RecupWhereEcriture   ( vBoForceDetailAno : boolean = false ) : string ;
      function RecupCriteresCommuns : string ;
      procedure InitColFListe;
      procedure CumulsANAClick(Sender: TObject);
      function PositionneExo        : TExoDate;
      procedure RechercheEcrituresClick ( Sender : Tobject ); // Recherche d'ecritures

      function Axe : TFichierBase;
      function ChoixAxe(index : integer) : string;

      procedure OnElipsisClickS_Section( Sender : TObject );
      procedure OnElipsisClickY_General(Sender: TObject);

      procedure OnExitS_Section( Sender : TObject );

      procedure MiseAJourCaptionEcran;
      procedure GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

      // Initialisation des PopUp
      procedure InitPopUP(vActivation: Boolean; vNumPopUp : integer =  0);

      // POPUPOUTIL
      procedure OnPopUpPopUpTraitement(Sender: TObject);
      procedure OnPopUpPopUpUtilitaire(Sender: TObject);
      procedure OnPopUpPopUpFctComp(Sender: TObject);
      procedure OnPopUpPopUpEdition(Sender: TObject);

      // Traitements Comptables
      procedure SaisieAnaClick(Sender : TObject); // Saisie analytique
      procedure ReVentAnaClick(Sender : TObject); // Re-ventilations analytique
      procedure ReImpAnaClick(Sender: TObject); // Re-imputations analytiques

      // Utilitaires
      procedure OnClickCompteSection ( Sender : TObject ); // Paramètres du compte Section

      // Fonctions complementaire
      procedure OnClickCommentSect(Sender : TObject);

      procedure OnClickCubDecAna(Sender : TObject);
      procedure OnClickAnaStaAna(Sender : TObject);
      procedure OnClickAnaStaAnaGen(Sender : TObject);


      Procedure PrepareArgumentGLG ( vBoJustifSolde : Boolean );
      procedure OnClickGLAna(Sender: TObject);
      procedure LancementGLAnaGen( vFichierBase : TFichierBase );
      procedure OnClickGLAnaGen(Sender: TObject);
      procedure LancementGLGenAna( vFichierBase : TFichierBase );
      procedure OnClickGLGenAna(Sender: TObject);
      procedure PrepareArgumentBAL;
      procedure OnClickBalAna(Sender: TObject);
      procedure LancementBalAnaGen( vFichierBase : TFichierBase );
      procedure OnClickBalAnaGen(Sender: TObject);
      procedure LancementBalGenAna( vFichierBase : TFichierBase );
      procedure OnClickBalGenAna(Sender: TObject);
      function AxeAna(Alias : boolean) : string;
    protected
      procedure InitControl; override; // Init des composants de la fiche
      procedure RefreshFListe( vBoFetch : Boolean ) ; override;
      function  BeforeLoad : Boolean ; override;
      procedure RemplitATobFListe ; override;
      function  AjouteATobFListe( vTob : Tob ) : Boolean ; override;
      procedure CalculPourAffichage( vTob : Tob );   override;
      function  AfterLoad : boolean;                 override;
    public
  end ;

function CPLanceFiche_ConsEcrAna( vStParam : string = '' ) : string ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
    // *** Fonctionnement de l'unité
    uLibExercice,          // CInitComboExercice
    SaisBor,          // LanceSaisieFolio
    Saisie,           // TrouveSaisie
   // MulAna,     // MultiCritereAna
    CPMULANA_TOF,
    uLibWindows,      // AfficheDBCR  (testjoker)
    SaisComm,         // (GetO)
    SaiSUtil,         // pour le StrSO (RMVT)

     // *** Traitements comptables ***
    SaisODA,
    Reventilana,
    CPREIMPUTANA_TOF,
     // *** Fonctions complémentaires ***
    Cummens,          // CumulSectMensuel
     // *** Editions ***
    uTofCPGLAna,      // CPLanceFiche_CPGL
    CPBALANAL_TOF,    // CPLanceFiche_BalanceAnalytique
    CPBALANALGEN_TOF, // CPLanceFiche_BalanceAnalGen
    CPBALGENANAL_TOF, // CPLanceFiche_BalanceGenAnal
     // *** Utilitaires ***
    CPSECTION_TOM,    // fiche section
    CPGENERAUX_TOM;

var
  GTobDetailEcr : Tob;

const
  cFI_TABLE = 'CPCECRANA';

  cSelect = ', Y_EXERCICE CLE_EXERCICE, Y_NUMLIGNE CLE_NUMLIGNE, Y_NUMVENTIL, ' + //fb 14/06/2006 FQ 17683
            ' Y_QUALIFPIECE CLE_QUALIFPIECE '; // Champ de la clé obligatoire

function CPLanceFiche_ConsEcrAna( vStParam : string = '' ) : string ;
begin
{$IFDEF CCSTD}
  Result := AGLLanceFiche('CP', 'CPCECRANA_VIE', '', '', vStParam);
{$ELSE}
  if (VH^.CpExoRef.Code = VH^.Encours.Code) or (VH^.CpExoRef.Code = VH^.Suivant.Code) then begin
    Result := AGLLanceFiche('CP', 'CPCECRANA_VIE', '', '', vStParam);
   end
  else begin
    PgiInfo('Pour pouvoir utiliser cette fonction, l''exercice de référence doit ' +
            'être l''exercice en cours ou le suivant.', 'Consultation des écritures analytiques');
    end;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPCECRANA_VIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPCECRANA_VIE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPCECRANA_VIE.OnClose ;
begin
  FreeAndNil(FTobEcrSelect);
  FreeAndNil(FTobAna);
  FreeAndNil(FTobGen);
  FreeAndNil(FZListJournal);
  inherited;
end ;

procedure TOF_CPCECRANA_VIE.AfterShow;
begin
  inherited;
  if (Trim(Y_Section.Text) = '') and (Trim(Y_General.Text) = '') then begin
    if Y_Section.CanFocus then
      Y_Section.SetFocus;
    end;
  // FQ 19513
  gLoad := false;
end;

procedure TOF_CPCECRANA_VIE.OnLoad;
var
  lStArg         : string;
  lSt            : string;
  lStChampsTries : string;
begin
  // Init des variables globales
  FDroitEcritures := ExJaiLeDroitConcept(TConcept(ccSaisEcritures), False);
  MakeZoomOLE(ECRAN.Handle); // Appel depuis Excel

  FOkCreateModif := False;

  // Chargement des paramètres du OnArgument
  if FStArgumentTOF <> '' then begin
    lStArg := FStArgumentTOF;

    // Section analytique
    {$IFDEF BUREAU}
    Y_Section.Enabled := False;
    {$ENDIF}
    lSt := ReadTokenSt(lStArg);
    if lSt <> '' then begin
      AFiltreDisabled := True;
      Y_Section.Text  := lSt;
      if Y_Section.CanFocus then
        Y_Section.SetFocus;
      end;

    // Axe
    {$IFDEF BUREAU}
    SAxe.Enabled := False;
    {$ENDIF}
    lSt := ReadTokenSt(lStArg);
    if lSt <> '' then
    begin
      AFiltreDisabled := True;
      // GCO - 30/10/2007 - FQ 21724
      SAxe.Value  := lSt;
      if SAxe.CanFocus then
        SAxe.SetFocus;
    end;

    // Exercice
    lSt := ReadTokenSt(lStArg);
    if lSt <> '' then
    begin
      if lSt = '-2' then
      begin
        Y_Exercice.ItemIndex := 0;
       {$IFDEF BUREAU}
        Y_Exercice.Enabled := False;
        FBoBloqueExercice  := True;
       {$ENDIF}
       end
      else
      begin
        // lSt vaut 0,1 ou -1 si on vient de Excel
        if lSt = '0' then
          lSt := VH^.Encours.Code
        else
          if lSt = '1' then
           lSt := VH^.Suivant.Code
          else
            if lSt = '-1' then
              lSt := VH^.Precedent.Code;
        Y_Exercice.Value := CExerciceVersRelatif( lSt );
      end;
    end;

    // Champs de Tri
    lSt := ReadTokenSt(lStArg);
    if lSt <> '' then
      lStChampsTries := lSt;

    // Général
    {$IFDEF BUREAU}
    Y_General.Enabled := False;
    {$ENDIF}
    lSt := ReadTokenSt(lStArg);
    if lSt <> '' then begin
      AFiltreDisabled   := True;
      Y_General.Text := lSt;
      if Y_General.CanFocus then
        Y_General.SetFocus;
      end;

    // Appel depuis la saisie
    if ReadTokenSt(lStArg) = '-' then
      FFromSaisie := False
    else
      FFromSaisie := True;

    // Renvoi des écritures dans GTobDetailECr
    FBoRenvoiDetailEcr := (ReadTokenSt(lStArg) = 'X');

    // Nom de la base sur laquelle on doit faire la requête SELECT
    FStNomBase := ReadTokenSt(lStArg);
    if FStNomBase <> '' then
      FStNomBase := FStNomBase + '.dbo.';

    end;

  //
  if FBoRenvoiDetailEcr then begin
    SetControlProperty('BVALIDER', 'MODALRESULT', '1');
    // Réaffectation du BValider car on doit renvoyer la liste des écritures
    BValider.OnClick := OnClickBValider;
   end
  else
    SetControlProperty('BVALIDER', 'MODALRESULT', '0');

  // On force la consultation si on vient de la saisie Borderau ou si
  // présence d'un verrou de la saisie dans la table COURRIER
  if (FFromSaisie) or ExisteSql('SELECT MG_EXPEDITEUR FROM ' + FStNomBase + 'COURRIER WHERE MG_TYPE = 1000' +
                               ' AND ((MG_COMBO = "NRSAISIECREAT") OR (MG_COMBO = "NRSAISIEMODIF"))' +
                               ' AND MG_EXPEDITEUR = "' + V_Pgi.User + '" ORDER BY MG_EXPEDITEUR') then begin
    ComboAcces.Value   := '0';
    ComboAcces.Enabled := False;
    AFiltreDisabled    := True;
    FOkCreateModif     := False;
    end;

  InitPopUp(False);

  // Mise à blanc des composants Z_
  AfficheInfoSup;

  InitAutoSearch;

  inherited; // InitControl; Chargement du filtre 'DEFAUT';

  PositionneEtabUser(Y_Etablissement);
end ;

procedure TOF_CPCECRANA_VIE.OnElipsisClickS_Section(Sender: TObject);
begin
  LookUpList( THEdit(Sender),
              'Sections analytiques',
              'SECTION',
              'S_SECTION' ,
              'S_LIBELLE',
              'S_AXE="A' + IntToStr(SAxe.itemindex+1) + '" AND ' + CGenereSQLConfidentiel('S'),
              'S_SECTION' ,
              True,
              3 );
end;

procedure TOF_CPCECRANA_VIE.OnElipsisClickY_General(Sender: TObject);
begin
  LookUpList( THEdit(Sender),
              'Comptes généraux',
              'GENERAUX',
              'G_GENERAL' ,
              'G_LIBELLE',
              CGenereSQLConfidentiel('G'),
              'G_GENERAL' ,
              True,
              1 );
end;

function TOF_CPCECRANA_VIE.Axe : TFichierBase;
begin
  Result := fbAxe1;
  Case SAxe.ItemIndex of
    0 : Result := fbAxe1;
    1 : Result := fbAxe2;
    2 : Result := fbAxe3;
    3 : Result := fbAxe4;
    4 : Result := fbAxe5;
    end;
end;

procedure TOF_CPCECRANA_VIE.OnExitS_Section(Sender: TObject);
begin
  ChargeAnalytique(false);
end;

procedure TOF_CPCECRANA_VIE.OnExitG_General(Sender: TObject);
begin
  ChargeGeneral(False);
end;

procedure TOF_CPCECRANA_VIE.SaisieAnaClick(Sender: TObject);
{$IFDEF COMPTA}
var
  M : RMVT ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  FillChar(M,Sizeof(M),#0) ;
  M.Simul := 'N' ;
  M.CodeD := V_PGI.DevisePivot ;
  M.DateC := V_PGI.DateEntree ;
  M.TauxD := 1 ;
  M.DateTaux := M.DateC ;
  M.Valide := False ;
  M.Etabl := VH^.ETABLISDEFAUT ;

  LanceSaisieODA(Nil,taCreat,M) ;
  BCherche.Click;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.ReVentAnaClick(Sender: TObject);
begin
{$IFDEF COMPTA}
{b fb 16/06/2006 FQ17912}
  AGLLanceFiche('CP', 'CPREVENTILANA', '', '', FTobEcrSelect.GetValue('Y_JOURNAL')+';'+
                                       FTobEcrSelect.GetValue('Y_EXERCICE')+';'+
                                       FTobEcrSelect.GetValue('Y_QUALIFPIECE')+';'+
                                       FTobEcrSelect.GetValue('Y_GENERAL')+';'+
                                       Y_DateComptable.Text+';'+Y_DateComptable_.Text+';'+
                                       SAxe.Text
                                       +';'+IntToStr(FTobEcrSelect.GetValue('Y_NUMEROPIECE'))
                                       +';'+IntToStr(FTobEcrSelect.GetValue('Y_NUMLIGNE'))
                                       );
{e fb 16/06/2006 FQ17912}
  BCherche.Click;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.ReImpAnaClick(Sender: TObject);
begin
{$IFDEF COMPTA}
  AGLLanceFiche('CP', 'CPREIMPUTANA', '', '',SAxe.Text+';'+
                                       FTobEcrSelect.GetValue('Y_GENERAL')+';'+
                                       FTobEcrSelect.GetValue('Y_EXERCICE')+';'+
                                       FTobEcrSelect.GetValue('Y_SECTION')+';'+
                                       FTobEcrSelect.GetValue('Y_JOURNAL')+';'+
                                       FTobEcrSelect.GetValue('Y_DEVISE')+';'+
                                       FTobEcrSelect.GetValue('Y_QUALIFPIECE')+';'+
                                       IntToStr(FTobEcrSelect.GetValue('Y_NUMEROPIECE')));

{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.OnClickCommentSect(Sender: TObject);
var
  lAction: TActionFiche;
begin
  if Y_Section.Text = '' then
    Exit;

  inherited;
  lAction := taModif;
  if ((not ExJaiLeDroitConcept(TConcept(ccSecModif), False))
  or (not FOkCreateModif)) then
    lAction := taConsult;

  FicheSection( nil, ChoixAxe(SAxe.ItemIndex), Y_Section.Text, lAction, 2);
end;

procedure TOF_CPCECRANA_VIE.OnClickCubDecAna(Sender: Tobject);
begin
  AGLLanceFiche('CP','CPANALYTIQ_CUBE','','','') ;
end;

procedure TOF_CPCECRANA_VIE.OnClickAnaStaAna(Sender: Tobject);
begin
  AGLLanceFiche('CP','CPANALYTIQ_TOBV','','','') ;
end;

procedure TOF_CPCECRANA_VIE.OnClickAnaStaAnaGen(Sender: Tobject);
begin
  AGLLanceFiche('CP','CPANAGENE_TOBV','','','') ;
end;

Procedure TOF_CPCECRANA_VIE.PrepareArgumentGLG ( vBoJustifSolde : Boolean );
begin
{$IFDEF COMPTA}
  // Exercice
  ACritEdt.CritEdt.Exo.Code := CRelatifVersExercice(Y_Exercice.Value);

  // Début de la Date Comptable (1)
  ACritEdt.CritEdt.Date1 := StrToDate(Y_DateComptable.Text);

  // Fin de la Datecomptable (2)
  ACritEdt.CritEdt.Date2 := StrToDate(Y_DateComptable_.Text);

  // Numéro de Pièce (1)
  if Y_NumeroPiece.Text <> '' then
    ACritEdt.CritEdt.GL.NumPiece1 := StrToInt(Y_NumeroPiece.Text);

  // Numéro de Pièce (2)
  if Y_NumeroPiece_.Text <> '' then
    ACritEdt.CritEdt.GL.NumPiece2 := StrToInt(Y_NumeroPiece_.Text);

  // Référence Interne
  ACritEdt.CritEdt.RefInterne := Y_REFINTERNE.Text;

  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := Y_QualifPiece.Text;

  // Etablissement
  ACritEdt.CritEdt.Etab := Y_ETABLISSEMENT.Value;

  // Devise
  ACritEdt.CritEdt.DeviseSelect := Y_DEVISE.Value;

  // Ecritures Valides
  if Y_Valide.State = cbGrayed then
    ACritEdt.CritEdt.Valide := ''
  else
    ACritEdt.CritEdt.Valide := IIF( Y_Valide.Checked, 'OUI', 'NON' );
{$ENDIF}


  // Ecritures Valides
  ACritEdt.CritEdt.Valide := '';
end;

procedure TOF_CPCECRANA_VIE.OnClickGLAna(Sender: TObject);
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1   := Y_SECTION.Text;
  ACritEdt.CritEdt.Cpt2   := Y_SECTION.Text;

  TheData := ACritEdt;
  CPLanceFiche_CPGLANA;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.LancementGLAnaGen( vFichierBase : TFichierBase );
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1    := Y_Section.Text;
  ACritEdt.CritEdt.Cpt2    :=  Y_Section.Text;
  ACritEdt.CritEdt.sCpt1    := Y_General.Text;
  ACritEdt.CritEdt.sCpt2    := Y_General.Text;
  ACritEdt.CritEdt.GL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_CPGLANAPARGENE;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.OnClickGLAnaGen(Sender: TObject);
begin
  LancementGLAnaGen(Axe);
end;

procedure TOF_CPCECRANA_VIE.LancementGLGenAna( vFichierBase : TFichierBase );
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1    := Y_General.Text;
  ACritEdt.CritEdt.Cpt2    := Y_General.Text;
  ACritEdt.CritEdt.sCpt1    := Y_Section.Text;
  ACritEdt.CritEdt.sCpt2    := Y_Section.Text;
  ACritEdt.CritEdt.GL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_CPGLGENEPARANA;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.OnClickGLGenAna(Sender: TObject);
begin
  LancementGLGenAna(Axe);
end;

procedure TOF_CPCECRANA_VIE.PrepareArgumentBAL;
begin
{$IFDEF COMPTA}
  // Exercice
  ACritEdt.CritEdt.Exo.Code := CRelatifVersExercice(Y_Exercice.Value);
  // Début de la Date Comptable (1)
  ACritEdt.CritEdt.Date1 := StrToDate(Y_DateComptable.Text);
  // Fin de la Datecomptable (2)
  ACritEdt.CritEdt.Date2 := StrToDate(Y_DateComptable_.Text);
  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := Y_QualifPiece.Text;
  // Etablissement
  ACritEdt.CritEdt.Etab := Y_Etablissement.Value;
  // Devise
  ACritEdt.CritEdt.DeviseSelect := Y_Devise.Value;
{$ENDIF}

  // Options
  ACritEdt.CritEdt.Bal.SansAnoANA := true;
  ACritEdt.CritEdt.AvecComparatif := false;
end;

procedure TOF_CPCECRANA_VIE.OnClickBalAna(Sender: TObject);
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := Y_Section.Text;
  ACritEdt.CritEdt.Cpt2 := Y_Section.Text;
  ACritEdt.CritEdt.BAL.Axe := fbToAxe(Axe);

  TheData := ACritEdt;
  CPLanceFiche_BalanceAnalytique;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.LancementBalAnaGen( vFichierBase : TFichierBase );
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1    := Y_Section.Text;
  ACritEdt.CritEdt.Cpt2    := Y_Section.Text;
  ACritEdt.CritEdt.sCpt1    := Y_General.Text;
  ACritEdt.CritEdt.sCpt2    := Y_General.Text;
  ACritEdt.CritEdt.BAL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_BalanceAnalGen;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.OnClickBalAnaGen( Sender : TObject );
begin
  LancementBalAnaGen(Axe);
end;

procedure TOF_CPCECRANA_VIE.LancementBalGenAna( vFichierBase : TFichierBase );
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1    := Y_General.Text;
  ACritEdt.CritEdt.Cpt2    := Y_General.Text;
  ACritEdt.CritEdt.sCpt1    := Y_Section.Text;
  ACritEdt.CritEdt.sCpt2    := Y_Section.Text;
  ACritEdt.CritEdt.BAL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_BalanceGenAnal;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.OnClickBalGenAna( Sender : TObject );
begin
  LancementBalGenAna(Axe);
end;

procedure TOF_CPCECRANA_VIE.OnPopUpPopUpTraitement(Sender: TObject);
begin
  InitPopUp(True, 2);
end;

procedure TOF_CPCECRANA_VIE.OnPopUpPopUpUtilitaire(Sender: TObject);
begin
  InitPopUp(True, 3);
end;

procedure TOF_CPCECRANA_VIE.OnPopUpPopUpFctComp(Sender: TObject);
begin
  InitPopUp(True, 4);
end;

procedure TOF_CPCECRANA_VIE.OnPopUpPopUpEdition(Sender: TObject);
begin
  InitPopUp(True, 5);
end;

procedure TOF_CPCECRANA_VIE.OnPopUpPopF11(Sender: TObject);
begin
{$IFDEF COMPTA}
  InitPopUp(True, 0); // 0 pour tous les PopUpMenu Menu
  inherited; // Ajoute les élements visibles des menus de Ecran dans PopF11
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.InitPopUP(vActivation: Boolean; vNumPopUp : integer =  0);
var
  i : integer;
begin
  RecupEcrSelect;

  FZListJournal.Load([FTobEcrSelect.GetValue('Y_JOURNAL')]);

  // POPUPOUTIL INACTIF

  // POPUPTRAITEMENT
  for i := 0 to PopUpTraitement.Items.Count -1 do begin
    if PopUpTraitement.Items[i].Name = 'SAISANA' then begin
      if vActivation then
        PopUpTraitement.Items[i].Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
      else
        PopUpTraitement.Items[i].OnClick := SaisieAnaClick;
      Continue;
      end;

    if PopUpTraitement.Items[i].Name = 'REVENTANA' then begin
      if vActivation then
        PopUpTraitement.Items[i].Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
      else
        PopUpTraitement.Items[i].OnClick := ReVentAnaClick;
      Continue;
      end;

    if PopUpTraitement.Items[i].Name = 'REIMPANA' then begin
      if vActivation then
        PopUpTraitement.Items[i].Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
      else
        PopUpTraitement.Items[i].OnClick := ReImpAnaClick;
      Continue;
      end;
    end;

  // POPUP UTILITAIRE
  for i := 0 to PopUpUtilitaire.Items.Count -1 do begin
    if PopUpUtilitaire.Items[i].Name = 'PARSECANA' then begin
      if vActivation then
        PopUpUtilitaire.Items[i].Enabled := (Y_Section.Text <> '') and (not FBoConfidentielAna)
      else
        PopUpUtilitaire.Items[i].OnClick := OnClickCompteSection;
      Continue;
      end;
    end;

    // POPUP FONCTIONCOMPLEMENTAIRE
  for i := 0 to PopUpFctComp.Items.Count -1 do begin
    if PopUpFctComp.Items[i].Name = 'RECCOMPANA' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (FTobEcrSelect<>nil)
      else
        PopUpFctComp.Items[i].OnClick := RechercheEcrituresClick;
      Continue;
      end;

    if PopUpFctComp.Items[i].Name = 'COMMENTSECT' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (FTobEcrSelect<>nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickCommentSect;
      Continue;
      end;

    if PopUpFctComp.Items[i].Name = 'CUMSECANA' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (FTobEcrSelect<>nil)
      else
        PopUpFctComp.Items[i].OnClick := CumulsAnaClick;
      Continue;
      end;

    if PopUpFctComp.Items[i].Name = 'CUBDECANA' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (FTobEcrSelect<>nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickCubDecAna;
      Continue;
      end;

    if PopUpFctComp.Items[i].Name = 'ANASTAANA' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (FTobEcrSelect<>nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickAnaStaAna;
      Continue;
      end;

    if PopUpFctComp.Items[i].Name = 'ANASTAANAGEN' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (FTobEcrSelect<>nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickAnaStaAnaGen;
      Continue;
      end;
    end;

  // POPUP EDITION
  for i := 0 to PopUpEdition.Items.Count - 1 do begin
    // Grand-livre analytique
    if PopUpEdition.Items[i].Name = 'GLANA' then begin
      //if not vActivation then
      PopUpEdition.Items[i].OnClick := OnClickGLAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Grand-livre analytique par général
    if PopUpEdition.Items[i].Name = 'GLANAGEN' then begin
      PopUpEdition.Items[i].OnClick := OnClickGLAnaGen;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Grand-livre général par analytique
    if PopUpEdition.Items[i].Name = 'GLGENANA' then begin
      PopUpEdition.Items[i].OnClick := OnClickGLGenAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Balance analytique
    if PopUpEdition.Items[i].Name = 'BALANA' then begin
      PopUpEdition.Items[i].OnClick := OnClickBalAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Balance analytique par général sur l'axe sélectionné
    if PopUpEdition.Items[i].Name = 'BALANAGEN' then begin
      PopUpEdition.Items[i].OnClick := OnClickBalAnaGen;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Balance générale par analytique sur l'Axe sélectionné
    if PopUpEdition.Items[i].Name = 'BALGENANA' then begin
      PopUpEdition.Items[i].OnClick := OnClickBalGenAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;
    end;
end;

procedure TOF_CPCECRANA_VIE.OnClickBBlocNote(Sender: TObject);
var
  lQEcr : TQuery;
begin
  inherited;
  HPB.Visible := BBlocNote.Down;
  if BBlocNote.Down then begin
    FBLocNote.Clear;
    if (not RecupEcrSelect) or (FTobEcrSelect=nil) then
      Exit;
    lQEcr := nil;
    try
      lQEcr := OpenSql('SELECT Y_BLOCNOTE FROM ' + FStNomBase + 'ANALYTIQ WHERE ' +
        '(Y_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetValue('Y_DATECOMPTABLE')) + '") AND ' +
        '(Y_NUMEROPIECE = '    + IntToStr(FTobEcrSelect.GetValue('Y_NUMEROPIECE')) + ') AND ' +
        '(Y_NUMLIGNE = '       + IntToStr(FTobEcrSelect.GetValue('Y_NUMLIGNE')) + ') AND ' +
        '(Y_NUMVENTIL = '       + IntToStr(FTobEcrSelect.GetValue('Y_NUMVENTIL')) + ') AND ' + //fb 14/06/2006 FQ 17683
        '(Y_EXERCICE = "'      + FTobEcrSelect.GetValue('Y_EXERCICE') + '") AND ' +
        '(Y_QUALIFPIECE = "'   + FTobEcrSelect.GetValue('Y_QUALIFPIECE') + '") AND ' +
        '(Y_JOURNAL = "'       + FTobEcrSelect.GetValue('Y_JOURNAL') + '")', True);

      StringToRich(FBlocNote, lQEcr.FindField('Y_BLOCNOTE').AsString);
    finally
      Ferme(lQEcr);
      end;
    end;
end;

procedure TOF_CPCECRANA_VIE.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
  AfficheInfoSup;
  OnClickBBlocNote(nil);
end;

procedure TOF_CPCECRANA_VIE.MiseAJourCaptionEcran;
begin
  Ecran.Caption := TraduireMemoire('Consultation des écritures analytiques') + ' : ';

  if (Y_SECTION.Text <> '') and (Y_GENERAL.Text <> '') then begin
    Ecran.Caption := Ecran.Caption + Y_SECTION.Text + '  ' + Y_SECTION.Text + '  ' +
                     GLIBELLE.Caption + '   ' + LESOLDE.Text + '  du ' +
                     Y_DateComptable.Text + ' au ' + Y_DateComptable_.Text;
   end
  else begin
    if Y_GENERAL.Text <> '' then
      Ecran.Caption := Ecran.Caption + Y_GENERAL.Text + '  ' + GLIBELLE.Caption + '   ' +
                       LESOLDE.Text + '  du ' + Y_DateComptable.Text + ' au ' +
                       Y_DateComptable_.Text
    else
      if Y_SECTION.Text <> '' then
        Ecran.Caption := Ecran.Caption + Y_SECTION.Text + '  ' + SLIBELLE.Caption + '   ' +
                         LESOLDE.Text + '  du ' + Y_DateComptable.Text + ' au ' +
                         Y_DateComptable_.Text;
    end;
  UpDateCaption(Ecran);
end;

procedure TOF_CPCECRANA_VIE.RefreshFListe;
{$IFDEF EAGLCLIENT}
var
  lQuery : TQuery;
{$ENDIF}
begin
  // On force le tri sur un champ supplémentaire afin que le ATobFListe.Detail.Sort
  // renvoi toujours ATobFListe triée dans le même ordre
  AStTriTobFListe := AStTriTobFListe + ' ;CLE_NUMLIGNE';

  inherited; // Appel du inherited de la fiche Ancêtre pour trier ATobFListe et
             // l'afficher dans la grille

  if not vBoFetch then begin
    LeSolde.Text := '0,00';
    {$IFDEF EAGLCLIENT}
    if not vBoFetch then begin
      if AStSqlWhereTobFListe <> '' then begin
        lQuery := OpenSql('SELECT SUM(Y_DEBIT)-SUM(Y_CREDIT) TOTAL FROM ' + FStNomBase + 'ANALYTIQ ' + AStSqlWhereTobFListe , True);
        LeSolde.Text := AfficheDBCR( lQuery.FindField('TOTAL').AsFloat );
        Ferme( lQuery );
        end;
      end;
    {$ELSE}
    // Mise à jour des informations du compte affiché
    if (ATobFListe <> nil) and (ATobFListe.Detail.Count <> 0) then
      LeSolde.Text := AfficheDBCR( ATobFListe.Detail[ATobFListe.Detail.Count-1].GetValue('SOLDEE'));
    {$ENDIF}

    // Mise à Jour du Caption de l' écran
    MiseAJourCaptionEcran;
    end;
end;

procedure TOF_CPCECRANA_VIE.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then
    Exit;

  if (ACol = FColSoldePro) and (FListe.Row <> ARow) then begin
    if Trim(Y_General.Text) <> '' then begin
      if FTobGen.GetValue('G_NATUREGENE') = 'TID' then
        Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed, ClGreen)
      else
        if FTobGen.GetValue('G_NATUREGENE') = 'TIC' then
          Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClGreen, ClRed);
     end
    else begin
      if Trim(Y_Section.Text) <> '' then begin
        if FTobAna.GetValue('S_SENS') = 'D' then
          Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed, ClGreen)
        else
          if FTobAna.GetValue('S_SENS') = 'C' then
            Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClGreen, ClRed);
        end;
      end;
    end;
end;

function TOF_CPCECRANA_VIE.AxeAna(Alias : Boolean) : string;
var
  Sql, Sql2 : string;
begin
  Sql:='';
  Sql2:='';
  if VH^.AnaCroisaxe then begin
    if GetParamSocSecur('SO_VENTILA'+IntToStr(SAxe.ItemIndex+1), false) then begin
      Sql := 'Y_SOUSPLAN'+IntToStr(SAxe.ItemIndex+1);
      if Alias then
        Sql2 := ' AS Y_SECTION';
      end;
   end
  else
    Sql := 'Y_SECTION';

  Result := Sql + Sql2;
end;

procedure TOF_CPCECRANA_VIE.RemplitATobFListe;
var
  lStWhere : string ;
begin
  // Traduction des champs dans la LISTE PARAMETRABLE
  AStSqlTobFListe := '' ;

  if FBoFaireRequete then begin
    lStWhere  :=  FaireStWhere ;
    if lStWhere <> '' then
       AStSqlTobFListe := 'SELECT ' + CSqlTextFromList(FStListeChamps) + ' ' +
                          cSelect + ' FROM ' + FStNomBase + 'ANALYTIQ ' + lStWhere +
                          ' ORDER BY Y_DATECOMPTABLE, Y_JOURNAL, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL ' ; //fb 14/06/2006 FQ 17683
    end;
end;

function TOF_CPCECRANA_VIE.AjouteATobFListe(vTob : Tob) : Boolean;
begin
  Result := True;
  vTob.AddChampSupValeur('SOLDEE', 0, False);
end;

function TOF_CPCECRANA_VIE.BeforeLoad : Boolean;
begin
  inherited BeforeLoad;

  // Test pour déterminer si les comptes saisies sont correctes
  ChargeAnalytique(False);
  ChargeGeneral(False);

  // Autorisation de faire la requête, tout les tests sont OK
  FBoFaireRequete := (not FBoConfidentielAna) and
                     (not FBoConfidentielGen) and
                     (FBoOkControleDate);

  // Modification des colonnes de la grille en fonction des comptes
  InitColFListe;

  //
  Result := True;
end;

function TOF_CPCECRANA_VIE.ChoixAxe(index : integer) : string;
begin
  Case index of
    0 : Result:='A1';
    1 : Result:='A2';
    2 : Result:='A3';
    3 : Result:='A4';
    4 : Result:='A5';
    end;
end;

procedure TOF_CPCECRANA_VIE.AfficheInfoSup;
var
  lPrixUnitaire : Double;

  function Affecte(psz: string): string;
  var
    v: Variant;
  begin
    v := FTobEcrselect.GetValue(psz);
    if VarIsNull(v) then
      Result := ''
    else
      Result := v;
  end;
begin
  THLabel(GetControl('Z_NumLigne')).Caption := '';
  THLabel(GetControl('Z_RefExterne')).Caption := '';
  THLabel(GetControl('Z_DateRefExterne')).Caption := '';
  THLabel(GetControl('Z_Devise')).Caption := '';
  THLabel(GetControl('Z_Debitdev')).Caption := '';

  THLabel(GetControl('Z_Qte1')).Caption := '';
  THLabel(GetControl('Z_Qte2')).Caption := '';
  THLabel(GetControl('Z_QualifQte1')).Caption := '';
  THLabel(GetControl('Z_QualifQte2')).Caption := '';

  THLabel(GetControl('Z_PrixUnitaire1')).Caption := '';
  THLabel(GetControl('Z_PrixUnitaire2')).Caption := '';

  THLabel(GetControl('Z_Affaire')).Caption := '';
  THLabel(GetControl('Z_Table0')).Caption := '';
  THLabel(GetControl('Z_Table1')).Caption := '';
  THLabel(GetControl('Z_Table2')).Caption := '';
  THLabel(GetControl('Z_Table3')).Caption := '';

  if (ComboInfoSup.ItemIndex = 0) or (not RecupEcrSelect) then
    Exit;

  THLabel(GetControl('Z_NumLigne')).Caption       := FTobEcrSelect.GetValue('Y_NUMLIGNE');
  THLabel(GetControl('Z_RefExterne')).Caption     := Affecte('Y_REFEXTERNE');

  THLabel(GetControl('Z_DateRefExterne')).Caption := FTobEcrSelect.GetValue('Y_DateRefExterne');
  THLabel(GetControl('Z_Devise')).Caption         := Affecte('Y_Devise');
  THLabel(GetControl('Z_Debitdev')).Caption       := IIF(FTobEcrSelect.GetValue('Y_CreditDev') <> 0,
                                                  FTobEcrSelect.GetValue('Y_CreditDev'),
                                                  FTobEcrSelect.GetValue('Y_DebitDev'));

  THLabel(GetControl('Z_Qte1')).Caption           := StrfMontant(FTobEcrSelect.GetDouble('Y_QTE1'),
                                                  13, V_Pgi.OkDecQ, '', True);
  THLabel(GetControl('Z_Qte2')).Caption           := StrfMontant(FTobEcrSelect.GetDouble('Y_QTE2'),
                                                  13, V_Pgi.OkDecQ, '', True);

  THLabel(GetControl('Z_Affaire')).Caption        := Affecte('Y_Affaire');
  THLabel(GetControl('Z_Table0')).Caption         := Affecte('Y_Table0');
  THLabel(GetControl('Z_Table1')).Caption         := Affecte('Y_Table1');
  THLabel(GetControl('Z_Table2')).Caption         := Affecte('Y_Table2');
  THLabel(GetControl('Z_Table3')).Caption         := Affecte('Y_Table3');

  if PageinfoSup.ActivePage = PageInfoSup.Pages[1] then begin
    if FTobEcrSelect.GetDouble('Y_QTE1') <> 0 then begin
      if FTobEcrSelect.GetString('Y_DEVISE') <> V_PGI.DevisePivot then
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('Y_DEBITDEV') -
                            FTobEcrSelect.GetDouble('Y_CREDITDEV')) /
                            FTobEcrSelect.GetDouble('Y_QTE1')
      else
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('Y_DEBIT') -
                            FTobEcrSelect.GetDouble('Y_CREDIT')) /
                            FTobEcrSelect.GetDouble('Y_QTE1');
     end
    else
      lPrixUnitaire := 0;

    THLabel(GetControl('Z_PrixUnitaire1')).Caption :=
      StrFMontant(lPrixUnitaire, 13, V_Pgi.OkDecP, '', True) + ' ' +
      RechDom('TTDEVISE', FTobEcrSelect.GetString('Y_DEVISE'), False);

    if lPrixUnitaire <> 0 then
      THLabel(GetControl('Z_PrixUnitaire1')).Caption :=
      THLabel(GetControl('Z_PrixUnitaire1')).Caption + ' / ' +
      RechDom('TTQUALUNITMESURE', FTobEcrSelect.GetString('Y_QUALIFQTE1'), False);

    if FTobEcrSelect.GetDouble('Y_QTE2') <> 0 then begin
      if FTobEcrSelect.GetString('Y_DEVISE') <> V_PGI.DevisePivot then
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('Y_DEBITDEV') -
                           FTobEcrSelect.GetDouble('Y_CREDITDEV')) /
                           FTobEcrSelect.GetDouble('Y_QTE2')
      else
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('Y_DEBIT') -
                           FTobEcrSelect.GetDouble('Y_CREDIT')) /
                           FTobEcrSelect.GetDouble('Y_QTE2');
     end
    else
      lPrixunitaire := 0;

    THLabel(GetControl('Z_PrixUnitaire2')).Caption :=
      StrFMontant(lPrixUnitaire, 13, V_Pgi.OkDecP, '', True) + ' ' +
      RechDom('TTDEVISE', FTobEcrSelect.GetString('Y_DEVISE'), False);

    if lPrixUnitaire <> 0 then
      THLabel(GetControl('Z_PrixUnitaire2')).Caption :=
      THLabel(GetControl('Z_PrixUnitaire2')).Caption + ' / ' +
      RechDom('TTQUALUNITMESURE', FTobEcrSelect.GetString('Y_QUALIFQTE2'), False);
    end;
end;

procedure TOF_CPCECRANA_VIE.IndiceColFListe;
var
  lStChamps, lStListeChamps: string;
  lInIndex: integer;
begin
  lStListeChamps := FStListeChamps;
  lInIndex := 1;
  while lStListeChamps <> '' do begin
     lStChamps := READTOKENST(lStListeChamps);
    if (lStChamps = 'Y_DEBIT') then
      FColDebitE := lInIndex
    else if (lStChamps = 'Y_CREDIT') then
      FColCreditE := lInIndex
    else if (lStChamps = 'SOLDEE') then
      FColSoldePro := lInIndex
    else if (lStChamps = 'Y_SECTION') then
      FColSection := lInIndex
    else if (lStChamps = 'Y_GENERAL') then
      FColGeneral := lInIndex;

    Inc(lInIndex);
    end; // while
end;

procedure TOF_CPCECRANA_VIE.OnDblClickFListe(Sender: TObject);
{$IFDEF COMPTA}
var
  AA: TActionFiche;
  M: RMVT;
  lQEcr: TQuery;
{$ENDIF}
begin
{$IFDEF COMPTA}
  if (not FListe.Focused) or (not RecupEcrSelect) then
    Exit;

  lQEcr := nil;
  try
    lQEcr := OpenSql('SELECT * FROM ' + FStNomBase + 'ANALYTIQ WHERE ' +
      '(Y_JOURNAL = "' + FTobEcrSelect.GetValue('Y_JOURNAL') + '") AND ' +
      '(Y_EXERCICE = "' + FTobEcrSelect.GetValue('Y_EXERCICE') + '") AND ' +
      '(Y_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetValue('Y_DATECOMPTABLE')) + '") AND ' +
      '(Y_NUMEROPIECE = ' + IntToStr(FTobEcrSelect.GetValue('Y_NUMEROPIECE')) + ') AND ' +
      '(Y_NUMLIGNE = ' + IntToStr(FTobEcrSelect.GetValue('Y_NUMLIGNE')) + ') AND ' +
      '(Y_NUMVENTIL = ' + IntToStr(FTobEcrSelect.GetValue('Y_NUMVENTIL')) + ') AND ' + //fb 14/06/2006 FQ 17683
      '(Y_QualifPiece = "' + FTobEcrSelect.GetValue('Y_QUALIFPIECE') + '")', True);

    AA := taModif;
    if not FOkCreateModif then
      AA := taConsult;

    if TrouveSaisieODA(lQEcr, M) then begin
      M.NumLigVisu := lQEcr.FindField('Y_NUMLIGNE').AsInteger;
      LanceSaisieODA(lQEcr, AA, M);
      end;

  finally
    Ferme(lQEcr);
  end;

  RefreshPclPGE;
{$ENDIF}
end;

procedure TOF_CPCECRANA_VIE.OnChangeComboY_Exercice(Sender: TObject);
begin
  if Y_Exercice.ItemIndex = 0 then begin
    if VH^.ExoV8.Code <> '' then
      Y_DateComptable.Text := DateToStr(VH^.ExoV8.Deb)
    else
      // Date de début du premier Exercice
      Y_DateComptable.Text := DateToStr(VH^.Exercices[1].Deb);
    Y_DateComptable_.Text  := DateToStr(iDate2099-1);
   end
  else begin
    if Y_Exercice.ItemIndex = -1 then begin
      if CtxPCl in V_Pgi.PgiContexte then begin
        if VH^.CPExoRef.Code <> '' then
          Y_Exercice.Value := CExerciceVersRelatif ( VH^.CPExoRef.Code )
        else
          Y_Exercice.Value := CExerciceVersRelatif ( VH^.Encours.Code );
       end
      else
        Y_Exercice.Value := CExerciceVersRelatif( VH^.Entree.Code );
      end;
    CExoRelatifToDates(Y_Exercice.Value, Y_DATECOMPTABLE, Y_DATECOMPTABLE_);
    end;

  VerifieLesDates;
end;

procedure TOF_CPCECRANA_VIE.OnClickBSUP(Sender: TObject);
begin
  AnaSuivant(False, True);
end;

procedure TOF_CPCECRANA_VIE.OnClickBSDOWN(Sender: TObject);
begin
  AnaSuivant(True, True);
end;

procedure TOF_CPCECRANA_VIE.OnClickBGUP(Sender: TObject);
begin
  GenSuivant(False, True);
end;

procedure TOF_CPCECRANA_VIE.OnClickBGDOWN(Sender: TObject);
begin
  GenSuivant(True, True);
end;

procedure TOF_CPCECRANA_VIE.OnExitY_DateComptable(Sender: TObject);
begin
  VerifieLesDates;
end;

procedure TOF_CPCECRANA_VIE.OnChangeComboInfoSup(Sender: TObject);
begin
  if ComboInfoSup.Value = '1' then begin
    TPageControl(GetControl('PCUMUL')).Visible := False;
    TPageControl(GetControl('PAGEINFOSUP')).Visible := True;
    TPageControl(GetControl('PCUMUL')).Visible := True;
    AfficheInfoSup;
   end
  else
    TPageControl(GetControl('PAGEINFOSUP')).Visible := False;
end;

procedure TOF_CPCECRANA_VIE.OnChangeComboAcces(Sender: TObject);
begin
  FOkCreateModif := (ComboAcces.Value = '1');
end;

procedure TOF_CPCECRANA_VIE.ApresChangementFiltre;
begin
  inherited;
  ChargeAnalytique(False);
  ChargeGeneral(False);
  CExoRelatifToDates(Y_EXERCICE.Value, Y_DATECOMPTABLE, Y_DATECOMPTABLE_, True);
end;

procedure TOF_CPCECRANA_VIE.OnClickBValider (Sender : TObject);
var
  i, lInPos      : integer;
  lStNomChamp    : string;
  lStListeChamps : string;
begin
  // Enregistrement du détail des écritures dans GTobDetailEcr pour les utiliser
  // avec un verbe OLE
  if GTobDetailEcr <> nil then begin
    if ATobFliste.Detail.Count > 0 then begin
      lStListeChamps := FStListeChamps;
      GTobDetailEcr.Dupliquer(ATobFListe, True, True, False);

      // Ajout de l'entête ANALYTIQUE, GENERAL au niveau de la MERE
      GTobDetailEcr.AddChampSupValeur('SECTION', Y_Section.Text, False);
      GTobDetailEcr.AddChampSupValeur('GENERAL', Y_General.Text, False);
      // Fin d'ajout de l'entête

      for i := 0 to FListe.ColCount-1 do begin
        lStNomChamp := ReadTokenSt( lStListeChamps );
        // Suppression dans GTobDetailEcr des colonnes non visibles de la grille
        if FListe.ColWidths[i+1] <= 0 then begin
          GTobDetailEcr.Detail[0].DelChampSup( lStNomChamp, True);
          end;
        end;

      // Suppression des champs CLE_
      lStListeChamps := cSelect;
      lStListeChamps := FindEtReplace( lStListeChamps, ',', ';', True);
      while lstListeChamps <> '' do begin
        lStNomChamp := ReadTokenSt( lStListeChamps );
        lInPos := Pos( 'CLE_', lStNomChamp );
        if lInPos > 0 then begin
          lStNomChamp := Copy( lStNomChamp, lInPos, Length( lStNomChamp ));
          GTobDetailEcr.Detail[0].DelChampSup( lStNomChamp, True);
          end;
        end;

      // Suppression systématique du champ Y_BLOCNOTE
      GTobDetailEcr.Detail[0].DelChampSup('Y_BLOCNOTE', True);

      end; // FIN if ATobFliste.Detail.Count > 0 then
    end;
end;

procedure TOF_CPCECRANA_VIE.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
end;

function TOF_CPCECRANA_VIE.AfterLoad: boolean;
begin
  Result := inherited AfterLoad;

  if FListe.Focused then begin
    if FOkLastRow then begin
      FListe.Row := FListe.RowCount - 1;
      FOkLastRow := False;
      end
    end;

  // Confidentialité de SECTION
  if FBoConfidentielAna then begin
    PgiInfo(TraduireMemoire('Le compte') + ' ' + Y_Section.Text + ' ' +
            TraduireMemoire('est confidentiel.'),
            TraduireMemoire('Consultation des écritures analytiques'));

    end;

  // Confidentialité du GENERAL
  if FBoConfidentielGen then begin
    PGIInfo(TraduireMemoire('Le compte') + ' ' + Y_General.Text + ' ' +
            TraduireMemoire('est confidentiel.'),
            TraduireMemoire('Consultation des écritures analytiques'));
    end;

  // Affichage des informations complémentaires de l'écriture sélectionnée
  AfficheInfoSup;
end;

procedure TOF_CPCECRANA_VIE.CalculPourAffichage( vTob : Tob );
var
  lSoldePro : Double;
begin
  // Calcul du solde progressif
  if vTob.GetIndex = 0 then
    lSoldePro := 0
  else
    lSoldePro := ATobFListe.Detail[vTob.GetIndex-1].GetDouble('SOLDEE');

  if lSoldePro >= 0 then
    lSoldePro := lSoldePro + vTob.GetDouble('Y_DEBIT') - vTob.GetDouble('Y_CREDIT')
  else
    lSoldePro := lSoldePro - vTob.GetDouble('Y_CREDIT') + vTob.GetDouble('Y_DEBIT');

  vTob.SetDouble('SOLDEE', lSoldePro);
end;

procedure TOF_CPCECRANA_VIE.CumulsANAClick(Sender: TObject);
var
  LeExo: TExoDate;
begin
  LeExo := PositionneExo;
  CumulCpteMensuel(Axe, FTobAna.GetValue('S_SECTION'), FTobAna.GetValue('S_LIBELLE'), LeExo);  //fb 29/05/2006 FQ17690
end;

function TOF_CPCECRANA_VIE.PositionneExo: TExoDate;
var
  LeExo: TExoDate;
begin
  LeExo := VH^.Entree;
  if Y_Exercice.Value = VH^.Encours.Code then
    LeExo := VH^.Encours
  else
    if VH^.Suivant.Code <> '' then begin
      if Y_Exercice.Value = VH^.Suivant.Code then
        LeExo := VH^.Suivant;
     end
    else
      if VH^.Precedent.Code <> '' then begin
        if Y_Exercice.Value = VH^.Precedent.Code then
          LeExo := VH^.Precedent;
        end;

  Result := LeExo;
end;

{$IFDEF COMPTA}
procedure TOF_CPCECRANA_VIE.OnClickCompteSection(Sender: TObject);
var
  lAction: TActionFiche;
begin
  if Y_Section.Text = '' then
    Exit;

  inherited;
  lAction := taModif;
  if ((not ExJaiLeDroitConcept(TConcept(ccSecModif), False))
  or (not FOkCreateModif)) then
    lAction := taConsult;

  FicheSection(nil, ChoixAxe(SAxe.ItemIndex), Y_Section.Text, lAction, 0);
  RefreshPclPGE;
end;
{$ENDIF}

{$IFDEF COMPTA}
procedure TOF_CPCECRANA_VIE.RechercheEcrituresClick(Sender: Tobject);
begin
  MultiCritereAna(taConsult);
end;
{$ENDIF}

function TOF_CPCECRANA_VIE.RecupWhereAGL : string ;
begin
  Result := RecupWhereCritere(PageControl);

  // Traduction des exercices en relatif
  Result := CMajRequeteExercice(Y_EXERCICE.Value, Result);

  // Suppression du WHERE ajouter par l'AGL
  Result := Copy(Result, 7, Length(Result));
end;

function TOF_CPCECRANA_VIE.RecupWhereEcriture ( vBoForceDetailAno : boolean = false ) : string;
var
  lSt  : string ;
  lSt1 : string ;
  lSt2 : string ;
  lDateComptable : TDateTime;
  lstOldValue    : string ;
begin

  lstOldValue       := Y_QualifPiece.Text ;
  Y_QualifPiece.Text := '' ;
  lSt               := RecupWhereAGL;
  Y_QualifPiece.Text := lstOldValue ;


  // Traitement du THMultiValCOmboBox du Y_QUALIFPIECE
  // Attention : vVideSiTous doit toujours être à false avec Y_QUALIFPIECE
  TraductionTHMultiValComboBox( Y_QualifPiece, lSt1, lSt2, 'Y_QUALIFPIECE', False);
  lSt := lSt + ' AND ' +  lSt1; //fb 09/06/2006 FQ17689

  lDateComptable := StrToDate(Y_DateComptable.Text);
  if (VH^.Suivant.Deb > 0) and (lDateComptable > VH^.Suivant.Deb) then
    lSt := lSt + ' AND Y_ECRANOUVEAU="N"'
  else begin
    if vBoForceDetailAno then
      lSt := lSt + ' AND (Y_ECRANOUVEAU="N" OR (Y_ECRANOUVEAU="H" AND Y_DATECOMPTABLE="' + UsDateTime(lDateComptable) + '"))'
    else
      lSt := lSt + ' AND (Y_ECRANOUVEAU="N" OR ((Y_ECRANOUVEAU="H" OR Y_ECRANOUVEAU="OAN") AND Y_DATECOMPTABLE="' + UsDateTime(lDateComptable) + '"))';
    end;

  Result := lSt;
end;

function TOF_CPCECRANA_VIE.RecupCriteresCommuns : string;
var lSt1, lSt2 : string;
    lExoDate : TExoDate;
begin
  Result := ' (Y_DEBIT <> 0 OR Y_CREDIT <> 0) AND Y_CREERPAR <> "DET"';

  // Section
  if Y_Section.Text <> '' then
    Result := Result  + ' AND ' + AxeAna(false) + ' = "' + Y_Section.Text + '"';

  // Général
  if Y_General.text <> '' then
    Result := Result + ' AND Y_GENERAL = "' + Y_General.Text + '"';

  // Etablissement
  TraductionTHMultiValComboBox( Y_Etablissement, lSt1, lSt2, 'Y_ETABLISSEMENT');
  if lSt1 <> '' then
    Result := Result + ' AND ' + lSt1;

  // Devise
  if Y_Devise.ItemIndex <> 0 then
    Result := Result + ' AND Y_DEVISE = "' + Y_DEVISE.Value + '"';

  if VH^.ExoV8.Code <> '' then begin
    if Y_Exercice.ItemIndex <> 0 then begin
      lExoDate.Code := CRelatifVersExercice(Y_Exercice.Value);
      RempliExoDate(lExoDate);
      if lExoDate.Deb >= VH^.ExoV8.Deb then
        Result := Result + ' AND Y_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '"';
      end;
    end;
end;

function TOF_CPCECRANA_VIE.FaireStWhere(vBoForceDetailAno: boolean): string;
var
  lStWhere           : string ;
begin
  lStWhere := '';
  if (Y_Section.Text <> '') or (Y_General.Text <> '') then begin
    lStWhere := ' WHERE (';

    lStWhere := lStWhere + RecupWhereEcriture(vBoForceDetailAno);

    // Ajout des critères communs à la requête
    lStWhere := lStwhere + ') AND ' + RecupCriteresCommuns;

    // On envoie le texte SQL de la requête à la fiche ANCETRE pour qu'elle fasse le OPENSQL
    Result := lStWhere;

   end
  else
  begin
    { BVE FQ 19513 14.05.07 }
    if not gLoad then
       PGIInfo('Vous devez renseigner au moins un des deux critères : Analytique ou Général');
    { END FQ 19513 }
    Result := '';
  end;
end;

function TOF_CPCECRANA_VIE.RecupEcrSelect(TheRow : integer = 0): Boolean;
var
  lTobEcr : Tob;
  lQuery  : TQuery;
  SQL     : string;
begin
  Result := False;
  FTobEcrSelect.InitValeurs; // Init des champs de la tob

  if TheRow = 0 then
   lTobEcr := GetO(FListe, FListe.Row)
  else
   lTobEcr := GetO(FListe, TheRow) ;

  if (lTobEcr=nil) then
    Exit;

  SQL := 'SELECT * FROM ' + FStNomBase + 'ANALYTIQ WHERE ' +
            'Y_JOURNAL= "' + lTobEcr.GetValue('Y_JOURNAL') + '" AND ' +
            'Y_EXERCICE = "' + lTobEcr.GetValue('CLE_EXERCICE') + '" AND ' +
            'Y_DATECOMPTABLE = "' + USDateTime(lTobEcr.GetValue('Y_DATECOMPTABLE')) + '" AND ' +
            'Y_NUMEROPIECE = ' + IntToStr(lTobEcr.GetValue('Y_NUMEROPIECE')) + ' AND ' +
            'Y_NUMLIGNE = ' + IntToStr(lTobEcr.GetValue('CLE_NUMLIGNE')) + ' AND ' +
            'Y_NUMVENTIL = ' + IntToStr(lTobEcr.GetValue('Y_NUMVENTIL')) + ' AND ' + //fb 14/06/2006 FQ 17683
            'Y_QUALIFPIECE = "' + lTobEcr.GetValue('CLE_QUALIFPIECE') + '"';

  lQuery := OpenSQL(SQL, True);
  try
    if not lQuery.Eof then
      Result := FTobEcrSelect.SelectDB('', lQuery );

  finally
    Ferme( lQuery );
    end;
end;

procedure TOF_CPCECRANA_VIE.InitColFListe;
var
  lVisuAna, lVisuGen: Boolean;
begin
  IndiceColFListe;

  lVisuAna := False;
  lVisuGen := False;

  if FTobAna.GetValue('S_SECTION') <> '' then
  begin
    lVisuAna := False;
    lVisuGen := True
  end
  else
  begin
    // GCO - 31/05/2007 - FQ 20499
    if FTobGen.GetValue('G_GENERAL') <> '' then
    begin
      lVisuAna := True;
      lVisuGen := False;
    end;
  end;

  FListe.ColWidths[FColSection]  := IIF(lVisuAna, FListe.ColWidths[FColSection], -1);
  FListe.ColWidths[FColGeneral]  := IIF(lVisuGen, FListe.ColWidths[FColGeneral], -1);
  FListe.ColWidths[FColSoldePro] := IIF(ComboSoldePro.Value = '0', -1, FListe.ColWidths[FColCreditE]);
end;

procedure TOF_CPCECRANA_VIE.VerifieLesDates;
begin
  FBoOkControleDate := True;
  if Y_EXERCICE.ItemIndex = 0 then begin
    if (VH^.ExoV8.Code <> '') then begin
      if (FTobGen.GetValue('G_GENERAL') <> '') or
        (FTobAna.GetValue('S_SECTION') <> '') then begin
        if StrToDate(Y_DateComptable.Text) < VH^.ExoV8.Deb then begin
          PgiInfo('Vous ne pouvez pas saisir une date inférieure au  ' + DateToStr(VH^.ExoV8.Deb) + '.', 'Consultation des écritures analytiques');
          Y_DateComptable.Text := DateToStr(VH^.ExoV8.Deb);
          FBoOkControleDate := False;
          end;
        end;
     end
    else begin
      if StrToDate(Y_DateComptable.Text) < VH^.Exercices[1].Deb then begin
        PgiInfo('Vous ne pouvez pas saisir une date inférieure au  ' + DateToStr(VH^.Exercices[1].Deb) + '.', 'Consultation des écritures analytiques');
        Y_DateComptable.Text := DateToStr(VH^.Exercices[1].Deb);
        FBoOkControleDate := False;
        end;
      end;
    end;
end;

function TOF_CPCECRANA_VIE.ChargeAnalytique( Force : boolean ) : Boolean;
var
  lQuery : TQuery;
begin
  Result := False;

  if ((FTobAna = nil) or (csDestroying in Y_Section.ComponentState)) then
    Exit;

  if Y_Section.Text = '' then begin
    FBoConfidentielAna := False;
    FTobAna.InitValeurs;
    SLibelle.Caption := '';
    Y_Section.DataType := 'TZSECTION';
    Y_Section.Plus := ' AND S_AXE="A' + IntToStr(SAxe.itemindex+1) + '" ';
    Exit;
    end;

  // Placement sur la dernière écriture si changement de compte
  if (not FOkLastRow) and (Y_Section.Text <> FTobAna.GetValue('S_SECTION')) then
    FOkLastRow := True;

  Y_Section.Text := BourreEtLess(Y_Section.Text, Axe);
  try
    lQuery := OpenSQL('SELECT * FROM ' + FStNomBase + 'SECTION WHERE ' +
                      'S_SECTION = "' + Y_Section.Text + '"' +
                      ' AND S_AXE="A' + IntToStr(SAxe.itemindex+1) + '" ', True);

    if not lQuery.Eof then begin
      FTobAna.SelectDB('', lQuery);
      // Test de la confidentialité des comptes
      FBoConfidentielAna := EstConfidentiel(FTobAna.GetValue('S_CONFIDENTIEL'));
      if not FBoConfidentielAna then begin
        SLibelle.Caption := FTobAna.GetValue('S_LIBELLE');
        VerifieLesDates;
        end;
     end
    else begin
      Y_General.DataType := 'TZGENERAL';
{b fb 09/06/2006 FQ17684}
      //      Y_General.ElipsisClick(nil);
      Y_Section.ElipsisClick(nil);
{e fb 09/06/2006 FQ17684}
      end;
  finally
    Ferme( lQuery );
    end;
end;

procedure TOF_CPCECRANA_VIE.AnaSuivant(vBoSuiv : Boolean; vBoResteSurZone : Boolean);
var
  lQuery: TQuery;
  lSQL, lCptAna, lNewCpt, lCptLibelle: string;
begin
  AFocusFListe := False;
  lCptAna      := Y_Section.text;
  lNewCpt      := lCptAna;

  if vBOSuiv then
    lSQL := 'SELECT S_SECTION,S_LIBELLE FROM ' + FStNomBase + 'SECTION WHERE S_SECTION>"' + lCptAna + '"'
  else
    lSQL := 'SELECT S_SECTION,S_LIBELLE FROM ' + FStNomBase + 'SECTION WHERE S_SECTION<"' + lCptAna + '"';

  // Défilement des comptes
  case ComboDefilCpt.ItemIndex of
    0 : ; // Tous les Comptes
    1 : lSQL := lSQL + ' AND (S_TOTDEBE - S_TOTCREE) <> 0 ';
    2 : lSQL := lSQL + ' AND ((S_TOTDEBE <> 0 OR S_TOTCREE <> 0) AND (S_DATEDERNMVT > "' + UsDateTime(iDate1900) + '"))';
    end;

  // Gestion du V_Pgi.Confidentiel
  lSql := lSql + ' AND ' + CGenereSQLConfidentiel('S');

  // Order By
  if vBoSuiv then
    lSQL := lSQL + ' ORDER BY S_SECTION '
  else
    lSQL := lSQL + ' ORDER BY S_SECTION DESC ';

  lQuery := OpenSQL(lSQL, True);
  if not lQuery.EOF then begin
    lNewCpt     := lQuery.FindField('S_SECTION').AsString;
    lCptLibelle := lQuery.FindField('S_LIBELLE').AsString;
    end;
  Ferme(lQuery);

  if lNewCpt <> lCptAna then begin
    AFocusFListe := True;
    Y_Section.Text := lNewCpt;
    SLibelle.Caption := lCptLibelle;
    ChargeAnalytique(False);
    RefreshPclPge;
    if vBoResteSurZone then
      SetFocusControl('YSECTION');
    end;
end;

function TOF_CPCECRANA_VIE.ChargeGeneral(Force: boolean): Boolean;
var
  lQuery : TQuery;
begin
  Result := False;

  if ((FTobGen = nil) or (csDestroying in Y_General.ComponentState)) then
    Exit;

  if Y_General.Text = '' then begin
    FBoConfidentielGen := False;
    FTobGen.InitValeurs;
    GLibelle.Caption := '';
    Y_General.DataType := 'TZGENERAL';
    Exit;
    end;

  // Placement sur la dernière écriture si changement de compte
  if (not FOkLastRow) and (Y_General.Text <> FTobGen.GetValue('G_GENERAL')) then
    FOkLastRow := True;

  Y_General.Text := BourreEtLess(Y_General.Text, fbgene);
  try
    lQuery := OpenSQL('SELECT * FROM ' + FStNomBase + 'GENERAUX WHERE ' +
                      'G_GENERAL = "' + Y_GENERAL.Text + '"', True);

    if not lQuery.Eof then begin
      FTobGen.SelectDB('', lQuery);
      // Test de la confidentialité des comptes
      // GCO - 03/01/2007 - FQ 19393
      // FBoConfidentielGen := EstConfidentiel(FTobAna.GetValue('G_CONFIDENTIEL'));
      FBoConfidentielGen := EstConfidentiel(FTobGen.GetValue('G_CONFIDENTIEL'));
      if not FBoConfidentielGen then begin
        GLibelle.Caption := FTobGen.GetValue('G_LIBELLE');
        VerifieLesDates;
        end;
     end
    else begin
//      Y_General.DataType := 'TZGENERAL'; //fb 09/06/2006 FQ17684
      Y_General.ElipsisClick(nil);
      end;
  finally
    Ferme( lQuery );
    end;
end;

procedure TOF_CPCECRANA_VIE.GenSuivant(vBoSuiv : Boolean; vBoResteSurZone : Boolean);
var
  lQ: TQuery;
  lSQL, lCptGen, lNewCpt: string;
  lCptLibelle: string;
begin
  AFocusFListe := False;
  lCptGen := Y_General.Text;
  lNewCpt := lCptGen;

  if vBoSuiv then
    lSQL := 'SELECT G_GENERAL, G_LIBELLE FROM ' + FStNomBase + 'GENERAUX WHERE G_GENERAL > "' + lCptGen + '" '
  else
    lSQL := 'SELECT G_GENERAL, G_LIBELLE FROM ' + FStNomBase + 'GENERAUX WHERE G_GENERAL < "' + lCptGen + '" ';

  // Défilement des comptes
  case ComboDefilCpt.ItemIndex of
    0: ; // Tous les Comptes
    1: lSQL := lSQL + ' AND (G_TOTDEBE - G_TOTCREE) <> 0 ';
    2: lSQL := lSQL + ' AND ((G_TOTDEBE <> 0 OR G_TOTCREE <> 0) AND (G_DATEDERNMVT > "' + UsDateTime(iDate1900) + '"))';
    end;

  // Gestion du V_Pgi.Confidentiel
  lSql := lSql + ' AND ' + CGenereSQLConfidentiel('G');

  // Order by
  if vBoSuiv then
    lSQL := lSQL + ' ORDER BY G_GENERAL '
  else
    lSQL := lSQL + ' ORDER BY G_GENERAL DESC ';

  lQ := OpenSQL(lSQL, True);
  if not lQ.EOF then begin
    lNewCpt := lQ.FindField('G_GENERAL').AsString;
    lCptLibelle := lQ.FindField('G_LIBELLE').AsString;
    end;
  Ferme(lQ);

  if lNewCpt <> lCptGen then begin
    AFocusFListe := True;
    Y_General.Text := lNewCpt;
    SLibelle.Caption := lCptLibelle;
    ChargeGeneral(True);
    RefreshPclPge;
    if vBoResteSurZone then
      SetFocusControl('YGENERAL');
    end;
end;

procedure TOF_CPCECRANA_VIE.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
  case Key of
    VK_DOWN, VK_F4 : begin
                       if Y_SECTION.Focused then begin
                         Key := 0;
                         AnaSuivant(True, ssCtrl in Shift);
                         end;

                       if Y_GENERAL.Focused then begin
                         Key := 0;
                         GenSuivant(True, ssCtrl in Shift);
                         end;
                       end;

     VK_UP, VK_F3 : begin
                      if Y_SECTION.Focused then begin
                        Key := 0;
                        AnaSuivant(False, ssCtrl in Shift);
                        end;

                      if Y_GENERAL.Focused then begin
                        Key := 0;
                        GenSuivant(False, ssCtrl in Shift);
                        end;
                      end;
    end;
end;

procedure TOF_CPCECRANA_VIE.OnColumnWidthsChanged( Sender : TObject );
begin
  Fliste.ColWidths[0]:=10;
end;

procedure TOF_CPCECRANA_VIE.OnArgument (S : String ) ;
begin
  // FQ 19513
  gLoad := true;
  
  FFI_Table := cFI_TABLE;
  FStListeParam := 'CPCECRANA';
  inherited;
  Ecran.HelpContext := 7602100 ; {Lek 040706}
  BBlocNote.Down:=false;  //fb 29/05/2006 FQ17688
  IndiceColfListe;

  // Récupération des arguments
  FStArgumentTOF := S;

  { FQ 19512 BVE 14.05.07 }
  SAxe:= THValComboBox(GetControl('Y_AXE', True));
  { END FQ 19512 }

  Fliste.SynEnabled:=false;
  FListe.OnColumnWidthsChanged := OnColumnWidthsChanged;
  // Elements de la fiche Ancêtre
  FListe.GetCellCanvas   := GetCellCanvasFListe;
  FListe.OnDblClick      := OnDblClickFListe;
  FListe.OnRowEnter      := OnRowEnterFListe;

  FListe.OnKeyDown       := OnKeyDownFListe;

  // Elements de la TOF
  FTobEcrSelect   := Tob.Create('ANALYTIQ', nil, -1);
  FTobAna         := Tob.Create('SECTION', nil, -1);
  FTobGen         := Tob.Create('GENERAUX', nil, -1);
  FZListJournal   := TZListJournal.Create;

  SLibelle        := THLabel(GetControl('SLIBELLE'));
  GLibelle        := THLabel(GetControl('GLIBELLE'));

  BSUP            := TToolBarButton97(GetControl('BSUP', True));
  BSDOWN          := TToolBarButton97(GetControl('BSDOWN', True));
  BGUP            := TToolBarButton97(GetControl('BGUP', True));
  BGDOWN          := TToolBarButton97(GetControl('BGDOWN', True));

  PopUpTraitement := TPopUpMenu(GetControl('PopUpTraitement', True));
  PopUpUtilitaire := TPopUpMenu(GetControl('PopUpUtilitaire', True));
  PopUpFctComp    := TPopUpMenu(GetControl('PopUpFctComp', True));
  PopUpEdition    := TPopUpMenu(GetControl('PopUpEdition', True));

  Y_SECTION        := THEdit(GetControl('YSECTION', True));
  Y_General        := THEdit(GetControl('YGENERAL', True)) ;
  Y_DateComptable  := THEdit(GetControl('Y_DATECOMPTABLE', True));
  Y_DateComptable_ := THEdit(GetControl('Y_DATECOMPTABLE_', True));
  LESOLDE          := THEdit(GetControl('LESOLDE', True));
  Y_NumeroPiece    := THEdit(GetControl('Y_NumeroPiece', True));
  Y_NumeroPiece_   := THEdit(GetControl('Y_NumeroPiece_', True));
  Y_RefInterne     := THEdit(GetControl('Y_REFINTERNE', True));

  Y_Valide         := TCheckBox(GetControl('Y_VALIDE', True));

  Y_Journal        := THMultiValComboBox(GetControl('Y_Journal', True));
  Y_NaturePiece    := THMultiValComboBox(GetControl('Y_NaturePiece', True));
  Y_Devise         := THValComboBox(GetControl('YDEVISE'));
  Y_Exercice       := THValComboBox(GetControl('Y_EXERCICE', True));
  CInitComboExercice(Y_EXERCICE);

  Y_QualifPiece    := THMultiValComboBox(GetControl('Y_QUALIFPIECE', False));
  if Y_QualifPiece = nil then begin
    THMultiValComboBox(GetControl('Y_QUALIFPIECE')).Name := 'YQUALIFPIECE';
    Y_QualifPiece   := THMultiValComboBox(GetControl('YQUALIFPIECE', False));
    end;
  Y_QualifPiece.ReadOnly := True;

  Y_Etablissement  := THMultiValComboBox(GetControl('YETABLISSEMENT', True));

  ComboInfoSup     := THValComboBox(GetControl('COMBOINFOSUP', True));
  ComboAcces       := THValComboBox(GetControl('COMBOACCES', True));
  ComboDefilCpt    := THValComboBox(GetControl('COMBODEFILCPT', True));
  ComboSoldePro    := THValComboBox(GetControl('COMBOSOLDEPRO', True));

  PageInfoSup      := TPageControl(GetControl('PAGEINFOSUP', True));

  //----------------------- Branchement des événements -----------------------//
  Ecran.OnKeyDown     := OnKeyDownEcran;
  Y_Exercice.OnChange := OnChangeComboY_Exercice;

  BSUP.OnClick         := OnClickBSUP;
  BSDOWN.OnClick       := OnClickBSDOWN;
  BGUP.OnClick         := OnClickBGUP;
  BGDOWN.OnClick       := OnClickBGDOWN;

  Y_Section.OnElipsisClick := OnElipsisClickS_Section;
  Y_Section.OnDblClick     := OnElipsisClickS_Section;
  Y_Section.OnExit         := OnExitS_Section;

  Y_General.OnElipsisClick   := OnElipsisClickY_General;
  Y_General.OnDblClick       := OnElipsisClickY_General;
  Y_General.OnExit           := OnExitG_General;

  Y_DateComptable.OnExit        := OnExitY_DateComptable;

  ComboInfoSup.OnChange         := OnChangeComboInfoSup;
  ComboAcces.OnChange           := OnChangeComboAcces;

  PopF11.OnPopup                := OnPopUpPopF11;
  PopUpTraitement.OnPopUp       := OnPopUpPopUpTraitement;
  PopUpUtilitaire.OnPopUp       := OnPopUpPopUpUtilitaire;
  PopUpFctComp.OnPopUp          := OnPopUpPopUpFctComp;
  PopUpEdition.OnPopUp          := OnPopUpPopUpEdition;

  {$IFDEF COMPTA}

  {$ELSE}
    PopF11.Items.Clear;
    TToolBarButton97(GetControl('BOutil')).Visible := False;
    TToolBarButton97(GetControl('BTraitement')).Visible := False;
    TToolBarButton97(GetControl('BUtilitaire')).Visible := False;
    TToolBarButton97(GetControl('BFctComplementaire')).Visible := False;
    TToolBarButton97(GetControl('BEdition')).Visible := False;
    TToolBarButton97(GetControl('BFctMultiSelect')).Visible := False;
  {$ENDIF}

  Y_Section.MaxLength    := VH^.CPta[Axe].Lg;
  Y_General.MaxLength := VH^.Cpta[fbGene].Lg;
  Y_Exercice.ItemIndex   := -1;

  ComboInfoSup.Value := '0';

  if ComboAcces.Enabled then
    ComboAcces.Value := '1';

  ComboDefilCpt.Value := '0';
  ComboSoldePro.Value := '1';

  if SAxe<>nil then
    SAxe.ItemIndex:=0;

  if Y_Journal<>nil then
    Y_Journal.value:=TraduireMemoire('<<Tous>>');

  if Y_Exercice<>nil then
    Y_Exercice.ItemIndex:=0;

  if Y_NaturePiece<>nil then
    Y_NaturePiece.value:=TraduireMemoire('<<Tous>>');

  if Y_QualifPiece<>nil then
    Y_QualifPiece.value:=TraduireMemoire('<<Tous>>');

  if Y_Etablissement<>nil then
    Y_Etablissement.value:=TraduireMemoire('<<Tous>>');

{$IFDEF BUREAU}
  FBoBloqueExercice := False;
{$ENDIF}
  Fliste.SynEnabled:=true;

  InitControl; //fb 19/05/2006 FQ17693
end ;

procedure TOF_CPCECRANA_VIE.InitControl;
begin
  OnChangeComboY_Exercice( nil );
end;

Initialization
  registerclasses ( [ TOF_CPCECRANA_VIE ] ) ;
end.


