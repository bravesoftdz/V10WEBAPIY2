{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 24/07/2003
Modifié le ... : 03/10/2007
Description .. : Source TOF de la FICHE : CPBALANCE ()
Suite ........ : 
Suite ........ : Source commun pour QRS1 des balances :
Suite ........ : TOF_METH --> TOF_CPBALANCE --> xxx
Suite ........ : GCO - 11/03/2004
Suite ........ : Ajout de la gestion du critEdt pour appel des Balances
Suite ........ :
Suite ........ : JP 03/10/07 : FQ 21560 : gestion de l'appel depuis
Suite ........ : CPTOTRUB_TOF
Mots clefs ... : TOF;CPBALALANCE
*****************************************************************}
Unit CPBALANCE_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl, eQRS1,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,
{$ENDIF}
     sysutils, ComCtrls, Spin, Ent1,
     Filtre,
     HCtrls, HEnt1, HMsgBox, UTOF, hqry,
     uTob,
     TofMeth,
     uLibExercice,      // CInitComboExercice(E_EXERCICE)
     ed_tools,          // barre de progression (form)
     hstatus,           // barre de progression inside
     AGLInit,           // TheData
     CritEdt,           // ClassCritEdt
     Htb97,             // TToolBarButton97
     uMultiDossierUtil, // RequeteMultiDossier
     utilPGI,           // EstTablePartagee
    {$IFDEF eAGLCLIENT}
    MenuOLX,
    {$ELSE}
    MenuOLG,
    {$ENDIF eAGLCLIENT}
     forms,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     dialogs
     ;

Type TTypeBal = ( balGene, balAuxi, balAnal, balAgee, balVentil, balCumul ) ;

Type
  TOF_CPBALANCE = Class (TOF_Meth)

    // Page control principal
    Pages                       : TPageControl;

    // ---------------
    // Onglet standard
    CompteDe, CompteA           : THEdit;
    DateComptaDe, DateComptaA   : THEdit;
    QualifPiece                 : THMultiValcomboBox;
    Exercice                    : THValComboBox;
    ModeSelect                  : THValComboBox;
    NatureCpt                   : THValComboBox;

    // ---------------
    // Onglet Complément
    Devise                      : THValcomboBox;
    CptExcept                   : THEdit;

    // ---------------
    // Onglet options
    Affichage                   : THRadioGroup;
    AvecComparatif              : TCheckBox;

    // ---------------
    // Onglet mise en page
    Rupture                     : THRadioGroup;

    // ---------------
    // Onglet Comparatif
    RuptureType, ComparType     : THRadioGroup;
    // --> période
    ComparExo                   : THValcomboBox;
    ComparDateDe, ComparDateA   : THEdit;
    // --> balsit
    BalSit                      : THEdit;

    // ---------------
    // Onglet Rupture
    // --> Table libre
    CptLibresExist              : TCheckBox;
    TableLibre                  : THValComboBox;
    LibreDe, LibreA             : THEdit;
    // --> N° compte
    NivoRupture                 : TSpinEdit;
    RuptureMax                  : TCheckBox ;
    // --> Plan corresp
    CorrespDe, CorrespA         : THValcomboBox;
    CptCorrespExist             : TCheckBox;
    Corresp                     : THValComboBox;

    // ---------------
    // Champs Invisibles
    AvecNivoRupture             : THEdit;
    AvecRupture                 : THEdit;
    AvecRupType                 : THEdit;
    AvecQualifPiece             : THEdit;
    AvecAffichage               : THEdit;
    XXRupture                   : THEdit;
    XXWhere                     : THEdit;
    SectionsFermees             : THEdit;
    
    // EVT TOF
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure ChargementCritEdt        ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnAfterFormShow          ; override ; 
    procedure SetTypeBalance           ; virtual ;

    // EVT FICHE
    // -> Onglet standard
    procedure CompteOnExit       (Sender: TObject) ;  virtual ;
    Function  CompleteAuto       (Sender: TObject ; CompteType : TFichierBase ) : Boolean ; virtual;
    procedure ExoOnChange        (Sender: TObject) ; virtual;
    procedure ModeSelectChanged  (Sender: TObject) ;
    procedure DateOnExit         (Sender: TObject) ;
    procedure QualifPieceEnter   (Sender: TObject) ;
    // -> Onglet Compléments
    procedure DeviseChanged      (Sender: TObject) ;
    // -> Onglet avancés
    procedure AvecComparatifClick(Sender: TObject) ;
    // -> Onglet mise en page
    procedure RuptureClick       (Sender: TObject) ; virtual ;
    // -> Onglet rupture
    procedure PlanCorrespChanged (Sender: TObject) ;
    procedure RuptureTypeClick   (Sender: TObject) ;
    procedure CorrespChanged     (Sender: TObject) ;
    procedure LibreChanged       (Sender: TObject) ;
    procedure RuptureMaxClick    (Sender: TObject) ;
    // -> Onglet comparatif
    procedure ComparTypeClick    (Sender: TObject) ;
    procedure ComparExoOnChange  (Sender: TObject) ;
    procedure ComparDateOnExit   (Sender: TObject) ;
    // -> Pour états chaînés NE PAS VIRER !!!!!!!!
    procedure FTimerTimer        (Sender: TObject) ;

    // INIT FORM
    procedure InitComposants     ; virtual ;

    // PARAM ETAT
    procedure ParamRuptures      ; virtual ;
    procedure ParamTitresCols    ; virtual ;
    procedure ParamComparatif    ; virtual ;
    procedure ParamDivers        ; virtual ;

    // FILTRES
    procedure InitZones          ; virtual ;
    procedure NouvRechClick      (Sender: TObject) ;
    procedure SupprFiltreClick   (Sender: TObject) ;
    {$IFNDEF GCGC}
    procedure OnChangeFiltre     (Sender: TObject) ; override;
    {$ENDIF}

    // TRAITEMENT

    // -> Méthodes de constitution de la table temporaire
    procedure RemplirEDTBALANCE             ; virtual ;
    procedure DeleteCEDTBALANCE ;
    procedure UpdateCumulsCEDTBALANCE       ; virtual ;
    // -> Construit la requête de l'état
    function  GenererRequeteBAL   : String  ; virtual ;
    function  GetConditionRupture : String  ; virtual ;
    // -> Construction des conditions du WHERE pour les cas spéciaux
    function  RecupWhereQualifPiece : String ;   virtual; // pour éviter les 'like 'xx%' or like...)


    // EVT A REDEFINIR
    procedure NatureCptChanged   (Sender: TObject); virtual ;
    procedure NatureAxeChanged   (Sender: TObject); virtual ;
    procedure TableLibreChanged  (Sender: TObject); virtual ;
    Function  GetNumTableLibre : String ;           override ;

    // NOUVEAU TRAITEMENT

    // -> Construit la requête de constitution de la liste des comptes cibles
    function  GenererInsertCPT         (vBoCompar : Boolean = False) : String ; virtual;
    function  GenererInsertCPTBalSit                                 : String ; virtual;
    function  GenererInsertBalanceCompar                             : String ; virtual;

    // Gestion Regroupement
    function  GetTablePourBase         ( vStTable : String )         : String ;

    //Gestion des filtres :
    function  GetValeurFiltre          ( vStZone : String )          : String ;
    procedure GereAffichageGlobal ;
    procedure MySelectFiltre                                         ; override ;

    // Nouvelles gestion des requêtes de maj
    function  GetConditionSQL( vBoAvecCpt : Boolean ; vInPer : Integer )     : String ; virtual ;
    function  GetConditionSQLEcr                                             : String ; virtual ;
    function  GetConditionSQLCpt                                             : String ; virtual ;
    {JP 03/10/07 : FQ 21560 : depuis CPTOTRUB_TOF, on n'a pas une fourchette de comptes,
                   mais une clause where avec des LIKE, BETWEEN}
    function  GetConditionSQLCptSpecif                                       : String ; virtual ;
    function  GetConditionSQLPeriode ( vInPer : Integer )                    : String ; virtual ;
    function  GetConditionSQLDateC   ( vInPer : Integer )                    : String ; virtual ;
    function  GetSQLCumul            ( vInPer : Integer ; vInCol : Integer ) : String ; virtual ;
    function  GetJointureCpt                                                 : String ; virtual ;
    function  GetWhereNature                                                 : String ; virtual ;
    function  GetBaseInsertCpt                                               : String ;
    function  GetCompte2                                                     : string ; virtual ;
    function  GetLibelle2                                                    : string ; virtual ;
    function  GetCollectif                                                   : string ; virtual ;
    function  GetGroupBy                                                     : string ; virtual ;

    // Divers
    function GetNumAxe      : String ;  // retourne l'axe en fonction de l'index


  private
    // Gestion Regroupement
    FBoRegroupement             : Boolean ;
    FStBase                     : String ;

    // Variable gestion des dates
    DateD                       : TDatetime;
    DateF                       : TDatetime;
    DateDCompar                 : TDatetime;
    DateFCompar                 : TDatetime;

    // Type de balance
    FTypeBal                    : TTypeBal ;

    // Nouvelle gestion des requête de maj
    FStTableEcr                 : String ;
    FStPfEcr                    : String ;
    FStTableCpt                 : String ;
    FStPfCpt                    : String ;
    FStChampCpt                 : String ;
    FStChampNatCpt              : String ;

  public

    DateDebutPer                : TDatetime;
    DateFinPer                  : TDatetime;
    DateDebutCompar             : TDatetime;
    DateFinCompar               : TDatetime;

    // ======================
    // ===== PROPRIETES =====
    // ======================
    property TypeBal     : TTypeBal   read  FTypeBal        write  FTypeBal ;
    property TableEcr    : String     read  FStTableEcr     write  FStTableEcr ;
    property PfEcr       : String     read  FStPfEcr        write  FStPfEcr ;
    property TableCpt    : String     read  FStTableCpt     write  FStTableCpt ;
    property PfCpt       : String     read  FStPfCpt        write  FStPfCpt ;
    property ChampCpt    : String     read  FStChampCpt     write  FStChampCpt ;
    property ChampNatCpt : String     read  FStChampNatCpt  write  FStChampNatCpt ;
    property MultiSoc    : Boolean    read  FBoRegroupement write  FBoRegroupement ;

  end ;



Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  uLibWindows,        // TraductionTHMultiValComboBox
  paramsoc;            //b FP GetPAramSoc


{TOF_CPBALANCE}

procedure TOF_CPBALANCE.OnNew ;
begin
  Inherited ;
  // Initialisation des champs si pas de filtres
  if FFiltres.Text = '' then
    NouvRechClick( nil ) ;
end ;

procedure TOF_CPBALANCE.OnUpdate ;
var lStListeBases : String ;
    lStTitre      : String ;
begin
  Inherited ;

  // Ecran de patience
  lStTitre      := Ecran.Caption ;
  Ecran.Caption := lStTitre + '  -  Préparation de l''état en cours...';
  UpdateCaption( Ecran ) ;
  SourisSablier ;
  InitMove( 100 , TraduireMemoire('Préparation de l''état en cours...') ) ;

  FStBase := '' ;
  if FBoRegroupement and ( GetControlText('MULTIDOSSIER') <> '' ) then
    begin
    // Récupération des bases de regroupement
    RequeteMultiDossier( Pages, lStListeBases ) ;

    // Remplissage table temporaire pour chaque base :
    FStBase := ReadTokenSt( lStListeBases ) ;
    while FStBase <> '' do
      begin
      RemplirEDTBALANCE ;
      FStBase := ReadTokenSt( lStListeBases ) ;
      end ;

    end
  else
    // Remplissage table temporaire
    RemplirEDTBALANCE;

  // Requête finale exécutée par l'état
  TFQRS1(Ecran).WhereSQL := GenererRequeteBAL ;

  // Ecran de patience
  Ecran.Caption := lStTitre ;
  UpdateCaption( Ecran ) ;
  FiniMove;
  SourisNormale;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPBALANCE.ChargementCritEdt;
begin
  inherited;

  if (TheData <> nil) and (TheData is ClassCritEdt) then
    begin
    // Comptes
    CompteDe.Text := ClassCritEdt(TheData).CritEdt.Cpt1;
    CompteA.Text  := ClassCritEdt(TheData).CritEdt.Cpt2;
    // Qualifpièce
    QualifPiece.Text := ClassCritEdt(TheData).CritEdt.QualifPiece;
    // Etablissement
    ComboEtab.Value := ClassCritEdt(TheData).CritEdt.Etab;
    // Devise
    Devise.Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;

    if not ( FTypeBal in [balAgee,balVentil] ) then
      begin
      // Exercice
      Exercice.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);
      // Date de Début de l'édition
      DateComptaDe.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
      // Date de Fin de l'édition
      DateComptaA.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);
      if ClassCritEdt(TheData).CritEdt.AvecComparatif then
      begin
        AvecComparatif.Checked := True;
        ComparExo.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.ExoComparatif.Code);
        ComparDateDe.Text := DateToStr(ClassCritEdt(TheData).CritEdt.ExoComparatif.Deb);
        ComparDateA.Text := DateToStr(ClassCritEdt(TheData).CritEdt.ExoComparatif.Fin);
      end
      end ;

    TheData := nil;

    end
  // FQ17010 Attention le type utilisé diffère pour les états chaînés
  else if (TheData <> nil) and (TheData is ClassCritEdtChaine) then
    begin

    with ClassCritEdtChaine(TheData).CritEdtChaine do
      if UtiliseCritStd then
        begin
        // Transformation ancien crit >> nouveaui crit
        // NOR : Normal
        // NSS : Normal, situation, simulation
        // PRE : prévision
        // SSI : Situation et simulation
        // TOU : Tout
        if TypeEcriture = 'NOR' then
          QualifPiece.Text := 'N;'
        else if TypeEcriture = 'NSS' then
          QualifPiece.Text := 'N;S;U;'
        else if TypeEcriture = 'PRE' then
          QualifPiece.Text := 'P;'
        else if TypeEcriture = 'SSI' then
          QualifPiece.Text := 'S;U;'
        else if TypeEcriture = 'TOU' then
          QualifPiece.Text := TraduireMemoire('<<Tous>>') ;

        if FTypeBal = balGene then
          NatureCpt.Value := NatureCompte ;

        if not ( FTypeBal in [balAgee,balVentil] ) then
          begin
          // Exercice
          self.Exercice.Value := Exercice.Code ;
          // Date de Début de l'édition
          DateComptaDe.Text := DateToStr( Exercice.Deb );
          // Date de Fin de l'édition
          DateComptaA.Text := DateToStr( Exercice.Fin );
          // ModeSelection
          ModeSelect.Value := ModeSelection ;
          end ;

        end ;

    TheData := nil;

    end;



end;
////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPBALANCE.OnLoad ;
begin

  NextPrevControl(TFQRS1(Ecran));

  Inherited ;

  // Les ruptures...
  ParamRuptures ;

  // Le comparatif...
  ParamComparatif ;

  // Paramétrage des Titre des Colonnes
  ParamTitresCols ;

  // Paramétrage divers
  ParamDivers ;

end ;

procedure TOF_CPBALANCE.OnArgument (S : String ) ;
begin

  // ============================
  // ==== Ajout type balance ====
  // ============================
  SetTypeBalance ;

  Inherited ;

  // Gestion du mode "MULTI-DOSSIER"
  FBoRegroupement := pos( 'MULTIDOSSIER', S ) > 0 ;

  // Modification du nom du filtre pour les éditions du menu "multidossier"
  if FBoRegroupement then
    TFQRS1(Ecran).FNomFiltre   := TFQRS1(Ecran).FNomFiltre + 'MD' ;

  // =====> Récupération des contrôles...
  InitComposants ;
end ;

//==============================================================================

procedure TOF_CPBALANCE.RuptureTypeClick(Sender: TObject);
begin

  SetControlVisible('RUPGROUPES',False);
  SetControlVisible('RUPLIBRES',False);
  SetControlVisible('RUPCORRESP',False);
  if RuptureType.value <> '' then
    SetControlVisible(RuptureType.value,True);

  //SG6 14.03.05 Gestion des filtres FQ 14996
  if cLoadFiltre then Exit;

  if RuptureType.value = 'RUPCORRESP' then
    begin
    if Corresp.Items.Count >= 1
    	then Corresp.ItemIndex := 1
    	else Corresp.ItemIndex := 0;
    PlanCorrespChanged(nil);
    end;
  if RuptureType.value = 'RUPLIBRES' then
    begin
    TableLibre.ItemIndex := 0;
    TableLibreChanged(nil);
    end;
  if RuptureType.value = 'RUPGROUPES' then
    begin
    NivoRupture.Value := 1 ;
    RuptureMax.Checked := False ;     
    end ;
end;

procedure TOF_CPBALANCE.AvecComparatifClick(Sender: TObject);
begin
  if AvecComparatif.checked then
    begin
    SetControlVisible('COMPARATIF',True) ;
    //SG6 14.03.05 Gestion des filtres FQ 14996
    if cLoadFiltre then Exit;
    ComparTypeClick(nil) ;
    end
  else
    begin
    SetControlVisible('COMPARATIF',False) ;
    end ;
end;

procedure TOF_CPBALANCE.RuptureClick(Sender: TObject);
begin
  if Rupture.value <> 'SANS' then
    begin
    SetControlVisible('RUPTURES',True);
    RuptureTypeClick(nil);
    end
  else
    begin
    SetControlVisible('RUPTURES',False);
    RuptureMax.Checked := False ;
    end;
end;

procedure TOF_CPBALANCE.DateOnExit(Sender: TObject) ;
begin
  DoDateOnExit(THEdit(Sender), DateComptaDe, DateComptaA, DateD, DateF);
end;

procedure TOF_CPBALANCE.ExoOnChange(Sender: TObject) ;
begin
  CExoRelatifToDates( Exercice.Value, DateComptaDe, DateComptaA);
  DateD := StrToDate(DateComptaDe.Text);
  DateF := StrToDate(DateComptaA.Text);
end;

procedure TOF_CPBALANCE.PlanCorrespChanged(Sender: TObject);
begin
  if RuptureType.value <> 'RUPCORRESP' then Exit;
  CorrespToCodes(Corresp,TComboBox(CorrespDe),TComboBox(CorrespA));
  CorrespDe.ItemIndex	:= 0 ;
  CorrespA.ItemIndex	:= CorrespA.Items.Count - 1 ;
end;

procedure TOF_CPBALANCE.TableLibreChanged(Sender: TObject);
begin
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 23/01/2003
Modifié le ... :   /  /
Description .. : Limite l'affichage en devise en fonction de la devise
Suite ........ : sélectionnée
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANCE.DeviseChanged(Sender: TObject);
begin
  if Devise.ItemIndex > 0 then
    Affichage.Enabled := True
  else
    begin
    Affichage.Value := 'EURO';
    Affichage.Enabled := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 23/01/2003
Modifié le ... :   /  /
Description .. : Restreint la tablette des comptes en fonction de la nature
Suite ........ : de compte sélectionnée
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANCE.NatureCptChanged(Sender: TObject);
var lStWhere : String ;
begin

  // Cas spécifique pour les sections
  if TypeBal = balAnal then
    begin
    NatureAxeChanged( Sender );
    Exit ;
    end ;

  // Pourr les généraux et les auxiliaires...
  CompteDe.Text := '';
  CompteA.Text := '';
  lStWhere := GetWhereNature ;
  if lStWhere <> '' then
    begin
    CompteDe.Plus := ' AND ' + lStWhere ;
    CompteA.Plus  := ' AND ' + lStWhere ;
    end
  else
    begin
    CompteDe.Plus := '' ;
    CompteA.Plus := '';
    end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 23/01/2003
Modifié le ... : 06/09/2005
Description .. : Utilisé pour les états chaînés
Suite ........ : NE PAS OTER !!!!!!!!!!!!!!!
Suite ........ : (même si pas utilisé en apparence...)
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANCE.FTimerTimer(Sender: TObject);
begin
  if FCritEdtChaine <> nil then
  begin
    with FCritEdtChaine do
    begin
      if CritEdtChaine.UtiliseCritStd then
      begin
        Exercice.Value 	  := CritEdtChaine.Exercice.Code;
        DateComptaDe.Text := DateToStr(CritEdtChaine.Exercice.Deb) ;
        DateComptaA.Text  := DateToStr(CritEdtChaine.Exercice.Fin) ;
      end;
    end;
  end;
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 12/12/2002
Modifié le ... :   /  /
Description .. : Construit la condition sur E_QUALIFPIECE.
Suite ........ : Extrait du RecupWhereCritere pour eviter les ( LIKE OR
Suite ........ : LIKE ) dans les jointures...
Mots clefs ... :
*****************************************************************}
function TOF_CPBALANCE.RecupWhereQualifPiece: String;
Var lStSql   : String ;
    lstTmp   : String ;
begin
  TraductionTHMultiValComboBox( QualifPiece, lStSql , lstTmp, PfEcr + '_QUALIFPIECE', False );
  Result := lStSql ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 23/01/2003
Modifié le ... :   /  /
Description .. : Interdit le mode de selection des compte "Comptes non
Suite ........ : soldés" pour interbase
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANCE.ModeSelectChanged(Sender: TObject);
begin
  if (ModeSelect.value = 'NSL') and (V_PGI.Driver = dbINTRBASE) then
    begin
    { ==== Pour info : liste des drivers de SGDB possibles ===
        Type TDBDriver = (dbINTRBASE,dbMSSQL,dbORACLE7,dbORACLE8,dbDB2,
                          dbINFORMIX,dbMSACCESS,dbPARADOX,dbSQLANY,
                          dbSQLBASE,dbPOL,dbSYBASE,dbMySQL,dbPROGRESS) ; }
    PGIBox ('Cette option n''est pas autorisée sous Interbase', TFQRS1(Ecran).Caption ) ;
    ModeSelect.ItemIndex 	:= 0 ;
    end ;
end;

procedure TOF_CPBALANCE.ComparExoOnChange(Sender: TObject);
begin
  CExoRelatifToDates(ComparExo.Value, ComparDateDe, ComparDateA);
  DateDCompar := StrToDate(ComparDateDe.Text);
  DateFCompar := StrToDate(ComparDateA.Text);
end;

procedure TOF_CPBALANCE.ComparDateOnExit(Sender: TObject);
begin
  DoDateOnExit(THEdit(Sender), ComparDateDe, ComparDateA, DateDCompar, DateFCompar);
end;

procedure TOF_CPBALANCE.ComparTypeClick(Sender: TObject);
begin
  SetControlVisible('COMPARBALSIT',False);
  SetControlVisible('COMPARPERIODE',False);
  if ComparType.value <> '' then
    SetControlVisible(ComparType.value,True);
  //SG6 14.03.05 Gestion des filtres FQ 14996
  if cLoadfiltre then exit;
  if ComparType.Value = 'COMPARPERIODE' then
    if VH^.Precedent.Code <> ''
      then ComparExo.Value := CExerciceVersRelatif (VH^.Precedent.Code)
      else ComparExo.Value := CExerciceVersRelatif (VH^.Entree.Code) ;
end;


procedure TOF_CPBALANCE.DeleteCEDTBALANCE ;
begin
  ExecuteSQL('DELETE FROM ' + GetTablePourBase('CEDTBALANCE') + ' WHERE CED_USER = "' + V_PGI.User + '"');
end;

procedure TOF_CPBALANCE.RemplirEDTBALANCE ;
begin
// 1. Effacement des enregistrements present pour le user
  DeleteCEDTBALANCE ;

// 2. Insertion de la liste des comptes cibles

  // 2.1 : Comptes du comparatifs si besoin (EN PREMIER)
  if not (FTypeBal in [balVentil, balAgee]) then
    if AvecComparatif.Checked then
      begin
      if (ComparType.Value = 'COMPARBALSIT') and (BALSIT.text <> '')
        then ExecuteSQL(GenererInsertCPTBalSit) ;
      if (ComparType.Value = 'COMPARPERIODE') and (ComparExo.Value <> '')
        then ExecuteSQL(GenererInsertBalanceCompar) ;
      end ;

  // 2.2 : Comptes des généraux (COMPLETE les Comptes du compratifs au besoin)
  ExecuteSQL(GenererInsertCPT);

// 3. Update des totaux / soldes
  UpdateCumulsCEDTBALANCE ;

end;


{JP 28/06/05 : Cette fonction est maintenant surchargée dans TOF_CPCUMULPERIODIQUE.PAS}
function TOF_CPBALANCE.GenererRequeteBAL: String;
begin
  // Champs
  result := 'SELECT CED_COMPTE CED_' + ChampCpt + ', CED_NATURE, CED_LIBELLE' ;
  result := result
               + ', CED_DEBIT1, CED_CREDIT1'
               + ', CED_DEBIT2, CED_CREDIT2'
               + ', CED_NATURE CED_AXE'
               + ', CED_RUPTURE, CED_COLLECTIF ';

  if Rupture.Value <> 'SANS' then
    begin
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0)
      then result := result + ', CR_LIBELLE RUPTURELIB '
      else if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0)
        then result := result + ', NT_LIBELLE RUPTURELIB '
        else result := result + ', "&#@" RUPTURELIB ' ;
    end
  else result := result + ', "&#@" RUPTURELIB ' ;

  // Tables
  result := result + ' FROM CEDTBALANCE ' ;
  if Rupture.Value <> 'SANS' then
    begin
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      result := result + ' LEFT JOIN CORRESP ON CR_CORRESP=CED_RUPTURE '
                       + 'AND CR_TYPE="' + Corresp.Value + '"' ;
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      result := result + ' LEFT JOIN NATCPTE ON NT_NATURE=CED_RUPTURE '
                       + 'AND NT_TYPECPTE="' + TableLibre.Value + '"' ;
    end ;

  // Conditions
  result := result + ' WHERE CED_USER = "' + V_PGI.User + '"' ;
  // Tri
  result := result + ' ORDER BY CED_RUPTURE, CED_' + ChampCpt ;

end;


function TOF_CPBALANCE.GetConditionRupture: String;
begin
  result := ' ' ;
  // Conditions sur champ de rupture
  if Rupture.Value <> 'SANS' then
    begin
    // Conditions sur table libre
    if (RuptureType.Value = 'RUPLIBRES') AND (TableLibre.ItemIndex >= 0) then
      begin
      if Trim(LibreDe.Text) <> '' then
        Result := Result  + ' AND ' + XXRupture.Text + ' >= "' + LibreDe.Text + '"';
      if Trim(LibreA.Text) <> '' then
        Result := Result + ' AND ' + XXRupture.Text + ' <= "' + LibreA.Text + '"';
      if CptLibresExist.Checked then
        Result := Result + ' AND ' + XXRupture.Text + '<>""';
      end;
    // Conditions sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
        begin
        if CorrespDe.ItemIndex > 0 then
          Result := Result + ' AND ' + XXRupture.Text + ' >= "' + CorrespDe.Text + '"';
        if CorrespA.ItemIndex < (CorrespA.Items.count - 1) then
          Result := Result + ' AND ' + XXRupture.Text + ' <= "' + CorrespA.Text + '"';
        if CptCorrespExist.Checked then
          Result := Result + ' AND ' + XXRupture.Text + '<>""';
        end;
    end;
end;

procedure TOF_CPBALANCE.CorrespChanged(Sender: TObject);
begin
  if (CorrespDe.ItemIndex > 0) or (CorrespA.ItemIndex < (CorrespA.Items.count - 1)) then
    begin
    CptCorrespExist.Checked := True ;
    CptCorrespExist.Enabled := False ;
    end
  else
    CptCorrespExist.Enabled := True ;
end;

procedure TOF_CPBALANCE.LibreChanged(Sender: TObject);
begin
  if (LibreDe.Text<>'') or (LibreA.Text<>'') then
    begin
    CptLibresExist.Checked := True ;
    CptLibresExist.Enabled := False ;
    end
  else
    CptLibresExist.Enabled := True ;
end;

procedure TOF_CPBALANCE.CompteOnExit(Sender: TObject);
begin
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 07/11/2003
Modifié le ... : 07/11/2003
Description .. : Essaye de faire la complétion auto du numéro de compte
Suite ........ : saisie dans la zone "Sender".
Suite ........ : Retourne True si complétion réalisée, False sinon.
Suite ........ : Fontion générique pour tous les type de compte, reconnue
Suite ........ : grâce au TFichierBase passé en paramètre.
Suite ........ :
Suite ........ : Utilisé dans les fonctions de type OnCompteExit
Suite ........ : implémentées dans les sous-classes.
Mots clefs ... :
*****************************************************************}
function TOF_CPBALANCE.CompleteAuto( Sender: TObject ; CompteType : TFichierBase ): Boolean;
begin
  Result := False ;
  Case CompteType of
    // Cas des compte généraux
    fbGene: begin
            if ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + BourreEtLess(THEdit(Sender).Text, fbGene) + '"') then
              begin
              Result := True ;
              if Length(THEdit(Sender).Text) < VH^.Cpta[fbGene].Lg then
                THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbGene);
              end
            end ;
  // Cas des auxiliaires
    fbAux : begin
            if ExisteSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = "' + BourreEtLess(THEdit(Sender).Text, fbAux) + '"') then
              begin
              Result := True ;
              if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
                THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
              end
            end ;
  // Cas des sections
    fbAxe1..fbAxe5: begin
            if ExisteSql('SELECT S_SECTION FROM SECTION WHERE S_SECTION = "' + BourreEtLess(THEdit(Sender).Text, CompteType) + '"') then
              begin
              Result := True ;
              if Length(THEdit(Sender).Text) < VH^.Cpta[CompteType].Lg then
                THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, CompteType);
              end
            end ;
    end ;
end;

////////////////////////////////////////////////////////////////////////////////
{$IFNDEF GCGC}
procedure TOF_CPBALANCE.OnChangeFiltre(Sender: TObject);
begin
  inherited;

  if not ( FTypeBal in [balAgee,balVentil] ) then
  begin
    CExoRelatifToDates( Exercice.Value, DateComptaDe, DateComptaA, True);
    CExoRelatifToDates( ComparExo.Value, ComparDateDe, ComparDateA, True);
  end;  

  GereAffichageGlobal ;

end;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPBALANCE.QualifPieceEnter(Sender: TObject);
begin
 CSelectionTextControl( Sender ) ;
end;

procedure TOF_CPBALANCE.NouvRechClick(Sender: TObject);
begin

  videFiltre( FFiltres, Pages ) ;
  TFQRS1(Ecran).ListeFiltre.new ;

  // === INIT ZONES
  InitZones ;  

  // === MAJ INTERFACE ===
  GereAffichageGlobal ;

end;

procedure TOF_CPBALANCE.OnAfterFormShow;
begin
  inherited;

  {JP 24/10/05 : FQ 16931 : Comme les champs du QRS1 ne portent pas sur la table de l'état :
                 la requête générée par le RecupWhereCriteres n'est pas valable lorsque l'on
                 utilise les champs avancés (CEDTBLANCE) : on cache l'onglet. D'autre part tous
                 les champs de CEDTBALANCE figurent sur les autres Onglets, donc l'onglet avancés
                 n'est pas utile dans les balances}
  TFQRS1(Ecran).Avances.Visible := False;
  TFQRS1(Ecran).Avances.TabVisible := False;
  TFQRS1(Ecran).CritAvancesVisibled := False;

  TFQRS1(Ecran).ListeFiltre.OnItemNouveau   := NouvRechClick ;
  TFQRS1(Ecran).ListeFiltre.OnItemSupprimer := SupprFiltreClick ;
end;

function TOF_CPBALANCE.GetTablePourBase(vStTable: String): String;
begin
  result := vStTable ;

  // Si la table est partagée, on laisse la main à l'AGL...
  if EstTablePartagee( vStTable ) then Exit ;

  if FStBase <> '' then
    if IsMSSQL  //10/08/2006 YMO Ajout test SQL2005
      then result := FStBase + '.DBO.' + vStTable
      else result := FStBase + '.' + vStTable ;
end;

function TOF_CPBALANCE.GetValeurFiltre(vStZone: String): String;
var lTobFiltre : TOB ;
    lTobZone   : TOB ;
begin
  result := '' ;

  lTobFiltre := TFQRS1(Ecran).ListeFiltre.Current ;
  if lTobFiltre = nil then Exit ;

  lTobZone := lTobFiltre.FindFirst( [ 'N' ] , [ vStZone ] , False ) ;
  if lTobZone = nil then Exit ;

  result := lTobZone.GetString( 'V' ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 19/04/2005
Modifié le ... :   /  /
Description .. : Raffraichit l'affichage de toutes les zones du QRS1
Suite ........ : suite au chargement d'un Filtre par exemple.
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANCE.GereAffichageGlobal;
begin

  // ===================================================
  // === Gestion de l'affichage des pages dynamiques ===
  // ===================================================

  // --> Onglet Rupture
  SetControlVisible('RUPTURES', (Rupture.value <> 'SANS') );

  // --> Onglet Comparatif
  if not ( FTypeBal in [ balVentil, balAgee ] ) then
    SetControlVisible('COMPARATIF', AvecComparatif.checked ) ;

  // ======================================================
  // === Gestion de l'affichage des groupbox dynamiques ===
  // ======================================================
  // --> Onglet Rupture
  SetControlVisible('RUPGROUPES',False);
  SetControlVisible('RUPLIBRES', False);
  SetControlVisible('RUPCORRESP',False);
  if RuptureType.value <> '' then
    SetControlVisible(RuptureType.value,True);

  // --> Onglet Comparatif
  if not ( FTypeBal in [ balVentil, balAgee ] ) then
    begin
    SetControlVisible('COMPARBALSIT',False);
    SetControlVisible('COMPARPERIODE',False);
    if ComparType.value <> '' then
      SetControlVisible(ComparType.value,True);
    end ;

  // ========================================
  // === Gestion des tablettes dynamiques ===
  // ========================================
  // --> Onglet Général
  NatureCptChanged( nil ) ;
  // --> Onglet Rupture
  TableLibreChanged( nil ) ;
  PlanCorrespChanged( nil) ;
  // ... puis remise en place des champs réinit par leur valeur contenu dans le filtre
  if cLoadFiltre or ( FFiltres.Text <> '') then
    begin
    if (CompteDe <> nil) then
      CompteDe.Text  := GetValeurFiltre('COMPTE') ;
    if (CompteA <> nil) then
      CompteA.Text   := GetValeurFiltre('COMPTE_') ;
    if (CorrespDe <> nil) then
      CorrespDe.Text := GetValeurFiltre('CORRESPDE');
    if (CorrespA <> nil) then
      CorrespA.Text  := GetValeurFiltre('CORRESPA');
    if (LibreDe <> nil) then
      LibreDe.Text   := GetValeurFiltre('LIBREDE');
    if (LibreA <> nil) then
      LibreA.Text    := GetValeurFiltre('LIBREA');
    end ;

end;

procedure TOF_CPBALANCE.SupprFiltreClick(Sender: TObject);
begin
  if not TFQRS1(Ecran).ListeFiltre.Delete then Exit ;
  videFiltre( FFiltres, Pages ) ;

  InitZones ;
end;

procedure TOF_CPBALANCE.InitZones;
begin

  if not ( FTypeBal in [balAgee,balVentil] ) then
    begin
    // pages standards
    if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' )
      then Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
      else Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;
    // Mode de sélection
    ModeSelect.ItemIndex   := 0;
    AvecComparatif.Checked := False ;
    end ;

  // Type de compte
  if Assigned( NatureCpt ) then
    NatureCpt.ItemIndex    := 0;

  // Type de pièce
  QualifPiece.value 	 := 'N;';

  Devise.ItemIndex 	 := 0;

  // Options
  Rupture.Value          := 'SANS' ;
  RuptureMax.Checked     := False ;

  GereEtablissement; {JP 29/06/06 : FQ 16149}
end;

procedure TOF_CPBALANCE.RuptureMaxClick(Sender: TObject);
begin

  NivoRupture.Enabled := not RuptureMax.Checked ;

  //SG6 14.03.05 Gestion des filtres FQ 14996
  if cLoadFiltre then Exit;

  if RuptureMax.Checked then
    NivoRupture.Value := 1 ;

end;


function TOF_CPBALANCE.GetNumTableLibre : String ;
begin
  result := '' ;
  if TableLibre.ItemIndex < 0 then Exit ;
  result := Copy( TableLibre.Value, 3, 1 )  ;
end;

procedure TOF_CPBALANCE.MySelectFiltre;
var lZExercice : TZExercice;
    lExoDate : TExoDate;
begin
  inherited;

  if FTypeBal in [balAgee,balVentil] then exit ;

  lZExercice := TZExercice.Create(False);
  QuelDateDeExo(CRelatifVersExercice(Exercice.Value), lExoDate);

  if (lZExercice.QuelExo(DateComptaDe.Text, False) <> lExoDate.Code) and
     (lZExercice.QuelExo(DateComptaA.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(Exercice.Value, DateComptaDe, DateComptaA);
  end;

  // Comparatif
  QuelDateDeExo(CRelatifVersExercice(ComparExo.Value), lExoDate);

  if (lZExercice.QuelExo(ComparDateDe.Text, False) <> lExoDate.Code) and
     (lZExercice.QuelExo(ComparDateA.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(ComparExo.Value, ComparDateDe, ComparDateA);
  end;

  FreeAndNil(lZExercice);
end;

procedure TOF_CPBALANCE.InitComposants;
begin
  // PageControl principal
  Pages                      := TPageControl(GetControl('Pages'));

  // ===========================================
  // ==== Contrôles de l'onglet - Standards ====
  // ===========================================
  CompteDe	             := THEdit(GetControl('COMPTE', true));
  CompteDe.OnExit 	     := CompteOnExit;

  CompteA	             := THEdit(GetControl('COMPTE_', true));
  CompteA.OnExit 	     := CompteOnExit;

  if Getcontrol('NATURECPT', true) is THValComboBox then
    begin
    NatureCpt	             := THValComboBox(Getcontrol('NATURECPT', true)) ;
    NatureCpt.OnChange 	     := NatureCptChanged;
    end
  else NatureCpt             := nil ;

  QualifPiece	             := THMultiValComboBox(GetControl('QUALIFPIECE', true));
  QualifPiece.ReadOnly       := True ;  // Rendre non saisissable la zone des qualifpiece !
  QualifPiece.OnEnter        := QualifPieceEnter ; // Pb visibilité du focus avec le readonly

  if not (FTypeBal in [balAgee,balVentil]) then
    begin
    ModeSelect	             := THValComboBox(Getcontrol('MODESELECT', true));
    ModeSelect.OnChange	     := ModeSelectChanged;

    Exercice := THValComboBox(Getcontrol('EXERCICE', true));
    CInitComboExercice( Exercice ); // Init de la combo Exercice en Relatif
    Exercice.onChange 	     := ExoOnChange;

    DateComptaDe	     := THEdit(GetControl('DATECOMPTABLE', true));
    DateComptaDe.OnExit	     := DateOnExit;

    DateComptaA	             := THEdit(GetControl('DATECOMPTABLE_', true));
    DateComptaA.OnExit 	     := DateOnExit;
    end ;

  // =============================================
  // ==== Contrôles de l'onglet - Compléments ====
  // =============================================
  // Contrôles critères et options - Page avancés
  CptExcept	     := THEdit(GetControl('CPTEXCEPT', true));
  Devise             := THValComboBox(Getcontrol('DEVISE', true));
  Devise.OnChange    := DeviseChanged;

  // ==============================================
  // ==== Contrôles de l'onglet - Mise en page ====
  // ==============================================
  Rupture 	             := THRadioGroup(GetControl('RUPTURE', true));
  Rupture.OnClick	     := RuptureClick;

  // =========================================
  // ==== Contrôles de l'onglet - Options ====
  // =========================================
  Affichage	             := THRadioGroup(GetControl('AFFICHAGE', true));
  Affichage.Items[0]         := '&' + VH^.LibDevisePivot ;   // Affichage de la devise pivot

  if not ( FTypeBal in [balAgee,balVentil] ) then
    begin
    AvecComparatif             := TCheckBox(GetControl('AVECCOMPARATIF', true));
    AvecComparatif.OnClick     := AvecComparatifClick;
    end ;

  SetControlVisible( 'MULTIDOSSIER',   FBoRegroupement ) ;
  SetControlVisible( 'TMULTIDOSSIER',  FBoRegroupement ) ;
  SetControlVisible( 'BVMULTIDOSSIER', FBoRegroupement ) ;
  SetControlVisible( 'RUPTUREMAX',     FBoRegroupement ) ;


  // ==========================================
  // ==== Contrôles de l'onglet - Ruptures ====
  // ==========================================
  RuptureType 	             := THRadioGroup(GetControl('RUPTURETYPE', true));
  RuptureType.OnClick 	     := RuptureTypeClick;

  NivoRupture 	             := TSpinEdit(GetControl('NIVORUPTURE', true));
  RuptureMax                 := TCheckBox(GetControl('RUPTUREMAX', true));
  RuptureMax.OnClick         := RuptureMaxClick;

  TableLibre                 := THValComboBox(Getcontrol('TABLELIBRE', true));
  TableLibre.OnChange  	     := TableLibreChanged;
  LibreDe	             := THEdit(GetControl('LIBREDE', true));
  LibreDe.OnChange           := LibreChanged ;
  LibreA	             := THEdit(GetControl('LIBREA', true));
  LibreA.OnChange            := LibreChanged ;
  CptLibresExist             := TCheckBox(GetControl('CPTLIBRESEXIST', true));


  Corresp	             := THValComboBox(Getcontrol('CORRESP', true));
  Corresp.OnChange  	     := PlanCorrespChanged;
  CptCorrespExist            := TCheckBox(GetControl('CPTCORRESPEXIST', true));
  CorrespDe	             := THValcomboBox(GetControl('CORRESPDE', true));
  CorrespDe.OnChange         := CorrespChanged ;
  CorrespA	             := THValcomboBox(GetControl('CORRESPA', true));
  CorrespA.OnChange          := CorrespChanged ;

  // ============================================
  // ==== Contrôles de l'onglet - COMPARATIF ====
  // ============================================

  if not ( FTypeBal in [balAgee,balVentil] ) then
    begin
    ComparType                 := THRadioGroup(GetControl('COMPARTYPE', true));
    ComparType.OnClick 	       := ComparTypeClick;

    BalSit                     := THEdit(GetControl('BALSIT', true));

    ComparExo	               := THValComboBox(Getcontrol('COMPAREXO', true));
    CInitComboExercice( ComparExo ); // Init de la combo Exercice en Relatif
    ComparExo.onChange         := ComparExoOnChange;

    ComparDateDe	       := THEdit(GetControl('COMPARDATEDE', true));
    ComparDateDe.OnExit        := ComparDateOnExit;

    ComparDateA	               := THEdit(GetControl('COMPARDATEA', true));
    ComparDateA.OnExit         := ComparDateOnExit;
    end ;


  // ========================================
  // ==== Contrôles de gestion INVISIBLE ====
  // ========================================
  AvecNivoRupture            := THEdit(GetControl('AVECNIVORUPTURE', true));
  AvecRupture                := THEdit(GetControl('AVECRUPTURE', true));
  AvecRupType 	             := THEdit(GetControl('AVECRUPTYPE', true));
  XXRupture 	             := THEdit(GetControl('XX_RUPTURE', true));
  XXWhere 	             := THEdit(GetControl('XX_WHERE', true));
  AvecQualifPiece            := THEdit(GetControl('AVECQUALIFPIECE', true));
  AvecAffichage              := THEdit(GetControl('AVECAFFICHAGE', true));


  // ===========================
  // ==== Autres Evènements ====
  // ===========================

  // GCO - 08/07/2004 - FQ 13484
  FTimer.OnTimer             := FTimerTimer;
end;

procedure TOF_CPBALANCE.ParamTitresCols;
var  Q : TQuery;
begin
  if AvecComparatif.Checked then
    begin     // Cas du comparatif
    SetControlText('TitreCol1','Solde du ' + DateToStr(DateDebutPer) + ' au ' + DateToStr(DateFinPer)) ;
    if ComparType.Value = 'COMPARBALSIT' then   // BALANCE DE SITUATION
      begin
      Q := OpenSQL('SELECT BSI_ABREGE FROM CBALSIT WHERE BSI_CODEBAL="'+BalSit.Text+'"', True);
      if not Q.Eof
        then SetControlText('TitreCol2',Q.FindField('BSI_ABREGE').AsString)
        else SetControlText('TitreCol2','Balance de situation');
      Ferme(Q);
  // En attente du portage de GetColonneSQL en eAGL
  //  	SetControlText('TitreCol1',GetColonneSQL('CBALSIT','BSI_ABREGE','BSI_CODEBAL="'+BalSit.Text+'"'));
      end
    else                                // COMPARAISON SUR PERIODE
      begin
      SetControlText('TitreCol2','Solde du ' + DateToStr(DateDebutCompar) + ' au ' + DateToStr(DateFinCompar));
      end ;
    end
  else
    begin    // Cas des totaux
    SetControlText('TitreCol1','Total du ' + DateToStr(DateDebutPer) + ' au ' + DateToStr(DateFinPer)) ;
    SetControlText('TitreCol2','Solde du ' + DateToStr(DateDebutPer) + ' au ' + DateToStr(DateFinPer)) ;
    end ;
end;

procedure TOF_CPBALANCE.ParamComparatif;
begin
  DateDebutPer          := StrToDate(DateComptaDe.Text);
  DateFinPer            := StrToDate(DateComptaA.Text);
  if AvecComparatif.Checked and (ComparType.Value = 'COMPARPERIODE') then
    begin
    DateDebutCompar     := StrToDate(ComparDateDe.Text);
    DateFinCompar       := StrToDate(ComparDateA.Text);
    end ;
end;

procedure TOF_CPBALANCE.ParamRuptures;
begin
  XXWhere.Text          := '';
  XXRupture.Text        := '';
  AvecNivoRupture.Text  := '0';
  AvecRupture.Text      := Rupture.Value ;  // FQ 19056 SBO 30/10/2006

  if Rupture.Value = 'SANS'
    then AvecRupType.Text      := ''
    else
      begin
        AvecRupType.Text := RuptureType.value;

        // rupture sur n° compte
        if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.value > 0) then
          AvecNivoRupture.Text := IntToStr(NivoRupture.value);

        // rupture sur champ libre tiers
        if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
          XXRupture.Text := PfCpt + '_TABLE' + GetNumTableLibre ;

        // rupture sur plan de correspondance
        if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
          XXRupture.Text := PfCpt + '_CORRESP' + IntToStr(Corresp.ItemIndex + 1);

      end;

end;

procedure TOF_CPBALANCE.ParamDivers;
var stTypeEcr : String;
begin

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(CompteDe.Text) then begin
    // Auto-remplissage si comptes non renseignés
    if (Trim(CompteDe.Text) = '') then
      CompteDe.Text := GetMinCompte( TableCpt, PfCpt + '_' + ChampCpt, '', GetWhereNature );
    if (Trim(CompteA.Text) = '') then
      CompteA.Text  := GetMaxCompte( TableCpt, PfCpt + '_' + ChampCpt, '', GetWhereNature );
  end;

  // libellés pour les types d'écritures...
  AvecQualifPiece.Text := '';
  stTypeEcr := QualifPiece.Value;
  if stTypeEcr <> '' then
    while stTypeEcr <> '' do
      AvecQualifPiece.Text := AvecQualifPiece.Text
      			      + RechDom(QualifPiece.DataType,ReadTokenSt(stTypeEcr),False)
                              + ' '
  else
    AvecQualifPiece.Text := '<<Tous>>';

  // Affichage devise...
  if Affichage.Value = 'EURO'
    then AvecAffichage.Text := Devise.Text + ' / Affichage ' + VH^.LibDevisePivot
    else AvecAffichage.Text := Devise.Text + ' / Affichage ' + Devise.Items[Devise.ItemIndex] ;

end;

function TOF_CPBALANCE.GetConditionSQL( vBoAvecCpt : Boolean; vInPer : Integer ) : String;
var lStWhereDate : String ;
begin

  result := '' ;

  // =====================
  // ==== LES COMPTES ====
  // =====================
  if vBoAvecCpt then
    begin
    result := GetConditionSQLCpt ;
    if not (FTypeBal in [ balAgee, balVentil ]) and (ModeSelect.value = 'FOU') then Exit ;
    result := result + ' AND ' ;
    end ;

  // ============================================
  // ==== Conditions sur la table principale ====
  // ============================================
  result := result + GetConditionSQLEcr ;

  // Critères Dates Extériosiés pour gestion des périodes
  if vBoAvecCpt
    then lStWhereDate := GetConditionSQLDateC( vInPer )
    else lStWhereDate := GetConditionSQLPeriode( vInPer ) ;
  if lStWhereDate <> '' then
    result := result + ' AND ' + lStWhereDate ;
end;

{JP 03/10/07 : FQ 21560 : depuis CPTOTRUB_TOF, on n'a pas une fourchette de comptes, mais
               une clause where avec des LIKE, BETWEEN. Cette fonction surchargée et appelée 
               dans GetConditionSQLCpt permet de courcircuiter le traitement standard}
{---------------------------------------------------------------------------------------}
function TOF_CPBALANCE.GetConditionSQLCptSpecif : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
end;

function TOF_CPBALANCE.GetConditionSQLCpt: String;
var listeCptExcept : String ;
    lStWhereNature : String ;
    stCpt          : String ;
begin

  // ===================
  // ==== LES TIERS ====
  // ===================
  // Comptes...
  {JP 03/10/07 : FQ 21560 : gestion d'éventuels cas spécifiques}
  Result := GetConditionSQLCptSpecif;

  {FQ 21560 : Si pas de gestion d'une clause where spécifique, on reprend le traitement standard}
  if Trim(Result) = '' then begin
    result := ' ' + ConvertitCaractereJokers(CompteDe, CompteA, PfCpt + '_' + ChampCpt );

    // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
    if CptExcept.Text <> '' then
      begin
      listeCptExcept := FindEtReplace(CptExcept.Text,',',';', True);
      while (listeCptExcept <> '') do
        begin
        stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
        if stCpt <> '' then
          result := result + ' AND ' + PfCpt + '_' + ChampCpt + ' NOT LIKE "' + stCpt + '%"';
        end ;
      end;
   end;

  // Nature
  lStWhereNature := GetWhereNature ;
  if lStWhereNature <> '' then
    result := result + ' AND ' + lStWhereNature ;

  // Rupture
  result := result + GetConditionRupture ;
  
end;


function TOF_CPBALANCE.GetConditionSQLEcr: String;
begin

  // ============================================================
  // === CONDITIONS AUTOMATIQUES : E_ETABLISSEMENT, E_DEVISE ====
  // ============================================================
  if ComboEtab.Value <> '' then
    Result := Result + ' ' + PfEcr + '_ETABLISSEMENT = "' + ComboEtab.Value + '" AND ';

  if Devise.Value <> '' then
    Result := Result + ' ' + PfEcr + '_DEVISE = "' + Devise.Value + '" AND ';

  if TypeBal=balAnal then {YMO 17/01/07 prise en compte de l'axe concerné uniqt FQ19289}
    result := result + ' ' + PfEcr + '_AXE ="' + NatureCpt.value + '" AND' ;

  // ============================
  // ==== LES TYPES DE PIECES ===
  // ============================
  Result := Result + RecupWhereQualifPiece ;

end;

function TOF_CPBALANCE.GetConditionSQLPeriode(vInPer: Integer): String;
begin

  result := '' ;
  if FTypeBal in [ balAgee, balVentil ] then Exit ;

  Case vInPer of

      // période principale (exo + date)
      1 : result := ' ' + FStPfEcr + '_EXERCICE ="' + CRelatifVersExercice(Exercice.Value) + '" AND '
                        + FStPfEcr + '_DATECOMPTABLE BETWEEN "' + USDateTime( DateDebutPer )
                                             + '" AND "'     + USDateTime( DateFinPer ) + '"' ;

      // Période optionnelle
      2 : result := ' ' + FStPfEcr + '_EXERCICE ="' + CRelatifVersExercice(ComparExo.Value) + '" AND '
                        + FStPfEcr + '_DATECOMPTABLE BETWEEN "' + USDateTime( DateDebutCompar )
                                             + '" AND "'     + USDateTime( DateFinCompar ) + '"' ;

      end ;

end;

function TOF_CPBALANCE.GetSQLCumul(vInPer, vInCol: Integer): String;
begin
  // --> Champs
  { FQ 20923 BVE 04.07.07 }
  Case vInCol of
    // Crédit + débit
    0 : if Affichage.Value = 'EURO'
          then result := 'SELECT SUM(' + FStPfEcr + '_DEBIT) D, SUM(' + FStPfEcr + '_CREDIT) C'
          else result := 'SELECT SUM(' + FStPfEcr + '_DEBITDEV) D, SUM(' + FStPfEcr + '_CREDITDEV) C';

    // Débit uniquement
    1 : if Affichage.Value = 'EURO'
          then result := 'SELECT SUM(' + FStPfEcr + '_DEBIT)'
          else result := 'SELECT SUM(' + FStPfEcr + '_DEBITDEV)';

    // Crédit uniquement
    2 : if Affichage.Value = 'EURO'
          then result := 'SELECT SUM(' + FStPfEcr + '_CREDIT)'
          else result := 'SELECT SUM(' + FStPfEcr + '_CREDITDEV)';
  end ;
  // --> Table
  result := result + ' FROM ' + GetTablePourBase(FStTableEcr) ;

  // --> Conditions
  result := result + ' WHERE ' + GetConditionSQL( False, vInPer ) ;

  // ---> Lien Compte // champ CED_COMPTE
  result := result + ' AND ' + PfEcr + '_' + ChampCpt + ' = CED_COMPTE ' ;

end;

function TOF_CPBALANCE.GetConditionSQLDateC(vInPer: Integer): String;
var lQExo  : TQuery ;
    lStExo : String ;
begin

  result := '' ;
  if (FTypeBal in [balAgee,balVentil]) then Exit ;

{b fb 22/05/2006 FQ17797}
//  if ( ModeSelect.value = 'EXO' ) or ( ModeSelect.value = 'PER' ) then
//    result := ' ' + PfEcr + '_EXERCICE ="' + CRelatifVersExercice(Exercice.Value) + '" AND ' ;

  if ( ModeSelect.value = 'EXO' ) or ( ModeSelect.value = 'PER' ) then
  { FQ 21348 BVE 07.09.07
    if Exercice.Value<>'' then
      result := ' ' + PfEcr + '_EXERCICE ="' + CRelatifVersExercice(Exercice.Value) + '" '; }
    if vInPer = 2
      then result := ' ' + PfEcr + '_EXERCICE ="' + CRelatifVersExercice( ComparExo.value ) + '" '
      else result := ' ' + PfEcr + '_EXERCICE ="' + CRelatifVersExercice( Exercice.Value ) + '" ';
  { END FQ 21348 }
{e fb 22/05/2006 17797}

  // ====================================
  // ==== Mouvementés sur l'exercice ====
  // ====================================
  if ModeSelect.value = 'EXO' then
    begin
    if vInPer = 2
      then lStExo := CRelatifVersExercice( ComparExo.value )
      else lStExo := CRelatifVersExercice( Exercice.value ) ;
    lQExo:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="' + lStExo +'"' , True ) ;
    if Not lQExo.EOF then
      begin
      result := result + ' AND ' + PfEcr + '_DATECOMPTABLE>="' + USDateTime(lQExo.FindField('EX_DATEDEBUT').asDateTime)
      	     	      + '" AND ' + PfEcr + '_DATECOMPTABLE<="' + USDateTime(lQExo.FindField('EX_DATEFIN').asDateTime) + '"';
      end;
    Ferme(lQExo) ;
    end

  // ====================================
  // ==== Mouvementés sur la période ====
  // ====================================
  else if (ModeSelect.value = 'PER') or (ModeSelect.value = 'NSL') then
    result := GetConditionSQLPeriode( vInPer ) ;

end;

procedure TOF_CPBALANCE.UpdateCumulsCEDTBALANCE;
Var lStReq      : String ;
begin

  try  // finally
    try  // except

      // -----------------------------------
      // ---- Calcul période principale ----
      // -----------------------------------
      // >> Champ Débit
      lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = ( ' + GetSQLCumul( 1, 1 ) + ' )'
                                                            + ' WHERE CED_USER="'+V_PGI.User+'"' ;
      ExecuteSQL( lStReq ) ;
      // Le problème de la requête précédente est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
      ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = 0 WHERE CED_DEBIT1 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;

      // >> Champ Crédit
      lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = ( ' + GetSQLCumul( 1, 2 ) + ' )'
                                                            + ' WHERE CED_USER="'+V_PGI.User+'"' ;
      ExecuteSQL( lStReq ) ;
      // Le problème des requêtes précédentes est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
      ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = 0 WHERE CED_CREDIT1 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;


      // ------------------------------------
      // ---- Calcul période optionnelle ----
      // ------------------------------------
      if AvecComparatif.checked and (ComparType.Value = 'COMPARPERIODE') and (ComparExo.Value <> '') then
        begin
        // >> Champ Débit 2
        lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = ( ' + GetSQLCumul( 2, 1 ) + ' )'
                                                              + ' WHERE CED_USER="'+V_PGI.User+'"' ;
        ExecuteSQL( lStReq ) ;
        // Le problème des requêtes précédentes est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
        ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = 0 WHERE CED_DEBIT2 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;

        // >> Champ Crédit 2
        lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT2 = ( ' + GetSQLCumul( 2, 2 ) + ' )'
                                                              + ' WHERE CED_USER="'+V_PGI.User+'"' ;
        ExecuteSQL( lStReq ) ;
        // Le problème des requêtes précédentes est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
        ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT2 = 0 WHERE CED_CREDIT2 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;

        end ;

      // ---------------------------------------------------------------------
      // ---- En cas de comparatif, les colonnes sont présentées en solde ----
      // ---------------------------------------------------------------------
      // Crédit 1 et Débit 1 --> tout le temps
      if AvecComparatif.checked then
          begin
          // >> Maj des soldes débiteurs
          lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = (CED_DEBIT1-CED_CREDIT1), CED_CREDIT1=0 '
                  + ' WHERE ( CED_DEBIT1 >= CED_CREDIT1 ) AND CED_USER="'+V_PGI.User+'"' ;
          ExecuteSQL( lStReq ) ;
          // >> Maj des soldes créditeurs
          lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = 0, CED_CREDIT1 = (CED_CREDIT1-CED_DEBIT1) '
                  + ' WHERE ( CED_DEBIT1 < CED_CREDIT1 ) AND CED_USER="'+V_PGI.User+'"' ;
          ExecuteSQL( lStReq ) ;

          // Crédit 2 et Débit 2 --> uniquement pour les comparatifs sur périodes
          if (ComparType.Value = 'COMPARPERIODE') and (ComparExo.Value <> '') then
              begin
              // >> Maj des soldes débiteurs
              lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = (CED_DEBIT2-CED_CREDIT2), CED_CREDIT2=0 '
                      + ' WHERE ( CED_DEBIT2 >= CED_CREDIT2 ) AND CED_USER="'+V_PGI.User+'"' ;
              ExecuteSQL( lStReq ) ;
              // >> Maj des soldes créditeurs
              lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = 0, CED_CREDIT2 = (CED_CREDIT2-CED_DEBIT2) '
                      + ' WHERE ( CED_DEBIT2 < CED_CREDIT2 ) AND CED_USER="'+V_PGI.User+'"' ;
              ExecuteSQL( lStReq ) ;
              end ;
          end
      else
          begin
          // >> Maj des soldes débiteurs
          lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = (CED_DEBIT1-CED_CREDIT1), CED_CREDIT2=0 '
                  + ' WHERE ( CED_DEBIT1 >= CED_CREDIT1 ) AND CED_USER="'+V_PGI.User+'"' ;
          ExecuteSQL( lStReq ) ;
          // >> Maj des soldes créditeurs
          lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = 0, CED_CREDIT2 = (CED_CREDIT1-CED_DEBIT1) '
                  + ' WHERE ( CED_DEBIT1 < CED_CREDIT1 ) AND CED_USER="'+V_PGI.User+'"' ;
          ExecuteSQL( lStReq ) ;
          end ;

    // Try Except
    except
      on E : Exception do
      begin
        PgiError( E.Message, Ecran.Caption );
      end;
    end;
  // Try Finally
  finally
//    FiniMoveProgressForm ;
  end ;

end;

function TOF_CPBALANCE.GenererInsertCPT( vBoCompar : Boolean ): String;
begin
  // ===============================================
  // ==== DEBUT DE LA CLAUSE INSERT des comptes ====
  // ===============================================
  result := GetBaseInsertCpt ;

  // ============================
  // ==== SELECT des comptes ====
  // ============================
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", ' + PfCpt + '_' + ChampCpt + ', ' ;
  if ChampNatCpt <> ''
    then result := result + ChampNatCpt + ', ' + PfCpt + '_LIBELLE'
    else result := result + '"", ' + PfCpt + '_LIBELLE';
  // Montants
  result := result + ',0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text
    else result := result + ', ""';
  // Collectif, Compte et libelle combinatoire
  result := result + ', ' + GetCollectif + ', ' + GetCompte2 + ', ' + GetLibelle2 ;

  // ================
  // ==== TABLES ====
  // ================
  result := result + ' FROM ' + GetTablePourBase( TableCpt );

  // Ajout Jointure si besoin
  result := result + GetJointureCpt ;

  // Table Ecriture si besoin
  if ( TypeBal in [balAgee, balVentil] ) or ( ModeSelect.value <> 'FOU') then
    result := result + ' LEFT JOIN ' + GetTablePourBase(TableEcr) + ' ON ' + PfEcr + '_' + ChampCpt + '=' + PfCpt + '_' + ChampCpt ;

  // ===============
  // ==== WHERE ====
  // ===============
  if vBoCompar
    then result := result + ' WHERE ' + GetConditionSQL( True, 2 )
    else result := result + ' WHERE ' + GetConditionSQL( True, 1 ) ;

  // SI COMPARATIF, on se contente d'ajouter les comptes non présent dans CEDTBALANCE
  if not ( TypeBal in [balAgee, balVentil] ) then
    if AvecComparatif.Checked and not vBoCompar then
      begin
      Result := result + ' AND (' + PfCpt + '_' + ChampCpt + ' not in (SELECT CED_COMPTE FROM ' + GetTablePourBase('CEDTBALANCE')
                                                 + ' WHERE CED_USER = "'+V_PGI.User+'"))'
      end ;

  { FQ 19388 BVE 06.06.07 }
  // Gestion des confidentiels
  result := result + ' AND ' + CGenereSQLConfidentiel(PfCpt);
  { END FQ 19388 }     

  // Group by
  result := result + GetGroupBy ;

  // Mode de sélection non soldé
  if ( TypeBal in [balAgee, balVentil] ) or (ModeSelect.value = 'NSL') then
    Result := Result + ' HAVING SUM(' + PfEcr + '_CREDIT) <> SUM(' + PfEcr + '_DEBIT)';

end;

function TOF_CPBALANCE.GetBaseInsertCpt: String;
begin
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_NATURE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ',CED_CREDIT3, CED_DEBIT3, CED_CREDIT4, CED_DEBIT4';
  result := result + ',CED_CREDIT5, CED_DEBIT5, CED_CREDIT6, CED_DEBIT6';
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
end;

function TOF_CPBALANCE.GetJointureCpt: String;
begin
  result := '' ;
end;

procedure TOF_CPBALANCE.SetTypeBalance;
begin
  FTypeBal    := balGene ;
  TableEcr    := 'ECRITURE' ;
  PfEcr       := 'E' ;
  TableCpt    := 'GENERAUX' ;
  PfCpt       := 'G' ;
  ChampCpt    := 'GENERAL' ;
  ChampNatCpt := 'G_NATUREGENE' ;
end;

function TOF_CPBALANCE.GetWhereNature: String;
var lMinIndex : Integer ;
begin

  result := '' ;
  if ChampNatCpt = '' then Exit ;

  // FQ 18190 : En bal. analytique, la valeur tous n'existe pas,
  //            ce qui entraînait l'oublie de la condition sur l'AXE
  if NatureCpt.Vide
    then lMinIndex := 1
    else lMinIndex := 0 ;
  if NatureCpt.ItemIndex >= lMinIndex
    then result := result + ' ' + ChampNatCpt + '="' + NatureCpt.value + '" ' ;

end;

function TOF_CPBALANCE.GenererInsertBalanceCompar: String;
begin
  result := GenererInsertCPT( True ) ;
end;

function TOF_CPBALANCE.GenererInsertCPTBalSit: String;
//var lStWhereNature : String ;
begin
  // ===============================================
  // ==== DEBUT DE LA CLAUSE INSERT des comptes ====
  // ===============================================
  result := GetBaseInsertCpt ;

  // ============================
  // ==== SELECT des comptes ====
  // ============================
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", BSE_COMPTE1, ' ;
  if ChampNatCpt <> ''
    then result := result + ChampNatCpt + ', ' + PfCpt + '_LIBELLE'
    else result := result + '"", ' + PfCpt + '_LIBELLE';
  // --> Alimentation des montants
  if Affichage.Value = 'EURO'
    then result := result + ',0, 0, BSE_CREDIT, BSE_DEBIT, 0, 0, 0, 0, 0, 0, 0, 0'         // montants Euro
    else result := result + ',0, 0, BSE_CREDITDEV, BSE_DEBITDEV, 0, 0, 0, 0, 0, 0, 0, 0';  // montants Devise
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text
    else result := result + ', ""';
  // Collectif, Compte et libelle combinatoire
  result := result + ', ' + GetCollectif + ', ' + GetCompte2 + ', ' + GetLibelle2 ;

  // ================
  // ==== TABLES ====
  // ================
  result := result + ' FROM CBALSITECR ' ;
  result := result + ' LEFT JOIN ' + GetTablePourBase(TableCpt) + ' ON BSE_COMPTE1 = ' + PfCpt + '_' + ChampCpt;

  // Ajout Jointure si besoin
  result := result + GetJointureCpt ;

  // ===============
  // ==== WHERE ====
  // ===============
  result := result + ' WHERE BSE_CODEBAL="' + BALSIT.text + '"';
//  Result := Result + ' AND ' + ConvertitCaractereJokers(CompteDe, CompteA, 'BSE_COMPTE1');
  Result := Result + ' AND ' + GetConditionSQLCpt ;

  // Nature Auxi ?
(*  lStWhereNature := GetWhereNature ;
  if lStWhereNature <> '' then
    result := result + ' AND ' + lStWhereNature ;
*)
  // Seulement les comptes mouvementés ?
  if ModeSelect.Value = 'NSL' then
    result := result + ' AND (BSE_CREDIT<>0 OR BSE_DEBIT<>0)';

  // *> Conditions sur champ de rupture
  result := result + GetConditionRupture ;
  // *> Fin des conditions sur champ de rupture

end;

function TOF_CPBALANCE.GetNumAxe: String;
begin
  result := '' ;
  if NatureCpt.ItemIndex < 0 then Exit ;
  result := copy( NatureCpt.Value, 2, 1 ) ;
end;

procedure TOF_CPBALANCE.NatureAxeChanged(Sender: TObject);
var stAxe  : string;
    lTFB   : TFichierBase;
begin

  // ---------------------
  // Au changement d'axe :
  // ---------------------
  if NatureCpt.ItemIndex < 0 then Exit;
  stAxe := GetNumAxe ;

  // 1. Type de section
  CompteDe.Text := '';
  CompteA.Text := '';
  if stAxe = '1' then
  begin
    CompteDe.DataType := 'TZSECTION';
    CompteA.DataType := 'TZSECTION';
  end
  else
  begin
    CompteDe.DataType := 'TZSECTION' + stAxe;
    CompteA.DataType := 'TZSECTION' + stAxe;
  end;

  // 2. Type de Plan de correspondance
  case V_PGI.LaSerie of
    S7:
      begin
        Corresp.plus := 'AND (CO_CODE = "A' + stAxe + '1" OR CO_CODE = "A' + stAxe + '2")';
      end;
  else
    begin
      Corresp.plus := 'AND CO_CODE = "A' + stAxe + '1"';
    end;
  end;

  // 3. Longueur des sections
  lTFB := AxeToFb('A' + stAxe);
  CompteDe.MaxLength := VH^.Cpta[lTFB].Lg;
  CompteA.MaxLength := VH^.Cpta[lTFB].Lg;

end;

function TOF_CPBALANCE.GetCompte2: string;
begin
  result := '""' ;
end;

function TOF_CPBALANCE.GetLibelle2: string;
begin
  result := '""' ;
end;

function TOF_CPBALANCE.GetGroupBy: string;
begin
  if ChampNatCpt <> ''
    then  result := ' GROUP BY ' + PfCpt + '_' + ChampCpt + ', ' + ChampNatCpt + ', ' + PfCpt + '_LIBELLE'
    else  result := ' GROUP BY ' + PfCpt + '_' + ChampCpt + ', ' + PfCpt + '_LIBELLE' ;
  if XXRupture.Text <> '' then
    result := result + ', ' + XXRupture.Text;
end;

function TOF_CPBALANCE.GetCollectif: string;
begin
result := '"-"' ;
end;

Initialization
  registerclasses ( [ TOF_CPBALANCE ] ) ;
end.

