{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 29/07/2003
Modifié le ... : 31/07/2003
Description .. : Source TOF de la FICHE : CPBALANCECOMBI
Suite ........ :
Suite ........ : Source commun pour QRS1 des balances combinatories :
Suite ........ : TOF_METH --> TOF_CPBALANCECOMBI --> xxx
Mots clefs ... : TOF;CPBALALANCECOMBI
*****************************************************************}
Unit CPBALANCECOMBI_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl, eQRS1,
{$ELSE}
     Fe_Main,
     db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,
{$ENDIF}
     sysutils, ComCtrls, Spin, Ent1,
     Htb97,
     Filtre,
     HCtrls, HEnt1, HMsgBox, UTOF, hqry,
     TofMeth,
     uTob,
     uLibExercice,       // CInitComboExercice(E_EXERCICE)
     ed_tools,           // barre de progression
     hstatus,            // barre de progression inside
     uMultiDossierUtil,  // RequeteMultiDossier
     utilPGI,            // EstTablePartagee
     uLibWindows,        // TraductionTHMultiValComboBox
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     Paramsoc   ,UentCommun;

Type
  TOF_CPBALANCECOMBI = Class (TOF_Meth)

    Pages : TPageControl;

    // ---------------
    // Onglet standard
    CompteDe, CompteA           : THEdit;
    Compte2De, Compte2A         : THEdit;
    DateComptaDe, DateComptaA   : THEdit;
    QualifPiece                 : THMultiValcomboBox;
    Exercice                    : THValComboBox;
    ModeSelect                  : THValComboBox;
    NatureCpt                   : THValComboBox;
    // ---------------
    // Onglet Complément
    Devise                      : THValcomboBox;
    CptExcept                   : THEdit;
    CptExcept2                  : THEdit;
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
    NivoRupture : TSpinEdit;
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

    // EVT TOF
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure ChargementCritEdt        ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnAfterFormShow          ; override ;

    // EVT FICHE
    // -> Onglet standard
    procedure CompteOnExit(Sender: TObject) ; virtual ;
    procedure Compte2OnExit(Sender: TObject) ; virtual ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure ModeSelectChanged(Sender : TObject);
    procedure DateOnExit(Sender: TObject) ;
    procedure QualifPieceEnter(Sender: TObject) ;
    // -> Onglet Compléments
    procedure DeviseChanged(Sender : TObject);
    // -> Onglet avancés
    procedure AvecComparatifClick(Sender: TObject);
    procedure ComparTypeClick(Sender : TObject);
    // -> Onglet mise en page
    procedure RuptureClick(Sender : TObject);
    // -> Onglet rupture
    procedure PlanCorrespChanged(Sender : TObject);
    procedure RuptureTypeClick(Sender : TObject);
    procedure CorrespChanged(Sender : TObject);
    procedure LibreChanged(Sender : TObject);
    // -> Onglet comparatif
    procedure ComparExoOnChange(Sender: TObject) ;
    procedure ComparDateOnExit(Sender: TObject) ;
    // -> Filtre
    procedure InitZones ;                     virtual ;
    procedure NouvRechClick(Sender:TObject) ; virtual ;
    procedure SupprFiltreClick(Sender:TObject) ;

    procedure AuxiElipsisClick      (Sender: TObject) ;

    // TRAITEMENT
    // -> Méthodes de constitution de la table temporaire
    // YMO 22/09/2006 FQ18810 méthode à surcharger dans CPBALAUXIGEN,CPBALGENAUXI,CPBALANALGEN et CPBALGENANAL
    procedure RemplirEDTBALANCE ; virtual ;
    procedure DeleteCEDTBALANCE ;
    procedure UpdateCumulsCEDTBALANCE ;  
    // -> Construit la requête de l'état
    function  GenererRequeteBAL : String;
    function  GetNomChampCPT : String ; virtual ;
    function  GetNomChampCPTCombi : String ; virtual ;
    function  GetConditionRupture : String ;
    // -> Calcul du solde à partir d'un débit / crédit
    procedure CalculSolde( vDebit, vCredit : Double ; var vSoldeDebit, vSoldeCredit : Double ) ;
    // -> Pour états chaînés NE PAS VIRER !!!!!!!!
    procedure FTimerTimer(Sender: TObject);
    // -> Construction des conditions du WHERE pour les cas spéciaux
    function  RecupWhereQualifPiece : String ;   virtual; // pour éviter les 'like 'xx%' or like...)
    function  RecupWhereBalance( vBoAvecDate : Boolean = True ): String; virtual;
    function  RecupWhereBalanceCompar( vBoAvecDate : Boolean = True ) : String ; virtual; // Condition sur la période comparée

    // EVT A REDEFINIR
    procedure NatureCptChanged  (Sender : TObject);    virtual;
    procedure TableLibreChanged (Sender : TObject);    virtual;
    procedure AxeChanged        (Sender : TObject) ;   virtual;


    // TRAITEMENTS A REDEFINIR
    // -> Procedure d'auto complétion des comptes
    Function CompleteAuto( Sender : TObject ; CompteType : TFichierBase ) : Boolean ; virtual;

    // -> Procedure de calcul des totaux / soldes pour alimentation table temporaire
    procedure CalculDebitCredit( vStCompte, vStCompte2: String ; vBoCompar : Boolean ; var vDebit, vCredit: Double );  virtual;

    // -> Procédure d'insertion des comptes cibles
    procedure InsererComptes ;             
    procedure InsererComptesComparatif ;   virtual;
    function  GenererRequeteComptes       : String ;      virtual; abstract ;
    function  GenererRequeteComptesCompar : String ;      virtual; abstract ;
    function  GenererRequeteCombis ( vStCode, vStLib, vStRup : String )       : String ; virtual; abstract ;
    function  GenererRequeteCombisCompar ( vStCode, vStLib, vStRup : String ) : String ; virtual; abstract ;
    function  GenererRequeteCombisBalsit ( vStCode, vStLib, vStRup : String ) : String ; virtual; abstract ;
    function  GenererRequeteVide ( vStCode, vStLib, vStRup : String )         : String ;

    // Récup des min max compte géné ventilables
    function  GetMinCompteVentil( vStNature, vStAxe : string ) : String ;
    function  GetMaxCompteVentil( vStNature, vStAxe : string ) : String ;
    function  GetNumAxe : String ; virtual ;
    function  GetNumTableLibre : String ; override ;

    {$IFNDEF GCGC}
    procedure OnChangeFiltre ( Sender : TObject ) ; override;
    {$ENDIF}

    // Gestion Regroupement
    function  GetTablePourBase( vStTable : String ) : String ;

    //Gestion des filtres :
    function  GetValeurFiltre( vStZone : String ) : String ;
    procedure GereAffichageGlobal ;
    procedure MySelectFiltre; override;


  public
    // Gestion Regroupement
    FBoRegroupement             : Boolean ;
    FStBase                     : String ;

    // Variable gestion des dates
    DateDebutPer                : TDatetime;  // Date début de la période
    DateFinPer                  : TDatetime;  // Date de fin de la période
    DateDebutCompar             : TDatetime;  // Date début de la période comparative
    DateFinCompar               : TDatetime;  // Date de fin de la période comparative
    DateD                       : TDatetime;  // Date début d'exercice
    DateF                       : TDatetime;  // Date fin d'exercice
    DateDCompar                 : TDatetime;  // Date début d'exercice comparatif
    DateFCompar                 : TDatetime;  // Date fin d'exercice comparatif
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  AGLInit,      // TheData
  CritEdt,      // ClassCritEdt
  UTofMulParamGen;

{TOF_CPBALANCECOMBI}

procedure TOF_CPBALANCECOMBI.OnNew ;
begin
  Inherited ;

  // Init Exercice si combo vide..Ne devrait pas arriver ! (Pour corriger un ancien dysfontionnement)
{  if Exercice.Value = '' then
    begin
    // pages standards
    if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' )
      then Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
      else Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;
    ExoOnChange(nil);
    end ;
}

  // Initialisation des champs si pas de filtres
  if FFiltres.Text = '' then
    NouvRechClick( nil ) ;
end ;

procedure TOF_CPBALANCECOMBI.OnUpdate ;
var lStListeBases : String ;
    lStTitre      : String ;
begin
  Inherited ;

  // Ecran de patience
  lStTitre      := Ecran.Caption ;
  Ecran.Caption := lStTitre + '  -  Préparation de la balance en cours, veuillez patienter...';
  UpdateCaption( Ecran ) ;
  SourisSablier ;
  InitMove( 100 , TraduireMemoire('Préparation de la balance en cours...') ) ;

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
  FiniMove ;
  Ecran.Caption := lStTitre ;
  UpdateCaption( Ecran ) ;
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
procedure TOF_CPBALANCECOMBI.ChargementCritEdt;
begin
  inherited;

  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    // Comptes
    CompteDe.Text := ClassCritEdt(TheData).CritEdt.Cpt1;
    CompteA.Text  := ClassCritEdt(TheData).CritEdt.Cpt2;
{b fb 24/05/2006 FQ17691}
   //Comptes 2
    Compte2De.Text := ClassCritEdt(TheData).CritEdt.sCpt1;
    Compte2A.Text := ClassCritEdt(TheData).CritEdt.sCpt2;
{e fb 24/05/2006 FQ17691}
    // Exercice
    Exercice.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);
    // Date de Début de l'édition
    DateComptaDe.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
    // Date de Fin de l'édition
    DateComptaA.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);
    // Qualifpièce
    QualifPiece.Text := ClassCritEdt(TheData).CritEdt.QualifPiece;
    // Etablissement
    ComboEtab.Value := ClassCritEdt(TheData).CritEdt.Etab;
    // Devise
    Devise.Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;
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

        if GetNomChampCPTCombi = 'CED_GENERAL' then
          NatureCpt.Value := NatureCompte ;

          // Exercice
          self.Exercice.Value := Exercice.Code ;
          // Date de Début de l'édition
          DateComptaDe.Text := DateToStr( Exercice.Deb );
          // Date de Fin de l'édition
          DateComptaA.Text := DateToStr( Exercice.Fin );
          // ModeSelection
          ModeSelect.Value := ModeSelection ;

          end ;

    TheData := nil;

    end ;


end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPBALANCECOMBI.OnLoad ;
var Q         : TQuery;
    stTypeEcr : String;
begin

  NextPrevControl(TFQRS1(Ecran));

  Inherited ;

  // Les ruptures...
  XXWhere.Text          := '';
  XXRupture.Text        := '';
  AvecNivoRupture.Text  := '0';
  AvecRupture.Text      := Rupture.value;
  if Rupture.Value = 'SANS'
    then AvecRupType.Text := ''
    else AvecRupType.Text := RuptureType.value;

  // Les dates
  DateDebutPer          := StrToDate(DateComptaDe.Text);
  DateFinPer            := StrToDate(DateComptaA.Text);
  if AvecComparatif.Checked and (ComparType.Value = 'COMPARPERIODE') then
    begin
    DateDebutCompar     := StrToDate(ComparDateDe.Text);
    DateFinCompar       := StrToDate(ComparDateA.Text);
    end ;

  // libellés pour les types d'écritures...
  AvecQualifPiece.Text := '';
  stTypeEcr := QualifPiece.Value;
  if stTypeEcr <> '' then
    while stTypeEcr <> '' do
      AvecQualifPiece.Text := AvecQualifPiece.Text
      			      + RechDom(QualifPiece.DataType,ReadTokenSt(stTypeEcr),False) + ' '
  else
    AvecQualifPiece.Text := '<<Tous>>';

  // Paramétrage des Titre des Colonnes
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
      end
    else SetControlText('TitreCol2','Solde du ' + DateToStr(DateDebutCompar) + ' au ' + DateToStr(DateFinCompar));
    end
  else
    begin    // Cas des totaux
 		SetControlText('TitreCol1','Total du ' + DateToStr(DateDebutPer) + ' au ' + DateToStr(DateFinPer)) ;
    SetControlText('TitreCol2','Solde du ' + DateToStr(DateDebutPer) + ' au ' + DateToStr(DateFinPer)) ;
    end ;

  // Paramétrage de l'affichage de la devise
  if Affichage.Value = 'EURO'
    then AvecAffichage.Text := Devise.Text + ' / Affichage ' + VH^.LibDevisePivot
    else AvecAffichage.Text := Devise.Text + ' / Affichage ' + Devise.Items[Devise.ItemIndex] ;

end ;

procedure TOF_CPBALANCECOMBI.OnArgument (S : String ) ;
begin
  Inherited ;

  // Gestion du mode "MULTI-DOSSIER"
  FBoRegroupement := pos( 'MULTIDOSSIER', S ) > 0 ;
  SetControlVisible( 'PREGROUPEMENT',  FBoRegroupement )  ;
  SetControlVisible( 'TPREGROUPEMENT', FBoRegroupement )  ;
  SetControlVisible( 'MULTIDOSSIER', FBoRegroupement )  ;
  // Modification du nom du filtre pour les éditions du menu "multidossier"
  if FBoRegroupement then
    TFQRS1(Ecran).FNomFiltre   := TFQRS1(Ecran).FNomFiltre + 'MD' ;


  // =====> Récupération des contrôles...
  Pages := TPageControl(GetControl('Pages'));
  // Contrôles critères et options - Page standard
  CompteDe	  := THEdit(GetControl('COMPTE', true));
  CompteA	  := THEdit(GetControl('COMPTE_', true));
  Compte2De	  := THEdit(GetControl('COMPTE2', true));
  Compte2A	  := THEdit(GetControl('COMPTE2_', true));
  NatureCpt	  := THValComboBox(Getcontrol('NATURECPT', true));
  ModeSelect	  := THValComboBox(Getcontrol('MODESELECT', true));
  QualifPiece	  := THMultiValComboBox(GetControl('QUALIFPIECE', true));
  DateComptaDe	  := THEdit(GetControl('DATECOMPTABLE', true));
  DateComptaA	  := THEdit(GetControl('DATECOMPTABLE_', true));
  // Contrôles critères et options - Page avancés
  CptExcept	  := THEdit(GetControl('CPTEXCEPT', true));
  CptExcept2  	  := THEdit(GetControl('CPTEXCEPT2', true));
  // Contrôles critères et options - Mise en page
  Rupture 	  := THRadioGroup(GetControl('RUPTURE', true));
  // Contrôles critères et options - Options d'éditions
  Affichage	  := THRadioGroup(GetControl('AFFICHAGE', true));
  AvecComparatif  := TCheckBox(GetControl('AVECCOMPARATIF', true));
  // Contrôles critères et options - Page ruptures
  RuptureType 	  := THRadioGroup(GetControl('RUPTURETYPE', true));
  NivoRupture 	  := TSpinEdit(GetControl('NIVORUPTURE', true));
  TableLibre      := THValComboBox(Getcontrol('TABLELIBRE', true));
  LibreDe	  := THEdit(GetControl('LIBREDE', true));
  LibreA	  := THEdit(GetControl('LIBREA', true));
  CptLibresExist  := TCheckBox(GetControl('CPTLIBRESEXIST', true));
  Corresp	  := THValComboBox(Getcontrol('CORRESP', true));
  CptCorrespExist := TCheckBox(GetControl('CPTCORRESPEXIST', true));
  CorrespDe	  := THValcomboBox(GetControl('CORRESPDE', true));
  CorrespA	  := THValcomboBox(GetControl('CORRESPA', true));
  // Contrôles invisibles utilisés dans l'état
  AvecNivoRupture := THEdit(GetControl('AVECNIVORUPTURE', true));
  AvecRupture     := THEdit(GetControl('AVECRUPTURE', true));
  AvecRupType 	  := THEdit(GetControl('AVECRUPTYPE', true));
  XXRupture 	  := THEdit(GetControl('XX_RUPTURE', true));
  XXWhere 	  := THEdit(GetControl('XX_WHERE', true));
  AvecQualifPiece := THEdit(GetControl('AVECQUALIFPIECE', true));
  AvecAffichage   := THEdit(GetControl('AVECAFFICHAGE', true));
  // Contrôles de la page COMPARATIF
  ComparDateDe	  := THEdit(GetControl('COMPARDATEDE', true));
  ComparDateA	  := THEdit(GetControl('COMPARDATEA', true));
  ComparExo	  := THValComboBox(Getcontrol('COMPAREXO', true));
  ComparType      := THRadioGroup(GetControl('COMPARTYPE', true));

  CInitComboExercice( Exercice ); // Init de la combo Exercice en Relatif
  CInitComboExercice( ComparExo ); // Init de la combo Exercice en Relatif

// Affichage de la devise pivot
  Affichage.Items[0]:='&'+VH^.LibDevisePivot ;

//=====> Evènements
  AvecComparatif.OnClick 	:= AvecComparatifClick;
  Rupture.OnClick 		:= RuptureClick;
  RuptureType.OnClick 		:= RuptureTypeClick;
  CompteDe.OnExit 		:= CompteOnExit;
  CompteA.OnExit 		:= CompteOnExit;
  Compte2De.OnExit 		:= Compte2OnExit;
  Compte2A.OnExit 		:= Compte2OnExit;
  DateComptaDe.OnExit  		:= DateOnExit;
  DateComptaA.OnExit 	 	:= DateOnExit;
  Exercice.onChange		:= ExoOnChange;
  Corresp.OnChange		:= PlanCorrespChanged;
  TableLibre.OnChange		:= TableLibreChanged;
  Devise.OnChange		:= DeviseChanged;
  NatureCpt.OnChange		:= NatureCptChanged;
  ModeSelect.OnChange		:= ModeSelectChanged;
  ComparDateDe.OnExit  		:= ComparDateOnExit;
  ComparDateA.OnExit 	 	:= ComparDateOnExit;
  ComparExo.onChange		:= ComparExoOnChange;
  CorrespDe.OnChange            := CorrespChanged ;
  CorrespA.OnChange             := CorrespChanged ;
  LibreDe.OnChange              := LibreChanged ;
  LibreA.OnChange               := LibreChanged ;

  // GCO - 26/11/2004 - FQ 13484
  FTimer.OnTimer := FTimerTimer;

  // Rendre non saisissable la zone des qualifpiece !
  QualifPiece.ReadOnly := True ;
  QualifPiece.OnEnter := QualifPieceEnter ; // Pb visibilité du focus avec le readonly

  TFQRS1(Ecran).OnAfterFormShow := OnAfterFormShow ;

  {YMO 12/04/2007 Branchement de la fiche auxi selon le ThEdit}
  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    If Copy(THEdit(GetControl('COMPTE', true)).datatype,1,3) = 'TZT' then
    begin
      THEdit(GetControl('COMPTE', true)).OnElipsisClick:=AuxiElipsisClick;
      THEdit(GetControl('COMPTE_', true)).OnElipsisClick:=AuxiElipsisClick;
    end
    else
    begin
      { FQ 22071 BVE 20.12.07 } {FQ22426 14.02.08 YMO Restriction aux tiers }
      if Copy(THEdit(GetControl('COMPTE2', true)).datatype,1,3) = 'TZT' then
      begin
        THEdit(GetControl('COMPTE2', true)).OnElipsisClick:=AuxiElipsisClick;
        THEdit(GetControl('COMPTE2_', true)).OnElipsisClick:=AuxiElipsisClick;
      end;
      { END FQ 22071 }
    end
  end;

end ;

//==============================================================================

procedure TOF_CPBALANCECOMBI.RuptureTypeClick(Sender: TObject);
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
    end
  else if RuptureType.value = 'RUPLIBRES' then
    begin
    TableLibre.ItemIndex := 0;
    TableLibreChanged(nil);
    end
  else if RuptureType.value = 'RUPGROUPES' then
    NivoRupture.Value := 1 ;
end;

procedure TOF_CPBALANCECOMBI.AvecComparatifClick(Sender: TObject);
begin
  if AvecComparatif.checked then
    begin
    SetControlVisible('COMPARATIF',True) ;
    //SG6 14.03.05 Gestion des filtres FQ 14996
    if cLoadFiltre then Exit;
    if VH^.Precedent.Code <> ''
      then ComparExo.Value := CExerciceVersRelatif (VH^.Precedent.Code)
      else ComparExo.Value := CExerciceVersRelatif (VH^.Entree.Code) ;
    end
  else
    SetControlVisible('COMPARATIF',False) ;
end;

procedure TOF_CPBALANCECOMBI.RuptureClick(Sender: TObject);
begin
  if Rupture.value <> 'SANS' then
    begin
    SetControlVisible('RUPTURES',True);
    RuptureTypeClick(nil);
    end
  else
    SetControlVisible('RUPTURES',False);
end;

procedure TOF_CPBALANCECOMBI.CompteOnExit(Sender: TObject) ;
begin
end;

procedure TOF_CPBALANCECOMBI.DateOnExit(Sender: TObject) ;
begin
  DoDateOnExit(THEdit(Sender), DateComptaDe, DateComptaA, DateD, DateF);
end;

procedure TOF_CPBALANCECOMBI.ExoOnChange(Sender: TObject) ;
begin
  CExoRelatifToDates( Exercice.Value, DateComptaDe, DateComptaA);
  DateD := StrToDate(DateComptaDe.Text);
  DateF := StrToDate(DateComptaA.Text);
end;

procedure TOF_CPBALANCECOMBI.PlanCorrespChanged(Sender: TObject);
begin
  if RuptureType.value <> 'RUPCORRESP' then Exit;
  CorrespToCodes(Corresp,TComboBox(CorrespDe),TComboBox(CorrespA));
  CorrespDe.ItemIndex	:= 0 ;
  CorrespA.ItemIndex	:= CorrespA.Items.Count - 1 ;
end;

procedure TOF_CPBALANCECOMBI.TableLibreChanged(Sender: TObject);
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
procedure TOF_CPBALANCECOMBI.DeviseChanged(Sender: TObject);
begin
  if Devise.ItemIndex > 0
    then Affichage.Enabled := True
    else
      begin
      Affichage.Value   := 'EURO';
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
procedure TOF_CPBALANCECOMBI.NatureCptChanged(Sender: TObject);
begin
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
procedure TOF_CPBALANCECOMBI.FTimerTimer(Sender: TObject);
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
function TOF_CPBALANCECOMBI.RecupWhereQualifPiece: String;
Var lStSql   : String ;
    lstTmp  : String ;
begin
  TraductionTHMultiValComboBox( QualifPiece, lStSql , lstTmp, 'E_QUALIFPIECE', False );
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
procedure TOF_CPBALANCECOMBI.ModeSelectChanged(Sender: TObject);
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

procedure TOF_CPBALANCECOMBI.ComparExoOnChange(Sender: TObject);
begin
  CExoRelatifToDates(ComparExo.Value, ComparDateDe, ComparDateA);
  DateDCompar := StrToDate(ComparDateDe.Text);
  DateFCompar := StrToDate(ComparDateA.Text);
end;

procedure TOF_CPBALANCECOMBI.ComparDateOnExit(Sender: TObject);
begin
  DoDateOnExit(THEdit(Sender), ComparDateDe, ComparDateA, DateDCompar, DateFCompar);
end;


function TOF_CPBALANCECOMBI.RecupWhereBalanceCompar( vBoAvecDate : Boolean = True ): String;
begin
  // Exercice
  Result := ' E_EXERCICE = "' + CRelatifVersExercice(ComparExo.Value) + '" ' ;
  // Dates comptables
  if vBoAvecDate then
    Result := Result + ' AND E_DATECOMPTABLE >= "' + UsDateTime(DateDebutCompar) + '" '
                     + ' AND E_DATECOMPTABLE <= "' + UsDateTime(DateFinCompar) + '" ' ;
  // QualifPiece
  Result := Result + ' AND ' + RecupWhereQualifPiece ;
  // Etablissement
  if ComboEtab.Value<>'' then
    Result := result + ' AND E_ETABLISSEMENT = "' + ComboEtab.Value + '" ' ;
  // Devise
  if Devise.Value<>'' then
    Result := result + ' AND E_DEVISE = "' + Devise.Value + '" ' ;
end;

procedure TOF_CPBALANCECOMBI.DeleteCEDTBALANCE;
begin
  ExecuteSQL('DELETE FROM ' + GetTablePourBase('CEDTBALANCE') + ' WHERE CED_USER = "' + V_PGI.User + '"');
end;

procedure TOF_CPBALANCECOMBI.RemplirEDTBALANCE ;
begin
// 1. Effacement des enregistrements present pour le user
  DeleteCEDTBALANCE ;

// 2. Insertion de la liste des comptes cibles
  // 2.1 : Comptes du comparatifs si besoin (EN PREMIER)
  if (AvecComparatif.Checked) then
    begin
    if   ( (ComparType.Value = 'COMPARPERIODE') and (ComparExo.Value <> '') )
      or ( (ComparType.Value = 'COMPARBALSIT') and (BALSIT.text <> '') )
      then InsererComptesComparatif ;
    end ;

  // 2.2 : Comptes des généraux (COMPLETE les Comptes du comparatifs au besoin)
  InsererComptes;

// 3. Update des totaux / soldes
  UpdateCumulsCEDTBALANCE ;

end;

procedure TOF_CPBALANCECOMBI.UpdateCumulsCEDTBALANCE;
Var QGene       : TQuery ;
    lStReq      : String ;
    lStUpdate   : String ;
    lCredit1    : Double ;
    lDebit1     : Double ;
    lCredit2    : Double ;
    lDebit2     : Double ;
    lStCompte   : String ;
    lStCompte2  : String ;
    lStLibelle  : String ;
    lBoMajCol2  : Boolean ;
begin

  try  // finally
    try  // except

      // -----------------------------------------
      // Chargement de la liste des comptes cibles
      lStReq := 'SELECT CED_COMPTE, CED_LIBELLE, CED_COMPTE2 '
              + 'FROM ' + GetTablePourBase('CEDTBALANCE')
              + ' WHERE CED_USER="'+V_PGI.User+'"' ;
      QGene := OpenSQL(lStReq, True ) ;

      // Barre de progession
{
      InitMoveProgressForm( Ecran,
                            Ecran.Caption,
                            'Traitement des écritures en cours...',
                            100 ,
                            True,
                            True);
}

      // Parcours des comptes
      While not QGene.Eof do
        begin
        lDebit1     := 0 ;
        lCredit1    := 0 ;
        lDebit2     := 0 ;
        lCredit2    := 0 ;
        lStCompte   := QGene.FindField('CED_COMPTE').AsString ;
        lStLibelle  := QGene.FindField('CED_LIBELLE').AsString ;
        lStCompte2  := QGene.FindField('CED_COMPTE2').AsString ;  // Compte combinatoire
        lBoMajCol2  := false ;

//        if Not MoveCurProgressForm(lStCompte + ' - ' + lStLibelle) then break;
        MoveCur( False ) ;


        // ----------------------------------------------------------------
        // Calcul des totaux / soldes suivant paramètrage et type de compte
        if AvecComparatif.checked then
          begin
          // ==> Calcul du solde de la période éditée <==
          CalculDebitCredit( lStCompte,
                             lStCompte2,
                             False,
                             lDebit1,
                             lCredit1 ) ;
          // --> Calcul du solde
          CalculSolde( lDebit1, lCredit1, lDebit1, lCredit1 ) ;
          // ==> Calcul du solde de la période comparée (uniquement si comparatif à une période) <==
          if (ComparType.Value = 'COMPARPERIODE') and (ComparExo.Value <> '') then
            begin
            lBoMajCol2  := true ;
            // --> Calcul des totaux
            CalculDebitCredit( lStCompte,
                               lStCompte2,
                               True,
                               lDebit2,
                               lCredit2 ) ;
            // --> Calcul du solde
            CalculSolde( lDebit2, lCredit2, lDebit2, lCredit2 ) ;
            end ;
          end
        else
          begin
          // ==> cas simple : pas de comparatif totaux + solde <==
          lBoMajCol2  := true ;

          // --> Calcul des totaux
          CalculDebitCredit( lStCompte,
                             lStCompte2,
                             False,
                             lDebit1,
                             lCredit1 ) ;
          // --> Calcul du solde
          CalculSolde( lDebit1, lCredit1, lDebit2, lCredit2 ) ;
          end ;

        // -----------
        // REQUETE MAJ
        lStUpdate := 'UPDATE ' + GetTablePourBase('CEDTBALANCE')
                               + ' SET CED_CREDIT1=' + VariantToSQL(lCredit1)
                                  + ', CED_DEBIT1=' + VariantToSQL(lDebit1);
        if lBoMajCol2 then
          lStUpdate := lStUpdate + ',CED_CREDIT2=' + VariantToSQL(lCredit2)
                                + ', CED_DEBIT2=' + VariantToSQL(lDebit2);
        lStUpdate := lStUpdate + ' WHERE CED_COMPTE = "' + lStCompte + '"'
                                 + ' AND CED_COMPTE2 = "' + lStCompte2 + '"'
                                 + ' AND CED_USER = "' + V_PGI.User + '"' ;
        // --------
        // MAJ BASE
        ExecuteSQL( lStUpdate ) ;

        //---------
        // Suivant
        QGene.Next ;

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
    Ferme(QGene) ;
//    FiniMoveProgressForm ;
  end ;


end;

procedure TOF_CPBALANCECOMBI.CalculDebitCredit( vStCompte, vStCompte2: String ;  vBoCompar : Boolean ; var vDebit, vCredit: Double );
begin
  vDebit  := 0 ;
  vCredit := 0 ;
end;


procedure TOF_CPBALANCECOMBI.CalculSolde( vDebit, vCredit : Double ; var vSoldeDebit, vSoldeCredit : Double );
begin
  if vDebit >= vCredit then
    begin
    vSoldeDebit  := vDebit - vCredit ;
    vSoldeCredit := 0 ;
    end
  else
    begin
    vSoldeDebit  := 0 ;
    vSoldeCredit := vCredit - vDebit ;
    end ;
end;

function TOF_CPBALANCECOMBI.GenererRequeteBAL: String;
begin
  // Champs
  if FBoRegroupement and ( GetControlText('MULTIDOSSIER') <> '' )
    then result := 'SELECT SYSDOSSIER, CED_COMPTE ' + getNomChampCPT + ', CED_LIBELLE'
    else result := 'SELECT CED_COMPTE ' + getNomChampCPT + ', CED_LIBELLE' ;

  result := result + ', CED_COMPTE2 ' + getNomChampCPTCombi + ', CED_LIBELLE2'
                   + ', CED_DEBIT1, CED_CREDIT1'
                   + ', CED_DEBIT2, CED_CREDIT2'
                   + ', CED_NATURE CED_AXE' // GCO - 28/11/2006 - FQ 18927
                   + ', CED_RUPTURE, CED_COLLECTIF ' ;
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
  result := result + ' FROM CEDTBALANCE' ;
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
  result := result + ' WHERE CED_USER = "' + V_PGI.User + '" ' ;
  // Tri
  result := result + ' ORDER BY CED_RUPTURE, ' + getNomChampCPTCombi + ', ' + getNomChampCPT ;

end;


procedure TOF_CPBALANCECOMBI.InsererComptes;
var	QCpt2						: TQuery ;
    stCpt           : String ;
    stLib           : String ;
    stRup           : String ;
    lInNbEnreg      : LongInt ;
begin
  // Récup comptes principaux
  QCpt2 := OpenSQL ( GenererRequeteComptes , True ) ;

  While not QCpt2.Eof do
    begin
    // Recup champ
    stCpt   := QCpt2.FindField('CODE').AsString ;
    stLib   := QCpt2.FindField('LIBELLE').AsString ;
    if ( XXRupture.Text <> '' )
      then stRup := QCpt2.FindField('RUPTURE').AsString
      else stRup := '' ;
    // Insertion combis cpt1 / cpt 2
    lInNbEnreg := ExecuteSQL ( GenererRequeteCombis( stCpt, stLib, stRup ) ) ;
    if lInNbEnreg = 0 then
      ExecuteSQL ( GenererRequeteVide( stCpt, stLib, stRup ) ) ;
    // suivant
    QCpt2.next ;
    end;
  Ferme( QCpt2 ) ;

end;

procedure TOF_CPBALANCECOMBI.InsererComptesComparatif;
var QCpt2	    : TQuery ;
    stCpt           : String ;
    stLib           : String ;
    stRup           : String ;
begin
  // Récup comptes principaux
  QCpt2 := OpenSQL ( GenererRequeteComptesCompar , True ) ;
  While not QCpt2.Eof do
    begin
    // Recup champ
    stCpt   := QCpt2.FindField('CODE').AsString ;
    stLib   := QCpt2.FindField('LIBELLE').AsString ;
    if ( XXRupture.Text <> '' )
      then stRup := QCpt2.FindField('RUPTURE').AsString
      else stRup := '' ;
    // Insertion combis cpt1 / cpt 2
    if (ComparType.Value = 'COMPARPERIODE')
       then ExecuteSQL ( GenererRequeteCombisCompar( stCpt, stLib, stRup ) )
    else if (ComparType.Value = 'COMPARBALSIT')
       then ExecuteSQL ( GenererRequeteCombisBalsit( stCpt, stLib, stRup ) ) ;

    // suivant
    QCpt2.next ;
    end;
  Ferme( QCpt2 ) ;
end;

function TOF_CPBALANCECOMBI.GenererRequeteVide(vStCode, vStLib, vStRup: String): String;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// VALEURS
  // Champs à vide
	result := result + 'VALUES ( "' + V_PGI.User + '", "", "", 0, 0, 0, 0';
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '" )' ;

end;

function TOF_CPBALANCECOMBI.GetMaxCompteVentil(vStNature, vStAxe: string): String;
var lQuery : TQuery ;
    lStReq : String ;
begin
  Result := '';
  lStReq := 'SELECT MAX( G_GENERAL ) CODE '
            + 'FROM GENERAUX '
            + 'WHERE G_VENTILABLE="X" AND G_VENTILABLE' + vStAxe + '="X"' ;
  if vStNature <> '' then
    lStReq  := lStReq + ' AND G_NATUREGENE="' + vStNature + '"' ;
  try
    try
      lQuery := OpenSQL( lStReq, True);
      if not lQuery.Eof then
        Result := lQuery.FindField('CODE').AsString;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : GetMaxCompteVentil');
    end;

  finally
    Ferme( lQuery );
  end;
end;

function TOF_CPBALANCECOMBI.GetMinCompteVentil(vStNature, vStAxe: string): String;
var lQuery : TQuery ;
    lStReq : String ;
begin
  Result := '';
  lStReq := 'SELECT MIN( G_GENERAL ) CODE '
            + 'FROM GENERAUX '
            + 'WHERE G_VENTILABLE="X" AND G_VENTILABLE' + vStAxe + '="X"' ;
  if vStNature <> '' then
    lStReq  := lStReq + ' AND G_NATUREGENE="' + vStNature + '"' ;

  try
    try
      lQuery := OpenSQL( lStReq, True);
      if not lQuery.Eof then
        Result := lQuery.FindField('CODE').AsString;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : GetMaxCompteVentil');
    end;

  finally
    Ferme( lQuery );
  end;

end;

function TOF_CPBALANCECOMBI.GetNomChampCPT: String;
begin

end;

function TOF_CPBALANCECOMBI.GetNomChampCPTCombi: String;
begin

end;

procedure TOF_CPBALANCECOMBI.CorrespChanged(Sender: TObject);
begin
  if (CorrespDe.ItemIndex > 0) or (CorrespA.ItemIndex < (CorrespA.Items.count - 1)) then
    begin
    CptCorrespExist.Checked := True ;
    CptCorrespExist.Enabled := False ;
    end
  else
    CptCorrespExist.Enabled := True ;
end;

procedure TOF_CPBALANCECOMBI.LibreChanged(Sender: TObject);
begin
  if (LibreDe.Text<>'') or (LibreA.Text<>'') then
    begin
    CptLibresExist.Checked := True ;
    CptLibresExist.Enabled := False ;
    end
  else
    CptLibresExist.Enabled := True ;
end;

function TOF_CPBALANCECOMBI.GetConditionRupture: String;
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
      if (Trim(LibreA.Text) = '') and (Trim(LibreDe.Text) = '') and CptLibresExist.Checked then
        Result := Result + ' AND ' + XXRupture.Text + '<>""';
      end;
    // Conditions sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
        begin
        if CorrespDe.ItemIndex > 0 then
          Result := Result + ' AND ' + XXRupture.Text + ' >= "' + CorrespDe.Value + '"';
        if CorrespA.ItemIndex < (CorrespA.Items.count - 1) then
          Result := Result + ' AND ' + XXRupture.Text + ' <= "' + CorrespA.Value + '"';
        if (CorrespDe.ItemIndex = 0) and (CorrespA.ItemIndex = (CorrespA.Items.count - 1))
           and CptCorrespExist.Checked then
          Result := Result + ' AND ' + XXRupture.Text + '<>""';
        end;
    end;
end;

function TOF_CPBALANCECOMBI.RecupWhereBalance( vBoAvecDate : Boolean ): String;
begin
  Result := RecupWhereCritere( Pages);
  Result := CMajRequeteExercice ( Exercice.Value, Result);
  // Dates comptables
  if vBoAvecDate then
    Result := Result + ' AND E_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DateComptaDe.Text)) + '" '
                     + ' AND E_DATECOMPTABLE <= "' + UsDateTime(StrToDate(DateComptaA.Text)) + '" ' ;
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
function TOF_CPBALANCECOMBI.CompleteAuto( Sender: TObject ; CompteType : TFichierBase ): Boolean;
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

procedure TOF_CPBALANCECOMBI.Compte2OnExit(Sender: TObject);
begin
end;

procedure TOF_CPBALANCECOMBI.ComparTypeClick(Sender: TObject);
begin
  SetControlVisible('COMPARBALSIT',False);
  SetControlVisible('COMPARPERIODE',False);
  if ComparType.value <> '' then
    SetControlVisible(ComparType.value,True)
  else if ComparType.Value = 'COMPARPERIODE' then
    if VH^.Precedent.Code <> ''
      then ComparExo.Value := CExerciceVersRelatif (VH^.Precedent.Code)
      else ComparExo.Value := CExerciceVersRelatif (VH^.Entree.Code) ;
end;

////////////////////////////////////////////////////////////////////////////////
{$IFNDEF GCGC}
procedure TOF_CPBALANCECOMBI.OnChangeFiltre(Sender: TObject);
begin
  inherited;
  CExoRelatifToDates ( Exercice.Value, DateComptaDe, DateComptaA, True);
  GereAffichageGlobal ;
end;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPBALANCECOMBI.QualifPieceEnter(Sender: TObject);
begin
  CSelectionTextControl( Sender ) ;
end;

procedure TOF_CPBALANCECOMBI.NouvRechClick(Sender: TObject);
begin

  videFiltre( FFiltres, Pages ) ;
  TFQRS1(Ecran).ListeFiltre.new ;

  // INIT ZONES
  InitZones ;

  // === MAJ INTERFACE ===
  GereAffichageGlobal ;

end;

procedure TOF_CPBALANCECOMBI.OnAfterFormShow;
begin
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

function TOF_CPBALANCECOMBI.GetTablePourBase(vStTable: String): String;
begin
  result := vStTable ;

  // Si la table est partagée, on laisse la main à l'AGL...
  if EstTablePartagee( vStTable ) then Exit ;
  
  if FStBase <> '' then
    if IsMSSQL  //10/08/2006 YMO Ajout test SQL2005
      then result := FStBase + '.DBO.' + vStTable
      else result := FStBase + '.' + vStTable ;
end;

procedure TOF_CPBALANCECOMBI.GereAffichageGlobal;
begin

  // ===================================================
  // === Gestion de l'affichage des pages dynamiques ===
  // ===================================================
  // --> Onglet Rupture
  SetControlVisible('RUPTURES', (Rupture.value <> 'SANS') );
  // --> Onglet Comparatif
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
  SetControlVisible('COMPARBALSIT',False);
  SetControlVisible('COMPARPERIODE',False);
  if ComparType.value <> '' then
    SetControlVisible(ComparType.value,True);

  // ========================================
  // === Gestion des tablettes dynamiques ===
  // ========================================
  // --> Onglet Général
  AxeChanged( nil ) ;        // Sert Uniquement pour la bal gen / ana
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
    if (Compte2De <> nil) then
      Compte2De.Text  := GetValeurFiltre('COMPTE2') ;
    if (Compte2A <> nil) then
      Compte2A.Text   := GetValeurFiltre('COMPTE2_') ;
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

function TOF_CPBALANCECOMBI.GetValeurFiltre(vStZone: String): String;
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

procedure TOF_CPBALANCECOMBI.AxeChanged(Sender: TObject);
begin

end;

procedure TOF_CPBALANCECOMBI.InitZones;
begin
  // pages standards
  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' )
    then Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
    else Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;
//  ExoOnChange(nil); // ?? auto ??

  // Type de compte
  NatureCpt.ItemIndex  := 0;

  // Mode de sélection
  ModeSelect.ItemIndex := 0;

  // Type de pièce
  QualifPiece.value 	 := 'N;';

  // pages avancés
  Devise.ItemIndex 	 := 0;
//  DeviseChanged(nil);            // ?? Auto

  // Options
  Rupture.Value          := 'SANS' ;
  AvecComparatif.Checked := False ;

  GereEtablissement; {JP 29/06/06 : FQ 16149}
end;

procedure TOF_CPBALANCECOMBI.SupprFiltreClick(Sender: TObject);
begin
  if not TFQRS1(Ecran).ListeFiltre.Delete then Exit ;
  videFiltre( FFiltres, Pages ) ;
  InitZones ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANCECOMBI.AuxiElipsisClick(Sender: TObject);
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

function TOF_CPBALANCECOMBI.GetNumAxe: String;
begin
  result := '' ;
  if NatureCpt.ItemIndex < 0 then Exit ;
  result := copy( NatureCpt.Value, 2, 1 ) ;
end;

function TOF_CPBALANCECOMBI.GetNumTableLibre: String;
begin
  result := '' ;
  if TableLibre.ItemIndex < 0 then Exit ;
  result := Copy( TableLibre.Value, 3, 1 )  ;
end;

procedure TOF_CPBALANCECOMBI.MySelectFiltre;
var lZExercice : TZExercice;
    lExoDate : TExoDate;
begin
  inherited;
  lZExercice := TZExercice.Create(False);
  QuelDateDeExo(CRelatifVersExercice(Exercice.Value), lExoDate);

  if (lZExercice.QuelExo(DateComptaDe.Text, False) <> lExoDate.Code) and
     (lZExercice.QuelExo(DateComptaA.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(Exercice.Value, DateComptaDe, DateComptaA);
  end;
  FreeAndNil(lZExercice);
end;

Initialization
  registerclasses ( [ TOF_CPBALANCECOMBI ] ) ;
end.

