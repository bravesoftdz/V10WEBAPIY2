{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :   /  /
Description .. : Gestion des la grille multicritères des affaires
Mots clefs ... : TOF;UTOFAFFAIRE_MUL
*****************************************************************}
unit UTofAffaire_Mul;

interface

uses  StdCtrls,
			Controls,
      Classes,
      M3FP,
      HTB97,
      HPanel,
      windows,
      messages,
      Ent1,
{$IFDEF EAGLCLIENT}
      eMul,
      Maineagl,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      DBGrids,
      mul,FE_Main,
{$ENDIF}
      HQry,
			Choix,
      forms,
      sysutils,
      HDB,
      ParamSoc,
      utob,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOF,
      ConfidentAffaire,
      AffaireUtil,
      AffaireRegroupeUtil,
      Dicobtp,
      SaisUtil,
      EntGC,
      BTPUtil,
      utofAfBaseCodeAffaire,
      utilpgi,
      AglInit,
      UtilGc,
      UtofAfAvenant, // à laisser car du scirpt, appel d'une fiche avec cette tof. Pour Ok dans projet
      AssistCreationAffaire, // mcd 10/06/02 appelé du script. a laisser
      StrUtils,
      UplanningBTP,
      TraducAffaire,
           Graphics
;
Type
		 TModegestion = (TTgStd,TTgCotraitance);
     TGridDrawState = set of (gdSelected, gdFocused, gdFixed);

     TOF_AFFAIRE_MUL = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnLoad ; override ;
        procedure OnClose; Override;
        procedure AfterShow;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        procedure TraiteRetour;
        procedure TraiteInvite;
     private
        fModeVisuPLanCharge : TModeGestionPC;
        fTOBSelect : TOB;
       	fModeGestion : TModeGestion;
     		MAffaire0   : THValComboBox;
        VUserInvite : boolean;
        AppelPlanning : Boolean;
        GNbLoads 		: integer;
{$IFDEF CTI}
				Contact : string;
				Numtel : string;
{$ENDIF}
        BOuvrir1   	: TToolBarButton97;
        BInsert			: TToolBarButton97;
        BDuplication: TToolBarButton97;

        DateFin     : THEdit;
        Datefin_    : THEdit;
        DateAAA     : THEdit;
        DateAAA_    : THEdit;

        Fonction 		: string;
        Affaire0		: String;
        Statut			: String;
        sPlus       : String;
				NomFiche		: String;

        Top 				: Integer;
        bPasSiUn		: Boolean;
        bItDatesMax : Boolean;
        AffEtat     : Boolean;
        ChangeStatut: Boolean;
        CREERPAR    : THCheckbox;
        EtatAffaire : THMultiValComboBox;

        StatutAff		: THValComboBox;

        Fliste 			: THDbGrid;

        TResponsable      : THEdit;
        TTiers            : THEdit;
        TDomaine          : THEdit;
        TEtablissement    : THEdit;
        LibResponsable    : THLabel;
        LibDomaine        : THLabel;
        LibEtablissement  : THLabel;
        LibClient         : THLabel;

		    procedure AffectationRessource(Sender : TOBJect);
				procedure ControleCritere(Critere, Champ, Valeur : string);
		    procedure StatutExit(Sender: TObject);
        procedure PositionneAffaire0 (Sender : Tobject);
        //
        //Modif FV
				Procedure FlisteDblClick (Sender : TObject);
		    Procedure BInsert_OnClick(Sender: TObject);
        procedure ExitResponsable(Sender: TObject);
        procedure ExitTiers(Sender: TObject);
        procedure BRechResponsable(Sender: TObject);
        procedure BDuplication_OnClick(Sender: TObject);
		    procedure ChargeEcranStatutAffaire;
        procedure VisuAffaireClick(Sender: TObject);
				procedure LanceChangeTvaContrats;
        procedure DateFinOnExit(Sender: TObject);
        procedure CREERPAROnClick(Sender: TOBject);
        procedure ChantierSelect (Sender : TObject);
        procedure Previsionnelpresent (Sender : TObject);
        procedure GetObjects;
        procedure ChargeResponsable;
        Procedure ChargeDomaine;
        procedure ChargeEtablissement(Abrege : Boolean=False);
     END ;
Type
     TOF_APPORTEUR_MUL = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
     END ;

Procedure AFLanceFiche_MulEAffaire;
Procedure AFLanceFiche_MulAffaire(Range,Argument:string);
Procedure AFLanceFiche_MulAffCompta;
Procedure AFLanceFiche_Mul_Apporteur;
Procedure AFLanceFiche_MulAvenantAff(Range,Argument:string);
Procedure AFLanceFiche_IntervAff;

Function  AFLanceFiche_AffaireRech(Range,Argument:string):variant;


implementation
uses  UtilRessource,
      uTOFComm,
      CalcOLEGenericBTP,
      PiecesRecalculs,
      AffEcheanceUtil
      ;


procedure TOF_AFFAIRE_MUL.AfterShow;
begin

with TFMul(Ecran) do
  begin
  if Q.RecordCount=1 then
    begin
    if isInside(TFmul(Ecran)) then
      begin
      THPanel(parent).CloseInside;
      THPanel(parent).VideToolBar ;
      end ;
{$IFNDEF EAGLCLIENT}
    Retour := Fliste.datasource.dataset.findfield('AFF_AFFAIRE').text+';'+Fliste.datasource.dataset.findfield('AFF_TIERS').text;
{$ELSE}
    Retour := Q.TQ.FindField('AFF_AFFAIRE').AsString+';'+Q.TQ.FindField('AFF_TIERS').AsString;
{$ENDIF}
//    if Not (V_PGI.ModeTSE) then
      PostMessage(TFmul(Ecran).Handle, WM_CLOSE, 0, 0);
//    PostMessage(Handle, WM_CLOSE, 0, 0);
    end;
  end;
end;

procedure TOF_AFFAIRE_MUL.OnArgument(stArgument : String );
var Critere {, sValeursEtatsAffaire} : string;
    champ,valeur : string;
    X : integer;
Begin
  fModeGestion := TTgStd;
	fMulDeTraitement  := true;
Inherited;
   fTOBSelect := TOB.Create ('LES AFFAIRES',nil,-1);
   fTOBSelect.AddChampSupValeur('TOUS','X');
   fTOBSelect.AddChampSupValeur('CRITERES','');
   fTOBSelect.AddChampSupValeur('TYPERESSOURCES','');
   fTOBSelect.AddChampSupValeur('ETABLISSEMENT','');

	 fTableName := 'AFFAIRE';
//
   AppelPlanning := false;
//
   GNbLoads := 0;
   bPasSiUn := false;
   bItDatesMax := false;
   VUserInvite:=False;
   top :=0;
   Statut :='';
   sPlus := '';
   Fonction := '';
   //
   ChangeStatut := True;
   top :=1;
   //

  GetObjects;
   Critere:=(Trim(ReadTokenSt(stArgument)));

   While (Critere <>'') do
       BEGIN
    Champ   := '';
    Valeur  := '';
       X:=pos(':',Critere);
       if x=0 then X:=pos('=',Critere);
       if x<>0 then
          begin
          Champ  :=copy(Critere,1,X-1);
          Valeur :=Copy (Critere,X+1,length(Critere)-X);
          end
       else
          champ  := critere;
       ControleCritere(Critere, Champ, Valeur);
       Critere:=(Trim(ReadTokenSt(stArgument)));
       END;

  if Statut = 'APP' then
    NomFiche := 'BTAPPELINT'
  else
 	 Nomfiche := 'BTAFFAIRE';

   {$IFDEF EAGLCLIENT}
   TraduitAFLibGridSt(TFMul(Ecran).FListe);
   {$ELSE}
   TraduitAFLibGridDB(TFMul(Ecran).FListe);
   {$ENDIF}
  if Not(ModifAffaireAutorise) and (GetControlText('XXAction') <> 'ACTION=RECH') then SetControlText('XXAction','ACTION=CONSULTATION');

  If GetControl ('XXACTION') <> Nil then
  begin
  if GetControlText('XXAction') <> '' then
   begin
   if (StringToAction(GetControlText('XXAction')) = taconsult) then
      begin
    {$IFDEF BTP}
      SetControlVisible ('BInsert',true); // on autorise la creation d'affaire en recherche
    {$ELSE}
      SetControlVisible ('BInsert',False);
    {$ENDIF}
      SetControlVisible ('BInsert1',False);
      end;
   end;
  end;

  ChargeEcranStatutAffaire;

  if AppelPlanning then
  begin
      Ecran.Caption     := TraduitGA('Réalisation par lots des Appels');
      CREERPAR.Visible  := AppelPlanning;
      SetControlproperty('FListe', 'Multiselection', False);
      SetControlVisible ('BInsert',False);
      SetControlVisible ('BDuplication',False);
      SetControlVisible ('BOuvrir',False);
      SetControlVisible ('BSelectAll',False);
      THMultiValComboBox(GetControl('AFF_ETATAFFAIRE')).Value := 'ECO;';
      ETATAFFAIRE.Enabled := False;
  end;

  // Caption de la form
  if Ecran.Name = 'AFFAIREAVT_MUL' then
   begin // Avenants sur affaires
   if Statut = 'PRO' then
   	  Ecran.Caption := TraduitGA('Avenants sur propositions d''affaires')
   else if Statut = 'INT' then
      Ecran.Caption := TraduitGA('Avenants sur contrat')
   else if Statut = 'APP' then
      Ecran.Caption := TraduitGA('Compléments d''appel')
   else
   	  Ecran.Caption := TraduitGA('Avenants sur affaires');
   end;

  UpdateCaption(Ecran);

  if Ecran.Name = 'BTPREVFACT_MUL' then
  begin
    ChargeResponsable;
    ChargeEtablissement(False);
    Chargedomaine;
    SetControlVisible('BDuplication',False);
    SetControlVisible ('BInsert',false);
  end;

  // on peut modifier le statut, il faut donc avoir les 2 boutons d'insert
  if (top = 0) then
   begin
   if fonction <> '' then
      Begin
      SetControlVisible ('BInsert',false);
      SetControlVisible ('BInsert1',False);
      end
   else
      Begin
	    SetControlVisible ('BInsert',true);
	    SetControlVisible ('BInsert1',true);
      end;
  end;

// Gestion des sous affaires
if not(GereSousAffaire) then
   begin
   SetcontrolVisible('TAFF_AFFAIREREF',False);
   SetcontrolVisible('AFFAIREREF1',False);
   SetcontrolVisible('AFFAIREREF2',False);
   SetcontrolVisible('AFFAIREREF3',False);
   SetcontrolVisible('AFFAIREREF4',False);
   SetcontrolVisible('AFF_ISAFFAIREREF',False);
   SetcontrolVisible('BSELECTAFF2',False);
   end;

{$IFDEF BTP}
//SetcontrolVisible('PCOMPLEMENT',False);
SetcontrolVisible('TAFF_TOTALHTGLO',False);
SetcontrolVisible('AFF_TOTALHTGLO',False);
SetcontrolVisible('TAFF_TOTALHTGLO_',False);
SetcontrolVisible('AFF_TOTALHTGLO_',False);
SetcontrolVisible('AFF_MODELE',False);
SetcontrolVisible('AFF_ADMINISTRATIF',False);

SetcontrolVisible('TAFF_STATUTAFFAIRE',ChangeStatut);
SetcontrolVisible('AFF_STATUTAFFAIRE',ChangeStatut);

SetcontrolVisible('TAFF_ETATAFFAIRE',AffEtat);
SetcontrolVisible('AFF_ETATAFFAIRE',AffEtat);

SetControlVisible('TAFF_CHANTIER', False);
SetControlVisible('AFF_CHANTIER', False);

{$ENDIF}

// Test droit d'accès en création
if not(CreationAffaireAutorise) then
   begin
   SetControlVisible('bAssistantCreation',False);
   SetControlVisible('bInsert',False);
   SetControlVisible ('BInsert1',False);
   SetControlVisible('BDuplication',False);
   end;
{$IFDEF EAGLCLIENT}
//SetControlVisible ('BASSISTANTCREATION',False); portage en Eagl de l'assistant création affaire
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}

// PL le 25/06/03 : gestion de la non ouverture si un seul element est trouvé
if bPasSiUn then
  begin
    TFmul(Ecran).OnAfterFormShow := AfterShow;
  end;

	// PL le 09/07/03 : on force l'intervalle de dates au maximum pour ramener toutes les missions dans certains cas (saisie d'activité par exemple)
  // On réduit également le tiers au code tiers exact et non pas au like "ZZZ%"
	if bItDatesMax then
  	 begin
     SetControlText ('AFF_DATEDEBUT', datetostr (idate1900));
     SetControlText ('AFF_DATEDEBUT_',  datetostr (idate2099));
   	 SetControlProperty( 'AFF_TIERS', 'operateur', Egal);
     TFMul(Ecran).FListe.SetFocus;
  	 end;

  {$IFDEF CCS3}
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  if (getcontrol('PSTAT') <> Nil) and  (Ecran.Name <> 'AFFAIREAVT_MUL') then SetControlVisible ('PSTAT', False);
  {$ENDIF}

  if (ecran.name = 'BTCONTRAT_MUL') then
  begin
    if THMultiValComboBox(GetControl('AFF_TYPEAFFAIRE')) <> nil then
    begin
      THMultiValComboBox (GetControl('AFF_TYPEAFFAIRE')).value := '';
      THLabel(GetControl('TAFFAIRE0')).caption := 'Type contrat :'
    end;
  end
  else if (Ecran.name = 'BTMULAPPELS') then
  begin
    if THMultiValComboBox(GetControl('AFF_TYPEAFFAIRE')) <> nil then
    begin
      THMultiValComboBox (GetControl('AFF_TYPEAFFAIRE')).value := '';
      THLabel(GetControl('TAFF_TYPEAFFAIRE')).caption := 'Type Intervention :'
    end;

  end else if (Ecran.name = 'BTTVACONTRAT_MUL') then
  begin
    if THMultiValComboBox(GetControl('AFF_TYPEAFFAIRE')) <> nil then
    begin
      THMultiValComboBox (GetControl('AFF_TYPEAFFAIRE')).value := '';
      THLabel(GetControl('TAFFAIRE0')).caption := 'Type contrat :'
    end;
    SetControlVisible('BInsert',false);
    SetControlVisible('BVisuAFFAIRE',true);
    TToolbarButton97(getControl('BVISUAFFAIRE')).onclick := VisuAffaireClick;
  end else if (Ecran.name = 'BTAFFAIRE_SEL') then
  begin
    SetControlVisible('BInsert',false);
    SetControlVisible('BVisuAFFAIRE',true);
    SetControlVisible('BSELECTALL',true);
    SetControlVisible('BDUPLICATION',false);
    SetControlEnabled('AFFAIRE0',false);
    THMultiValComboBox(GetControl('AFF_ETATAFFAIRE')).Plus := 'AND CC_CODE="ACP"';
    THMultiValComboBox(GetControl('AFF_ETATAFFAIRE')).Text := 'ACP';
    SetControlEnabled('AFF_ETATAFFAIRE',false);
    TToolbarButton97(getControl('BOUVRIR')).OnClick := ChantierSelect;
    if fModeVisuPLanCharge = TmPcPlanCharge then ecran.caption := 'Plan de charge / Chantiers'
                                            else ecran.caption := 'Plan de charge / Fonctions';
    ThEdit(getControl('XX_WHERESEL')).Text := ' AND EXISTS (SELECT 1 FROM PIECE WHERE GP_NATUREPIECEG="PBT" AND GP_AFFAIRE=AFF_AFFAIRE)';
    UpdateCaption(ecran);
  end;


//uniquement en line
{*
  if THLabel(GetControl('TAFF_RESPONSABLE')) <> nil then THLabel(GetControl('TAFF_RESPONSABLE')).visible := false;
  if THLabel(GetControl('AFF_RESPONSABLE')) <> nil then THLabel(GetControl('AFF_RESPONSABLE')).visible := false;
  if THLabel(GetControl('TAFFAIRE0')) <> nil then THLabel(GetControl('TAFFAIRE0')).visible := false;
  if THValComboBox(GetControl('AFFAIRE0')) <> nil then THValComboBox(GetControl('AFFAIRE0')).visible := false;
  if THLabel(GetControl('TAFF_ETABLISSEMENT')) <> nil then THLabel(GetControl('TAFF_ETABLISSEMENT')).visible := false;
  if THValComboBox(GetControl('AFF_ETABLISSEMENT')) <> nil then THValComboBox(GetControl('AFF_ETABLISSEMENT')).visible := false;
	TFMUL(Ecran).SetDbliste('BTMULAFFAIRE_S1');
*}

End;

Procedure TOf_AFFAIRE_MUL.GetObjects;
begin

  if assigned(GetControl('AFF_MANDATAIRE')) then THDBCheckbox (GetControl('AFF_MANDATAIRE')).State := cbGrayed;

  if Assigned(GetControl('AFF_RESPONSABLE')) Then
  begin
      TResponsable                := THEdit(GetControl('AFF_RESPONSABLE'));
      TResponsable.OnElipsisClick := BRechResponsable;
      TResponsable.OnEXit         := ExitResponsable;
  end;

  If Assigned(GetControl('AFF_TIERS')) then
  begin
      TTiers        := THEdit(GetControl('AFF_TIERS'));
      TTiers.OnEXit := ExitTiers;
  end;

  If Assigned(GetControl('AFF_DOMAINE')) then
  begin
      TDomaine        := THEdit(GetControl('AFF_DOMAINE'));
  end;

  If Assigned(GetControl('AFF_ETABLISSEMENT')) then
  begin
      TEtablissement        := THEdit(GetControl('AFF_ETABLISSEMENT'));
  end;

  if Assigned(GetControl('LIBRESSOURCE'))     Then LibResponsable   := THLabel(Getcontrol('LIBRESSOURCE'));
  if Assigned(GetControl('LIBDOMAINE'))       Then LibDomaine       := THLabel(Getcontrol('LIBDOMAINE'));
  if Assigned(GetControl('LIBETABLISSEMENT')) Then LibEtablissement := THLabel(Getcontrol('LIBETABLISSEMENT'));
  if Assigned(GetControl('LIBCLIENT'))        Then LibClient        := THLabel(Getcontrol('LIBCLIENT'));

  if copy(ecran.name ,1,12) <> 'BTAPPOFF_MUL' then
  begin
   if Assigned(GetControl('AFF_ETATAFFAIRE')) then EtatAffaire := THMultiValComboBox (GetControl ('AFF_ETATAFFAIRE'));
  end;

  if assigned(GetControl('CREERPAR')) then
  begin
    CREERPAR := THCheckbox (GetControl('CREERPAR'));
    CREERPAR.Checked := False;
    CREERPAR.Visible := AppelPlanning;
    CREERPAR.OnClick := CREERPAROnClick;
  end;

  if assigned(GetControl('AFF_MANDATAIRE')) then THDBCheckbox (GetControl('AFF_MANDATAIRE')).State := cbGrayed;

  //FV1 : 22/12/2015 - FS#1816 - VEODIS : multicritère contrats, Critère de recherche : Pb affichage date fin (pas dans la bonne case
  if assigned(GetControl('DATEAAA'))       then
  begin
    DateAAA  := THEdit(GetControl('DATEAAA'));
    DateAAA.OnExit := DateFinOnExit;
  end;
  //
  if assigned(GetControl('DATEAAA_'))      then
  begin
    DateAAA_ := THEdit(GetControl('DATEAAA_'));
    DateAAA_.OnExit := DateFinOnExit;
  end;

  if assigned(GetControl('AFF_DATEFIN'))   then DateFin  := THEdit(GetControl('AFF_DATEFIN'));
  if assigned(GetControl('AFF_DATEFIN_'))  then DateFin_ := THEdit(GetControl('AFF_DATEFIN_'));
  //
  if copy(ecran.name ,1,12) <> 'BTAPPOFF_MUL' then
  begin
    Fliste := THDbGrid (GetControl('FLISTE'));
    Fliste.OnDblClick := FlisteDblClick;
  end;

  //Déclarations et procédures des zones ecran
  if Assigned(GetControl('AFF_STATUTAFFAIRE')) then
  Begin
    StatutAff := THValComboBox(GetControl('AFF_STATUTAFFAIRE'));
    StatutAff.OnExit := StatutExit;
  end;

  if Assigned(getControl('SPREVISIONNEL')) then TRadioButton(GetControl('SPREVISIONNEL')).OnClick := Previsionnelpresent;

  if Assigned(getControl('STOUS')) then TRadioButton(GetControl('STOUS')).OnClick := Previsionnelpresent;


end;

Procedure TOF_AFFAIRE_MUL.ChargeResponsable;
Var lib1, Lib2  : string;
    TitreMess   : string;
begin

  PositionneResponsableUser(TResponsable, True, True, False);

  if LibResponsable = nil then Exit;

  If TResponsable.text <> '' then
  Begin
    LibelleRessource(TResponsable.text, lib1, lib2);
    //
    Tresponsable.ElipsisButton  := False;
    LibResponsable.Visible      := True;
    LibResponsable.caption      := Lib1 + ' '  + lib2;
  end
  else
  Begin
    LibResponsable.Visible      := False;
    Tresponsable.ElipsisButton  := True;
  end;

end;

Procedure TOF_AFFAIRE_MUL.ChargeDomaine;
Var CC : THValCOmboBox;
    Lib1, Lib2  : string;
begin

	// gestion Domaine Activité (BTP)
	CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
	if CC<>Nil then	PositionneDomaineUser(CC) ;

  if LibDomaine = nil then Exit;

  If CC.Value <> '' then
  Begin
    LibelleDomaine(CC.Value, lib1, lib2);
    //
    LibDomaine.Visible := True;
    LibDomaine.caption := Lib1 + ' '  + lib2;
  end
  else
    LibDomaine.Visible    := False;

end;

Procedure TOF_AFFAIRE_MUL.ChargeEtablissement(Abrege : Boolean=False);
Var CC : THValCOmboBox;
    Lib1, Lib2  : string;
begin

	// gestion Etablissement (BTP)
	CC:=THValComboBox(GetControl('AFF_ETABLISSEMENT')) ;

	if CC<>Nil then
  begin
  	PositionneEtabUser(CC) ;
    if not VH^.EtablisCpta then
    begin
    	if LibEtablissement <> nil then LibEtablissement.Visible := false;
			CC.visible := false;
    end;
	end;

  If CC.Value <> '' then
  Begin
    LibelleEtablissement(CC.Value, lib1, lib2);
    //
    LibEtablissement.Visible := True;
    if abrege then
      LibEtablissement.caption := lib2
    else
      LibEtablissement.caption := Lib1;
  end
  else
    LibEtablissement.Visible    := False;

end;

//FV1 : 22/12/2015 - FS#1816 - VEODIS : multicritère contrats, Critère de recherche : Pb affichage date fin (pas dans la bonne case
Procedure TOF_AFFAIRE_MUL.DateFinOnExit(Sender : TObject);
begin

  DateFin.text  := DateAAA.text;
  DateFin_.Text := DateAAA_.text;

end;

Procedure TOF_AFFAIRE_MUL.CREERPAROnClick(Sender : TOBject);
Var StWhere : string;
    Condition : string;
    Long    : Integer;

begin

  Condition := ' AND AND AFF_CREERPAR="TAC"';

  if Assigned(THEdit(GetCOntrol('XX_WHERE'))) then StWhere := THedit(GetControl('XX_WHERE')).Text;

  if CREERPAR.Checked then
    StWhere :=  StWhere + Condition
  else
  begin
    Stwhere := '';
  end;

  if Assigned(THEdit(GetCOntrol('XX_WHERE'))) then THEDIT(GetControl('XX_WHERE')).text := StWhere;

end;

Procedure TOF_AFFAIRE_MUL.ChargeEcranStatutAffaire;
Var StWhere : String;
Begin       

  if Assigned(THEdit(GetCOntrol('XX_WHERE'))) then StWhere := THedit(GetCOntrol('XX_WHERE')).Text;

  //chargement du combo en avec le code statut... (AFFAIRE0)
  MAffaire0 := THValComboBox(GetControl('AFFAIRE0'));
  //
  if Maffaire0 <> nil then
  begin
     MAffaire0.OnChange := PositionneAffaire0;
     Maffaire0.plus     := SetPlusAffaire0;
     MAffaire0.enabled  := ChangeStatut;
  end;

  if statutAff <> nil then StatutAff.Value := Statut;
  //SetControlProperty('AFF_STATUTAFFAIRE','Value', Statut);

  //Si statut Affaire = contrat d'Intervention
	if Statut = 'INT' then   //contrats
  begin
    Ecran.Caption := TraduitGA('Interventions - Contrats');
    SetControlProperty('BInsert','Hint', TraduitGA('Nouveau contrat'));
    SetControlVisible ('BInsert',False);
    sPlus := sPlus + ' AND (CC_LIBRE IN ("BTP","CON"))';
    Statut := 'INT';
    Affaire0 := 'I';
    SetControlText ('TAFF_AFFAIRE', TraduitGa ('Contrat') ) ;
    ETATAFFAIRE.Value := 'ACP';
    ETATAFFAIRE.Plus  := sPlus;
    ETATAFFAIRE.Enabled := True;
    if Maffaire0 <> nil then Maffaire0.Value := Affaire0;
  End Else if Statut = 'APP' then  //appels
  begin
    if fonction = 'AFFECTATION' then
    begin
      Ecran.Caption := TraduitGA('Affectation par Lots des Appels');
      SetControlVisible ('BInsert',False);
      SetControlVisible('BDuplication',False);
      ETATAFFAIRE.Text := 'ECO;ECR';
      ETATAFFAIRE.Enabled := false;
    end Else if fonction = 'REALISATION' then
    Begin
      Ecran.Caption := TraduitGA('Réalisation par lots des Appels');
      BOuvrir1 := TToolbarButton97(ecran.FindComponent('BOuvrir1'));
      BOuvrir1.OnClick := AffectationRessource;
      BOuvrir1.Visible := True;
      BOuvrir1.hint := 'Indiquer la réalisation des appels' ;
      SetControlVisible ('BInsert',False);
      SetControlVisible('BDuplication',False);
      SetControlVisible ('BOuvrir',False);
      SetControlVisible ('BSelectAll',True);
      ETATAFFAIRE.Text := 'AFF;ECR';
      ETATAFFAIRE.Enabled := false;
      SetControlproperty('FListe', 'Multiselection', true);
    End Else if fonction = 'TERMINE' then
    Begin
      Ecran.Caption := TraduitGA('Clôture par Lots des Appels');
      BOuvrir1 := TToolbarButton97(ecran.FindComponent('BOuvrir1'));
      BOuvrir1.OnClick := AffectationRessource;
      BOuvrir1.Visible := True;
      BOuvrir1.hint := 'clôturer les appels' ;
      SetControlVisible ('BInsert',False);
      SetControlVisible('BDuplication',False);
      SetControlVisible ('BOuvrir',False);
      SetControlVisible ('BSelectAll',True);
      SetControlVisible('BDuplication',False);
      ETATAFFAIRE.Text := 'REA';
      ETATAFFAIRE.Enabled := false;
      StWhere := StWhere + ' AND (NOT AFF_ETATAFFAIRE IN ("ANN","CL1","ECR","FIN","TER","FAC"))';
      SetControlproperty('FListe', 'Multiselection', true);
    end Else
    begin
      //FV1 : 30/07/2014 - FS#1154 - MULTIPHONE NETCOM : Pb avec la loupe sur les devis et les appels
      Ecran.Caption := TraduitGA('Appels - Interventions');
      //ETATAFFAIRE.Text := '';
    end;
    //
    SetControlProperty('BInsert','Hint', TraduitGA('Nouvel Appel')) ;
    sPlus := sPlus + ' AND (ISNUMERIC(CC_LIBRE)=1)';
    Statut := 'APP';
    Affaire0 := 'W';
    SetControlText ('TAFF_AFFAIRE', TraduitGa ('Appel') ) ;
    SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
    SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critéres Appel');
    SetControlProperty('AFF_PRIOCONTRAT','Value','');
    ETATAFFAIRE.Plus := sPlus;

    if Maffaire0 <> nil then Maffaire0.Value := Affaire0;
  end else if Statut = 'AFF' then // affaires - chantiers
  begin
     Ecran.Caption := TraduitGA('Chantiers');
     SetControlProperty('BInsert','Hint', TraduitGA('Nouveau Chantier')) ;
     SetControlVisible('BDuplication',True);
     SetControlVisible ('BInsert1',False);
     sPlus := sPlus + ' AND (CC_LIBRE="BTP" AND CC_CODE <> "ATT")';
     Statut := 'AFF';
     Affaire0 := 'A';
     SetControlText ('TAFF_AFFAIRE', TraduitGa ('Chantier') ) ;
     ETATAFFAIRE.Plus := sPlus;
     Thedit(GetControl('AFF_AFFAIRE')).Text := '';
     if Maffaire0 <> nil then Maffaire0.Value := Affaire0;
  end else if statut = 'PRO' then   //Si statut Affaire = Appel d'offre
  begin
    Ecran.Caption := TraduitGA('Etudes Appel d''offres');
    SetControlProperty('BInsert','Hint', TraduitGA('Nouvel appel d''offre')) ;
    SetControlVisible ('BInsert',False);
    SetControlVisible('BDuplication',False);
    SetControlText ('TAFF_AFFAIRE', TraduitGa ('Appel d''offre') ) ;
    SetControlText ('AFF_AFFAIRE', '');
    Affaire0 := 'P';
    if Maffaire0 <> nil then
    begin
      Maffaire0.Value := Affaire0;
      Maffaire0.Text := Affaire0;
    end;
  end else
  begin
    SetControlVisible ('BInsert1',False);
//    sPlus := sPlus + ' AND CC_LIBRE="BTP" ';
    sPlus := sPlus + ' AND CC_LIBRE="BTP"';
    Statut := 'AFF';
    Affaire0 := 'A';
    SetControlText ('TAFF_AFFAIRE', TraduitGa ('Affaire') ) ;
    ETATAFFAIRE.Plus := sPlus;

    if Maffaire0 <> nil then
       begin
       Thedit(GetControl('AFF_AFFAIRE')).Text := '';
       Maffaire0.Value := Affaire0;
       Maffaire0.Text := Affaire0;
       MAffaire0.enabled := ChangeStatut;
       end;
  end;

  if Assigned(THEdit(GetCOntrol('XX_WHERE'))) then THEDIT(GetControl('XX_WHERE')).text := StWhere;

end;

procedure TOF_AFFAIRE_MUL.FlisteDblClick(Sender: TObject);
var Affaire				: String;
		Tiers					: String;
    EAffaire		  : String;
    Action 				: string;
    Statut        : string;
begin

   //SI pas d'enreg double click interdit ==> FV 05112008
	 if FListe.datasource.DataSet.RecordCount = 0  then exit;

   //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
   //Statut := StatutAff.Value;
   Affaire:=Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
   Statut := Copy(Affaire, 1,1);

   //FV1 - 10/04/2017 - FS#2489 - CGV ENERGIE - non visibilité de certaines zones onglet Complément dans un contrat
   if Assigned(StatutAff) then
   begin
     if Statut = 'I' then StatutAff.Value := 'INT'
     else if Statut = 'A' then StatutAff.Value := 'AFF'
     else if statut = 'W' then StatutAff.Value := 'APP';
   end;

   if Fliste.datasource.dataset.FindField ('AFF_ETATAFFAIRE') <> nil then
   begin
	  EAffaire:=Fliste.datasource.dataset.FindField('AFF_ETATAFFAIRE').AsString;
   end else
   begin
	  EAffaire:='';
   end;

	 Tiers:=Fliste.datasource.dataset.FindField('AFF_TIERS').AsString;

   Action:=GetControlText('XXAction');
   if Action = '' then Action := 'ACTION=MODIFICATION';

   if (Action='ACTION=RECH') Or AppelPlanning then
      begin
      TFMul(Ecran).Retour :=Affaire+';'+Tiers;
      TFMUL(Ecran).Close;
      end
   else
      begin
      if (Pos(EAffaire,'TER;CLO')>0) and (not V_PGI.SAV) then Action := 'ACTION=CONSULTATION';
      //if pos('APP',StatutAff.value)>0 then
      if Ecran.Name = 'BTPREVFACT_MUL' then
        AGLLanceFiche ('BTP','BTPREVFAC','','','CODEAFFAIRE=' + Affaire + ';ACTION=MODIFICATION')
      else
      begin
        //FV1 - 20/06/2016 - FS#1968 - ALGO BATI - confusion affaire et appel pour la partie 1 d'un code chantier
      if Statut = 'W' then //pos('W',Affaire)>0 then
         AGLLanceFiche ('BTP','BTAPPELINT','','','CODEAPPEL:' + Affaire + ';ETAT:'+ EAffaire + ';ACTION=MODIFICATION')
      else
         AGLLanceFiche ('BTP',Nomfiche,'',Affaire,'STATUT:'+StatutAff.Value+';ETAT:'+ EAffaire +';'+Action);
      end;
      RefreshDb;
   end;

end;

procedure TOF_AFFAIRE_MUL.VisuAffaireClick(Sender: TObject);
var Affaire				: String;
		Tiers					: String;
    EAffaire		  : String;
    Statut        : string;
begin

  //SI pas d'enreg double click interdit ==> FV 05112008
  if FListe.datasource.DataSet.RecordCount = 0  then exit;

  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  //Statut := StatutAff.Value;
  Affaire:=Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
  Statut := Copy(Affaire, 1,1);

  if Fliste.datasource.dataset.FindField ('AFF_ETATAFFAIRE') <> nil then
  begin
  	EAffaire:=Fliste.datasource.dataset.FindField('AFF_ETATAFFAIRE').AsString;
  end else
  begin
  	EAffaire:='';
  end;

  Tiers:=Fliste.datasource.dataset.FindField('AFF_TIERS').AsString;

  Action:=GetControlText('XXAction');

  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  //if pos('APP',StatutAff.value)>0 then
  if statut = 'W' then
  	AGLLanceFiche ('BTP','BTAPPELINT','','','CODEAPPEL:' + Affaire + ';ETAT:'+ EAffaire + ';ACTION=CONSULTATION')
  else
  	AGLLanceFiche ('BTP',Nomfiche,'',Affaire,'STATUT:'+StatutAff.Value+';ETAT:'+ EAffaire +';ACTION=CONSULTATION');
  refreshdb;

end;

procedure TOF_AFFAIRE_MUL.BDuplication_OnClick(Sender: TObject);
var CodeAffaire : String;
    Fliste	    : THDbGrid;
begin

  Fliste := THdbgrid(GetControl('FLISTE'));

  //SI pas d'enreg double click interdit ==> FV 05112008
  if FListe.datasource.DataSet.RecordCount = 0  then exit;

  codeAffaire := Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;

  AGLLanceFiche('BTP',Nomfiche,'','','ACTION=CREATION;DUPLICATION='+CodeAffaire);

  refreshdb;

end;

procedure TOF_AFFAIRE_MUL.BInsert_OnClick(Sender: TObject);
var natpiece			: String;
    StatutAffaire	: String;
    Statut        : string;
    Affaire       : string;
    EAffaire		: String;
    Tiers,stTiers	: String;
    stCTI 				: string;
    RetourAff     : string;
    ModeGestionCotrait : string;
begin

  ModeGestionCotrait := '';

  stCTI := '';

  //StatutAffaire	:= GetControlText('AFF_STATUTAFFAIRE');
  StatutAffaire := StatutAff.Value;

  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  Statut :=GetControlText('AFF_AFFAIRE0');

  if (statutaffaire='') and (statut<>'') then StatutAffaire := Statut;

  if assigned(EtatAffaire) then EAffaire		:= 'ETAT :' + EtatAffaire.Value;

  //Tiers 			:= 'AFF_TIERS :' + GetControlText('AFF_TIERS');
  
  //FV1 : 10/01/2017 - FS#2144 - ESPACS - Code client non repris en création d'affaire depuis un devis
  Tiers 			:= 'AFF_TIERS=' + GetControlText('AFF_TIERS');

  if fModeGestion = TTgCotraitance then
  begin
    ModeGestionCotrait := ';COTRAITANCE';
  end;

  if (StatutAffaire = 'PRO') then
     natpiece := 'ETU'
  else if (StatutAffaire = 'APP') then
     natpiece := 'DAP'
  else if (StatutAffaire = 'INT') then
     natpiece := 'AFF'
  else
     natpiece := 'DBT';

{$IFDEF CTI}
	stCTI := ';NUMTEL='+NumTel+ ';APPELANT='+Contact;
  stTiers := 'AFF_TIERS=' + GetControlText('AFF_TIERS');
{$ENDIF}
   
  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  if Statut = 'W' then
  begin
  	stTiers := 'CODETIERS=' + GetControlText('AFF_TIERS');
    AGLLanceFiche('BTP','BTAPPELINT','','','ACTION=CREATION;' + stTiers+stCti);
  end else
  begin
  	retourAff := AGLLanceFiche('BTP',NomFiche,'','','ACTION=CREATION;STATUT:'+ StatutAffaire + ';' + Tiers + ';'+ EAffaire + ';NATURE:'+natpiece+ModeGestionCotrait);
    if (GetParamSocSecur ('SO_BTREOUVNEWAFF',False)) and (retouraff <> '') then
    begin
    	StatutAffaire := ReadTokenSt(RetourAff);
    	AGLLanceFiche ('BTP',Nomfiche,'',retouraff,'ACTION=MODIFICATION;'+'STATUT:'+ StatutAffaire + ModeGestionCotrait );
    end;
  end;

	// TtoolBarButton97(GetCOntrol('Bcherche')).Click;
  RefreshDB(retouraff);

end;



procedure TOF_AFFAIRE_MUL.ControleCritere(Critere, Champ, Valeur : string);
Var TobEtatsBloques : TOB;
		I								: Integer;
    Action          : TActionFiche;
    Affaire         : String;
    Aff0            : String;
    Aff1            : String;
    Aff2            : String;
    Aff3            : String;
    Avenant         : String;
Begin

  //Modification pour prise en compte de l'appel de la fiche à partir du planning.
  //Modification du double -click...
  //Possibilité d'avoir une sélection multiple...
  //Modification de la validation..
  if Champ = 'PLANNING' then
  begin
    AppelPlanning := True;
  end;

  if Champ = 'XX_WHERE' then
  begin
    SetControlText('XX_WHERE', Valeur);
  end;

	// MCD 29/09/00 tout revu pour prendre en compte ce qui était fait dans le script
  // Xlequel peut permettre de passer ce code tiers
	// au prog de saisie pour forcer MAJ sur ce code
  // gm 02/09/02 , quand on veut les affaires d'un client, il les faut toutes
  // les  voir qu'elles soient administratives ou non
  if Champ ='AFF_TIERS' then
     begin
     SetControlText('XXlequel' ,Critere);
     SetControltext ('AFF_TIERS',valeur);
   	 SetControlProperty( 'AFF_TIERS', 'operateur', Egal);
     SetControlVisible ('AFF_TIERS',False);
     SetControlVisible ('TAFF_TIERS',False);
     SetControlVisible ('BEffaceaff1',False);
     TCheckBox(GetControl('AFF_ADMINISTRATIF')).State := cbGrayed;
     end;

  if champ ='ACTION' then SetControlText('XXAction',Critere);

  // Modif BTP
  if champ ='ETAT' then
  begin
    if Statut = 'APP' then Valeur := StringReplace(Valeur, ',', ';', [rfReplaceAll]);
    if ETATAFFAIRE <> nil then ETATAFFAIRE.Value := Valeur;
    end;

  // --
  // -- INT = Gestion des contrats
  // -- APP = Gestion des  Appels d'Interventions
  // -- PRO = Gestion des Appels d'Offre
  //sPlus := EtatAffaire.plus;
  if (champ = 'STATUT') then
  begin
  	Statut := Valeur;
    if Assigned(GetControl('TAFF_AFFAIRE')) then
    begin
      if      Valeur = 'AFF' then SetControlText('TAFF_AFFAIRE', 'Code Chantier :')
      Else if Valeur = 'APP' then SetControlText('TAFF_AFFAIRE', 'Code Appel :')
      Else if Valeur = 'INT' then SetControlText('TAFF_AFFAIRE', 'Code Contrat :')
      else if Valeur = 'PRO' then SetControlText('TAFF_AFFAIRE', 'Code Appel d''Offre : ')
      Else                        SetControlText('TAFF_AFFAIRE', 'Code Affaire : ');
    end;
    top := 0;
  end;

  // Réduction de la multicombo des états de l'affaire en fonction des blocages affaire
  if (Champ = 'CTX') and (EtatAffaire <> nil) then
     begin
     TobEtatsBloques := nil;
     try
        TobEtatsBloques := TOBEtatsBloquesAffaire (copy (critere, 5, 3), '-', V_PGI.groupe);
        sPlus := EtatAffaire.plus;
        if (TobEtatsBloques <> nil) then
           for i:=0 to TobEtatsBloques.Detail.count - 1 do
               begin
               sPlus := sPlus  + ' AND CC_CODE<>"' + TobEtatsBloques.Detail[i].GetValue('ABA_ETATAFFAIRE') + '"';
               end;
          EtatAffaire.Plus := sPlus;
     finally
        TobEtatsBloques.Free;
     end;
     end;

  if critere = 'APPOFFACCEPT' then
     begin
     SetControlVisible ('BInsert',False);
     SetControlVisible ('BSelectAll',True);
     SetControlVisible ('BDuplication',False);
{$IFDEF BTP}
     SetControlProperty('AFF_STATUTAFFAIRE','Value','PRO');
{$ENDIF}
     Statut := 'PRO';
     SetControlProperty('Fliste','Multiselection',true);
     end;

  if critere = 'NOCHANGETIERS' then BEGIN SetControlEnabled('AFF_TIERS',False); bChangetiers := False; END;
  if critere = 'MODELEONLY' then SetControlProperty ('AFF_MODELE','Checked',True);
  if critere = 'ADMINGRAYED' then TCheckBox(GetControl('AFF_ADMINISTRATIF')).State := cbGrayed;
  if critere = 'NOFILTRE' then   Tfmul(Ecran).FiltreDisabled:=true;
  if critere = 'PASSIUN' then   bPasSiUn := true;
  if critere = 'ITDATESMAX' then   bItDatesMax := true;
  if critere = 'REDUIT' then
     if Not(GetParamsoc('SO_AfSaisAffInterdit')) then VUserInvite := True; //mcd 18/10/02 ajout test paramsoc

  if assigned(GetControl('AFF_MANDATAIRE')) then
  begin
    if Champ = 'MANDATAIRE' then
    begin
      if Valeur = 'X' then
        THDBCheckbox (GetControl('AFF_MANDATAIRE')).State := cbchecked
      else
        THDBCheckbox (GetControl('AFF_MANDATAIRE')).State := cbUnchecked;
    end;
  end;

  if Champ = 'COTRAITANCE'   then
  begin
    if assigned(GetControl('XX_WHERE')) then Setcontroltext('XX_WHERE', 'AND AFF_MANDATAIRE IN ("X", "-")');
    fModeGestion := TTgCotraitance;
  end;

  if critere = 'NOCHANGESTATUT' then
     BEGIN
     ChangeStatut := False;
     top :=1;
  END;

  //FV1 : 17/11/2014 - FS#1309 - DELABOUDINIERE : la recherche des affaires ne permet plus l’accès à la sélection du type d’affaire;
  //tout le else était en commentaire il suffit de tester si TOUS est chargé (????)
  if champ = 'TOUS' then
  begin
     ChangeStatut := True;
     top :=1;
  end;
  {*if Statut='' then
  begin
     ChangeStatut := True;
     top :=1;
  end;*}

  if Statut = 'NOAFFETAT' then
		 AffEtat := false
  else
	   AffEtat := True;

	// gm 04/07/02 pour gerer le mul avec stat client
  if critere = 'PASMULTI'    then
     begin
     TFMul(Ecran).caption:=TraduitGa('Affaires (avec critères clients)');
     SetControlChecked('Multi',false);
     SetControlVisible('PCLIENTSTAT',true);
     TFMUL(Ecran).Dbliste:='AFMULAFFAIREMULTI';
     if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'AFMULAFFAIREMULTI';
     UpdateCaption(TFMul(Ecran)) ;
     end ;

  if critere = 'AFFECTATION' then
     begin
     Fonction := critere;
     BOuvrir1 := TToolbarButton97(ecran.FindComponent('BOuvrir1'));
     BOuvrir1.OnClick := AffectationRessource;
     SetControlVisible ('BOuvrir',False);
     BOuvrir1.Visible := True;
     SetControlVisible ('BSelectAll',True);
     ETATAFFAIRE.Value := 'ECO;';
     ETATAFFAIRE.Enabled := false;
     SetControlproperty('FListe', 'Multiselection', true);
     end;

  if critere = 'REALISATION' then  Fonction := critere;

  if critere = 'TERMINE' then Fonction := critere;
{$IFDEF CTI}
  if champ = 'APPELANT' then Contact := Valeur;
  if champ = 'NUMTEL' then Numtel := Valeur;
{$ENDIF}
  BInsert := TToolbarButton97(ecran.FindComponent('BInsert'));
  BInsert.OnClick := BInsert_OnClick;
  //
  if assigned(GetControl('BDUPLICATION')) then
  begin
    BDuplication := TToolbarButton97(ecran.FindComponent('BDuplication'));
    BDuplication.OnClick := BDuplication_OnClick;
  end;
  //
  if TToolBarButton97(GetControl('BVISUAFFAIRE'))<> nil then
  begin
  	TToolBarButton97(GetControl('BVISUAFFAIRE')).OnClick := VisuAffaireClick;
  end;
  
//uniquement en line
//	if TTabSheet(GetControl('PCOMPLEMENT'))<> nil then TTabSheet(GetControl('PCOMPLEMENT')).tabVisible := false;

  if champ = 'MODEPC' then // mode de gestion du plan de charge
  begin
    if valeur = 'CHANTIER' then fModeVisuPLanCharge := TmPcPlanCharge
    else if valeur = 'FONCTION' then fModeVisuPLanCharge := TmcPlanCharFnc;
  end;
End;

procedure TOF_AFFAIRE_MUL.OnUpdate;
{$IFDEF BTP}
var
  i : integer;
  StChamp : String;
  StChamp2: String;
  NumPart : Integer;
  ind     : Integer;
  lg      : integer;
  Libelle : String;
  Nomchamp: String;
{$ENDIF}
Begin
// Gestion repositionnement auto sur l'affaire en cours si sortie rapide / bug par prg ( ou Eagl)

inherited;
if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('AFF_AFFAIRE2','');
if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('AFF_AFFAIRE3','');

{$IFDEF BTP}
if EtatAffaire <> nil then                   
begin
  if EtatAffaire.Value = 'ACP' then
  begin
  {$IFDEF EAGLCLIENT}
    for i:=0 to TFMul(Ecran).FListe.ColCount-1 do
    begin
  // AFAIREEAGL
    end;
  {$ELSE}
  For i:=0 to TFMul(Ecran).FListe.Columns.Count-1 do
    Begin
    If (TFMul(Ecran).FListe.Columns[i].FieldName = 'AFF_REFERENCE') Then
      TFMul(Ecran).FListe.columns[i].visible:=False ;
    end;
  {$ENDIF}
  end;
end;
{$ENDIF}

 {*
  for i := 0 to TFMul(Ecran).fliste.Columns.Count-1 do
  begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(TFMul(Ecran).FListe.Columns[i].FieldName);
    StChamp2 := TFMul(Ecran).FListe.Columns[i].FieldName;
    lg := Length(TFMul(Ecran).Fliste.Columns[i].FieldName);
    ind := Pos ('_',TFMul(Ecran).Fliste.Columns[i].FieldName);
    if ((ind = 0) or (lg =0)) then  continue;
    if (TFMul(Ecran).Fliste.Columns[i].Field)=Nil then Continue;
    Nomchamp := Copy (TFMul(Ecran).Fliste.Columns[i].FieldName,ind+1, lg-ind);
    if Nomchamp = 'PREVISIONNEL' then
    begin
      TFMul(Ecran).fliste.Columns[i].ColFormats := 'CB=BTCTRLPREVFACT';
      Break;
    end;
  end;  *}


End;

procedure TOF_AFFAIRE_MUL.TraiteRetour;
Var stRetour : string;
    Champ : string;
begin
stRetour := GetControlText ('RETOUR');
if stRetour <> '' then
   begin
   // On récupére dans l'ordre le client + Affaire 1,2,3
   Champ:=(Trim(ReadTokenSt(stRetour)));
   if Champ <> '' then SetControltext('AFF_TIERS',Champ);
   Champ:=(Trim(ReadTokenSt(stRetour)));
   if Champ <> '' then SetControltext('AFF_AFFAIRE1',Champ);
   Champ:=(Trim(ReadTokenSt(stRetour)));
   if (Champ <> '') And (VH_GC.CleAffaire.Co2Visible) then SetControltext('AFF_AFFAIRE2',Champ);
   Champ:=(Trim(ReadTokenSt(stRetour)));
   if (Champ <> '') And (VH_GC.CleAffaire.Co3Visible) then SetControltext('AFF_AFFAIRE3',Champ);
   end;
SetControlText('RETOUR', '');
end;


procedure TOF_AFFAIRE_MUL.OnLoad;
Var //RazAffRef : Boolean;
		//Affaire0 : string;
    Affaire, Affaire1  ,Affaire2 ,Affaire3,Avenant,Aff0 : string;
begin
inherited;
Inc (GNbLoads);

if VUserInvite and ( not V_PGI.Superviseur) and ( not V_PGI.Controleur) then TraiteInvite;

if GereSousAffaire then
   begin
     if (GetControl('AFFAIREREF1')<>NIL) then  // gm 25/08/03
     begin
      if (GetControlText('AFFAIREREF1')='') then
        SetControlText('AFF_AFFAIREREF','')
      else
        begin
        Aff0 := Affaire0;
        Affaire1 := GetControlText('AFFAIREREF1');
        Affaire2 := GetControlText('AFFAIREREF2');
        Affaire3 := GetControlText('AFFAIREREF3');
        Avenant := GetControlText('AFFAIREREF4');
        Affaire := CodeAffaireRegroupe (Aff0, Affaire1, Affaire2, Affaire3, Avenant, taModif, false, false, false);
        if not ExisteAffaire (Affaire, '') then SetControlText('AFF_AFFAIREREF',Trim(Copy(Affaire,1,15)));
        end;
     end;
   end;

// PL le 26/06/03 : en scot on agrandit l'intervalle quand le client est préselectionné et seulement
// la première fois
if (ctxScot In V_PGI.PGIContexte) and (GetControlText('AFF_TIERS') <> '') and (GNbLoads <= 1) then
  begin
    SetControlText ('AFF_DATEDEBUT', datetostr (idate1900));
    SetControlText ('AFF_DATEDEBUT_', datetostr (idate2099));
  end;

end;

procedure TOF_AFFAIRE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
// MODIF LS
Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));
// --
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
// affaire de référence pour recherche
Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
Aff1_:=THEdit(GetControl('AFFAIREREF1'));
Aff2_:=THEdit(GetControl('AFFAIREREF2'));
Aff3_:=THEdit(GetControl('AFFAIREREF3'));
Aff4_:=THEdit(GetControl('AFFAIREREF4'));
end;

procedure TOF_AFFAIRE_MUL.TraiteInvite;
var
lstzone,zone : string;
begin
  lstzone := 'AFF_RESPONSABLE;AFF_RESSOURCE1;AFF_RESSOURCE2;AFF_RESSOURCE3';
  while lstzone <> '' do
    begin
     zone := ReadTokenSt(Lstzone);
     if ThEdit(GetControl(zone))<> nil then
      begin
        SetControlText(zone,'');
        SetControlEnabled(zone,True);
      end;
    end;
  //ThDBGrid(GetControl('FListe')).OnDblClick := nil;
  TToolBarButton97(GetControl('BInsert')).enabled := False;
  TToolBarButton97(GetControl('BInsert1')).enabled := False;
  TToolBarButton97(GetControl('BVOIRAFFAIRE')).enabled := False;
  TToolBarButton97(GetControl('BDuplication')).enabled := False;
end;

procedure TOF_AFFAIRE_MUL.StatutExit(Sender : TObject);
Begin

	if StatutAff.value='PRO' then
		 begin
     SetcontrolVisible('TAFF_ETATAFFAIRE',True);
    EtatAffaire.Visible := True;
   	 end
	else
		 begin
     SetcontrolVisible('TAFF_ETATAFFAIRE',False);
    EtatAffaire.Visible := True;
    //EtatAffaire.value := '';
     end;

  //ChargeEcranStatutAffaire;

end;

Procedure TOF_AFFAIRE_MUL.AffectationRessource(Sender : TOBJect);
var Arg,Affaire      : String;
    TOBparam : Tob;
    TOB1Appel: Tob;
    TOBDAppel: Tob;
    F 			 : TFMul;
    i 			 : integer;
    QQ			 : TQuery ;
{$IFDEF EAGLCLIENT}
    L 			 : THGrid;
{$ELSE}
    Ext 		 : String;
    L 			 : THDBGrid;
{$ENDIF}

begin
  Inherited ;

  Arg := '';
  Ext := '';

	F:=TFMul(Ecran);

	if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné','');
    exit;
  end;

  if Fonction = 'AFFECTATION' then
  Begin
    TOBParam := TOB.Create ('AFFECT ',nil,-1);
    TOBParam.AddChampSupValeur ('RESSOURCE','');
    TOBParam.AddChampSupValeur ('RETOUR','-');
    TheTOB := TOBparam;
    AGLLanceFiche ('BTP','BTAFFECTATION','','','');
    TheTOB := nil;
    if TOBparam.GetValue('RETOUR')='-' then
    Begin
      TobParam.free;
      exit ;
    end;
  end;

	if PGIAsk('Confirmez-vous le traitement ?','') <> mrYes then
  Begin
    TobParam.free;
  	exit ;
  end;

  TOBDAppel := tob.create('LES APPELS', nil, -1);

	L:= F.FListe;
  SourisSablier;

  try
	if L.AllSelected then
  begin
  QQ:=F.Q;
  QQ.First;
  while Not QQ.EOF do
  Begin
    Affaire := QQ.FindField('AFF_AFFAIRE').AsString;
    TOB1Appel := TOB.Create ('AFFAIRE',TOBDappel,-1);
    TOB1Appel.putValue('AFF_AFFAIRE',Affaire);
    if TOB1Appel.LoadDb (false) then
    begin
    if Fonction = 'AFFECTATION' then
       Begin
       TOB1Appel.PutValue('AFF_RESPONSABLE',TOBParam.getValue ('RESSOURCE'));
       TOB1Appel.putValue('AFF_ETATAFFAIRE','AFF');
       End
    else if Fonction = 'REALISATION' then
       Begin
       TOB1Appel.putValue('AFF_ETATAFFAIRE','REA');
       TOB1Appel.putValue('AFF_DATEFIN', now);
       end
    else if Fonction = 'TERMINE' then
       Begin
       TOB1Appel.putValue('AFF_ETATAFFAIRE','CL1');
       TOB1Appel.putValue('AFF_DATEFIN', now);
       end;
    end
    else
      TOB1Appel.free;
    QQ.Next;
  end;
  L.AllSelected:=False;
  end
  else
     begin
     for i:=0 to L.NbSelected-1 do
         begin
	       L.GotoLeBookmark(i);
         Affaire := TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
         TOB1Appel := TOB.Create ('AFFAIRE',TOBDappel,-1);
         TOB1Appel.putValue('AFF_AFFAIRE',Affaire);
         if TOB1Appel.LoadDb (false) then
            Begin
            if Fonction = 'AFFECTATION' then
         	 	   begin
	             TOB1Appel.PutValue('AFF_RESPONSABLE',TOBParam.getValue ('RESSOURCE'));
             	 TOB1Appel.putValue('AFF_ETATAFFAIRE','AFF');
            	 end
            else if Fonction = 'REALISATION' then
               begin
   	           TOB1Appel.putValue('AFF_ETATAFFAIRE','REA');
	             TOB1Appel.putValue('AFF_DATEFIN', now);
               end
            else if Fonction = 'TERMINE' then
               Begin
               TOB1Appel.putValue('AFF_ETATAFFAIRE','CL1');
               TOB1Appel.putValue('AFF_DATEFIN', now);
               end;
            end
         else
            TOB1Appel.free;
      	 end;
	   		 L.ClearSelected;
	   end;
  if TOBDappel.detail.count > 0 then TOBDappel.UpdateDB (true);
	finally
  TOBDappel.free;
  TobParam.free;
	SourisNormale;
  F.BChercheClick(ecran);
  end;

end;

//******************************************************************************
//************************ TOF Multicritères apporteur *************************
//******************************************************************************
procedure TOF_APPORTEUR_MUL.OnArgument(stArgument : String );

Begin
Inherited;
  if ctxScot In V_PGI.PGIContexte then
    BEGIN
    SetcontrolVisible('TGCL_TYPECOMMERCIAL',False);
    SetcontrolVisible('GCL_TYPECOMMERCIAL',False);
    SetcontrolVisible('TGCL_ZONECOM',False);
    SetcontrolVisible('GCL_ZONECOM',False);
    End;

  // Type commercial forcé à apporteur gm25/08/03
  SetControlText('GCL_TYPECOMMERCIAL', 'APP');
END;


//******************************************************************************
//************* Fonction appellées depuis le Scripts ***************************
//******************************************************************************
procedure MajMulAffaire( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
if (LaTof is TOF_AFFAIRE_MUL) then TOF_AFFAIRE_MUL(LaTof).TraiteRetour else exit;
end;


procedure AglChangeTVAContrats( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
if (LaTof is TOF_AFFAIRE_MUL) then TOF_AFFAIRE_MUL(LaTof).LanceChangeTvaContrats else exit;
end;


Procedure AFLanceFiche_MulEAffaire;
begin
AGLLanceFiche ('AFF','AFEAFFAIRE_MUL','','','');
end;
Procedure AFLanceFiche_MulAffaire(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFFAIRE_MUL',Range,'',Argument);
end;
Function AFLanceFiche_AffaireREch(Range,Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFFAIRERECH_MUL',Range,'',Argument);
end;
Procedure AFLanceFiche_MulAffCompta;
begin
AGLLanceFiche ('AFF','AFFAIRE_MUL_COMPT','','','');
end;
Procedure AFLanceFiche_Mul_Apporteur;
begin
AGLLanceFiche ('AFF','AFAPPORTEUR_Mul','','','');
end;
Procedure AFLanceFiche_MulAvenantAff(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFFAIREAVT_MUL',Range,'',Argument);
end;
Procedure AFLanceFiche_IntervAff;
begin
AGLLanceFiche ('AFF','AFINTERVENANTSAFF','','','');
end;



procedure TOF_AFFAIRE_MUL.PositionneAffaire0(Sender: Tobject);
var splus : string;
begin

	THEdit(GetControl('AFF_AFFAIRE0')).Text := mAffaire0.Value;

	if Length(THEdit(GetControl('AFF_AFFAIRE')).Text) <= 1 then
     begin
  	 THEdit(GetControl('AFF_AFFAIRE')).Text := '';
     if maffaire0.value = 'I' then
        begin
        SetControlProperty('AFF_STATUTAFFAIRE','Value','INT');
        sPlus := sPlus + ' AND (CC_LIBRE IN ("BTP","CON"))' ;
        SetControlProperty('AFF_ETATAFFAIRE','Value','ENC');
        end
     else if maffaire0.value = 'W' then
        begin
        SetControlProperty('AFF_STATUTAFFAIRE','Value','APP');
        sPlus := sPlus + ' AND (ISNUMERIC(CC_LIBRE)=1)';
        SetControlProperty('AFF_ETATAFFAIRE','Value','AFF');
        end
     else if maffaire0.value = 'P' then
        begin
        SetControlProperty('AFF_STATUTAFFAIRE','Value','PRO');
        sPlus := sPlus + ' AND (CC_LIBRE="BTP")';
        SetControlProperty('AFF_ETATAFFAIRE','Value','ENC');
        end
     else if maffaire0.value = 'A' then
        begin
      	SetControlProperty('AFF_STATUTAFFAIRE','Value','AFF');
        sPlus := sPlus + ' AND (CC_LIBRE="BTP" AND CC_CODE <> "ATT")';
        SetControlProperty('AFF_ETATAFFAIRE','Value','');
        end;
     SetControlProperty ('AFF_ETATAFFAIRE','Plus', sPlus);
//     THValComboBox(GetControl('AFF_ETATAFFAIRE')).ReLoad ;
     end;

end;

Procedure TOF_AFFAIRE_MUL.ExitResponsable(Sender: TObject);
Var Lib1, Lib2 : string;
begin

  if LibResponsable = nil then Exit;

  IF TResponsable.text <> '' then
  begin
    LibelleRessource(TResponsable.text, lib1, lib2);
    LibResponsable.Visible := True;
    LibResponsable.caption := Lib1 + ' ' + Lib2;
  end
  else
    LibResponsable.Visible := False;

end;

Procedure TOF_AFFAIRE_MUL.ExitTiers(Sender: TObject);
begin

  if LibClient = nil then Exit;

  IF TTiers.text <> '' then
  begin
    LibClient.caption := RecupLibelleTiers(TTiers.text);
    LibClient.Visible := True;
  end
  else
    LibClient.Visible := False;

end;


procedure TOF_AFFAIRE_MUL.BRechResponsable(Sender: TObject);
Var Lib1, Lib2 : String;
begin

  GetRessourceRecherche(TResponsable,'ARS_TYPERESSOURCE="SAL"', '', '');

  if LibResponsable = nil then Exit;

  IF TResponsable.text <> '' then
  begin
    LibelleRessource(TResponsable.text, lib1, lib2);
    LibResponsable.Visible := True;
    LibResponsable.caption := Lib1 + ' ' + Lib2;
  end
  else
    LibResponsable.Visible := False;

end;


procedure TOF_AFFAIRE_MUL.LanceChangeTvaContrats;

	function ChargePiece (TOBpiece,TOBAffaire: TOB; var DEV : Rdevise; Affaire : string) : boolean;
  var QQ : TQuery;
  		StatutAffaire,NaturePiece : string;
      okOk : Boolean;
  begin
    Result := false;
    OKOK := false;
		QQ := OpenSQL('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"',True,1,'',true);
    if Not QQ.eof then
    begin
			StatutAffaire:=QQ.findField('AFF_STATUTAFFAIRE').AsString;
      TOBAffaire.SelectDB('',QQ);
      OkOk := true;
    end;
    ferme (QQ);
    if not Okok then Exit;
    //
    If (StatutAffaire='AFF') Then
       NaturePiece:= VH_GC.AFNatAffaire
    {$IFDEF BTP}
    else if (StatutAffaire='INT') Then
       NaturePiece:= 'AFF'
    else if (StatutAffaire='APP') Then
       NaturePiece:= 'DAP'
    {$ENDIF}
    Else
       NaturePiece:= VH_GC.AFNatProposition;
    QQ := OpenSQL('SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+NaturePiece+'" AND GP_AFFAIRE="'+Affaire+'"',True,1,'',true);
    if Not QQ.eof then
    begin
      TOBpiece.SelectDB('',QQ);
      Result := True;
    end;
    ferme (QQ);
    FillChar (DEV,SizeOf(DEV),#0);
    if Result then
    begin
    	DEV.Code := TOBpiece.GetString('GP_DEVISE');
      GetInfosDevise (DEV) ;
    end;
  end;

  procedure EnregMessage(TOBPiece,TOBAffaire : TOB);
  var TobJnal: TOB;
    Qry: TQuery;
    NumEvt: integer;
    Nature, ActionSurPiece, TypeEvt: string;
    BlocNote: TStringList;
  begin
    Nature := RechDom('GCNATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'), False);
    if Nature = '' then nature := 'Détail Contrat';
    BlocNote := TStringList.Create;
    BlocNote.Add(Nature + TraduireMemoire(' numéro ') + IntToStr(TOBPiece.GetValue('GP_NUMERO')));
    TobJnal := TOB.Create('JNALEVENT', nil, -1);
    TobJnal.PutValue('GEV_TYPEEVENT', 'ATR');
    TobJnal.PutValue('GEV_LIBELLE', 'Chang. TVA / affaire '+BTPCodeAffaireAffiche(TOBAffaire.GetString('AFF_AFFAIRE')));
    TobJnal.PutValue('GEV_DATEEVENT', Date);
    TobJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
    TobJnal.PutValue('GEV_ETATEVENT', 'OK');
    TobJnal.PutValue('GEV_BLOCNOTE', BlocNote.Text);
    Qry := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
    if not Qry.EOF then
      NumEvt := Qry.Fields[0].AsInteger
    else
      NumEvt := 0;
    Ferme(Qry);
    Inc(NumEvt);
    TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
    TobJnal.InsertDB(nil);
    TobJnal.Free;
    BlocNote.Free;
  end;


var F : TFMul;
		L : THDBGrid;
    NomChoixCod,newCode,Affaire : string;
    QQ : TQuery;
    TOBpiece,TOBAffaire : TOB;
    II : Integer;
    RecalcEches : boolean;
    DEV : RDevise;
begin
  RecalcEches := false;
  If THCheckBox(GetControl('CBRECALCECHES')) <> nil then
  begin
    RecalcEches := THCheckBox(GetControl('CBRECALCECHES')).Checked;
  end;

	F:=TFMul(Ecran);

	if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné','');
    exit;
  end;
  TOBpiece := TOB.Create('PIECE',nil,-1);
  TOBAffaire := TOB.Create('AFFAIRE',nil,-1);
  TRY
    NomChoixcod := 'CHOIXCOD';
    QQ := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="TTTVA"', True);
    if not QQ.EOF then
    begin
      NomChoixCod := QQ.FindField('DS_NOMBASE').AsString+'.DBO.CHOIXCOD';
    end;
    Ferme (QQ);

    NewCode := Choisir('Base de T.V.A', NomChoixcod, 'CC_LIBELLE', 'CC_CODE', 'CC_TYPE="TX1"', 'CC_CODE');
    if NewCode = '' then Exit;
    if PgiAsk ('Vous allez appliquer cette base de T.V.A sur les contrats sélectionnés.#13#10Confirmez-vous ?')=mryes then
    begin
      L:= F.FListe;
      SourisSablier;

      if L.AllSelected then
      begin
        QQ:=F.Q;
        QQ.First;
        while Not QQ.EOF do
        Begin
          TOBpiece.InitValeurs();
          TOBAffaire.InitValeurs();

          Affaire := QQ.FindField('AFF_AFFAIRE').AsString;
          if ChargePiece (TOBpiece,TOBAffaire,DEV,Affaire) then
          begin
            AffecteTvaPiece (TOBpiece,NewCode);
            if RecalcEches then
            begin
							LanceRecalcEches (TOBAffaire,DEV);
            end;
            EnregMessage (TOBPiece,TOBAffaire);
          end;
          QQ.Next;
        end;
      	L.AllSelected:=False;
        PgiInfo ('Traitement terminé');
      end
      else
      begin
        for II :=0 to L.NbSelected-1 do
        begin
					TOBpiece.InitValeurs();
          TOBAffaire.InitValeurs();
          L.GotoLeBookmark(II);
          Affaire := TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
          if ChargePiece (TOBpiece,TOBAffaire,DEV,Affaire) then
          begin
            AffecteTvaPiece (TOBpiece,NewCode);
            if RecalcEches then
            begin
							LanceRecalcEches (TOBAffaire,DEV);
            end;
            EnregMessage (TOBPiece,TOBAffaire);
          end;
        end;
        L.ClearSelected;
        PgiInfo ('Traitement terminé');
      end;
    end;
  FINALLY
    TOBPiece.free;
    TOBAffaire.Free;
  	SourisNormale;
	  refreshdb;
	END;

end;

procedure TOF_AFFAIRE_MUL.ChantierSelect(Sender: TObject);
var Q :Tquery;
    OneTOB : TOB;
    i : integer;
    bCont : boolean;
    WhereSql : string;
begin
  bCont := false;
  if assigned(fTOBSelect) then
  begin
    fTOBselect.ClearDetail;
    fTOBSELECT.SetString('TOUS','-');
    fTOBSELECT.SetString('CRITERES','');
    fTOBSELECT.SetString('TYPERESSOURCES',GetControlText('TYPERESSOURCE') );
    fTOBSELECT.SetString('ETABLISSEMENT',GetControlText('AFF_ETABLISSEMENT') );
    //
    if fliste.AllSelected then
    begin
      fTOBSELECT.SetString('TOUS','X');
      WhereSql := RecupWhereCritere (TPageControl(GetCOntrol('PAGES')));
      fTOBSELECT.SetString('CRITERES',WhereSql);
      bCont := true;
    end else
    begin
      if Fliste.nbSelected > 0 then
      begin
        for i:=0 to Fliste.nbSelected-1 do
        begin
          Fliste.GotoLeBookmark(i);
          OneTOB := TOB.Create ('UN CHANTIER',fTOBSelect,-1);
          OneTOB.AddChampSupValeur('CHANTIER', Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString);
        end;
        bCont := true;
      end;
    end;
    if bCont  then
    begin
      OpenPlanningBtp (fModeVisuPLanCharge,fTOBSelect) ;
      if fliste.NbSelected >0 then fliste.ClearSelected
      else if fListe.AllSelected then TToolbarButton97 (GetControl('bSELECTALL')).Click;
    end;
  end;
end;

procedure TOF_AFFAIRE_MUL.OnClose;
begin
  fTOBSelect.free;
  inherited;
end;

procedure TOF_AFFAIRE_MUL.Previsionnelpresent(Sender: TObject);
begin
  ThEdit(getControl('XX_WHERESEL')).Text := '';
  if TRadioButton(GetControl('SPREVISIONNEL')).Checked then
  begin
    ThEdit(getControl('XX_WHERESEL')).Text := ' AND EXISTS (SELECT 1 FROM PIECE WHERE GP_NATUREPIECEG="PBT" AND GP_AFFAIRE=AFF_AFFAIRE)';
  end;
end;


Initialization
registerclasses([TOF_AFFAIRE_MUL]);
registerclasses([TOF_APPORTEUR_MUL]);
RegisterAglProc( 'MajMulAffaire', TRUE , 0, MajMulAffaire);
RegisterAglProc( 'BTChangeTVAContrats', TRUE , 0, AglChangeTVAContrats);
end.
