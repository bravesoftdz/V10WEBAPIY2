{***********UNITE*************************************************
Auteur  ...... : ??
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : TOM de la table DPSOCIAL
Mots clefs ... :

LM20070404 : TabSheet pGeneral tout le temps

*****************************************************************}
Unit dpTOMSocial ;

Interface

Uses
   UTOM, db, HCtrls, Classes, sysutils, UTOB,
   HEnt1, dbctrls, controls, htb97,
{$IFDEF EAGLCLIENT}
   eFiche,
   MaineAGL,
{$ELSE}
   Fiche,
   FE_Main,
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
   ParamSoc,
   EventDecla
   ;

Type
   TOM_DPSOCIAL = Class (TOM)
        procedure OnArgument(sParam_p : String )  ; override;
        procedure OnNewRecord  ; override ;
        procedure OnLoadRecord; override ;
        procedure OnChangeField(fChamp_p : TField); override;
        procedure OnUpdateRecord; override;
        procedure OnAfterUpdateRecord; override;
        procedure OnClose; override;//LM20070404

      protected

        procedure DFI_TAXESALAIRES_OnClick(Sender : TObject);
        procedure DSO_PAIECAB_OnClick(Sender : TObject);
        procedure DSO_PAIEENT_OnClick(Sender : TObject);
        procedure DSO_CONVENCOLLEC_OnClick(Sender : TObject);
        procedure BWebDiodeClick (Sender:TObject);
        procedure showGED (Sender: TObject);//LMO20060901
        procedure RAZLookup (Sender: TObject);//LMO20060901
        procedure DSO_CONJOINTAVEC_OnClick (sender:TObject);//LM20070404
        procedure BLIENS_OnClick(sender:TObject); //MD 20070524
        procedure bValiderOnClick (sender: TObject);//LM20070712

      private

        LoadRecord : boolean ; //LM20070404
        sAction_c, sGuidPer_c, sNoDossier_c : string;
        bVoirDiode : Boolean;
        bPaie_c    : boolean;
        iMaxCaisses_c  : integer;
        m_bBaseExists  :boolean;
        EvDecla : TEventDecla ;//LM20070404
        DFI_TAXESALAIRESini : string; //LM20070404
        DefTabSheet :  string ;

        procedure Initialisation(sGuidPer_p : string);
        function  InitEffectifs(sNoDossier_p : string;
                                dtDateFinEx_p, dtExSocial_p : TDateTime;
                                var dEffectif_p : Double;
                                var dtDateEffectif_p, dtDateFin_p : TDateTime) : boolean;
        function  DonnerCodeDiode (UnCodeConvention : String): String;


   end;


/////////// IMPLEMENTATION ////////////
Implementation

uses
   dpOutils, DpJurOutils, DpJurOutilsGrille, web, galOutil, EntDp,
   extctrls, UGedViewer, hMsgBox, stdCtrls, annoutils;//LMO20060901

{***********A.G.L.Privé.*****************************************
Auteur  ...... : ??
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.OnArgument(sParam_p : String );
var
    bVoirLiens : Boolean;

    procedure initSocGed (nom:string);//LM20060901
    var ctrl : TControl ;
    begin
      ctrl:=GetControl('RAZ' + nom);
      if ctrl<>nil then TImage(ctrl).OnClick:=RAZLookup ;
      ctrl:=GetControl('GED' + nom);
      if ctrl<>nil then TImage(ctrl).OnClick:=showGed ;
    end ;

begin
  LoadRecord:=true ;//LM20070404

  Inherited ;

  DefTabSheet := ArgumentToTabsheet (sParam_p) ;//LM20070528

  EvDecla := TEventDecla.create(Ecran) ;//LM20070404
  EvDecla.Rebranche('DFI_TAXESALAIRES', 'OnClick', DFI_TAXESALAIRES_OnClick); //LM20070404
  EvDecla.Rebranche('DSO_CONJOINTAVEC', 'OnClick', DSO_CONJOINTAVEC_OnClick); //LM20070404
  EvDecla.Rebranche('DSO_PAIECAB',      'OnClick', DSO_PAIECAB_OnClick);
  EvDecla.Rebranche('DSO_PAIEENT',      'OnClick', DSO_PAIEENT_OnClick);
  EvDecla.Rebranche('DSO_CONVENCOLLEC', 'OnClick', DSO_CONVENCOLLEC_OnClick);

  EvDecla.Rebranche('bValider', 'OnClick', bValiderOnClick);
  
  m_bBaseExists := TRUE;
  iMaxCaisses_c := 10;

  sAction_c := ReadTokenSt(sParam_p);
  ReadTokenSt(sParam_p);
  ReadTokenSt(sParam_p);
  sGuidPer_c := ReadTokenSt(sParam_p);
  if sGuidPer_c = '' then
     sGuidPer_c := TFFiche(Ecran).FLequel;

  // $$$ JP 21/07/06: savoir s'il y a une base ou pas
  if ReadTokenSt (sParam_p) = 'SANSBASE' then
     m_bBaseExists := FALSE;

{$IFDEF DP}
  bVoirDiode := VH_DP.SeriaLienDiode;
{$ELSE}
  bVoirDiode := False;
{$ENDIF}

  Initialisation(sGuidPer_c);
  Ecran.Caption := 'Social : ' + GetNomPer(sGuidPer_c);
  UpdateCaption(Ecran);

  bVoirLiens := JaiLeDroitConceptBureau(ccVoirLesLiens);
  SetControlVisible('BLIENS', bVoirLiens);

  SetControlVisible('LCODEDIODE',   bVoirDiode);
  SetControlVisible('DSO_ORIGDADS', bVoirDiode);
  SetControlVisible('BWEBDIODE',    bVoirDiode);

  if bVoirDiode then
    TToolbarButton97 (GetControl ('BWEBDIODE')).OnClick := BWebDiodeClick
  else if not vh_dp.Group then //LM20060901
  begin
    // Recalage à gauche
    SetControlProperty('TDSO_CONVENCOLLEC', 'Left', 20) ;
    SetControlProperty('DSO_CONVENCOLLEC', 'Left', 100) ;
  end;

  if vh_dp.group then //LMO20060901
  begin
    SetControlVisible('BINSERT', false);//LMO>MD  : pourquoi la possibilité de créer?
    initSocGed('DSO_GEDPARTICIPAT') ;
    initSocGed('DSO_ORIGDADS') ;
    initSocGed('DSO_GEDCOMITEENT') ;
  end ;

  // #### MD 20070524 - En attendant refonte des liens
  if GetControl('BLIENS')<>Nil then
  begin
    TToolbarButton97(GetControl('BLIENS')).OnClick := BLIENS_OnClick;
    TToolbarButton97(GetControl('BLIENS')).Hint := 'Liens organismes sociaux';
  end;

//  LoadRecord:=false;//LM20070404
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_DPSOCIAL.OnNewRecord;
var
   dEffectif_l : Double;
   dtDateFinEx_l, dtDateFin_l, dtExSocial_l, dtDateEffectif_l : TDateTime;
   UnCodeDiode  : String;
begin
  inherited;
  SetField('DSO_GUIDPER', sGuidPer_c);

   // Traitement de la paie
   if bPaie_c then
   begin
      SetField('DSO_PAIECAB', 'X');
      SetField('DSO_PAIEENT', '-');

      if (GetParamSocDossier('SO_PGTYPECDETICKET', 'DB'+sNoDossier_c) <> '') then
          SetField('DSO_EXISTTICKREST', 'X')
      else
          SetField('DSO_EXISTTICKREST', '-');

      if (GetParamSocDossier('SO_PGDECALAGE', 'DB'+sNoDossier_c)) then
          SetField('DSO_PAIEDECALEE', 'X')
      else
          SetField('DSO_PAIEDECALEE', '-');

      if (GetParamSocDossier('SO_PGANALYTIQUE', 'DB'+sNoDossier_c)) then
          SetField('DSO_PAIEANALYTIQUE', 'X')
      else
          SetField('DSO_PAIEANALYTIQUE', '-');

      if (GetParamSocDossier('SO_PGTICKETRESTAU', 'DB'+sNoDossier_c)) then
          SetField('DSO_TICKETREST', 'X')
      else
          SetField('DSO_TICKETREST', '-');
   end
   else
   begin
      SetField('DSO_PAIECAB', '-');
      SetField('DSO_PAIEENT', 'X');
   end;

   SetField('DSO_TXSALPERIODIC', 'MEN');
   SetField('DSO_PAIEENTSYS', '001');
   SetField('DSO_REGPERS', '001');
   SetField('DSO_DECUNEMB', '001');

   // Effectifs
   dtDateFinEx_l := GetField('DSO_DATEFINEX');
   dtExSocial_l := GetField('DSO_DATEEXSOC');
   dEffectif_l := GetField('DSO_EFFECTIF');
   dtDateEffectif_l := GetField('DSO_DATEEFFECTIF');
   dtDateFin_l := iDate1900;

   if InitEffectifs(sNoDossier_c, dtDateFinEx_l, dtExSocial_l,
                    dEffectif_l, dtDateEffectif_l, dtDateFin_l) then
   begin
      SetField('DSO_EFFECTIF', dEffectif_l);
      SetField('DSO_DATEEFFECTIF', dtDateFin_l);
      SetField('DSO_DATEEXSOC', dtDateFin_l);
      SetField('DSO_DATEEFFECTIF', dtDateFin_l);
   end;

   if (BVoirDiode) and (bpaie_c) then
   begin
      UnCodeDiode := DonnerCodeDiode(GetField('DSO_CONVENCOLLEC'));
      SetField('DSO_ORIGDADS', UnCodeDiode);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : ??
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.OnLoadRecord;
var sRequete_l  : String;
    q:TQuery ;
    ctrl : TControl ;
    b:boolean ;
begin
//   LoadRecord:=true ;  //LM20070404
   inherited;
   if Ecran.Name = 'DPCOMSOCIAL' then
   begin
      ModeEdition(DS);
      Exit;
   end;

   SetActiveTabSheet('PGeneral');

   // Taxes et participation
   q:=opensql('select DFI_TAXESALAIRES from DPFISCAL '+ //+LM20070404
              'where DFI_GUIDPER = "' + sGuidPer_c + '"', true ) ;
   if not q.eof then DFI_TAXESALAIRESini := q.FindField('DFI_TAXESALAIRES').AsString
                else DFI_TAXESALAIRESini :='-' ;
   setControlText('DFI_TAXESALAIRES', DFI_TAXESALAIRESini) ;
   ferme(q);

   ctrl :=GetControl('DFI_TAXESALAIRES') ;
   if ctrl<>nil then
   SetControlEnabled('DSO_TXSALPERIODIC', TCheckBox(ctrl).checked);//-LM20070404

   // Traitement de la paie
   SetControlEnabled('DSO_PAIECAB', True); //Modif CAT
   SetControlEnabled('TDSO_REMBULLPAIE', TDBCheckBox(GetControl('DSO_PAIECAB')).checked); //Ajout CAT
   SetControlEnabled('DSO_REMBULLPAIE', TDBCheckBox(GetControl('DSO_PAIECAB')).checked); //Ajout CAT
   SetControlEnabled('DSO_PAIEENT', not bPaie_c);
   SetControlEnabled('DSO_PAIEENTSYS', TDBCheckBox(GetControl('DSO_PAIEENT')).checked);

   // $$$ JP 05/05/06 - FQ 10936
   SetControlEnabled ('TDSO_CONVENCOLLEC', not TDBCheckBox(GetControl('DSO_PAIECAB')).checked);
   SetControlEnabled ('DSO_CONVENCOLLEC',  not TDBCheckBox(GetControl('DSO_PAIECAB')).checked);
   SetControlEnabled ('DSO_ORIGDADS',  not TDBCheckBox(GetControl('DSO_PAIECAB')).checked);

   SetControlEnabled('DSO_PAIEDECALEE', not bPaie_c);
   SetControlEnabled('DSO_PAIEANALYTIQUE', not bPaie_c);
   SetControlEnabled('DSO_PAIEAPPOINTS', not bPaie_c);
   SetControlEnabled('DSO_DATEDEBEX', not bPaie_c);
   SetControlEnabled('DSO_DATEFINEX', not bPaie_c);
   SetControlEnabled('DSO_DATEDERPAIE', not bPaie_c);
   SetControlEnabled('DSO_GESTCONGES', not bPaie_c);
   SetControlEnabled('DSO_MUTSOCAGR', not bPaie_c);
   SetControlEnabled('DSO_INTERMSPEC', not bPaie_c);
   SetControlEnabled('DSO_BTP', not bPaie_c);
   SetControlEnabled('DSO_TICKETREST', not bPaie_c);
   SetControlEnabled('DSO_PAIEGENECRIT', not bPaie_c);
   SetControlEnabled('DSO_GESTIONETS', not bPaie_c);
   SetControlEnabled('DSO_TELEADS', not bPaie_c);
   SetControlEnabled('DSO_PLANPAIEACT', not bPaie_c);

   // Cotisations
   SetControlVisible('TSCOTISATIONS', bPaie_c);
   if not vh_dp.group then //LMO
    SetControlVisible ('TSCOTISATIONSBIS',TDBCheckBox(GetControl('DSO_PAIEENT')).checked); //Ajout CAT

   if bPaie_c then
   begin
        // $$$ JP 12/10/2004 - il faut taper finalement dans DPSOCIALCAISSE car alimentée par PAIE depuis v6 récente
        // $$$ JP 15/12/2004 - fq 10520 - pas de jointures, car pas des codes, mais des données brutes sans lien
        sRequete_l := 'SELECT DSC_NOMCAISSE,DSC_NATUREDUCS,DSC_PERIODICITE FROM DPSOCIALCAISSE WHERE DSC_GUIDPER = "' + sGuidPer_c + '" ORDER BY DSC_NOMCAISSE';
        OnLoadGrille (ThGrid (GetControl ('FLISTECAISSES')), sRequete_l, 'DSC_NOMCAISSE;DSC_NATUREDUCS;DSC_PERIODICITE', 'Caisse;Nature DUCS;Périodicité', '16;6;6', 'G.0O ---;G.0O ---;G.00 ---', TRUE, iMaxCaisses_c);
{        sRequete_l := 'SELECT DSC_NOMCAISSE,CO1.CO_LIBELLE AS CO1_NATURE,CO2.CO_LIBELLE AS CO2_PERIODICITE FROM DPSOCIALCAISSE '
                   +  'LEFT JOIN COMMUN AS CO1 ON DSC_NATUREDUCS=CO1.CO_CODE AND CO1.CO_TYPE="PDO" '
                   +  'LEFT JOIN COMMUN AS CO2 ON DSC_PERIODICITE=CO2.CO_CODE AND CO2.CO_TYPE="PPD" '
                   +  'WHERE DSC_GUIDPER="' + TFFiche (Ecran).FLequel + '" ORDER BY DSC_NOMCAISSE';
        OnLoadGrille (ThGrid (GetControl ('FLISTECAISSES')), sRequete_l, 'DSC_NOMCAISSE;CO1_NATURE;CO2_PERIODICITE', 'Caisse;Nature DUCS;Périodicité', '16;6;6', 'G.0O ---;G.0O ---;G.00 ---', TRUE, iMaxCaisses_c);}
//        OnLoadGrille (ThGrid (GetControl ('FLISTECAISSES')), 'DPSOCIALCAISSE', 'DSC_GUIDPER="' + TFFiche (Ecran).FLequel + '"', FALSE, iMaxCaisses_c);

{        sRequete_l := 'SELECT POG_ETABLISSEMENT,POG_ORGANISME,POG_LIBELLE,POG_NATUREORG,POG_PERIODICITDUCS,'
                    + 'CO1.CO_LIBELLE AS CO1_LIBELLE,CO2.CO_LIBELLE AS CO2_LIBELLE FROM ORGANISMEPAIE'
                    + 'LEFT JOIN COMMUN AS CO1 ON POG_NATUREORG=CO1.CO_CODE AND CO1.CO_TYPE="PDO" '
                    + 'LEFT JOIN COMMUN AS CO2 ON POG_PERIODICITDUCS=CO2.CO_CODE AND CO2.CO_TYPE= "PPD" '+
                    'WHERE POG_ETABLISSEMENT = "' + GetParamSocSecur('SO_ETABLISDEFAUT', '') + '" ' +
                    'ORDER BY POG_LIBELLE';
        OnLoadGrille (ThGrid(GetControl('FLISTECAISSES')), sRequete_l, 'POG_LIBELLE;CO1_LIBELLE;CO2_LIBELLE', 'Nom caisse;Nature DUCS;Périodicité', '16;8;8', 'G.0O ---;G.0O ---;G.00 ---', TRUE, iMaxCaisses_c);}
   end;

  b:=presenceComplexe('ANNUAIRE', ['ANN_GUIDPER', 'ANN_PPPM'], ['=','='], [sGuidPer_c,'PP'],['S','S']);
  SetControlProperty( 'TSSTATUTSOCIAL','TabVisible',b ); //LM20070404

  DSO_CONJOINTAVEC_OnClick(nil);//LM20070404

  if DefTabSheet<>'' then SetActiveTabSheet(DefTabSheet); //LM20070528

  LoadRecord := false ;  //LM20070404
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 23/06/2004
Modifié le ... : 23/06/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.OnChangeField(fChamp_p : TField);
begin
   if Ecran.Name = 'DPCOMSOCIAL' then
      exit;
  inherited;   
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.OnUpdateRecord;
begin
  inherited;

  if DFI_TAXESALAIRESini <> getcontroltext('DFI_TAXESALAIRES') then //LM20070404
    try
      // L'enregistrement fiscal n'existe pas forcément
      ExecuteSql('update DPFISCAL set DFI_TAXESALAIRES = "' + getcontroltext('DFI_TAXESALAIRES') + '" ' +
                'where DFI_GUIDPER = "' + sGuidPer_c + '"' );
    except
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.OnAfterUpdateRecord;
begin
   inherited;
   if IsInside(Ecran) then
      ReloadTomInsideAfterInsert(TFFiche(Ecran), DS, ['DSO_GUIDPER'], [sGuidPer_c]);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 23/06/2004
Modifié le ... :
Description .. : Corrige des valeurs sur l'enregistrement demandé dans
Suite ........ : FLequel du AglLanceFiche, avant même son chargement.
Suite ........ :
Suite ........ : MD : devrait être fait dans OnLoadRecord mais comme la
Suite ........ : fiche ne permet pas de scroller sur d'autres enregistrements,
Suite ........ : on corrige uniqt l'enreg en cours, avant même de le charger
Suite ........ : (cela évite le message "voulez-vous enregistrer les
Suite ........ : modifications" alors qu'on n'a rien modifié.)
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.Initialisation( sGuidPer_p : string );
var
   OBSocial_l : TOB;
   dEffectif_l : Double;
   dtDateFinEx_l, dtDateFin_l, dtExSocial_l, dtDateEffectif_l : TDateTime;
   UnCodeDiode  : String;   
begin
   if sGuidPer_p='' then exit;
   sNoDossier_c := GetNoDossierFromGuidPer(sGuidPer_p);

   // $$$ JP 21/07/06: si pas de base, c'est comme s'il n'y avait pas de paie traitée pour le dossier
   bPaie_c := FALSE;
   if m_bBaseExists = TRUE then
      bPaie_c := ExisteSQL('SELECT DAP_NOMEXEC FROM DOSSAPPLI ' +
                          'WHERE DAP_NOMEXEC = "CPS5.EXE" ' +
                          'AND DAP_NODOSSIER = "' + sNoDossier_c + '"');


   // on ne fait rien sur un enregistrement pas encore créé
   if TFFiche(Ecran).TypeAction in [taCreat..taCreatOne] then exit;
//   if sAction_c = 'ACTION=CREATION' then exit;

   OBSocial_l := TOB.Create('DPSOCIAL', nil, -1);
   OBSocial_l.LoadDetailDBFromSQL('DPSOCIAL', 'SELECT * FROM DPSOCIAL WHERE DSO_GUIDPER = "' + sGuidPer_p + '"');

   if OBSocial_l.Detail.Count = 0 then
   begin
      OBSocial_l.Free;
      exit;
   end;

   // Traitement de la paie
   if bPaie_c then
   begin
      if OBSocial_l.Detail[0].GetValue('DSO_PAIECAB') <> 'X' then
         OBSocial_l.Detail[0].PutValue('DSO_PAIECAB', 'X');

      if OBSocial_l.Detail[0].GetValue('DSO_PAIEENT') <> '-' then
         OBSocial_l.Detail[0].PutValue('DSO_PAIEENT', '-');

      if (GetParamSocDossier('SO_PGTYPECDETICKET', 'DB'+sNoDossier_c) <> '') then
         OBSocial_l.Detail[0].PutValue('DSO_EXISTTICKREST', 'X')
      else
         OBSocial_l.Detail[0].PutValue('DSO_EXISTTICKREST', '-');

      if (GetParamSocDossier('SO_PGDECALAGE', 'DB'+sNoDossier_c)) then
         OBSocial_l.Detail[0].PutValue('DSO_PAIEDECALEE', 'X')
      else
         OBSocial_l.Detail[0].PutValue('DSO_PAIEDECALEE', '-');

      if (GetParamSocDossier('SO_PGANALYTIQUE', 'DB'+sNoDossier_c)) then
         OBSocial_l.Detail[0].PutValue('DSO_PAIEANALYTIQUE', 'X')
      else
         OBSocial_l.Detail[0].PutValue('DSO_PAIEANALYTIQUE', '-');

      if (GetParamSocDossier('SO_PGTICKETRESTAU', 'DB'+sNoDossier_c)) then
         OBSocial_l.Detail[0].PutValue('DSO_TICKETREST', 'X')
      else
         OBSocial_l.Detail[0].PutValue('DSO_TICKETREST', '-');
   end
   else
   begin
      // FQ 11602 - Mise en commentaire
      // car si coché, on laisse coché (on n'a pas forcément la paie Cegid Expert)
      //if OBSocial_l.Detail[0].GetValue('DSO_PAIECAB') <> '-' then
      //   OBSocial_l.Detail[0].PutValue('DSO_PAIECAB', '-');

      // FQ 11602 - Et si c'est coché, il ne faut pas cocher paie entreprise
      if (OBSocial_l.Detail[0].GetValue('DSO_PAIEENT') <> '-') and (OBSocial_l.Detail[0].GetValue('DSO_PAIECAB') <> 'X') then
         OBSocial_l.Detail[0].PutValue('DSO_PAIEENT', 'X');
   end;

   if (OBSocial_l.Detail[0].GetValue('DSO_TXSALPERIODIC') = '') then
      OBSocial_l.Detail[0].PutValue('DSO_TXSALPERIODIC', 'MEN');

   if (OBSocial_l.Detail[0].GetValue('DSO_PAIEENTSYS') = '') then
      OBSocial_l.Detail[0].PutValue('DSO_PAIEENTSYS', '001');

   if (OBSocial_l.Detail[0].GetValue('DSO_REGPERS') = '') then
      OBSocial_l.Detail[0].PutValue('DSO_REGPERS', '001');

   if (OBSocial_l.Detail[0].GetValue('DSO_DECUNEMB') = '') then
      OBSocial_l.Detail[0].PutValue('DSO_DECUNEMB', '001');

   // Effectifs
   dtDateFinEx_l := OBSocial_l.Detail[0].GetValue('DSO_DATEFINEX');
   dtExSocial_l := OBSocial_l.Detail[0].GetValue('DSO_DATEEXSOC');
   dEffectif_l := OBSocial_l.Detail[0].GetValue('DSO_EFFECTIF');
   dtDateEffectif_l := OBSocial_l.Detail[0].GetValue('DSO_DATEEFFECTIF');
   dtDateFin_l := iDate1900;

   if InitEffectifs(sNoDossier_c, dtDateFinEx_l, dtExSocial_l,
                    dEffectif_l, dtDateEffectif_l, dtDateFin_l) then
   begin
      if (OBSocial_l.Detail[0].GetValue('DSO_EFFECTIF') = 0.0) then
      begin
         OBSocial_l.Detail[0].PutValue('DSO_EFFECTIF', dEffectif_l);
         OBSocial_l.Detail[0].PutValue('DSO_DATEEFFECTIF', dtDateFin_l);
      end;
      if (OBSocial_l.Detail[0].GetValue('DSO_DATEEXSOC') = iDate1900) then
      begin
         OBSocial_l.Detail[0].PutValue('DSO_DATEEXSOC', dtDateFin_l);
         OBSocial_l.Detail[0].PutValue('DSO_DATEEFFECTIF', dtDateFin_l);
      end;
   end;

   if (BVoirDiode) and (bpaie_c) then
    begin
     UnCodeDiode := DonnerCodeDiode(OBSocial_l.Detail[0].GetValue('DSO_CONVENCOLLEC'));
     OBSocial_l.Detail[0].PutValue('DSO_ORIGDADS',UnCodeDiode);
    end;

   if OBSocial_l.Detail[0].Modifie then
      OBSocial_l.Detail[0].UpdateDB;

   OBSocial_l.Free;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_DPSOCIAL.InitEffectifs(sNoDossier_p : string;
                                    dtDateFinEx_p, dtExSocial_p : TDateTime;
                                    var dEffectif_p : Double;
                                    var dtDateEffectif_p, dtDateFin_p : TDateTime) : boolean;
var
   OBTabPaie_l : TOB;
begin
   // Effectifs
   result := false;
   if bPaie_c and (dtDateFinEx_p <> iDate1900) then
   begin
      OBTabPaie_l := TOB.Create('DPTABPAIE', nil, -1);
      OBTabpaie_l.LoadDetailDBFromSQL('DPTABPAIE',
                              'SELECT * FROM DPTABPAIE ' +
                              'WHERE DTP_NODOSSIER = "' + sNoDossier_p + '" ' +
                              ' AND DTP_DATEFIN = "' + USDATETIME(dtDateFinEx_p) + '"', true );

      if OBTabpaie_l.Detail.Count > 0 then
      begin
         if (dEffectif_p = 0.0) then
         begin
            dEffectif_p := OBTabpaie_l.Detail[0].GetValue('DTP_EFFECTIF');
            dtDateEffectif_p := Date;
         end;
         if (dtExSocial_p = iDate1900) then
         begin
            dtDateFin_p := OBTabpaie_l.Detail[0].GetValue('DTP_DATEDEB');
            dtDateEffectif_p := Date;
         end;
      end;
      result := (OBTabpaie_l.Detail.Count > 0);
      OBTabpaie_l.Free;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.DFI_TAXESALAIRES_OnClick(Sender : TObject);
begin
  if LoadRecord then exit ;//+LM20070404

  modeEdition(DS);

  // #### MD>LM ça sent la catastrophe : ça veut dire qu'on passe pas par OnNewRecord ... ???
  if (DS.State in [dsInsert, dsEdit]) then
    SetField('DSO_GUIDPER', sGuidPer_c);

   //LM20070404 if (DS.State = dsBrowse) then Exit;
   SetControlEnabled('DSO_TXSALPERIODIC', TCheckBox(GetControl('DFI_TAXESALAIRES')).checked);//-LM20070404
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.DSO_PAIEENT_OnClick(Sender : TObject);
begin
   if LoadRecord then exit ;

   modeEdition(DS);

   SetControlEnabled ('DSO_PAIEENTSYS', TDBCheckBox(GetControl('DSO_PAIEENT')).checked);
   SetControlVisible ('TSCOTISATIONSBIS',TDBCheckBox(GetControl('DSO_PAIEENT')).checked); //Ajout CAT

   // si on a coché l'une, il faut décocher l'autre
   if TDBCheckBox (GetControl ('DSO_PAIEENT')).Checked and TDBCheckBox (GetControl ('DSO_PAIECAB')).Checked then
   begin
     SetField('DSO_PAIECAB', '-');
     TDBCheckBox (GetControl ('DSO_PAIECAB')).Checked := False;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : CATALA DAvid
Créé le ...... : 20/06/2005
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPSOCIAL.DSO_PAIECAB_OnClick(Sender : TObject);
var
   bPaieCab  :boolean;
begin
   if LoadRecord then exit ;

   modeEdition(DS);

   bPaieCab := TDBCheckBox (GetControl ('DSO_PAIECAB')).Checked;

   SetControlEnabled ('TDSO_REMBULLPAIE', bPaieCab); //Ajout CAT
   SetControlEnabled ('DSO_REMBULLPAIE', bPaieCab);

   // $$$ JP 05/05/06 - FQ 10936: autoriser la combo convention si la paie n'est PAS gérée par le cabinet
   SetControlEnabled ('DSO_ORIGDADS', not bPaieCab);
   SetControlEnabled ('DSO_CONVENCOLLEC', not bPaieCab);
   SetControlEnabled ('TDSO_CONVENCOLLEC', not bPaieCab);

   // si on a coché l'une, il faut décocher l'autre
   if (bPaieCab) and (TDBCheckBox (GetControl ('DSO_PAIEENT')).Checked) then
   begin
      SetField('DSO_PAIEENT', '-');
      TDBCheckBox (GetControl ('DSO_PAIEENT')).Checked := False;
   end;
end;

procedure TOM_DPSOCIAL.BWebDiodeClick (Sender:TObject);
var
   strWeb        :string;
   strCodeDiode  :string;
begin
     strCodeDiode := Trim (GetControlText ('DSO_ORIGDADS'));
     if strCodeDiode <> '' then
     begin
          // $$$ JP 09/01/06 - adresse par défaut
          strWeb := Trim (GetParamSocSecur ('SO_DIODEWEBCODE', 'www.diode.fr/index_cegid.php?num_cc=')); //'www.diode.fr\cc_obligation_index.html'));
          if strWeb = '' then
             strWeb := 'www.diode.fr/index_cegid.php?num_cc=';
          LanceWeb (strWeb + strCodeDiode, FALSE)
     end
     else
     begin
          // $$$ JP 09/01/06 - adresse par défaut
          strWeb := Trim (GetParamSocSecur ('SO_DIODEWEBSTD', 'www.diode.fr/index_cegid.php'));
          if strWeb = '' then
             strWeb := 'www.diode.fr/index_cegid.php';
          LanceWeb (strWeb, FALSE);
     end;
end;


//----------------------------
//--- Nom : DonnerCodeDiode
//----------------------------
function TOM_DPSOCIAL.DonnerCodeDiode (UnCodeConvention : String): String;
var Numero, SSql     : String;
    RSql             : TQuery;
begin
 Numero:='';
 SSql:='SELECT PCV_NUMERODIV FROM CONVENTIONCOLL WHERE PCV_CONVENTION="'+UnCodeConvention+'"';
 RSql:=OpenSql (SSql, True);
 if (not RSql.Eof) then
  Numero:=RSql.FindField ('PCV_NUMERODIV').AsString;
 Ferme (RSql);
 Result:=Numero;
end;

//--------------------------------------
//--- Nom : DSO_CONVENCOLLECT_ONCLICK
//--------------------------------------
procedure TOM_DPSOCIAL.DSO_CONVENCOLLEC_OnClick (Sender : TObject);
var UneTHValComboBox : THValComboBox;
begin
 UneTHValComboBox:=THValComboBox (GetControl ('DSO_CONVENCOLLEC'));
 SetField ('DSO_ORIGDADS',DonnerCodeDiode (UneTHValComboBox.Values[UneTHValComboBox.itemIndex]));
end;

procedure TOM_DPSOCIAL.RAZLookup (Sender: TObject);//LMO20060901
var nom : string ;
begin
  if sender=nil then exit ;
  nom := TControl(sender).name ;
  nom := copy(nom,4,length(nom)-3) ;//hypothèse : nom de la gomme = RAZ + nom du champ(=nom zone)
  setControltext(nom,'') ;
  setField(nom,#0);
end ;

procedure TOM_DPSOCIAL.showGED (Sender: TObject);//LMO20060901
var nom : string ;
begin
  if sender=nil then exit ;
  nom := TControl(sender).name ;
  nom := copy(nom,4,length(nom)-3) ;//hypothèse : nom de l'aperçu = GED + nom du champ(=nom zone)
  showGedViewer(getControltext(nom), true) ;

end ;

procedure TOM_DPSOCIAL.OnClose;//LM20070404
begin
  ////LM20070712 détruit par le free de la fiche if EvDecla<>Nil then EvDecla.free ;
  inherited ;
end ;

procedure TOM_DPSOCIAL.DSO_CONJOINTAVEC_OnClick (sender:TObject);
var b : boolean;
begin
  b:=(getControlText('DSO_CONJOINTAVEC')='X');
  setControlVisible('DSO_CONJOINTSTATUT', b) ;
  setControlVisible('TDSO_CONJOINTSTATUT', b) ;
end ;

procedure TOM_DPSOCIAL.BLIENS_OnClick(sender: TObject);
begin
  if GetField('DSO_GUIDPER')<>'' then
    AGLLanceFiche('DP','LIENINTERVENANT','ANL_GUIDPERDOS='+GetField('DSO_GUIDPER')+';ANL_FONCTION=SOC','',GetField('DSO_GUIDPER')+';SOC');
end;

procedure TOM_DPSOCIAL.bValiderOnClick (sender: TObject); //LM20070712
begin
  nextprevcontrol(Ecran, true); //Ne fonctionne pas bien dans le contexte => on le lance en forcant le chgt de zone...
  EvDecla.Exec('bValider', 'OnClick', sender);

end ;


Initialization
  registerclasses ( [ TOM_DPSOCIAL ] ) ;
end.

