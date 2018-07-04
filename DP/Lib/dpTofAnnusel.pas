//----------------------------------------------------------------------------------------------
//--- NOM   : DPTofAnnusel
//--- OBJET : TOF de la fiche YY ANNUAIRE_SEL
//---         Sélection d'une personne sans critères particuliers sur les fonctions
//---         Exemple : rattachement d'une personne à un Tiers (voir ellipsis dans UTomTiers)
//----------------------------------------------------------------------------------------------
unit dpTOFAnnusel;

interface

uses
   UTOF, Controls, AGLInit, HCtrls, HDB, Windows, menus, comctrls,HTB97,
   AnnOutils, HEnt1,
{$IFDEF EAGLCLIENT}
   MaineAGL, eMul,
{$ELSE}
   fe_main, Mul,
{$ENDIF}
   sysutils, Classes, hmsgbox, dpoutils, galOutil, stdctrls
{$IFDEF AFFAIRE}
   ,UtofZonesLibres
{$ENDIF}

{$IFDEF VER150}
   ,Variants
{$ENDIF}

 ;

Type
{$ifdef AFFAIRE}
   TOF_ANNUSEL = Class (TOF_ZONESLIBRES)    //mcd 07/03/06   10937 mise tof_zonelibre au lieu tof
{$else}
   TOF_ANNUSEL = Class (TOF)
{$endif}
      procedure OnArgument(Arguments : String ) ; override ;
      procedure OnLoad; override ;
   private
      TypeFiche   : String;
      TypeRequete : String;
      RComplement : String;
      m_bLibre1Empty, m_bLibre2Empty :boolean;
      bSansDossier_l : boolean;

      procedure FListe_OnFlipSelection(Sender : TObject);
      procedure FListe_OnDblClick(Sender : TObject);
      procedure BVALIDER_OnClick(Sender : TObject);
      procedure OnFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

      procedure BCHERCHE_OnClick(Sender: TObject);
      procedure BCRITINFOSLIBRES_OnClick(Sender: TObject);
      procedure BCRITACTIVITE_OnClick(Sender: TObject);
      procedure BCRITBNC_OnClick(Sender: TObject);
      procedure BCRITASSO_OnClick(Sender: TObject);
      procedure BCRITTIERS_OnClick(Sender: TObject);
      procedure BCRITSOCIAL_OnClick(Sender: TObject);
      procedure BCRITAGRICOLE_OnClick(Sender: TObject);
      procedure SANSGRPCONF_OnClick(Sender: TObject);

      procedure BCULTURES_OnClick (Sender : TObject);
      procedure BCULTURESP_OnClick (Sender : TObject);
      procedure BANIMAUX_OnClick (Sender : TObject);
      procedure BTRANSFORMATION_OnClick (Sender : TObject);
      procedure BAUTRES_OnClick (Sender : TObject);
      function  DonnerListeCode (NomControle : String) : String;
      procedure ConstruitXXWhereGlobal;
      procedure MasqueZone(zonesAmasquer : String ) ; //LM20071126
   end;

function LancerAnnuSel(Param1 : String; Param2 : String; Param3 : String) : String;

implementation

uses
    UDroitsAcces; // dans Commun\Lib

procedure TOF_ANNUSEL.MasqueZone(zonesAmasquer : String ) ; //LM20071126
var st : string;
    //i:integer ;
    //v : variant ;
begin
  if zonesAmasquer ='' then exit ;
  repeat
    st:=readTokenstV(zonesAmasquer) ;
    setcontrolVisible(st, false) ;
    (*à revoir : étendre le getPropertyVal au tkClass
    //recherche si caption associé à la zone
    for i:=0 to TFmul(Ecran).ComponentCount-1 do
      if isPropertyExists(TFmul(Ecran).Components[i], 'FocusControl') then
      begin
        getPropertyVal (TFmul(Ecran).Components[i], 'FocusControl', v) ;
        if uppercase(st) = upperCase(v) then
        begin
          setcontrolVisible(TFmul(Ecran).Components[i].Name, false) ;
          break ;
        end ;
      end ;
    *)
  until st='' ;
end ;

//------------------------------------------------------------------------
//--- Nom   : OnArgument
//--- Objet : Récupération des arguments
//---         Argument ACTION (TIERS,DOUBLONS,DOUBLONSIRET,LIENPERSONNE)
//------------------------------------------------------------------------
procedure TOF_ANNUSEL.OnArgument(Arguments : String ) ;
var TypeAction, sSansDossier_l, st : string;
    ii : integer;
begin
 inherited;
 MasqueZone (RecupArgument('MASQUE', Arguments));//LM20071126

 //--- Récupére les paramètres
 if (UpperCase(Copy(Arguments,1,6))='ACTION') then
  TypeAction:=ReadTokenSt(Arguments);  // Inutilisé
 // TypeFiche peut recevoir : DOUBLONS / DOUBLONSIRET / TIERS / LIENSPERSONNE
 TypeFiche:=ReadTokenSt (Arguments);
 TypeRequete:=ReadTokenSt (Arguments);
 // #### MD 24/03/06 : pas rencontré de cas où on se servait de ce dernier paramètre !!??
 sSansDossier_l := ReadTokenSt (Arguments);
 bSansDossier_l := (sSansDossier_l = 'X');
 RComplement:='';

 //--- Construction des requetes en fonction des types de fiche appelante
  if (TypeFiche='DOUBLONS') or (TypeFiche='DOUBLONSIRET') then
   begin
    //--- On recherche les homonymes sur la concaténation de (ANN_NOM1.Text+ANN_NOMPER.Text)
    //--- ou sur le numéro de SIRET (ANN_SIREN+ANN_CLESIRET)
    SetControlText('XX_WHERE',TypeRequete);
    TFMul(Ecran).Caption := 'Homonymes pour l''intervenant ou l''organisme saisi : ' ;
    SetControlVisible('BAGRANDIR',false);
    SetControlVisible('BREDUIRE',false);
    SetControlVisible('BRECHERCHER',false);
    SetControlVisible('BEXPORT',false);
    SetControlVisible('BSELECTALL',false);
    SetControlVisible('BDELETE',false);
    SetControlVisible('BINSERT',false);
    SetControlVisible('BZOMM',false);
    THDBGrid(GetControl('FLISTE')).OnFlipSelection := FListe_OnFlipSelection;
   end
  else
   begin
    SetControlVisible('PAnnexe',False);

    // Droits de création,
    GereDroitsConceptsAnnuaire(Self, Ecran);

    //--- Libellés et valeurs des zones libres (si aucun dans un onglet, on empêche de montrer l'onglet)
    AfficheLibTablesLibres(Self);
    m_bLibre1Empty := IsTabEmpty (TTabSheet (GetControl ('PLIBRE1')));
    m_bLibre2Empty := IsTabEmpty (TTabSheet (GetControl ('PLIBRE2')));
    if (m_bLibre1Empty = TRUE) AND (m_bLibre1Empty = TRUE) then
     TMenuItem (GetControl('CRIT_INFOSLIBRES')).Visible := FALSE;

    // Initialisation
    InitialiserComboGroupeConf(THMultiValCombobox(GetControl('GROUPECONF')));

    // Filtrage/positionnement
    ConstruitXXWhereGlobal;

    //--- On recherche des personnes qui ne sont pas déjà affectées à des tiers...
    if (TypeFiche='TIERS') then
     begin
      // MD => fait dans ConstruitXXWhereGlobal
      // RComplement:=' AND (ANN_TIERS="")';
      // ChXXWhere:=ChXXWhere+RComplement;
      //mcd 22/03/2006 docn dans ce cas, il ne faut pas les valeurs par défaut sur champ tierscompl
      if ((Ecran.FindComponent('YTC_TABLELIBRETIERS1')) <> nil) then
      begin
        for ii := 1 to 9 do
          SetControlText('YTC_TABLELIBRETIERS' + IntToStr(ii),'');
        SetControlText('YTC_TABLELIBRETIERSA','');
        if TPopupMenu(GetControl('POPUPCRIT'))<> nil  then  //on cache acces option client sur bouon ajout critère
          for ii:=0 to TPopupMenu(GetControl('POPUPCRIT')).items.count-1 do
             begin
             st:=uppercase(TPopupMenu(GetControl('POPUPCRIT')).items[ii].name);
             if st='CRIT_TIERS' then
                 begin
                 TPopupMenu(GetControl('POPUPCRIT')).items[ii].visible:=False;
                 break;
                 end;
             end;
      end;
     end
     ;
     //--- Fin typefich Tiers ---
    //--- On recherche des personnes faisant l'objet d'un lien avec le guidperdos en param
    //else if (TypeFiche='LIENSPERSONNE') then
          //begin
           // MD => fait dans ConstruitXXWhereGlobal
           // RComplement:=' AND (ANN_GUIDPER IN (SELECT ANL_GUIDPER FROM ANNULIEN WHERE ANL_GUIDPERDOS="'+TypeRequete+'"))';
           // ChXXWhere:=ChXXWhere+RComplement;
          //end;

    TMenuItem(GetControl('CRIT_INFOSLIBRES')).OnClick := BCRITINFOSLIBRES_OnClick;
    TMenuItem(GetControl('CRIT_ACTIVITE')).OnClick    := BCRITACTIVITE_OnClick;
    TMenuItem(GetControl('CRIT_BNC')).OnClick         := BCRITBNC_OnClick;
    TMenuItem(GetControl('CRIT_ASSO')).OnClick        := BCRITASSO_OnClick;
    TMenuItem(GetControl('CRIT_TIERS')).OnClick       := BCRITTIERS_OnClick;
    TMenuItem(GetControl('CRIT_SOCIAL')).OnClick      := BCRITSOCIAL_OnClick;
    TMenuItem(GetControl('CRIT_AGRICOLE')).Onclick    := BCRITAGRICOLE_OnClick;
    TCheckbox(GetControl('SANSGRPCONF')).OnClick      := SANSGRPCONF_OnClick;

    TCheckBox(GetControl('BCULTURES')).OnClick        := BCULTURES_OnClick;
    TCheckBox(GetControl('BCULTURESP')).OnClick       := BCULTURESP_OnClick;
    TCheckBox(GetControl('BANIMAUX')).OnClick         := BANIMAUX_OnClick;
    TCheckBox(GetControl('BTRANSFORMATION')).OnClick  := BTRANSFORMATION_OnClick;
    TCheckBox(GetControl('BAUTRES')).OnClick          := BAUTRES_OnClick;

    TToolBarButton97(GetControl('BCHERCHE')).OnClick := BCHERCHE_OnClick;
   end;

  //--- Affectation des événements
  THDBGrid(GetControl('FLISTE')).OnDblClick:=FListe_OnDblClick;
  TToolBarButton97(GetControl('BVALIDER')).OnClick:=BValider_OnClick;
  Ecran.OnKeyDown:=OnFormKeyDown;

  //--- Initialisation de la multiValCombobox LISTEAPPLICATION
  if THMultiValComboBox (GetControl ('ListeApplication'))<>Nil then
   InitialiserComboApplication (THMultiValComboBox (GetControl ('LISTEAPPLICATION')));
end;

//------------------------------------
//--- Nom   : OnLoad
//--- Objet : Affichage des onglets
//------------------------------------
procedure TOF_ANNUSEL.OnLoad;
begin
 inherited;

 SetControlVisible('PCritere',(TypeFiche<>'DOUBLONS') and (TypeFiche<>'DOUBLONSIRET'));
 SetControlVisible('PRegFiscDir',(TypeFiche<>'DOUBLONS') and (TypeFiche<>'DOUBLONSIRET'));
 SetControlVisible('PRegFisInDir',(TypeFiche<>'DOUBLONS') and (TypeFiche<>'DOUBLONSIRET'));
 SetControlVisible('PAvance',(TypeFiche<>'DOUBLONS') and (TypeFiche<>'DOUBLONSIRET'));
 SetControlVisible('PComplement',(TypeFiche<>'DOUBLONS') and (TypeFiche<>'DOUBLONSIRET'));
 SetControlVisible('BCritComplement',(TypeFiche<>'DOUBLONS') and (TypeFiche<>'DOUBLONSIRET'));


 if (TypeFiche='DOUBLONSIRET') then
  begin
   SetControlCaption('LBLMESSAGE1', 'Le numéro de SIRET saisi présente un ou plusieurs homonymes.');
   SetControlCaption('LBLMESSAGE4', 'Sinon <Annuler> et corriger le numéro de siret.');
  end;
end;

//------------------------------------------------------------------------
//--- Nom   : FListe_OnDblClick
//--- Objet : Sélection du ligne
//------------------------------------------------------------------------
procedure TOF_ANNUSEL.FListe_OnDblClick(Sender : TObject);
var
    SGuidPer : string;
begin
 if (TypeFiche='DOUBLONS') or (TypeFiche='DOUBLONSIRET') then
  begin
   {$IFDEF EAGLCLIENT}
    THGrid(GetControl('FLISTE')).ClearSelected;
    THGrid(GetControl('FLISTE')).FlipSelection(THGrid(GetControl('FLISTE')).Row);
   {$ELSE}
    THDBGrid(GetControl('FLISTE')).ClearSelected;
    THDBGrid(GetControl('FLISTE')).FlipSelection;
   {$ENDIF}
  end
 else
  begin
   SGuidPer:=VarToStr(GetField('ANN_GUIDPER'));
   if SGuidPer<>'' then AGLLanceFiche('YY','ANNUAIRE',SGuidPer,SGuidPer,'');
    AglRefreshDB([LongInt(Ecran),'FListe'], 2);
  end;
end;

//-------------------------------------------
//--- Nom   : FListe_OnFlipSelection
//--- Objet : Vérification de la sélection
//-------------------------------------------
procedure TOF_ANNUSEL.FListe_OnFlipSelection(Sender : TObject);
begin
 if THDBGrid(GetControl('FLISTE')).nbSelected > 1 then
  begin
   PGIInfo( 'Vous ne pouvez sélectionner qu''une seule ligne.', Ecran.Caption );
   THDBGrid(GetControl('FLISTE')).ClearSelected;
  end;
end;

//-------------------------------------------
//--- Nom   : BValider_OnClick
//--- Objet : Validation de la fiche
//-------------------------------------------
procedure TOF_ANNUSEL.BVALIDER_OnClick(Sender : TObject);
begin
 if ((TypeFiche='DOUBLONS') or (TypeFiche='DOUBLONSIRET')) and (THDBGrid(GetControl('FLISTE')).nbSelected<1) then
  begin
   PGIInfo( 'Vous n''avez sélectionné aucune ligne.#13#10'+' L''intervenant ou organisme ne sera pas remplacé.',Ecran.Caption);
   Ecran.Close;
  end
 else
  begin
     // $$$ JP 23/06/06: pour vega, il faut respecter le format de retour: "nomcle=valcle" de l'enreg sélectionné: ici, ANN_GUIDPER=...
{$IFDEF VEGA}
     TFMul (Ecran).Retour := 'ANN_GUIDPER=' + VarToStr (GetField ('ANN_GUIDPER'));
{$ELSE}
     TFMul(Ecran).Retour := VarToStr(GetField('ANN_GUIDPER'));
{$ENDIF}
     Ecran.Close;
  end;
end;

//----------------------------------
//--- Nom   : OnFormKeyDown
//--- Objet : Gestion des touches
//----------------------------------
procedure TOF_ANNUSEL.OnFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
 BtnIns : TToolbarButton97;
begin
 Case Key of
  //--- VK_F10 = lancer traitement
  VK_F10 : begin
             Key:=0;
             // Une icône BValider a été rajoutée par dessus le bouton BOuvrir !!!
             if GetControlVisible('BValider') then
               TToolBarButton97(GetControl('BValider')).Click;
           end;
  //--- CTRL+N = Nouveau
  Ord('N') : begin
               BtnIns := TToolbarButton97(GetControl('BINSERT'));
               if (ssCtrl in Shift) and (BtnIns <> nil) and (BtnIns.Visible) then
                 BtnIns.Click;
             end;

  else
    TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
  end;
end;

//------------------------------------------------------------------------------
//--- Nom   : BCHERCHE_ONCLICK
//--- Objet :
//------------------------------------------------------------------------------
procedure TOF_ANNUSEL.BCHERCHE_OnClick(Sender: TObject);
begin
 NextPrevControl(Ecran, True, True); // FQ 11630 mais inutile pour l'instant :
 // la liste des groupes de travail affichée dans les mul YY ANNUAIRE_SELxxxx
 // est une liste complète (non réduite à ceux auxquel l'utilisateur a droit)

 ConstruitXXWhereGlobal;

 TFMul(Ecran).BChercheClick(Sender);
end;

function TOF_ANNUSEL.DonnerListeCode (NomControle : String) : String;
var Indice     : Integer;
    SListeCode : String;
begin
 Result:=THMultiValCombobox (GetControl (NomControle)).text;
 if (Result='<<Tous>>') then
  begin
   SListeCode:='';
   for Indice:=0 to THMultiValCombobox (GetControl (NomControle)).items.count-1 do
    SListeCode:=SListeCode+THMultiValCombobox (GetControl (NomControle)).Values [Indice]+';';
   Result:=SListeCode;
  end;
end;


//------------------------------------------------------------------------------
//--- Nom   : BCRITINFOLIBRE_OnClick
//--- Objet :
//------------------------------------------------------------------------------
procedure TOF_ANNUSEL.BCRITINFOSLIBRES_OnClick(Sender: TObject);
begin
 if m_bLibre1Empty = FALSE then
  SetControlVisible ('PLIBRE1', Not (TTabSheet (GetControl ('PLIBRE1')).TabVisible));
 if m_bLibre2Empty = FALSE then
  SetControlVisible ('PLIBRE2', Not (TTabSheet (GetControl ('PLIBRE2')).TabVisible));
 TMenuItem (GetControl ('CRIT_INFOSLIBRES')).Checked := (TTabSheet (GetControl ('PLIBRE1')).TabVisible or TTabSheet (GetControl ('PLIBRE2')).TabVisible);
end;

//------------------------------------------------------------------------------
//--- Nom   : BCRITACTIVITE_OnClick
//--- Objet :
//------------------------------------------------------------------------------
procedure TOF_ANNUSEL.BCRITACTIVITE_OnClick(Sender: TObject);
begin
 SetControlVisible ('PACTIVITE', Not (TTabSheet (GetControl ('PACTIVITE')).TabVisible));
 TMenuItem (GetControl ('CRIT_ACTIVITE')).Checked := TTabSheet (GetControl ('PACTIVITE')).TabVisible;
end;

//------------------------------------------------------------------------------
//--- Nom   : BCRITBNC_OnClick
//--- Objet :
//------------------------------------------------------------------------------
procedure TOF_ANNUSEL.BCRITBNC_OnClick(Sender: TObject);
begin
 SetControlVisible ('PBNC', Not (TTabSheet (GetControl ('PBNC')).TabVisible));
 TMenuItem (GetControl ('CRIT_BNC')).Checked := TTabSheet (GetControl ('PBNC')).TabVisible;
end;

//------------------------------------------------------------------------------
//--- Nom   : BCRITASSO_OnClick
//--- Objet :
//------------------------------------------------------------------------------
procedure TOF_ANNUSEL.BCRITASSO_OnClick(Sender: TObject);
begin
 SetControlVisible ('PASSOCIATION', Not (TTabSheet (GetControl ('PASSOCIATION')).TabVisible));
 TMenuItem (GetControl ('CRIT_ASSO')).Checked := TTabSheet (GetControl ('PASSOCIATION')).TabVisible;
end;

//------------------------------------------------------------------------------
//--- Nom   : BCRITTIERS_OnClick
//--- Objet :
//------------------------------------------------------------------------------
procedure TOF_ANNUSEL.BCRITTIERS_OnClick(Sender: TObject);
begin
 SetControlVisible ('PSTATTIERS',  Not (TTabSheet (GetControl ('PSTATTIERS')).TabVisible));
 SetControlVisible ('PCOMPLTIERS', Not (TTabSheet (GetControl ('PCOMPLTIERS')).TabVisible) );
 TMenuItem(GetControl ('CRIT_TIERS')).Checked := TTabSheet(GetControl ('PSTATTIERS')).TabVisible;
end;

//------------------------------------------------------------------------------
//--- Nom   : BCRITSOCIAL_OnClick
//--- Objet :
//------------------------------------------------------------------------------
procedure TOF_ANNUSEL.BCRITSOCIAL_OnClick(Sender: TObject);
begin
 SetControlVisible ('PDPSOCIAL', Not (TTabSheet (GetControl ('PDPSOCIAL')).TabVisible));
 TMenuItem (GetControl ('CRIT_SOCIAL')).Checked := TTabSheet (GetControl('PDPSOCIAL')).TabVisible;
end;

procedure TOF_ANNUSEL.BCRITAGRICOLE_OnClick(Sender: TObject);
begin
     SetControlVisible ('PDPAGRICOLE', Not (TTabSheet (GetControl ('PDPAGRICOLE')).TabVisible));
     TMenuItem (GetControl ('CRIT_AGRICOLE')).Checked := TTabSheet (GetControl('PDPAGRICOLE')).TabVisible;
end;

procedure TOF_ANNUSEL.SANSGRPCONF_OnClick(Sender: TObject);
begin
  GereCheckboxSansGrpConf(TCheckbox(GetControl('SANSGRPCONF')), THMultiValCombobox(GetControl('GROUPECONF')) );
end;

procedure TOF_ANNUSEL.BCULTURES_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BCULTURES')).Checked;
 THLabel (GetControl ('LACTIVITE1')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBCULTURE')).Enabled:=Etat;
end;

procedure TOF_ANNUSEL.BCULTURESP_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BCULTURESP')).Checked;
 THLabel (GetControl ('LACTIVITE2')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBCULTUREP')).Enabled:=Etat;
end;

procedure TOF_ANNUSEL.BANIMAUX_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BANIMAUX')).Checked;
 THLabel (GetControl ('LACTIVITE3')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBANIMAUX')).Enabled:=Etat;
end;

procedure TOF_ANNUSEL.BTRANSFORMATION_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BTRANSFORMATION')).Checked;
 THLabel (GetControl ('LACTIVITE4')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBTRANSFORMATION')).Enabled:=Etat;
end;

procedure TOF_ANNUSEL.BAUTRES_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BAUTRES')).Checked;
 THLabel (GetControl ('LACTIVITE5')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBAUTRES')).Enabled:=Etat;
end;

//------------------------------------------------------------------------------
//--- Nom   : LancerAnnuSel
//--- Objet :
//------------------------------------------------------------------------------
function LancerAnnuSel(Param1 : String; Param2 : String; Param3 : String) : String;
begin
{$IFDEF BUREAU}
  if ExExJaiLeDroitConceptBureau(ccListesAnnSimplifiees) then
   Result:=AGLLanceFiche('YY','ANNUAIRE_SELLITE', Param1, Param2, Param3)
  else
   Result:=AGLLanceFiche('YY','ANNUAIRE_SEL', Param1, Param2, Param3);
{$ELSE}
   Result:=AGLLanceFiche('YY','ANNUAIRE_SELLITE', Param1, Param2, Param3);
{$ENDIF}
end;


procedure TOF_ANNUSEL.ConstruitXXWhereGlobal;
var
    ChXXWhere : String;
begin
  ChXXWhere:='';

  // On recherche des personnes qui ne sont pas déjà affectées à des tiers...
  if (TypeFiche='TIERS') then
    RComplement:=' AND (ANN_TIERS="")'
  // On recherche des personnes faisant l'objet d'un lien avec le guidperdos en param
  else if (TypeFiche='LIENSPERSONNE') then
    RComplement:=' AND (ANN_GUIDPER IN (SELECT ANL_GUIDPER FROM ANNULIEN WHERE ANL_GUIDPERDOS="'+TypeRequete+'"))'
  // Cas de requête SQL directement passé en paramètre
  else if (TypeRequete<>'') then
    RComplement:=' AND ('+TypeRequete+')';

  if (Not (THMultiValCombobox(GetControl('GROUPECONF')).Tous) or (TcheckBox (GetControl ('SANSGRPCONF')).Checked)) then
  begin
    ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),TcheckBox (GetControl ('SANSGRPCONF')).Checked);
    ChXXWhere:=ChXXWhere+GererCritereApplication (THMultiValComboBox (GetControl ('LISTEAPPLICATION')));
    ChXXWhere:=ChXXWhere+GererCritereDivers (False, False, bSansDossier_l);
  end;

  // FQ 11490
  if ChXXWhere<>'' then ChXXWhere := '(' + ChXXWhere + ') AND ';
  ChXXWhere := '(NOT EXISTS (SELECT 1 FROM DOSSIER WHERE DOS_GUIDPER=ANN_GUIDPER AND DOS_NODOSSIER="000STD") )';

  // Critères agricoles
  if (TcheckBox (GetControl ('BCULTURES')).Checked) then
    ChXXWhere:=ChXXWhere+GererCritereAgricole (DonnerListeCode ('CBCULTURE'));
  if (TcheckBox (GetControl ('BCULTURESP')).Checked) then
    ChXXWhere:=ChXXWhere+GererCritereAgricole (DonnerListeCode ('CBCULTUREP'));
  if (TcheckBox (GetControl ('BANIMAUX')).Checked) then
    ChXXWhere:=ChXXWhere+GererCritereAgricole (DonnerListeCode ('CBANIMAUX'));
  if (TcheckBox (GetControl ('BTRANSFORMATION')).Checked) then
    ChXXWhere:=ChXXWhere+GererCritereAgricole (DonnerListeCode ('CBTRANSFORMATION'));
  if (TcheckBox (GetControl ('BAUTRES')).Checked) then
    ChXXWhere:=ChXXWhere+GererCritereAgricole (DonnerListeCode ('CBAUTRES'));

  if (TCheckBox (GetControl ('CHKINTVDOSSIER')).Checked) then
  ChXXWhere:=ChXXWhere+' AND ANN_FAMPER=""';

  if (TCheckBox (GetControl ('CHKINTVDISPO')).Checked) then
  ChXXWhere:=ChXXWhere+' AND EXISTS (SELECT DOS_GUIDPER FROM DOSSIER WHERE DOS_GUIDPER=ANN_GUIDPER)';

  if RComplement<>'' then ChXXWhere:=ChXXWhere+RComplement;

  // $$$ JP 04/09/06 - plus autorisé en D7/Unicode ... TEdit(GetControl('XX_WHERE')).Text:=ChXXWhere;
  SetControlText ('XX_WHERE', ChXXWhere);
end;

Initialization
registerclasses([TOF_ANNUSEL]) ;
end.
