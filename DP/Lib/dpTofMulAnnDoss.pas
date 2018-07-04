{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYMULANNDOSS
Mots clefs ... : TOF;MULANNDOSS
*****************************************************************}
Unit dpTofMulAnnDoss ;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
uses UTOF, UTOFMailingStd, CBPPath;
//////////////////////////////////////////////////////////////////
Type
  TOF_MULANNDOSS = Class (TOF_MAILINGSTD)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    sWhere_Dossiers, sWhere_FamPer, sWhere_ClientProspect : string;
    m_bLibre1Empty, m_bLibre2Empty   :boolean;

    procedure BCRITINFOSLIBRES_OnClick(Sender: TObject);
    procedure BCRITACTIVITE_OnClick(Sender: TObject);
    procedure BCRITBNC_OnClick(Sender: TObject);
    procedure BCRITASSO_OnClick(Sender: TObject);
    procedure BCRITTIERS_OnClick(Sender: TObject);
    procedure BCRITSOCIAL_OnClick(Sender: TObject);
    procedure BCRITNETEXPERT_OnClick(Sender: TObject);
    procedure BCRITINTERLOC_OnClick (Sender: TObject);
    procedure BCRITAGRICOLE_OnClick(Sender: TObject);    

//    procedure BNOUVRECH_OnClick(Sender: TObject);
    procedure CHKINTVDISPO_OnClick(Sender: TObject);
    procedure CHKINTVDOSSIER_OnClick(Sender: TObject);
    procedure CHKCLIENTPROSPECT_OnClick(Sender: TObject);
    procedure SANSGRPCONF_OnClick(Sender: TObject);

    procedure BCULTURES_OnClick (Sender : TObject);
    procedure BCULTURESP_OnClick (Sender : TObject);
    procedure BANIMAUX_OnClick (Sender : TObject);
    procedure BTRANSFORMATION_OnClick (Sender : TObject);
    procedure BAUTRES_OnClick (Sender : TObject);
    function  DonnerListeCode (NomControle : String) : String;

    procedure GereIntvDispo;

    procedure AfterShow;   // $$$ JP 03/12/2004 - pour le menu "nouvelle recherche" des filtres, trop tôt dans OnArgument
    procedure OnNouvelleRecherche (Sender:TObject);
    procedure OnSupprimeFiltre (Sender:TObject);
  end ;

///////////////////////// IMPLEMENTATION /////////////////////////
Implementation
//////////////////////////////////////////////////////////////////
Uses
   HEnt1, HDB, StdCtrls, Classes, ComCtrls, HCtrls, Menus, HTB97,
{$IFDEF EAGLCLIENT}
   eMul,
{$ELSE}
   Mul,
{$ENDIF}
   ParamSoc, galOutil, dpOutils, entdp;

procedure TOF_MULANNDOSS.AfterShow;
begin
     TFMul (Ecran).ListeFiltre.OnItemNouveau   := OnNouvelleRecherche;
     TFMul (Ecran).ListeFiltre.OnItemSupprimer := OnSupprimeFiltre;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnArgument (S : String ) ;
var repmail: String;
begin
  Inherited ;
//  TMenuItem(GetControl('BNOUVRECH')).OnClick := BNOUVRECH_OnClick; // $$$ JP 03/12/04 - remplacé par OnNouvelleRecherche

  // Bouton d'affichage des onglets de critères complémentaires
  TMenuItem (GetControl ('CRIT_INFOSLIBRES')).OnClick := BCRITINFOSLIBRES_OnClick;
  TMenuItem (GetControl ('CRIT_ACTIVITE')).OnClick    := BCRITACTIVITE_OnClick;
  TMenuItem (GetControl ('CRIT_BNC')).OnClick         := BCRITBNC_OnClick;
  TMenuItem (GetControl ('CRIT_ASSO')).OnClick        := BCRITASSO_OnClick;
  TMenuItem (GetControl ('CRIT_TIERS')).OnClick       := BCRITTIERS_OnClick;
  TMenuItem (GetControl ('CRIT_SOCIAL')).OnClick      := BCRITSOCIAL_OnClick;
  TMenuItem(GetControl('CRIT_AGRICOLE')).Onclick     := BCRITAGRICOLE_OnClick;
  if (VH_DP.EwsActif) or (GetParamSocSecur ('SO_NECWASURL', '') <> '') then
      TMenuItem (GetControl ('CRIT_NETEXPERT')).OnClick := BCRITNETEXPERT_OnClick
  else
      TMenuItem (GetControl ('CRIT_NETEXPERT')).Visible := FALSE;
  TMenuItem (GetControl ('CRIT_INTERLOC')).OnClick := BCRITINTERLOC_OnClick;

  TCheckBox(GetControl('CHKINTVDISPO')).OnClick := CHKINTVDISPO_OnClick;
  TCheckBox(GetControl('CHKINTVDOSSIER')).OnClick := CHKINTVDOSSIER_OnClick;

  if TCheckBox(GetControl('CHKCLIENTPROSPECT'))<> nil then
  TCheckBox(GetControl('CHKCLIENTPROSPECT')).OnClick := CHKCLIENTPROSPECT_OnClick;

  TCheckBox(GetControl('SANSGRPCONF')).OnClick := SANSGRPCONF_OnClick;

  TCheckBox(GetControl('BCULTURES')).OnClick        := BCULTURES_OnClick;
  TCheckBox(GetControl('BCULTURESP')).OnClick       := BCULTURESP_OnClick;
  TCheckBox(GetControl('BANIMAUX')).OnClick         := BANIMAUX_OnClick;
  TCheckBox(GetControl('BTRANSFORMATION')).OnClick  := BTRANSFORMATION_OnClick;
  TCheckBox(GetControl('BAUTRES')).OnClick          := BAUTRES_OnClick;

{$IFDEF EAGLCLIENT}
  THDBGrid(GetControl('FLISTE')).MultiSelect := false;
{$ELSE}
  THDBGrid(GetControl('FLISTE')).MultiSelection := false;
{$ENDIF}

//  ZonesLibresOnglets ('PLibre');
  //ZonesLibresInit    ('ANNUPARAM', 'YYANNUPARAM');

  //Initialisation du mailing
  repmail := GetParamSocSecur('SO_MDREPDOCUMENTS', TCbpPath.GetCegidDistriStd+'\BUREAU');
  RepertoireMailing (repmail, repmail, repmail);

  // Libellés et valeurs des zones libres (si aucun dans un onglet, on empêche de montrer l'onglet)
  AfficheLibTablesLibres (Self);
  m_bLibre1Empty := IsTabEmpty (TTabSheet (GetControl ('PLIBRE1')));
  m_bLibre2Empty := IsTabEmpty (TTabSheet (GetControl ('PLIBRE2')));
  if (m_bLibre1Empty = TRUE) AND (m_bLibre1Empty = TRUE) then
     TMenuItem (GetControl('CRIT_INFOSLIBRES')).Visible := FALSE;

  // filtrage/positionnement => rajout crit. pour voir l'annuaire même sans dossiers en face
  InitialiserComboGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')));

   // Affectation des procédures
   Case TFMul(Ecran).AutoSearch of
      asChange : ; // voir onclick de la case
      asExit   :   // voir proc. InitAutoSearch du mul
      begin
         TCheckBox(GetControl('CHKINTVDISPO')).onExit := TFMul(Ecran).SearchTimerTimer ;
         TCheckBox(GetControl('CHKINTVDOSSIER')).onExit := TFMul(Ecran).SearchTimerTimer ;

         if TCheckBox(GetControl('CHKCLIENTPROSPECT'))<> nil then
          TCheckBox(GetControl('CHKCLIENTPROSPECT')).onExit := TFMul(Ecran).SearchTimerTimer ;
      end;
      asTimer  : ; // voir onclick de la case
   end;

  // $$$ JP 03/12/2004 - Màj callback du menu filtre "nouvelle recherche" trop tôt dans OnArgument, il faut le faire dans AfterShow
  TFmul (Ecran).OnAfterFormShow := AfterShow;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnNew ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnLoad ;
var
   sXXWhere_l : String;
   ChXXWhere  : String;
begin
   inherited;
   if sWhere_FamPer <> '' then
      sXXWhere_l := sWhere_FamPer;

   // Si uniquement les dossiers clients
   if sWhere_Dossiers <> '' then
   begin
      if sXXWhere_l <> '' then
         sXXWhere_l := sXXWhere_l + ' AND ';
      sXXWhere_l := sXXWhere_l + sWhere_Dossiers;
   end;

     // Si client prospect
  if sWhere_ClientProspect<>'' then
    begin
    if sXXWhere_l<>'' then sXXWhere_l := sXXWhere_l + ' AND ';
    sXXWhere_l := sXXWhere_l + sWhere_ClientProspect;
    end;

  // Si click SansGroupConf
  if (Not (THMultiValCombobox(GetControl('GROUPECONF')).Tous) or (TcheckBox (GetControl ('SANSGRPCONF')).Checked)) then
   begin
    ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),TcheckBox (GetControl ('SANSGRPCONF')).Checked);
    ChXXWhere:=ChXXWhere+GererCritereDivers ();
    if (sXXWhere_l<>'') then sXXWhere_l:=ChXXWhere+' AND '+sXXWhere_l else sXXWhere_l:=ChXXWhere;
   end;

  // Critères agricoles
  if (TcheckBox (GetControl ('BCULTURES')).Checked) then
   sXXWhere_l:=sXXWhere_l+GererCritereAgricole (DonnerListeCode ('CBCULTURE'));
  if (TcheckBox (GetControl ('BCULTURESP')).Checked) then
   sXXWhere_l:=sXXWhere_l+GererCritereAgricole (DonnerListeCode ('CBCULTUREP'));
  if (TcheckBox (GetControl ('BANIMAUX')).Checked) then
   sXXWhere_l:=sXXWhere_l+GererCritereAgricole (DonnerListeCode ('CBANIMAUX'));
  if (TcheckBox (GetControl ('BTRANSFORMATION')).Checked) then
   sXXWhere_l:=sXXWhere_l+GererCritereAgricole (DonnerListeCode ('CBTRANSFORMATION'));
  if (TcheckBox (GetControl ('BAUTRES')).Checked) then
   sXXWhere_l:=sXXWhere_l+GererCritereAgricole (DonnerListeCode ('CBAUTRES'));

  // FQ 11490
  if sXXWhere_l<>'' then sXXWhere_l := '(' + sXXWhere_l + ') AND ';
  sXXWhere_l:=sXXWhere_l+'(NOT EXISTS (SELECT 1 FROM DOSSIER WHERE DOS_GUIDPER=ANN_GUIDPER AND DOS_NODOSSIER="000STD") )';

  SetControlText( 'XX_WHERE', sXXWhere_l );

  // $$$ JP - il faut cacher par défaut les onglets "plibre_"
//  m_bLibre1 := GetControlVisible ('PLIBRE1');
//  m_bLibre2 := GetControlVisible ('PLIBRE2');
//  SetControlProperty ('PLIBRE1', 'TabVisible', FALSE);
//  SetControlProperty ('PLIBRE2', 'TabVisible', FALSE);
//  if (m_bLibre1 = FALSE) and (m_bLibre2 = FALSE) then
  //   TMenuItem (GetControl('CRIT_INFOSLIBRES')).Visible := FALSE;
end ;

function TOF_MULANNDOSS.DonnerListeCode (NomControle : String) : String;
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

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MULANNDOSS.BCRITINFOSLIBRES_OnClick(Sender: TObject);
begin
     if m_bLibre1Empty = FALSE then
        SetControlVisible ('PLIBRE1', Not (TTabSheet (GetControl ('PLIBRE1')).TabVisible));
     if m_bLibre2Empty = FALSE then
        SetControlVisible ('PLIBRE2', Not (TTabSheet (GetControl ('PLIBRE2')).TabVisible));
     TMenuItem (GetControl ('CRIT_INFOSLIBRES')).Checked := (TTabSheet (GetControl ('PLIBRE1')).TabVisible or TTabSheet (GetControl ('PLIBRE2')).TabVisible);
end;

procedure TOF_MULANNDOSS.BCRITACTIVITE_OnClick(Sender: TObject);
begin
     SetControlVisible('PACTIVITE', Not (TTabSheet(GetControl('PACTIVITE')).TabVisible) );
     TMenuItem(GetControl('CRIT_ACTIVITE')).Checked := TTabSheet(GetControl('PACTIVITE')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCRITBNC_OnClick(Sender: TObject);
begin
     SetControlVisible('PBNC', Not (TTabSheet(GetControl('PBNC')).TabVisible) );
     TMenuItem (GetControl ('CRIT_BNC')).Checked := TTabSheet(GetControl('PBNC')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCRITASSO_OnClick(Sender: TObject);
begin
     SetControlVisible('PASSOCIATION', Not (TTabSheet(GetControl('PASSOCIATION')).TabVisible) );
     TMenuItem (GetControl ('CRIT_ASSO')).Checked := TTabSheet(GetControl('PASSOCIATION')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCRITTIERS_OnClick(Sender: TObject);
begin
     SetControlVisible ('PSTATTIERS', Not (TTabSheet(GetControl('PSTATTIERS')).TabVisible) );
     SetControlVisible ('PCOMPLTIERS', Not (TTabSheet(GetControl('PCOMPLTIERS')).TabVisible) );
     TMenuItem (GetControl ('CRIT_TIERS')).Checked := TTabSheet(GetControl('PSTATTIERS')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCRITSOCIAL_OnClick(Sender: TObject);
begin
     SetControlVisible('PDPSOCIAL', Not (TTabSheet(GetControl('PDPSOCIAL')).TabVisible) );
     TMenuItem (GetControl ('CRIT_SOCIAL')).Checked := TTabSheet(GetControl('PDPSOCIAL')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCRITNETEXPERT_OnClick(Sender: TObject);
begin
     SetControlVisible('PNETEXPERT', Not (TTabSheet(GetControl('PNETEXPERT')).TabVisible) );
     TMenuItem (GetControl ('CRIT_NETEXPERT')).Checked := TTabSheet(GetControl('PNETEXPERT')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCRITINTERLOC_OnClick(Sender: TObject);
begin
     SetControlVisible('PINTERLOC', Not (TTabSheet(GetControl('PINTERLOC')).TabVisible) );
     TMenuItem (GetControl ('CRIT_INTERLOC')).Checked := TTabSheet(GetControl('PINTERLOC')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCRITAGRICOLE_OnClick(Sender: TObject);
begin
     SetControlVisible ('PDPAGRICOLE', Not (TTabSheet (GetControl ('PDPAGRICOLE')).TabVisible));
     TMenuItem (GetControl ('CRIT_AGRICOLE')).Checked := TTabSheet (GetControl('PDPAGRICOLE')).TabVisible;
end;

procedure TOF_MULANNDOSS.BCULTURES_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BCULTURES')).Checked;
 THLabel (GetControl ('LACTIVITE1')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBCULTURE')).Enabled:=Etat;
end;

procedure TOF_MULANNDOSS.BCULTURESP_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BCULTURESP')).Checked;
 THLabel (GetControl ('LACTIVITE2')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBCULTUREP')).Enabled:=Etat;
end;

procedure TOF_MULANNDOSS.BANIMAUX_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BANIMAUX')).Checked;
 THLabel (GetControl ('LACTIVITE3')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBANIMAUX')).Enabled:=Etat;
end;

procedure TOF_MULANNDOSS.BTRANSFORMATION_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BTRANSFORMATION')).Checked;
 THLabel (GetControl ('LACTIVITE4')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBTRANSFORMATION')).Enabled:=Etat;
end;

procedure TOF_MULANNDOSS.BAUTRES_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BAUTRES')).Checked;
 THLabel (GetControl ('LACTIVITE5')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBAUTRES')).Enabled:=Etat;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{procedure TOF_MULANNDOSS.BNOUVRECH_OnClick(Sender: TObject);
begin
  TFMul(Ecran).BNouvRechClick(Sender);
  // la recherche le remet à 0 ...
  SetControlText('DFI_JOURDECLA', '-1');
end;}

procedure TOF_MULANNDOSS.OnNouvelleRecherche (Sender:TObject);
begin
     // Traitement AGL
     TFMul(Ecran).BNouvRechClick (Sender);

     // On veut -1 et pas 0 par défaut dans jour déclaration TVA
     SetControlText ('DFI_JOURDECLA', '-1');
end;

procedure TOF_MULANNDOSS.OnSupprimeFiltre (Sender:TObject);
begin
     // Traitement AGL
     TFMul(Ecran).BDelFiltreClick (Sender);

     // On veut -1 et pas 0 par défaut dans jour déclaration TVA
     SetControlText ('DFI_JOURDECLA', '-1');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 20/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.CHKINTVDISPO_OnClick(Sender: TObject);
begin
   if TCheckBox(GetControl('CHKINTVDISPO')).checked then
   begin
      // obligé de purger Plus pour éviter que le mul fasse un OR de toutes les valeurs
      // possibles affichées dans la combo...
      SetControlText('ANN_TYPEPER', '');
      // les intervenants disponibles = ceux sans famille
      sWhere_FamPer := 'ANN_FAMPER=""';
   end
   else
   begin
      // #### que devient TYPEPER qui devrait dépendre de ce choix...
      sWhere_FamPer := '';
   end;
   GereIntvDispo ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 20/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.CHKINTVDOSSIER_OnClick(Sender: TObject);
begin
  if TCheckBox(GetControl('CHKINTVDOSSIER')).checked then
    // Uniquement les clients
    sWhere_Dossiers := 'EXISTS (SELECT DOS_GUIDPER FROM DOSSIER WHERE DOS_GUIDPER=ANN_GUIDPER)'
  else
    // Tous
    sWhere_Dossiers := '';
end;

procedure TOF_MULANNDOSS.CHKCLIENTPROSPECT_OnClick(Sender: TObject);
begin
  if TCheckBox(GetControl('CHKCLIENTPROSPECT')).checked then
    // Uniquement les clients
    sWhere_ClientProspect := 'ANN_TIERS<>'' AND ANN_TIERS IS NOT NULL'
  else
    // Tous
    sWhere_ClientProspect := '';
end;

procedure TOF_MULANNDOSS.SANSGRPCONF_OnClick(Sender: TObject);
begin
  GereCheckboxSansGrpConf(TCheckbox(GetControl('SANSGRPCONF')), THMultiValCombobox(GetControl('GROUPECONF')) );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 20/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULANNDOSS.GereIntvDispo;
var actif: Boolean;
begin
  actif := Not (TCheckBox(GetControl('CHKINTVDISPO')).checked);
  SetControlEnabled('ANN_TYPEPER', actif);
end;


//////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_MULANNDOSS ] ) ;
end.

