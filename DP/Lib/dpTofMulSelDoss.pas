{***********UNITE*************************************************
Auteur  ...... : MD
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYMULSELDOSS
Mots clefs ... : TOF;MULSELDOSS
*****************************************************************}
Unit dpTofMulSelDoss ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     MaineAGL,
     MenuOLX,
{$ELSE}
     Mul,
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     FE_Main,
     MenuOLG,
{$ENDIF}
     PGIAppli,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF, EntGC,
     UGedDP,
     HTB97, AGLInit, HDB, Menus, Windows, HStatus, UIUtil,
     ExtCtrls, UTob, HQry,uWAIni,uWAIniBase,uhttp,uhttpCS,
     utofzoneslibres,
     GalSelectGroupeConf,
     GalTraitement,
     WinnerUtilGin,
     winnerInfosInstall, //PGR 08/2007 Infos Install Winner
     galSystem;

Type
  TOF_MULSELDOSS = class (TOF_ZONESLIBRES) // TOF_AFTRADUCCHAMPLIBRE) //Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    { Déclarations privées }
    ChXXWhere   : String;
    Qry         : THQuery;
    Lst         : THDBGrid;
{$IFNDEF GALDISPATCH}
    DossierDest : String;
    TimerDest   : TTimer;
{$ENDIF}
    InfosWinner : TWinnerInfo;
    m_bLibre1Empty, m_bLibre2Empty : boolean;
    mnMarquePourTransport, mnSynthDossier : TMenuItem;
    mnBlocageFonctionnel                  : TMenuItem; // $$$ JP 25/04/06
    bDroitActiverDossierAsp : Boolean;
    procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BLANCEDOSSIER_OnClick(Sender: TObject);
    procedure DOS_NODOSSIER_OnExit(Sender: TObject);
    procedure FLISTE_OnDblClick(Sender: TObject);
    procedure BSUPPRIMER_OnClick(Sender: TObject);
    procedure BINSERT_OnClick(Sender: TObject);
    procedure BFICHEDOSSIER_OnClick(Sender: TObject);
    procedure BOUVRIR_OnClick(Sender: TObject);
    procedure BCHERCHE_OnClick(Sender: TObject);
    procedure ShowTabInSql (strSql:string; strTab1Name:string; strTab2Name:string; strMenuName:string); // $$$ JP 26/02/07: gestion de 2 onglets liés

    procedure BCRITINFOSLIBRES_OnClick(Sender: TObject);
    procedure BCRITACTIVITE_OnClick(Sender: TObject);
    procedure BCRITBNC_OnClick(Sender: TObject);
    procedure BCRITASSO_OnClick(Sender: TObject);
    procedure BCRITTIERS_OnClick(Sender: TObject);
    procedure BCRITSOCIAL_OnClick(Sender: TObject);
    procedure BCRITAGRICOLE_OnClick(Sender: TObject);
    procedure BCRITNETEXPERT_OnClick(Sender: TObject);
    procedure SANSGRPCONF_OnClick(Sender: TObject);

    procedure BCULTURES_OnClick (Sender : TObject);
    procedure BCULTURESP_OnClick (Sender : TObject);
    procedure BANIMAUX_OnClick (Sender : TObject);
    procedure BTRANSFORMATION_OnClick (Sender : TObject);
    procedure BAUTRES_OnClick (Sender : TObject);
    function DonnerListeCode (NomControle : String) : String;

    procedure mnSynthDossierClick(Sender: TObject);
    procedure mnConfDossierClick(Sender: TObject);
    procedure mnConfProfilDossierClick(Sender: TObject);
    procedure mnPwdDossierClick(Sender: TObject);
    procedure mnMarquePourTransportClick(Sender: TObject);
    procedure mnBlocageFonctionnelClick (Sender: TObject); // $$$ JP 25/04/06
    procedure mnLiensannuaireClick(Sender: TObject);
    procedure mnLiensdelapersonneClick(Sender: TObject);
    procedure mnEvenementsClick(Sender: TObject);
    procedure PPM_OnPopup(Sender: TObject);
    procedure LocaliseNoDossier;
    function  GetLibDossier : String;
    function  ConfirmeAvecListeAppli: String;
    procedure ResetTimer(Sender: TObject);
{$IFNDEF GALDISPATCH}
    procedure TimerDest_OnTimer(Sender : TObject);
{$ELSE}
    procedure LanceFicheDossier (strDossier:string);
{$ENDIF}
    procedure MajMenuRepertoireTel(sGuidPer: String);
    procedure mnActiverClick(Sender: TObject);
    procedure mnTabComptaClick(Sender : TObject);
    procedure mnTabPaieClick(Sender : TObject);
    procedure mnTabActiviteClick (Sender:TObject);
    procedure mnEchangesEtNetExpertClick(Sender: TObject); //MP 07-2004
    procedure MnEnCoursClientClick (Sender : TObject); //CAT 01-2005
    procedure MnNouveauMessageClick (Sender : TObject); //CAT 03-2005
    procedure MnNouvActionClick (Sender : TObject); //mcd 06/07/2005
    procedure MnNouvelleActiviteClick (Sender : TObject); //CAT 03-2005
    procedure MnApplicationsClick (Sender : TObject); //CAT 04-2005

    procedure MnPriseEnCompteEDIClick (Sender : TObject); //CAT 12-2007
    procedure MnEnvoiEDIClick (Sender : TObject); //CAT 12-2007
    procedure MnJournalEDIClick (Sender : TObject); //CAT 12-2007
    procedure MnJournalSyntheseEDIClick (Sender : TObject); //CAT 19-2007

    function  GetNoDossDansListeAvecPwd: String;
    function  PasSurUnDossier : Boolean;
    procedure ActiveContenuPopup(bActif : Boolean);

    procedure AfterShow;   // $$$ JP 02/12/2004 - pour le menu "nouvelle recherche" des filtres, trop tôt dans OnArgument
    procedure OnNouvelleRecherche (Sender:TObject);
    procedure OnSupprimeFiltre (Sender:TObject);

    procedure VireOngletNetExpert;

{$IFDEF BUREAU}
    procedure MContactClick (Sender:TObject);
{$ENDIF}
    //PGR 08/2007 Infos Install Winner
    procedure BINFOSWINNER_OnClick (Sender:TObject);
  end ;


///////////// IMPLEMENTATION //////////////
Implementation

uses
     galMenuDisp, UtilTranspDossier,
{$IFDEF EAGLCLIENT}
{$ELSE}
     galReparationSql,
{$ENDIF}
{$IFDEF EWS}
     UtileWS,
{$ENDIF}
     dpTableauBordGeneral, dpTableauBordLibrairie, dpTableauBord,

     {$IFDEF VER150}
     Variants,
     {$ENDIF}

     UDossierSelect,
     galAssistDossier, galOutil, dpOutils, AnnOutils,
     EntDP, ChangePwdDossier, UtilMulTraitmt, ParamSoc, PwdDossier,
     galParamS1ASP, dpTableauBordDetailSolde,
     YNewMessage, Mailol, DPOutilsAgenda, galSynthese, GRPTRAVAIL_TOF,
     TDI_LibFonction, TDI_Receptionner, TDI_JOURNAL_TOF, TDI_ENVOYER_TOF;


procedure TOF_MULSELDOSS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULSELDOSS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULSELDOSS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULSELDOSS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULSELDOSS.OnClose ;
begin
  Inherited ;
  InfosWinner.Free;
end ;

procedure TOF_MULSELDOSS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULSELDOSS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULSELDOSS.AfterShow;
begin
     TFMul (Ecran).ListeFiltre.OnItemNouveau   := OnNouvelleRecherche;
     TFMul (Ecran).ListeFiltre.OnItemSupprimer := OnSupprimeFiltre;
end;

procedure TOF_MULSELDOSS.OnArgument (S : String ) ;
var
   DFI_JOURDECLA   : THDBSpinEdit;
   Mn,MnRacine     : TMenuItem;
begin
  Inherited ;
  // composants
  Qry := TFMul(Ecran).Q;
  Lst := TFMul(Ecran).FListe;
  InfosWinner := TWinnerInfo.Create;
  Ecran.Caption := 'Liste des dossiers';
  UpdateCaption(Ecran);

  // Timer pour accès dossier client, sinon access vio0 si on passe
  // par BOuvrir au lieu de double-clic sur la liste
{$IFNDEF GALDISPATCH}
  TimerDest := TTimer.Create(Ecran);
  TimerDest.Interval := 300;
  TimerDest.Enabled := False;
  TimerDest.OnTimer := TimerDest_OnTimer;
{$ENDIF}

  // évènements
  Ecran.OnKeyDown := Form_OnKeyDown;
  TToolBarButton97(GetControl('bLanceDossier')).OnClick := BLANCEDOSSIER_OnClick;
  THCritMaskEdit(GetControl('DOS_NODOSSIER')).OnExit := DOS_NODOSSIER_OnExit;
  Lst.OnDblClick := FLISTE_OnDblClick;
  if Not JaiLeDroitTag(26051) then SetControlVisible('BSUPPRIMER', False)
  else TToolBarButton97(GetControl('BSUPPRIMER')).OnClick := BSUPPRIMER_OnClick;
  if Not JaiLeDroitTag(26053) then SetControlVisible('BINSERT', False)
  else TToolBarButton97(GetControl('BINSERT')).OnClick := BINSERT_OnClick;
  TToolBarButton97(GetControl('BFICHEDOSSIER')).OnClick := BFICHEDOSSIER_OnClick;
  TToolBarButton97(GetControl('BOUVRIR')).OnClick       := BOUVRIR_OnClick;
  TToolBarButton97(GetControl('BCHERCHE')).OnClick      := BCHERCHE_OnClick;
  TCheckBox(GetControl('SANSGRPCONF')).OnClick          := SANSGRPCONF_OnClick;
  //PGR 08/2007 Infos Install Winner
  if GetControl('BINFOSWINNER')<>nil then
    TToolBarButton97(GetControl('BINFOSWINNER')).OnClick      := BINFOSWINNER_OnClick;
  if GetParamSocSecur ('SO_MDLIENWINNER', False) then
    SetControlVisible('BINFOSWINNER', True);

  TCheckBox(GetControl('BCULTURES')).OnClick        := BCULTURES_OnClick;
  TCheckBox(GetControl('BCULTURESP')).OnClick       := BCULTURESP_OnClick;
  TCheckBox(GetControl('BANIMAUX')).OnClick         := BANIMAUX_OnClick;
  TCheckBox(GetControl('BTRANSFORMATION')).OnClick  := BTRANSFORMATION_OnClick;
  TCheckBox(GetControl('BAUTRES')).OnClick          := BAUTRES_OnClick;

  // Bouton d'affichage des onglets de critères complémentaires
  TMenuItem(GetControl('CRIT_INFOSLIBRES')).OnClick := BCRITINFOSLIBRES_OnClick;
  TMenuItem(GetControl('CRIT_ACTIVITE')).OnClick    := BCRITACTIVITE_OnClick;
  TMenuItem(GetControl('CRIT_BNC')).OnClick         := BCRITBNC_OnClick;
  TMenuItem(GetControl('CRIT_ASSO')).OnClick        := BCRITASSO_OnClick;
  TMenuItem(GetControl('CRIT_TIERS')).OnClick       := BCRITTIERS_OnClick;
  TMenuItem(GetControl('CRIT_SOCIAL')).OnClick      := BCRITSOCIAL_OnClick;
  TMenuItem(GetControl('CRIT_AGRICOLE')).Onclick    := BCRITAGRICOLE_OnClick;
  VireOngletNetExpert;

  // mcd 06/07/2005
  if not VH_GC.GRCSeria then
   TMenuItem(GetControl('mnNouvAction')).Visible:=False
  else
   if GetControl('mnNouvAction')<>nil then
    TMenuItem(GetControl('mnNouvAction')).OnClick := mnNouvActionClick;

  // $$$ JP 24/10/05 - il faut rendre invisible les menuitem, et non les "disabled", car re-"enabled" à chaque OnPopup
  if (not VH_DP.SeriaMessagerie) then
  begin
       TMenuItem(GetControl('mnNouveauMessage')).Visible := FALSE;
       TMenuItem(GetControl('mnNouveauRdv')).Visible     := FALSE;
  end
  else
  begin
       if GetControl('mnNouveauMessage')<>nil then
          TMenuItem(GetControl('mnNouveauMessage')).OnClick := mnNouveauMessageClick;

       if GetControl('mnNouveauRdv')<>nil then
       begin
            MnRacine:=TMenuItem (GetControl ('MnNouveauRdv'));
            Mn:=TMenuItem.create (MnRacine);
            Mn.Name:='MnActiviteInterne';
            Mn.Caption:='Activité interne';
            Mn.OnClick:=mnNouvelleActiviteClick;
            MnRacine.Add(Mn);
            Mn:=TMenuItem.create (MnRacine);
            Mn.Name:='MnActiviteExterne';
            Mn.Caption:='Activité externe';
            Mn.OnClick:=mnNouvelleActiviteClick;
            MnRacine.Add(Mn);
            Mn:=TMenuItem.create (MnRacine);
            Mn.Name:='MnAbsence';
            Mn.Caption:='Absence';
            Mn.OnClick:=mnNouvelleActiviteClick;
            MnRacine.Add(Mn);
       end;
  end;

  TMenuItem(GetControl('mnEchangesEtNetExpert')).OnClick := mnEchangesEtNetExpertClick;
  TMenuItem(GetControl('mnEnCoursClient')).OnClick := mnEnCoursClientClick; // CAT 01-2005

{ if GetControl('mnPriseEnCompteEDI')<>Nil then
    TMenuItem(GetControl('mnPriseEnCompteEDI')).OnClick := mnPriseEnCompteEDIClick; // CAT 12-2007
  if GetControl('mnEnvoiEDI')<>Nil then
    TMenuItem(GetControl('mnEnvoiEDI')).OnClick := mnEnvoiEDIClick; // CAT 12-2007 }
  if GetControl('mnJournalEDI')<>Nil then
    TMenuItem(GetControl('mnJournalEDI')).OnClick := mnJournalEDIClick; // CAT 12-2007

  if GetControl ('MnJournalSyntheseEDI')<>Nil then
    TMenuItem(GetControl('mnJournalSyntheseEDI')).OnClick := mnJournalSyntheseEDIClick; // CAT 12-2007

  // $$$ JP 02/12/2004 - trop tôt, fait dans AfterShow
//  TMenuItem(GetControl('BNOUVRECH')).OnClick := BNOUVRECH_OnClick;
  mnSynthDossier := TMenuItem(GetControl('mnSynthDossier'));
  mnSynthDossier.OnClick := mnSynthDossierClick;
  TMenuItem(GetControl('mnConfDossier')).OnClick := mnConfDossierClick;
  if GetControl('mnConfProfilDossier') <> nil then
  begin
    TMenuItem(GetControl('mnConfProfilDossier')).OnClick := mnConfProfilDossierClick;
    TMenuItem(GetControl('mnConfProfilDossier')).Visible := V_pgi.Superviseur or JaiLeDroitAdminDossier ;
  end;


  // $$$ JP 25/04/06 - invisible si pas le droit changer mdp
  if GetControl('mnPwdDossier')<>Nil then
     with TMenuItem (GetControl ('mnPwdDossier')) do
          if JaiLeDroitConceptBureau (26055) = TRUE then
              OnClick := mnPwdDossierClick
          else
              Visible := FALSE;

  mnMarquePourTransport := TMenuItem(GetControl('mnMarquePourTransport'));
  mnMarquePourTransport.OnClick := mnMarquePourTransportClick;

  // $$$ JP 25/04/06
  mnBlocageFonctionnel := TMenuItem (GetControl ('mnBlocageFonctionnel'));
  mnBlocageFonctionnel.OnClick := mnBlocageFonctionnelClick;

  TMenuItem(GetControl('mnLiensannuaire')).OnClick := mnLiensannuaireClick;
  TMenuItem(GetControl('mnLiensdelapersonne')).OnClick := mnLiensdelapersonneClick;
  if VH_DP.SeriaMessagerie then
    TMenuItem(GetControl('mnEvenements')).OnClick := mnEvenementsClick
  else
    TMenuItem(GetControl('mnEvenements')).Visible := False;
  TMenuItem(GetControl('mnActiver')).OnClick := mnActiverClick;
  TMenuItem(GetControl('mnTabCompta')).OnClick := mnTabComptaClick;
  TMenuItem(GetControl('mnTabPaie')).OnClick := mnTabPaieClick;
  TMenuItem(GetControl('mnTabActivite')).OnClick := mnTabActiviteClick;

  TPopupMenu(GetControl('PPM')).OnPopup := PPM_OnPopup;
  // Fonctionnalité non prête
  TMenuItem(GetControl('mnLettreMission')).Visible := False;
  // pas de DP hors PCL
  if (V_PGI.ModePCL<>'1')
  or (Not VH_DP.SeriaDP) then
    begin
    TMenuItem(GetControl('mnLiensannuaire')).Visible := False;
    TMenuItem(GetControl('mnLiensdelapersonne')).Visible := False;
    end;
  // pas de conf si interdiction du concept "Confidentialité d'un dossier"
  if (Not JaiLeDroitTag(26052)) then
    TMenuItem(GetControl('mnConfDossier')).Visible := False;

  DFI_JOURDECLA := THDBSpinEdit(GetControl('DFI_JOURDECLA'));
  // idem InitAutoSearch dans Mul, car ne traite pas le THDBSpinEdit
  Case TFMul(Ecran).AutoSearch of
   asChange : DFI_JOURDECLA.OnChange := TFMul(Ecran).SearchTimerTimer ;
   asExit   : DFI_JOURDECLA.OnExit := TFMul(Ecran).SearchTimerTimer ;
   asTimer  : DFI_JOURDECLA.OnChange := ResetTimer ;
  end;

  // Libellés et valeurs des zones libres (si aucun dans un onglet, on empêche de montrer l'onglet)
  AfficheLibTablesLibres(Self);
  m_bLibre1Empty := IsTabEmpty (TTabSheet (GetControl ('PLIBRE1')));
  m_bLibre2Empty := IsTabEmpty (TTabSheet (GetControl ('PLIBRE2')));
  if (m_bLibre1Empty = TRUE) AND (m_bLibre1Empty = TRUE) then
     TMenuItem (GetControl('CRIT_INFOSLIBRES')).Visible := FALSE;

  // filtrage/positionnement
  InitialiserComboGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')));
  ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),TcheckBox (GetControl ('SANSGRPCONF')).Checked);
  ChXXWhere:=ChXXWhere+GererCritereDivers ();

  //PGR 10/07/2006 Ajout critère base dossier = CEGIDEWEENVEXP
  //pgr 14/12/2007 Afficher systématiquement le dossier 000000
  if GetParamSocSecur ('SO_MDLIENWINNER', False) and (GetEnvVar('CEGIDEWEENVEXP') <> '') then
  begin
   if FileExists(ExtractFilePath(Application.ExeName)+ 'Trait000.nop') then
     ChXXWhere:=ChXXWhere+ ' AND (DOS_WINSTALL="' + GetEnvVar('CEGIDEWEENVEXP') + '" OR DOS_WINSTALL="" )'
   else ChXXWhere:=ChXXWhere+ ' AND ((DOS_WINSTALL="' + GetEnvVar('CEGIDEWEENVEXP') + '") OR (DOS_WINSTALL="") OR (DOS_NODOSSIER="000000"))';
  end;

  // $$$ JP 04/09/06: pour D7/Unicode: ne plus faire ce genre de cast... TEdit(GetControl('XX_WHERE')).Text:=ChXXWhere;
  SetControlText ('XX_WHERE', chXXWhere);

  // $$$ JP 02/12/2004 - Màj callback du menu filtre "nouvelle recherche" trop tôt dans OnArgument, il faut le faire dans AfterShow
  TFmul (Ecran).OnAfterFormShow := AfterShow;

  //--- Initialisation de la multiValCombobox LISTEAPPLICATION
  if THMultiValComboBox (GetControl ('ListeApplication'))<>Nil then
   begin
    case TFMul(Ecran).AutoSearch of
     asChange : THMultiValComboBox (GetControl ('ListeApplication')).OnChange := TFMul(Ecran).SearchTimerTimer ;
     asExit   : THMultiValComboBox (GetControl ('ListeApplication')).OnExit := TFMul(Ecran).SearchTimerTimer ;
     asTimer  : THMultiValComboBox (GetControl ('ListeApplication')).OnChange := ResetTimer ;
    end;
    InitialiserComboApplication (THMultiValComboBox (GetControl ('LISTEAPPLICATION')));
   end;

  if vh_dp.Group then//LM20071108
  begin
    setcontrolvisible('mnLettreMission', false) ;
    setcontrolvisible('mnEncoursClient', false) ;
    setcontrolvisible('mnSep4', false) ;
    setcontrolvisible('mnTabCompta', false );
    setcontrolvisible('mnTabPaie', false) ;
    setcontrolvisible('mnTabActivite', false) ;
    setcontrolvisible('mnBlocageFonctionnel', false) ;
    setcontrolvisible('mnSep6', false) ;
    setcontrolvisible('mnEchangesEtNetExpert', false) ;
    setcontrolvisible('mnSep8', false) ;
  end ;

{$IFDEF EWS}
  bDroitActiverDossierAsp := JaiLeDroitConceptBureau(ccActiverDossierAsp);
{$ELSE}
  bDroitActiverDossierAsp := False;
{$ENDIF}
end;

procedure TOF_MULSELDOSS.Form_OnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of

  // F11 = affichage du popup
  VK_F11 :  begin
            TPopupMenu(GetControl('PPM')).Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
            // évite de traiter le F11 par défaut de Pgi (=vk_choixmul=double clic ds grid)
            Key := 0;
            end;

  // F5 = sélection de dossier = double-click sur la grille
  // (dans pgi, c'est le F11 qui fait ça...)
  VK_F5 : if (Shift = []) and (TFMul(Ecran).FListe.Focused) then
            begin
            Key := 0;
            Lst.OnDblClick(Lst);
            exit; // pas d'inherited, car le bSelectAll.Down := FListe.AllSelected ;
                  // du mul fait access vio car le dbl click ferme la fiche
            end;

  // Suppression : Ctrl + Suppr
  VK_DELETE : if (Shift = [ssCtrl]) and (TFMul(Ecran).FListe.Focused) then
            BSUPPRIMER_OnClick(Nil);

{ automatiques
  // Appliquer les critères
  //  VK_F9: TToolbarButton97(GetControl('BCHERCHE')).Click;

    VK_F10: BValider.Click;

    VK_F12: if FListe.Focused then
        PageControl.SetFocus
      else
        FListe.SetFocus;

    // Ctrl + A
    65: if (FListe.Focused) and (Shift = [ssCtrl]) then
        BSelectAll.Click;

    // Ctrl + F
    ...

  else
  end; }
  end;

  TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
end;

procedure TOF_MULSELDOSS.BLANCEDOSSIER_OnClick(Sender: TObject);
// Effectue l'ouverture d'un dossier, MEME S'IL n'est pas dans
// le jeu d'enregistrement du multi-critère
var
  nodoss, nodossrech: String;
begin
  nodossrech := '';
  nodoss := GetControlText('DOS_NODOSSIER');

  if not IsNoDossierOk (nodoss) then
    begin
    PgiInfo('Ce numéro de dossier n''est pas correct (uniquement majuscules ou chiffres).', TitreHalley);
    exit;
    end;
  if ExisteSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_NODOSSIER="'+nodoss+'"') then
    nodossrech := nodoss;
  if nodossrech='' then
    begin
    PgiInfo ('Ce numéro de dossier n''existe pas.', TitreHalley);
    exit;
    end;
  if JaiLeDroitDossier(nodossrech) then
    begin
    // C'est le timer qui fera l'accès au dossier client,
    // sinon access vio en changeant de fiche inside sur bouvrir
{$IFDEF GALDISPATCH}
    LanceFicheDossier (nodossrech);
{$ELSE}
    DossierDest := nodossrech;
    TimerDest.Enabled := True;
{$ENDIF}
    end
  else
    PGIInfo('Vous ne possédez pas les autorisations suffisantes pour accéder à ce dossier.',TitreHalley);
end;

procedure TOF_MULSELDOSS.DOS_NODOSSIER_OnExit(Sender: TObject);
var nodoss: String;
begin
  nodoss := GetControlText('DOS_NODOSSIER');
  if nodoss = '' then exit;
  if not IsNoDossierOk (nodoss) then
    begin
    PgiInfo('Le numéro de dossier n''est pas correct (uniquement majuscules ou chiffres).', TitreHalley);
    SetFocusControl('DOS_NODOSSIER');
    exit;
    end;
end;

procedure TOF_MULSELDOSS.FLISTE_OnDblClick(Sender: TObject);
// sélectionne le dossier cliqué, et passe au dossier client
begin
     if PasSurUnDossier then exit;

{$IFDEF GALDISPATCH}
   LanceFicheDossier (GetField ('DOS_NODOSSIER'));
{$ELSE}
    // C'est le timer qui fera l'accès au dossier client
    DossierDest := GetField('DOS_NODOSSIER'); // (*)
    // (*) => ne pas mettre Qry.FindField('DOS_NODOSSIER').AsString
    // car, en eagl, cela renvoit toujours sur la 1ère ligne du mul
    // à moins de faire Qry.TQ.Seek(...)
    TimerDest.Enabled := True;
{$ENDIF}
end;

procedure TOF_MULSELDOSS.BSUPPRIMER_OnClick(Sender: TObject);
var
  sNoDossier, sGuidPer, txt: String;
  marq, usr, login : String;
  bSansBase : Boolean;
  i: Integer;
  procedure SupprimeDossierDansListe;
    begin
    sGuidPer := Qry.FindField('DOS_GUIDPER').AsString;
    sNoDossier := Qry.FindField('DOS_NODOSSIER').AsString;
    // état du dossier
    marq := EtatMarqueDossier(sNoDossier, usr, login, bSansBase);
    if (marq = 'PAR') then
      PGIInfo('Le dossier '+GetLibDossier+' a été emporté par '+login+'. Vous ne pouvez pas le supprimer.', TitreHalley)
    else
      begin
      txt := ConfirmeAvecListeAppli;
      // suppression DP + base physique
      if (PgiAsk(txt, TitreHalley)=mrYes) then SupprimeInfoDp(sGuidPer);
      // si on supprime le dossier en cours, on le déselectionne !
      if (VH_Doss<>Nil) and (sNoDossier=VH_Doss.NoDossier) then
        VH_Doss.NoDossier := '';
      end;
    end;

BEGIN
  inherited;
  if Not JaiLeDroitTag(26051) then
    begin
    PGIInfo('Vous n''avez pas accès à cette fonctionnalité.', TitreHalley);
    exit;
    end;
  if (Lst.nbSelected=0) and (Not Lst.AllSelected) then
    begin
    PGIInfo('Aucun dossier sélectionné.', TitreHalley);
    exit;
    end;

  // liste des dossiers à supprimer
  if Lst.AllSelected then
    BEGIN
{$IFDEF EAGLCLIENT}
    if not TFMul(Ecran).Fetchlestous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
      begin
      Qry.First;
      while Not Qry.EOF do
        begin
        SupprimeDossierDansListe;
        Qry.Next;
        end;
      end;
    END
  else
    BEGIN
    InitMove(Lst.NbSelected,'');
    for i:=0 to Lst.NbSelected-1 do
      begin
      MoveCur(False);
      Lst.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Qry.TQ.Seek(Lst.Row - 1) ;
{$ENDIF}
      SupprimeDossierDansListe;
      end;
    FiniMove;
    END;
  // déselectionne
  FinTraitmtMul(TFMul(Ecran));
  // actualisation
  BCHERCHE_OnClick(Sender);
END;

procedure TOF_MULSELDOSS.BFICHEDOSSIER_OnClick(Sender: TObject);
var nodp: String;
begin
  inherited;

  // consultation du dossier en cours
  // $$$ JP 10/08/05 - non, il faut voir celui en sélection visuel, pas celui actif
  if VH_Doss=Nil then exit;
  nodp := GetField ('DOS_GUIDPER'); //IntToStr(VH_Doss.Guidper);
  AGLLanceFiche('YY','ANNUAIRE',nodp,nodp,'ACTION=CONSULTATION');
end;

// $$$ JP 24/08/05 - on repère les critères qui sont dans la reqûete active du mul
procedure TOF_MULSELDOSS.ShowTabInSql (strSql:string; strTab1Name:string; strTab2Name:string; strMenuName:string); // $$$ JP 26/02/07: gestion de 2 onglets liés
//procedure TOF_MULSELDOSS.ShowTabInSql (strSql:string; strTabName:string; strMenuName:string; strTab2Name:string='');
var
   Tab1, Tab2   :TTabSheet;
   i            :integer;
   strName      :string;
begin
     Tab1 := nil;
     if strTab1Name <> '' then
        Tab1 := TTabSheet (GetControl (strTab1Name));
     Tab2 := nil;
     if strTab2Name <> '' then
        Tab2 := TTabSheet (GetControl (strTab2Name));

     if Tab1 <> nil then
     begin
          for i := 0 to Tab1.ControlCount-1 do
          begin
               strName := Tab1.Controls [i].Name;
               if Pos (strName, strSql) > 0 then
               begin
                    Tab1.TabVisible := TRUE;

                    // $$$ JP 26/02/07: l'onglet lié doit également être affiché
                    if Tab2 <> nil then
                       Tab2.TabVisible := TRUE;

                    TMenuItem (GetControl (strMenuName)).Checked := TRUE;
                    exit;
               end;
          end;
     end;

     // $$$ JP 26/02/07: deuxième onglet si défini
     if Tab2 <> nil then
     begin
          for i := 0 to Tab2.ControlCount-1 do
          begin
               strName := Tab2.Controls [i].Name;
               if Pos (strName, strSql) > 0 then
               begin
                    Tab2.TabVisible := TRUE;

                    // $$$ JP 26/02/07: l'onglet lié doit également être affiché
                    if Tab1 <> nil then
                       Tab1.TabVisible := TRUE;

                    TMenuItem (GetControl (strMenuName)).Checked := TRUE;
                    exit;
               end;
          end;
     end;
end;

//-----------------------------
//--- Nom : BCherche_OnClick
//-----------------------------
procedure TOF_MULSELDOSS.BCHERCHE_OnClick(Sender: TObject);
var ChXXWhere : String;
    StrSQL    : String;
begin
 NextPrevControl(Ecran, True, True); // FQ 11630

 ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),TcheckBox (GetControl ('SANSGRPCONF')).Checked);

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

 ChXXWhere:=ChXXWhere+GererCritereApplication (THMultiValComboBox (GetControl ('LISTEAPPLICATION')));
 ChXXWhere:=ChXXWhere+GererCritereDivers ();

 //PGR 10/07/2006 Ajout critère base dossier = CEGIDEWEENVEXP
 //PGR 30/08/06 Récupération uniquement des dossiers correspondant à DOS_WINSTALL
 if GetParamSocSecur ('SO_MDLIENWINNER', False) and (GetEnvVar('CEGIDEWEENVEXP') <> '') then
   //ChXXWhere:=ChXXWhere+ 'AND (DOS_WINSTALL="' + GetEnvVar('CEGIDEWEENVEXP') + '" OR DOS_WINSTALL="")';
 //pgr 14/12/2007 Afficher systématiquement le dossier 000000
 begin
   if FileExists(ExtractFilePath(Application.ExeName)+ 'Trait000.nop') then
     ChXXWhere:=ChXXWhere+ ' AND (DOS_WINSTALL="' + GetEnvVar('CEGIDEWEENVEXP') + '")'
   else ChXXWhere:=ChXXWhere+ ' AND ((DOS_WINSTALL="' + GetEnvVar('CEGIDEWEENVEXP') + '") OR (DOS_NODOSSIER="000000"))';
 end;

 // $$$ JP 04/09/06: pour D7/Unicode: ne plus faire ce genre de cast... TEdit(GetControl('XX_WHERE')).Text:=ChXXWhere;
 SetControlText ('XX_WHERE', chXXWhere);

 // Traitement générique
 TFMul(Ecran).BChercheClick(Sender);

 // Repositionnement sur le dossier en cours
 LocaliseNoDossier;

 // $$$ JP 23/08/05 - il faut afficher les onglets manquants, s'ils ont un critère dans la requête active
 strSQL := UpperCase (Trim (Self.Qry.SQL.Text));
 System.Delete (strSQL, 1, Pos ('FROM', strSQL));
 ShowTabInSql (strSQL, 'PLIBRE1',      '',            'CRIT_INFOSLIBRES');
 ShowTabInSql (strSQL, 'PLIBRE2',      '',            'CRIT_INFOSLIBRES');
 ShowTabInSql (strSQL, 'PACTIVITE',    '',            'CRIT_ACTIVITE');
 ShowTabInSql (strSQL, 'PBNC',         '',            'CRIT_BNC');
 ShowTabInSql (strSQL, 'PASSOCIATION', '',            'CRIT_ASSO');
 ShowTabInSql (strSQL, 'PDPSOCIAL',    '',            'CRIT_SOCIAL');
 ShowTabInSql (strSQL, 'PSTATTIERS',   'PCOMPLTIERS', 'CRIT_TIERS'); // $$$ JP 26/02/07: PCOMPLTIERS doit avoir le même état que PSTATTIERS //ShowTabInSql (strSQL, 'PCOMPLTIERS',  'CRIT_TIERS', 'PSTATTIERS');
 ShowTabInSql (strSQL, 'PNETEXPERT',   '',            'CRIT_NETEXPERT');
 ShowTabInSql (strSQL, 'PDPAGRICOLE',   '',            'CRIT_AGRICOLE');
end;

function TOF_MULSELDOSS.DonnerListeCode (NomControle : String) : String;
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

procedure TOF_MULSELDOSS.BINSERT_OnClick(Sender: TObject);
begin
     if Not JaiLeDroitTag(26053) then
     begin
          PGIInfo('Vous n''avez pas accès à cette fonctionnalité.', TitreHalley);
          exit;
     end;

     // assistant création de dossier
     DP_LanceAssistDossier ('','');

     // actualisation
     BCHERCHE_OnClick(Sender);
end;

procedure TOF_MULSELDOSS.BCRITINFOSLIBRES_OnClick(Sender: TObject);
begin
     if m_bLibre1Empty = FALSE then
        SetControlVisible ('PLIBRE1', Not (TTabSheet (GetControl ('PLIBRE1')).TabVisible));
     if m_bLibre2Empty = FALSE then
        SetControlVisible ('PLIBRE2', Not (TTabSheet (GetControl ('PLIBRE2')).TabVisible));
     TMenuItem (GetControl ('CRIT_INFOSLIBRES')).Checked := (TTabSheet (GetControl ('PLIBRE1')).TabVisible or TTabSheet (GetControl ('PLIBRE2')).TabVisible);
end;

procedure TOF_MULSELDOSS.BCRITACTIVITE_OnClick(Sender: TObject);
begin
     SetControlVisible ('PACTIVITE', Not (TTabSheet (GetControl ('PACTIVITE')).TabVisible));
     TMenuItem (GetControl ('CRIT_ACTIVITE')).Checked := TTabSheet (GetControl ('PACTIVITE')).TabVisible;
end;

procedure TOF_MULSELDOSS.BCRITBNC_OnClick(Sender: TObject);
begin
     SetControlVisible ('PBNC', Not (TTabSheet (GetControl ('PBNC')).TabVisible));
     TMenuItem (GetControl ('CRIT_BNC')).Checked := TTabSheet (GetControl ('PBNC')).TabVisible;
end;

procedure TOF_MULSELDOSS.BCRITASSO_OnClick(Sender: TObject);
begin
     SetControlVisible ('PASSOCIATION', Not (TTabSheet (GetControl ('PASSOCIATION')).TabVisible));
     TMenuItem (GetControl ('CRIT_ASSO')).Checked := TTabSheet (GetControl ('PASSOCIATION')).TabVisible;
end;

procedure TOF_MULSELDOSS.BCRITTIERS_OnClick(Sender: TObject);
begin
     SetControlVisible ('PSTATTIERS',  Not (TTabSheet (GetControl ('PSTATTIERS')).TabVisible));
     SetControlVisible ('PCOMPLTIERS', Not (TTabSheet (GetControl ('PCOMPLTIERS')).TabVisible) );
     TMenuItem(GetControl ('CRIT_TIERS')).Checked := TTabSheet(GetControl ('PSTATTIERS')).TabVisible;
end;

procedure TOF_MULSELDOSS.BCRITSOCIAL_OnClick(Sender: TObject);
begin
     SetControlVisible ('PDPSOCIAL', Not (TTabSheet (GetControl ('PDPSOCIAL')).TabVisible));
     TMenuItem (GetControl ('CRIT_SOCIAL')).Checked := TTabSheet (GetControl('PDPSOCIAL')).TabVisible;
end;

procedure TOF_MULSELDOSS.BCRITNETEXPERT_OnClick(Sender: TObject);
begin
     SetControlVisible ('PNETEXPERT', Not (TTabSheet (GetControl ('PNETEXPERT')).TabVisible));
     TMenuItem (GetControl ('CRIT_NETEXPERT')).Checked := TTabSheet (GetControl('PNETEXPERT')).TabVisible;
end;

procedure TOF_MULSELDOSS.BCRITAGRICOLE_OnClick(Sender: TObject);
begin
     SetControlVisible ('PDPAGRICOLE', Not (TTabSheet (GetControl ('PDPAGRICOLE')).TabVisible));
     TMenuItem (GetControl ('CRIT_AGRICOLE')).Checked := TTabSheet (GetControl('PDPAGRICOLE')).TabVisible;
end;

procedure TOF_MULSELDOSS.SANSGRPCONF_OnClick(Sender: TObject);
begin
  GereCheckboxSansGrpConf(TCheckbox(GetControl('SANSGRPCONF')), THMultiValCombobox(GetControl('GROUPECONF')) );
end;

procedure TOF_MULSELDOSS.BCULTURES_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BCULTURES')).Checked;
 THLabel (GetControl ('LACTIVITE1')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBCULTURE')).Enabled:=Etat;
end;

procedure TOF_MULSELDOSS.BCULTURESP_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BCULTURESP')).Checked;
 THLabel (GetControl ('LACTIVITE2')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBCULTUREP')).Enabled:=Etat;
end;

procedure TOF_MULSELDOSS.BANIMAUX_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BANIMAUX')).Checked;
 THLabel (GetControl ('LACTIVITE3')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBANIMAUX')).Enabled:=Etat;
end;

procedure TOF_MULSELDOSS.BTRANSFORMATION_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BTRANSFORMATION')).Checked;
 THLabel (GetControl ('LACTIVITE4')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBTRANSFORMATION')).Enabled:=Etat;
end;

procedure TOF_MULSELDOSS.BAUTRES_OnClick (Sender : TObject);
var Etat : Boolean;
begin
 Etat:=TcheckBox (GetControl ('BAUTRES')).Checked;
 THLabel (GetControl ('LACTIVITE5')).Enabled:=Etat;
 THMultiValComboBox (GetControl ('CBAUTRES')).Enabled:=Etat;
end;

{procedure TOF_MULSELDOSS.BNOUVRECH_OnClick(Sender: TObject);
begin
     TFMul(Ecran).BNouvRechClick(Sender);

     // la recherche le remet à 0 ...                                  
     SetControlText('DFI_JOURDECLA', '-1');
end;}

procedure TOF_MULSELDOSS.OnNouvelleRecherche (Sender:TObject);
begin
     // Traitement AGL
     TFMul(Ecran).BNouvRechClick (Sender);

     // On veut -1 et pas 0 par défaut dans jour déclaration TVA
     SetControlText ('DFI_JOURDECLA', '-1');
end;

procedure TOF_MULSELDOSS.OnSupprimeFiltre (Sender:TObject);
begin
     // Traitement AGL
     TFMul(Ecran).BDelFiltreClick (Sender);

     // On veut -1 et pas 0 par défaut dans jour déclaration TVA
     SetControlText ('DFI_JOURDECLA', '-1');
end;

procedure TOF_MULSELDOSS.mnNouveauMessageClick(Sender: TObject);
Var NumDossier : String;
begin
 NumDossier := GetNoDossDansListeAvecPwd;
 if NumDossier='' then Exit;

 if (V_PGI.MailMethod=mmSMTP) and (VH_DP.SeriaMessagerie) then
  ShowNewMessage ('','','',DonnerAdresseEmail (NumDossier),NumDossier)
 else
  SendMail ('','','',nil,'',False);
end;

procedure TOF_MULSELDOSS.mnNouvActionClick(Sender: TObject);
Var Noclt,NumDossier : String;
    QQ : Tquery;
begin  //mcd 06/07/2005
    NumDossier := GetNoDossDansListeAvecPwd;
    if NumDossier='' then Exit;

    Noclt :='';
    QQ := Opensql ('SELECT ANN_TIERS FROM ANNUAIRE, DOSSIER WHERE ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="'+NumDossier +'"',true);
    if not QQ.eof then NoClt := QQ.FindField('ANN_TIERS').AsString ;
    Ferme(QQ);
    if Noclt ='' then PgiInfo('Pas de tiers associé à votre dossier', TitreHalley)
     else AglLanceFiche ('RT','RTACTIONS','','','FICHETIERS;RAC_TIERS='+NoClt+';ACTION=CREATION');
end;

procedure TOF_MULSELDOSS.mnApplicationsClick (Sender:TObject);
var
   UneAppliDoss : TAppliDoss;
   strAppli     :string;
begin
 if PasSurUnDossier then exit;
 // Obligation de sélectionner le dossier car tout le contexte de lancement des applis en dépend
 If Not LanceContexteDossier (GetNoDossDansListeAvecPwd) then exit;

 // $$$ JP 13/04/06 - pour gérer les multiples instances d'une même appli (sans doute pas normal, mais faut le gérer)
 strAppli := TMenuItem (Sender).Name;
 while strAppli [Length (strAppli)] = '_' do
       system.Delete (strAppli, Length (strAppli), 1);

 // Ne rien faire si appli winner alors qu'elles sont désactivées
 if Not GetParamSocSecur('SO_MDLIENWINNER', False) And (pos('WISUITE', strAppli) > 0) then exit;
 if (pos('WISUITE', TMenuItem(Sender).Name) > 0) then
   UneAppliDoss := TAppliDoss.Create (VH_Doss.NoDossier, True, 'WISUITE.EXE /'+copy (strAppli, 10, 5), nil, nil, InfosWinner)
 else //autre appli
   UneAppliDoss := TAppliDoss.Create (VH_Doss.NoDossier,True, copy (strAppli, 3, Length (strAppli)-2) + '.exe');

 if (UneAppliDoss.TesterEtatDossier (True)) then
  begin
   if not UneAppliDoss.AutoriserExecution(UneAppliDoss.InfoAppli.Nom) then exit;
   if not UneAppliDoss.AutoriserExecutionSiTransport then exit;
   if not UneAppliDoss.VerifierPwd then exit;

   {$IFDEF EAGLCLIENT}
   if not UneAppliDoss.VerifierInstallation then exit;
   {$ENDIF}
   if pos('WISUITE', UneAppliDoss.InfoAppli.Nom) > 0 then
     InfosWinner.LanceAppliWinner(UneAppliDoss.ChNoDossier, UneAppliDoss.InfoAppli)
   else
     LanceAppliMultiDossier(UneAppliDoss.InfoAppli);
  end;

 UneAppliDoss.Free;
end;

procedure TOF_MULSELDOSS.mnNouvelleActiviteClick (Sender:TObject);
var
   strDossier         :string;
begin
     // Récupération du dossier en sélection par clic droit
     strDossier := GetNoDossDansListeAvecPwd;
     if strDossier <> '' then
        AgendaNewActivite (V_PGI.User, strDossier, TMenuItem (Sender).Name = 'MnActiviteExterne', TMenuItem (Sender).Name = 'MnAbsence');
end;

procedure TOF_MULSELDOSS.mnEchangesEtNetExpertClick(Sender: TObject);
var NumDossier : String;
    bComptaActive : Boolean;
begin
  inherited;
  NumDossier := GetNoDossDansListeAvecPwd;
  if NumDossier='' then exit;

  // 21/10/05 : on sélectionne le dossier demandé, car si après
  // on va sur les services ASP, on sera pas sur le bon dossier...
  if Not LanceContexteDossier(NumDossier) then exit;

  //--- Activation ASP
  if NE_Connecter (NumDossier) then
   begin
    if VH_DP.EwsActif then
     begin
      if not EWS_ConnecterBusinessLine (NumDossier) then
       begin
        NE_Deconnecter ();
        exit;
       end;

      //--- Ouvre la fiche pour choix base modèle et rappel du service actif
      // FQ 11863 - pas de création du dossier BL depuis le bureau sans base expert compta
      // (= mode avec application Compta Winner qui se chargera de créer)
      if ExisteSQL('SELECT 1 FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+NumDossier+'" AND DAP_NOMEXEC="CCS5.EXE"')
      and (EWS_TypeService (EWSCOMPTA,ParamEws.Produit) or EWS_TypeService (EWSPAYE,ParamEws.Produit)) then
       LancerParamS1ASP (NumDossier);
     end
    else
     begin
      // sans eWS, on est forcément ici avec de la compta
      //--- Ouvre la fiche pour saisie CD key et choix base modèle
      LancerParamS1ASP (NumDossier);
     end;
    //--- Fermeture des connexions ASP
    NE_Deconnecter ();
   end;
end;

procedure TOF_MULSELDOSS.MnEnCoursClientClick (Sender : TObject);
Var NumDossier : String;
begin
 Inherited;
 NumDossier:=GetNoDossDansListeAvecPwd;
 if NumDossier='' then exit;
 ConsulterSolde (NumDossier);
end;

procedure TOF_MULSELDOSS.mnSynthDossierClick(Sender: TObject);
begin
  inherited;
  // la fiche synthèse ne sera pas en modale mais plein écran
  // donc sélectionne le dossier au préalable (= double clic)
  Lst.OnDblClick(Lst);
end;

procedure TOF_MULSELDOSS.mnConfDossierClick(Sender: TObject);
var
    sNoDoss, sGroupes : String;
begin
  inherited;
  // récupère l'enreg. en cours
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;

  sGroupes := LanceFicheSelectGroupeConf(sNoDoss);
  if sGroupes<>'##ABANDON##' then
    begin
    AffecteGroupe (sNoDoss, sGroupes);
    AffecteDosGrpConf(sNoDoss);
    // actualisation
    BCHERCHE_OnClick(Sender);
    end;
end;

procedure TOF_MULSELDOSS.mnPwdDossierClick(Sender: TObject);
var sNoDoss: String;
begin
  inherited;
  // récupère l'enreg. en cours
  if PasSurUnDossier then exit;
  sNoDoss := GetField('DOS_NODOSSIER');
  // propose la modif du mot de passe
  OuvreChangePwdDossier(sNoDoss);
  // actualisation
  BCHERCHE_OnClick(Sender);
end;

procedure TOF_MULSELDOSS.mnMarquePourTransportClick(Sender: TObject);
var sNoDoss, dossverrou: String;
// ENL : en ligne
// MAR : marqué pour transport
// PAR : parti
begin
  inherited;
  // récupère l'enreg. en cours
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;
  if IsDossierUtilise (sNoDoss) then
    begin
    PGIInfo ('Le dossier '+sNoDoss+' est en cours d''utilisation et ne peut donc pas être marqué pour transport.', TitreHalley);
    exit;
    end;
  // marquage/démarquage
  dossverrou := MarqueDossier(sNoDoss);
end;

procedure TOF_MULSELDOSS.mnBlocageFonctionnelClick (Sender: TObject);
var
   sNoDoss   :string;
begin
  // récupère l'enreg. en cours
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;
  if IsDossierUtilise (sNoDoss) then
    begin
    PGIInfo ('Le dossier '+sNoDoss+' est en cours d''utilisation et ne peut donc pas être bloqué fonctionnellement.', TitreHalley);
    exit;
    end;

  // Blocage ou déblocage
  if mnBlocageFonctionnel.Checked = FALSE then
  begin
       //if PgiAsk ('Confirmez-vous le blocage fonctionnel du dossier: ' + sNoDoss + ' ?') = mrYes then
       if ExecuteSQL ('UPDATE DOSSIER SET DOS_VERROU="BLO",DOS_UTILISATEUR="' + V_PGI.User + '" WHERE DOS_NODOSSIER="' + sNoDoss + '" AND (DOS_VERROU="ENL" OR DOS_VERROU="")') < 1 then
          PgiInfo ('Blocage fonctionnel impossible (dossier peut être marqué: transport, mise à jour...)');
  end
  else
  begin
       //if PgiAsk ('Confirmez-vous le déblocage fonctionnel du dossier: ' + sNoDoss + ' ?') = mrYes then
       if ExecuteSQL ('UPDATE DOSSIER SET DOS_VERROU="ENL",DOS_UTILISATEUR="' + V_PGI.User + '" WHERE DOS_NODOSSIER="' + sNoDoss + '" AND DOS_VERROU="BLO"') < 1 then
          PgiInfo ('Déblocage fonctionnel impossible (dossier peut être marqué: transport, mise à jour...)');
  end;
end;

procedure TOF_MULSELDOSS.mnLiensannuaireClick(Sender: TObject);
// clic-droit, Annuaire du dossier
// = Liste des liens rattachés au dossier
var sNoDoss, GuidPerDos: String;
begin
  inherited;
  // récupère l'enreg. en cours
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;

  GuidPerDos := GetField('DOS_GUIDPER');
  // sans le test ANL_TRI, on voit les lignes génériques "INTERVENANT" de juri, mais pas grave...
  // AglLanceFiche('DP','LIENINTERDOS','ANL_GUIDPERDOS='+GUIDPerDos+' AND ANL_TRI>10','',GUIDPerDos+';DOS')
  AglLanceFiche('DP','LIENINTERDOS','ANL_GUIDPERDOS='+GuidPerDos,'',GuidPerDos+';DOS');
end;

procedure TOF_MULSELDOSS.mnLiensdelapersonneClick(Sender: TObject);
// clic-droit, Liens de la personne
// attention : reçoit un GuidPer dans les paramètres, or la tof annulien
// attend un Guidperdos à cet endroit, mais normal car ...
var sNoDoss, GuidPerDos: String;
begin
  inherited;
  // récupère l'enreg. en cours
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;

  GuidPerDos :=GetField('DOS_GUIDPER');
  // ... le test est fait sur ANL_GUIDPER, car il s'agit d'une liste des liens
  // dans laquelle INTERVIENT la personne du dossier en cours
  // NPC avec juste au dessus : liste des personnes LIEES au dossier en cours
  AglLanceFiche('DP','LIENPERSONNE','ANL_GUIDPER='+GuidPerDos,'',GuidPerDos+';DOS');
end;

procedure TOF_MULSELDOSS.mnEvenementsClick(Sender: TObject);
// en fait l'intitulé du menu est "Messages"
var nodoss : String;
begin
  if PasSurUnDossier then exit;
  nodoss := GetField('DOS_NODOSSIER');

  // Sélectionne le dossier demandé
  if Not LanceContexteDossier(nodoss) then exit;

  // Messagerie du dossier
  // => dispatch permet de ne pas l'ouvrir en modale (=referme fiche en cours)
  FMenuG.LanceDispatch(76530);
end;

procedure TOF_MULSELDOSS.PPM_OnPopup(Sender: TObject);
var nodoss, marq, usr, login: String;
    mnEncoursClient : TMenuItem;
    bMenuActif, bBaseExiste, bSansBase, bVersionOk : Boolean;
    VDoss : Integer;
    Mn,MnPlaq,MnRacine : TMenuItem;
    Indice             : Integer;
    Appli              : TPgiAppli;
    i : integer;
    strMnName          :string;
begin
  inherited;
  if PasSurUnDossier then
    begin
    ActiveContenuPopup(False);
    mnSynthDossier.Caption := 'Synthèse du dossier ?';
    exit;
    end
  else
    ActiveContenuPopup(True);

  nodoss := GetField('DOS_NODOSSIER');
  mnSynthDossier.Caption := 'Synthèse du dossier '+nodoss;

  //---- Menu répertoire téléphonique du dossier en cours
  MajMenuRepertoireTel(GetField('DOS_GUIDPER'));

  //---- Menu "Marquer pour transport" (dynamique)
  // Existence de la base
  bBaseExiste := DBExists ('DB' + NoDoss);

  // Version de la base
  if bBaseExiste then
   begin
    //--- Version de la base = version base en cours
    VDoss := VersionDossier(nodoss);
    bVersionOk := (VDoss=V_PGI.NumVersionSoc);
   end
  else
   bVersionOk := False;

  // par défaut
  mnMarquePourTransport.Enabled := False;
  mnMarquePourTransport.Caption := 'Marquer pour transport';
  mnMarquePourTransport.Checked := False;
  mnMarquePourTransport.Visible := True;
  // $$$ JP 25/04/06 SetControlVisible('mnSep5', True);

  if nodoss='' then exit;

  // état du marquage/démarquage
  marq := EtatMarqueDossier(nodoss, usr, login, bSansBase);

  // Correction anomalie éventuelle
  if bBaseExiste and bSansBase then
  begin
       ExecuteSQL('UPDATE DOSSIER SET DOS_ABSENT="-" WHERE DOS_NODOSSIER="'+nodoss+'"');
       bSansBase := False;
  end;

  // ancien statut obsolète
  if marq='ABS' then marq := 'ENL';

  // si dossier du cabinet, on interdit le marquage
  if (nodoss='000000') or (nodoss=NoDossierBaseCommune) then
  begin
       mnMarquePourTransport.Caption := 'Dossier non transportable';
  end
  else
  begin
       // dossier présent (bloquer fonctionnellement ou pas)
       // $$$ JP 26/04/06
       if (marq = 'ENL') or (marq = 'BLO') then
       begin
            // Dossier sans base
            //PGR 06/09/2006 Transport dossiers Winner
            //if Not bBaseExiste then
            //  mnMarquePourTransport.Caption := 'Dossier sans base de production'
            //else if Not bVersionOk then
            if (Not bBaseExiste) and (GetParamSocSecur ('SO_MDLIENWINNER', False) = False) then
                 mnMarquePourTransport.Caption := 'Dossier sans base de production'
            else if (bBaseExiste) and (Not bVersionOk) then
                 mnMarquePourTransport.Caption := 'Dossier à mettre à jour'
            // autorise le marquage (sauf sur une base locale !)
            else if (VH_DP.LeMode<>'L') and (marq = 'ENL') then // $$$ JP 26/04/06
                 mnMarquePourTransport.Enabled := True
            // rien à indiquer
            else
            //begin
                 mnMarquePourTransport.Visible := False;
                 // $$$ JP 25/04/06 SetControlVisible('mnSep5', False);
            //end;
       end
       // dossier marqué pour transport
       else if (marq = 'MAR') then
       begin
            mnMarquePourTransport.checked := True;
            mnMarquePourTransport.caption := 'Dossier marqué par '+login;
            // marqué par soi ou par un inconnu => dégriser menu
            // (sauf sur une base locale !)
            if ((usr=V_PGI.User) or (login='')) and (VH_DP.LeMode<>'L') then
               mnMarquePourTransport.enabled := True;
       end
       // dossier parti
       else if (marq = 'PAR') then
       //begin
            mnMarquePourTransport.Caption := 'Dossier emporté par '+login
       //end
       // dossier en cours de réparation
       else if (marq = 'SOS') then
       //begin
            mnMarquePourTransport.Caption := 'Dossier en cours de réparation par '+login
       else
           mnMarquePourTransport.Visible := FALSE; // $$$ JP 25/04/06

       //end
  end; // $$$ JP 25/04/06

  // dossier bloqué
  // $$$ JP 25/04/06 - FQ 10144, et surtout utilisation d'un autre menu pour le blocage fonctionnel
  if (marq = 'BLO') then
  begin
       mnBlocageFonctionnel.Caption := 'Bloqué fonctionnellement par ' + login;
       mnBlocageFonctionnel.Checked := TRUE;
       if usr <> V_PGI.User then
          mnBLocageFonctionnel.Enabled := FALSE;
       mnBlocageFonctionnel.Visible := TRUE;
  end
  else
  begin
       if (JaiLeDroitConceptBureau (26056) = TRUE) and ( (marq = '') or (marq = 'ENL')) then
       begin
            mnBlocageFonctionnel.Caption := 'Bloquer fonctionnellement';
            mnBlocageFonctionnel.Checked := FALSE;
            mnBlocageFonctionnel.Enabled := TRUE;
            mnBlocageFonctionnel.Visible := TRUE;
       end
       else
           mnBlocageFonctionnel.Visible := FALSE;
  end;

  // $$$ JP 25/04/06
  SetControlVisible ('mnSep5', mnMarquePourTransport.Visible and mnBlocageFonctionnel.Visible);
  
  //--- Menu En cours client : affiche juste le solde
  mnEncoursClient := TMenuItem(GetControl('mnEncoursClient'));
  mnEncoursClient.Caption := 'En-cours client : '+ FormaterMontant(CalculerSolde(nodoss, False));
  if ExisteSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_NODOSSIER="'+NoDoss+'" AND DOS_CABINET="X"') then
   MnEncoursClient.Enabled:=False;

  //--- Menu EDI
  SetControlVisible('mnSep8',VH_DP.SeriaTDI);
  SetControlVisible('mnPriseEnCompteEDI',False); // VH_DP.SeriaTDI);
  SetControlVisible('mnEnvoiEDI',False); // VH_DP.SeriaTDI);
  SetControlVisible('mnJournalEDI',VH_DP.SeriaTDI);
  SetControlVisible('mnJournalSyntheseEDI',VH_DP.SeriaTDI);

  //--- Menu Activer dossier ASP : uniquement si on a le droit, et URL existe
  bMenuActif:= bDroitActiverDossierAsp and VH_DP.NetExpertActif
  // et dossier accessible et (eWS actif ou compta activée)
  and ( (marq = 'ENL') and (bVersionOk) and ( VH_DP.EwsActif or ExisteSQL('SELECT 1 FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+nodoss+'" AND DAP_NOMEXEC="CCS5.EXE"') )
  //    ou dossier sans base et eWS actif
     or ( bSansBase and VH_DP.EwsActif ) );
  SetControlVisible('mnEchangesEtNetExpert', bMenuActif);

  //--- Menu Application
  if (GetControl('mnApplications')<>Nil) and (GetParamSocSecur('SO_MDLISTEAPPLIS', False)) then
   begin
    MnRacine:=TMenuItem (GetControl ('MnApplications'));
    MnRacine.Clear;

    //--- Ajouter les applis Winner au tableau.
    if GetParamSocSecur('SO_MDLIENWINNER', False) then
      if InfosWinner.getInfosInstallWinner(NoDoss) then
        begin
          For i := 0 to InfosWinner.listeApplisWinnerCodes.Count -1 do
            begin
            InfosWinner.getAppliWinner (InfosWinner.listeApplisWinnerCodes[i]);
            end;
        end;

    for Indice:=0 to V_Applis.Applis.count-1 do
     begin
      Appli:=TPGIAppli (V_Applis.Applis [Indice]);
      //--- Ne pas prendre les applis eAgl
      {$IFDEF EAGLCLIENT}
       if Not appli.eAgl then Continue;
      {$ELSE}
       if Appli.eAgl then Continue;
      {$ENDIF}

      //--- Ne pas prendre les gestion des standards
      if Appli.Std then Continue;
      //--- Sauf Bureau PGI (mnu rajouté pour contenir la clé de séria)
      {$IFDEF EAGLCLIENT}
       if appli.Nom='ECEGIDPGI.EXE' then Continue;
      {$ELSE}
       if appli.Nom='CEGIDPGI.EXE' then Continue;
      {$ENDIF}

      //--- Teste autorisation du groupe sur cette appli
      if (not JaiLeDroitTag(ValeurI(Appli.Tag))) then Continue;

      if Appli.Visible then
       begin
        // les SetFlagAppli sont faits avec les noms eagl, donc on compare le nom de l'appli en enlevant le e
        {$IFDEF EAGLCLIENT}
        if ExisteSQl ('Select DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+NoDoss+'" AND DAP_NOMEXEC="'+ copy (Appli.Nom, 2, Length (Appli.Nom)-1) +'"') then
        {$ELSE}
        if ExisteSQl ('Select DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+NoDoss+'" AND DAP_NOMEXEC="'+Appli.Nom+'"') then
        {$ENDIF}
         begin
          if (not Appli.plaq) then
           begin
            // $$$ JP 13/04/06 Mn:=TMenuItem.create (MnRacine);
            //si appli winner, mettre un autre nom
            if Not GetParamSocSecur('SO_MDLIENWINNER', False) And (pos('WISUITE', appli.nom) > 0) then continue;
            if pos('WISUITE', appli.nom) > 0 then
              strMnName {Mn.Name} := 'MnWISUITE'+InfosWinner.ExtractAppli(appli.nom)
            else//autre appli
              strMnName {Mn.Name} :='Mn'+ChangeFileExt (Appli.Nom,'');

            // $$$ JP 13/04/06 - au cas où une appli existerait plusieurs fois dans la liste (apparement, ça arrive parfois)
            while (mnRacine.FindComponent (strMnName) <> nil) and (Length (strMnName) < 255) do
                  strMnName := strMnName + '_';
            Mn:=TMenuItem.create (MnRacine);
            Mn.Name := strMnName;
            Mn.Caption:=Appli.Titre;
            Mn.OnClick:=MnApplicationsClick;
            MnRacine.Add(Mn);
           end
          else
           begin
            //--- Construction du sous-menu "plaquette"
            Mn:=MnRacine;
            MnPlaq:=Mn.Find ('Plaquettes');
            //--- Si le sous menu "plaquette" n'existe pas déjà on le créé
            if (MnPlaq=nil) then
             begin
              MnPlaq:=TMenuItem.create (MnRacine);
              MnPlaq.Name:='MnPlaquettes';
              MnPlaq.Caption:='Plaquettes';
              MnRacine.Add (MnPlaq);
             end;
            strMnName := 'Mn'+ChangeFileExt (Appli.Nom,'');

            // $$$ JP 13/04/06 - au cas où une appli existerait plusieurs fois dans la liste (apparement, ça arrive parfois)
            while (mnRacine.FindComponent (strMnName) <> nil) and (Length (strMnName) < 255) do
                  strMnName := strMnName + '_';
            Mn:=TMenuItem.create (MnPlaq);
            Mn.Name:= strMnName; //'Mn'+ChangeFileExt (Appli.Nom,'');
            Mn.Caption:=Appli.Titre;
            Mn.Onclick:=MnApplicationsClick;
            MnPlaq.Add (Mn);
           end
         end;
       end;
     end;

    SetControlVisible('mnSep7',MnRacine.count>0);
    SetControlVisible('mnApplications',MnRacine.count>0);
   end
  else
   begin
    SetControlVisible('mnSep7',False);
    SetControlVisible('mnApplications',False);
   end;
end;

procedure TOF_MULSELDOSS.LocaliseNoDossier;
// Redonne le focus dans la grille au dossier précédemment sélectionné,
// sauf si le dossier ne fait plus partie de la nouvelle recherche...
var sNoDossier : String;
begin
{$IFDEF EAGLCLIENT}
  sNoDossier := '';
  if VH_Doss=Nil then exit;
  if (VH_Doss<>Nil) and (VH_Doss.NoDossier<>'') then sNoDossier := VH_Doss.NoDossier;
  if sNoDossier<>'' then
    begin
    // En eAGL, le mul n'affiche qu'une 1ère page limitée d'enregistrements :
    // - si on trouve l'enreg dans TQ, c'est qu'il est affiché dans la grille,
    // - sinon, on ne va pas chercher à charger tous les enreg, ils n'ont qu'à
    //   mettre des critères de recherche plus limités (ou alors on peut mettre
    //   sNoDossier dans le critère dos_nodossier inférieur et rebalancer bchercheclick,
    //   mais si ils retournent en sélection de dossier, c'est pour changer de dossier !)
    if Qry.TQ.Locate('DOS_NODOSSIER', sNoDossier, []) then
      Lst.Row := Qry.TQ.CurrentFilleIndex + 1;
    end;
{$ELSE}
  sNoDossier := '';
  if (VH_Doss<>Nil) and (VH_Doss.NoDossier<>'') then sNoDossier := VH_Doss.NoDossier;
  if sNoDossier<>'' then
    begin
    if Qry.Active then // #### contourne bug de l'onglet "avancé" utilisé seul
      Qry.Locate('DOS_NODOSSIER', sNoDossier, []);
    end;
  // 12/07/04 NE PAS SELECTIONNER UN DOSSIER PAR DEFAUT
 { else GetField('DOS_NODOSSIER')<>'' then
    // pas de dossier en cours, on en sélectionne un par défaut, celui pointé par la grille
    begin
    sNoDossier := VarToStr(GetField('DOS_NODOSSIER'));
    if Not LanceContexteDossier(sNoDossier) then exit;
    end; }
{$ENDIF}
end;

// $$$ JP 20/10/2004 - ANN_NOM1 prioritaire par rapport à ANN_NOMPER dans le message
function TOF_MULSELDOSS.GetLibDossier: String;
var
   msg    :string;
begin
     msg := '';

     try
        msg := Trim (GetField ('DOS_LIBELLE'));
     except
     end;

     if msg='' then
        try
           msg := Trim (GetField ('ANN_NOM1'));
        except
        end;

     if msg='' then
        try
           msg := Trim (GetField ('ANN_NOMPER'));
        except
        end;

     if msg <> '' then
        msg := ' (' + msg + ')';

     Result := Trim (GetField ('DOS_NODOSSIER')) + msg;
end;

function TOF_MULSELDOSS.ConfirmeAvecListeAppli: String;
// Message de confirmation pour la suppression du dossier en cours
var Qappli : TQuery;
    txt, nodoss : String;
    appli: TPGIAppli;
begin
  if PasSurUnDossier then exit;
  nodoss := GetField('DOS_NODOSSIER');
  Qappli := OpenSQL('SELECT DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+nodoss+'"', True, -1, '', True);
  txt :='';
  If Not Qappli.eof then
  txt := txt +'Le dossier '+GetLibDossier+' a utilisé les modules suivants :#13';
  while Not Qappli.eof do
    begin
{$IFDEF EAGLCLIENT}
    appli := V_Applis.Lappli('E'+Qappli.FindField('DAP_NOMEXEC').AsString);
{$ELSE}
    appli := V_Applis.Lappli(Qappli.FindField('DAP_NOMEXEC').AsString);
{$ENDIF}
    if appli<>Nil then txt := txt +'    - '+ appli.Titre + #13;
    Qappli.Next;
    end;
  If ExisteSQL('SELECT 1 FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+nodoss+'" AND DAP_NOMEXEC="CPS5.EXE"') then
  txt := txt +' Attention : les informations propres au dossier de paie (paramétrages prédéfinis dossier, etc...)#13'
      +' vont être définitivement supprimées.#13';
  Ferme(Qappli);
  
  // confirmation
  txt := txt + ' Voulez-vous vraiment supprimer le dossier '+GetLibDossier + ' ainsi que sa base de production ?';
  Result := txt;
end;

procedure TOF_MULSELDOSS.ResetTimer(Sender: TObject);
begin
  // idem proc dans Mul
  TTimer(GetControl('SearchTimer')).Enabled:=FALSE ;
  TTimer(GetControl('SearchTimer')).Enabled:=TRUE ;
end;

{$IFNDEF GALDISPATCH}
procedure TOF_MULSELDOSS.TimerDest_OnTimer(Sender: TObject);
// Accède au dossier client
begin
  TimerDest.Enabled := False;
  // si pas vrai dossier, on ne fait rien (normalement, jamais le cas)
  if DossierDest='' then exit;

  // sélectionne le dossier demandé
  if VH_Doss=Nil then VH_Doss := TDossSelect.Create;
  if VH_Doss.NoDossier<>DossierDest then
    if Not LanceContexteDossier(DossierDest) then exit;

  // Passe au module "Dossier client"
  FMenuG.LanceDispatch(76501);
end;
{$ELSE}
procedure TOF_MULSELDOSS.LanceFicheDossier (strDossier:string);
begin
    // sélectionne le dossier demandé
    if strDossier <> '' then
    begin
         if VH_Doss = nil then
            VH_Doss := TDossSelect.Create;
         if VH_Doss.NoDossier <> strDossier then
            if not LanceContexteDossier (strDossier) then
               exit;

         // Dispatch le tag correspondant à la fiche client (fiche synthèse)
         GalDispatch (76501);
    end;
end;
{$ENDIF}

procedure TOF_MULSELDOSS.BOUVRIR_OnClick(Sender: TObject);
begin
  if PasSurUnDossier then exit;
  // C'est le timer qui fera l'accès au dossier client,
  // sinon access vio en changeant de fiche inside sur bouvrir
{$IFNDEF GALDISPATCH}
  DossierDest := GetField('DOS_NODOSSIER');
  TimerDest.Enabled := True;
{$ELSE}
   LanceFicheDossier (GetField ('DOS_NODOSSIER'));
{$ENDIF}
end;

procedure TOF_MULSELDOSS.MajMenuRepertoireTel(sGuidPer : String);
// Ajoute le répertoire téléphonique de la personne, en sous-menu du popup
var
    MnRep, MN: TMenuItem;
    TobRep, T : TOB;
    i : Integer;
    ligne : String;

    procedure AjouteLigneMnRep (strTel:string='');
    begin
      MN := TMenuItem.Create(MnRep);
      MN.Caption := ligne;
      MN.Checked := ( T.GetValue('PRINC')='X' );

      // $$$ JP: sur click, il faut soit ouvrir la fiche, soit téléphoner (si cti activé)
{$IFDEF BUREAU}
      strTel := Trim (strTel);
      if (VH_DP.ctiAlerte <> nil) and (strTel <> '') then
      begin
           MN.Hint    := strTel;
           MN.OnClick := MContactClick;
      end;
{$ENDIF}

      MnRep.Add(MN);
    end;

begin
  if GetControl('mnRepertoiretel')=Nil then exit;
  MnRep := TMenuItem(GetControl('mnRepertoiretel'));

  // Supprime les anciennes lignes
  for i:=MnRep.Count-1 downto 0 do MnRep.Delete(i);

  // Cherche liste des contacts
  TobRep := TOB.Create('', Nil, -1);

  TobRep.LoadDetailFromSQL(
     'SELECT "X" AS PRINC, ANN_CV||" "||ANN_NOM1||" "||ANN_NOM2 AS NOM,'
             //mcd 12/2005  +' ANN_TEL1 AS ANI_TEL1, ANN_TEL2 AS ANI_TEL2, ANN_FAX AS ANI_FAX'
    +' ANN_TEL1 AS C_TELEPHONE, ANN_TEL2 AS C_TELEX, ANN_FAX AS C_FAX'
    +' FROM ANNUAIRE WHERE ANN_GUIDPER="'+sGuidPer+'"'
    +' UNION '
(* mcd 12/2005    +'SELECT "-" AS PRINC, ANI_CV||" "||ANI_NOM||" "||ANI_PRENOM AS NOM,'
    +' ANI_TEL1, ANI_TEL2, ANI_FAX'
    +' FROM ANNUINTERLOC'
    +' WHERE ANI_CODEPER='+sCodePer  *)
    +'SELECT "-" AS PRINC, C_CIVILITE||" "||C_NOM||" "||C_PRENOM AS NOM,'
    +' C_TELEPHONE, C_TELEX, C_FAX'
    +' FROM CONTACT'
    +' WHERE C_GUIDPER="'+sGuidPer
    +'" ORDER BY NOM');

  for i:=0 to TobRep.Detail.Count-1 do
    begin
    T := TobRep.Detail[i];
    // Attention : on n'a droit qu'à une seule tabulation pour que le popup
    // gère correctement le multi-colonnage (les autres seraient vus
    // comme des carrés)
    if T.GetValue('C_TELEPHONE')<>'' then
      begin
      ligne := Trim(T.GetValue('NOM'))+char(9)+'Téléphone : '+T.GetValue('C_TELEPHONE');
      AjouteLigneMnRep (T.GetString ('C_TELEPHONE'));
      end;
    if T.GetValue('C_TELEX')<>'' then
      begin
      ligne := Trim(T.GetValue('NOM'))+char(9)+'Autre tél. :   '+T.GetValue('C_TELEX');
      AjouteLigneMnRep (T.GetString ('C_TELEX'));
      end;
    if T.GetValue('C_FAX')<>'' then
      begin
      ligne := Trim(T.GetValue('NOM'))+char(9)+'Fax :            '+T.GetValue('C_FAX');
      AjouteLigneMnRep;
      end;
    end;
  TobRep.Free;

end;

procedure TOF_MULSELDOSS.mnActiverClick(Sender: TObject);
var nodoss : String;
begin
  nodoss := GetField('DOS_NODOSSIER');
  if nodoss = '' then exit;

  // sélectionne le dossier demandé
  LanceContexteDossier(nodoss);
end;

procedure TOF_MULSELDOSS.mnTabComptaClick(Sender: TObject);
var sNoDoss: String;
begin
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;

  Aff_TableauBord(sNoDoss, COMPTA);
end;

procedure TOF_MULSELDOSS.mnTabPaieClick(Sender: TObject);
var sNoDoss: String;
begin
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;

  Aff_TableauBord(sNoDoss, PAIE);
end;

procedure TOF_MULSELDOSS.mnTabActiviteClick (Sender:TObject);
var sNoDoss: String;
begin
  sNoDoss := GetNoDossDansListeAvecPwd;
  if sNoDoss='' then exit;

  Aff_TableauBordGeneral (sNoDoss);
end;

function TOF_MULSELDOSS.GetNoDossDansListeAvecPwd: String;
var sNoDoss : String;
begin
  Result := '';

  // récupère l'enreg. en cours
  if PasSurUnDossier then exit;
  sNoDoss := GetField('DOS_NODOSSIER');

  // si ce n'est pas le dossier déjà sélectionné, on vérifie le password
  if (VH_Doss.NoDossier<>sNoDoss) and (Not VerifPwdDossier(sNoDoss)) then exit;

  Result := sNoDoss;
end;


function TOF_MULSELDOSS.PasSurUnDossier: Boolean;
begin
{
  Result := True;
  // Test si AUCUN enreg dans la liste
  if (Qry.Eof) and (Qry.Bof) then exit;
  // Teste si l'enreg en cours anormal (?)
  if GetField('DOS_NODOSSIER')='' then exit;
  Result := False;
} // plus simple !!! :
  Result := VarIsNull(GetField('DOS_NODOSSIER'));
end;

procedure TOF_MULSELDOSS.ActiveContenuPopup(bActif: Boolean);
var i : Integer;
begin
  for i := 0 to TPopupMenu(GetControl('PPM')).Items.Count-1 do
    if TPopupMenu(GetControl('PPM')).Items[i].Visible then
      TPopupMenu(GetControl('PPM')).Items[i].Enabled := bActif;
end;

procedure TOF_MULSELDOSS.VireOngletNetExpert;
begin
 // #### A VIRER QUAND ON SAURA QUOI FAIRE DE L'ONGLET !
 TMenuItem (GetControl ('CRIT_NETEXPERT')).Visible:=False;
 SetControlVisible ('PNETEXPERT',False);
 TMenuItem(GetControl('mnEchangesEtNetExpert')).Caption:='Activation du dossier en ASP';
end;

{$IFDEF BUREAU}
procedure TOF_MULSELDOSS.MContactClick (Sender:TObject);
var
   strTel   :string;
begin
     strTel := TMenuItem (Sender).Hint;
     if strTel <> '' then
        if PgiAsk ('Appeler le ' + strTel + ' ?') = mrYes then
           VH_DP.ctiAlerte.MakeCall (strTel, 'ANN_GUIDPER=' + GetField ('DOS_GUIDPER'));
end;
{$ENDIF}

procedure TOF_MULSELDOSS.mnConfProfilDossierClick(Sender: TObject);
begin
  AGLLanceFiche('YY','YYUSERGROUP','','','DOSSIER='+GetField ('DOS_NODOSSIER')) ;
end;

//PGR 08/2007 Infos Install Winner
procedure TOF_MULSELDOSS.BInfosWinner_OnClick(Sender: TObject);
begin
  LanceFiche_WinnerInfosInstall;
end;

procedure TOF_MULSELDOSS.MnPriseEnCompteEDIClick (Sender : TObject);
var SNoDoss : String;
begin
 sNoDoss := GetField('DOS_NODOSSIER');

 AFF_Receptionner (SNoDoss);
end;

procedure TOF_MULSELDOSS.MnEnvoiEDIClick (Sender : TObject);
var SNoDoss : String;
begin
 sNoDoss := GetField('DOS_NODOSSIER');

 AglLanceFiche ('FIS','TDI_ENVOYER','','',SNoDoss);
end;

procedure TOF_MULSELDOSS.MnJournalEDIClick (Sender : TObject);
var SNoDoss : String;
begin
 sNoDoss := GetField('DOS_NODOSSIER');

 AglLanceFiche ('FIS','TDI_JOURNAL','','','FALSE|FALSE|'+SNoDoss);
end;


procedure TOF_MULSELDOSS.MnJournalSyntheseEDIClick (Sender : TObject);
var SNoDoss : String;
begin
 sNoDoss := GetField('DOS_NODOSSIER');
 AglLanceFiche ('FIS','TDI_SYNTHESE_DOS','','','|'+SNoDoss);
end;


initialization
  registerclasses ( [TOF_MULSELDOSS] ) ;


end.

