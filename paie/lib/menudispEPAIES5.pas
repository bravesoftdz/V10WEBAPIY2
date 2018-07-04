{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 06/07/2001
Modifié le ... :   /  /
Description .. : Menu de la paie
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 06/08/2001 V547 PH rajout fonctions menu 43 eAgl car il n'a plus besoins
                           de script dans cws tout est dans la procedure
                           dispatch
PT2   : 17/09/2001 V547 SB Copy des .dot dans le répertoire modèles
                           Ajout de la procédure CopyDotInModele
                           15/11/2002 suppression de la procédure
PT3   : 26/09/2001 V562 SB ReConception module absence
                          Modif. Multicritère
                          Suppression ligne de menu 42254 Gestion absence par
                          salarie
PT4   : 08/10/2001 V562 SB Mise en commentaire de l'appel de la fonction
                          CopyDotInModele..
                          Le copie se fera au niveau du multicritère de
                          selection des salaries voir source UTOFPGSaisieSalRtf
PT5   : 25/10/2001 V562 SB Maj de l'ordreetat des rubriques de cotisations
PT6   : 15/11/2001 V562 SB Ajout d'une ligne de menu : Paramétrage des CP
PT7   : 23/01/2002 V571 SB Ajout point d'entrée pour paramétrage des pays..
PT8   : 14/02/2002 V571 MF Ajout d'une ligne de menu : Groupes Internes
PT9   : 22/02/2002 V571 VG Suppression du calcul DADS-U BTP si SO_PGBTP = False
                           + Gestion de la saisie BTP
PT10  : 27/03/2002 V571 VG Purge des envois DADS-U et prud'hommes
PT11  : 30/04/2002 V582 PH Activation désactivation des lignes de menus
                           correspondant à des options gérées
PT12  : 13/05/2002 V582 MF Ajout menu 41650 : Qualifiant de cotisation
PT13  : 15/05/2002 V582 MF Désactivation menus 42358 et 41840 (DucsEdi)
PT14-1: 21/05/2002 V582 VG Lancement de l'assistant après connexion s'il
                           n'existe pas d'établissement complémentaire
PT14-2: 21/05/2002 V582 VG Version S3
PT15  : 04/06/2002 V582 JL Fiche de bug n°10019 : Si pas d'éxercice social alors ouverture en mode création
PT16  : 11/06/2002 V582 PH Gestion des controles avec saisie des primes
PT17  : 13/06/2002 V582 MF Ajout menu 41841,41843,41845 Envois Ducs EDI
PT18  : 13/08/2002 V582 JL Ajout Menu 47 pour formation
PT19  : 15/11/2002 V591 SB Ajout Traitement par lot pour état chainés
PT20  : 11/12/2002 V591 SB FQ 10326 Ajout Menu Calendrier & jours feriés
PT21  : 06/01/2003 V591 PH Suppression des accès à la compta ==> pgimajver
PT22  : 30/01/2003 V912 JL FQ 10444 Suppression tablette activité MS du menu si paramsoc MSA non coché
PT23  : 10/04/2003 V_42 MF Ajout Menu 42359 : Ducs récapitulative
PT24  : 05/05/2003 V_42 MF Ajout Menu 41487 : Tablette Codes distribution tickets restau./
PT25  : 05/06/2003 V_42 MF Ajout Menus 46961 à 46966 : Tickets restaurant
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT26  : 17/06/2003 V_421 PH initialisation Préférence et gestion des favoris
PT27  : 21/07/2003 V_421 MF Modification appel menu 41260 (ORGANISMES)
PT28  : 04/08/2003 V_421 MF Menus "titres restaurant" déplacés
PT29  : 29/04/2004 V_50  Mf Ajout Menus "Evolution masse salariale et effectifs"
PT30  : 09/06/2004 V_50  MF Ajout Menu "Traitement des IJSS"
PT31  : 20/07/2004 V_50  MF Correction : affichages des modules "bilan social" &
                            "Gestion des compétnece" qd sérialisés.
PT32  : 29/07/2004 V_50  MF Complément Menu "Traitement des IJSS & maintien"
PT33  : 07/09/2004 V_50  MF Complément Menu "Traitement des IJSS & maintien"
}

{======= PENSER A
Rajouter le menu 43 conditionné au code seria

}
unit MenuDispEPAIES5;

interface
uses HEnt1,
  Forms,
  sysutils,
  HMsgBox,
  Classes,
  Controls,
  HPanel,
  UIUtil,
  ImgList,
  Windows,
  Hdebug,
  UTreeLinks,
  UTOB,
{$IFNDEF CCS3}
{$IFNDEF EAGLCLIENT}
{$IFNDEF EABSENCES}
  UNetListe,
{$ENDIF}
{$ENDIF}
{$ENDIF}
  ParamSoc, AGLUtilOLE, Hctrls, Ent1, EntPGI
{$IFDEF EAGLCLIENT}
  , MenuOLX, M3FP, UtileAGL, MaineAgl, eTablette
{$ELSE}
  , {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, EdtEtat, AGLInit, Tablette, AssistSOPaie, hplanning, ImportISIS
  , UTomPaie, MenuOLG, PgiExec, FichComm, UtilSoc,
  galoutil, galenv,
  Doc_parser, UEdtComp
{$IFNDEF SEUL}
  , EtbMce, CODEAFB
{$ENDIF}
{$ENDIF}

{$IFDEF ETEMPS}
  , eTempsObject, eTempsMenu { Source de saisie des temps }
{$ENDIF}

{$IFDEF MEMCHECK}
  , memcheck
{$ENDIF}
  , YNewMessage // pour ShowNewMessage
  , AglMail // pour AglMailForm
  ;

procedure InitApplication;
procedure RechargeMenuPopPaie;
function PgAnalConnect: boolean;
// PT16  : 11/06/2002 V582 PH Gestion des controles avec saisie des primes
function PgAnalConnectVar: boolean;

procedure DispatchTraitementLot; //PT19

{$IFDEF EAGLCLIENT}
// Nothing !
{$ELSE}
function MyAglMailForm(var Sujet, A, CC: string; Body: TStringList; var Files: string;
  ZoneTechDataVisible: boolean = True): boolean;
{$ENDIF}

type
  TFMenuDisp = class(TDatamodule)
    procedure FMenuDispCreate(Sender: TObject);
  public
    SeriaMessOk: Boolean;
  end;

var
  FMenuDisp: TFMenuDisp;

implementation
uses
  uMultiDossier,
  P5Def,
  P5Util,
  EntPaie,
  PgOutils,
  PGEdtEtat,
  GrpFavoris
{$IFDEF PAIEGED}
  , Utilged
  , UtilDossier
  , YNewDocument
  , YDocuments_Tom
{$ENDIF}
  , UGedFiles
  , HDTLINKS
{$IFNDEF EAGLCLIENT}
  , CreerSoc, CopySoc, Reseau, Mullog, InitSoc, UtilPgi, UtomUtilisat
  //  , MajHalley
  , uBob
{$ENDIF}
{$IFDEF PAIEOLE}
  , OlePaiePgi
{$ENDIF}
{$IFDEF WEBSERVICES}
  , uMajServices
{$ENDIF}
{$IFDEF TESTSIC}
{$IFDEF EAGLCLIENT}
  , CegidIEU
{$ENDIF EAGLCLIENT}
{$ENDIF TESTSIC}
  ;
//PT26  : 17/06/2003 V_421 PH gestion des favoris

function PgChargeFavoris: boolean;
begin
  AddGroupFavoris(FMenuG, ''); // 2eme parama lesTagsToRemove
  result := true;
end;

procedure MyAfterImport(Sender: TObject; FileID: Integer; var Cancel: Boolean);
// Evenement après scannérisation : retourne le FileID du fichier scanné
{$IFDEF PAIEGED}
var
  LastError: string;
  ParGed: TParamGedDoc;
  TobDoc, TobDocGed: TOB;
  DocId: Integer;
{$ENDIF}
begin
  if FileId = -1 then exit;
{$IFDEF PAIEGED}
  // Propose le rangement dans la GED
  ParGed.DocId := 0;
  ParGed.NoDossier := V_PGI.NoDossier;
  ParGed.CodeGed := '';
  // FileId est le n° de fichier obtenu par la GED suite à l'insertion
  ParGed.FileId := FileId;

  // Description par défaut du document à archiver...
  if Sender is TForm then ParGed.DefName := TForm(Sender).Caption;

  Application.BringToFront;
  if ShowNewDocument(ParGed) = '##NONE##' then
    // Fichier refusé, suppression dans la GED
    V_GedFiles.Erase(FileId);
{$ENDIF}
end;

{$R *.DFM}
// Procedure qui indique que le dossier en mode nomade . Uniquement PCL !!!

procedure AfficheMessPaie;
begin
  PGIBox('Attention, le dossier est parti en nomade.#13#10 Vous ne pouvez pas accèder à la comptablilité.',
    'Gestion du Dossier');
end;

procedure Salarie_ressource;
var
  Tob_Sal: TOB;
  St: string;
  Rep: Integer;
  Q: TQuery;
begin
  Rep := PGIAsk('Vous allez exporter vos salariés pour gérer vos ressources', 'Exportation salariés');
  if Rep <> mrYes then exit;
  st := 'SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_LIBELLEEMPLOI FROM SALARIES';
  Tob_Sal := tob.create('Mon export', nil, -1);
  if St <> '' then
  begin
    Q := OpensQl(st, True);
    Tob_Sal.loaddetaildb('SALARIE', '', '', Q, false);
  end;
  if TOB_Sal.detail.count >= 1 then TOB_Sal.SaveToFile('C:\PGI00\STD\EXPORTSR.TXT', False, True, True, '');
  if st <> '' then Ferme(Q);
  FreeAndNil(Tob_Sal);
  PGIBOX('Traitement terminé', 'Exportation salariés');
end;

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
var
  Q: TQuery;
  Nbre, NbreI: Integer;
  st, LesModulPaie: string;
begin
  case Num of
    // PT-1 : 06/08/2001 : V547 : PH Rajout clause identifictaion eAGL pour supprimer accès au script
    10:
      begin
        PgModifNomChamp;
       // Fonction analyse de la connection eAGL pour savoir si les descriptifs outlook des salaries sont corrects
{$IFDEF EABSENCES}
        if PgAnalConnect = FALSE then sortiehalley := TRUE; // Attention, non autorisation connection à la base eAGL
{$ENDIF}
{$IFDEF EPRIMES}
        if PgAnalConnectVar = FALSE then sortiehalley := TRUE; // Attention, non autorisation connection à la base eAGL
{$ENDIF}

        if V_PGI.ModePcl = '1' then
        begin
          V_PGI.QRPdf := True;
          // (Le container de la GED est la base commune)
          InitializeGedFiles(V_PGI.DeFaultSectionName, MyAfterImport);
        end;

        FMEnuG.ChoixModule; // Pour afficher la liste des modules
        DispatchTraitementLot; //PT19
{$IFDEF EAGLCLIENT}
{$ELSE}
        if FMenuDisp.SeriaMessOK then OnAglMailForm := MyAglMailForm;
{$ENDIF}
      end;
    11: VideLesTOBPaie(FALSE); //Après déconnection
    12:
      begin // avant connexion
        if FileExists('CEGID.par') then DroitCegPaie := TRUE
        else DroitCegPaie := FALSE;
        //  Modif Pour tester la seria PCL , pour avoir les variables V_PGI... initilalisees on fait le test aprs la connexion et pas dans init de l'appli

{$IFDEF CCS3}
        FMenuG.SetSeria('00016011', ['00071060'], ['Paie']); // Paie entreprise
{$ELSE}
        if (V_PGI.ModePcl = '1') then
          FMenuG.SetSeria('00280011', ['00250060', '00033060', '00092060', '00042060', '05051060', '00247060', '00248060', '00127060'],
            ['Paie et GRH', 'Analyses', 'Gestion de la formation', 'Saisie déportée', 'DADS Bilatérale', 'Bilan social', 'Gestion des compétences', 'Messagerie/Agenda']) //PCL
        else
          FMenuG.SetSeria('05970012', ['05050060', '05060060', '00092060', '05061060', '05051060', '00247060', '00248060'],
            ['Paie et GRH', 'Analyses', 'Gestion de la formation', 'Saisie déportée', 'DADS Bilatérale', 'Bilan social', 'Gestion des compétences']); // Paie entreprise
{$ENDIF}
        // Interdiction de lancer en direct
        if (not V_PGI.RunFromLanceur) and (V_PGI.ModePCL = '1') then
        begin
          FMenuG.FermeSoc;
          FMenuG.Quitter;
          exit;
        end;
        // A rajouter ici tous les tests sur modepcl ou bien voir si tjrs possible dans InitLaVariablePGI car modif AGL
{$IFNDEF CCS3}
        if (V_PGI.ModePcl = '1') then TitreHalley := 'PAIE-GRH PGI Expert'
        else TitreHalley := 'PAIE-GRH S5';
{$ELSE}
        TitreHalley := 'PAIE S3';
{$ENDIF}
      end;
    13:
    16:
      begin
        NotSalAss := False;
        SalAdm := 'X';
        SalVal := 'X';
      end;
    {------------------------------------------------------------
                                STANDARDS
    -------------------------------------------------------------}
    100: // action A Faire quand on passe par le lanceur
      begin
      end;

    {------------------------------------------------------------
                         MODULE FORMATION MENU 47                     //PT18
    ------------------------------------------------------------}
    43000..43999 : Paie_Menus (Num, PRien);
    47000..47999 : Paie_Menus (Num, PRien);

    {--------------------------------------------------------------------------------
       Fonctionnalités de la gestion eAgl des éléments variables
     -----------------------------------------------------------------------------------}
    40110: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P;X'); // Saisie des primes
    40111: AglLanceFiche('PAY', 'A', '', '', ''); // Saisie des primes
    40114: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P'); // Saisie des primes
    40120: AglLanceFiche('PAY', 'PGMULHISTOSAL', '', '', 'P'); // ANAL_HISTPRIMconsultation historique des primes

  else HShowMessage('2;?caption?;' + TraduireMemoire('Fonction non disponible : ') + ';W;O;O;O;', TitreHalley, IntToStr(Num));
  end;
end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT: string; Range: string);
begin
  case Num of
    1: ;
    900: ParamRegroupementMultiDossier(Lequel);
    999: ParamTable(TT, taCreat, 0, nil); {06/08/2003 SB FQ 10058}
    40135: AglLanceFiche('PAY', 'PROFIL_MUL', '', '', '');
    40190: ParamTable('PGTYPEORGANISME', taCreat, 0, nil);
    40476: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'COE');
    40477: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'IND');
    40478: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'NIV');
    40479: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'QUA');
    40480: ParamTable('PGLIBEMPLOI', taCreat, 0, nil);
    42300: ParamTable('PGQUALDECLARANT', taCreat, 0, nil);
    // Types de voies
    85951: ParamTable('YYVOIENOCOMPL', taCreat, 0, nil, 3);
    85952: ParamTable('YYVOIETYPE', taCreat, 0, nil, 3);

  end;
end;

procedure ZoomEdt(Qui: integer; Quoi: string);
var
  predefini, Dossier, param, nat: string;
  j: integer;
  Ok: Boolean;
begin
  Ok := False;
  Param := Quoi;
  if (Param = '') then exit;
  j := Pos(';', Param);
  if j > 0 then Ok := True;
  if Ok = True then
  begin
    Predefini := ReadTokenSt(Param);
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    if Predefini = 'DOS' then Dossier := PgRendNoDossier()
    else dossier := '000000';
    Quoi := ReadTokenSt(Param);
    Nat := Param;
    if (Nat = 'COT') or (Nat = 'BAS') then Qui := 470;
    if (Nat = 'AAA') then Qui := 471;
  end;

  case Qui of
    440: AglLanceFiche('PAY', 'VARIABLE', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
    450: AglLanceFiche('PAY', 'SALARIE', '', Quoi, 'ACTION=CONSULTATION');
    460: AglLanceFiche('PAY', 'CUMUL', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
    470: AglLanceFiche('PAY', 'COTISATION', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
    471: AglLanceFiche('PAY', 'REMUNERATION', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
    480: AglLanceFiche('PAY', 'PROFIL', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
    490: AglLanceFiche('PAY', 'ORGANISME', '', Quoi, 'ACTION=CONSULTATION');

  end;
  Screen.Cursor := crDefault;
end;

procedure ZoomEdtEtat(Quoi: string);
var
  i, Qui: integer;
begin
  i := Pos(';', Quoi);
  Qui := StrToInt(Copy(Quoi, 1, i - 1));
  Quoi := Copy(Quoi, i + 1, Length(Quoi) - i);
  ZoomEdt(Qui, Quoi);
end;

procedure ChangeMenu41;
var
  i: Integer;
  libelle: string;
begin
  FMenuG.RenameItem(41701, VH_Paie.PgLibCombo1); //Tablette Libre
  FMenuG.RenameItem(41702, VH_Paie.PgLibCombo2);
  FMenuG.RenameItem(41703, VH_Paie.PgLibCombo3);
  FMenuG.RenameItem(41704, VH_Paie.PgLibCombo4);
  FMenuG.RenameItem(41471, VH_Paie.PGLibelleOrgStat1); //Organisation
  FMenuG.RenameItem(41472, VH_Paie.PGLibelleOrgStat2);
  FMenuG.RenameItem(41473, VH_Paie.PGLibelleOrgStat3);
  FMenuG.RenameItem(41474, VH_Paie.PGLibelleOrgStat4);
  FMenuG.RenameItem(41480, VH_Paie.PGLibCodeStat);
  if VH_Paie.PGLibCodeStat = '' then FMenuG.RemoveItem(41480); //Code Stat
  // PT28 if VH_Paie.PGTicketRestau = False then FMenuG.RemoveItem (41487); //PT24

  for i := 1 to 4 do
  begin
    Libelle := '';
    if i = 1 then libelle := VH_Paie.PGLibelleOrgStat1;
    if i = 2 then libelle := VH_Paie.PGLibelleOrgStat2;
    if i = 3 then libelle := VH_Paie.PGLibelleOrgStat3;
    if i = 4 then libelle := VH_Paie.PGLibelleOrgStat4;
    if (Libelle = '') then FMenuG.RemoveItem(i + 41470);
  end;
  if VH_Paie.PgNbCombo < 5 then
    for i := VH_Paie.PgNbCombo + 1 to 4 do
    begin
      FMenuG.RemoveItem(i + 41700);
    end;
  //ChargeMenuPop(hm6,FMenuG.Dispatch) ;
end;


// d PT25  Tickets restaurant    PT28

procedure ChangeMenu42;
begin
  if (VH_Paie.PGCodeClient <> VH_Paie.PGCodeRattach) then
  begin
    {confection du fichier final uniquement si client principal}
    FMenuG.RemoveItem(42456);
    {Analyse du récapitulatif uniquement si client principal}
    FMenuG.RemoveItem(42462);
  end;
  if (VH_Paie.PGTicketRestau = False) then
    FMenuG.RemoveGroup(-42450, TRUE);
  {Edition en en prévisions pour V6 caché pour le moment}
  FMenuG.RemoveItem(42158);
  // d PT33
  if VH_Paie.PGGestIJSS = False then
  begin
    FMenuG.RemoveGroup(-42800, TRUE);
  end;
  if VH_Paie.PGMaintien = False then
  begin
    FMenuG.RemoveItem(-42820);
  end;
  // f PT33

end;
// f PT25

procedure ChangeMenu47; //PT18 Affichage des libellés des tablettes
var
  i: Integer;
  libelle: string;
begin
  FMenuG.RenameItem(47111, VH_Paie.FormationLibre1);
  FMenuG.RenameItem(47112, VH_Paie.FormationLibre2);
  FMenuG.RenameItem(47113, VH_Paie.FormationLibre3);
  FMenuG.RenameItem(47114, VH_Paie.FormationLibre4);
  FMenuG.RenameItem(47115, VH_Paie.FormationLibre5);
  FMenuG.RenameItem(47116, VH_Paie.FormationLibre6);
  FMenuG.RenameItem(47117, VH_Paie.FormationLibre7);
  FMenuG.RenameItem(47118, VH_Paie.FormationLibre8);
  for i := 1 to 8 do
  begin
    if i = 1 then Libelle := VH_Paie.FormationLibre1;
    if i = 2 then Libelle := VH_Paie.FormationLibre2;
    if i = 3 then Libelle := VH_Paie.FormationLibre3;
    if i = 4 then Libelle := VH_Paie.FormationLibre4;
    if i = 5 then Libelle := VH_Paie.FormationLibre5;
    if i = 6 then Libelle := VH_Paie.FormationLibre6;
    if i = 7 then Libelle := VH_Paie.FormationLibre7;
    if i = 8 then Libelle := VH_Paie.FormationLibre8;
    if Libelle = '' then
    begin
      FMenuG.RemoveItem(i + 47110);
    end;
  end;
  if VH_Paie.PGForPrevisionnel = False then
  begin
    FMenuG.RemoveGroup(-47200, TRUE); //Prévisionnel
    FMenuG.RemoveGroup(-47570, TRUE); //Editions comparatif
    FMenuG.RemoveItem(47620); //Analyse
    FMenuG.RemoveGroup(-47155, TRUE); //Gestion des cursus
    FMenuG.RemoveItem(47162); //Salaires animateurs
    FMenuG.RemoveItem(47131); //Forfait frais
    FMenuG.RemoveItem(47913); //Maj coûts
    FMenuG.RemoveItem(47923); //Reaffectation
    FMenuG.RemoveItem(47930); //Récupération
  end;
  if VH_Paie.PGForValidSession = False then FMenuG.RemoveItem(47332);
  if VH_Paie.PGForGestionOPCA = False then FMenuG.RemoveGroup(-47800, TRUE);
  if VH_Paie.PGForValidPrev = False then
  begin
    FMenuG.RemoveItem(47220);
    FMenuG.RemoveItem(47250);
  end;
  //SCORING
  FMenuG.RenameItem(47713, RechDom('PGZONELIBRE', 'SC1', False));
  FMenuG.RenameItem(47714, RechDom('PGZONELIBRE', 'SC2', False));
  FMenuG.RenameItem(47715, RechDom('PGZONELIBRE', 'SC3', False));
  FMenuG.RenameItem(47716, RechDom('PGZONELIBRE', 'SC4', False));
  FMenuG.RenameItem(47717, RechDom('PGZONELIBRE', 'SC5', False));
  FMenuG.RemoveItem(47553);
  FMenuG.RemoveItem(47582);
  FMenuG.RemoveItem(47525);
  FMenuG.RemoveItem(47742);
  FMenuG.RemoveItem(47743);
  FMenuG.RemoveItem(47573);
  FMenuG.RemoveItem(47571);
  FMenuG.RemoveItem(47255);
end;

procedure ChangeMenu48;
begin
  FMenuG.RenameItem(48191, RechDom('GCZONELIBRE', 'RH1', False));
  FMenuG.RenameItem(48192, RechDom('GCZONELIBRE', 'RH2', False));
  FMenuG.RenameItem(48193, RechDom('GCZONELIBRE', 'RH3', False));
  FMenuG.RenameItem(48194, RechDom('GCZONELIBRE', 'RH4', False));
  FMenuG.RenameItem(48195, RechDom('GCZONELIBRE', 'RH5', False));
  FMenuG.RenameItem(48196, RechDom('GCZONELIBRE', 'CP1', False));
  FMenuG.RenameItem(48197, RechDom('GCZONELIBRE', 'CP2', False));
  FMenuG.RenameItem(48198, RechDom('GCZONELIBRE', 'CP3', False));
  FMenuG.RemoveItem(48136);
  FMenuG.RemoveItem(48137);
end;

procedure RechargeMenuPopPaie;
begin
  //
end;
//

procedure AfterChangeModule(NumModule: Integer);
var
  CEG, STD, DOS: Boolean;
begin
{$IFDEF ENCOURS}
  VH_Paie.PgeAbsences := TRUE;
{$ENDIF}

  // Nécessaire car les menus sont chargés à chaque fois... !!!! ????
{$IFDEF eTemps}
  ET_Menu(-1);
{$ENDIF}
{$IFDEF EABSENCES}
  if SalAdm = '-' then FMenuG.RemoveGroup(-43800, TRUE); //  désactiver menu administrateur
  if SalVal = '-' then FMenuG.RemoveGroup(43700, TRUE); //  désactiver menu validation absences salarie
  if NotSalAss then FMenuG.RemoveItem(43165);
{$ENDIF}


  // A VOIR pour rajouter les suppressions de lignes de menus spéciques eAgl
  if VH_Paie.PGEcabMonoBase then
  begin //si applicatif monobase, desactive menu import export
    FMenuG.RemoveItem(43810); //  Import des données via ECAB
    FMenuG.RemoveItem(43820); //  Export des absences via ECAB
    FMenuG.RemoveItem(49950); //  Export des données via la paie
  end;
  if not VH_Paie.PgLienRessource then FMenuG.RemoveItem(42170);
{$IFNDEF CCS3}
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';
  case NumModule of
    47: Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CFORMS5.HLP';
    48: Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CRHCAR5.HLP';
  end;
{$ENDIF}

  case NumModule of
    41: ChangeMenu41;
    42: ChangeMenu42; // PT25 Tickets restaurant PT28
    47: ChangeMenu47; // PT18 Affichage des libellés des tablettes
    48: ChangeMenu48;
  end;
  /// Menu Pop général
  case NumModule of
    41, 42: ChargeMenuPop(6, FMenuG.DispatchX);
  else ChargeMenuPop(NumModule, FMenuG.DispatchX);
  end;
{$IFDEF ENCOURS}
  VH_Paie.PgSeriaAnalyses := True;
  VH_Paie.PgSeriaDADSB := True;
{$ENDIF}
  if VH_Paie.PGPCS2003 then FMenuG.RemoveItem(-42180);

  if not VH_Paie.PgSeriaAnalyses then // Module Analyse non sérialisé
  begin
    FMenuG.RemoveItem(46710);
    FMenuG.RemoveItem(46720);
    FMenuG.RemoveGroup(46700, TRUE);
    FMenuG.RemoveItem(46810);
    FMenuG.RemoveItem(46820);
    FMenuG.RemoveGroup(46800, TRUE);
  end;

  if not VH_Paie.PgSeriaDADSB then // Module DADS Bilatérale non sérialisés
  begin
    FMenuG.RemoveItem(-41816);
    FMenuG.RemoveGroup(-42560, TRUE);
  end;

  AccesPredefini('CEG', CEG, STD, DOS);
  if not CEG then FMenuG.RemoveItem(49210) // désactiver serialisation car dans barre outils
  else FMenuG.RenameItem(49210, 'Utilitaire de contrôle');
{$IFDEF EAGLCLIENT}
  FMenuG.RemoveGroup(49100, TRUE);
  FMenuG.RemoveItem(49220);
  FMenuG.RemoveItem(49230);
  FMenuG.RemoveItem(49240);
  FMenuG.RemoveItem(49250);
  FMenuG.RemoveItem(49260);
  FMenuG.RemoveItem(42910);
  FMenuG.RemoveItem(42336); //Calendrier bimensuel inopérant CWAS
{$ENDIF}
  //  FMenuG.RenameItem(49810, 'S1 et NetService');
    //FMenuG.RemoveItem(43730);   //  Planning
  FMenuG.RemoveItem(49980); //  Planning 41835
  FMenuG.RemoveItem(41835); //  Planning 41835
  //ANCIEN MENU ETEMPS
  FMenuG.RemoveItem(43210); // Gestion Etemps collaborateur
  FMenuG.RemoveGroup(-43240); // Gestion Etemps responsable
  FMenuG.RemoveItem(-43240); // Gestion Etemps responsable



{$IFNDEF ENCOURS}
  //Spécificité CEGID SA
  if not GetParamSocSecur('SO_IFDEFCEGID',False) then   { 13/10/2005 norme AGL }
  begin
    FMenuG.RemoveItem(42257); //  Calendrier excel
  end;
  FMenuG.RemoveItem(42390); //  Calendrier annuel
  FMenuG.RemoveItem(42395); //  Point d'entrée
  FMenuG.RemoveItem(42255); //  Planning des absences
  FMenuG.RemoveItem(49852); //  Annulation clôture CP
  FMenuG.RemoveItem(49853); //  Génération reliquat
  if VH_Paie.PGMODIFCOEFF = False then FMenuG.RemoveGroup(-46400, TRUE); // Evolution de modif
  if not VH_Paie.PgeAbsences then FMenuG.RemoveGroup(-46600, TRUE); // organisation
  if not VH_Paie.PGSaisieArret then FMenuG.RemoveGroup(-42700, TRUE); // saisie arret
  FMenuG.RemoveItem(42153); //  Congés spectacles
  // Accès compatibilité supprimé
  //FMenuG.RemoveItem(41710);
  FMenuG.RemoveItem(41715);
  FMenuG.RemoveItem(41720);
  FMenuG.RemoveItem(41730);

  FMenuG.RemoveItem(41491); //  Motif Entree
  FMenuG.RemoveItem(41492); //  Motif sortie
  //FMenuG.RemoveItem(42255);   //  Planning des absences
  //FMenuG.RemoveItem(-41840);   //  PT14  (ducsEdi)
  //FMenuG.RemoveItem(42358);   //  PT14 (ducsEdi)

  if VH_Paie.PGBTP = False then
  begin
    FMenuG.RemoveItem(42140); //  Compléments BTP salarié
    FMenuG.RemoveItem(42512); //  Calcul dads-u CP BTP
    FMenuG.RemoveItem(42532); //  Edition DADS-U CP BTP
  end;

{$IFNDEF VG}
  FMenuG.RemoveItem(42524); //  Multi-critère Lexique DADS-U
{$ELSE}
  FMenuG.RenameItem(42524, 'Lexique DADS-U');
{$ENDIF}

  if VH_Paie.PGMODIFCOEFF = False then
    FMenuG.RemoveItem(-46610); // Evolution du coefficient

  //PT14-2
{$IFDEF CCS3}
  FMenuG.RemoveItem(41485); //  Activité MSA OK
  FMenuG.RemoveItem(41197); //  Contrat intérimaire OK
  FMenuG.RemoveItem(41270); //  Banques des organismes
  FMenuG.RemoveItem(-41840); // Ducs-EDI OK
  FMenuG.RemoveItem(-41860); // Congés spectacles
  //FMenuG.RemoveItem(42140);// Compléments BTP salarié OK
  FMenuG.RemoveItem(-42150); // Intermittents OK
  FMenuG.RemoveItem(-42160); // MSA OK
  FMenuG.RemoveItem(42217); //  Création bulletin complémentaire OK
  FMenuG.RemoveItem(-42220); //  IJSS & maintien du salaire OK
  FMenuG.RemoveItem(42255); //  Supervision OK
  FMenuG.RemoveItem(42351); //  Déclaration MO OK
  FMenuG.RemoveItem(42358); //  Ducs-EDI OK
  FMenuG.RemoveItem(42380); //  Etat libre OK
  FMenuG.RemoveItem(42390); //  Calendrier annuel OK
  FMenuG.RemoveItem(42394); //  Etats chaînés OK
  FMenuG.RemoveItem(42395); //  Edition en masse OK
  FMenuG.RemoveGroup(-42450, TRUE); //  Titres restaurant
  //FMenuG.RemoveItem(42512);         //  Calcul dads-u CP BTP OK
  //FMenuG.RemoveItem(42532);         //  Edition DADS-U CP BTP
  FMenuG.RemoveGroup(-42700, TRUE); // saisie arret
  FMenuG.RemoveGroup(46100, TRUE); //  Gestion des contrats OK
  FMenuG.RemoveGroup(-46200, TRUE); //  Travailleurs handicapés OK
  FMenuG.RemoveGroup(-46300, TRUE); //  Médecine du travail OK
  FMenuG.RemoveGroup(-46500, TRUE); //  Intérimaires et stagiaires OK
  FMenuG.RemoveGroup(-46600, TRUE); //  Organisation OK
  FMenuG.RemoveGroup(46800, TRUE); //  Synthèses OK
  FMenuG.RemoveItem(49810); //  NetService OK
  //FMenuG.RemoveItem(49860);         //  Gestion des standards
  FMenuG.RemoveItem(-49880); //  Basculement en Euro OK
  FMenuG.RemoveGroup(49900, TRUE); //  ePaie OK
{$ELSE}
  FMenuG.RemoveItem(-42385); //  Analyse et Synthèse OK
{$ENDIF}
  FMenuG.RenameItem(49240, 'Utilisateurs &connectés'); // Faute orthographe
  //FIN PT14-2
{$IFDEF EAGLCLIENT}
  FMenuG.RemoveItem(41420); //  Annuaire PCL non porté CWAS
{$ENDIF}
{$IFDEF ENCOURS}
  VH_Paie.PgeAbsences := TRUE; // @@@@@@@@@@@@@@@@@@@@@@@@@ ne passe jamais!!
{$ENDIF}
  // Suppression des lignes de menu concernant le module recrutement
  FMenuG.RemoveItem(-48270); // Gestion des entretiens
  FMenuG.RemoveGroup(-48500, TRUE); // Recrutement
  FMenuG.RemoveGroup(-46400, TRUE); // Participation
  FMenuG.RemoveItem(-42280); // Frais professionnel

  // Suppression des lignes de menu concernant la TOX
  FMenuG.RemoveGroup(49300, TRUE);
  FMenuG.RemoveItem(49865); //  Web Services en devenir en case 30ou 31
  // if NOT V_PGI.DebugSQL then FMenuG.RemoveItem  (49861);  // Transfert dossier
  if not V_PGI.DebugSQL then FMenuG.RemoveGroup(49400); // menu hiérarchie
{$IFDEF EAGLCLIENT}
  FMenuG.RemoveItem(42394); // En CWAS, non applicable
{$ENDIF}
  if not VH_Paie.PgeAbsences then
    FMenuG.RemoveGroup(49900, TRUE) // EPaie si non séria
  else
    if VH_Paie.PGEcabMonoBase then
  begin //si applicatif monobase, desactive menu import export
    FMenuG.RemoveItem(43810); //  Import des données via ECAB
    FMenuG.RemoveItem(43820); //  Export des absences via ECAB
    FMenuG.RemoveItem(49950); //  Export des données via la paie
  end;
  //   PT11 30/04/2002 V582 PH Activation désactivation des lignes de menus aux différents cas gérés
  // Gestion des intérimaires
  if not VH_Paie.PGInterimaires then
    FMenuG.RemoveGroup(-46500, TRUE);

  // Gestion MSA
  if not VH_Paie.PGMsa then
  begin
    FMenuG.RemoveItem(-42160);
    FMenuG.RemoveItem(41485); // PT22
  end;

  // Gestion des intermittants du spectacle
  if not VH_Paie.PGIntermittents then
    FMenuG.RemoveItem(-42150);

  // Gestion des Services
{$IFDEF SERVICE}
  VH_Paie.PGRESPONSABLES := True;
{$ENDIF}
  if not VH_Paie.PGRESPONSABLES then
    FMenuG.RemoveGroup(-46600, TRUE);
  // FIN PT11
{$ENDIF}
  //FORMATION
  FMenuG.RemoveItem(47170);
  FMenuG.RemoveItem(-47190);
  {futur menu V6}
  FMenuG.RemoveItem(47255);
  FMenuG.RemoveItem(47913);
  FMenuG.RemoveItem(47923);
  FMenuG.RemoveItem(47157);

  //SB 15/09/03 FQ 10781 Affichage module CP si gestion CP Paramètre Soc
  if not VH_PAie.PGCongesPayes then
  begin
    FMenuG.RemoveItem(-42130); //Module Gestion par salarié
    FMenuG.RemoveItem(42334); //Provision CP
    FMenuG.RemoveItem(42335); //Solde CP
  end;
  //FMenuG.RemoveItem (49413); // tablettes hierarchiques
  //FMenuG.RemoveItem (49414); // tablettes hierarchiques

  if (V_PGI.Driver = dbORACLE7) then FMenuG.RemoveItem(42337);
  if VH_Paie.PGPCS2003 then FMenuG.RemoveItem(-41170);
  // d PT30 - IJSS
  if VH_Paie.PGGestIJSS = False then
  begin
    FMenuG.RemoveGroup(-42800, TRUE); // PT33
  end;
  // f PT30 - IJSS
  // d PT32 - Maintien
  if VH_Paie.PGMaintien = False then
  begin
    FMenuG.RemoveItem(-42820); // PT33
  end;
  // f PT32 - Maintien

  // Menu multi dossier n'existe pas en PCL du moins dans la V6
  if V_PGI.ModePCL = '1' then FMenuG.RemoveItem(-46800);
end;

procedure AfterProtec(sAcces: string); //Sérialisation des modules
begin
  // cas où on coche "version non sérialisée' dans
  // l'écran de séria. On force la non sérialisation

  if V_PGI.NoProtec then V_PGI.VersionDemo := TRUE
  else V_PGI.VersionDemo := False;
  if (sAcces[1] = '-') then V_PGI.VersionDemo := TRUE;
  VH_Paie.PgSeriaPaie := (sAcces[1] = 'X');
  VH_Paie.PgSeriaAnalyses := (sAcces[2] = 'X');
  VH_Paie.PgSeriaFormation := (sAcces[3] = 'X');
  VH_Paie.PgeAbsences := (sAcces[4] = 'X');
  VH_Paie.PgSeriaDADSB := (sAcces[5] = 'X');
  VH_Paie.PgSeriaBilan := (sAcces[6] = 'X'); // PT31
  VH_Paie.PgSeriaCompetence := (sAcces[7] = 'X'); // PT31
  if V_PGI.ModePCL = '1' then
  begin // Agenda PCL
    FMenuDisp.SeriaMessOK := False;
    if (sAcces[8] = 'X') then FMenuDisp.SeriaMessOK := TRUE;
  end;
{$IFDEF ENCOURS}
  VH_Paie.PgeAbsences := TRUE;
  VH_Paie.PgSeriaFormation := TRUE;
  VH_Paie.PgSeriaAnalyses := TRUE;
  VH_Paie.PgSeriaDADSB := TRUE;
  VH_Paie.PgSeriaCompetence := TRUE;
  VH_Paie.PgSeriaBilan := TRUE;
{$ENDIF}
  //PT14-2
{$IFNDEF CCS3}
{$IFNDEF EABSENCES}
{$IFNDEF EPRIMES}
{$IFNDEF EFORMATION}
  ChargeModules_PAIE;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
  //FIN PT14-2

{$IFDEF ETEMPS}
  V_PGI.VersionDemo := False;
{$ENDIF}

end;

procedure AGLAfterChangeModule(Parms: array of variant; nb: integer);
begin
  AfterChangeModule(Integer(Parms[0]));
  ProcCalcEdt := CalcOLEEtatPG;
end;
{ Fonction qui recupere le numero de Salarie dans le cas d'une connection NT avec
  de façon à savoir, s'il est adminstrateur et/ou responsable validation absences
}

function PgAnalConnect: boolean;
var
  Q: TQuery;
  LaChaine, LeResp: string;
begin
  LaChaine := string(GetParamSoc('SO_PGCHEMINEAGL'));
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
      VH_Paie.PGCheminEagl := Q.FindField('SOC_DATA').AsString;
    Ferme(Q);
  end;
  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if LeSalarie = '' then
  begin
    PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', 'Saisie des absences');
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    Q := OpenSql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
    begin
      SalAdm := Q.FindField('PSE_ADMINISTABS').AsString;
      SalVal := Q.FindField('PSE_OKVALIDABS').AsString;
      result := true;
    end
    else
    begin
      PgiBox('Vous êtes inconnu comme salarié#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
      Ferme(Q);
      result := false;
      exit;
    end;
    LeResp := Q.FindField('PSE_RESPONSABS').AsString;
    Ferme(Q);
    if LeResp = '' then
    begin
      PgiBox('Vous n''avez pas de responsable affecté !#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
      result := false;
      exit;
    end;
    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
    else
    begin
      PgiBox('Vous êtes inconnu comme salarié de la paie#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
      Ferme(Q);
      result := false;
      exit;
    end;
    Ferme(Q);
  end;
  // Desactivation des fonctionnalités administrateur ou responsable validation selon le salarié
  if SalAdm <> 'X' then FMenuG.RemoveGroup(-43800, TRUE); //  désactiver menu administrateur
  if SalVal <> 'X' then FMenuG.RemoveGroup(43700, TRUE); //  désactiver menu validation absences salarie
  //Salarié secrétaire
  NotSalAss := not (ExisteSql('SELECT PSE_ASSISTABS FROM DEPORTSAL WHERE PSE_ASSISTABS="' + LeSalarie + '"'));
  if NotSalAss then FMenuG.RemoveItem(43165);

  result := true;
  //BaseEPaie := (GetParamSoc ('SO_PGBASEEPAIE'));

{$IFDEF etemps}
  // XP le 24-07-2003
  ET_Menu(-1);
  if ldSaisieTemps in LeTemps.MesDroits then
  begin
    FMenuG.ChangeModule(263);
    ET_Menu(263001);
  end
  else
  begin
    if result and (SalAdm <> 'X') and (SalVal <> 'X') then
      AglLanceFiche('PAY', 'EABSENCE_MUL', '', '', 'SAL');
  end;
{$ELSE}
  if result and (SalAdm <> 'X') and (SalVal <> 'X') then
    AglLanceFiche('PAY', 'EABSENCE_MUL', '', '', 'SAL');
{$ENDIF}
end;
// PT16  : 11/06/2002 VDev PH Gestion des controles avec saisie des primes
{ Fonction qui recupere le numero de Salarie dans le cas d'une connection NT avec
  de façon à savoir, s'il est adminstrateur et/ou responsable validation des primes
}
{***********A.G.L.***********************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 15/11/2002
Modifié le ... : 15/11/2002
Description .. : Traitement par lots
Mots clefs ... : PAIE;LOT
*****************************************************************}

function PgAnalConnectVar: boolean;
var
  Q: TQuery;
  LaChaine, Lib: string;
  BaseEPaie: Boolean;
begin
  Lib := 'Saisie ' + VH_Paie.PGLibSaisPrim;
  LaChaine := string(GetParamSoc('SO_PGCHEMINEAGL'));
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
      VH_Paie.PGCheminEagl := Q.FindField('SOC_DATA').AsString;
    Ferme(Q);
  end;
  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if LeSalarie = '' then
  begin
    PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', Lib);
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    Q := OpenSql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
    begin
      SalAdm := Q.FindField('PSE_ADMINISTABS').AsString; // Administrateur eAgl des absences et des primes
      SalVal := Q.FindField('PSE_OKVALIDVAR').AsString; // Autorisation de valider les primes = saisie des primes
      ConsultP := FALSE;
      // On va regarder si le salarie valide les absences pour autoriser les accès en consultation uniquement
      if SalVal <> 'X' then
      begin
        SalVal := Q.FindField('PSE_OKVALIDABS').AsString;
        ConsultP := TRUE;
      end;
      result := true;
    end
    else
    begin
      PgiBox('Vous êtes inconnu comme salarié#13#10 Renseignez vous au service du personnel', Lib);
      Ferme(Q);
      result := false;
      exit;
    end;
    if (SalAdm <> 'X') and (SalVal <> 'X') then
    begin
      PgiBox('Vous n''êtes pas autorisé à vous connecter, !#13#10Renseignez vous au service du personnel', Lib);
      result := false;
      exit;
    end;
    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
    else
    begin
      PgiBox('Vous êtes inconnu comme salarié de la paie#13#10 Renseignez vous au service du personnel', Lib);
      Ferme(Q);
      result := false;
      exit;
    end;
    Ferme(Q);
  end;
  result := true;
  BaseEPaie := GetParamSoc('SO_PGBASEEPAIE');
  if (not BaseEPaie) then
    if (SalAdm <> 'X') then
    begin
      PgiBox('L''accès à la saisie ' + VH_Paie.PGLibSaisPrim + ' est interdit#13#10 Reconnectez vous ultérieurement SVP', Lib);
      Result := FALSE; // Base non accessible en tant que base de saisie déportée
    end;
end;
// FIN PT16


{DEB PT19 : Traitement par lots}

procedure DispatchTraitementLot;
var
  i: Integer;
  St, Nom, Value: string;
begin
  for i := 1 to ParamCount do
  begin
    St := ParamStr(i);
    Nom := UpperCase(Trim(ReadTokenPipe(St, '=')));
    Value := UpperCase(Trim(St));
    if Nom = '/PGETATSCHAINES' then
    begin
      // Lancement d' une fiche pour mettre en inside, sinon le premier état chainé de s' imprime pas
      Value := Value + ';PG';
      AGLLanceFiche('CP', 'ETATSCHAINES', '', '', Value);
      // Fermeture de l'application
      Application.ProcessMessages;
{$IFNDEF EAGLCLIENT}
      FMenuG.ForceClose := True;
{$ENDIF}
      FMenuG.Close;
      Exit;
    end;
  end;
end;
{FIN}

{$IFDEF EAGLCLIENT}
// Nothing !
{$ELSE}

function MyAglMailForm(var Sujet, A, CC: string; Body: TStringList; var Files: string;
  ZoneTechDataVisible: boolean = True): boolean;
begin
  ;
  // Routage vers la messagerie PGI au lieu du send mail agl
  Result := ShowNewMessage(0, 0, '', A);
end;
{$ENDIF}

procedure InitApplication;
begin
{$IFDEF MEMCHECK}
  MemCheckLogFileName := ChangeFileExt(Application.exename, '.log');
  MemChk;
{$ENDIF}

  ProcCalcEdt := CalcOLEEtatPG;
  ProcZoomEdt := ZoomEdtEtat;

  //ProcGetVH:=GetVS1 ;
  //ProcGetDate:=GetDate ;
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
{$IFNDEF EAGLCLIENT}
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
{$ENDIF}
  // Gestion de la seria Modif faite pour V1.1 du 12-01-2001 pour reconnaitre Seria PCL
  // Deplaceé dans apres connexion pour avoir les variables V_PGI... initialisées
  // Mais initialisation faite quand meme pour avoir la variable sAcces definie pour la fonction AfterProtect
{$IFDEF CCS3}
  FMenuG.SetSeria('00016011', ['00071060'], ['Paie']); // Paie entreprise
{$ELSE}
  if (V_PGI.ModePcl = '1') then
    FMenuG.SetSeria('00280011', ['00250060', '00033060', '00092060', '00042060', '05051060', '00119060', '00247060', '00248060', '00127060'],
      ['Paie et GRH', 'Analyses', 'Gestion de la formation', 'Saisie déportée', 'DADS Bilatérale', 'GED', 'Bilan social', 'Gestion des compétences', 'Messagerie/Agenda']) //PCL
  else
    FMenuG.SetSeria('05970012', ['05050060', '05060060', '00092060', '05061060', '05051060', '00247060', '00248060'],
      ['Paie et GRH', 'Analyses', 'Gestion de la formation', 'Saisie déportée', 'DADS Bilatérale', 'Bilan social', 'Gestion des compétences']); // Paie entreprise
{$ENDIF}

  FMenuG.OnAfterProtec := AfterProtec;
  FMenuG.SetModules([45], [45]);
{$IFNDEF EABSENCES}
  // Menu 40 reservé à la gestion des standards
  // VH_Paie.PgeAbsences := TRUE; // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  //PT14-2
{$IFDEF CCS3}
  FMenuG.SetModules([42, 41, 49], [42, 41, 49]);
{$ELSE}
{$IFNDEF EPRIMES}
{$IFNDEF EFORMATION}
  ChargeModules_PAIE;
  //  if VH_Paie.PgeAbsences then FMenuG.SetModules([42,41,46,49,43,47],[42,41,46,49,43,47])//PT18
  //  else FMenuG.SetModules([42,41,46,47,49],[42,41,46,47,49]) ;  //PT18
{$ELSE}
  FMenuG.SetModules([47], [47]);
{$ENDIF}
{$ELSE}
  FMenuG.SetModules([40], [40]);
{$ENDIF}
{$ENDIF}
  //FIN PT14-2
{$ELSE}
  // Cas eAgl pour la saisie décentralisée des absences, c'est le menu 43 qui lui est reservé afin de ne pas melanger les applications
{$IFDEF etemps}
  FMenuG.SetModules([43, 263], [43, 73]);
{$ELSE}
  FMenuG.SetModules([43], [43]);
{$ENDIF}
{$ENDIF}

  FMenuG.OnChangeModule := AfterChangeModule;
  V_PGI.DispatchTT := DispatchTT;
  //PT26  : 17/06/2003 V_421 PH initialisation Préférence et gestion des favoris
{$IFNDEF EAGLCLIENT}
  FMenuG.SetPreferences(['Divers'], False);
{$ENDIF}
  FMenuG.OnChargeFavoris := PgChargeFavoris;
{$IFDEF TESTSIC}
{$IFDEF EAGLCLIENT}
  SaveSynRegKey('eAGLHost', 'SRV-LYO-TECH9:8096', true);
  FCegidIE.HostN.Enabled := false;
{$ENDIF EAGLCLIENT}
{$ENDIF TESTSIC}
  V_PGI.DispatchTT := DispatchTT;
end;

{ TFMenuDisp }

procedure TFMenuDisp.FMenuDispCreate(Sender: TObject);
begin
  // PgiAppAlone:= TRUE;
  // CreatePGIApp;          supprimer à partir de la version 422 AGL
  SeriaMessOk := False;
end;


procedure InitLaVariablePGI;
begin
  Apalatys := 'CEGID';
  HalSocIni := 'cegidpgi.ini';
  NomHalley := 'CPS5';

{$IFDEF EABSENCES}
  HalSocIni := 'cegidpgi.ini';
  NomHalley := 'eCABS5';
{$ENDIF}
{$IFDEF ABSENCES}
  HalSocIni := 'cegidpgi.ini';
  NomHalley := 'CABS5';
{$ENDIF}
{$IFDEF ETEMPS}
  NomHalley := 'eCABS5';
  HalSocIni := 'IlEstTemps.ini';
{$ENDIF}
  Copyright := '© Copyright ' + Apalatys;
  // Serie S5
  V_PGI.OutLook := TRUE;
  V_PGI.LaSerie := S5;
  if (V_PGI.ModePcl = '1') then TitreHalley := 'PAIE-GRH PGI Expert'
  else TitreHalley := 'PAIE-GRH S5';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';
{$IFDEF EFORMATION}
  TitreHalley := 'FORMATION S5';
  NomHalley := 'eCFORS5';
  HalSocIni := 'eFormation.ini';
{$ENDIF}

{$IFDEF EPRIMES}
  TitreHalley := 'SAISIE DES PRIMES S5';
  NomHalley := 'eCPRIS5';
  HalSocIni := 'TheSaiPrim.ini';
{$ENDIF}
{$IFDEF CCS3}
  // Serie S3
  V_PGI.OutLook := TRUE;
  V_PGI.LaSerie := S3;
  TitreHalley := 'PAIE S3';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS3.HLP';
  NomHalley := 'CPS3';
{$ELSE}
{$IFDEF SERIES7}
  // Serie S7
  V_PGI.OutLook := FALSE;
  V_PGI.LaSerie := S7;
  TitreHalley := 'PAIE-GRH S7';
  NomHalley := 'CPS7';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS7.HLP';
{$ENDIF}
{$ENDIF}
  V_PGI.ToolsBarRight := TRUE;
  if (V_PGI.ModePcl = '1') then ForceCascadeForms;
{$IFDEF EAGLCLIENT}
  V_PGI.AutoSearch := False;
{$ENDIF}
  V_PGI.MultiUserLogin := true;
  V_PGI.CegidUpdateServerParams := 'www.update.cegid.fr';
  ChargeXuelib;
  V_PGI.OfficeMsg := TRUE;
  V_PGI.NoModuleButtons := FALSE;
  V_PGI.NbColModuleButtons := 2;
  V_PGI.VersionDemo := True;
  V_PGI.MenuCourant := 0;
  V_PGI.VersionReseau := True;
  V_PGI.NumVersion := '6.0.0';
  V_PGI.NumBuild := '000.014';
{$IFDEF eTemps}
  V_PGI.NumVersionBase := 620;
{$ELSE}
  V_PGI.NumVersionBase := 656;
{$ENDIF}

  V_PGI.DateVersion := EncodeDate(2004, 09, 10);
  //V_PGI.SAV:=True ;
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  //V_PGI.NiveauAcces:=1 ;
  V_PGI.MenuCourant := 0;
  //V_PGI.Decla100:=TRUE ;
  { Pour faire du DEBUG  }
{$IFDEF ENCOURS}
  V_PGI.DebugSQL := False;
  V_PGI.Debug := False;
  V_PGI.SAV := False;
{$ELSE}
  V_PGI.DebugSQL := False;
  V_PGI.Debug := False;
  V_PGI.SAV := False;
{$ENDIF}
  V_PGI.QRPdf := True;
  V_PGI.NumMenuPop := 27;
  //  version 422 AGL
  V_PGI.CegidBureau := True;
  V_PGI.CegidApalatys := FALSE;
  // pour autoriser l'acces au dp comme dossier de travail,au multi dossier
  V_PGI.StandardSurDP := true;
  // pour autoriser l'acces au dp comme base de reference pour tous les dossiers ==> Table DP+STD
  V_PGI.MajPredefini := True;
  // Confidentialité
  V_PGI.PGIContexte := [CtxPaie];
  V_PGI.RAZForme := TRUE;
  V_PGI.DispatchTT := DispatchTT;
  V_PGI.TabletteHierarchiques := TRUE;
  // Blocage des mises à jour de structure par l'appli. Elle doit etre faite par PgiMajVEr fourni dans le kit socref
  V_PGI.BlockMAJStruct := TRUE;
{$IFDEF PAIEOLE}
  RegisterPaiePGI;
{$ENDIF}
  if V_PGI.ModePcl = '1' then
  begin
{$IFNDEF EAGLCLIENT}
    V_PGI.ToolBarreUp := False;
    V_PGI.PortailWeb := 'http://experts.cegid.fr/home.asp';
{$ENDIF}
  end
  else V_PGI.PortailWeb := 'http://utilisateurs.cegid.fr';
{$IFNDEF EAGLCLIENT}
  if V_PGI.ModePcl = '1' then

    V_PGI.SansMessagerie := true
  else V_PGI.SansMessagerie := FALSE;
{$ENDIF}
  V_PGI.CodeProduit := '013';
{$IFDEF TESTSIC}
  TitreHalley := 'TEST ' + TitreHalley;
{$IFNDEF EAGLCLIENT}
  NomHalley := 'CPS5SIC';
{$ELSE}
  NomHalley := 'eCPS5SIC';
{$ENDIF}

{$ENDIF}

end;

initialization
  ProcChargeV_PGI := InitLaVariablePGI;
end.

