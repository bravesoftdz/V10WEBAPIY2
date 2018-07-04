{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/07/2001
Modifié le ... : 23/07/2001
Description .. : Gestion du menu du Front Office
Mots clefs ... : FO
*****************************************************************}
unit MDispFO;

interface

uses
  Controls,                                                
  {$IFDEF EAGLCLIENT}
  MenuOLX, MaineAGL, eTablette, UtileAGL,
  {$ELSE}
  MenuOLG, FE_Main, Tablette, EdtQR, EdtEtat, Edtdoc, UEdtComp,    
  {$ENDIF}
  Forms, Sysutils, HMsgBox, HEnt1, HPanel, AGLInit, Hctrls, UiUtil, GrpFavoris;

procedure InitApplication;

type
  TFMenuDisp = class(TDatamodule)
  end;

var FMenuDisp: TFMenuDisp;

implementation

uses
  {$IFNDEF EAGLCLIENT}
  UtilSoc, InitSoc, CreerSoc, CopySoc, Reseau, Mullog, Traduc,
  UtilPGI, UtomUtilisat, Devise, Souche, ListeDeSaisie,
  LP_Param, LP_Base,
  galOutil,
  {$ENDIF}
  Ent1, TiersUtil, FichCOmm, ZoomEdtGescom, ParamSoc, Confidentialite_TOF,
  EntGC, UtilGC, Facture, FactUtil, Transfert, UtilArticle, UtilDispGC,
  {$IFDEF TOXCLIENT}
  FOExportCais, FOImportCais,
  ToxSim, ToxDetop, uToxConf, ToxConsultBis, GCTOXWORK, GCToxCtrl, uToxFiches,
  uToxConst, FOToxUtil, RECUPFICHIER_TOF, AssistImportGPAO,
  {$ENDIF}
  AssistFinInv, FactPiece,
  FODefi, FOChoixCaisse, FOUtil, TickUtilFO, TicketFO, FOOuvreJour, FOFermeJour,
  FOConsultCaisse, FOParamTicket, CalcOLEGescom, CalcMulPGI, GrStaMR, GrJH,
  Reassort;
  

{$R *.DFM}
const
  // Libellés des messages
  TexteMessage: array[1..13] of string = (
    {1}'Fonction non disponible : ',
    {2}'Fonction non autorisée : ',
    {3}'Pour accéder à cette fonction, la caisse doit être ouverte !',
    {4}'Pour accéder à cette fonction, la caisse doit être fermée !',
    {5}'Aucun tiers par défaut n''a été défini dans les paramètres sociétés.#13#13Vous ne pouvez donc pas créer de transfert !',
    {6}'La nature de pièce des transferts n''est pas correctement renseignée.#13#13Vous ne pouvez donc pas créer de transfert !',
    {7}'Pour accéder à cette fonction, la caisse doit avoir été ouverte au moins une fois !',
    {8}'Aucun chèque n''est à remettre en banque !',
    {9}'Fonction interdite !',
    {10}'La caisse est déjà ouverte !',
    {11}'La caisse est déjà fermée !',
    {12}'Attention, la date d''ouverture de la caisse %% ne correspond pas à la date système $$',
    {13}'Voulez-vous continuer ?'
    );

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 11/03/2003
  Modifié le ... : 11/03/2003
  Description .. : Vérifie si la date d''ouverture de la caisse correspond bien à
  Suite ........ : la date système.
  Mots clefs ... :
  *****************************************************************}

function VerifDateJournee: boolean;
var Stg: string;
begin
  Result := True;
  if V_PGI.DateEntree <> Date then
  begin
    Stg := TraduireMemoire(TexteMessage[12]);
    Stg := FindEtReplace(Stg, '%%', '(' + DateToStr(V_PGI.DateEntree) + ')', False);
    Stg := FindEtReplace(Stg, '$$', '(' + DateToStr(Date) + ')', False);
    if FOJaiLeDroit(81, False, False) then
    begin
      Stg := Stg + #10#10 + TraduireMemoire(TexteMessage[13]);
      Result := (HShowMessage('2;?caption?;' + Stg + ';Q;YN;Y;N', TitreHalley, '') = mrYes);
    end else
    begin
      Result := FOJaiLeDroit(81, True, True, Stg);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Vérifie l'état de la caisse selon la ligne de menu choisie.
Suite ........ :
Suite ........ : Si Verifie est à vrai on contrôle la cohérence de l'état de la
Suite ........ : journée par rapport à la ligne de menu choisie, sinon on
Suite ........ : indique si la ligne de menu nécessite une vérification de
Suite ........ : l'état de la journée.
Mots clefs ... :
*****************************************************************}

function VerifEtatCaisse(Num: integer; Verifie: boolean): boolean;
var EtatCaisse: string;
  NoErr: integer;
begin
  NoErr := 2;
  // Rechargement des paramètres et de l'état de la caisse (ouverte, fermée, non ouverte...)
  EtatCaisse := ETATJOURCAISSE[1];
  if Verifie then EtatCaisse := FOEtatCaisse(FOCaisseCourante);
  // Vérification de l'état de la caisse en fonction de la ligne de menu choisie
  case Num of
    107110, // Saisie des tickets de vente
    107130, // Annulation d'un ticket
    107150, // Modification des règlements
    107180, // Rattachement d'un ticket à un client
    107310, // Ticket X
    108210, // Saisie des Transferts
    108310, // Saisie des Commandes de réassort
    108320, // Modification des Commandes de réassort
    108410, // Saisie des Bons de réception
    108420, // Modification des Bons de réception
    108450, // Génération des Bons de réception à partir des annonces de livraison
    108230, // Modification des Transferts Inter-Boutiques
    108240: // Validation des Transferts Inter-Boutiques
      begin
        if Verifie then
          Result := (EtatCaisse = ETATJOURCAISSE[2]) // la journée doit être ouverte
        else
          Result := True;
        NoErr := 3;
        if (Num <> 107310) and (Result) and (Verifie) then
        begin
          Result := VerifDateJournee;
          NoErr := 0; // le message d'erreur a déjà été affiché
        end;
      end;
    107220: // Fermeture de la journée de vente
      begin
        if Verifie then
          Result := (EtatCaisse = ETATJOURCAISSE[2]) // la caisse est déjà fermée
        else
          Result := True;
        NoErr := 11;
      end;
    107210: // Ouverture de la journée de vente
      begin
        if Verifie then
          Result := (EtatCaisse <> ETATJOURCAISSE[2]) // la caisse est déjà ouverte
        else
          Result := True;
        NoErr := 10;
      end;
    107320: // Ticket Z
      begin
        if Verifie then
          Result := (EtatCaisse <> ETATJOURCAISSE[2]) // la journée doit être fermée
        else
          Result := True;
        NoErr := 4;
      end;
    107230: // Etat de la journée
      begin
        if Verifie then
          Result := (EtatCaisse <> ETATJOURCAISSE[1]) // la journée doit être définie
        else
          Result := True;
        NoErr := 7;
      end;
    107120, // Consultation d'un ticket
    107140, // Liste des règlements
    107160, // Liste des chèques différés
    107170, // Liste des articles vendus
    107240, // Remise en banque
    107250, // Impression du bordereau de remise en banque
    107330, // récapitulatif par vendeurs
    107340, // liste des règlements
    107410, // Situation Flash
    107420, // Répartition par mode de paiement
    107430, // Meilleure journée / Meilleure heure
    107440, // Situation sur plusieurs clôtures
    107450, // Statistiques par famille article
    107460, // Statistiques sur les remises
    108220, // Consultation des Transferts Inter-Boutiques
    108330, // Consultation des Commandes de réassort
    108340, // Consultation des lignes de Commandes de réassort
    108430, // Consultation des Bons de réception
    108440, // Consultation des lignes de Bons de réception
    112341, // Reprise des liaisons des opérations de caisse
    112342, // Reprise des liaisons des règlements
    112530, // Paramétrage de la saisie de ticket
    109310: // Mise à jour manuelle de la fidélité
      Result := True; // la connexion à une caisse est obligatoire
  else
    begin
      if Verifie then
        Result := True
      else
        Result := False; // la connexion à une caisse n'est pas nécessaire
    end;
  end;
  if (not Result) and (Verifie) and (NoErr > 0) then
    HShowMessage('2;?caption?;' + TraduireMemoire(TexteMessage[NoErr]) + ';W;O;O;O;', TitreHalley, '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Vérifie si on s'est bien connecé à une caisse
Mots clefs ... :
*****************************************************************}

function VerifConnexionCaisse(Num: integer): boolean;
var Ok: boolean;
begin
  Result := True;
  // Pas de connexion pour les lignes du menu AGL et de gestion de la caisse
  if Num < 1000 then Exit;
  if (Num = 112110) or (Num = 112145) or (Num = 112160) then Exit;
  // Connexion systèmatiquement à chaque lancement d'une ligne du menu
  if FOOptionConnexionCaisse(cfoChaqueFois) then FOLibereVHGCCaisse;
  // Connexion pour toutes les lignes du menu
  if FOOptionConnexionCaisse(cfoToutesLignes) then
    Ok := True
  else
    Ok := VerifEtatCaisse(Num, False);
  if Ok then
  begin
    // Vérification si une caisse est sélectionnée
    if (not FOVerifieVHGCCaisse) or (not FOCaisseDisponible) then
    begin
      Result := FOChoixCodeCaisse;
      if Result then FOChgStatus;
    end;
    // Vérification de l'état de la journée
    if Result then Result := VerifEtatCaisse(Num, True);
  end;
  {$IFNDEF EAGLCLIENT}
  if not Result then FMenuG.VireInside(nil);
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Lancement de l'éditeur de ticket
Mots clefs ... :
*****************************************************************}

procedure LanceParamModele(TypeModele: TTypeEtatFO; PRien: THPanel; var RetourForce: boolean);
var sNature, sCodeEtat, sTitre: string;
begin
  RetourForce := True;
  FODonneCodeEtat(TypeModele, True, sNature, sCodeEtat, sTitre);
  {$IFNDEF EAGLCLIENT}
  Param_LPTexte(nil, 'T', sNature, sCodeEtat, True);
  LibereSauvgardeLP;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Lancement du paramétrage du poste de travail
Mots clefs ... :
*****************************************************************}

procedure LanceParamPoste(PRien: THPanel; var ReaffBoutons: boolean);
var sCaisse: string;
begin
  ReaffBoutons := True;
  PRien.InsideTitle.Caption := TraduireMemoire('Paramétrage du poste de travail');
  V_PGI.ZoomOLE := True;
  sCaisse := AGLLanceFiche('MFO', 'PARAMPOSTE', '', '', ActionToString(taModif));
  V_PGI.ZoomOLE := False;
  if (sCaisse <> '') and (sCaisse <> FOCaisseCourante) then
    FOLibereVHGCCaisse;
  {$IFDEF EAGLCLIENT}
  if PRien <> nil then PRien.CloseInside;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Lancement e la remise en banque
Mots clefs ... :
*****************************************************************}

procedure LanceParamRemisBanque(Num: integer; var ReaffBoutons: boolean);
var
  sLequel, sSql, sNat: string;
begin
  sNat := FOGetNatureTicket(False, False) ;
  sLequel := GetSoucheG(sNat, VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'), '');
  sSQL := 'SELECT GPE_NUMERO FROM GCREGLEMENTFO WHERE GPE_NATUREPIECEG="'+ sNat +'" '
    + 'AND GPE_SOUCHE="' + sLequel + '" '
    + 'AND (MP_TYPEMODEPAIE="' + TYPEPAIECHEQUE + '" '
    + 'OR MP_TYPEMODEPAIE="' + TYPEPAIECHQDIFF + '") '
    + 'AND GPE_CHQDIFUTIL="-"';
  if ExisteSQL(sSql) then
    AGLLanceFiche('MFO', 'REMISEBANQUE', '', '', 'CAISSE=' + FOCaisseCourante + ';NATUREPIECE='+ sNat)
  else
  begin
    PGIBox(TexteMessage[8], 'REMISE EN BANQUE');
    ReaffBoutons := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Lancement d'une édition au format ticket
Mots clefs ... :
*****************************************************************}

procedure LanceEdition(Num: integer; PRien: THPanel; var ReaffBoutons: boolean);
var sLequel, sTitre: string;
  TypeEtat: TTypeEtatFO;
  iNumZ: integer;
begin
  ReaffBoutons := True;
  case Num of
    107250: // Impression du bordereau de remise en banque
      begin
        sTitre := 'Impression du bordereau de remise en banque';
        sLequel := '(MP_TYPEMODEPAIE = "' + TYPEPAIECHEQUE + '" OR '
          + 'MP_TYPEMODEPAIE = "' + TYPEPAIECHQDIFF + '") AND '
          + 'GPE_DATEREMBNQ ="' + USDateTime(Now) + '" AND '
          + 'GPE_CHQDIFUTIL="X" AND GPE_NATUREPIECEG=' + FOGetNatureTicket(False, True);
        TypeEtat := efoRemiseBq;
      end;
    107310: // Ticket X
      begin
        sTitre := 'Ticket X';
        sLequel := FOMakeWhereTicketX('GP');
        TypeEtat := efoTicketX;
      end;
    107320: // Ticket Z
      begin
        iNumZ := FOGetNumZCaisse(FOCaisseCourante);
        sTitre := 'Ticket Z';
        sLequel := FOMakeWhereTicketZ('GP', '', iNumZ, iNumZ);
        TypeEtat := efoTicketZ
      end;
    107330: // récapitulatif par vendeurs
      begin
        sTitre := 'Récapitulatif par vendeurs';
        sLequel := FOMakeWhereTicketX('GP');
        TypeEtat := efoRecapVend;
      end;
    107340: // liste des règlements
      begin
        sTitre := 'Liste des règlements';
        sLequel := FOMakeWhereTicketX('GP');
        TypeEtat := efoListeRegle;
      end;
  else Exit;
  end;
  PRien.InsideTitle.Caption := TraduireMemoire(sTitre);
  FOLanceImprimeLP(TypeEtat, sLequel, True, nil);
  {$IFDEF EAGLCLIENT}
  if PRien <> nil then PRien.CloseInside;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Lancement d'un MUL sur les tickets de caisse
Mots clefs ... :
*****************************************************************}

procedure LanceMULTicket(Num: integer; var ReaffBoutons: boolean);
///////////////////////////////////////////////////////////////////////////////////////
  function MakeLequel(Prefixe: string; AvecNumZ: boolean = False; AvecNumCais: boolean = True): string;
  var sDateOuv: string;
  begin
    sDateOuv := DateToStr(FOGetParamCaisse('GPK_DATEOUV'));
    Result := Prefixe + '_NATUREPIECEG='+ FOGetNatureTicket(False, False) +';'
      + Prefixe + '_DATEPIECE=' + sDateOuv + ';'
      + Prefixe + '_DATEPIECE_=' + sDateOuv;
    if AvecNumCais then
      Result := Result + ';' + Prefixe + '_CAISSE=' + FOCaisseCourante;
    if AvecNumZ then
      Result := Result + ';' + Prefixe + '_NUMZCAISSE=' + IntToStr(FOGetParamCaisse('GPK_NUMZCAISSE'));
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
var sLequel: string;
begin
  case Num of
    107120: // Consultation d'un ticket
      begin
        sLequel := MakeLequel('GP');
        AGLLanceFiche('MFO', 'PIECEDEV_MUL', sLequel, '', 'TICKET');
      end;
    107130: // Annulation d'un ticket
      begin
        sLequel := MakeLequel('GP') + ';GP_TICKETANNULE=-'
                 + ';NEWNATURE=' + FOGetNatureTicket(False, False);
        AGLLanceFiche('MFO', 'ANNULTIC_MUL', sLequel, '', 'ANNULTIC');
      end;
    107140: // Liste des règlements
      begin
        sLequel := MakeLequel('GP');
        AGLLanceFiche('MFO', 'PIEDECHE_MUL', sLequel, '', 'ECHEANCE');
      end;
    107150: // Modification des règlements
      begin
        sLequel := MakeLequel('GP', True) + ';GP_TICKETANNULE=-';
        AGLLanceFiche('MFO', 'PIEDECHE_MUL', sLequel, '', 'MODIFICATION');
      end;
    107160: // Liste des chèques différés
      begin
        sLequel := MakeLequel('GP') + ';MP_TYPEMODEPAIE=' + TYPEPAIECHQDIFF;
        AGLLanceFiche('MFO', 'CHQDIF_MUL', sLequel, '', 'ECHEANCE');
      end;
    107170: // Liste des articles vendus
      begin
        sLequel := MakeLequel('GL') + ';GL_TYPEARTICLE=MAR';
        if (V_PGI.LaSerie > S3) and (GetParamSoc('SO_ARTLOOKORLI')) then
          AGLLanceFiche('MFO', 'PIECARTS5_MUL', sLequel, '', 'TICKET')
        else
          AGLLanceFiche('MFO', 'PIECART_MUL', sLequel, '', 'TICKET');
      end;
    107180: // Rattachement d'un ticket à un client
      begin
        sLequel := MakeLequel('GP', True) + ';GP_TIERS=' + FODonneClientDefaut
          + ';GP_TICKETANNULE=-;NEWNATURE=' + FOGetNatureTicket(False, False);
        AGLLanceFiche('MFO', 'ANNULTIC_MUL', sLequel, '', 'ASSIGNTIC');
      end;
    112341: // Reprise des liaisons des opérations de caisse
      begin
        sLequel := MakeLequel('GL', False, False)
          + ';GL_TYPEARTICLE=FI;GL_ETABLISSEMENT=' + VH^.EtablisDefaut
          + ';XX_WHERE=GA_TYPEARTFINAN IN ("ABA","VAR")';
        AGLLanceFiche('MFO', 'MFOREPRISELIENOPC', sLequel, '', '');
      end;
    112342: // Reprise des liaisons des règlements
      begin
        sLequel := MakeLequel('GPE', False, False)
          + ';GP_ETABLISSEMENT=' + VH^.EtablisDefaut
          + ';XX_WHERE=MP_TYPEMODEPAIE IN ("002","007")';
        AGLLanceFiche('MFO', 'MFOREPRISELIENREG', sLequel, '', '');
      end;
  else Exit;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Fabrique la liste des tags à supprimer
Mots clefs ... :
*****************************************************************}

function DoTagToRemove: string;
///////////////////////////////////////////////////////////////////////////////////////
  procedure AddTagToListe(var Liste: string; Tag: string);
  var Sep: string;
  begin
    Sep := ';';
    if (Liste = '') or (Liste[Length(Liste)] = ';') then Sep := '';
    if Tag[1] = ';' then Delete(Tag, 1, 1);
    Liste := Liste + Sep + Tag;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
begin
  Result := '';
  AddTagToListe(Result, '107900');
  if FOComptaEstActive(FOGetNatureTicket(False, False)) then AddTagToListe(Result, '107150;107180');
  AddTagToListe(Result, '111230,111240,111330');
  {$IFDEF MODES3}
  AddTagToListe(Result, '108300;108125;108130;-108140;108146;108400;108500;108551;108552');
  AddTagToListe(Result, '109200');
  AddTagToListe(Result, '112442;112433;112434;112435;112436');
  {$ENDIF}

  {$IFNDEF TOXCLIENT}
  AddTagToListe(Result, '107260');
  {$ENDIF}

  {$IFDEF EAGLCLIENT}
  AddTagToListe(Result, '107350');
  AddTagToListe(Result, '108125;108591');
  AddTagToListe(Result, '112115;112140;112180;112190;112510;112522;112523;112524;112200;112300;112400');
  {$ENDIF}

  if GetParamSoc('SO_GCFOBASEFORMATION') then AddTagToListe(Result, '112330');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Lancement d'une fonction avec un écran maximisé
Mots clefs ... :
*****************************************************************}

procedure LanceMaximise(Num: integer; PRien: THPanel; var ReaffBoutons: boolean);
var
  FF: Forms.TForm;
begin
  FF := nil;
  if PRien <> nil then
  begin
    FF := Forms.TForm.Create(Application);
    InitInside(FF, PRien);
    FF.Show;
  end;

  case Num of
    107110: // Saisie des tickets de vente
      begin
        CreerTicketFO(FOGetNatureTicket(False, False));
        ReaffBoutons := True;
      end;
    107210: // Ouverture de la journée de vente
      begin
        FOOuvreJournee(PRien);
        ReaffBoutons := True;
      end;
    107220: // Fermeture de la journée de vente
      begin
        ReaffBoutons := FOFermeJournee(PRien);
      end;
    107410: // Situation Flash
      begin
        FOConsultationCaisse(FOCaisseCourante, '', '', nil, True);
        ReaffBoutons := True;
      end;
    112530: // Saisie de ticket
      begin
        SaisieParamTicketFO(FOGetNatureTicket(False, False), taModif, FOCaisseCourante);
        ReaffBoutons := True;
      end;
  end;

  if FF <> nil then
  begin
    FF.Close;
    FF.Free;
  end;
  if PRien <> nil then PRien.CloseInside;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Dispatch
///////////////////////////////////////////////////////////////////////////////////////

procedure Dispatch(Num: Integer; PRien: THPanel; var RetourForce, SortieHalley: boolean);
var sCaisse, NumZ: string;
  {$IFNDEF EAGLCLIENT}
  StVire, StSeule, stMessage: string;
  {$ENDIF}
  ReaffBoutons: Boolean;
begin
  // Connexion à la caisse pour toutes les lignes de menu sauf la libération des caisses
  if not VerifConnexionCaisse(Num) then
  begin
    {$IFDEF EAGLCLIENT}
    if PRien <> nil then PRien.CloseInside;
    {$ENDIF}
    Exit;
  end;
  ReaffBoutons := False;
  case Num of
    10:
      begin //  après connexion
      {$IFNDEF EAGLCLIENT}
        if PCL_IMPORT_BOB(NomHalley) = 1 then
        begin
          stMessage := 'LE SYSTEME A DETERMINE UNE MAJ DE MENU';
          stMessage := stMessage +#13#10 +'POUR ETRE PRISE EN COMPTE';
          stMessage := stMessage +#13#10 +'VOUS DEVREZ RELANCER L''APPLICATION';
          PGIInfo(stMessage,'MAJ MENU');
        end;
      {$ENDIF}
        UpdateCombosGC;
        GCModifLibelles;
        GCTripoteStatus;
      {$IFDEF MODE}
        VH_GC.GcMarcheVentilAna:='MODE';
      {$ELSE}
        VH_GC.GcMarcheVentilAna:='AFF';     // gm 09/01/03 même pour la GC ,il faut mettre AFF
      {$ENDIF}
        {$IFDEF TOXCLIENT}
        FOGetToxIdApp; // initialisation éventuelle des paramètres de la Tox
        FOToxServeurMajConf;
        {$ENDIF}
        FOApresConnexion;
      end;
    11: FOApresDeConnexion; // après déconnexion
    12: if FOAvantConnexion(SortieHalley) then Exit; // avant connexion et séria
    13:
      begin // avant déconnexion
        // Liberation de la caisse
        FOLibereVHGCCaisse;
        ChargeMenuPop(-1, FMenuG.DispatchX);
      end;
    15: ; ///////////
    16: ; // OT 29/04/2003 le nouvel agl par paquet provoque une fonction 16 non disponible !!
    18: ; ///////////
    19: ; ///////////
    ///////  PAS ENCAISSEMENT  ////////////
    ///////////////////////////////////
        ///// //////////
    {$IFNDEF EAGLCLIENT}
    112421: // Paramètres - Société
      begin
        BrancheParamSocAffiche(StVire, StSeule);
        ParamSociete(False, StVire, StSeule, '', nil, ChargePageSoc, SauvePageSoc, InterfaceSoc, 301300200);
      end;
    112422: // Gestionnaire société
      GestionSociete(PRien, @InitSociete, nil);
    112423: // Recopie société à société
      RecopieSoc(PRien, True);
    112431: // groupe utilisateurs
      begin
        AGLLanceFiche('YY', 'YYUSERGROUP', '', '', '');
        FOLibereVHGCCaisse;
      end;
    112432: // Utilisateurs
      begin
        FicheUSer(V_PGI.User);
        ControleUsers;
      end;
    112433: // Restrictions utilisateurs
      AGLLanceFiche('YY', 'PROFILETABL', '', '', 'ETA');
    112434: // Utilisateurs connectés
      ReseauUtilisateurs(False);
    112435: // Journal des évènements
      AGLLanceFiche('GC', 'GCJNALEVENT_MUL', '', '', '');
    112436: // Suivi d'activité
      VisuLog;
    112437: // RAZ connexions
      ReseauUtilisateurs(True);
    112438: // Gestion droits d'accès
      GCLanceFiche_Confidentialite( 'YY', 'YYCONFIDENTIALITE', '', '', '26;27;107;108;109;111;112'  );
    112442: // Multilingue - personnalisation
      ParamTraduc(TRUE, PRien);
    112441: // Multilingue - Traduction
      ParamTraduc(FALSE, PRien);
    {$ENDIF}

    ///////  107 : ENCAISSEMENT  ////////////
    /////////////////////////////////////////
    107110: // Saisie des tickets de vente
      LanceMaximise(Num, PRien, ReaffBoutons);
    107120: // Consultation d'un ticket
      LanceMULTicket(Num, ReaffBoutons);
    107130: // Annulation d'un ticket
      LanceMULTicket(Num, ReaffBoutons);
    107140: // Liste des règlements
      LanceMULTicket(Num, ReaffBoutons);
    107150: // Modification des règlements
      LanceMULTicket(Num, ReaffBoutons);
    107160: // Liste des chèques différés
      LanceMULTicket(Num, ReaffBoutons);
    107170: // Liste des articles vendus
      LanceMULTicket(Num, ReaffBoutons);
    107180: // Rattachement d'un ticket à un client
      LanceMULTicket(Num, ReaffBoutons);
    107210: // Ouverture de la journée de vente
      LanceMaximise(Num, PRien, ReaffBoutons);
    107220: // Fermeture de la journée de vente
      LanceMaximise(Num, PRien, ReaffBoutons);
    107230: // Etat de la journée
      begin
        sCaisse := FOCaisseCourante;
        NumZ := IntToStr(FOGetNumZCaisse(sCaisse));
        AGLLanceFiche('MFO', 'JCAISSE', '', '', 'GJC_CAISSE=' + sCaisse + ';GJC_NUMZCAISSE=' + NumZ);
      end;
    107240: // Remise en banque
      LanceParamRemisBanque(Num, ReaffBoutons);
    107250: // Impression du bordereau de remise en banque
      LanceEdition(Num, PRien, ReaffBoutons);
    {$IFDEF TOXCLIENT}
    107260: // Retansmission d'une journée
      AGLLanceFiche('MFO', 'RETRANSJOUR', '', '', '');
    {$ENDIF}
    107310: // Ticket X
      LanceEdition(Num, PRien, ReaffBoutons);
    107320: // Ticket Z
      LanceEdition(Num, PRien, ReaffBoutons);
    107330: // récapitulatif par vendeurs
      AGLLanceFiche('MFO', 'RECAPVEN_MUL', 'GJC_CAISSE=' + FOCaisseCourante, '', ActionToString(taConsult));
    107340: // liste des règlements
      LanceEdition(Num, PRien, ReaffBoutons);
    {$IFNDEF EAGLCLIENT}
    107350: // Lanceur d'états libres utilisateurs
      begin
        LanceEtatLibreGC;
        ReaffBoutons := True;
      end;
    {$ENDIF}
    107410: // Situation Flash
      LanceMaximise(Num, PRien, ReaffBoutons);
    107420: // Répartition par mode de paiement
      StatparModeReglement(PRien, FOGetNatureTicket(False, False));
    107430: // Meilleure journée / Meilleure heure
      MeuilleurJouretHour(PRien, FOGetNatureTicket(False, False));
    107440: // Situation sur plusieurs clôtures
      begin
        sCaisse := FOCaisseCourante;
        NumZ := IntToStr(FOGetNumZCaisse(sCaisse));
        AGLLanceFiche('MFO', 'SITUFLASH_MUL', 'GJC_CAISSE=' + sCaisse + ';GJC_NUMZCAISSE_=' + NumZ, '', '');
      end;
    107450: // Statistiques par famille article
      AGLLanceFiche('MFO', 'GRSTATFAM', '', '', 'GL_CAISSE=' + FOCaisseCourante + ';NATUREPIECE='+ FOGetNatureTicket(False, False)) ;
    107460: // Statistiques sur les remises
      AGLLanceFiche('MFO', 'GRSTATREM', '', '', 'GL_CAISSE=' + FOCaisseCourante+ ';NATUREPIECE='+ FOGetNatureTicket(False, False));
    ///////  108 : GESTION  /////////////////
    /////////////////////////////////////////
      /////// 108100 : Stocks ////////////////////
    108120: // Disponible article par dimension
      DispatchArtMode(3, 'GQ_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    {$IFNDEF EAGLCLIENT}
    108125: // Stock distant
      AGLLanceFiche('MBO', 'DISPODIST_MUL', '', '', 'MULDISTANT');
    {$ENDIF}
    108127: // Stock local
      AGLLanceFiche('MBO', 'DISPODIST_MUL', '', '', 'MULLOCAL');
    108130: // Statistique stock disponible
      DispatchArtMode(13, 'GQ_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    108141: // Etiquettes articles sur catalogue
      DispatchArtMode(1, '', '', 'ETIQARTCAT');
    108142: // Etiquettes articles à la demande
      AGLLanceFiche('GC', 'GCETIARTDEM_MODE', '', '', '');
    108143: // Etiquettes articles sur retour de vente
      AGLLanceFiche('GC', 'GCETIARTFFO_MODE', 'GL_NATUREPIECEG='+ FOGetNatureTicket(False, False) +';GL_TYPEARTICLE=MAR;GL_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'FFO');
    108144: // Etiquettes articles sur commandes
      AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=FCF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'FCF');
    108145: // Etiquettes articles sur receptions
      AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=BLF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'BLF');
    108146: // Etiquettes articles sur transferts
      AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=TRE;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'TRE');
    108147: // Etiquettes articles sur stock
      AGLLanceFiche('GC', 'GCETIARTSTK_MODE', 'ETABLISSEMENT=' + VH^.EtablisDefaut, '', '');
    108148: // Etiquettes articles sur annonce de livraison
      AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=ALF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'ALF');
    /////// 108200 : Transferts Inter-Boutiques ////////////////////
    108210: // Saisie des Transferts Inter-Boutiques
      ReaffBoutons := not CreerTransfert('TEM', 'TEM_TRV_NUMIDENTIQUE');
    108220: // Consultation des Transferts Inter-Boutiques
      AGLLanceFiche('GC', 'GCTRANSFERT_MUL', 'GP_NATUREPIECEG=TEM;GP_DEPOT=' + VH_GC.GCDepotDefaut, '', 'CONSULTATION');
    108230: // Modification des Transferts Inter-Boutiques
      AGLLanceFiche('GC', 'GCTRANSFERT_MUL', 'GP_NATUREPIECEG=TEM', '', 'MODIFICATION');
    108240: // Validation des Transferts Inter-Boutiques
      AGLLanceFiche('GC', 'GCTRANSFERT_MUL', 'GP_NATUREPIECEG=TRV', '', 'VALIDATION');
    /////// 108300 : Commandes de réassort ////////////////////
    108310: // Saisie des Commandes de réassort
      CreerCommandeReassort('FCF');
    108320: // Modification des Commandes de réassort
      AGLLanceFiche('MFO', 'REASSORT_MUL', 'GP_NATUREPIECEG=FCF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'MODIFICATION');
    108330: // Consultation des Commandes de réassort
      AGLLanceFiche('MFO', 'REASSORT_MUL', 'GP_NATUREPIECEG=FCF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'CONSULTATION');
    108340: // Consultation des lignes de Commandes de réassort
      AGLLanceFiche('MFO', 'REASSORTLIGNE_MUL', 'GL_NATUREPIECEG=FCF;GL_TYPELIGNE=ART;GL_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'CONSULTATION');
    /////// 108400 : Bons de réception ////////////////////
    108410: // Saisie des Bons de réception
      CreerPiece('BLF');
    108420: // Modification des Bons de réception
      AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=BLF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'MODIFICATION');
    108430: // Consultation des Bons de réception
      AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=BLF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'CONSULTATION');
    108440: // Consultation des lignes de Bons de réception
      AGLLanceFiche('GC', 'GCLIGNEACH_MUL', 'GL_NATUREPIECEG=BLF;GL_TYPELIGNE=ART;GL_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'CONSULTATION');
    108450: // Génération des Bons de réception à partir des annonces de livraison
      AGLLanceFiche('GC', 'GCTRANSACH_MUL', 'GP_NATUREPIECEG=ALF;GP_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'BLF');
    /////// 108500 : Inventaire ////////////////////
    108512: // Edition de l'état préparatoire
      AGLLanceFiche('GC', 'GCLISTEINVPRE', '', '', '');
    108540: // Saisie de l'inventaire
      AGLLanceFiche('GC', 'GCSAISIEINV_MUL', 'GIE_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    108560: // Validation de l'inventaire
      AGLLanceFiche('GC', 'GCVALIDINV_MUL', 'GIE_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    108521: // Transmission inventaire
      AGLLanceFiche('MBO', 'TRANSTPI', '', '', '');
    108522: // Inventaires transmis
      AGLLanceFiche('MBO', 'TRANSINV_MUL', 'GIT_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    108523: // Intégration
      AGLLanceFiche('GC', 'GCINTEGRINV_MUL', 'GIT_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    108571: // Consultation de la liste des écarts d'inventaires générés
      AGLLanceFiche('GC', 'GCLIGNEINV_MUL', 'GL_ETABLISSEMENT=' + VH^.EtablisDefaut, '', 'CONSULTATION');
    108511: // Génération liste préparatoire d'inventaire
      AGLLanceFiche('MBO', 'LISTEINV', '', '', '');
    108580: // Fin d'inventaire
      Assist_FinInvent;
    108572: // Edition de la liste des écarts d'inventaire générés
      AGLLanceFiche('MBO', 'EDTECARTINV', 'GL_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    108551: // Edition de l'inventaire
      AGLLanceFiche('MBO', 'EDTLISTEINV', 'GIL_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    108552: // Etat comparatif Stock/Inventaire
      AGLLanceFiche('MBO', 'EDTCOMPSTKINV', 'GIL_DEPOT=' + VH_GC.GCDepotDefaut, '', '');
    {$IFNDEF EAGLCLIENT}
    108591:  // Intégration du fichier inventoriste
      begin
        Assist_ImportGPAO(True);
        RetourForce := TRUE;
      end; // Reprise du fichier inventaire
    {$ENDIF}
    108592: // Paramétrage du format du fichier inventaire
      AGLLanceFiche('GC', 'GCREPRISEDONNEES', 'SAISIE_INVENTAIRE=X', '', ''); // Paramétrage du fichier inventaire
    /////// 109 : CLIENTS ////////////////////
    //////////////////////////////////////////
    109110: // Modification des fiches clients
      AGLLanceFiche('GC', 'GCCLIMUL_MODE', 'T_NATUREAUXI=CLI', '', '');
    109120: // Historique des achats d'un client
      AGLLanceFiche('GC', 'GCCLIMUL_MODE', 'T_NATUREAUXI=CLI', '', 'HISTORIQUE');
    109130: // Liste des articles d'un client
      AGLLanceFiche('GC', 'GCCLIMUL_MODE', 'T_NATUREAUXI=CLI', '', 'LISTEARTICLE');
    109320: // Paramétrage de la fidélité
      AGLLanceFiche('MBO','MBOREGLEFID','','','ACTION=CONSULTATION') ;
    109310: // Consultation des cartes de fidélité
      AGLLanceFiche('MBO','FIDELITEENT','','','') ;
    109210: // Etiquettes clients
      AGLLanceFiche('GC', 'GCETIQCLI', '', '', '');

    /////// 112 : PARAMETRES /////////////////
    //////////////////////////////////////////
    112110: // Paramétrage des caisses
      AGLLanceFiche('MFO', 'PCAISSE', '', '', ActionToString(taModif));
    112115: // Modification en série du paramétrage des caisses
      AGLLanceFiche('MFO', 'MODIFCAIS_MUL', '', '', ActionToString(taModif));
    112125: // Etablissement
      AGLLanceFiche('GC', 'GCETABLISS_MUL', '', '', '');
    112140: // Démarques
      AGLLancefiche('MFO', 'DEMARQUE', '', '', ActionToString(taConsult));
    112145: // Libération des caisses
      AGLLanceFiche('MFO', 'LIBERECAISSE', '', '', ActionToString(taModif));
    112150: // Vendeurs
      AglLanceFiche('MFO', 'VENDEUR_MUL', 'GCL_TYPECOMMERCIAL=VEN;GCL_ETABLISSEMENT=' + VH^.EtablisDefaut, '', ActionToString(taConsult));
    112160: // Paramétrage du poste de travail
      LanceParamPoste(PRien, ReaffBoutons);
    112170: // Modes de Paiement
      AglLanceFiche('YY', 'YYMODEPAIE', '', '', ActionToString(taConsult));
    112180: // Détail des espèces
      AglLanceFiche('MFO', 'DETAILESPECES', V_PGI.DevisePivot + ';BIL', '', 'MODE=FO;DEVISE=' + V_PGI.DevisePivot);
    {$IFNDEF EAGLCLIENT}
    112190: // Devises
      FicheDevise(V_PGI.DevisePivot, taModif, False);
    112210: // Modèles d'impression texte : Tickets
      LanceParamModele(efoTicket, PRien, ReaffBoutons);
    112220: // Modèles d'impression texte : Chèques
      LanceParamModele(efoCheque, PRien, ReaffBoutons);
    112230: // Modèles d'impression texte : Tickets X
      LanceParamModele(efoTicketX, PRien, ReaffBoutons);
    112240: // Modèles d'impression texte : Tickets Z
      LanceParamModele(efoTicketZ, PRien, ReaffBoutons);
    112250: // Modèles d'impression texte : bordereau de remise en banque
      LanceParamModele(efoRemiseBq, PRien, ReaffBoutons);
    112261: // Modèles d'impression texte : récapitulatif par vendeur
      LanceParamModele(efoRecapVend, PRien, ReaffBoutons);
    112262: // Modèles d'impression texte : liste des règlements
      LanceParamModele(efoListeRegle, PRien, ReaffBoutons);
    112263: // Modèles d'impression texte : statistisques par famille article
      LanceParamModele(efoStatFam, PRien, ReaffBoutons);
    112264: // Modèles d'impression texte : statistisques remise
      LanceParamModele(efoStatRem, PRien, ReaffBoutons);
    112265: // Modèles d'impression texte : liste des articles vendus
      LanceParamModele(efoListeArtVendu, PRien, ReaffBoutons);
    112266: // Modèles d'impression texte : liste du fond de caisse
      LanceParamModele(efoFdCaisse, PRien, ReaffBoutons);
    112271: // Modèles d'impression texte : bon d'achat
      LanceParamModele(efoBonAchat, PRien, ReaffBoutons);
    112272: // Modèles d'impression texte : bon d'avoir
      LanceParamModele(efoBonAvoir, PRien, ReaffBoutons);
    112273: // Modèles d'impression texte : bon de versement d'arrhes
      LanceParamModele(efoBonArrhes, PRien, ReaffBoutons);
    112281: // Modèles d'impression texte : ticket de transfert
      LanceParamModele(efoTransfert, PRien, ReaffBoutons);
    112282: // Modèles d'impression texte : ticket de commande
      LanceParamModele(efoCommande, PRien, ReaffBoutons);
    112283: // Modèles d'impression texte : ticket de réception
      LanceParamModele(efoReception, PRien, ReaffBoutons);
    {$IFDEF TOXCLIENT}
    112310: // Export d'une caisse
      begin
        FOExportCaisse('');
        ReaffBoutons := True;
      end;
    112320: // Import d'une caisse
      if FOImportCaisse then SortieHalley := True else ReaffBoutons := True;
    {$ENDIF}
    112330: // Génération de la base de formation
      AGLLanceFiche('MFO', 'GENEREDBFORM', '', '', '');
    {$ENDIF}
    112341: // Reprise des liaisons des opérations de caisse
      LanceMULTicket(Num, ReaffBoutons);
    112342: // Reprise des liaisons des règlements
      LanceMULTicket(Num, ReaffBoutons);
    112410: // Paramétrage du groupe Favori
      ParamFavoris('', DoTagToRemove, False, False);
    {$IFNDEF EAGLCLIENT}
    112510: // Compteurs
      FicheSouche('GES');
    {$ENDIF}
    112521: // Natures
      {$IFDEF VRELEASE}
      {$IFNDEF EAGLCLIENT}
      if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
        AglLanceFiche('GC', 'GCPARPIECE', '', '', '') // LanceparPiece
      else AglLanceFiche('GC', 'GCPARPIECEUSR', '', '', ''); // LanceparPiece  utilisateur
    {$ELSE}
      AglLanceFiche('GC', 'GCPARPIECEUSR', '', '', ''); // LanceparPiece  utilisateur
    {$ENDIF} // EAGLCLIENT
    {$ELSE}
      AglLanceFiche('GC', 'GCPARPIECE', '', '', ''); // LanceparPiece
    {$ENDIF} // VRELEASE
    {$IFNDEF EAGLCLIENT}
    112522: // Modèles d'édition document
      begin
        EditDocument('L', 'GPI', '', True);
        ReaffBoutons := True;
      end;
    112523: // Modèles d'édition état
      begin
        EditEtat('E', 'GPJ', '', True, nil, '', '');
        ReaffBoutons := True;
      end;
    112524: // Listes de saisie
      EntreeListeSaisie('');
    {$ENDIF}
    112530: // Saisie de ticket
      LanceMaximise(Num, PRien, ReaffBoutons);
    {***
      112510 :       // Editeur d'états libres utilisateurs
             BEGIN
             EditEtatS5S7('E', 'UFO', '', True, Nil, '', '') ;
             ReaffBoutons := True;
             END ;
    ***}
    /////// 111 : LIAISON SITES DISTANTS /////////////////
    //////////////////////////////////////////////////////
    {$IFDEF TOXCLIENT}
    111100: AglToxSaisieParametres; // Paramètres par défaut
    111110: AglToxSaisieVariables; // Variables
    111120: AglToxMulSites; // Sites
    111130: AglToxSaisieGroupes; //  Groupes
    111140: AglToxMulConditions; // Requêtes
    111150: AglToxMulEvenements; // Evénements
    111200: // Démarrage/Arrêt des échanges
      begin
        //
        // 2 cas pour l'instant en fonction des paramètres sociétés
        //     A - On demande confirmation
        //     B - On intègre sans demander
        //
        if GetParamSoc('SO_GCTOXCONFIRM') = True then
        begin
          AglToxConf(aceConfigure, '', nil, GescomModeTraitementTox, nil);
        end else
        begin
          AglToxConf(aceConfigure, '', GescomModeNotConfirmeTox, GescomModeTraitementTox, nil);
        end;
      end;
    111210: AglToxMulChronos; // Consultation des échanges
    111220: AfficheCorbeille; // Consultation des corbeilles
    111300: ToxSimulation; // Intégration d'un fichier
    111310: ConsultManuelTox; // Visualisation d'un fichier
    111320: DetopeTox; // Détopage d'un fichier
    {$IFNDEF FOS5}
    111330: LanceGenereBE; // Génération des BE
    {$ENDIF}
    111340: LanceMiniFTP; // Mini FTP pour récupérer des fichiers sur le central
    {$ENDIF}

  else HShowMessage('2;?caption?;' + TraduireMemoire(TexteMessage[1]) + ';W;O;O;O;', TitreHalley, IntToStr(Num));
  end;
  {$IFNDEF EAGLCLIENT}
  if ReaffBoutons then FMenuG.VireInside(nil);
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DispatchTT
///////////////////////////////////////////////////////////////////////////////////////

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT, Range: string);
var
  ii: integer;
  Arg5LanceFiche: string;
begin
  // Sur le Front Office on ne peut créer ou modifier que des clients mais aucun paramètre.
  if (Num <> 8) and (Num <> 16) and (Num <> 28) and (Num <> 900) then
  begin
    if Action = taModif then
    begin
      Action := taConsult;
    end else
    if Action = taConsult then
    begin
      // action interdite
      PGIBox(TexteMessage[9], TitreHalley);
      Exit;
    end;
  end;

  if ((Num = 8) and (TT = 'GCTIERSSAISIE')) then
  begin
    ii := TTToNum(TT);
    if (ii > 0) and (pos('T_NATUREAUXI="FOU"', V_PGI.DECombos[ii].Where) > 0) then num := 12;
  end;
  case Num of
    {$IFNDEF GCGC}
    1: {Compte gene} FicheGene(nil, '', LeQuel, Action, 0);
    2: {Tiers compta} FicheTiers(nil, '', LeQuel, Action, 1);
    4: {Journal} FicheJournal(nil, '', Lequel, Action, 0);
    {$ENDIF}
    5:
      begin {Affaires}
        Arg5LanceFiche := ActionToString(Action);
        if (Range <> '') then Arg5LanceFiche := Arg5LanceFiche + ';' + Range;
        if (lequel = '') then lequel := 'AFF';
        AGLLanceFiche('AFF', 'AFFAIRE', '', Lequel, Arg5LanceFiche);
      end;
    6:
      begin {Ressources}
        Arg5LanceFiche := ActionToString(Action);
        if (Range <> '') then Arg5LanceFiche := Arg5LanceFiche + ';' + Range;
        AGLLanceFiche('AFF', 'RESSOURCE', '', Lequel, Arg5LanceFiche);
      end;
    7: {Article} DispatchTTArticle(Action, Lequel, TT, Range);
    8: {Clients par T_TIERS}
      begin
        Lequel := TiersAuxiliaire(Lequel, false);
        Arg5LanceFiche := ActionToString(Action) + ';MONOFICHE';
        if Range = '' then Arg5LanceFiche := Arg5LanceFiche + ';T_NATUREAUXI=CLI' else Arg5LanceFiche := Arg5LanceFiche + Range;
        if FOExisteClavierEcran then
          AGLLanceFiche('MFO', 'GCTIERSFO', '', Lequel, Arg5LanceFiche)
        else
          AGLLanceFiche('GC', 'GCTIERS', '', Lequel, Arg5LanceFiche);
      end;
    //9 : {Commerciaux} AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)+';GCL_TYPECOMMERCIAL="REP"') ;
    9: {Commerciaux} AGLLanceFiche('MFO', 'FOCOMMERCIAL', '', Lequel, ActionToString(Action) + ';GCL_TYPECOMMERCIAL="VEN"');
    10: {Conditionnement} AGLLanceFiche('GC', 'GCCONDITIONNEMENT', Range, Lequel, ActionToString(Action));
    //11 : {Catalogue} AGLLAnceFiche('GC','GCCATALOGU_NVFOUR','',Lequel,ActionToString(Action)) ;
    11: {Catalogue} AGLLAnceFiche('GC', 'GCCATALOGU_SAISI3', '', Lequel, ActionToString(Action));
    12: {Fournisseurs via T_TIERS}
      begin
        Lequel := TiersAuxiliaire(Lequel, false);
        AGLLanceFiche('GC', 'GCFOURNISSEUR', '', Lequel, ActionToString(Action) + ';MONOFICHE;T_NATUREAUXI=FOU');
      end;
    13: {Dépots} AGLLanceFiche('GC', 'GCDEPOT', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    14: {Apporteurs} AGLLanceFiche('GC', 'GCCOMMERCIAL', '', Lequel, ActionToString(Action) + ';GCL_TYPECOMMERCIAL="APP"');
    15: {Règlement} FicheRegle_AGL('', False, taModif);
    16: {contacts} AGLLanceFiche('YY', 'YYCONTACT', Lequel, Range, ActionToString(Action) + ';MONOFICHE');
    17: {Fonction des ressources} AGLLanceFiche('AFF', 'FONCTION', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    18: {Etablissements}
      if (CtxMode in V_PGI.PGIContexte) then
        AGLLanceFiche('MBO', 'ETABLISS', '', Lequel, ActionToString(Action) + ';MONOFICHE')
      else
        AGLLanceFiche('GC', 'GCETABLISS', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    19: {Contenu tables libres} AGLLanceFiche('GC', 'GCCODESTATPIECE', '', '', 'ACTION=CREATION;YX_CODE=' + Range);
    20: ;
    21: {Ports} AGLLanceFiche('GC', 'GCPORT', '', Lequel, ActionToString(Action) + ';MONOFICHE');

    22: {GRC-action} AGLLanceFiche('RT', 'RTACTIONS', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    23: {GRC-operation} AGLLanceFiche('RT', 'RTOPERATIONS', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    24: {GRC-actionsChainées}
      begin
        Arg5LanceFiche := ActionToString(Action) + ';MONOFICHE;' + Range;
        AGLLanceFiche('RT', 'RTACTIONSCHAINE', '', Lequel, Arg5LanceFiche);
      end;

    28: {Clients par T_AUXILIAIRE}
      begin
        Arg5LanceFiche := ActionToString(Action) + ';MONOFICHE';
        if Range = '' then Arg5LanceFiche := Arg5LanceFiche + ';T_NATUREAUXI=CLI' else Arg5LanceFiche := Arg5LanceFiche + Range;
        if (CtxMode in V_PGI.PGIContexte) then
          AGLLanceFiche('MFO', 'GCTIERSFO', '', Lequel, Arg5LanceFiche)
        else
          AGLLanceFiche('GC', 'GCTIERS', '', Lequel, Arg5LanceFiche);
      end;
    29: {Fournisseurs via T_AUXILIAIRE} AGLLanceFiche('GC', 'GCFOURNISSEUR', '', Lequel, ActionToString(Action) + ';MONOFICHE;T_NATUREAUXI=FOU');
    30: {GRC-Projet} AGLLanceFiche('RT', 'RTPROJETS', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    31:
      begin
        if Action = taConsult then Action := taModif;
        AGLLanceFiche('MBO', 'EXPORT', Range, Range + ';' + Lequel, ActionToString(Action) + ';MONOFICHE');
      end;

    112200: {Lancement de l'éditeur de ticket} FODispatchParamModele(Num, Action, Lequel, TT, Range);

    // *** maj des tablettes ***
    900: ParamTable('GCCOMMENTAIRELIGNE', taCreat, 0, nil, 6); // mis de manière explicite car appellé depuis un Grid

    // pour mise à jour par defaut des tablettes
    994: ParamTable(TT, taModif, 0, nil, 9); {choixext}
    995: ParamTable(TT, taCreat, 0, nil, 9); {choixext}
    996: ParamTable(TT, taModif, 0, nil, 6); {choixext}
    997: ParamTable(TT, taCreat, 0, nil, 6); {choixext}
    998: ParamTable(TT, taModif, 0, nil, 3); {choixcode ou commun}
    999: ParamTable(TT, taCreat, 0, nil, 3); {choixcode ou commun}
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : AfterChangeModule
Mots clefs ... :
*****************************************************************}

procedure AfterChangeModule(NumModule: integer);
begin
  ////AjusteMenu(NumModule);
  case NumModule of
    107:
      begin
        if FOComptaEstActive(FOGetNatureTicket(False, False)) then
        begin
          FMenuG.RemoveItem(107150); // Modification des règlements
          FMenuG.RemoveItem(107180); // Rattachement d'un ticket à un client
        end;
        {$IFNDEF TOXCLIENT}
        FMenuG.RemoveItem(107260); // Retansmission d'une journée
        {$ENDIF}

        {$IFDEF EAGLCLIENT}
        FMenuG.RemoveItem(107350); // Lanceur d'états libres utilisateurs
        {$ENDIF}

        FMenuG.RemoveGroup(107900, True); // Droits d’accès
      end;
    108:
      begin
        {$IFDEF MODES3}
        FMenuG.RemoveGroup(108300, True); // Commandes de réassort
        FMenuG.RemoveItem(108125); // Stock distant
        FMenuG.RemoveItem(108130); // Statistique Stock disponible
        FMenuG.RemoveItem(-108140); // Etiquettes articles
        FMenuG.RemoveItem(108146); // Etiquettes articles sur transfert
        FMenuG.RemoveGroup(108400, True); // Réceptions
        FMenuG.RemoveGroup(108500, True); // Inventaire
        {$ENDIF}

        {$IFDEF EAGLCLIENT}
        FMenuG.RemoveItem(108125); // Stock distant
        FMenuG.RemoveItem(108591); // Intégration du fichier inventoriste
        {$ENDIF}
      end;
    109:
      begin
        {$IFDEF MODES3}
        FMenuG.RemoveGroup(109200, True); // Edition des étiquettes clients
        {$ENDIF}
      end;
    111:
      begin
        FMenuG.RemoveItem(111230); // Liaison FO-BO - Echanges - Etat de consolidation
        FMenuG.RemoveItem(111240); // Liaison FO-BO - Echanges - Etat des intégrations
        FMenuG.RemoveItem(111330); // Génération des BE
      end;
    112:
      begin
        {$IFDEF MODES3}
        FMenuG.RemoveItem(112442); // Personnalisation
        FMenuG.RemoveItem(112433); // Restrictions utilisateurs
        FMenuG.RemoveItem(112434); // Utilisateurs connectés
        FMenuG.RemoveItem(112435); // Journal des évènements
        FMenuG.RemoveItem(112436); // Suivi d'activité
        {$ENDIF}

        {$IFDEF EAGLCLIENT}
        FMenuG.RemoveItem(112115); // Modification en série du paramétrage des caisses
        FMenuG.RemoveItem(112140); // Démarques
        FMenuG.RemoveItem(112180); // Détail des espèces
        FMenuG.RemoveItem(112190); // Devises
        FMenuG.RemoveItem(112510); // Compteurs
        FMenuG.RemoveItem(112522); // Modèles d'édition document
        FMenuG.RemoveItem(112523); // Modèles d'édition état
        FMenuG.RemoveItem(112524); // Listes de saisie
        FMenuG.RemoveGroup(112200, True); // Modèles d'impression texte
        FMenuG.RemoveGroup(112300, True); // Récupérations
        //FMenuG.RemoveGroup(112400, True) ;     // Administration
        FMenuG.RemoveGroup(112420, True); // Société
        FMenuG.RemoveItem(-112420); // Société
        FMenuG.RemoveItem(112421); // Paramètres société
        FMenuG.RemoveItem(112422); // Gestionnaire société
        FMenuG.RemoveItem(112423); // Recopie société à société
        FMenuG.RemoveItem(112431); // Groupes utilisateurs
        FMenuG.RemoveItem(112432); // Utilisateurs
        FMenuG.RemoveItem(112433); // Restrictions utilisateurs
        FMenuG.RemoveItem(112434); // Utilisateurs connectés
        FMenuG.RemoveItem(112435); // Journal des évènements
        FMenuG.RemoveItem(112436); // Suivi d'activité
        FMenuG.RemoveItem(112437); // Remise à zéro connexions
        FMenuG.RemoveItem(112438); // Gestion droits d'accès
        FMenuG.RemoveGroup(112440, True); // Multilingue
        FMenuG.RemoveItem(-112440); // Multilingue
        FMenuG.RemoveItem(112441); // Traduction
        FMenuG.RemoveItem(112442); // Personnalisation
        {$ENDIF}

        if GetParamSoc('SO_GCFOBASEFORMATION') then
          FMenuG.RemoveItem(112330); // Génération de la base de formation
      end;
  end;

  // Menu Pop général
  case NumModule of
    30, 31, 32, 33, 34, 65, 70:
      ChargeMenuPop(5, FMenuG.DispatchX);
    107,108,109,111,112 : ChargeMenuPop(25,FMenuG.DispatchX) ;
  else
    ChargeMenuPop(NumModule, FMenuG.DispatchX);
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : ChargeFavoris
Mots clefs ... :
*****************************************************************}

function ChargeFavoris: boolean;
var lesTagsToRemove: string;
begin
  Result := True;
  lesTagsToRemove := DoTagToRemove;
  AddGroupFavoris(FMenuG, lesTagsToRemove);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Changement d'utilisateur
Mots clefs ... :
*****************************************************************}

{$IFNDEF EAGLCLIENT}
function ChangeUser: boolean;
begin
  Result := True;
  ChargeProfilUser;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : AfterProtec
Mots clefs ... :
*****************************************************************}

procedure AfterProtec(sAcces: string);
var FOSeria: boolean;
begin
  FOSeria := False;
  if Length(sAcces) > 0 then FOSeria := (sAcces[1] = 'X'); // Front Office Mode sérialisé

  if V_PGI.NoProtec then
  begin
    V_PGI.VersionDemo := True;
    FOSeria := False;
  end
  else
    V_PGI.VersionDemo := False;

  if (FOSeria = False) then
    V_PGI.VersionDemo := True; // on repasse en demo dans ce cas

  if (V_PGI.VersionDemo) and (Pos(' (DEMO)', TitreHalley) = 0) then
    TitreHalley := TitreHalley + ' (DEMO)';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Cette procédure permet d'initiliaser certaines références de
Suite ........ : fonction, les modules des menus gérés par l'application, ...
Suite ........ :
Suite ........ : Cette procédure est appelée directement dans le source du
Suite ........ : projet.
Mots clefs ... : INITILISATION
*****************************************************************}

procedure InitApplication;
begin
  ProcZoomEdt := GCZoomEdtEtat;
  ProcCalcEdt := GCCalcOLEEtat;
  ProcCalcMul := PGIProcCalcMul; //FOCalcMul;
  //ProcGetVH := GetVS1 ;
  //ProcGetDate := GetDate ;
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
  {$IFDEF EAGLCLIENT}
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
  {$ELSE}
  FMenuG.OnChangeUser := ChangeUser;
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
  FMenuG.OnMajPendant := nil;
  {$ENDIF}
  FMenuG.OnAfterProtec := AfterProtec;

  {$IFDEF MODES3}
  FMenuG.SetSeria('30330010', ['30310050'], ['CEGID Front Office S3']);
  {$ELSE}
  FMenuG.SetSeria('30330010', ['30320050'], ['CEGID Front Office S5']);
  {$ENDIF}

  {$IFDEF TOXCLIENT}
  FMenuG.SetModules([107, 108, 109, 111, 112], [94, 111, 78, 69, 74]);
  {$ELSE}
  FMenuG.SetModules([107, 108, 109, 112], [94, 111, 78, 74]);
  {$ENDIF}
  //V_PGI.NbColModuleButtons := 2 ;

  FMenuG.OnChangeModule := AfterChangeModule;
  FMenuG.OnChargeFavoris := ChargeFavoris;
  FMenuG.SetPreferences(['Pièces', '', '', '', 'Tickets'], False);
  V_PGI.DispatchTT := DispatchTT;
end;

end.
