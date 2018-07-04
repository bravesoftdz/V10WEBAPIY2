{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2003
Modifié le ... : 23/08/2005 - GCO - FQ 16460
Description .. : Source TOF de la FICHE : POINTAGEECR ()
Suite ........ : GCO - 16/09/2005 - FQ 16460
Mots clefs ... : TOF;POINTAGEECR
*****************************************************************}
unit uTOFPointageEcr;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  Maineagl,     // AGLLanceFiche EAGL
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_main,      // AGLLanceFiche AGL
{$ENDIF}
{$IFDEF VER150}
   Variants,
 {$ENDIF}
  uTofViergeMul, // Fiche Vierge MUL
  sysutils,
  ComCtrls,
  HQry,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  Ent1,            // VH^.
  Graphics,        // TCanvas
  Grids,           // TGridDrawState
  Windows,         // TheRect, VK_
  uLibWindows,     // IIF
  ParamSoc,        // GetParamSocSecur
  uTob,            // TOB
  Saiscomm,        // GetO
  SaisUtil,        // RMVT
  Htb97,           // TToolBarButton97
  utilPGI,         // _Blocage
  HSysMenu,        // HSystemMenu
  Menus,           // TPopUpMenu
  utilSais,        // CUpdateCumulsPointeMS
  Forms;

type
  TOF_POINTAGEECR = class(TOF_ViergeMul)

    E_DateComptable_: THEdit;
    BSaisie         : TToolBarButton97;
    BValideEtRappro : TToolBarButton97;

    procedure OnNew                  ; override;
    procedure OnDelete               ; override;
    procedure OnUpdate               ; override;
    procedure OnLoad                 ; override;
    procedure OnArgument(S: string)  ; override;
    procedure OnDisplay              ; override;
    procedure OnClose                ; override;
    procedure OnCancel               ; override;
    procedure OnClickBSelectAll      (Sender: TObject); override;
    procedure OnClickBValider        (Sender: TObject); override;
    procedure OnClickBSaisie         (Sender: TObject);
    procedure OnClickBValideEtRappro (Sender: TObject);
    procedure OnKeyDownFListe        (Sender: TObject; var Key: Word; Shift: TShiftState); override;
    procedure OnKeyDownEcran         (Sender: TObject; var Key: Word; Shift: TShiftState); override;
    procedure OnExitE_DateComptable_ (Sender: TObject);
    procedure OnMouseDownFListe      (Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure OnDblClickFListe       (Sender: TObject);
    procedure OnResizeEcran          (Sender: TObject);
    procedure GetCellCanvasFListe    (ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure OnClickBImprimer       (Sender: TObject); override;

  protected
    procedure RemplitATobFListe;                         override;
    procedure InitControl;                               override; // Init des composants de la fiche
    procedure RefreshFListe( vBoFetch : Boolean );       override; // Recharge la grille avec le PutGridDetail
    function  BeforeLoad: boolean;                       override;
    function  AfterLoad : boolean;                       override;

  private
    FStDevise : string;            // Devise du compte général ou du compte de contrepartie du journal
    FStEE_General     : string;    // Attention ! peut contenir un GENERAL ou un JOURNAL
    FStEE_RefPointage : string;    // Contient la Réference de Pointage à utiliser
    FDtEE_DatePointage: TDatetime; // Contient la Date de Pointage
    FItEE_Numero : integer;        // Contien le numero interne du relevé
    FBoReadOnly : boolean;         // Si pointage auto : readonly

    FBoFaireRequete   : Boolean;   // Faut il faire la requete SQL?
    FBoAllSelected    : Boolean;
    FBoDroitEcritures : Boolean;   // Droit de créer ou de modifier des écritures
    FBoCptEnDevise    : Boolean;   // Afficahge des montants en DEVISE
    FBoAffichePointe  : Boolean;   // Affichage des écritures pointées
    FBoConsultation   : Boolean;   // Blocage des fonctions de Pointage ou Dépointage
    FBoSecurise       : Boolean;   // Pointage Securise OUI ou NON

    FTotDebPointeRef  : Double;   // Total Crédit des mvts pointés de la référence
    FTotCrePointeRef  : Double;   // Total Débit  des mvts pointés de la référence
    FTotNbPointeRef   : Integer;  // Nombre d'écritures pointées de la référence

    // Totaux de pointage du compte général
    FTotDebPTP, FTotDebPTD : Double;
    FTotCrePTP, FTotCrePTD : Double;

    FTotDebDejaPointe      : Double;  // Total Débit des des mvts pointés
    FTotCreDejaPointe      : Double;  // Total Crédit des mvts pointés
    FTotNbDejaPointe       : Integer; // Nombre d'écritures pointées

    FStLeDebit, FStLeCredit: string; // E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV

    // Variables contentant le numéro des colonnes imortantes de la grille
    FColRefPointage, FColSoldePro, FColDatePointage : integer;
    FColLeDebit, FColLeCredit : integer;

    FTobGene: TOB; // TOB sur les GENERAUX
    FTobSave: TOB; // TOB pour sauvegarder le pointage en cours
    FTobRefPointage: TOB; // TOB sur EEXBQ

    ////////////////////////////////////////////////////////////////////////////
    procedure IndiceColFListe;
    procedure PointeEcriture(vRow: integer; vOkPointe: Boolean);
    procedure AnnuleDerniereOperation;
    procedure DepointeTout;
    procedure CacheEcriturepointee;
    procedure AfficheEcriturepointee;
    procedure PasseLibelleSuivant(vStLibelle: string = '');
    procedure PasseSelectionSuivante;
    procedure PasseJourSuivant;
    procedure PasseMoisSuivant;
    procedure PassePrecedent;
    procedure PasseSuivant;
    procedure SelectionRapide;
    procedure SelectionJour;
    procedure SelectionMois;
    procedure TraiteDeviseDuCompte;
    procedure CalculAvancement;
    procedure AligneControleSurGrille(vGrille: THGrid; vCol: integer; vNomControl: string; vAvecWidth : Boolean = False);
    procedure SauvePointageEnCours;
    procedure ChargePointageEnCours;
    procedure ValidationPointage;    // GCO - 26/21/2005

    // Procédure pour appeler par le POPF11
    procedure OnClickAnnuleDerniereOperation ( Sender : TObject );
    procedure OnClickDepointeTout            ( Sender : TObject );
    procedure OnClickCacheEcriturepointee    ( Sender : TObject );
    procedure OnClickAfficheEcriturepointee  ( Sender : TObject );
    procedure OnClickPasseLibelleSuivant     ( Sender : TObject );
    procedure OnClickPasseSelectionSuivante  ( Sender : TObject );
    procedure OnClickPasseJoursuivant        ( Sender : TObject );
    procedure OnClickPasseMoisSuivant        ( Sender : TObject );
    procedure OnClickPassePrecedent          ( Sender : TObject );
    procedure OnClickPasseSuivant            ( Sender : TObject );
    procedure OnClickSelectionRapide         ( Sender : TObject );
    procedure OnClickSelectionJour           ( Sender : TObject );
    procedure OnClickSelectionMois           ( Sender : TObject );

    ////////////////////////////////////////////////////////////////////////////
    function RecupAutreWhere: string;
    function GereAffichageEcriture(vTobEcriture: Tob): string;
    function ControleRefPointage  : Boolean;
    function ControleContrepartie : Boolean;
    function ControleDevise       : Boolean;

  end;

  /////////////////////////////////////////////////////////////////////
function CPLanceFiche_Pointage(vStParam: string = ''): string;
/////////////////////////////////////////////////////////////////////

//--------------------- Clé primaire Table ECRITURE --------------------------
// E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE
//----------------------------------------------------------------------------

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  UlibPointage,
  {$ENDIF MODENT1}
  Saisie,          // TrouveSaisie
  SaisBor,         // LanceSaisieFolio
  uTofCPMulMvt,    // MultiCritereMvt
  uTofPointageMul, // CDepointeEcriture
  CPRapproDet_Tof; // CC_LanceFicheEtatRapproDet

const
  cAutreSelect = 'E_JOURNAL CLE_JOURNAL, E_EXERCICE CLE_EXERCICE, ' +
    'E_DATECOMPTABLE CLE_DATECOMPTABLE, E_NUMEROPIECE CLE_NUMEROPIECE, ' +
    'E_NUMLIGNE CLE_NUMLIGNE, E_NUMECHE CLE_NUMECHE, ' +
    'E_QUALIFPIECE CLE_QUALIFPIECE, E_DATEMODIF CLE_DATEMODIF, ' +
    'E_DATEECHEANCE CLE_DATEECHEANCE ';

////////////////////////////////////////////////////////////////////////////////

function CPLanceFiche_Pointage(vStParam: string = ''): string;
begin
  // if _Blocage(['nrCloture'], False, 'nrPointage') then Exit;
  AGLLanceFiche('CP', 'CPPOINTAGEECR', '', '', vStParam);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnArgument(S: string);

  procedure _RecupereArgument(vS: string);
  var
    lStArgumentTOF: string;
    lSt: string;
  begin
    lStArgumentTOF := vS;

    // Compte GENERAL ou JOURNAL
    lSt := ReadTokenSt(lStArgumentTOF);
    FStEE_General := IIF(lSt <> '', lSt, '');

    // Date de Pointage
    lSt := ReadTokenSt(lStArgumentTOF);
    FDtEE_DatePointage := StrToDate(IIF(lSt <> '', lSt, ''));

    // Référence de POINTAGE
    lSt := ReadTokenSt(lStArgumentTOF);
    FStEE_RefPointage := IIF(lSt <> '', lSt, '');

    // numero interne du relevé
    lSt := ReadTokenSt(lStArgumentTOF);
    FItEE_Numero := StrToInt(IIF(lSt <> '', lSt, '1'));

    // mode consultation ??
    lSt := ReadTokenSt(lStArgumentTOF);
    FBoReadOnly := (lSt = 'ACTION=CONSULTATION');
  end;

  function _RecupereCompte: string;
  var
    lQuery: TQuery;
  begin
    Result := '';
    try
      lQuery := nil;
      try
        lQuery := OpenSql('SELECT G_GENERAL FROM GENERAUX LEFT JOIN JOURNAL ' +
          'ON J_CONTREPARTIE = G_GENERAL WHERE J_JOURNAL = "' + FStEE_General + '"', True);

        Result := lQuery.FindField('G_GENERAL').AsString;
      except
        on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction _RecupereCompte');
      end;

    finally
      Ferme(lQuery);
    end;
  end;

  procedure _ChargeRefPointage;
  var
    lQuery: TQuery;
  begin
    try
      lQuery := nil;
      try
        lQuery := OpenSql('SELECT * FROM EEXBQ WHERE EE_GENERAL = "' + FStEE_General + '"' +
                          ' AND EE_REFPOINTAGE = "' + FStEE_RefPointage + '"' +
                          ' AND EE_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '"' +
                          ' AND EE_NUMERO = ' + IntToStr(FItEE_Numero), True);

        if not lQuery.Eof then
        begin
          if FBoCptEnDevise then
          begin
            SetControlText('TDSOLDEBANCAIRE', StrfMontant(lQuery.FindField('EE_NEWSOLDEDEB').AsFloat, 15, V_PGI.OkDecV, '', True));
            SetControlText('TCSOLDEBANCAIRE', StrfMontant(lQuery.FindField('EE_NEWSOLDECRE').AsFloat, 15, V_PGI.OkDecV, '', True));
          end
          else
          begin
            SetControlText('TDSOLDEBANCAIRE', StrfMontant(lQuery.FindField('EE_NEWSOLDEDEBEURO').AsFloat, 15, V_PGI.OkDecV, '', True));
            SetControlText('TCSOLDEBANCAIRE', StrfMontant(lQuery.FindField('EE_NEWSOLDECREEURO').AsFloat, 15, V_PGI.OkDecV, '', True));
          end;
        end;

      except
        on E: Exception do
          PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : _ChargeRefPointage');
      end;

    finally
      Ferme(lQuery);
    end;
  end;

begin
  FStListeParam := 'CPPOINTAGEECR';
  FBoFaireRequete := True;
  FBoAllSelected := False;
  FBoDroitEcritures := ExJaiLeDroitConcept(TConcept(ccSaisEcritures), False);
  FBoAffichePointe := True;
  inherited;

  // Interdit à la fiche ancetre de faire l'affichage dans la grille
  FBoOkAncetreFaitAffichage := False;

  // Pointage sécurisé OUI ou NON
  FBoSecurise := (GetParamSocSecur('SO_CPPOINTAGESECU', False) = True);

  // Decoupe l'argument passé à la TOF et charge la variable globale
  _RecupereArgument(S);

  // Création de la TOB sur les GENERAUX et récuperation des infos du compte
  FTobGene := Tob.Create('GENERAUX', nil, -1);
  if not VH^.PointageJal then
    FTobGene.SelectDB('"' + FStEE_GENERAL + '"', nil, false)
  else
    FTobGene.SelectDB('"' + _RecupereCompte + '"', nil, false);
  // FQ 17073 : Pb relecture des totaux piontés des comptes en multisoc entreprise SBO 17/01/2007
  if EstTablePartagee('GENERAUX') then
    CChargeCumulsMS( fbGene, FTobGene.GetString('G_GENERAL'), '', FTobGene ) ;

  // TOB pour sauvegarder le pointage en cours avant passage à la saisie BOR
  FTobSave := Tob.Create('', nil, -1);

  // récupération des totaux pointés du compte général
  FTotDebPTP := FTobGene.GetValue('G_TOTDEBPTP');
  FTotCrePTP := FTobGene.GetValue('G_TOTCREPTP');
  FTotDebPTD := FTobGene.GetValue('G_TOTDEBPTD');
  FTotCrePTD := FTobGene.GetValue('G_TOTCREPTD');

  // Création TOb sur EEXBQ pour récuperation des infos de la Ref de Pointage
  FTobRefPointage := Tob.Create('', nil, -1);

  TraiteDeviseDuCompte;
  FStLeDebit := IIF(FBoCptEnDevise, 'E_DEBITDEV', 'E_DEBIT');
  FStLeCredit := IIF(FBoCptEnDevise, 'E_CREDITDEV', 'E_CREDIT');

  // Charge la référence de pointage et affiche le solde bancaire
  _ChargeRefPointage;

  /////////////////////////////////////////////////////////////////
  E_DateComptable_        := THEdit(GetControl('E_DATECOMPTABLE_', True));
  E_DateComptable_.Text   := DateToStr(FDtEE_DatePointage);
  E_DateComptable_.OnExit := OnExitE_DateComptable_;
  BSaisie                 := TToolBarButton97(GetControl('BSAISIE', True));
  BValideEtRappro         := TToolBarButton97(GetControl('BVALIDEETRAPPRO', True));

  Ecran.OnResize          := OnResizeEcran;
  BSaisie.OnClick         := OnClickBSaisie;
  BCherche.OnClick        := OnClickBCherche;
  BSelectAll.OnClick      := OnClickBSelectAll;
  BValider.OnClick        := OnClickBValider;
  BValideEtRappro.OnClick := OnClickBValideEtRappro;

  FListe.OnKeyDown        := OnKeyDownFListe;
  FListe.OnMouseDown      := OnMouseDownFListe;
  FListe.OnDblClick       := OnDblClickFListe;
  FListe.GetCellCanvas    := GetCellCanvasFListe;

  // Popup Menu sur F11
  // Débranchement de l'événement de la fiche ancêtre
  PopF11.OnPopup := nil ;
  POPF11.Items[0].OnClick  := OnDblClickFListe;
  POPF11.Items[1].OnClick  := OnClickBSelectAll;
  POPF11.Items[2].OnClick  := OnClickBSaisie;
  POPF11.Items[3].OnClick  := OnClickAnnuleDerniereOperation;
  POPF11.Items[4].OnClick  := OnClickDepointeTout;
  POPF11.Items[5].OnClick  := OnClickBValider;
  POPF11.Items[6].OnClick  := OnClickBValideEtRappro;
  POPF11.Items[7].OnClick  := OnClickSelectionRapide;
  POPF11.Items[8].OnClick  := OnClickSelectionJour;
  POPF11.Items[9].OnClick  := OnClickSelectionMois;
  POPF11.Items[10].OnClick := OnClickPasseSelectionSuivante;
  POPF11.Items[11].OnClick := OnClickPasseSuivant;
  POPF11.Items[12].OnClick := OnClickPassePrecedent;
  POPF11.Items[13].OnClick := OnClickPasseLibelleSuivant;
  POPF11.Items[14].OnClick := OnClickPasseMoisSuivant;
  POPF11.Items[15].OnClick := OnClickPasseJourSuivant;
  POPF11.Items[16].OnClick := OnClickCacheEcriturepointee;
  POPF11.Items[17].OnClick := OnClickAfficheEcriturepointee;

  // Recupere les numeros de colonne des champs importants
  IndiceColfListe;

  //
  AFiltreDisabled := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnLoad;
var
  lSt: string;
begin
  inherited;
  lSt := IIF(not VH^.PointageJal, 'Compte général', 'Journal');

  Ecran.Caption := lSt + ' ' + FStEE_General + ' - Pointage au ' +
                   DateToStr(FDtEE_DatePointage) + ' (' + FStEE_RefPointage + ')';

  UpDateCaption(Ecran);

  SetControlText('TRECAPITULATIF', 'Pour la référence ' + FStEE_RefPointage +
                 ' au ' + DateToStr(FDtEE_DatePointage));

  SetControlText('TAVANCEMENT', 'Reste à pointer');

  FBoConsultation := (FBoReadOnly or CEstPointageEnConsultationSurDossier);
  // Si Le dossier autorise les modifs de pointage, Test du pointage lui même
  if not FBoConsultation then
  begin
    if FBoSecurise then
      FBoConsultation := ControleRefPointage
    else
      FBoConsultation := False;
  end;

  // Vérouillage des fonctions de pointage si on est en mode Consultation
  if FBoConsultation then
  begin
    SetControlProperty('TCONSULTATION', 'CAPTION', 'Modification impossible...');
    BValider.Enabled := False;
    BValideEtRappro.Enabled := False;
    POPF11.Items[3].Enabled := False;
    POPF11.Items[4].Enabled := False;
    POPF11.Items[5].Enabled := False;
    POPF11.Items[6].Enabled := False;
  end
  else
    SetControlProperty('TCONSULTATION', 'CAPTION', '');

  SetControlProperty('TCONTREPARTIE', 'CAPTION', '');
  if (VH^.PointageJal) and (ControleContrepartie) then
    SetControlProperty('TCONTREPARTIE', 'CAPTION', 'Présence d''écritures sur un autre journal');

  SetControlProperty('TDEVISE', 'CAPTION', '');

  // FB 12782 - GCO - 13/10/2003
  if FStDevise <> V_PGI.DevisePivot then
  begin
    if ControleDevise then
     SetControlProperty('TDEVISE', 'CAPTION', 'Présence d''écritures saisies dans une autre devise');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnDisplay();
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnNew;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnDelete;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnUpdate;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClose;
begin
  // _Bloqueur('nrPointage',False) ;
  FreeAndNil(FTobGene);
  FreeAndNil(FTobRefPointage);
  FreeAndNil(FTobSave);
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnCancel();
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_POINTAGEECR.BeforeLoad: boolean;
begin
  inherited BeforeLoad;
  // Modification des colonnes de la grille en fonction des comptes
  IndiceColFListe;

  Result := FBoFaireRequete;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_POINTAGEECR.AfterLoad: boolean;
begin
  Result := inherited AfterLoad;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.InitControl;
begin
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/09/2002
Modifié le ... :   /  /
Description .. : Récupère d'autres paramètres Where pour la requete sur les
Suite ........ : écritures
Mots clefs ... :
*****************************************************************}
function TOF_POINTAGEECR.RecupAutreWhere: string;
var
  lSt: string;
begin
  if not VH^.PointageJal then
    lSt := ' AND E_GENERAL = "' + FStEE_General + '"'
  else
    lSt := ' AND E_GENERAL <> "' + FTobGene.GetValue('G_GENERAL') + '" AND ' +
      ' E_JOURNAL = "' + FStEE_General + '"';

  lSt := lSt + 'AND E_QUALIFPIECE = "N"';

  lSt := lSt + ' AND ((E_REFPOINTAGE = "" AND E_DATEPOINTAGE = "' +
    UsDateTime(iDate1900) + '") OR ' +
    '(E_REFPOINTAGE = "' + FStEE_RefPointage + '" AND ' +
    ' E_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '" ) OR ' +
    '(E_DATECOMPTABLE <= "' + UsDateTime(FDtEE_DatePointage) + '" AND ' +
    'E_DATEPOINTAGE  > "' + UsDateTime(FDtEE_DatePointage) + '"))';

  // FB 12782 - GCO - 13/10/2003
  if FStDevise <> V_PGI.DevisePivot then
    lSt := lSt + ' AND E_DEVISE = "' + FStDevise + '"';

  lSt := lSt + ' AND (E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H") AND E_TRESOLETTRE = "-"';
  lSt := lSt + ' AND ((E_DEBIT <> 0) OR (E_CREDIT <> 0)) AND E_CREERPAR <> "DET"';

  try
    // Gestion du VH^.ExoV8
    if VH^.ExoV8.Code <> '' then
      lSt := lSt + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '" ';
  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Erreur ExoV8');
  end;

  Result := lSt;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... :   /  /
Description .. : Génère la requete SQL en fonction des crtières
Suite ........ : Charge la grille avec le résultat de la requête
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.RemplitATobFListe;
var
  lQuery          : TQuery;
  lStListeChamps  : string;
  lSoldeComptable : Double;
  lStSqlWhereTobFListe : string;
  lSt : string;
begin
  if not FBoFaireRequete then Exit;
  FTotDebDejaPointe := 0;
  FTotCreDejaPointe := 0;
  FTotNbDejaPointe  := 0;

  try
    try
      // Solde DEBIT, CREDIT de tout les des mvts déjà pointés
      if not VH^.PointageJal then
      begin
        lQuery := OpenSql('SELECT COUNT(*) NBDEJAPOINTE, SUM(' + FStLeDebit + ') TOTDEB, '+
                          'SUM(' + FStLeCredit + ') TOTCRE FROM ECRITURE WHERE ' +
                          'E_GENERAL = "' + FTobGene.GetValue('G_GENERAL') + '" AND ' +
                          '(E_QUALIFPIECE = "N") AND ' +
                          '(E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H") AND ' +
                          '(E_REFPOINTAGE = "' + FStEE_RefPointage + '" AND ' +
                          'E_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '" OR ' +
                          '(E_DATEPOINTAGE > "' + UsDateTime(FDtEE_DatePointage) + '"))', True);
      end
      else
      begin
        lQuery := OpenSql('SELECT COUNT(*) NBDEJAPOINTE, SUM(' + FStLeDebit + ') TOTDEB, ' +
                          'SUM(' + FStLeCredit + ') TOTCRE FROM ECRITURE WHERE ' +
                          'E_JOURNAL = "' + FStEE_General + '" AND ' +
                          'E_GENERAL <> "' + FTobGene.GetValue('G_GENERAL') + '" AND ' +
                          '(E_QUALIFPIECE = "N") AND ' +
                          '(E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H") AND ' +
                          '(E_REFPOINTAGE = "' + FStEE_RefPointage + '" AND ' +
                          'E_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '" OR ' +
                          '(E_DATEPOINTAGE > "' + UsDateTime(FDtEE_DatePointage) + '"))', True);
      end;

      // Total DEBIT Pointés de la requête SQL Principale
      FTotDebDejaPointe := lQuery.Findfield('TOTDEB').AsFloat;
      // Total CREDIT Pointés de la requête SQL Principale
      FTotCreDejaPointe := lQuery.Findfield('TOTCRE').AsFloat;
      // Nombre d'écritures déjà Pointées
      FTotNbDejaPointe  := lQuery.Findfield('NBDEJAPOINTE').AsInteger;
    except
      on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, ' Total Déjà Pointé');
      end;
    end;
  finally
    Ferme( lQuery );
  end;

  // Solde DEBIT, CREDIT des mvts pointés de la référence
  try
    FTotDebPointeRef  := 0;
    FTotCrePointeRef  := 0;
    FTotNbPointeRef   := 0;
    try
      if not VH^.PointageJal then
      begin
         lSt := 'SELECT COUNT(*) NBDEJAPOINTE, SUM(' + FStLeDebit + ') TOTDEB, '+
                'SUM(' + FStLeCredit + ') TOTCRE FROM ECRITURE WHERE ' +
                'E_GENERAL = "' + FTobGene.GetValue('G_GENERAL') + '" AND ' +
                'E_REFPOINTAGE = "' + FStEE_RefPointage + '" AND ' +
                'E_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '" AND ' +
                '(E_ECRANOUVEAU = "H" OR E_ECRANOUVEAU = "N")';
      end
      else
      begin
        lSt := 'SELECT COUNT(*) NBDEJAPOINTE, SUM(' + FStLeDebit + ') TOTDEB, ' +
               'SUM(' + FStLeCredit + ') TOTCRE FROM ECRITURE WHERE ' +
               'E_JOURNAL = "' + FStEE_General + '" AND ' +
               'E_GENERAL <> "' + FTobGene.GetValue('G_GENERAL') + '" AND ' +
               'E_REFPOINTAGE = "' + FStEE_RefPointage + '" AND ' +
               'E_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '"';
      end;

      lQuery := OpenSql(lSt , True);

      FTotDebPointeRef := lQuery.FindField('TOTDEB').AsFloat;
      FTotCrePointeRef := lQuery.FindField('TOTCRE').AsFloat;
      FTotNbPointeRef  := lQuery.FindField('NBDEJAPOINTE').AsInteger;

    except
      on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, ' Total Pointé Réference');
      end;
    end;
  finally
    Ferme( lQuery );
  end;

  // Récupération des champs de la LISTE PARAMETRABLE
  lStListeChamps := CSqlTextFromList( FStListeChamps );

  // Condition where de la Requête SQL
  lStSqlWhereTobFListe := RecupWhereCritere(PageControl) + ' ' + RecupAutreWhere;

  // Requête SQL complète à éxécuter
  AStSqlTobFListe := 'SELECT ' + lStListeChamps + ',' + cAutreSelect + ' FROM ECRITURE ' + lStSqlWhereTobFListe +
                     ' ORDER BY E_GENERAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE';

  try
    // Calcul des Totaux des Mouvements Affichés de la Requête Principale
    lQuery := OpenSql('SELECT COUNT(*) LENOMBRE' +
                      ', SUM(' + FStLeDebit + ') LEDEBIT' +
                      ', SUM(' + FStLeCredit + ') LECREDIT FROM ECRITURE ' +
                      lStSqlWhereTobFListe + ' GROUP BY ' +
                      IIF( VH^.PointageJal, 'E_JOURNAL', 'E_GENERAL'), True);

    if not lQuery.Eof then
    begin
      // Remplissage des Zones des Mouvements Pointables
      SetControlText('TDMVTAFFICHE', StrfMontant(lQuery.FindField('LEDEBIT').AsFloat,  15, V_PGI.OkDecV, '', True));
      SetControlText('TCMVTAFFICHE', StrfMontant(lQuery.FindField('LECREDIT').AsFloat, 15, V_PGI.OkDecV, '', True));
      SetControlText('TMVTAFFICHE',  TraduireMemoire('Mouvements pointables') + ' : ' + IntToStr(lQuery.FindField('LENOMBRE').AsInteger));
    end
    else
    begin
      SetControlText('TDMVTAFFICHE', StrfMontant(0,  15, V_PGI.OkDecV, '', True));
      SetControlText('TCMVTAFFICHE', StrfMontant(0, 15, V_PGI.OkDecV, '', True));
      SetControlText('TMVTAFFICHE',  TraduireMemoire('Mouvements pointables') + ' : ' + IntToStr(0));
    end;

    // Calcul du Solde Comptable à la date de la Référence de Pointage
    lSoldeComptable := 0;
    try
      if not VH^.PointageJal then
      begin
        // Pointage Sur Compte
        lSoldeComptable := IIF(FBoCptEnDevise, FTotDebPTD - FTotCrePTD, FTotDebPTP - FTotCrePTP);
        lSoldeComptable := lSoldeComptable - (FTotDebDejaPointe - FTotCreDejaPointe);
        lSoldeComptable := lSoldeComptable + (lQuery.FindField('LEDEBIT').AsFloat - lQuery.FindField('LECREDIT').AsFloat);
      end
      else
      begin
        // Pointage Sur Journal
        lSoldeComptable := IIF(FBoCptEnDevise, FTotCrePTD - FTotDebPTD, FTotCrePTP - FTotDebPTP);
        lSoldeComptable := lSoldeComptable - (FTotCreDejaPointe - FTotDebDejaPointe);
        lSoldeComptable := lSoldeComptable + (lQuery.FindField('LECREDIT').AsFloat - lQuery.FindField('LEDEBIT').AsFloat);
      end;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Calcul Solde Comptable');
    end;
  finally
    Ferme( lQuery );
  end;

  if lSoldeComptable >= 0 then
  begin
    SetControlText('TDSOLDECOMPTABLE', StrfMontant(lSoldeComptable, 15, V_PGI.OkDecV, '', True));
    SetControlText('TCSOLDECOMPTABLE', StrfMontant(0, 15, V_PGI.OkDecV, '', True));
  end
  else
  begin
    SetControlText('TDSOLDECOMPTABLE', StrfMontant(0, 15, V_PGI.OkDecV, '', True));
    SetControlText('TCSOLDECOMPTABLE', StrfMontant(Abs(lSoldeComptable), 15, V_PGI.OkDecV, '', True));
  end;

  FBoFaireRequete := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2003
Modifié le ... : 25/01/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.RefreshFListe( vBoFetch : Boolean );
var
  i: integer;
  lNbAffiche: integer;
  lSoldeProgressif: Double;
  lInDepBoucle : integer;
  lStRefPointage : string;
begin
  inherited;
  lSoldeProgressif := 0;

  // Base de Depart du Solde Progressif ( on prend l'antérieurité déjà Pointée )
  if not vBoFetch then
  begin
    if not VH^.PointageJal then
    begin
      if FBoCptEnDevise then
        lSoldeProgressif := FTobGene.GetValue('G_TOTDEBPTD') - FTobGene.GetValue('G_TOTCREPTD')
      else
        lSoldeProgressif := FTobGene.GetValue('G_TOTDEBPTP') - FTobGene.GetValue('G_TOTCREPTP');
      lSoldeProgressif := lSoldeProgressif - (FTotDebDejaPointe - FTotCreDejaPointe);
    end
    else
    begin
      // ATTENTION SUR JOURNAL C EST INVERSE ( CREDIT - DEBIT )
      if FBoCptEnDevise then
        lSoldeProgressif := FTobGene.GetValue('G_TOTCREPTD') - FTobGene.GetValue('G_TOTDEBPTD')
      else
        lSoldeProgressif := FTobGene.GetValue('G_TOTCREPTP') - FTobGene.GetValue('G_TOTDEBPTP');
      lSoldeProgressif := lSoldeProgressif - (FTotCreDejaPointe - FTotDebDejaPointe);
    end;
  end;

  HGBeginUpdate(FListe);
  try
    if ATobFListe.Detail.Count > 0 then
    begin
      lInDepBoucle := 0;
      lNbAffiche   := 0;

      {$IFDEF EAGLCLIENT}
        if vBoFetch then
          lInDepBoucle := ATobFListe.GetValue('DEPARTBOUCLE');
      {$ENDIF}

      FListe.RowCount := ATobFListe.Detail.Count + 1;
      RemplitListeChamps( ATobFListe.Detail[0] ) ;
      for i := lInDepBoucle to (ATobFListe.Detail.Count - 1) do
      begin
        // Sauvegarde du champ
        if not (ATobFListe.Detail[i].FieldExists('REFPOINTAGESAV')) then
        begin
          lStRefPointage := ATobFListe.Detail[i].GetString('E_REFPOINTAGE');
          ATobFListe.Detail[i].AddChampSupValeur('REFPOINTAGESAV', lStRefPointage, False);
          ATobFListe.Detail[i].AddChampSupValeur('DEJACOMPARE', '-', False);
        end;

        ATobFListe.Detail[i].AddChampSupValeur('OKAFFICHE', GereAffichageEcriture(ATobFListe.Detail[i]), False);

        if ATobFListe.Detail[i].GetValue('OKAFFICHE') = 'X' then
        begin
          if (vBoFetch) and ( i = lInDepBoucle) then
            lSoldeProgressif := ATobFListe.detail[i-1].GetValue('SOLDEE');

          if not VH^.PointageJal then
            lSoldeProgressif := lSoldeProgressif + ATobFListe.Detail[i].GetValue(FStLeDebit) - ATobFListe.Detail[i].GetValue(FStLeCredit)
          else
            lSoldeProgressif := lSoldeProgressif + ATobFListe.Detail[i].GetValue(FStLeCredit) - ATobFListe.Detail[i].GetValue(FStLeDebit);

          // indexs enreg dans Q
          ATobFListe.Detail[i].AddChampSupValeur('SOLDEE', lSoldeProgressif, False);

          // Affichage Grille
          Inc(lNbAffiche);
          FListe.Row := lInDepBoucle + lNbAffiche;
          TOBVersTHGrid(ATobFListe.Detail[i]);

        end; //
      end; // for i := lInDepBoucle to...

      // GCO - 16/09/2005 - FQ 16460
      if lNbAffiche = 0 then
        FListe.VidePile(False)
      else
        FListe.RowCount := lInDepBoucle + IIF(lNbAffiche = 0, 2, lNbAffiche + 1);
    end; // if ATobFListe.Detail.Count > 0 then

  finally
    HGEndUpdate(FListe);

    SetControlText('TDMVTPOINTE', StrfMontant(FTotDebPointeRef, 15, V_PGI.OkDecV, '', True));
    SetControlText('TCMVTPOINTE', StrfMontant(FTotCrePointeRef, 15, V_PGI.OkDecV, '', True));
    SetControlText('TMVTPOINTE', TraduireMemoire('Mouvements pointés') + ' : ' + IntToStr(FTotNbPointeRef));

    CalculAvancement;
    SetControlProperty('PCUMUL', 'VISIBLE', False);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/02/2003
Modifié le ... :   /  /
Description .. : Recherche la position de certains champs de la liste pour
Suite ........ : les modifications d'affichage
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.IndiceColFListe;
var
  lStChamp, lStListeChamps: string;
  lInIndex: integer;
begin
  FColLeDebit      := 0;
  FColLeCredit     := 0;
  FColRefPointage  := 0;
  FColSoldePro     := 0;
  FColDatePointage := 0;

  lStListeChamps := FStListeChamps;
  lInIndex := 1;
  while lStListeChamps <> '' do
  begin
    lStChamp := ReadTokenSt(lStListeChamps);
    if (lStChamp = 'E_DEBIT') then
    begin
      if lStChamp <> FStLeDebit then
      begin
        FListe.ColLengths[lInIndex] := -1;
        FListe.ColWidths[lInIndex] := -1;
      end
      else
        FColLeDebit := lInIndex;
    end
    else if (lStChamp = 'E_CREDIT') then
    begin
      if lStChamp <> FStLeCredit then
      begin
        FListe.ColLengths[lInIndex] := -1;
        FListe.ColWidths[lInIndex] := -1;
      end
      else
        FColLeCredit := linIndex;
    end
    else if (lStChamp = 'E_DEBITDEV') then
    begin
      if lStChamp <> FStLeDebit then
      begin
        FListe.ColLengths[lInIndex] := -1;
        FListe.ColWidths[lInIndex] := -1;
      end
      else
        FColLeDebit := lInIndex;
    end
    else if (lStChamp = 'E_CREDITDEV') then
    begin
      if lStChamp <> FStLeCredit then
      begin
        FListe.ColLengths[lInIndex] := -1;
        FListe.ColWidths[lInIndex] := -1;
      end
      else
        FColLeCredit := linIndex;
    end
    else if (lStChamp = 'E_REFPOINTAGE') then
      FColRefPointage := lInIndex
    else if (lStChamp = 'E_DATEPOINTAGE') then
      FColDatePointage := lInIndex
    else if (lStChamp = 'SOLDEE') then
      FColSoldePro := lInIndex;

    Inc(lInIndex);
  end; // while
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;

  case Key of

    // Shift + Inser --> Toutes les écritures deviennent visibles
    45: if (FListe.Focused) and (Shift = [ssShift]) then
          AfficheEcriturePointee;

    // Shift + Suppr --> Rend non visible les ecritures en cours de pointage ou pointées
    46: if (FListe.Focused) and (Shift = [ssShift]) then
          CacheEcriturePointee;

    // Déselectionne toutes les écritures CTRL + D
    68: if (FListe.Focused) and (Shift = [ssCtrl]) then
        begin
          if FBoConsultation then Exit;
          DepointeTout;
          FBoAllSelected := False;
          BCherche.Click;
        end;

    // Ctrl + J
    74: if (FListe.Focused) and (Shift = [ssCtrl]) then
      begin
        if FBoConsultation then Exit;
        SelectionJour;
      end;

    // Ctrl + M
    77: if (FListe.Focused) and (Shift = [ssCtrl]) then
      begin
        if FBoConsultation then Exit;
        SelectionMois;
      end;

    // Ctrl + N
    78: if (FListe.Focused) and (Shift = [ssCtrl]) then
      begin
        if FBoConsultation then Exit;
        SelectionRapide;
      end;

    // Ctrl + Z
    90: if (FListe.Focused) and (Shift = [ssCtrl]) then
      begin
        if FBoConsultation then Exit;
        AnnuleDerniereOperation;
      end;

  else
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2003
Modifié le ... : 26/12/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;

  case key of

    VK_F10 : begin // Attention F10 fait aussi par l'ancetre
               if (Shift = [ssAlt]) then BValideEtRappro.Click;
             end;

    // ESPACE
    32:
      begin
        if FBoConsultation then  Exit;

        if FListe.Cells[FColRefPointage, FListe.Row] = FStEE_RefPointage then
          PointeEcriture(FListe.Row, False)
        else
          PointeEcriture(FListe.Row, True);
        if FListe.AllSelected then
          Exit;
        if FListe.Row < FListe.RowCount - 1 then
          FListe.Row := FListe.Row + 1;
      end;

    // C --> Première ligne suivante sélectionée en pointage
    67: PasseSelectionSuivante;

    // J --> Premier mouvement du jour suivant
    74: PasseJourSuivant;

    // L --> Mouvement suivant ayant le même libellé
    76: PasseLibelleSuivant('');

    // M --> Premier mouvement du mois suivant
    77: PasseMoisSuivant;

    // P --> Mouvement non pointé précédent de la liste
    80: PassePrecedent;

    // S --> Mouvement non pointé suivant de la liste
    83: PasseSuivant;
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/05/2003
Modifié le ... :   /  /    
Description .. : Pointe l'écriture avec la combinaison Ctrl + Click
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnMouseDownFListe(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (ATobFliste.Detail.Count = 0) or (FBoConsultation) then Exit;

  if (GetKeyState( VK_CONTROL ) < 0) and (Button = mbLeft) then
  begin
    if FListe.Cells[FColRefPointage, FListe.Row] = FStEE_RefPointage then
      PointeEcriture(FListe.Row, False)
    else
      PointeEcriture(FListe.Row, True);
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : Double Click sur la grille
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.OnDblClickFListe(Sender: TObject);
var
  AA: TActionFiche;
  M: RMVT;
  lTob: Tob;
  lQuery: TQuery;
begin
  if ATobFliste.Detail.Count = 0 then
    Exit;

  // Sortie si pas d'écriture dans la grille
  lTob := Tob(FListe.Objects[0, FListe.Row]);
  if lTob = nil then
    Exit;

  SauvePointageEnCours;
  // Ajout GCO - 14/02/2006
  PgiInfo('L''insertion ou la suppression de mouvements et la modification des montants du folio/pièce' + #13+#10 +
          'entraînent la suppression du pointage en cours des mouvements présents dans ce folio/pièce.', Ecran.Caption);
  // FIN GCO

  lQuery := nil;
  try
    lQuery := OpenSql('SELECT * FROM ECRITURE WHERE ' +
      '(E_Journal = "' + lTob.GetValue('CLE_JOURNAL') + '") and ' +
      '(E_Exercice = "' + lTob.GetValue('CLE_EXERCICE') + '") and ' +
      '(E_DateComptable = "' + USDateTime(lTob.GetValue('CLE_DATECOMPTABLE')) + '") and ' +
      '(E_NumeroPiece = ' + IntToStr(lTob.GetValue('CLE_NUMEROPIECE')) + ') and ' +
      '(E_NumLigne = ' + IntToStr(lTob.GetValue('CLE_NUMLIGNE')) + ') and ' +
      '(E_NumEche = ' + IntToStr(lTob.GetValue('CLE_NUMECHE')) + ') and ' +
      '(E_QualifPiece = "' + lTob.GetValue('CLE_QUALIFPIECE') + '")', True);

    AA := taModif;
    if (not FBoDroitEcritures) or (FBoConsultation) then
      AA := taConsult;

    if (lQuery.FindField('E_MODESAISIE').AsString <> '-') and
      (lQuery.FindField('E_MODESAISIE').AsString <> '') then // Bordereau
      LanceSaisieFolio(lQuery, AA)
    else
    begin
      // GC-26/03/2003
      if TrouveSaisie(lQuery, M, lQuery.FindField('E_QUALIFPIECE').Asstring) then
      begin
        M.NumLigVisu := lQuery.FindField('E_NUMLIGNE').AsInteger;
        LanceSaisie(lQuery, AA, M);
      end;
    end;

  finally
    Ferme(lQuery);
  end;

  FBoFaireRequete := True;
  BCherche.Click;
  ChargePointageEnCours;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
  lOldBrush: TBrush;
  lOldPen: TPen;
begin
  inherited;

  if ARow = 0 then Exit;

  lOldBrush := TBrush.Create;
  lOldPen := TPen.Create;
  // sauvegarde des valeurs courantes
  lOldBrush.assign(FListe.Canvas.Brush);
  lOldPen.assign(FListe.Canvas.Pen);

  try
    if (ACol = FColSoldePro) and (FListe.Row <> ARow) then
    begin
      if FTobGene.GetValue('G_SENS') = 'D' then
        Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed,
          ClGreen)
      else if FTobGene.GetValue('G_SENS') = 'C' then
        Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0,
          ClGreen, ClRed)
    end;

  finally
    // réaffectation des valeurs du canevas
    FListe.Canvas.Brush.Assign(lOldBrush);
    FListe.Canvas.Pen.Assign(lOldPen);
    if assigned(lOldBrush) then lOldBrush.Free;
    if assigned(lOldPen) then lOldPen.Free;
  end; // try

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2003
Modifié le ... :   /  /
Description .. : Pointe ou Depointe l'écriture dans la TOB et la GRILLE
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.PointeEcriture(vRow: integer; vOkPointe: Boolean);
var lTobEcr: Tob;
begin
  lTobEcr := GetO(FListe, vRow);
  if lTobEcr = nil then Exit;

  if vOkPointe then
  begin
    if Trim(FListe.Cells[FColRefPointage, vRow]) = '' then
    begin
      Inc(FTotNbPointeRef);
      //Total FStLeDebit, FStLeCredit des mvts pointés
      FTotDebPointeRef := FTotDebPointeRef + lTobEcr.GetValue(FStLeDebit);
      FTotCrePointeRef := FTotCrePointeRef + lTobEcr.GetValue(FStLeCredit);

      FListe.Cells[FColRefPointage, vRow] := FStEE_RefPointage;
      lTobEcr.PutValue('E_REFPOINTAGE', FStEE_RefPointage);
      if FColDatePointage > 0 then
      begin
        FListe.Cells[FColDatePointage, vRow] := DateToStr(FDtEE_DatePointage);
        lTobEcr.PutValue('E_DATEPOINTAGE', FDtEE_DatePointage);
      end;
    end;
  end
  else
  begin
    if FListe.Cells[FColRefPointage, vRow] = FStEE_RefPointage then
    begin
      Dec(FTotNbPointeRef);
      //Total FStLeDebit FStLeCredit des mvts pointés
      FTotDebPointeRef := FTotDebPointeRef - lTobEcr.GetValue(FStLeDebit);
      FTotCrePointeRef := FTotCrePointeRef - lTobEcr.GetValue(FStLeCredit);

      FListe.Cells[FColRefPointage, vRow] := '';
      lTobEcr.PutValue('E_REFPOINTAGE', '');
      if FColDatePointage > 0 then
      begin
        FListe.Cells[FColDatePointage, vRow] := DateToStr(iDate1900);
        lTobEcr.PutValue('E_DATEPOINTAGE', DateToStr(iDate1900));
      end;
    end;
  end;

  // Affichage des mouvements pointés
  SetControlText('TDMVTPOINTE', StrfMontant(FTotDebPointeRef, 15, V_PGI.OkDecV, '', True));
  SetControlText('TCMVTPOINTE', StrfMontant(FTotCrePointeRef, 15, V_PGI.OkDecV, '', True));
  SetControlText('TMVTPOINTE', TraduireMemoire('Mouvements pointés') + ' : ' + IntToStr(FTotNbPointeRef));

  CalculAvancement;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. : Pointage de tous les éléments de la grille
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.OnClickBSelectAll(Sender: TObject);
var i: integer;
begin
  inherited;
  if FBoConsultation then Exit;

  if FBoAllSelected then
    for i := 1 to FListe.RowCount - 1 do
      PointeEcriture(i, False)
  else
    for i := 1 to FListe.RowCount - 1 do
      PointeEcriture(i, True);

  FBoAllSelected := not FBoAllSelected;
  //BCherche.Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.AnnuleDerniereOperation;
begin
  if (FListe.Row <> 1) then
  begin
    if FListe.Cells[FColRefPointage, FListe.Row - 1] = FStEE_REFPointage
      then
      FListe.Row := FListe.Row - 1;
  end;
  if FListe.Cells[FColRefPointage, FListe.Row] = FStEE_REFPointage then
    PointeEcriture(FListe.Row, False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/05/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.DepointeTout;
var i : integer;
begin
  for i := 1 to FListe.RowCount - 1 do
    PointeEcriture(i, False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.SelectionRapide;
var i : integer;
begin
  for i := (FListe.Row) to FListe.RowCount - 1 do
    PointeEcriture(i, True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.SelectionJour;
var
  lTob: Tob;
  i: integer;
  lDtComptable: TDateTime;
begin
  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then Exit;

  lDtComptable := lTob.GetValue('E_DATECOMPTABLE');
  for i := 1 to FListe.RowCount - 1 do
  begin
    lTob := GetO(FListe, i);
    if lTob <> nil then
    begin
      if lTob.GetValue('E_DATECOMPTABLE') = lDtComptable then
        PointeEcriture(i, True);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.SelectionMois;
var
  lTob: Tob;
  i: integer;
  lJour, lMois, lAnnee: Word;
  lJour_, lMois_, lAnnee_: Word;
begin
  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then Exit;

  DecodeDate(lTob.GetValue('E_DATECOMPTABLE'), lAnnee, lMois, lJour);
  for i := 1 to FListe.RowCount - 1 do
  begin
    lTob := GetO(FListe, i);
    if lTob <> nil then
    begin
      DecodeDate(lTob.GetValue('E_DATECOMPTABLE'), lAnnee_, lMois_, lJour_);
      if lMois = lMois_ then
        PointeEcriture(i, True);
    end;    
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/05/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.CacheEcriturepointee;
begin
  SetControlText('TRECAPITULATIF', 'Pour la référence ' + FStEE_RefPointage
          + ' au ' + DateToStr(FDtEE_DatePointage) +
          ' avec les écritures pointées non visibles');
  FBoAffichePointe := False;
  RefReshFListe( False );
  //BCherche.Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.AfficheEcriturepointee;
begin
  SetControlText('TRECAPITULATIF', 'Pour la référence ' + FStEE_RefPointage
          + ' au ' + DateToStr(FDtEE_DatePointage));
  FBoAffichePointe := True;
  RefReshFListe( False );
  //BCherche.Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/03/2003
Modifié le ... : 25/03/2003
Description .. : Le curseur se deplace se le prochain libelle identique.
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.PasseLibelleSuivant(vStLibelle: string = '');
var
  Crit: string;
  lTob: TOB;

  //----------------------------------------------------------------------------
  procedure _RechercheLib(vNbBoucle: integer; vStLibelle: string = '');
  var
    i: integer;
  begin
    if vNbBoucle > 1 then
      exit; // on a deja fait une fois le tour des enregistrements de la grille on sort

    for i := (FListe.Row + 1) to (FListe.RowCount - 1) do
    begin
      lTob := GetO(FListe, i);
      if lTob = nil then
        Continue;
      if lTob.GetValue('E_LIBELLE') = Crit then
        break;
    end; // for

    if (i <= (FListe.RowCount - 1)) and (i <> 0) then
      FListe.Row := i
    else
    begin
      FListe.Row := 1;
      Inc(vNbBoucle);
      _RechercheLib(vNbBoucle, Crit);
    end;
  end;
  //----------------------------------------------------------------------------

begin
  if (FListe.Row <= 0) then
    Exit;

  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then
    Exit;

  FListe.SynEnabled := False;
  FListe.BeginUpdate;

  if vStLibelle = '' then
    Crit := lTob.GetValue('E_LIBELLE')
  else
    Crit := vStLibelle;

  if Crit <> '' then
    _RechercheLib(0, Crit);

  FListe.SynEnabled := True;
  FListe.EndUpdate;
  FListe.Refresh;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEECR.PasseSelectionSuivante;
var
  lTob: TOB;

  //----------------------------------------------------------------------------
  procedure _RechercheSelectionSuivante(vNbBoucle: integer; vStRefPointage:
    string);
  var
    i: integer;
  begin
    if vNbBoucle > 1 then
      exit; // on a deja fait une fois le tour des enregistrements de la grille on sort

    for i := (FListe.Row + 1) to (FListe.RowCount - 1) do
    begin
      lTob := GetO(FListe, i);
      if lTob = nil then
        Continue;
      if lTob.GetValue('E_REFPOINTAGE') = vStRefPointage then
        break;
    end; // for

    if (i <= (FListe.RowCount - 1)) and (i <> 0) then
      FListe.Row := i
    else
    begin
      FListe.Row := 1;
      Inc(vNbBoucle);
      _RechercheSelectionSuivante(vNbBoucle, vStRefPointage);
    end;
  end;
  //----------------------------------------------------------------------------

begin
  if (FListe.Row <= 0) then
    Exit;

  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then
    Exit;

  FListe.SynEnabled := False;
  FListe.BeginUpdate;

  if FStEE_Refpointage <> '' then
    _RechercheSelectionSuivante(0, FStEE_RefPointage);

  FListe.SynEnabled := True;
  FListe.EndUpdate;
  FListe.Refresh;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.PasseJourSuivant;
var
  lTob: TOB;
  lDtComptable: TDatetime;

  //----------------------------------------------------------------------------
  procedure _RecherchePasseJourSuivant;
  var
    i: integer;
  begin
    for i := (FListe.Row + 1) to (FListe.RowCount - 1) do
    begin
      lTob := GetO(FListe, i);
      if lTob = nil then
        Continue;
      if lTob.GetValue('E_DATECOMPTABLE') > lDtComptable then
        break;
    end; // for

    if (i <= (FListe.RowCount - 1)) and (i <> 0) then
      FListe.Row := i;
  end;
  //----------------------------------------------------------------------------

begin
  if (FListe.Row <= 0) or (FListe.Row = FListe.RowCount - 1) then
    Exit;

  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then
    Exit;

  lDtComptable := lTob.GetValue('E_DATECOMPTABLE');

  FListe.SynEnabled := False;
  FListe.BeginUpdate;

  _RecherchePasseJourSuivant;

  FListe.SynEnabled := True;
  FListe.EndUpdate;
  FListe.Refresh;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.PasseMoisSuivant;
var
  lTob: TOB;
  lJour, lMois, lAnnee: Word;

  //----------------------------------------------------------------------------
  procedure _RecherchePasseMoisSuivant;
  var
    i: integer;
    lAnnee_, lMois_, lJour_: Word;

  begin
    for i := (FListe.Row + 1) to (FListe.RowCount - 1) do
    begin
      lTob := GetO(FListe, i);
      if lTob = nil then
        Continue;

      DecodeDate(lTob.GetValue('E_DATECOMPTABLE'), lAnnee_, lMois_, lJour_);

      if (lMois_ > lMois) or ((lMois_ <= lMois) and (lAnnee_ > lAnnee)) then
        break;
    end; // for

    if (i <= (FListe.RowCount - 1)) and (i <> 0) then
      FListe.Row := i;
  end;
  //----------------------------------------------------------------------------

begin
  if (FListe.Row <= 0) or (FListe.Row = FListe.RowCount - 1) then
    Exit;

  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then
    Exit;

  DecodeDate(lTob.GetValue('E_DATECOMPTABLE'), lAnnee, lMois, lJour);

  FListe.SynEnabled := False;
  FListe.BeginUpdate;

  _RecherchePasseMoisSuivant;

  FListe.SynEnabled := True;
  FListe.EndUpdate;
  FListe.Refresh;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.PassePrecedent;
var
  lTob: TOB;

  //----------------------------------------------------------------------------
  procedure _RecherchePassePrecedent(vNbBoucle: integer);
  var
    i: integer;
  begin
    if vNbBoucle > 1 then
      exit; // on a deja fait une fois le tour des enregistrements de la grille on sort

    for i := (FListe.Row - 1) downto (1) do
    begin
      lTob := GetO(FListe, i);
      if lTob = nil then
        Continue;
      if lTob.GetValue('E_REFPOINTAGE') = '' then
        break;
    end; // for

    if (i <= (FListe.RowCount - 1)) and (i <> 0) then
      FListe.Row := i
    else
    begin
      FListe.Row := 1;
      Inc(vNbBoucle);
      _RecherchePassePrecedent(vNbBoucle);
    end;
  end;
  //----------------------------------------------------------------------------

begin
  if (FListe.Row <= 0) then
    Exit;

  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then
    Exit;

  FListe.SynEnabled := False;
  FListe.BeginUpdate;

  _RecherchePassePrecedent(0);

  FListe.SynEnabled := True;
  FListe.EndUpdate;
  FListe.Refresh;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.PasseSuivant;
var
  lTob: TOB;

  //----------------------------------------------------------------------------
  procedure _RecherchePasseSuivant;
  var
    i: integer;
  begin
    for i := (FListe.Row + 1) to (FListe.RowCount - 1) do
    begin
      lTob := GetO(FListe, i);
      if lTob = nil then
        Continue;
      if lTob.GetValue('E_REFPOINTAGE') = '' then
        break;
    end; // for

    if (i <= (FListe.RowCount - 1)) and (i <> 0) then
      FListe.Row := i;
  end;
  //----------------------------------------------------------------------------

begin
  if (FListe.Row <= 0) or (FListe.Row = FListe.RowCount - 1) then
    Exit;

  lTob := GetO(FListe, FListe.Row);
  if lTob = nil then
    Exit;

  FListe.SynEnabled := False;
  FListe.BeginUpdate;

  _RecherchePasseSuivant;

  FListe.SynEnabled := True;
  FListe.EndUpdate;
  FListe.Refresh;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/04/2003
Modifié le ... :   /  /
Description .. : Determine si le compte est tenue en monnaie de tenue du dossier
Suite ........ : ou en devise afin de savoir si on doit utiliser E_DEBIT, E_CREDIT
Suite ........ : ou E_DEBITDEV, E_CREDITDEV
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.TraiteDeviseDuCompte;
var
  lQuery: Tquery;
begin
  FBoCptEnDEvise := False;
  FStDevise := V_PGI.DevisePivot;

  lQuery := nil;
  try
    try
      lQuery := OpenSql('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL="' +
                FTobGene.GetValue('G_GENERAL')+'" AND BQ_NODOSSIER="'+V_PGI.NoDossier // 19/10/2006 YMO Multisociétés
                + '" ORDER BY BQ_GENERAL', True);

      if not lQuery.Eof then
      begin
        if lQuery.FindField('BQ_DEVISE').AsString <> V_PGI.DevisePivot then
        begin
          FBoCptEnDevise := True;
          FStDevise := lQuery.FindField('BQ_DEVISE').AsString; // Enregistre la devise du compte
        end;
      end;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : TraiteDeviseDuCompte');
    end;

  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Calcul de l' avancement du Pointage
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.CalculAvancement;
var
  lDebitAffiche: Double;
  lCreditAffiche: Double;
  lDebitPointe: Double;
  lCreditPointe: Double;
  lDebitSoldeComptable: Double;
  lCreditSoldeComptable: Double;
  lDebitSoldeBancaire: Double;
  lCreditSoldeBancaire: Double;
  lSoldeAvancement: Double;

begin
  lDebitAffiche  := Valeur(GetControlText('TDMVTAFFICHE'));
  lCreditAffiche := Valeur(GetControlText('TCMVTAFFICHE'));

  lDebitPointe  := Valeur(GetControlText('TDMVTPOINTE'));
  lCreditPointe := Valeur(GetControlText('TCMVTPOINTE'));

  lDebitSoldeComptable  := Valeur(GetControlText('TDSOLDECOMPTABLE'));
  lCreditSoldeComptable := Valeur(GetControlText('TCSOLDECOMPTABLE'));

  lDebitSoldeBancaire  := Valeur(GetControlText('TDSOLDEBANCAIRE'));
  lCreditSoldeBancaire := Valeur(GetControlText('TCSOLDEBANCAIRE'));

  if VH^.PointageJal then
  begin
    lSoldeAvancement := lCreditAffiche - lCreditPointe -
                        lDebitAffiche + lDebitPointe -
                        lDebitSoldeComptable + lCreditSoldeComptable +
                        lCreditSoldeBancaire - lDebitSoldeBancaire;
  end
  else
  begin
    lSoldeAvancement := lCreditSoldeBancaire - lDebitSoldeBancaire +
                        lDebitAffiche - lDebitPointe -
                        lCreditAffiche + lCreditPointe -
                        lDebitSoldeComptable + lCreditSoldeComptable;
  end;

  // Affichage de l'avancement du Pointage
  if (lSoldeAvancement) > 0 then
  begin
    SetControlText('TDAVANCEMENT', StrfMontant(lSoldeAvancement, 15, V_PGI.OkDecV, '', True));
    SetControlText('TCAVANCEMENT', StrfMontant(0, 15, V_PGI.OkDecV, '', True));
  end
  else
  begin
    SetControlText('TDAVANCEMENT', StrfMontant(0, 15, V_PGI.OkDecV, '', True));
    SetControlText('TCAVANCEMENT', StrfMontant(Abs(lSoldeAvancement), 15, V_PGI.OkDecV, '', True));
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Aligne LEFT et WIDTH sur la colonne d'une grille
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.AligneControleSurGrille(vGrille: THGrid; vCol: integer; vNomControl: string; vAvecWidth : boolean = False);
var lRect: TRect;
begin
  lRect := vGrille.CellRect(vCol, 0);
  SetControlProperty(vNomControl, 'LEFT', lRect.Left + 1);
  if vAvecWidth then
    SetControlProperty(vNomControl, 'WIDTH', (lRect.Right - lRect.Left)+1);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/12/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.ValidationPointage;
var
  lModalResult : TModalResult;
  i : integer;
  lStSql : string;
  lSt    : string;
  lStRefPointage    : string;
  lStRefPointageSav : string;
  lStlibelle        : string;

  function _Pointe: string;
  begin
    Result := 'UPDATE ECRITURE SET E_REFPOINTAGE= "' + FStEE_RefPointage + '",' +
      {$IFDEF TRSYNCHRO}
      {FQ Tréso 10019 : maj du champ E_TRESOSYNCHRO, Pour que lors de prochaine synchronisation
                        de la trésorerie le champ TE_DATERAPPRO puisse être mis à jour}
      'E_TRESOSYNCHRO = "MOD", ' +
      {$ENDIF}
      'E_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '",' +
      'E_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' +
      'E_JOURNAL = "' + ATobFListe.Detail[i].GetValue('CLE_JOURNAL') + '" AND ' +
      'E_EXERCICE = "' + ATobFListe.Detail[i].GetValue('CLE_EXERCICE') + '" AND ' +
      'E_DATECOMPTABLE = "' + UsDateTime(ATobFListe.Detail[i].GetValue('CLE_DATECOMPTABLE')) + '" AND ' +
      'E_NUMEROPIECE = ' + IntToStr(ATobFListe.Detail[i].GetValue('CLE_NUMEROPIECE')) + ' AND ' +
      'E_NUMLIGNE = ' + IntToStr(ATobFListe.Detail[i].GetValue('CLE_NUMLIGNE')) + ' AND ' +
      'E_NUMECHE = ' + IntToStr(ATobFListe.Detail[i].GetValue('CLE_NUMECHE')) + ' AND ' +
      'E_QUALIFPIECE = "' + ATobFListe.Detail[i].GetValue('CLE_QUALIFPIECE') + '" AND ' +
      'E_DATEMODIF = "' + UsTime(ATobFListe.Detail[i].GetValue('CLE_DATEMODIF')) + '"';
  end;

  function _Depointe: string;
  begin
    Result := 'UPDATE ECRITURE SET E_REFPOINTAGE= "",' +
      {$IFDEF TRSYNCHRO}
      {FQ Tréso 10019 : maj du champ E_TRESOSYNCHRO, Pour que lors de prochaine synchronisation
                        de la trésorerie le champ TE_DATERAPPRO puisse être mis à jour}
      'E_TRESOSYNCHRO = "MOD", ' +
      {$ENDIF}
      'E_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '",' +
      'E_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' +
      'E_JOURNAL = "' + ATobFListe.Detail[i].GetValue('CLE_JOURNAL') + '" AND ' +
      'E_EXERCICE = "' + ATobFListe.Detail[i].GetValue('CLE_EXERCICE') + '" AND ' +
      'E_DATECOMPTABLE = "' + UsDateTime(ATobFListe.Detail[i].GetValue('CLE_DATECOMPTABLE')) + '" AND ' +
      'E_NUMEROPIECE = ' + IntToStr(ATobFListe.Detail[i].GetValue('CLE_NUMEROPIECE')) + ' AND ' +
      'E_NUMLIGNE = ' + IntToStr(ATobFListe.Detail[i].GetValue('CLE_NUMLIGNE')) + ' AND ' +
      'E_NUMECHE = ' + IntToStr(ATobFListe.Detail[i].GetValue('CLE_NUMECHE')) + ' AND ' +
      'E_QUALIFPIECE = "' + ATobFListe.Detail[i].GetValue('CLE_QUALIFPIECE') + '" AND ' +
      'E_DATEMODIF = "' + UsTime(ATobFListe.Detail[i].GetValue('CLE_DATEMODIF')) + '"';
  end;

  function _SynchroS1 : string;
  begin
    Result := 'UPDATE ECRITURE SET E_IO = "X" WHERE E_IO <> "X" AND ' +
              'E_JOURNAL = "' + ATobFListe.Detail[i].GetValue('CLE_JOURNAL') + '" AND ' +
              'E_EXERCICE = "' + ATobFListe.Detail[i].GetValue('CLE_EXERCICE') + '" AND ' +
              'E_PERIODE = ' + IntToStr(GetPeriode(ATobFListe.Detail[i].GetValue('CLE_DATECOMPTABLE'))) + ' AND ' +
              'E_NUMEROPIECE = ' + IntToStr(ATobFListe.Detail[i].GetValue('CLE_NUMEROPIECE')) + ' AND ' +
              'E_QUALIFPIECE = "' + ATobFListe.Detail[i].GetValue('CLE_QUALIFPIECE') + '"';
  end;

begin
  LastError := 0;
  lModalResult := MrYes;

  if FBoSecurise then
  begin
    if Valeur(GetControlText('TDAVANCEMENT')) - Valeur(GetControlText('TCAVANCEMENT')) <> 0 then
      lModalResult := PgiAskCancel('Attention : l''avancement du pointage n''est pas équilibré.' + #13 + #10 +
        'Confirmez vous la validation du pointage?', Ecran.Caption);
  end
  else
  begin
    // On recherche si des références de pointage postérieures ont déjà été crées
    if ExisteSql('SELECT EE_DATEPOINTAGE FROM EEXBQ WHERE' +
                ' EE_GENERAL = "' + FStEE_General + '"' +
                ' AND EE_DATEPOINTAGE > "' + UsDateTime(FDtEE_DatePointage) + '"' +
                ' AND (EE_ORIGINERELEVE<>"INT" OR EE_ORIGINERELEVE IS NULL) ' +
                ' AND EE_NUMERO = 1') then
    begin
      if (Valeur(GetControlText('TDAVANCEMENT')) - Valeur(GetControlText('TCAVANCEMENT')) <> 0) then
      begin
        lModalResult := PgiAskCancel('Validation d''une référence de pointage non équilbrée.' + #13 + #10 +
                                     'Toutes les sessions de pointage suivantes, ne seront pas équilibrées.' + #13 + #10 +
                                     'Confirmez vous la validation du pointage?', Ecran.Caption);
      end
      else
      begin
        lModalResult := PgiAskCancel('Validation d''une référence de pointage équilbrée.' + #13 + #10 +
                                     'Toutes les sessions de pointage suivantes devront être revalidées.' + #13 + #10 +
                                     'Confirmez vous la validation du pointage?', Ecran.Caption);
      end;
    end;
  end;

  // On a demander l'annulation
  if lModalResult = MrCancel  then
      LastError := -1;

  if (lModalResult = MrYes) then
  begin
    try
      BeginTrans;
      for i := 0 to ATobFListe.Detail.Count - 1 do
      begin
        lStSql := '';

        lStRefpointage    := VarToStr(ATobFListe.Detail[i].GetValue('E_REFPOINTAGE'));
        lStRefPointageSav := varToStr(ATobFListe.Detail[i].GetValue('REFPOINTAGESAV'));
        lStLibelle        := VarToStr(ATobFListe.Detail[i].GetValue('E_LIBELLE'));

        if (lStRefPointage <> lStRefPointageSav) and (not VarIsNull(ATobFListe.Detail[i].GetValue('E_REFPOINTAGE'))) then
        begin
          if ATobFListe.Detail[i].GetValue('E_REFPOINTAGE') = FStEE_refPointage then
          begin
            lStSql     := _Pointe;
            FTotDebPTP := FTotDebPTP  + ATobFListe.Detail[i].GetValue('E_DEBIT');
            FTotCrePTP := FTotCrePTP  + ATobFListe.Detail[i].GetValue('E_CREDIT');
            FTotDebPTD := FTotDebPTD  + ATobFListe.Detail[i].GetValue('E_DEBITDEV');
            FTotCrePTD := FTotCrePTD  + ATobFListe.Detail[i].GetValue('E_CREDITDEV');
          end
          else
          begin
            lStSql     := _Depointe;
            FTotDebPTP := FTotDebPTP - ATobFListe.Detail[i].GetValue('E_DEBIT');
            FTotCrePTP := FTotCrePTP - ATobFListe.Detail[i].GetValue('E_CREDIT');
            FTotDebPTD := FTotDebPTD - ATobFListe.Detail[i].GetValue('E_DEBITDEV');
            FTotCrePTD := FTotCrePTD - ATobFListe.Detail[i].GetValue('E_CREDITDEV');
          end;
        end;

        if lStSql <> '' then
        begin
          if ExecuteSql(lStSql) <> 1 then
            raise Exception.Create('Une écriture a été modifiée par un autre utilisateur')
          else
          begin
            // Ecriture sauvegardée, on change E_IO de la pièce pour Synchro S1
            ExecuteSQL(_SynchroS1);
          end;

        end;
      end;

      // Update du solde du compte général
      if not EstTablePartagee( 'GENERAUX' ) then
        begin
        lStSql := 'UPDATE GENERAUX SET ' +
                  'G_TOTDEBPTP = ' + VariantToSql(FTotDebPTP) + ', ' +
                  'G_TOTCREPTP = ' + VariantToSql(FTotCrePTP) + ', ' +
                  'G_TOTDEBPTD = ' + VariantToSql(FTotDebPTD) + ', ' +
                  'G_TOTCREPTD = ' + VariantToSql(FTotCrePTD) + ' ' +
                  'WHERE G_GENERAL = "' + FTobGene.GetValue('G_GENERAL') + '"';

        if ExecuteSql(lStSql) <> 1 then
          raise Exception.Create('Erreur de mise à jour du compte général') ;
        end
      // MAJ table cumul en multi-soc
      else if not CUpdateCumulsPointeMS( FTobGene.getString('G_GENERAL'), FTotDebPTP, FTotCrePTP, FTotDebPTD, FTotCrePTD )
             then raise Exception.Create('Erreur de mise à jour du cumul multi-société (table CUMULS)') ;

      // Mise a jour du Statut de la référence de pointage
      if (Valeur(GetControlText('TDAVANCEMENT')) - Valeur(GetControlText('TCAVANCEMENT'))) <> 0 then
        lSt := '-'
      else
        lSt := 'X';

      lStSql := 'UPDATE EEXBQ SET EE_AVANCEMENT = "' + lSt + '" WHERE ' +
                'EE_GENERAL = "' + FStEE_General + '"';

      if (not FBoSecurise) and (Valeur(GetControlText('TDAVANCEMENT')) <> Valeur(GetControlText('TCAVANCEMENT'))) then
        lStSql := lStSql + ' AND EE_DATEPOINTAGE >="' + UsDateTime(FDtEE_DatePointage) + '"'
      else
        lStSql := lStSql + ' AND EE_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '"' +
                           ' AND EE_REFPOINTAGE = "' + FStEE_RefPointage + '"';

      //StSql := lStSql + ' AND (EE_ORIGINERELEVE<>"INT" OR EE_ORIGINERELEVE IS NULL)' +
      //                   ' AND EE_NUMERO=1';

      if ExecuteSql(lStSql) = 0 then
        raise Exception.Create('Erreur de mise à jour de la référence de pointage');

      CommitTrans;
    except
      on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Attention');
        RollBack;
      end;
    end;
  end; // if lModalResult = MrYes
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Valide le pointage des écritures
Suite ........ : Pas d' héritage de la fiche ancetre car dans celle ci on
Suite ........ : faisait le même traitement que le DoubleClique sur la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnClickBValider(Sender: TObject);
begin
  if FBoConsultation then Exit;
  ValidationPointage;
  if LastError = 0 then Ecran.Close;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/12/2005
Modifié le ... :   /  /
Description .. : FQ 16593
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnClickBValideEtRappro(Sender: TObject);
begin
  if FBoConsultation then Exit;
  ValidationPointage;
  if LastError = 0 then
  begin
    CC_LanceFicheEtatRapproDet(FStEE_General + ';' + DateToStr(FDtEE_DatePointage) + ';X');
    Ecran.Close;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/04/2003
Modifié le ... :   /  /
Description .. : Gere le Champ OKAFFICHE des écritures de TOBFListe afin
Suite ........ : de savoir si l'on doit les afficher en fonction des critères
Suite ........ : de l'utilisateur
Mots clefs ... :
*****************************************************************}

function TOF_POINTAGEECR.GereAffichageEcriture(vTobEcriture: Tob): string;
begin
  Result := '-';

  // GCO - 23/08/2005 - FQ 16460
  if (not FBoAffichePointe) and (vTobEcriture.GetString('E_REFPOINTAGE') <> '') then Exit;

  // E_DATECOMPTABLE
  if (GetControlText('E_DATECOMPTABLE') <> DateToStr(iDate1900)) and
    (vTobEcriture.GetValue('E_DATECOMPTABLE') <
    StrToDate(GetControlText('E_DATECOMPTABLE'))) then
    Exit;

  // E_DATECOMPTABLE_
  if (GetControlText('E_DATECOMPTABLE_') <> DateToStr(iDate2099)) and
    (vTobEcriture.GetValue('E_DATECOMPTABLE') >
    StrToDate(GetControlText('E_DATECOMPTABLE_'))) then
    Exit;

  // E_DATEECHEANCE
  if (GetControlText('E_DATEECHEANCE') <> DateToStr(iDate1900)) and
    (vTobEcriture.GetValue('CLE_DATEECHEANCE') >
    StrToDate(GetControlText('E_DATEECHEANCE'))) then
    Exit;

  // E_DATEECHEANCE_
  if (GetControlText('E_DATEECHEANCE_') <> DateToStr(iDate2099)) and
    (vTobEcriture.GetValue('CLE_DATEECHEANCE') >
    StrToDate(GetControlText('E_DATEECHEANCE_'))) then
    Exit;

  // E_NUMEROPIECE
  if (GetControlText('E_NUMEROPIECE') <> '') and
    (vTobEcriture.GetValue('CLE_NUMEROPIECE') <
    StrToInt(GetControlText('E_NUMEROPIECE'))) then
    Exit;

  // E_NUMEROPIECE_
  if (GetControlText('E_NUMEROPIECE_') <> '') and
    (vTobEcriture.GetValue('CLE_NUMEROPIECE') >
    StrToInt(GetControlText('E_NUMEROPIECE_'))) then
    Exit;

  // E_LIBELLE
  if (GetControlText('E_LIBELLE') <> '') and
    (Pos(GetControlText('E_LIBELLE'), vTobEcriture.GetValue('E_LIBELLE')) = 0)
      then
    Exit;

  // E_REFINTERNE
  if (GetControlText('E_REFINTERNE') <> '') and
    (Pos(GetControlText('E_REFINTERNE'), vTobEcriture.GetValue('E_REFINTERNE')) = 0) then
    Exit;

  Result := 'X';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. : Controle si la référence de pointage est la dernière crée
Mots clefs ... :
*****************************************************************}

function TOF_POINTAGEECR.ControleRefPointage: Boolean;
var
  lQuery: TQuery;
  lMaxDatePointage: TDateTime;
begin
  Result := True;
  lMaxdatePointage := iDate1900;

  try
    lQuery := nil;
    try
      lQuery := OpenSql('SELECT MAX(EE_DATEPOINTAGE) DATEPOINTAGE FROM EEXBQ WHERE ' +
                        'EE_GENERAL = "' + FStEE_General + '"', True);

      if IsValidDate(lQuery.Findfield('DATEPOINTAGE').AsString) then
        lMaxDatePointage := lQuery.Findfield('DATEPOINTAGE').AsDateTime;

    except
      on E: Exception do
        PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : ControleRefPointage');
    end;

  finally
    Ferme(lQuery);
  end;

  if lMaxDatePointage > FDtEE_DatePointage then
    Exit;

  // OK on travaille sur la dernière référence de pointage
  Result := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/05/2003
Modifié le ... :   /  /
Description .. : Vérification effectue uniquement en Pointage sur journal
Suite ........ :
Mots clefs ... :
*****************************************************************}

function TOF_POINTAGEECR.ControleContrepartie: Boolean;
begin
  Result := False;
  try
    Result := ExisteSql('SELECT E_GENERAL FROM ECRITURE WHERE ' +
      'E_GENERAL = "' + FTobGene.GetValue('G_GENERAL') + '" AND ' +
      'E_JOURNAL <> "' + FStEE_General + '" AND ' +
      'E_DATECOMPTABLE <="' + UsDateTime(FDtEE_DatePointage) + '"');
  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : ControleContrepartie');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/09/2003
Modifié le ... :   /  /
Description .. : Verifie la présence d'écritures dans une devise differente
Suite ........ : de celle du compte de BANQUE(RIB) ou de la Devise PIVOT
Suite ........ : uniquement dans le cas où la devise utilisée est différente
Suite ........ : de la monnaie de tenue du DOSSIER.
Mots clefs ... :
*****************************************************************}
function TOF_POINTAGEECR.ControleDevise : Boolean;
begin
  Result := False;
  try
    Result := ExisteSql('SELECT E_GENERAL FROM ECRITURE WHERE ' +
      'E_GENERAL = "' + FTobGene.GetValue('G_GENERAL') + '" AND ' +
      'E_DEVISE <> "' + FStDevise + '" AND ' +
      'E_DATECOMPTABLE <="' + UsDateTime(FDtEE_DatePointage) + '"');
  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : ControleDevise');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/05/2003
Modifié le ... : 13/02/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.OnClickBSaisie(Sender: TObject);
var lOldSC: Boolean;
begin
  lOldSC := VH^.BouclerSaisieCreat;
  SauvePointageEnCours;

  VH^.BouclerSaisieCreat := False;
  MultiCritereMvt(taCreat, 'N', False);
  VH^.BouclerSaisieCreat := lOldSC;

  FBoFaireRequete := True;
  BCherche.Click;
  ChargePointageEnCours;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/05/2003
Modifié le ... :   /  /
Description .. : On sauvegarde le contexte des écritures avant de partir en
Suite ........ : saisie bordereau ou pièce afin de ne pas perdre, le pointage
Suite ........ : que l'on a commencé.
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.SauvePointageEnCours;
var
  i: integer;
  lTobfille: Tob;
  lVariant : variant;
begin
  for i := 0 to ATobFListe.Detail.Count - 1 do
  begin
    ATobFliste.Detail[i].SetString('DEJACOMPARE', '-');

    if VarIsNull(ATobFListe.Detail[i].GetValue('E_REFPOINTAGE')) then
      lVariant := ''
    else
      lVariant := ATobFListe.Detail[i].GetValue('E_REFPOINTAGE');

    if lVariant <> ATobFListe.Detail[i].GetValue('REFPOINTAGESAV') then
    begin
      lTobFille := Tob.Create('ECRITURE', FTobSave, -1);
      lTobFille.Dupliquer(ATobFListe.Detail[i], False, True, False);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/05/2003
Modifié le ... : 19/02/2004
Description .. : On recharge le contexte des écritures avant d' être partie en
Suite ........ : saisie bordereau ou pièce afin de ne pas perdre, le pointage
Suite ........ : que l'on avait commencé.
Suite ........ : GCO - 19/02/2004
Suite ........ : -> Ajout du test VarIsNull pour le champ E_REFPOINTAGE suite
Suite ........ : au problème ORACLE ( Null <> "" ), qui fesait pointer toutes
Suite ........ : les écritures en retour de saisie d' une nouvelle pièce
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEECR.ChargePointageEnCours;
var i, j : integer;
begin
  for i := 0 to FTobSave.Detail.Count - 1 do
  begin
    for j := 0 to ATobFListe.Detail.Count - 1 do
    begin
      if (ATobFListe.Detail[j].GetValue('DEJACOMPARE') = '-') and
         (ATobFListe.Detail[j].GetValue('CLE_JOURNAL') = FTobSave.Detail[i].GetValue('CLE_JOURNAL')) and
         (ATobFListe.Detail[j].GetValue('CLE_EXERCICE') =  FTobSave.Detail[i].GetValue('CLE_EXERCICE')) and
         (ATobFListe.Detail[j].GetValue('CLE_DATECOMPTABLE') = FTobSave.Detail[i].GetValue('CLE_DATECOMPTABLE')) and
         (ATobFListe.Detail[j].GetValue('CLE_NUMEROPIECE') = FTobSave.Detail[i].GetValue('CLE_NUMEROPIECE')) and
         (ATobFListe.Detail[j].GetValue('CLE_NUMLIGNE') = FTobSave.Detail[i].GetValue('CLE_NUMLIGNE')) and
         (ATobFListe.Detail[j].GetValue('CLE_NUMECHE') = FTobSave.Detail[i].GetValue('CLE_NUMECHE')) and
         (ATobFListe.Detail[j].GetValue('CLE_QUALIFPIECE') = FTobSave.Detail[i].GetValue('CLE_QUALIFPIECE')) and
         (ATobFListe.Detail[j].GetValue('E_DEBIT') = FTobSave.Detail[i].GetValue('E_DEBIT')) and
         (ATobFListe.Detail[j].GetValue('E_CREDIT') = FTobSave.Detail[i].GetValue('E_CREDIT')) then
      begin
        ATobFListe.Detail[j].SetString('DEJACOMPARE', 'X');

        if (not VarIsNull(FTobSave.Detail[i].GetValue('E_REFPOINTAGE'))) and
           (FTobSave.Detail[i].GetValue('E_REFPOINTAGE') <> '') then
          PointeEcriture(j + 1, True)
        else
          PointeEcriture(j + 1, False);

        Break;
      end;
    end;
  end;
  // Supression des ecritures sauvegardées car elles ont été restaurées
  FTobSave.ClearDetail;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. : Controle afin d'empecher de saisir une date supérieure à la
Suite ........ : date de la référence de pointage
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEECR.OnExitE_DateComptable_(Sender: TObject);
begin
  if not IsValidDate(E_DateComptable_.Text) then
    Exit;

  if StrToDate(E_DateComptable_.Text) > FDtEE_DatePointage then
  begin
    PgiInfo('La date saisie doit être inférieure au ' + DateToStr(FDtEE_DatePointage + 1), 'Erreur de saisie');
    SetFocusControl('E_DATECOMPTABLE_');
    SetControlProperty('E_DATECOMPTABLE_', 'TEXT', FDtEE_DatePointage);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickPasseJoursuivant(Sender: TObject);
begin
  PasseJourSuivant;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickPasseLibelleSuivant(Sender: TObject);
begin
  PasseLibelleSuivant;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickPasseMoisSuivant(Sender: TObject);
begin
  PasseMoisSuivant;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickPasseSelectionSuivante(Sender: TObject);
begin
  PasseSelectionSuivante;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickPassePrecedent(Sender: TObject);
begin
  PassePrecedent;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickPasseSuivant(Sender: TObject);
begin
  PasseSuivant;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickAnnuleDerniereOperation;
begin
  if not POPF11.Items[3].Enabled then Exit;
  AnnuleDerniereOperation;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickDepointeTout(Sender: TObject);
begin
  if not POPF11.Items[4].Enabled then Exit;
  DepointeTout;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickSelectionRapide(Sender: TObject);
begin
  SelectionRapide;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickCacheEcriturepointee(Sender: TObject);
begin
  CacheEcriturepointee;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickAfficheEcriturepointee(Sender: TObject);
begin
  AfficheEcriturepointee;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickSelectionJour(Sender: TObject);
begin
  SelectionJour;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnClickSelectionMois(Sender: TObject);
begin
  SelectionMois;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_POINTAGEECR.OnResizeEcran(Sender: TObject);
begin
  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);

  // Alignement des Soldes avec les colonnes de la grille
  AligneControleSurGrille(FListe, FColLeDebit, 'TDMVTAFFICHE'    , True);
  AligneControleSurGrille(FListe, FColLeDebit, 'TDMVTPOINTE'     , True);
  AligneControleSurGrille(FListe, FColLeDebit, 'TDSOLDECOMPTABLE', True);
  AligneControleSurGrille(FListe, FColLeDebit, 'TDSOLDEBANCAIRE' , True);
  AligneControleSurGrille(FListe, FColLeDebit, 'TDAVANCEMENT'    , True);

  // Alignement des Soldes avec les colonnes de la grille
  AligneControleSurGrille(FListe, FColLeCredit, 'TCMVTAFFICHE'    , True);
  AligneControleSurGrille(FListe, FColLeCredit, 'TCMVTPOINTE'     , True);
  AligneControleSurGrille(FListe, FColLeCredit, 'TCSOLDECOMPTABLE', True);
  AligneControleSurGrille(FListe, FColLeCredit, 'TCSOLDEBANCAIRE' , True);
  AligneControleSurGrille(FListe, FColLeCredit, 'TCAVANCEMENT'    , True);

  AligneControleSurGrille(FListe, FColRefPointage, 'TMVTAFFICHE'    , False);
  AligneControleSurGrille(FListe, FColRefPointage, 'TMVTPOINTE'     , False);
  AligneControleSurGrille(FListe, FColRefPointage, 'TSOLDECOMPTABLE', False);
  AligneControleSurGrille(FListe, FColRefPointage, 'TSOLDEBANCAIRE' , False);
  AligneControleSurGrille(FListe, FColRefPointage, 'TAVANCEMENT'    , False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_POINTAGEECR.OnClickBImprimer;
begin
  // GCO - 15/10/2004 - FQ 14812
  SauvePointageEnCours;
  inherited;
  FBoFaireRequete := True;
  BCherche.Click;
  ChargePointageEnCours;
end;
////////////////////////////////////////////////////////////////////////////////

initialization
  registerclasses([TOF_POINTAGEECR]);
end.

