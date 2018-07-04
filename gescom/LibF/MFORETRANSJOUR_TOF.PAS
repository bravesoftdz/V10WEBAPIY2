{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/09/2002
Modifié le ... : 24/09/2002
Description .. : Source TOF de la FICHE : MFORETRANSJOUR ()
Mots clefs ... : TOF;MFORETRANSJOUR
*****************************************************************}
unit MFORETRANSJOUR_TOF;

interface

uses
  Classes, sysutils, Controls, forms,
  {$IFNDEF EAGLCLIENT}
  dbtables,
  {$ENDIF}
  HCtrls, HTB97, HEnt1, Vierge, UTOB, UTOF;

type
  TOF_MFORETRANSJOUR = class(TOF)
  private
    {$IFDEF TOXCLIENT}
    TOBEvent: TOB; // TOB des événements
    TOBParam: TOB; // TOB des paramètres par défaut
    GsReq: THGrid; // Grille de sélection des requêtes
    BAVANCE: TToolbarButton97; // Bouton configuration avancée
    MajEnCours: boolean; // Mise à jour en cours
    procedure BAVANCEClick(Sender: TObject);
    procedure BAFFICINFOClick(Sender: TObject);
    procedure ChargeControls;
    procedure ChargeSiteCentral;
    function DonneSiteCentral: string;
    procedure ChargeRequetes;
    function DonneRequetes(AvecTri: boolean): string;
    procedure ChargeParamDefaut;
    procedure ChargeEvenement;
    function DonneEvenement: string;
    procedure MaJEvenement;
    procedure MajDates;
    procedure DemarreEchange;
    procedure AttenteFinEchange;
    procedure ArretEchange;
    procedure LanceTransmission;
    procedure MessTraitement(Indice: integer; Fini: boolean);
    procedure SauveConfiguration;
    procedure LibereTOB;
    {$ENDIF}
  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  end;

implementation

uses
  {$IFDEF TOXCLIENT}
  uTox, uToxNet, uToxClasses, uToxConf, uToxConst, UtoxFiche, uToxDecProc, FOToxUtil,
  {$ENDIF}
  FOUtil, FODefi;

const
  CodeSiteCentral = 'SC'; // code par défaut du site central
  CodeEvenement = 'RETRANSJOUR'; // code par défaut de l'événement de retransmission
  ListeReq = 'CTRLCAISMT_DEM;JOURCAISSE_DEM;JOURSETAB_DEM;PIECES_ALF_DEM;' // valeur par défaut de la liste des requêtes
  + 'PIECES_BLF_DEM;PIECES_FCF_DEM;PIECES_FFO_DEM;PIECES_TEM_DEM;'
    + 'PIECES_TRE_DEM;PIECES_TRV_DEM;PIECES_INV_DEM;TIERS_CLI_DEM;';

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 19/12/2002
  Modifié le ... : 19/12/2002
  Description .. : OnNew
  Mots clefs ... :
  *****************************************************************}

procedure TOF_MFORETRANSJOUR.OnNew;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.OnDelete;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.OnUpdate;
begin
  inherited;
  {$IFDEF TOXCLIENT}
  if MajEncours then Exit;
  MajEncours := True;
  if (BAVANCE <> nil) and (BAVANCE.Down) then BAVANCE.Click;
  SetControlEnabled('BValider', False);
  SetControlVisible('ENCOURS', True);
  SourisSablier;
  EnableControl(Ecran, False);
  try
    // Enregistre la configuration dans la base de registres
    SauveConfiguration;
    // Modification de l'événement
    MaJEvenement;
    // Modification des heures de début et fin
    MajDates;
    if GetControlText('FE_EXECNOW') = 'X' then
    begin
      // Déclenchement de l'événement
      DemarreEchange;
      // Attend la fin de l'événement
      AttenteFinEchange;
      // Arrêt des échanges FO - BO
      ArretEchange;
      // Demande au ToxServeur de déclencher une communication
      LanceTransmission;
    end;
  finally
    SetControlVisible('ENCOURS', False);
    SourisNormale;
    EnableControl(Ecran, True);
    MajEncours := False;
  end;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.OnLoad;
begin
  inherited;
  {$IFDEF TOXCLIENT}
  MajEncours := False;
  // Chargement des contrôles de saisie
  ChargeControls;
  // Chargement de l'événement
  ChargeEvenement;
  // Chargement des paramètres par défaut
  ChargeParamDefaut;
  // Chargement des requêtes
  ChargeRequetes;
  // Chargement du site central
  ChargeSiteCentral;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.OnArgument(S: string);
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.OnClose;
begin
  inherited;
  {$IFDEF TOXCLIENT}
  if MajEnCours then
  begin
    LastError := 999; // pour ne pas fermer la fiche
    Exit;
  end;
  LibereTOB;
  {$ENDIF}
end;

{$IFDEF TOXCLIENT}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : bouton BAVANCE
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.BAVANCEClick(Sender: TObject);
begin
  if not (Sender is TToolbarButton97) then Exit;
  if GsReq = nil then Exit;
  SetControlVisible('PAVANCE', TToolbarButton97(Sender).Down);
  SetControlVisible('PMESSAGE', not TToolbarButton97(Sender).Down);
  if (TToolbarButton97(Sender).Down) and (Ecran <> nil) and (Ecran is TFVierge) then
    TFVierge(Ecran).HMTrad.ResizeGridColumns(GsReq);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : bouton BAFFICINFO
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.BAFFICINFOClick(Sender: TObject);
begin
  if AglToxConf(aceStatusFenetre) = 0 then
    AglToxConf(aceActiveFenetre)
  else
    AglToxConf(aceFermeFenetre);
  //AglToxConf( aceConfigure, '', Nil, Nil, Nil ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Chargement des contrôles de saisie
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.ChargeControls;
var Ctrl: TControl;
  Ok: boolean;
  Stg: string;
begin
  Ctrl := GetControl('GSREQ');
  if (Ctrl <> nil) and (Ctrl is THGrid) then GsReq := THGrid(Ctrl);
  Ctrl := GetControl('BAVANCE');
  if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
  begin
    TToolbarButton97(Ctrl).OnClick := BAVANCEClick;
    BAVANCE := TToolbarButton97(Ctrl);
  end;
  Ctrl := GetControl('BAFFICINFO');
  if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
  begin
    TToolbarButton97(Ctrl).OnClick := BAFFICINFOClick;
    TToolbarButton97(Ctrl).Visible := V_PGI.SAV;
  end;
  Ok := True;
  Ok := FOGetFromRegistry(REGTRANSJOUR, REGENVOIIMMEDIAT, Ok);
  if Ok then Stg := 'X' else Stg := '-';
  SetControlText('FE_EXECNOW', Stg);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Recherche du site central
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.ChargeSiteCentral;
var LeSite: TCollectionSite;
  sSite, sCode: string;
begin
  sSite := '';
  with TCollectionSites.Create(TCollectionSite, True) do
  begin
    // recherche si le site SC existe
    sCode := CodeSiteCentral;
    sCode := FOGetFromRegistry(REGTRANSJOUR, REGCODECENTRAL, sCode);
    LeSite := Find(sCode);
    if LeSite <> nil then
    begin
      if LeSite.SSI_SITEENABLED then
      begin
        sSite := LeSite.SSI_CODESITE;
        if (LeSiteLocal <> nil) and (sSite = LeSiteLocal.SSI_CODESITE) then
          sSite := '';
      end;
    end;
    // recherche du 1er site actif
    if sSite = '' then
    begin
      LeSite := First;
      while (sSite = '') and (LeSite <> nil) do
      begin
        if LeSite.SSI_SITEENABLED then
        begin
          sSite := LeSite.SSI_CODESITE;
          if (LeSiteLocal <> nil) and (sSite = LeSiteLocal.SSI_CODESITE) then
            sSite := '';
        end;
        LeSite := Next(LeSite);
      end;
    end;
    Free;
  end;
  SetControlText('FE_SITE', sSite);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Retourne le code du site central
Mots clefs ... :
*****************************************************************}

function TOF_MFORETRANSJOUR.DonneSiteCentral: string;
begin
  Result := GetControlText('FE_SITE');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Recherche des requêtes
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.ChargeRequetes;
var sSql, sListe: string;
  TOBReq: TOB;
  Ind: integer;
begin
  if GsReq = nil then Exit;
  TOBReq := TOB.Create('', nil, -1);
  sSql := 'SELECT SQE_CODEREQUETE,SQE_LIBELLE from STOXQUERYS'
    + ' where SQE_CODEREQUETE like "%DEM"';
  TOBReq.LoadDetailFromSQL(sSql);
  TOBReq.PutGridDetail(GsReq, False, False, 'SQE_CODEREQUETE;SQE_LIBELLE');
  TOBReq.Free;
  sListe := ListeReq;
  sListe := FOGetFromRegistry(REGTRANSJOUR, REGLISTEREQUETES, sListe);
  GsReq.ClearSelected;
  for Ind := GsReq.FixedRows to GsReq.RowCount do
    if FOStrCmp(GsReq.CellValues[GsReq.FixedCols, Ind], sListe) then
      GsReq.FlipSelection(Ind);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Retourne la liste des requêtes
Mots clefs ... :
*****************************************************************}

function TOF_MFORETRANSJOUR.DonneRequetes(AvecTri: boolean): string;
var Liste, Sql, Stg, NomTable, FirstR, LastR: string;
  Ind: integer;
  TOBQuery, TOBQ: TOB;
  QQ: TQuery;
  Lst: TStringList;
begin
  TOBQuery := nil;
  Result := '';
  Liste := '';
  if GsReq = nil then Exit;
  // Recherche des requêtes sélectionnées
  for Ind := GsReq.FixedRows to GsReq.RowCount do
    if GsReq.IsSelected(Ind) then
      Liste := Liste + GsReq.CellValues[GsReq.FixedCols, Ind] + ';';
  if (AvecTri) and (Liste <> '') then
  begin
    // Lecture des requêtes sélectionnées
    Sql := 'SELECT SQE_CODEREQUETE,SQE_INFOS FROM STOXQUERYS WHERE SQE_TYPETRT="001"'
      + FOFabriqueSQLIN(Liste, 'SQE_CODEREQUETE', True, False, False);
    QQ := OpenSQL(Sql, True);
    if not QQ.EOF then
    begin
      TOBQuery := TOB.Create('Les requêtes', nil, -1);
      TOBQuery.LoadDetailDB('STOXQUERYS', '', '', QQ, False);
    end;
    Ferme(QQ);
  end;
  if TOBQuery <> nil then
  begin
    // Tri des requêtes sélectionnées pour mettre en tête celles qui concerne les TIERS
    FirstR := '';
    LastR := '';
    while Liste <> '' do
    begin
      NomTable := '';
      Stg := ReadTokenSt(Liste);
      // recherche de la requête
      if Stg <> '' then
        TOBQ := TOBQuery.FindFirst(['SQE_CODEREQUETE'], [Stg], False)
      else
        TOBQ := nil;
      if TOBQ <> nil then
      begin
        // recherche de la table de la requête
        Lst := TStringList.Create;
        try
          Lst.Text := TOBQ.GetValue('SQE_INFOS');
          NomTable := Lst.Values['TABLE']
        finally
          Lst.Free;
        end;
      end;
      if UpperCase(NomTable) = 'TIERS' then
      begin
        if FirstR <> '' then FirstR := FirstR + ';';
        FirstR := FirstR + Stg;
      end else
      begin
        if LastR <> '' then LastR := LastR + ';';
        LastR := LastR + Stg;
      end;
    end;
    TOBQuery.Free;
    if FirstR <> '' then Result := FirstR;
    if LastR <> '' then
    begin
      if Result <> '' then Result := Result + ';';
      Result := Result + LastR;
    end;
  end else
  begin
    Result := Liste;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Chargement des paramètres par défaut
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.ChargeParamDefaut;
begin
  TOBParam := TOB.Create('STOXPARMS', nil, -1);
  TOBParam.InitValeurs;
  TOBParam.SelectDB('"000"', nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Chargement de l'événement
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.ChargeEvenement;
var sCode: string;
begin
  TOBEvent := TOB.Create('STOXEVENTS', nil, -1);
  TOBEvent.InitValeurs;
  sCode := CodeEvenement;
  sCode := FOGetFromRegistry(REGTRANSJOUR, REGCODEEVENEMENT, sCode);
  SetControlText('FE_EVENEMENT', sCode);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Retourne le code de l'événement
Mots clefs ... :
*****************************************************************}

function TOF_MFORETRANSJOUR.DonneEvenement: string;
begin
  Result := GetControlText('FE_EVENEMENT');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Modification de l'événement
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.MaJEvenement;
var StgLst: TStringList;
begin
  MessTraitement(1, False);
  StgLst := TStringList.Create;
  StgLst.Clear;
  if TOBEvent.SelectDB('"001";"' + DonneEvenement + '"', nil) then
  begin
    TOBEvent.PutValue('SEV_ACTIF', 'X');
    TOBEvent.PutValue('SEV_TYPECONDITION', '003'); // Déclenchement : Tout de suite
    StgLst.Text := TOBEvent.GetValue('SEV_INFOS');
    // modification des requêtes
    StgLst[0] := DonneRequetes(True);
    // modification du site central
    StgLst[1] := DonneSiteCentral;
    TOBEvent.PutValue('SEV_INFOS', StgLst.Text);
    TOBEvent.UpdateDateModif;
    TOBEvent.UpdateDB;
  end else
  begin
    TOBEvent.InitValeurs;
    TOBEvent.PutValue('SEV_TYPETRT', '001');
    TOBEvent.PutValue('SEV_CODEEVENT', DonneEvenement);
    TOBEvent.PutValue('SEV_LIBELLE', TraduireMemoire('Retransmission d''une journée'));
    TOBEvent.PutValue('SEV_TOXSPEC', '-');
    TOBEvent.PutValue('SEV_LOADMODE', '001');
    TOBEvent.PutValue('SEV_TYPE', 'EM');
    TOBEvent.PutValue('SEV_LESMOIS', TOBParam.GetValue('STP_LESMOIS'));
    TOBEvent.PutValue('SEV_LESJOURS', TOBParam.GetValue('STP_LESJOURS'));
    TOBEvent.PutValue('SEV_TELJOUR', TOBParam.GetValue('STP_TELJOUR'));
    TOBEvent.PutValue('SEV_HEUREDEBUT', TOBParam.GetValue('STP_HEUREDEBUT'));
    TOBEvent.PutValue('SEV_HEUREFIN', TOBParam.GetValue('STP_HEUREFIN'));
    TOBEvent.PutValue('SEV_NBESSAI', TOBParam.GetValue('STP_NBESSAI'));
    TOBEvent.PutValue('SEV_TIMER', TOBParam.GetValue('STP_TIMER'));
    TOBEvent.PutValue('SEV_ACTIF', 'X');
    TOBEvent.PutValue('SEV_TYPECONDITION', '003'); // Déclenchement : Tout de suite
    // ajout des requêtes
    StgLst.Add(DonneRequetes(True));
    // ajout du site central
    StgLst.Add(DonneSiteCentral);
    TOBEvent.PutValue('SEV_INFOS', StgLst.Text);
    TOBEvent.InsertDB(nil);
  end;
  StgLst.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Modification des dates de début et fin
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.MajDates;
var dDate: TDateTime;
  Stg: string;
begin
  Stg := GetControlText('FE_DATEDEB');
  dDate := StrToDate(Stg);
  TOBParam.PutValue('STP_DATEDEB', dDate);
  Stg := GetControlText('FE_DATEDEB_') + ' 23' + TimeSeparator + '59' + TimeSeparator + '59';
  dDate := StrToDateTime(Stg);
  TOBParam.PutValue('STP_DATEFIN', dDate);
  TOBParam.UpdateDB;
  MessTraitement(1, True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Fonction permettant de ne pas traiter un fichier TOX
Mots clefs ... :
*****************************************************************}

procedure FctToxEvent(P: TCollectionEnveloppe; var Atraiter: boolean);
begin
  Atraiter := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Fonction permettant de ne pas intégrer un fichier TOX
Mots clefs ... :
*****************************************************************}

procedure FctToxWork(P: TCollectionEnveloppe; Index: integer; TheTob: TOB; var ttResult: TToxTraitement; ToFeel: TToxFeelTob = nil);
begin
  ttResult := [ttRewrite, ttRetry];
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Démarrage des échanges FO - BO et déclenchement de l'événement
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.DemarreEchange;
var IdApplication: string;
  WhenStart: TDateTime;
begin
  if TX_LesEvenements <> nil then TX_LesEvenements.Reload;
  if not AglStatusTox(WhenStart) then
  begin
    // Recherche de l'identifiant de l'application du ToxServeur
    IdApplication := FOGetToxIdApp;
    // Démarrage des échanges FO - BO
    AglToxConf(aceStart, IdApplication, FctToxEvent, FctToxWork, AglToxFormError);
  end;
  MessTraitement(2, False);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Attente de la fin d'exécution de l'événement
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.AttenteFinEchange;
var Ok: boolean;
  iNb: integer;
  LeEvent: TCollectionEvent;
begin
  if TX_LesEvenements = nil then Exit;
  if V_PGI.SAV then iNb := 10000 else iNb := 60000; // attente d'une minute
  iNb := FOGetFromRegistry(REGTRANSJOUR, REGTMODEMARRAGE, iNb);
  Delay(iNb);
  Ok := False;
  iNb := 10000; // attente de 10 secondes
  iNb := FOGetFromRegistry(REGTRANSJOUR, REGTMOATTENTEFIN, iNb);
  while (not Ok) and (TX_LesEvenements.First <> nil) do
  begin
    SourisSablier;
    LeEvent := TX_LesEvenements.Find(DonneEvenement);
    if LeEvent = nil then Ok := True
    else Ok := (LeEvent.Fini) or (not LeEvent.SEV_ACTIF);
    if not Ok then Delay(iNb);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Arrêt des échanges FO - BO
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.ArretEchange;
var WhenStart: TDateTime;
  iNb: integer;
begin
  if AglStatusTox(WhenStart) then
  begin
    AglToxTerminate(aceStop);
    iNb := 10000; // attente de 10 secondes
    iNb := FOGetFromRegistry(REGTRANSJOUR, REGTMOATTENTEFIN, iNb);
    Delay(iNb);
    AglToxConf(aceStop, '', nil, nil, nil);
  end;
  MessTraitement(2, True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Demande au ToxServeur de déclencher une communication
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.LanceTransmission;
var iNb: integer;
begin
  SourisSablier;
  MessTraitement(3, False);
  if not AglToxCommuniCam(UST_RUNNING, nil) then
  begin
    AglToxCommuniCam(UST_START, nil);
    iNb := 10000; // attente de 10 secondes
    iNb := FOGetFromRegistry(REGTRANSJOUR, REGTMOATTENTEFIN, iNb);
    Delay(iNb);
  end;
  AglToxCommuniCam(UST_ForceExecute, nil);
  MessTraitement(3, True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Affichage d'un message d'état du traitement
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.MessTraitement(Indice: integer; Fini: boolean);
var Libelle: string;
begin
  SetControlVisible('PHASE' + IntToStr(Indice), True);
  if Fini then Libelle := 'Ok' else Libelle := 'En cours...';
  SetControlText('ETAT' + IntToStr(Indice), TraduireMemoire(Libelle));
  SetControlVisible('ETAT' + IntToStr(Indice), True);
  Application.ProcessMessages;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Enregistre la configuration dans la base de registres
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.SauveConfiguration;
var Ok: boolean;
  Stg: string;
begin
  FODeleteFromRegistry(REGTRANSJOUR, REGCODECENTRAL);
  Stg := DonneSiteCentral;
  if CodeSiteCentral <> Stg then
    FOSaveInRegistry(REGTRANSJOUR, REGCODECENTRAL, Stg);
  FODeleteFromRegistry(REGTRANSJOUR, REGLISTEREQUETES);
  Stg := DonneRequetes(False);
  if ListeReq <> Stg then
    FOSaveInRegistry(REGTRANSJOUR, REGLISTEREQUETES, Stg);
  FODeleteFromRegistry(REGTRANSJOUR, REGCODEEVENEMENT);
  Stg := DonneEvenement;
  if CodeEvenement <> Stg then
    FOSaveInRegistry(REGTRANSJOUR, REGCODEEVENEMENT, Stg);
  Ok := (GetControlText('FE_EXECNOW') = 'X');
  FODeleteFromRegistry(REGTRANSJOUR, REGENVOIIMMEDIAT);
  if not Ok then
    FOSaveInRegistry(REGTRANSJOUR, REGENVOIIMMEDIAT, Ok);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : libère les TOB créées
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORETRANSJOUR.LibereTOB;
begin
  if TOBEvent <> nil then
  begin
    TOBEvent.Free;
    TOBEvent := nil;
  end;
  if TOBParam <> nil then
  begin
    TOBParam.Free;
    TOBParam := nil;
  end;
end;
{$ENDIF}

initialization
  registerclasses([TOF_MFORETRANSJOUR]);
end.
