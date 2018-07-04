{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/05/2001
Modifié le ... : 30/10/2003
Description .. : Assistant de fermeture d'une journée de vente du Front
Suite ........ : Office
Mots clefs ... : FO
*****************************************************************}
unit FOFermeJour;

interface

uses
  Forms, assist, ImgList, Controls, StdCtrls, ComCtrls, HRichEdt, HRichOLE,
  CheckLst, Hctrls, Spin, Mask, HSysMenu, hmsgbox, ExtCtrls, HTB97, Classes,
  Windows, SysUtils, Graphics, Hent1, ParamSoc, EntGC, AGLInit,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  dbtables, FE_Main,
  {$ENDIF}
  UTob, HPanel, kb_Ecran;

function FOFermeJournee(PRien: THPanel): Boolean;

type
  TFFermeJour = class(TFAssist)
    PChoixCaisse: TTabSheet;
    GJC_CAISSE: THValComboBox;
    TGPK_ETABLISSEMENT: THLabel;
    LGPK_ETABLISSEMENT: TEdit;
    TGPK_DEPOT: THLabel;
    LGPK_DEPOT: TEdit;
    GPK_DEPOT: THValComboBox;
    GPK_ETABLISSEMENT: THValComboBox;
    TGJC_JOURNEE: THLabel;
    GJC_JOURNEE: THCritMaskEdit;
    PEvent: TTabSheet;
    TGJE_METEO: THLabel;
    GJE_METEO: THValComboBox;
    TLST_CODEEVT: THLabel;
    PBlocNote: TTabSheet;
    TitreChoix: THLabel;
    TitreEvent: THLabel;
    TGJE_BLOCNOTE: THLabel;
    TGJC_CAISSE: THLabel;
    PResume: TTabSheet;
    GJE_BLOCNOTE: THRichEditOLE;
    CompteRendu: TRichEdit;
    GET_CODEEVENT: THMultiValComboBox;
    ListeImages: TImageList;
    LST_CODEEVT: TCheckListBox;
    LGJC_CAISSE: TEdit;
    GJE_NBENTREES: TSpinEdit;
    TGJE_NBENTREES: THLabel;
    PAction: TTabSheet;
    cbTicketZ: TCheckBox;
    cbRecapVd: TCheckBox;
    cbFOBO: TCheckBox;
    TitreAction: THLabel;
    TitreAction2: THLabel;
    cbFdCaisse: TCheckBox;
    TGJC_VENDFERME: THLabel;
    GJC_VENDFERME: THCritMaskEdit;
    PPave: TPanel;
    PCtrlCais: TTabSheet;
    BLEFT: TToolbarButton97;
    BRIGHT: TToolbarButton97;
    BUP: TToolbarButton97;
    BDOWN: TToolbarButton97;
    PPieceAtt: TTabSheet;
    TitrePieceAtt: THLabel;
    HLabel2: THLabel;
    NbPieceAtt: THNumEdit;
    TNbPieceAtt: THLabel;
    BTicketAttente: TToolbarButton97;
    cbListeArt: TCheckBox;
    cbListeReg: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure PChange(Sender: TObject); override;
    procedure GJC_CAISSEChange(Sender: TObject);
    procedure GJC_JOURNEEChange(Sender: TObject);
    procedure GJC_VENDFERMEElipsisClick(Sender: TObject);
    procedure GJC_VENDFERMEDblClick(Sender: TObject);
    procedure GJE_NBENTREESChange(Sender: TObject);
    procedure GJC_VENDFERMEEnter(Sender: TObject);
    procedure GJE_METEOEnter(Sender: TObject);
    procedure LST_CODEEVTEnter(Sender: TObject);
    procedure BLEFTClick(Sender: TObject);
    procedure BRIGHTClick(Sender: TObject);
    procedure BUPClick(Sender: TObject);
    procedure BDOWNClick(Sender: TObject);
    procedure BTicketAttenteClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    TOBCaisse: TOB; // TOB de la caisse choisie
    TOBJCaisse: TOB; // TOB de la journée de la caisse
    TOBJouEtab: TOB; // TOB de la journée de l'établissement
    MajJouEtab: Boolean; // modification de la journée établissement
    TOBCtrlCais: TOB; // TOB du contrôle caisse
    CtrlCaisOk: Boolean; // Eléments du contrôle déjà chargés en TOB
    NbPage: Integer; // Nombre de pages actives
    NoPage: Integer; // Numéro de la page courante
    CodeEvent: string; // Liste des événements de la journée (chaîne tokenisée) avant modification
    NoEvent: Integer; // N° d'événements avant modification
    MiseAttente: Boolean; // Mise en attente de communication avec le central
    TicketGenere: Boolean; // Des tickets ont été générés pour le fond de caisse
    PnlBoutons: TClavierEcran; // pavé tactile
    BoutonEncours: Boolean; // traitement d'un bouton en cours
    LanceCpteurFait: boolean; // Lancement du traitement des compteurs d'entrées fait
    MajDone: boolean; // Mise à jour faite
    procedure ChargeCaisse;
    procedure ChargeJCaisse;
    procedure ChargeJouEtab;
    procedure CreeTicketFondCaisse;
    function SauveCaisse: boolean;
    procedure EnregistreSaisie;
    procedure ImprimeTicket;
    procedure ChargeCompteRendu;
    procedure ActivebFin;
    procedure ChargeListeEvtEtab;
    procedure SauveListeEvtEtab;
    procedure EvtEtabToToken;
    procedure TokenToEvtEtab(CodeEtab: string; idEvt: Integer);
    function VerifieNbEntrees: Boolean;
    function VerifieVendFerme: Boolean;
    function VerifiePageActive(TT: TTabSheet): Boolean;
    procedure CalculNbPieceAttente;
    procedure LanceCtrlCaisse;
    procedure LanceCompteurs;
    procedure SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure BoutonCalculetteClick(Val: string);
    procedure SaisieCalculette(Val: Double);
    procedure ChoixVendeurPave;
  public
    { Déclarations publiques }
    function PreviousPage: TTabSheet; override;
    function NextPage: TTabSheet; override;
  end;

implementation
uses
  {$IFDEF TOXCLIENT}
  FOToxUtil, uToxConst, uToxConf, uToxFiche, uToxNet, gcToxWork, GCToxCtrl, UTox,
  {$ENDIF}
  MFOCTRLECAISSE_TOF, FOAttenteComm, TickUtilFO, FOUtil, FODefi;

{$R *.DFM}

///////////////////////////////////////////////////////////////////////////////////////
//  FOFermeJournee : fermeture d'une journée
///////////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/05/2001
Modifié le ... : 23/07/2001
Description .. : Fermeture d'une journée de vente du Front Office
Mots clefs ... : FO
*****************************************************************}

function FOFermeJournee(PRien: THPanel): Boolean;
var X: TFFermeJour;
  Titre: string;
begin
  Result := True;
  FOAlerteAnnoncesNonValidees(FOGetParamCaisse('GPK_ETABLISSEMENT'), FOGetParamCaisse('GPK_DEPOT'));
  Titre := TraduireMemoire('Fermeture de la journée de vente');
  PRien.InsideTitle.Caption := Titre;
  X := TFFermeJour.Create(Application);
  try
    X.ShowModal;
  finally
    if X.MiseAttente then Result := False;
    X.Free;
  end;
  if not Result then FOMiseAttenteComm(Titre);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetNoEvent : Retourne le n° d'événement disponible
///////////////////////////////////////////////////////////////////////////////////////

function GetNoEvent(CodeEtab: string): Integer;
var QQ: TQuery;
  Stg: string;
begin
  Result := 0;
  // Recherche du dernier numéro attribué
  Stg := 'SELECT max(GET_NOEVENT) FROM JOURSETABEVT where GET_ETABLISSEMENT="' + CodeEtab + '"';
  QQ := OpenSQL(Stg, True);
  if not QQ.EOF then Result := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  // Incrémentation du numéro
  Inc(Result);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeCaisse : charge la TOB de la caisse choisie
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ChargeCaisse;
var Action: Boolean;
begin
  // Lecture du paramétrage de la caisse
  TOBCaisse.SelectDB('"' + GJC_CAISSE.Value + '"', nil);
  // Lecture des actions complémentaires à la fermeture de journée
  Action := True;
  cbTicketZ.Checked := FOGetFromRegistry(REGJOURNEE, REGFERMETICKETZ, Action);
  if not (FOExisteCodeEtat(efoTicketZ)) then
  begin
    cbTicketZ.Visible := False;
    cbTicketZ.Checked := False;
  end;
  Action := True;
  cbFdCaisse.Checked := FOGetFromRegistry(REGJOURNEE, REGFERMEFDCAISSE, Action);
  if not (FOEstVrai(TOBCaisse, 'GPK_GEREFONDCAISSE')) and not (FOExisteCodeEtat(efoFdCaisse)) then
  begin
    cbFdCaisse.Visible := False;
    cbFdCaisse.Checked := False;
  end;
  Action := True;
  cbRecapVd.Checked := FOGetFromRegistry(REGJOURNEE, REGFERMERECAPVD, Action);
  if not (FOEstVrai(TOBCaisse, 'GPK_VENDSAISIE')) and not (FOEstVrai(TOBCaisse, 'GPK_VENDSAISLIG')) and
    not (FOExisteCodeEtat(efoRecapVend)) then
  begin
    cbRecapVd.Visible := False;
    cbRecapVd.Checked := False;
  end;
  Action := False;
  cbListeArt.Checked := FOGetFromRegistry(REGJOURNEE, REGFERMELISTEART, Action);
  if not (FOExisteCodeEtat(efoListeArtVendu)) then
  begin
    cbListeArt.Visible := False;
    cbListeArt.Checked := False;
  end;
  Action := False;
  cbListeReg.Checked := FOGetFromRegistry(REGJOURNEE, REGFERMELISTEREG, Action);
  if not (FOExisteCodeEtat(efoListeRegle)) then
  begin
    cbListeReg.Visible := False;
    cbListeReg.Checked := False;
  end;
  {$IFDEF TOXCLIENT}
  Action := True;
  cbFOBO.Checked := FOGetFromRegistry(REGJOURNEE, REGFERMEFOBO, Action);
  if not V_PGI.FSAV then
  begin
    cbFOBO.Enabled := False;
    cbFOBO.Checked := True;
  end;
  {$ELSE}
  cbFOBO.Visible := False;
  cbFOBO.Enabled := False;
  cbFOBO.Checked := False;
  {$ENDIF}
  // Affichage des données complémentaires
  LGJC_CAISSE.Text := GJC_CAISSE.Text;
  GPK_ETABLISSEMENT.Value := TOBCaisse.GetValue('GPK_ETABLISSEMENT');
  LGPK_ETABLISSEMENT.Text := GPK_ETABLISSEMENT.Text;
  GPK_DEPOT.Value := TOBCaisse.GetValue('GPK_DEPOT');
  LGPK_DEPOT.Text := GPK_DEPOT.Text;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeJCaisse : charge la TOB de la journée de la caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ChargeJCaisse;
var Caisse, NumZ: string;
begin
  Caisse := GJC_CAISSE.Value;
  NumZ := IntToStr(FOGetNumZCaisse(Caisse, 'MAX'));
  TOBJCaisse.SelectDB('"' + Caisse + '";"' + NumZ + '"', nil);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeJouEtab : charge la TOB de la journée de l'établissement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ChargeJouEtab;
var sDate, Stg: string;
  QQ: TQuery;
begin
  // Lecture de la journée de l'établissement
  if (FOEstVrai(TOBCaisse, 'GPK_GEREEVENT')) then
  begin
    sDate := USDate(GJC_JOURNEE);
    TOBJouEtab.SelectDB('"' + GPK_ETABLISSEMENT.Value + '";"' + sDate + '"', nil);
    GJE_NBENTREES.Value := TOBJouEtab.GetValue('GJE_NBENTREES');
    GJE_METEO.Value := TOBJouEtab.GetValue('GJE_METEO');
    GJE_BLOCNOTE.Text := TOBJouEtab.GetValue('GJE_BLOCNOTE');
    // Chargement des événements de la journée
    NoEvent := TOBJouEtab.GetValue('GJE_NOEVENT');
    if (NoEvent <> 0) then
    begin
      Stg := 'SELECT * FROM JOURSETABEVT where GET_ETABLISSEMENT="' + GPK_ETABLISSEMENT.Value + '" '
        + 'and GET_NOEVENT=' + IntToStr(NoEvent);
      QQ := OpenSQL(Stg, True);
      if not QQ.EOF then TOBJouEtab.LoadDetailDB('JOURSETABEVT', '', '', QQ, False);
      Ferme(QQ);
      EvtEtabToToken;
    end;
    ChargeListeEvtEtab;
  end;
  MajJouEtab := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeListeEvtEtab : charge la liste des événements de la journée
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ChargeListeEvtEtab;
var Ind: Integer;
begin
  LST_CODEEVT.Items.BeginUpdate;
  LST_CODEEVT.Items.Clear;
  for Ind := 0 to GET_CODEEVENT.Items.count - 1 do
  begin
    LST_CODEEVT.Items.Add(GET_CODEEVENT.Items[Ind]);
    LST_CODEEVT.Checked[Ind] := False;
    if (pos(';' + GET_CODEEVENT.values[Ind] + ';', ';' + GET_CODEEVENT.Text + ';') > 0) then LST_CODEEVT.checked[Ind] := TRUE;
  end;
  LST_CODEEVT.Items.EndUpdate;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauveListeEvtEtab : enregistre la liste des événements de la journée
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.SauveListeEvtEtab;
var Nbl: Integer;
begin
  GET_CODEEVENT.Text := '';
  for Nbl := 0 to LST_CODEEVT.Items.Count - 1 do
    if LST_CODEEVT.Checked[Nbl] then
    begin
      if GET_CODEEVENT.Text <> '' then GET_CODEEVENT.Text := GET_CODEEVENT.Text + ';';
      GET_CODEEVENT.Text := GET_CODEEVENT.Text + GET_CODEEVENT.Values[Nbl];
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EvtEtabToToken : place les événements de la journée dans une chaîne tokenisée
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.EvtEtabToToken;
var TOBL: TOB;
  Nbl: Integer;
begin
  CodeEvent := '';
  for Nbl := 0 to TOBJouEtab.Detail.Count - 1 do
  begin
    TOBL := TOBJouEtab.Detail[Nbl];
    if (Nbl > 0) then CodeEvent := CodeEvent + ';';
    CodeEvent := CodeEvent + TOBL.GetValue('GET_CODEEVENT');
  end;
  GET_CODEEVENT.Text := CodeEvent;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TokenToEvtEtab : transforme une chaîne tokenisée en TOB des événements de la journée
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.TokenToEvtEtab(CodeEtab: string; idEvt: Integer);
var TOBL: TOB;
  Nbl: Integer;
  Stg: string;
begin
  // Suppression des événements saisis précédemment
  for Nbl := TOBJouEtab.Detail.Count - 1 downto 0 do
  begin
    TOBL := TOBJouEtab.Detail[Nbl];
    if (TOBL <> nil) then TOBL.free;
  end;
  // Ajout des événements choisis dans la TOB
  Stg := GET_CODEEVENT.Text;
  while Stg <> '' do
  begin
    TOBL := TOB.Create('JOURSETABEVT', TOBJouEtab, -1);
    TOBL.PutValue('GET_ETABLISSEMENT', CodeEtab);
    TOBL.PutValue('GET_NOEVENT', idEvt);
    TOBL.PutValue('GET_CODEEVENT', ReadTokenSt(Stg));
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieNbEntrees : vérifie la cohérence du champ NbEntrees
///////////////////////////////////////////////////////////////////////////////////////

function TFFermeJour.VerifieNbEntrees: Boolean;
var Stg: string;
  Taille, Ind: Integer;
begin
  Result := True;
  Stg := GJE_NBENTREES.Text;
  Taille := Length(Stg);
  for Ind := 1 to Taille do if not IsDelimiter('0123456789', Stg, Ind) then
    begin
      Msg.Execute(8, Caption, '');
      if GJE_NBENTREES.CanFocus then GJE_NBENTREES.SetFocus;
      Result := False;
      Break;
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieVendFerme : vérifie la cohérence du code vendeur
///////////////////////////////////////////////////////////////////////////////////////

function TFFermeJour.VerifieVendFerme: Boolean;
begin
  Result := True;
  if GJC_VENDFERME.Visible then
  begin
    if GJC_VENDFERME.Text = '' then
    begin
      if FOEstVrai(TOBCaisse, 'GPK_VENDOBLIG') then
      begin
        //Msg.Execute(6, Caption, '') ;
        Result := False
      end;
    end else
    begin
      if not FORepresentantExiste(GJC_VENDFERME.Text, 'VEN', nil) then
      begin
        Msg.Execute(7, Caption, '');
        Result := False;
      end;
    end;
  end;
  if not Result then
  begin
    if GJC_VENDFERME.CanFocus then GJC_VENDFERME.SetFocus;
    GJC_VENDFERMEElipsisClick(Self);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CreeTicketFondCaisse : créé les tickets d'annulation du fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.CreeTicketFondCaisse;
var
  sSql, sArt, sVen, sMdp, sDev, sTiers: string;
  TOBFdc, TOBL: TOB;
  Ind: Integer;
  dMtFdc: Double;
begin
  sSql := 'select GJM_FDCAISSEDEV,GJM_MODEPAIE,GJM_DEVISE from CTRLCAISMT where GJM_FDCAISOUV<>0'
    + ' and GJM_CAISSE="' + GJC_CAISSE.Value + '"'
    + ' and GJM_NUMZCAISSE="' + IntToStr(TOBJcaisse.GetValue('GJC_NUMZCAISSE')) + '" '
    + FOFabriqueListeMDP('GPK_MDPCTRLCAIS', 'GJM_MODEPAIE');
  TOBFdc := TOB.Create('Les fonds de caisse', nil, -1);
  TOBFdc.LoadDetailFromSQL(sSql);
  if (TOBFdc <> nil) and (TOBFdc.Detail.Count > 0) then
  begin
    sArt := FOGetParamCaisse('GPK_FDCAISSE');
    sVen := GJC_VENDFERME.Text;
    for Ind := 0 to TOBFdc.Detail.Count - 1 do
    begin
      TOBL := TOBFdc.Detail[Ind];
      dMtFdc := TOBL.GetValue('GJM_FDCAISSEDEV');
      if dMtFdc <> 0 then
      begin
        sMdp := TOBL.GetValue('GJM_MODEPAIE');
        sDev := TOBL.GetValue('GJM_DEVISE');
//        {$IFDEF GESCOM}
//        sTiers := '&#@';
//        {$ELSE}
//        sTiers := '';
//        {$ENDIF}
        FOGenereTicket(sArt, sVen, sMdp, sDev, sTiers, Caption, (dMtFdc * -1), True, True, True);
        TicketGenere := True;
      end;
    end;
  end;
  if TOBFdc <> nil then TOBFdc.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauveCaisse : enregistre la TOB de la caisse choisie
///////////////////////////////////////////////////////////////////////////////////////

function TFFermeJour.SauveCaisse: boolean;
var idEvt: Integer;
  io: TIOErr;
begin
  Result := False;
  idEvt := TOBJouEtab.GetValue('GJE_NOEVENT');
  // Met à jour la journée de la caisse à l'état "Cloturée définitivement"
  TOBJcaisse.PutValue('GJC_ETAT', ETATJOURCAISSE[4]);
  //TOBJcaisse.PutValue('GJC_DATEFERME', NowH) ;
  TOBJcaisse.PutValue('GJC_DATEFERME', Date);
  TOBJcaisse.PutValue('GJC_HEUREFERME', TimeToStr(Time));
  TOBJcaisse.PutValue('GJC_USERFERME', V_PGI.FUser);
  if TOBJcaisse.FieldExists('GJC_VENDFERME') then TOBJcaisse.PutValue('GJC_VENDFERME', GJC_VENDFERME.Text);
  // Met à jour les événements de la journée
  SauveListeEvtEtab;
  if (FOEstVrai(TOBCaisse, 'GPK_GEREEVENT')) and (GET_CODEEVENT.Text <> CodeEvent) then
  begin
    if (idEvt = 0) and (GET_CODEEVENT.Text <> '') then idEvt := GetNoEvent(GPK_ETABLISSEMENT.Value);
    TokenToEvtEtab(GPK_ETABLISSEMENT.Value, idEvt);
    MajJouEtab := True;
  end;
  // Met à jour la journée de l'établissement
  if (FOEstVrai(TOBCaisse, 'GPK_GEREEVENT')) and
    ((GJE_NBENTREES.Value <> TOBJouEtab.GetValue('GJE_NBENTREES')) or
    (GJE_METEO.Value <> TOBJouEtab.GetValue('GJE_METEO')) or
    (GJE_BLOCNOTE.Text <> TOBJouEtab.GetValue('GJE_BLOCNOTE'))) then
  begin
    TOBJouEtab.PutValue('GJE_ETABLISSEMENT', TOBCaisse.GetValue('GPK_ETABLISSEMENT'));
    TOBJouEtab.PutValue('GJE_JOURNEE', StrToDate(GJC_JOURNEE.Text));
    TOBJouEtab.PutValue('GJE_NBENTREES', GJE_NBENTREES.Value);
    TOBJouEtab.PutValue('GJE_METEO', GJE_METEO.Value);
    TOBJouEtab.PutValue('GJE_BLOCNOTE', GJE_BLOCNOTE.Text);
    MajJouEtab := True;
  end;
  // On conserve le numéro si des événements ont été choisis
  if (GET_CODEEVENT.Text = '') then TOBJouEtab.PutValue('GJE_NOEVENT', 0)
  else TOBJouEtab.PutValue('GJE_NOEVENT', idEvt);
  // Mise à jour de la base
  io := Transactions(EnregistreSaisie, 2);
  case io of
    oeOk: Result := True;
    oeUnknown: MessageAlerte(Msg.Mess[3]);
    oeSaisie: MessageAlerte(Msg.Mess[4]);
  else MessageAlerte(Msg.Mess[3]);
  end;
  // Conserve le code du dernier vendeur
  FOMajChampSupValeur(VH_GC.TOBPCaisse, 'LASTVENDEUR', GJC_VENDFERME.Text);
  // Sauvegarde des actions complémentaires à la fermeture de journée
  FOSaveInRegistry(REGJOURNEE, REGFERMETICKETZ, cbTicketZ.Checked);
  FOSaveInRegistry(REGJOURNEE, REGFERMERECAPVD, cbRecapVd.Checked);
  FOSaveInRegistry(REGJOURNEE, REGFERMEFDCAISSE, cbFdCaisse.Checked);
  FOSaveInRegistry(REGJOURNEE, REGFERMELISTEART, cbListeArt.Checked);
  FOSaveInRegistry(REGJOURNEE, REGFERMELISTEREG, cbListeReg.Checked);
  FOSaveInRegistry(REGJOURNEE, REGFERMEFOBO, cbFOBO.Checked);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EnregistreSaisie : enregistre les données saisies
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.EnregistreSaisie;
var Stg: string;
begin
  // la caisse doit être ouverte
  if FOEtatCaisse(GJC_CAISSE.Value, False) <> ETATJOURCAISSE[2] then
  begin
    V_PGI.IoError := oeSaisie;
    Exit;
  end;
  // suppression des tickets en attente
  if NbPieceAtt.Value > 0 then
  begin
    FOSupprimePieceEnAttente(GJC_CAISSE.Value);
    Stg := TraduireMemoire('Suppression de ') + NbPieceAtt.Text
      + ' tickets en attente de la caisse ' + GJC_CAISSE.Value;
    FOIncrJoursCaisse(jfoNbTicattSup, Stg, StrToInt(NbPieceAtt.Text));
  end;
  // enregistre la TOB de la jounée de la caisse
  TOBJCaisse.UpdateDateModif;
  if not TOBJCaisse.InsertOrUpdateDB(False) then
  begin
    V_PGI.IoError := oeUnknown;
    Exit;
  end;
  // enregistre la journée niveau établissement
  if MajJouEtab then
  begin
    // Suppression des événements saisis précédemment
    Stg := 'DELETE FROM JOURSETABEVT where GET_ETABLISSEMENT="' + GPK_ETABLISSEMENT.Value + '" '
      + 'and GET_NOEVENT=' + IntToStr(NoEvent);
    ExecuteSQL(Stg);
    // enregistre la TOB des événements saisis
    TOBJouEtab.UpdateDateModif;
    if not TOBJouEtab.InsertOrUpdateDB(False) then
    begin
      V_PGI.IoError := oeUnknown;
      Exit;
    end;
  end;
  // enregistre le résultat du contrôle caisse
  FOMAJTablesCtrlCais(TOBCtrlCais, GJC_CAISSE.Value, GJC_VENDFERME.Text, 'Contrôle caisse',
    TOBJCaisse.GetValue('GJC_NUMZCAISSE'), True);
  // crée les ticket de fond de caisse
  if (FOEstVrai(TOBCaisse, 'GPK_GEREFONDCAISSE')) and (V_PGI.IoError = oeOk) then CreeTicketFondCaisse;
  // JTR - RD encaissés
  Stg := 'UPDATE OPERCAISSE SET GOC_ANTERIEUR="-"'
       + ' WHERE GOC_ETABLISSEMENT="'+ TOBCaisse.GetValue('GPK_ETABLISSEMENT') +'"'
       + ' AND GOC_CAISSE="'+ TOBJcaisse.GetValue('GJC_CAISSE') +'"'
       + ' AND GOC_NUMZCAISSE='+ IntToStr(TOBJcaisse.GetValue('GJC_NUMZCAISSE'))
       + ' AND GOC_DATEPIECE="'+ USDateTime(TOBJcaisse.GetValue('GJC_DATEOUV')) +'"'
       + ' AND GOC_ANTERIEUR="X"';
  ExecuteSQL(Stg);
  // Fin JTR - RD encaissés
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ImprimeTicket : impression des tickets de clôture
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ImprimeTicket;
var sWhere: string;
  NumZ: Integer;
  TWait: Integer;
begin
  TWait := 0;
  if TicketGenere then TWait := FOGetFromRegistry(REGJOURNEE, REGTMODEBIMP, 5000);
  NumZ := TOBJCaisse.GetValue('GJC_NUMZCAISSE');
  if (cbTicketZ.Checked) and (FOExisteCodeEtat(efoTicketZ)) then
  begin
    if FOExisteLP then Delay(TWait);
    CompteRendu.lines.Append(TraduireMemoire('Impression du ticket Z'));
    sWhere := FOMakeWhereTicketZ('GP', GJC_CAISSE.Value, Numz, NumZ);
    FOLanceImprimeLP(efoTicketZ, sWhere, False, nil);
    TWait := FOGetFromRegistry(REGJOURNEE, REGTMOFINIMP, 3000);
  end;
  if (cbRecapVd.Checked) and (FOExisteCodeEtat(efoRecapVend)) then
  begin
    if FOExisteLP then Delay(TWait);
    CompteRendu.lines.Append(TraduireMemoire('Impression du récapitulatif par vendeur'));
    sWhere := FOMakeWhereRecapVendeurs('GP', GJC_CAISSE.Value, NumZ, NumZ);
    FOLanceImprimeLP(efoRecapVend, sWhere, False, nil);
    TWait := FOGetFromRegistry(REGJOURNEE, REGTMOFINIMP, 3000);
  end;
  if (cbFdCaisse.Checked) and (FOExisteCodeEtat(efoFdCaisse)) then
  begin
    if FOExisteLP then Delay(TWait);
    CompteRendu.lines.Append(TraduireMemoire('Impression du fonds de caisse'));
    sWhere := 'GJM_CAISSE="' + GJC_CAISSE.Value + '" and GJM_NUMZCAISSE="' + IntToStr(NumZ) + '"';
    FOLanceImprimeLP(efoFdCaisse, sWhere, False, nil);
  end;
  if (cbListeArt.Checked) and (FOExisteCodeEtat(efoListeArtVendu)) then
  begin
    if FOExisteLP then Delay(TWait);
    CompteRendu.lines.Append(TraduireMemoire('Impression de la liste des articles vendus'));
    sWhere := 'GL_TYPEARTICLE="MAR" AND ' + FOMakeWhereTicketZ('GP', GJC_CAISSE.Value, Numz, NumZ);
    FOLanceImprimeLP(efoListeArtVendu, sWhere, False, nil);
  end;
  if (cbListeReg.Checked) and (FOExisteCodeEtat(efoListeRegle)) then
  begin
    if FOExisteLP then Delay(TWait);
    CompteRendu.lines.Append(TraduireMemoire('Impression de la liste des règlements'));
    sWhere := FOMakeWhereTicketZ('GP', GJC_CAISSE.Value, Numz, NumZ);
    FOLanceImprimeLP(efoListeRegle, sWhere, False, nil);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////  
//  ChargeCompteRendu : Initialisation des zones du compte rendu
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ChargeCompteRendu;
begin
  CompteRendu.Text := '';
  CompteRendu.SelAttributes.Style := [fsBold];
  CompteRendu.lines.Append('');
  CompteRendu.lines.Append(TraduireMemoire('Vous allez fermer la journée du'));
  CompteRendu.SelAttributes.Style := [fsBold, fsItalic];
  CompteRendu.lines.Append(GJC_JOURNEE.Text);
  CompteRendu.SelAttributes.Style := [fsBold];
  CompteRendu.lines.Append('');
  CompteRendu.lines.Append(TraduireMemoire('pour la caisse'));
  CompteRendu.SelAttributes.Style := [fsBold, fsItalic];
  CompteRendu.lines.Append(GJC_CAISSE.Text);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.FormCreate(Sender: TObject);
begin
  inherited;
  // Creation des TOB de la caisse
  TOBCaisse := TOB.Create('PARCAISSE', nil, -1);
  TOBJCaisse := TOB.Create('JOURSCAISSE', nil, -1);
  TOBJouEtab := TOB.Create('JOURSETAB', nil, -1);
  TOBCtrlCais := TOB.Create('Contrôle caisse', nil, -1);
  CtrlCaisOk := False;
  MiseAttente := False;
  TicketGenere := False;
  BoutonEncours := False;
  LanceCpteurFait := False;
  MajDone := False;
  PnlBoutons := nil;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then WindowState := wsMaximized;
  // Calcul du nombres d'étapes
  NoPage := 1;
  NbPage := P.PageCount;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormClose
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.FormClose(Sender: TObject; var Action: TCloseAction);
var Ok: boolean;
  Stg: string;
begin
  inherited;
  // enregistre le résultat du contrôle caisse
  if not (MajDone) and (TOBCtrlCais.Detail.Count > 0) then
  begin
    if FOJaiLeDroit(86, False, False) then
    begin
      Ok := (Msg.Execute(10, Caption, '') = mrYes)
    end else
    begin
      Stg := FOMsgBoxExec(Msg, 10, Caption, '');
      Ok := FOJaiLeDroit(86, True, True, Stg);
    end;
    if Ok then
    begin
      FOMAJTablesCtrlCais(TOBCtrlCais, GJC_CAISSE.Value, GJC_VENDFERME.Text, 'Contrôle caisse',
        TOBJCaisse.GetValue('GJC_NUMZCAISSE'), False);
      MajDone := True;
    end;
  end;
  // Libération des TOB de la caisse
  TOBCaisse.Free;
  TOBJCaisse.Free;
  TOBJouEtab.Free;
  TOBCtrlCais.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.FormShow(Sender: TObject);
var Ind: Integer;
  CashDef: string;
begin
  inherited;
  // Lecture de la caisse choisie à la connexion
  CashDef := VH_GC.TOBPCaisse.GetValue('GPK_LIBELLE');
  Ind := GJC_CAISSE.Items.IndexOf(CashDef);
  if Ind >= 0 then GJC_CAISSE.ItemIndex := Ind else GJC_CAISSE.ItemIndex := 0;
  GJC_Caisse.OnChange(nil);
  // On propose la date de la denière ouverture
  ChargeJCaisse;
  GJC_JOURNEE.Text := DateToStr(TOBJcaisse.GetValue('GJC_DATEOUV'));
  // Activation de la saisie du vendeur
  if TOBJcaisse.FieldExists('GJC_VENDFERME') and not (FOEstVrai(TOBCaisse, 'GPK_VENDSAISIE')) and
    not (FOEstVrai(TOBCaisse, 'GPK_VENDSAISLIG')) then
  begin
    GJC_VENDFERME.Visible := False;
    TGJC_VENDFERME.Visible := False;
  end;
  // On rend invisibles les champs si les tablettes associées sont vides
  if GJE_METEO.Values.Count <= 1 then
  begin
    GJE_METEO.Visible := False;
    TGJE_METEO.Visible := False;
    GJE_NBENTREES.Top := GJE_METEO.Top;
    TGJE_NBENTREES.Top := TGJE_METEO.Top;
  end;
  if GET_CODEEVENT.Values.Count <= 0 then
  begin
    LST_CODEEVT.Visible := False;
    TLST_CODEEVT.Visible := False;
  end;
  // Calcul du nombre de pièce en attente
  CalculNbPieceAttente;
  // On inactive le bouton FIN
  ActivebFin;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormResize :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.FormResize(Sender: TObject);
var
  AHeight, BtnTopValue: Integer;
///////////////////////////////////////////////////////////////////////////////////////
  procedure PlaceBoutons(Btn: TToolbarButton97; ALeft: Integer);
  begin
    Btn.Height := 50;
    Btn.Width := 120;
    Btn.Top := Height - 78;
    Btn.Left := ALeft;
    Btn.Font.Size := 11;
  end;
  ////////////////////////////////////// ////////////////////////////////////////////////
  procedure PlaceBtnFleche(Btn: TToolbarButton97; Rang: Integer);
  begin
    Btn.Height := 50;
    Btn.Width := 60;
    Btn.Top := Height - 78;
    Btn.Left := (Width - (Rang * (Btn.Width + 1)) - 14);
    Btn.Font.Size := 11;
    Btn.Visible := True;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
begin
  inherited;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then
  begin
    lEtape.Visible := False;
    GroupBox1.Top := Height - 85;
    //JTR - 13/08/04 - Look 2003 *** ATTENTION, MODIFS A EFFECTUER DANS LE PROCHAIN AGL ***
 {   if V_PGI.Draw2003 then
    begin
      Dock971.Top := GroupBox1.Top + GroupBox1.height ;
      Dock971.Height := Height-Dock971.Top;
      ToolWindow971.Height := 70;
      BtnTopValue := 2;
    end else    }
      BtnTopValue := 78;

    PlaceBoutons(bAide, 20);
    PlaceBoutons(bAnnuler, 145);
    PlaceBoutons(bPrecedent, 271);
    PlaceBoutons(bSuivant, 392);
    PlaceBoutons(bFin, 518);
    PlaceBtnFleche(BLEFT, 4);
    PlaceBtnFleche(BRIGHT, 3);
    PlaceBtnFleche(BUP, 2);
    PlaceBtnFleche(BDOWN, 1);
    PPave.Left := 0;
    PPave.Width := Width;
    PPave.Top := P.Top + P.Height;
    AHeight := GroupBox1.Top - PPave.Top;
    if AHeight > 0 then PPave.Height := AHeight;
    PPave.Visible := True;
    if PnlBoutons = nil then
    begin
      FOCreateClavierEcran(PnlBoutons, Self, PPave, False, True);
      PnlBoutons.LanceBouton := SaisieClavierEcran;
      PnlBoutons.LanceCalculette := SaisieCalculette;
      PnlBoutons.BoutonCalculette := BoutonCalculetteClick;

      if (P.ActivePage.Name = 'PChoixCaisse') and (GJC_VENDFERME.Focused) then ChoixVendeurPave;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActivebFin : Activation du bouton FIN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ActivebFin;
var Ok: Boolean;
begin
  Ok := (P.ActivePage.Name = 'PResume');
  bFin.Enabled := Ok;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then
  begin
    bFin.Visible := Ok;
    BLEFT.Visible := (not Ok);
    BRIGHT.Visible := (not Ok);
    BUP.Visible := (not Ok);
    BDOWN.Visible := (not Ok);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bFinClick : bouton FIN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.bFinClick(Sender: TObject);
{$IFDEF TOXCLIENT}
var IdApplication: string;
  WhenStart: TDateTime;
  {$ENDIF}
begin
  inherited;
  if not bFin.Enabled then Exit;
  // Interdit les boutons
  bAide.Enabled := False;
  bAnnuler.Enabled := False;
  bPrecedent.Enabled := False;
  bSuivant.Enabled := False;
  bFin.Enabled := False;
  // Enregistre les données saisies
  if not SauveCaisse then
  begin
    Close;
    Exit;
  end;
  MajDone := True;
  // Impression des tickets de clôture
  CompteRendu.lines.Append('');
  ImprimeTicket;
  {$IFDEF TOXCLIENT}
  if (cbFOBO.Checked) and (FOVerifSiteTox) and
    (GJC_CAISSE.Value = Trim(GetParamSoc('SO_GCFOCAISREFTOX'))) then
  begin
    // Balayage des corbeilles de la TOX pour la suppression des anciens fichiers
    PurgeLesCorbeillesTox;
    // Lancement du Tox Serveur
    //if not AglToxServeurActifNet then
    if not AglToxCommuniCam(UST_RUNNING, nil) then
    begin
      //AglToxServeurRunNet ;
      AglToxCommuniCam(UST_START, nil);
      CompteRendu.lines.Append(TraduireMemoire('Lancement du serveur de communication ...'));
      Delay(5000); // attente de 5 secondes
    end;
    // Recherche de l'identifiant de l'application du ToxServeur
    IdApplication := FOGetToxIdApp;
    // Démarrage des échanges FO - BO
    CompteRendu.lines.Append(TraduireMemoire('Démarrage des échanges FO-BO'));
    AglToxConf(aceStart, IdApplication, GescomModeNotConfirmeTox, GescomModeTraitementTox, AglToxFormError);
    // Demande l'affichage de l'écran de mise en attente de communication (patience).
    Delay(1000);
    if AglStatusTox(WhenStart) then MiseAttente := True;
  end;
  CompteRendu.lines.Append('');
  {$ENDIF}
  Delay(2000);
  Close;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bSuivantClick : bouton SUIVANT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.bSuivantClick(Sender: TObject);
begin
  if not bSuivant.Enabled then Exit;
  // Verification du code vendeur
  if (P.ActivePage.Name = 'PChoixCaisse') and (not VerifieVendFerme) then Exit;
  // Contrôle des événements saisis
  if (P.ActivePage.Name = 'PEvent') and (not VerifieNbEntrees) then Exit;
  Inc(Nopage);
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bPrecedentClick : bouton PRECEDENT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.bPrecedentClick(Sender: TObject);
begin
  if not bPrecedent.Enabled then Exit;
  // Décrementation du n° de page avant traitement du changement de page
  Dec(Nopage);
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.PChange(Sender: TObject);
begin
  inherited;
  // Affiche le n° réelle de la page, en fonction du nombre de pages à traiter.
  lEtape.Caption := Msg.Mess[0] + ' ' + IntToStr(Nopage) + '/' + IntToStr(NbPage);
  // Affichage de l'image de l'onglet
  if (P.ActivePage.ImageIndex >= 0) and (P.ActivePage.ImageIndex < Image.Images.Count) then
  begin
    Image.Visible := True;
    Image.ImageIndex := P.ActivePage.ImageIndex;
  end else
  begin
    Image.Visible := False;
  end;
  // Initialisation des zones du compte rendu
  if (P.ActivePage.Name = 'PResume') then ChargeCompteRendu;
  // Activation du bouton FIN
  ActivebFin;
  // Efface le pavé tactile
  if PnlBoutons <> nil then PnlBoutons.PageCourante := 1;
  // Donne le focus au 1er champ de la page
  if (P.ActivePage.Name = 'PChoixCaisse') and (GJC_VENDFERME.CanFocus) then GJC_VENDFERME.SetFocus;
  if (P.ActivePage.Name = 'PEvent') and (GJE_NBENTREES.CanFocus) then
  begin
    LanceCompteurs;
    GJE_NBENTREES.SetFocus;
  end;
  if (P.ActivePage.Name = 'PBlocNote') and (GJE_BLOCNOTE.CanFocus) then GJE_BLOCNOTE.SetFocus;
  if (P.ActivePage.Name = 'PAction') and (cbTicketZ.CanFocus) then cbTicketZ.SetFocus;
  if (P.ActivePage.Name = 'PResume') and (CompteRendu.CanFocus) then CompteRendu.SetFocus;
  // Lancement du contrôle caisse
  if (P.ActivePage.Name = 'PCtrlCais') then LanceCtrlCaisse;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PreviousPage
///////////////////////////////////////////////////////////////////////////////////////

function TFFermeJour.PreviousPage: TTabSheet;
var TT: TTabsheet;
  Boucle: Integer;
begin
  TT := inherited PreviousPage;
  // Active ou non les onglets en fonction du paramétrage de la caisse
  Boucle := -1;
  while (not VerifiePageActive(TT)) do
  begin
    TT := P.Pages[P.ActivePage.PageIndex - 1 + Boucle];
    Dec(Boucle);
  end;
  Result := TT;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  NextPage
///////////////////////////////////////////////////////////////////////////////////////

function TFFermeJour.NextPage: TTabSheet;
var TT: TTabsheet;
  Boucle: Integer;
begin
  TT := inherited NextPage;
  // Active ou non les onglets en fonction du paramétrage de la caisse
  Boucle := 1;
  while (not VerifiePageActive(TT)) do
  begin
    TT := P.Pages[P.ActivePage.PageIndex + 1 + Boucle];
    Inc(Boucle);
  end;
  Result := TT;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifiePageActive : vérifie si une page doit être affichée
///////////////////////////////////////////////////////////////////////////////////////

function TFFermeJour.VerifiePageActive(TT: TTabSheet): Boolean;
begin
  Result := True;
  // Active ou non les onglets en fonction du paramétrage de la caisse
  if (TT <> nil) and (TOBCaisse <> nil) then
  begin
    if TT.Name = 'PPieceAtt' then
    begin
      Result := (NbPieceAtt.Value <> 0);
    end;
    if TT.Name = 'PCtrlCais' then
    begin
      Result := ((FOEstVrai(TOBCaisse, 'GPK_CTRLCAISSE')) and ((CtrlCaisOk) or (ExisteSQL(FOSelectCtrlCaisse))));
    end else
      if TT.Name = 'PEvent' then
    begin
      Result := FOEstVrai(TOBCaisse, 'GPK_GEREEVENT');
    end else
      if TT.Name = 'PBlocNote' then
    begin
      Result := FOEstVrai(TOBCaisse, 'GPK_GEREEVENT');
    end else
      if TT.Name = 'PAction' then
    begin
      Result := FOJaiLeDroit(82, False, False);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_CAISSEChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.GJC_CAISSEChange(Sender: TObject);
var Ind: Integer;
begin
  inherited;
  // Recherche de la caisse choisie
  ChargeCaisse;
  // Calcul du nombres d'étapes
  NbPage := 0;
  for Ind := 0 to P.PageCount - 1 do
    if VerifiePageActive(P.Pages[Ind]) then
      Inc(NbPage);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_JOURNEEChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.GJC_JOURNEEChange(Sender: TObject);
begin
  inherited;
  // Recherche de la journée établissement
  ChargeJouEtab;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChoixVendeurPave : choix du vendeur avec le pavé tactile
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.ChoixVendeurPave;
var sPlus: string;
begin
  if PnlBoutons <> nil then
  begin
    sPlus := 'GCL_TYPECOMMERCIAL="VEN" and GCL_ETABLISSEMENT="' + GPK_ETABLISSEMENT.Value + '"'
      + ' and GCL_DATESUPP>"' + USDateTime(Date) + '"';
    PnlBoutons.LanceChargePageMenu('VEN', '', sPlus, clBtnFace, 40);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_VENDFERMEEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.GJC_VENDFERMEEnter(Sender: TObject);
begin
  inherited;
  ChoixVendeurPave;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_VENDFERMEElipsisClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.GJC_VENDFERMEElipsisClick(Sender: TObject);
begin
  inherited;
  FOGetVendeur(GJC_VENDFERME, Msg.Mess[5], '', GPK_ETABLISSEMENT.Value);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_VENDFERMEDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.GJC_VENDFERMEDblClick(Sender: TObject);
begin
  inherited;
  FOGetVendeur(GJC_VENDFERME, Msg.Mess[5], '', GPK_ETABLISSEMENT.Value);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJE_NBENTREESChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.GJE_NBENTREESChange(Sender: TObject);
begin
  inherited;
  VerifieNbEntrees;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJE_METEOEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.GJE_METEOEnter(Sender: TObject);
begin
  inherited;
  if PnlBoutons <> nil then PnlBoutons.LanceChargePageMenu('MET', 'GCTYPEMETEO', '', clBtnFace, 31);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LST_CODEEVTEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.LST_CODEEVTEnter(Sender: TObject);
begin
  inherited;
  if PnlBoutons <> nil then PnlBoutons.LanceChargePageMenu('EVT', 'GCCODEEVENT', '', clBtnFace, 43);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Calcul du nombre de pièce en attente
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.CalculNbPieceAttente;
var QQ: TQuery;
  sSql: string;
begin
  NbPieceAtt.Value := 0;
  if FOGetParamCaisse('GPK_GERETICKETATT') = '-' then Exit;
  sSql := 'SELECT COUNT(*) FROM PIECE WHERE GP_NATUREPIECEG='+ FOGetNatureTicket(True, True)
        + ' AND GP_CAISSE="' + GJC_CAISSE.Value + '"';
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then NbPieceAtt.Value := QQ.Fields[0].AsInteger;
  Ferme(QQ);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LanceCtrlCaisse : Lancement du contrôle caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.LanceCtrlCaisse;
var sChoix, sArg: string;
begin
  if (P.ActivePage.Name = 'PCtrlCais') and (FOEstVrai(TOBCaisse, 'GPK_CTRLCAISSE')) then
  begin
    if (CtrlCaisOk) or (ExisteSQL(FOSelectCtrlCaisse)) then
    begin
      // chargement des éléments du contrôle caisse
      if not CtrlCaisOk then
      begin
        SourisSablier;
        FOChargeTablesCtrlCaisse(TOBCtrlCais, GJC_CAISSE.Value, TOBJCaisse.GetValue('GJC_NUMZCAISSE'));
        SourisNormale;
        CtrlCaisOk := True;
      end;
      // ouverture du tiroir caisse si nécessaire
      FOOuvreTiroir(False, True);
      // lancement du contrôle caisse
      if GJC_VENDFERME.Text <> '' then sArg := 'VENDEUR=' + GJC_VENDFERME.Text;
      TheTOB := TOBCtrlCais;
      sChoix := AGLLanceFiche('MFO', 'CTRLECAISSE', '', '', sArg);
      if TheTOB <> nil then
      begin
        TOBCtrlCais := TheTOB;
        TheTOB := nil;
      end;
    end;
    if sChoix = 'OK' then
      bSuivantClick(nil)
    else if sChoix = 'PREV' then
      bPrecedentClick(nil)
    else
    begin
      NormalClose := True;
      Close;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/04/2003
Modifié le ... : 23/04/2003
Description .. : Lancement du programme de collecte des données issues
Suite ........ : des compteurs d'entrées
Mots clefs ... :
*****************************************************************}

procedure TFFermeJour.LanceCompteurs;
var
  ExecName: string;
  TT: TFOProgressForm;
begin
  if LanceCpteurFait then Exit;
  LanceCpteurFait := True;

  ExecName := Trim(GetParamSoc('SO_GCFOEXECCPTENTREE'));
  if ExecName = '' then Exit;

  if GetParamSoc('SO_GCFOWAITCPTENTREE') then
  begin
    TT := TFOProgressForm.Create(Self, Caption, Msg.Mess[9] + ' ' + ExecName);
    FileExecAndWait(ExecName);
    TT.Free;
  end else
    FOLanceProg(ExecName);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieClavierEcran : interprète les boutons du pavé
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
var Ind: Integer;
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  if Concept = 'VEN' then // changement de vendeur
  begin
    if P.ActivePage.Name = 'PChoixCaisse' then
    begin
      if GJC_VENDFERME.CanFocus then GJC_VENDFERME.SetFocus;
      GJC_VENDFERME.Text := Code;
      FOSimuleClavier(VK_TAB);
    end;
  end else
    if Concept = 'MET' then // saisie d'un code météo
  begin
    if P.ActivePage.Name = 'PEvent' then
    begin
      if GJE_METEO.CanFocus then GJE_METEO.SetFocus;
      GJE_METEO.Value := Extra;
      FOSimuleClavier(VK_TAB);
    end;
  end else
    if Concept = 'EVT' then // saisie d'un code événement
  begin
    if P.ActivePage.Name = 'PEvent' then
    begin
      if LST_CODEEVT.CanFocus then LST_CODEEVT.SetFocus;
      Ind := GET_CODEEVENT.Values.IndexOf(Extra);
      if (Ind >= 0) and (Ind < LST_CODEEVT.Items.Count) then
        LST_CODEEVT.Checked[Ind] := (not LST_CODEEVT.Checked[Ind]);
      FOSimuleClavier(VK_TAB);
    end;
  end else
    ;
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BoutonCalculetteClick : interprète les boutons de la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.BoutonCalculetteClick(Val: string);
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  if Val = 'ENTER' then
  begin
    bSuivantClick(nil);
  end else
    if Val = 'CLEAR' then
  begin
    Close;
  end;
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieCalculette : prend en compte un chiffre saisi sur la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.SaisieCalculette(Val: Double);
var Texte: string;
  CC: TWinControl;
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  Texte := FloatToStr(Val);
  CC := Screen.ActiveControl;
  if CC is TCustomEdit then TCustomEdit(CC).Text := Texte;
  if CC is THValComboBox then THValComboBox(CC).Value := Texte;
  FOSimuleClavier(VK_TAB);
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BLEFTClick : OnClick sur le bouton BLEFT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.BLEFTClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_TAB, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BRIGHTClick : OnClick sur le bouton BRIGHT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.BRIGHTClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_TAB, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BUPClick : OnClick sur le bouton BUP
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.BUPClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_UP, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BDOWNClick : OnClick sur le bouton BDOWN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.BDOWNClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_DOWN, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BTicketAttenteClick : OnClick sur le bouton BTicketAttente
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.BTicketAttenteClick(Sender: TObject);
var
  Stg : string;
begin
  inherited;
  Stg := 'GP_NATUREPIECEG='+ FOGetNatureTicket(True, False) +';GP_CAISSE=' + GJC_CAISSE.Value ;
  AGLLanceFiche('MFO', 'PIECEATTENTE_MUL', Stg, '', 'CONSULTATION');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bAnnulerClick : OnClick sur le bouton bAnnuler
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFermeJour.bAnnulerClick(Sender: TObject);
begin
  inherited;
  NormalClose := False;
end;

end.
