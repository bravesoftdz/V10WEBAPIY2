unit AssistPL;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  ActnList, Mask, HPanel, UTOB, ImgList, Menus,
  HEnt1, LicUtil, MajTable, ParamSoc, HStatus,
  Devise_TOM,
  {$IFDEF COMPTA}
  CONTABON_TOM,
 {$ENDIF}
 {$IFDEF EAGLCLIENT}
  MaineAGL,
  uLanceProcess, // LanceProcessServer
  {$IFDEF COMPTA}
  RefAuto_Tom,
  {$ENDIF} 
 {$ELSE}
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Db,
  {$IFDEF COMPTA}
  Refauto_Tom,  // ParamLib  BVE 25.06.07 FQ 20165
  {$ENDIF}
  fe_main,
  {$IFDEF COMPTA}
  // scenario, BVE 21.05.07 FQ 19684
  SUIVCPTA_TOM,
  {$ENDIF}
  {$ENDIF}
  // Uses Commun
  FichComm,
  // Uses Compta
  Ent1, UTOFMULPARAMGEN, AssistExo,
  ImSaiCoef,// A FAIRE
  Soldecpt,
  GalOutil,
{$IFDEF NETEXPERT}
  UAssistImport, // ajout me 19/11/2004
{$ENDIF}
  Spin,
  utof_asscomptespe,
  utof_asscomptaaux,  
  utof_assprefcompta, HImgList;

function LanceAssistantComptaPCL(bInit: boolean): boolean;
procedure InitSocietePCL;

type
  TFAssistPL = class(TFAssist)
    ActionList1: TActionList;
    PDebut: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    HLabel1: THLabel;
    TSO_LIBELLE: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    bCoordonnees: TToolbarButton97;
    LABEL_CREATION: THLabel;
    SO_LIBELLE: TEdit;
    SO_CODE: TEdit;
    GroupBox2: TGroupBox;
    SO_NUMPLANREF: THValComboBox;
    HLabel4: THLabel;
    bParamComptable: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    bFourchettes: TToolbarButton97;
    GroupBox3: TGroupBox;
    HLabel8: THLabel;
    bMonnaie: TToolbarButton97;
    bComptaAux: TToolbarButton97;
    bDevise: TToolbarButton97;
    bPreferences: TToolbarButton97;
    bEtablissement: TToolbarButton97;
    GroupBox4: TGroupBox;
    HLabel2: THLabel;
    bAnalytique: TToolbarButton97;
    bBDSDynamique: TToolbarButton97;
    bTVA: TToolbarButton97;
    bTresorerie: TToolbarButton97;
    bTiers: TToolbarButton97;
    bSuivi: TToolbarButton97;
    bImmobilisation: TToolbarButton97;
    PopParam: TPopupMenu;
    Activer: TMenuItem;
    ImageList: THImageList;
    HLabel3: THLabel;
    PopParamPlan: TPopupMenu;
    PlanComptesGeneraux: TMenuItem;
    JournauxComptables: TMenuItem;
    LibelleAuto: TMenuItem;
    GuideSaisie: TMenuItem;
    FICHIER_IMPORT: THCritMaskEdit;
    OpenDialog: TOpenDialog;
    Desactiver: TMenuItem;
    bContrat: TToolbarButton97;
    MONNAIE_TENUE: THValComboBox;
    HLabel9: THLabel;
    BParamDevise: TToolbarButton97;
    Label3: TLabel;
    FICHENAME: THCritMaskEdit;
    TMODECREAT: THLabel;
    MODECREAT: THValComboBox;
    bSaisieQte: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure OnBoutonActivationClick(Sender: TObject);
    procedure ActiverClick(Sender: TObject);
    procedure bCoordonneesClick(Sender: TObject);
    procedure PlanComptesGenerauxClick(Sender: TObject);
    procedure JournauxComptablesClick(Sender: TObject);
    procedure LibelleAutoClick(Sender: TObject);
    procedure GuideSaisieClick(Sender: TObject);
    procedure bParamComptableClick(Sender: TObject);
    procedure bFourchettesClick(Sender: TObject);
    procedure bPreferencesClick(Sender: TObject);
    procedure bMonnaieClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure FICHIER_IMPORTClick(Sender: TObject);
    procedure DesactiverClick(Sender: TObject);
    procedure OnBoutonActivationMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure bButton97MouseEnter(Sender: TObject);
    procedure bComptaAuxClick(Sender: TObject);
    procedure BParamDeviseClick(Sender: TObject);
    procedure MODECREATChange(Sender: TObject);

  private
    { Déclarations privées }
    fNumPlan: integer;
    CurToolbarButton: TToolbarButton97;
    CurToolBarButtonName: string;
    bComptaActivee: boolean;
    procedure InitMission;
    function ChargeStandard(NumPlan: integer): boolean;
{$IFDEF EAGLCLIENT}
{$ELSE}
    function ChargeStandardCompta : Boolean;
    procedure ChargeParamSocRef(NumPlan: integer);
    procedure AssistCreatCompte;
    procedure InitMonnaieDossier(Monnaie: string);
    procedure CreatComptegene(Compte: string);
{$ENDIF}
    procedure ChargeComboNumPlan;
    // ajout me
    function DeversePlanComptable ( NumPlan : integer; Monnaie : string) : boolean ;

  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
{$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  CPTypeCons,
{$ENDIF MODENT1}
  UtilPgi,
{$IFDEF COMPTA}
  uTOFCPMulGuide,
  uTofCPBalSitDyna,
{$ENDIF}
  uLibStdCpta;

{ TFAssistCompta }

procedure InitTablesLibresPCL;
var TPl : TOB;
    QCom: TQuery;
    i : integer;
    bVide : boolean;
    cVisible : Char;
begin
  TPl := TOB.Create ('', nil, -1);
  try
    TPl.LoadDetailDB ('PARAMLIB','','',nil,False);
    // Est-ce que PARAMLIB est vide ? cas création dossier avec version antérieure stricte à 591
    bVide := (TPl.Detail.Count = 0);
    if TPl.Detail.Count = 0 then  // si oui
    begin
      QCom := OpenSQL('SELECT * FROM ##DP##.PARAMLIB',True);
      TPl.LoadDetailDB('PARAMLIB','','',QCom,False);
      Ferme(QCom);
    end;
    for i:=0 to TPl.Detail.Count - 1 do
    begin
      if  ((((TPl.Detail[i].GetValue('PL_VISIBLE')='X') and (not bVide)) or (TPl.Detail[i].GetValue('PL_TABLE')='E')) and
          ((TPl.Detail[i].GetValue('PL_CHAMP')='E_TABLE0') or (TPl.Detail[i].GetValue('PL_CHAMP')='E_TABLE1'))) then
          cVisible := 'X'
      else cVisible := '-';
      TPl.Detail[i].PutValue ('PL_VISIBLE', cVisible);
    end;
    TPl.SetAllModifie (True);
    if bVide then TPl.InsertDB(nil,True)
    else TPl.UpdateDB;
  finally
    Tpl.Free;
  end;
end;

// ### si bInit = True , alors on est au lancement de l'application
//       ==> le dossier n'est pas encore initialisé pour la compta
// ### si bInit = False, on est dans l'application.
//       ==> le dossier est déjà initialisé pour la compta.

function LanceAssistantComptaPCL(bInit: boolean): boolean;
var
  FAssistPL: TFAssistPL;
  bRetour: boolean;
begin
  bRetour := false;
  if GetFlagAppli('CCS5.EXE') then
    bRetour := True;
  if (((bRetour) and (not bInit)) or ((not bRetour) and (bInit))) then
  begin
    FAssistPL := TFAssistPL.Create(Application);
    try
      InitSocietePCL;
      InitEtablissement;
      FAssistPL.bComptaActivee := bRetour;
      FAssistPL.InitMission;
      FAssistPL.ShowModal;
    finally
      FAssistPL.Free;
    end;
    bRetour := GetFlagAppli('CCS5.EXE');
  end;
  if bRetour then
  begin
    if not ExisteSQL('SELECT * FROM EXERCICE') then
      LanceAssistExo;
  end;
  result := ((bRetour) and (ExisteSQL('SELECT * FROM EXERCICE')));
end;

procedure TFAssistPL.FormShow(Sender: TObject);
var
  NumPlan: integer;
begin
  inherited;
  // Si on n'est pas dans la compta, on n'affiche pas l'onglet 4 pour éviter de compiler
  // des inutiles.
  // Basculement des standards en euro.
{$IFDEF IMP}
  TabSheet4.TabVisible := False;
{$ENDIF}
{$IFDEF COMPTA}
TabSheet4.TabVisible := False;
TabSheet3.TabVisible := False;
LibelleAuto.Visible  := False;
GuideSaisie.Visible  := False;
{$ENDIF}
  BasculeStandardEuro;
  if not StandardEnEuro then
  begin
    Close;
  end;
  SO_CODE.Text := GetParamSocSecur('SO_SOCIETE','');
  SO_LIBELLE.Text := GetParamSocSecur('SO_LIBELLE','');
  bAnnuler.Enabled := true;
  bPrecedent.Enabled := FALSE;
  bSuivant.Enabled := true;
  bFin.Enabled := true;
  NumPlan := GetParamSocSecur('SO_NUMPLANREF',0);
    // Pour l'instant pas d'autre mode de création - 19/01/2001
  SO_NUMPLANREF.Enabled := (not (bComptaActivee)) and (not (NumPlan > 0));
  MONNAIE_TENUE.Enabled := (not (bComptaActivee)) and (not (NumPlan > 0));
  BParamDevise.Enabled := MONNAIE_TENUE.Enabled;
  // ajout me 30/11/2004
  MODECREAT.Enabled := (not (bComptaActivee)) and (not (NumPlan > 0));
  MONNAIE_TENUE.Enabled := (not (bComptaActivee)) and (not (NumPlan > 0));

  // AJOUT ME 21/05/01  SO_LGCPTE.Enabled := (not (bComptaActivee)) and (not (NumPlan > 0));
    // CA - 07/12/2001 - Euro par défaut. ( case à cocher euro cochée par défaut )
  ChargeComboNumPlan;
  if NumPlan > 0 then
    SO_NUMPLANREF.Value := Format('%.03d', [integer(GetParamSocSecur('SO_NUMPLANREF','0'))])
  else if bComptaActivee then
    SO_NUMPLANREF.Value := '000'
  else
    SO_NUMPLANREF.Value := '007'; // Plan par défaut
  fNumplan := StrToint(SO_NUMPLANREF.Value);
  if MONNAIE_TENUE.Enabled then MONNAIE_TENUE.Value := 'EUR'
  else MONNAIE_TENUE.Value := GetParamSocSecur ('SO_DEVISEPRINC','');
  { Suppression des monnaie 'In' }
  ExecuteSQL ('DELETE FROM DEVISE WHERE D_MONNAIEIN="X" AND D_FONGIBLE="-"');
  AvertirMultiTable ('TTDEVISE');
  THValComboBox(MONNAIE_TENUE).Reload;
{$IFDEF NETEXPERT}
  FICHENAME.Visible := (MODECREAT.Value='FIC');
  Label3.Visible    :=  (MODECREAT.Value='FIC');
  SO_NUMPLANREF.Enabled := (MODECREAT.Enabled) and (MODECREAT.Value<>'FIC');
  MONNAIE_TENUE.Enabled := (MODECREAT.Enabled) and (MODECREAT.Value<>'FIC');
  BParamDevise.Enabled := MONNAIE_TENUE.Enabled;
{$ENDIF}
end;

procedure TFAssistPL.bAnnulerClick(Sender: TObject);
begin
  inherited;
  Close;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/09/2002
Modifié le ... : 24/10/2005
Description .. : - 17/09/2002 - ajout de la gestion des scenarios
Suite ........ : - 24/10/2005 - LG - ajout de la gestion des qt
Mots clefs ... :
*****************************************************************}
procedure TFAssistPL.OnBoutonActivationClick(Sender: TObject);
begin
{$IFDEF COMPTA}

  if Sender.ClassName <> 'TToolbarButton97' then
    exit;

  CurToolBarButtonName := TToolbarButton97(Sender).Name;
  CurToolbarButton := TToolbarButton97(Sender);
  if TToolbarButton97(Sender).ImageIndex = 0 then
  begin
    PopParam.Items[0].Visible := True;
    PopParam.Items[1].Visible := False;
    PopParam.Popup(mouse.CursorPos.x, mouse.CursorPos.y);
  end;

  if TToolbarButton97(Sender).ImageIndex = 1 then
  begin
    inherited;
    if CurToolBarButtonName = 'bDevise' then
      FicheDevise('', taModif, True)
    else if CurToolBarButtonName = 'bContrat' then
{$IFNDEF EAGLCLIENT}
      ParamAbonnement(TRUE, '', taModif)
{$ENDIF}
    else if CurToolBarButtonName = 'bEtablissement' then
      FicheEtablissement_AGL(taCreat)
    else if CurToolBarButtonName = 'bAnalytique' then
    begin
      // ajout me
      // GCO - 24/09/2007 - Force la lecture en base du PARAMSOC
      if GetParamSocSecur('SO_CPPCLSANSANA', True, True) = FALSE then
      begin
        SetParamSoc('SO_ZSAISIEANAL', TRUE);
        SetParamSoc('SO_ZGEREANAL', TRUE);
      end
      else
      begin
        SetParamSoc('SO_ZSAISIEANAL', FALSE);
        SetParamSoc('SO_ZGEREANAL', FALSE);
      end;
      {$IFDEF COMPTA}
      MAJHistoparam ('PARAMSOC', ''); //   AJOUT ME cas synchro BL
      {$ENDIF}
      AGLLanceFiche('CP', 'ASSANALYTIQUE', '', '', '');
    end
    else
    if CurToolBarButtonName = 'bBDSDynamique' then
    begin
      {$IFDEF COMPTA}
      CPLanceFiche_CPBALSITDYNA;
      {$ENDIF}
    end
    else
    if CurToolBarButtonName = 'bTVA' then
    begin
      if not GetParamSocSecur('SO_CPPCLSAISIETVA',False) then
      begin
        SetParamSoc('SO_CPPCLSAISIETVA', True);
        // Ajout GCO - 06/12/2005
        if GetParamSocSecur('SO_CPPCLSANSANA', True, True) then
        begin
          SetParamSoc('SO_ZSAISIEANAL', True);
          SetParamSoc('SO_ZGEREANAL', True);
          SetParamSoc('SO_CPPCLSANSANA', False);
        end;
      end;
    end
    else if CurToolBarButtonName = 'bTresorerie' then
      MessageAlerte('Non disponible')
    else if CurToolBarButtonName = 'bTiers' then
      MessageAlerte('Non disponible')
  {$IFNDEF EAGLCLIENT}
    else if CurToolBarButtonName = 'bSuivi' then
      MultiScenario
  {$ENDIF}
    else if CurToolBarButtonName = 'bImmobilisation' then
      MessageAlerte('Non disponible')
    else if CurToolBarButtonName = 'bSaisieQte' then
      SetParamSoc('SO_CPPCLSAISIEQTE', not GetParamSocSecur('SO_CPPCLSAISIEQTE', False) )
  end;

{$ENDIF}

end;

procedure TFAssistPL.bCoordonneesClick(Sender: TObject);
begin
  inherited;
{$IFNDEF EAGLCLIENT}
  ParamSociete(False, '', 'SCO_COORDONNEES', '', nil, nil, nil, nil, 0, taConsult);
{$ENDIF}
end;

procedure TFAssistPL.PlanComptesGenerauxClick(Sender: TObject);
begin
  inherited;
  CCLanceFiche_MULGeneraux('C;7112000');
end;

procedure TFAssistPL.JournauxComptablesClick(Sender: TObject);
begin
  inherited;
  CPLanceFiche_MULJournal('C;7205000');
end;

procedure TFAssistPL.LibelleAutoClick(Sender: TObject);
begin
  inherited;
  {$IFDEF COMPTA}
  ParamLibelle;
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 12/01/2006
Modifié le ... :   /  /
Description .. : - LG - 12/01/2005 - modif de l'appel de la fct pour ouvrir les 
Suite ........ : guides
Mots clefs ... : 
*****************************************************************}
procedure TFAssistPL.GuideSaisieClick(Sender: TObject);
begin
  inherited;
  {$IFDEF COMPTA}
  CPLanceFiche_MulGuide('','','') ;
  {$ENDIF}
end;

procedure TFAssistPL.bParamComptableClick(Sender: TObject);
begin
  inherited;
  // AJOUT ME 21/05/01
  AGLLanceFiche('CP', 'ASSCOMPTESPE', '', '', 'CREATION');
  ChargeSocieteHalley;
end;

procedure TFAssistPL.bFourchettesClick(Sender: TObject);
begin
  inherited;
{$IFDEF EAGLCLIENT}
  // A FAIRE
{$ELSE}
  ParamSociete(False, '', 'SCO_FOURCHETTES', '', nil, nil, nil, nil, 0);
{$ENDIF}
  ChargeSocieteHalley;
end;

procedure TFAssistPL.bMonnaieClick(Sender: TObject);
begin
  inherited;
{$IFDEF EAGLCLIENT}
  // A FAIRE
{$ELSE}
  ParamSociete(False, '', 'SCO_EURO', '', nil, nil, nil, nil, 0);
{$ENDIF}
end;

procedure TFAssistPL.bPreferencesClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('CP', 'ASSPREFCOMPTA', '', '', '');
end;

// ### CA le 07/08/2000 - A faire jusqu'à correction dans le lanceur.
// ### à supprimer ensuite.

procedure InitSocietePCL;
var
  OB, OB_DETAIL: TOB;
begin
  // Création de l'enregistrement
  OB := TOB.Create('La société', nil, -1);
  OB_DETAIL := TOB.Create('SOCIETE', OB, -1);
  OB_DETAIL.PutValue('SO_SOCIETE', GetParamSocSecur('SO_SOCIETE',''));
  OB_DETAIL.PutValue('SO_LIBELLE', GetParamSocSecur('SO_LIBELLE','')); // CA - 27/11/2001
  OB_DETAIL.PutValue('SO_ADRESSE1', GetParamSocSecur('SO_ADRESSE1',''));
  OB_DETAIL.PutValue('SO_ADRESSE2', GetParamSocSecur('SO_ADRESSE2',''));
  OB_DETAIL.PutValue('SO_ADRESSE3', GetParamSocSecur('SO_ADRESSE3',''));
  OB_DETAIL.PutValue('SO_CODEPOSTAL', GetParamSocSecur('SO_CODEPOSTAL',''));
  OB_DETAIL.PutValue('SO_VILLE', GetParamSocSecur('SO_VILLE',''));
  OB_DETAIL.PutValue('SO_PAYS', GetParamSocSecur('SO_PAYS',''));
  OB_DETAIL.PutValue('SO_TELEPHONE', GetParamSocSecur('SO_TELEPHONE',''));
  OB_DETAIL.PutValue('SO_FAX', GetParamSocSecur('SO_FAX',''));
  OB_DETAIL.PutValue('SO_TELEX', GetParamSocSecur('SO_TELEX',''));
  OB_DETAIL.PutValue('SO_SIRET', GetParamSocSecur('SO_SIRET',''));
  OB_DETAIL.PutValue('SO_APE', GetParamSocSecur('SO_APE',''));
  OB.InsertOrUpdateDB(True);
  OB.Free;
end;


procedure TFAssistPL.bSuivantClick(Sender: TObject);
var
  bRetour: boolean;
begin
  bRetour := True;
  bPrecedent.Enabled := true;
  if ((P.ActivePage = PDebut) and (SO_NUMPLANREF.Enabled)) then
  begin
    if (MODECREAT.Value='FIC') then
        bFinClick(Sender)
    else
        bRetour := DeversePlanComptable(StrToInt(SO_NUMPLANREF.Value),MONNAIE_TENUE.Value);
  end;
  if bRetour then
    inherited;
end;

procedure TFAssistPL.bFinClick(Sender: TObject);
var bRetour: boolean;
begin
  bRetour := True;
  if ((P.ActivePage = PDebut) and (SO_NUMPLANREF.Enabled)) then
  begin
{$IFDEF NETEXPERT}
    if (MODECREAT.Value='FIC') then
    begin
       if FICHENAME.Text = '' then
       begin
          PGIInfo ('Veuillez renseigner le fichier');
          P.ActivePage := PDebut;
          FICHENAME.SetFocus ;
          exit;
       end;
       bRetour := ImportDonnees(FICHENAME.Text+';Minimized',TRUE);
       if not bRetour then
       PGIBox('Chargement du fichier impossible. Veuillez recommencer.', 'Chargement d''un fichier');
    end
    else
{$ENDIF}
       bRetour := DeversePlanComptable(StrToInt(SO_NUMPLANREF.Value), MONNAIE_TENUE.Value);
  end
  else
  begin
{$IFDEF NETEXPERT}
       if (FICHENAME.Text <> '') then
       begin
             bRetour := ImportDonnees(FICHENAME.Text+';Minimized',TRUE);
             if not bRetour then
             PGIBox('Chargement du fichier impossible. Veuillez recommencer.', 'Chargement d''un fichier');
       end;
{$ENDIF}
  end;

  if bRetour then
  begin
    inherited;
    SetFlagAppli ('CCS5.EXE',True);
    SetFlagAppli ('COMSX.EXE',True);
    ChargeMagHalley;
    Close;
  end;

  // ### Afficher Message d'erreur ###
end;

procedure TFAssistPL.FICHIER_IMPORTClick(Sender: TObject);
begin
  inherited;
  OpenDialog.Execute;
  FICHIER_IMPORT.Text := OpenDialog.FileName;
end;

function TFAssistPL.ChargeStandard(NumPlan: integer): boolean;
var bRetour: boolean;
{$IFDEF EAGLCLIENT}
    lTobParam  : Tob;
    lTobResult : Tob;
{$ENDIF}
begin
  // GCO - 27/05/2004
  // Test de la présence de filtres ancien format avant déversement du Standard (uniquement pour les standards non CEGID
  if (NumPlan > 20) and (PresenceAncienFiltreDansSTD( NumPlan )) then
  begin
    PgiInfo( MessageAncienFiltreDansSTD(NumPlan, 1), 'Chargement d''un standard');
    Result := False;
    Exit;
  end;

  if _Blocage(['nrBatch', 'nrCloture'], True, 'nrCloture') then
  begin
    Result := False;
    Exit;
  end;

  fNumPlan := NumPlan;

{$IFDEF EAGLCLIENT}
  //lTobParam := InitTobParamProcessServer;
  lTobParam := TOB.Create('',nil,-1);
  lTobResult := LanceProcessServer('cgiStdCpta', 'CHARGESTD', IntToStr(NumPlan), lTobParam, True ) ;
  if lTobResult.GetValue('RESULT') = 'PASOKCHARGE' then
{$ELSE}
  if not ChargeStandardCompta then
{$ENDIF}
  begin
    // Erreur durant la mise à jour depuis les standards
    PGIBox('Chargement du standard impossible. Veuillez recommencer.', 'Chargement d''un standard');
    bRetour := false;
  end
  else
  begin
    SO_NUMPLANREF.Enabled := false;
    SetParamSoc('SO_NUMPLANREF', StrToInt(SO_NUMPLANREF.Value));
    bRetour := True;
  end;
{$IFDEF EAGLCLIENT}
  FreeAndNil( lTobParam );
  FreeAndNil( lTobResult );
{$ENDIF}

  _Bloqueur('nrCloture', False);
  Result := bRetour;
end;

{$IFDEF EAGLCLIENT}
{$ELSE}
function TFAssistPL.ChargeStandardCompta : Boolean;
begin
  BeginTrans;
  try
    ChargeParamSocRef(fNumPlan);
    ChargeSocieteHalley;
    // AJOUT ME 21/05/01
    SetParamSoc('SO_TENUEEURO ', (MONNAIE_TENUE.Value='EUR'));
    InitMonnaieDossier ( MONNAIE_TENUE.Value );
    RechargeParamSoc;
    VideTablettes;

    if not ChargeToutLeStandard( fNumPlan, False, False) then
      Raise Exception.Create('');

    CreationSectionAttente; // Ajout CA le 19/01/2001

    // Installation des coefficients dégressifs Immo
    InstalleLesCoefficientsDegressifs; // CA - 10/01/2002

    InitMission;

    AssistCreatCompte; // Ajpout Me

    CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      RollBack;
      Result := False;
      PgiError('Erreur de requête SQL : ' + V_PGI.LastSQLError, 'Fonction : ChargeStandardCompta');
    end;
  end;
end;

procedure TFAssistPL.ChargeParamSocRef(NumPlan: integer);
var
  QParamSocRef: TQuery;
  StVal: string;
begin
  // AJOUT ME 21/05/01
  QParamSocRef := OpenSQL('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
    IntToStr(NumPlan) + ' AND PRR_SOCNOM="SO_LGCPTEGEN"', True);
  if not QParamSocRef.EOF then
    SetParamSoc('SO_LGCPTEGEN', QParamSocRef.FindField('PRR_SOCDATA').AsString)
  else
    SetParamSoc('SO_LGCPTEGEN', 6);
  ferme(QParamSocRef);

  QParamSocRef := OpenSQL('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
    IntToStr(NumPlan) + ' AND PRR_SOCNOM="SO_LGCPTEAUX"', True);
  if not QParamSocRef.EOF then
    SetParamSoc('SO_LGCPTEAUX', QParamSocRef.FindField('PRR_SOCDATA').AsString)
  else
    SetParamSoc('SO_LGCPTEAUX', 6);
  ferme(QParamSocRef);

  ChargeSocieteHalley;

  QParamSocRef := OpenSQL('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
    IntToStr(NumPlan) + ' ORDER BY PRR_SOCNOM', True);
  InitMove(QParamSocRef.RecordCount, 'Chargement des paramètres société');
  while not QParamSocRef.EOF do
  begin
    StVal := QParamSocRef.FindField('PRR_SOCDATA').AsString;
    if (QParamSocRef.FindField('PRR_COMPTE').AsString = 'G') then
      StVal := BourreLaDonc(stVal, fbGene)
    else if (QParamSocRef.FindField('PRR_COMPTE').AsString = 'T') then
      StVal := BourreLaDonc(stVal, fbAux);
    SetParamSoc(QParamSocRef.FindField('PRR_SOCNOM').AsString, StVal);
    QParamSocRef.Next;
    MoveCur(False);
  end;
  Ferme(QParamSocRef);
  FiniMove;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/10/2005
Modifié le ... :   /  /    
Description .. : - 24/10/2005 - LG - ajout de la gestion des qt
Mots clefs ... : 
*****************************************************************}
procedure TFAssistPL.InitMission;
begin
  if GetParamSocSecur('SO_CPPCLSANSCONTRAT',False) = True then
    bContrat.ImageIndex := 0
  else
    bContrat.ImageIndex := 1;

  if GetParamSocSecur('SO_CPPCLSANSDEVISE', False) = True then
    bDevise.ImageIndex := 0
  else
    bDevise.ImageIndex := 1;

  // GCO - 24/09/2007 - FQ 21275
  // Erreur car SO_CPPCLSANSANA vaut "" au moment du lancement de l'assistant
  // donc correction du paramsoc
  if GetColonneSQL('PARAMSOC', 'SOC_DATA', 'SOC_NOM = "SO_CPPCLSANSANA"') = '' then
    SetParamSoc('SO_CPPCLSANSANA', 'X');

  if GetParamSocSecur('SO_CPPCLSANSANA', True, True) = True then
    bAnalytique.ImageIndex := 0
  else
    bAnalytique.ImageIndex := 1;

  if GetParamSocSecur('SO_CPPCLSANSANNEXE', False) = True then
    bSuivi.ImageIndex := 0
  else
    bSuivi.ImageIndex := 1 ;

  if GetParamSocSecur('SO_CPPCLSAISIEQTE', False) = True then
    bSaisieQte.ImageIndex := 1
  else
    bSaisieQte.ImageIndex := 0 ;

  if GetParamSocSecur('SO_CPPCLSAISIETVA',False) = True then
    bTva.ImageIndex := 1
  else
    bTva.ImageIndex := 0 ;

  if GetParamSocSecur('SO_CPBDSDYNA', False) = True then
    bBDSDynamique.ImageIndex := 1
  else
    bBDSDynamique.ImageIndex := 0;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/10/2005
Modifié le ... : 19/07/2006
Description .. : - 24/10/2005 - LG - ajout de la gestion des qt
Suite ........ : - 19/07/2006 - LG - si la tv n'est aps active, on mets la 
Suite ........ : bonne icone sur le bp tva
Mots clefs ... : 
*****************************************************************}
procedure TFAssistPL.ActiverClick(Sender: TObject);
begin
  inherited;
  CurToolBarButton.ImageIndex := 1;
  if CurToolBarButtonName = 'bDevise' then
    SetParamSoc('SO_CPPCLSANSDEVISE', False)
  else if CurToolBarButtonName = 'bContrat' then
    SetParamSoc('SO_CPPCLSANSCONTRAT', False)
  else if CurToolBarButtonName = 'bAnalytique' then
  begin
    // ajout me
    SetParamSoc('SO_ZSAISIEANAL', TRUE);
    SetParamSoc('SO_ZGEREANAL', TRUE);
    SetParamSoc('SO_CPPCLSANSANA', False);
                // ajout me E_IO
    ExecuteSQL('UPDATE  ECRITURE SET E_IO = "X"');
    {$IFDEF COMPTA}
      MAJHistoparam ('PARAMSOC', ''); //   AJOUT ME cas synchro BL
    {$ENDIF}
  end else if CurToolBarButtonName = 'bSuivi' then
  begin
    // Activation des tables libres
    SetParamSoc('SO_CPPCLSANSANNEXE', False);
    InitTablesLibresPCL;
  end
  else if CurToolBarButtonName = 'bSaisieQte' then
    SetParamSoc('SO_CPPCLSAISIEQTE', true )
  else
  if CurToolBarButtonName = 'bTVA' then
  begin
    if not ActivationTvaSurDossier then
     CurToolBarButton.ImageIndex := 0 ;
  end
  else
  if CurToolBarButtonName = 'bBDSDynamique' then
  begin
    {$IFDEF COMPTA}
    SetParamSoc('SO_CPBDSDYNA', True);
    CPLanceFiche_CPBALSITDYNA;
    {$ENDIF}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/10/2005
Modifié le ... :   /  /    
Description .. : - 24/10/2005 - LG - ajout de la gestion des qt
Mots clefs ... : 
*****************************************************************}
procedure TFAssistPL.DesactiverClick(Sender: TObject);
begin
  inherited;
  CurToolBarButton.ImageIndex := 0;
  if CurToolBarButtonName = 'bDevise' then
    SetParamSoc('SO_CPPCLSANSDEVISE', True)
  else
  if CurToolBarButtonName = 'bSaisieQte' then
     SetParamSoc('SO_CPPCLSAISIEQTE', false )
  else
  if CurToolBarButtonName = 'bTVA' then
  begin
   DesactivationTvasurDossier ;
    //Pgiinfo('Désactivation non disponible dans cette version.', 'Gestion de la TVA');
   // bTva.ImageIndex := 1;
    //SetParamSoc('SO_CPPCLSAISIETVA', false )
  end
  else
  if CurToolBarButtonName = 'bBDSDynamique' then
  begin
    if PgiAsk('Attention, les balances de situation dynamiques vont devenir statiques".' + #13#10 +
              'Voulez-vous continuer?', 'Balances de situation dynamiques') = MrYes then
    begin
      SetParamSoc('SO_CPBDSDYNA', False);
      SetParamSoc('SO_CPBDSMENSUELLE', False);
      SetParamSoc('SO_CPBDSTRIMESTRIELLE', False);
      SetParamSoc('SO_CPBDSSEMESTRIELLE', False);
      SetParamSoc('SO_CPBDSANNUELLE', False);
      ExecuteSQL('UPDATE CBALSIT SET BSI_TYPEBAL = "BDS" WHERE BSI_TYPEBAL = "DYN"');
    end
    else
      bBDSDynamique.ImageIndex := 1;
  end
  else if CurToolBarButtonName = 'bContrat' then
    SetParamSoc('SO_CPPCLSANSCONTRAT', True)
  else if CurToolBarButtonName = 'bSuivi' then
    SetParamSoc('SO_CPPCLSANSANNEXE', True)
  else if CurToolBarButtonName = 'bAnalytique' then
  begin
   if GetParamSocSecur('SO_CPPCLSAISIETVA',False) then
     begin
      PGIInfo('Vous ne pouvez pas désactiver l''analytique avec la TVA activée') ;
      CurToolBarButton.ImageIndex := 1;
      exit ;
     end ;
    if HShowMessage('0;Suppression analytiques;' +
      'Confimez-vous la suppression de l''ensemble des données analytiques' +
      ';Q;YN;N;N;', '', '') <> mrYes then
    begin
      CurToolBarButton.ImageIndex := 1;
      exit;
    end
    else
    begin
      if (GetParamSocSecur('SO_CPLIENGAMME','') = 'S1') and (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE) = TRUE) then
      begin
           if HShowMessage('0;Suppression analytiques;' +
           'Les données analytiques que vous apprêtez à supprimer intégralement ' +#10#13 +
           'le seront aussi dans le dossier Business Line de votre client.' +#10#13 +
           'Confimez-vous la suppression de l''ensemble des données analytiques' +
                          ';Q;YN;N;N;', '', '') <> mrYes then
           begin
            CurToolBarButton.ImageIndex := 1;
            exit;
           end;
      end;
      DesactivationAnalytique;
    end;
  end;
end;

procedure TFAssistPL.OnBoutonActivationMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if CurToolBarButton = nil then
    exit;
  if ((Button = mbRight) and (CurToolBarButton.ImageIndex = 1)) then
  begin
    PopParam.Items[0].Visible := False;
    PopParam.Items[1].Visible := True;
    PopParam.Popup(mouse.CursorPos.x, mouse.CursorPos.y);
  end;
  inherited;
end;

procedure TFAssistPL.bButton97MouseEnter(Sender: TObject);
begin
  inherited;
  CurToolbarButton := TToolbarButton97(Sender);
  CurToolBarButtonName := TToolbarButton97(Sender).Name;
end;

procedure TFAssistPL.ChargeComboNumPlan;
var
  Q: TQuery;
begin
  if bComptaActivee then
  begin
    // Pas de plan comptable : plan = 0
    SO_NUMPLANREF.Items.Add(TraduireMemoire('Pas de standard'));
    SO_NUMPLANREF.Values.Add('000');
  end;
  Q := OpenSQL('SELECT * FROM STDCPTA WHERE STC_PREDEFINI<>""', True);
  while not Q.Eof do
  begin
    SO_NUMPLANREF.Items.Add(Q.FindField('STC_LIBELLE').AsString);
    SO_NUMPLANREF.Values.Add(Format('%.03d',
      [Q.FindField('STC_NUMPLAN').AsInteger]));
    Q.Next;
  end;
  Ferme(Q);
end;

procedure TFAssistPL.bComptaAuxClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('CP', 'ASSCOMPTAAUX', '', '', '');
end;

{$IFDEF EAGLCLIENT}
{$ELSE}
procedure TFAssistPL.InitMonnaieDossier ( Monnaie : string );
var
  OBDevise, OBDetail: TOB;
  QDev : TQuery;
begin
  SetParamSoc('SO_TENUEEURO', (Monnaie='EUR'));
  if Monnaie='EUR' then SetParamSoc('SO_TAUXEURO', 1)
  else SetParamSoc('SO_TAUXEURO', 6.55957);
  OBDevise := TOB.Create('', nil, -1);
  OBDevise.LoadDetailDB('DEVISE', '', '', nil, True);
  OBDetail := OBDevise.FindFirst(['D_DEVISE'], ['EUR'], False);
  if Monnaie = 'EUR' then  // Initialisation du contexte devise Euro
  begin
    if (OBDetail = nil) then  // Le cas échéant, création de la devise Euro
    begin
      OBDetail := TOB.Create('DEVISE', OBDevise, -1);
      OBDetail.PutValue('D_DEVISE', 'EUR');
      OBDetail.PutValue('D_LIBELLE', 'Euro');
      OBDetail.PutValue('D_SYMBOLE', '€');
      OBDetail.PutValue('D_FERME', '-');
      OBDetail.PutValue('D_DECIMALE', 2);
      OBDetail.PutValue('D_QUOTITE', 1);
      OBDetail.PutValue('D_SOCIETE', V_PGI.CodeSociete);
      OBDetail.PutValue('D_MONNAIEIN', '-');
      OBDetail.PutValue('D_FONGIBLE', '-');
      OBDetail.PutValue('D_PARITEEURO', 1);
    end
    else   // Sinon mise à jour de la devise euro
    begin
      OBDetail.PutValue('D_FERME', '-');
      OBDetail.PutValue('D_DECIMALE', 2);
      OBDetail.PutValue('D_QUOTITE', 1);
      OBDetail.PutValue('D_MONNAIEIN', '-');
      OBDetail.PutValue('D_FONGIBLE', '-');
      OBDetail.PutValue('D_PARITEEURO', 1);
    end;
    OBDevise.InsertOrUpdateDB(True);
    SetParamSoc('SO_DEVISEPRINC', 'EUR');
  end
  else
  begin
    if OBDetail <> nil then
    begin
      OBDetail.PutValue('D_FERME', '-');
      OBDetail.PutValue('D_DECIMALE', 2);
      OBDetail.PutValue('D_QUOTITE', 1);
      OBDetail.PutValue('D_MONNAIEIN', '-');
      OBDetail.PutValue('D_FONGIBLE', '-');
      OBDetail.PutValue('D_PARITEEURO', 1);
      OBDevise.UpdateDB(False, True);
    end;
    { Suppression de la devise FRF }
    ExecuteSQL ('DELETE FROM DEVISE WHERE D_DEVISE="FRF"');
    SetParamSoc('SO_DEVISEPRINC', Monnaie);
    QDev:=OpenSQL('SELECT D_DECIMALE FROM DEVISE WHERE D_DEVISE="'+Monnaie+'"',True) ;
    If Not QDev.EOF then SetParamSoc('SO_DECVALEUR',QDev.Fields[0].AsInteger) ;
    Ferme(QDev) ;
  end;
  OBDevise.Free;
  // rechargement des tablette devises
  AvertirTable('TTDEVISE');
  AvertirTable('TTDEVISETOUTES');
  AvertirTable('TTDEVISEETAT');
  AvertirTable('TTDEVISEOUT');
end;

procedure TFAssistPL.AssistCreatCompte;
var
  Q1, QDos: TQuery;
  i, Lggene: integer;
  XWhere, Compte: string;
  QStd: TQuery;
  Suffixe, ChampDos, ValChamp: string;
  where: string;
begin
  // vérification des comptes dans les guides
  Q1 :=
    OpenSql('SELECT EG_GENERAL, EG_AUXILIAIRE,EG_TYPE, EG_GUIDE, EG_NUMLIGNE from ECRGUI', TRUE);
  InitMove(Q1.RecordCount, 'Vérification Comptes');
  while not Q1.EOF do
  begin
    ValChamp := BourreLaDonc(Q1.Fields[0].asstring, fbGene);
    if Length(ValChamp) > GetParamSocSecur('SO_LGCPTEGEN',0) then
    begin // AJOUT ME 21/05/01
      ValChamp := copy(Q1.Fields[0].asstring, 0,
        StrToInt(GetParamSocSecur('SO_LGCPTEGEN','0')));
      where := '" where EG_TYPE="' + Q1.findfield('EG_Type').asstring +
        '" and EG_GUIDE="' +
        Q1.findfield('EG_GUIDE').asstring + '" and EG_NUMLIGNE=' +
          IntToStr(Q1.findfield('EG_NUMLIGNE').asinteger);
      ExecuteSQL('UPDATE  ECRGUI SET EG_GENERAL="' + ValChamp +
        where);
    end;
    if (not Presence('GENERAUX', 'G_GENERAL', ValChamp))
      and (Q1.Fields[0].asstring <> '') then
      CreatComptegene(Q1.Fields[0].asstring);

    ValChamp := BourreLaDonc(Q1.Fields[1].asstring, fbAux);
    if Length(ValChamp) > GetParamSocSecur('SO_LGCPTEAUX',0) then
    begin // AJOUT ME 21/05/01
      ValChamp := copy(Q1.Fields[1].asstring, 0,
        StrToInt(GetParamSocSecur('SO_LGCPTEAUX','0')));
      where := '" where EG_TYPE="' + Q1.findfield('EG_Type').asstring +
        '" and EG_GUIDE="' +
        Q1.findfield('EG_GUIDE').asstring + '" and EG_NUMLIGNE=' +
          IntToStr(Q1.findfield('EG_NUMLIGNE').asinteger);
      ExecuteSQL('UPDATE  ECRGUI SET EG_AUXILIAIRE="' + ValChamp +
        where);
    end;

    if (not Presence('TIERS', 'T_AUXILIAIRE', ValChamp))
      and (Q1.Fields[1].asstring <> '') then
    begin
      XWhere := ' and ' + 'TRR_AUXILIAIRE="' + BourreLess(Q1.Fields[1].asstring,
        fbAux) + '"';
      QStd := OpenSQL('SELECT * FROM TIERSREF WHERE TRR_NUMPLAN=' +
        IntToStr(fNumPlan) + XWhere, True);
      if not QStd.Eof then // Si un tel standard existe
      begin
        QDos := OpenSQL('SELECT * FROM TIERS', FALSE);
        QDos.Insert;
        InitNew(QDos);
        for i := 1 to QStd.FieldCount - 1 do
        begin
          Suffixe := ExtractSuffixe(QStd.Fields[i].FieldName);
          if Suffixe = 'PREDEFINI' then
            continue
          else
            ChampDos := 'T_' + Suffixe;
          if ChampDos = 'T_AUXILIAIRE' then
            ValChamp := BourreLaDonc(QStd.FindField('TRR_AUXILIAIRE').AsString,
              fbAux)
          else if ChampDos = 'T_TIERS' then
            ValChamp := BourreLaDonc(QStd.FindField('TRR_TIERS').AsString, fbAux)
          else if ChampDos = 'T_COLLECTIF' then
            ValChamp := BourreLaDonc(QStd.FindField('TRR_COLLECTIF').AsString,
              fbGene)
          else
            ValChamp := QStd.FindField('TRR_' + Suffixe).AsVariant;
          QDos.FindField(ChampDos).AsVariant := ValChamp;
        end;
        QDOS.Post;
        Ferme(QDos);
      end;
      Ferme(QStd);
    end;

    Q1.next;
    MoveCur(False);
  end; // while
  Ferme(Q1);
  // vérification de l'existance du journal dans les guides
  Q1 := OpenSql('SELECT GU_JOURNAL from GUIDE', TRUE);
  InitMove(Q1.RecordCount, 'Vérification des guides');
  while not Q1.EOF do
  begin
    if not Presence('JOURNAL', 'J_JOURNAL', Q1.Fields[0].asstring) then
    begin
      XWhere := ' and ' + 'JR_JOURNAL="' + Q1.Fields[0].asstring + '"';
      QStd := OpenSQL('SELECT * FROM JALREF WHERE JR_NUMPLAN=' +
        IntToStr(fNumPlan) + XWhere, True);
      if not QStd.Eof then // Si un tel standard existe
      begin
        QDos := OpenSQL('SELECT * FROM JOURNAL', FALSE);
        QDos.Insert;
        InitNew(QDos);
        for i := 1 to QStd.FieldCount - 1 do
        begin
          Suffixe := ExtractSuffixe(QStd.Fields[i].FieldName);
          if Suffixe = 'PREDEFINI' then
            break
          else
            ChampDos := 'J_' + Suffixe;
          ValChamp := QStd.FindField('JR_' + Suffixe).AsVariant;
          QDos.FindField(ChampDos).AsVariant := ValChamp;
        end;
        QDOS.Post;
        Ferme(QDos);
      end;
      Ferme(QStd);
    end;
    Q1.next;
    MoveCur(False);
  end; // while
  Ferme(Q1);
  // vérification des journaux pour les comptes automatiques
  // de contrepartie , interdits et de régularisation
  Lggene := GetParamSocSecur('SO_LGCPTEGEN',0);
  Q1 := OpenSql('SELECT * from JOURNAL', TRUE);
  InitMove(Q1.RecordCount, 'Vérification du journal');
  while not Q1.EOF do
  begin
    Compte := BourreLaDonc(Q1.FindField('J_CONTREPARTIE').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_CONTREPARTIE').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_CONTREPARTIE').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_COMPTEINTERDIT').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_COMPTEINTERDIT').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_COMPTEINTERDIT').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_COMPTEAUTOMAT').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_COMPTEAUTOMAT').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_COMPTEAUTOMAT').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_CPTEREGULDEBIT1').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_CPTEREGULDEBIT1').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_CPTEREGULDEBIT1').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_CPTEREGULDEBIT2').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_CPTEREGULDEBIT2').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_CPTEREGULDEBIT2').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_CPTEREGULDEBIT3').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_CPTEREGULDEBIT3').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_CPTEREGULDEBIT3').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_CPTEREGULCREDIT1').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_CPTEREGULCREDIT1').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_CPTEREGULCREDIT1').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_CPTEREGULCREDIT2').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_CPTEREGULCREDIT2').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_CPTEREGULCREDIT2').asstring <> '') then
      CreatComptegene(Compte);

    Compte := BourreLaDonc(Q1.FindField('J_CPTEREGULCREDIT3').asstring, fbGene);
    if Length(Compte) > GetParamSocSecur('SO_LGCPTEGEN',0) then
      Compte := copy(Q1.FindField('J_CPTEREGULCREDIT3').asstring, 0, Lggene);

    if (not Presence('GENERAUX', 'G_GENERAL', Compte))
      and (Q1.FindField('J_CPTEREGULCREDIT3').asstring <> '') then
      CreatComptegene(Compte);
    Q1.next;
    MoveCur(False);
  end; // while
  Ferme(Q1);
  FiniMove;
end;

procedure TFAssistPL.CreatComptegene(Compte: string);
var
  XWhere: string;
  QStd, QDos: TQuery;
  i: integer;
  Suffixe: string;
  ValChamp: string;
  ChampDos: string;
begin
  XWhere := ' and ' + 'GER_GENERAL="' + BourreLess(Compte, fbGene) + '"';
  QStd := OpenSQL('SELECT * FROM GENERAUXREF WHERE GER_NUMPLAN=' +
    IntToStr(fNumPlan) + XWhere, True);
  if not QStd.Eof then // Si un tel standard existe
  begin
    QDos := OpenSQL('SELECT * FROM GENERAUX', FALSE);
    QDos.Insert;
    InitNew(QDos);
    for i := 1 to QStd.FieldCount - 1 do
    begin
      Suffixe := ExtractSuffixe(QStd.Fields[i].FieldName);
      if Suffixe = 'COMPTE' then
        ChampDos := 'G_GENERAL'
      else if Suffixe = 'REPORTDETAIL' then
        continue
      else if Suffixe = 'PREDEFINI' then
        break
      else
        ChampDos := 'G_' + Suffixe;
      if ChampDos = 'G_GENERAL' then
        ValChamp := BourreLaDonc(QStd.FindField('GER_GENERAL').AsString, fbGene)
      else
        ValChamp := QStd.FindField('GER_' + Suffixe).AsVariant;
      QDos.FindField(ChampDos).AsVariant := ValChamp;
    end;
    QDOS.Post;
    Ferme(QDos);
  end;
  Ferme(QStd);
end;
{$ENDIF}

procedure TFAssistPL.BParamDeviseClick(Sender: TObject);
begin
  inherited;
  FicheDevise('', taModif, True);
  AvertirMultiTable('TTDEVISE');
  THValComboBox(MONNAIE_TENUE).Reload;
end;

function TFAssistPL.DeversePlanComptable ( NumPlan : integer; Monnaie : string ) : boolean ;
begin
  SetParamSoc('SO_I_CPTAPGI',False);
  SetParamSoc('SO_LGCPTEGEN', 6);
  SetParamSoc('SO_LGCPTEAUX', 6);
  { On impose le bourrage des comptes avec le caractère '0' }
  SetParamSoc('SO_BOURREGEN', '0');
  SetParamSoc('SO_BOURREAUX', '0');
  { Rechargement des paramètres société }
  ChargeSocieteHalley;
  Result := ChargeStandard(NumPlan);
end;

procedure TFAssistPL.MODECREATChange(Sender: TObject);
begin
  inherited;
{$IFDEF NETEXPERT}
  FICHENAME.Visible := (MODECREAT.Value='FIC');
  Label3.Visible    := (MODECREAT.Value='FIC');
  SO_NUMPLANREF.Enabled :=(MODECREAT.Value<>'FIC');
  MONNAIE_TENUE.Enabled := (MODECREAT.Value<>'FIC');
  BParamDevise.Enabled := MONNAIE_TENUE.Enabled;
  bSuivant.Enabled := (MODECREAT.Value='PLA') or ((MODECREAT.Value='FIC') and (FileExists(FICHENAME.Text)));
  bFin.Enabled := (MODECREAT.Value='PLA') or ((MODECREAT.Value='FIC') and (FileExists(FICHENAME.Text)));
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////

end.





