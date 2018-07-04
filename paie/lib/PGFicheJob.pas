{***********UNITE*************************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : - XP 22-10-2004 : Bizarre, mais la séquence n'était pas reprise !!!
Mots clefs ... :
*****************************************************************}
unit PGFicheJob;

{$IFDEF HVCL}sdsd{$ENDIF}
{$IFDEF BASEEXT}sdfsdf{$ENDIF}

// Mul

interface

uses Windows,
  Messages,
  SysUtils,
  Classes,
  TradMini,
  Graphics,
  Controls,
  {$IFDEF eaglclient}
  {$ELSE}
  db,
  {$IFDEF ODBCDAC}
  odbcconnection,
  odbctable,
  odbcquery,
  odbcdac,
  {$ELSE}
  {$IFNDEF DBXPRESS}dbtables,
  {$ELSE}uDbxDataSet,
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
  Forms,
  Dialogs,
  Hent1,
  Hctrls,
  uTaskJob,
  hmsgbox,
  ExtCtrls,
  HPanel,
  HTB97,
  StdCtrls,
  utob,
  Mask,
  Spin,
  uLanceProcess,
  uFicheTache,
  HSysMenu;

function AglFicheJob(
  const ParamJobId: integer;
  const ParamAvecAction: TActionFiche;
  const ParamExeName, ParamAction: string;
  ParamTob: tob = nil;
  const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True): Boolean;

function AglFicheJobExt(
  const ParamJobId: integer;
  const ParamAvecAction: TActionFiche;
  const ParamExeName, ParamAction: string;
  ParamTob: tob = nil;
  const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const VisibleInsert: boolean = False;
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True): integer;

type
  TPFicheJob = class(THForm)
    PanelHaut: THPanel;
    PanelBas: TToolWindow97;
    Baide: TToolbarButton97;
    BFerme: TToolbarButton97;
    BValide: TToolbarButton97;
    Bsupprime: TToolbarButton97;
    Binsert: TToolbarButton97;
    GroupIdentification: THGroupBox;
    PanelJob: THPanel;
    HLabel1: THLabel;
    SKJ_LIBELLE: THCritMaskEdit;
    SKJ_ACTIF: THCheckBox;
    HLabel2: THLabel;
    SKJ_LIBELLEAUTO: THCritMaskEdit;
    GroupeCaracteristiques: THGroupBox;
    HLabel3: THLabel;
    SKJ_EMAIL: THCritMaskEdit;
    HLabel4: THLabel;
    SKJ_CONFIDENTIEL: THSpinEdit;
    SKJ_FORCEEXEC: THCheckBox;
    SKJ_EXENAME: THCritMaskEdit;
    HLabel5: THLabel;
    HLabel6: THLabel;
    SKJACTION: THCritMaskEdit;
    HMTrad: THSystemMenu;
    Dock971: TDock97;
    SKJ_NOMJOB: THCritMaskEdit;
    SKJ_LIBREJOB: THCritMaskEdit;
    SKJ_TYPEJOB: THCritMaskEdit;
    BAgrandir: TToolbarButton97;
    BReduire: TToolbarButton97;
    BPurge: TToolbarButton97;
    Mess: THMsgBox;
    PanelFrame: THPanel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BsupprimeClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure FormCanResize(Sender: TObject; var NewWidth, NewHeight: Integer; var Resize: Boolean);
    procedure ChangeField(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BPurgeClick(Sender: TObject);
  protected //@@ private
    { Déclarations privées }
    JobId: integer; // N° de job en cours
    TobJob: tob; // Tob réelle du job et de ses détails
    ExeName: string; // Nom de l'exécutable
    Action: string; // nom de l'action
    // XP 10-10-2005 FQ 11930
    FLibelleDuJob: string;
    AvecAction: TActionFiche; // Indique la méthode en cours (création,consultation,modification)
    tobperiode: tob; // Tob virtuelle des données de la périodicité
    TobParams: tob; // Tob virtuelle de l'application (sera enregistré dans le détail)
  private   //@@
    Frame: TFramePeriodicite; // Instance de la frame de saisie de la période
    procedure AfficheJob;
    procedure ChargeJob;
    function ControleFiche: boolean;
    procedure Changement(sender: tobject);
  public
    { Déclarations publiques }
  protected//@@ private
    FTypeJob, FNomJob, FLibreJob: string;
    FEnabledDataSys: boolean;
  protected//@@private
    FVisibleInsert: BOOLEAN;
  end;

implementation

{$R *.DFM}

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. : Rend le n° du dernier job sur leque la fiche a travaillé
Mots clefs ... :
*****************************************************************}
function AglFicheJobExt(const ParamJobId: integer; const ParamAvecAction: TActionFiche; const ParamExeName, ParamAction: string; ParamTob: tob = nil; const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const VisibleInsert: boolean = False;
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True): integer;
begin
  with TPFicheJob.Create(application) do
  begin
    ExeName := ParamExeName;
    action := ParamAction;
    JobId := ParamJobId;
    AvecAction := ParamAvecAction;
    FTypeJob := TypeJob;
    FNomJob := NomJob;
    FLibreJob := LibreJob;

    // XP 10-10-2005 FQ 11930
    FLibelleDuJob := LibelleDuJob;
    FEnabledDataSys := EnabledDataSys;

    // XP 04-07-2005 : Par défaut, pas de création en boucle
    FVisibleInsert := VisibleInsert;

    TobParams := PAramTob;

    // Il faut retourner le dernier n° du job travaillé
    if ShowModal = mrOk then
      result := TobJob.GetInteger('SKJ_JOBID')
    else
      result := -1;
    Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003  Reseau
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function AglFicheJob(const ParamJobId: integer;
  const ParamAvecAction: TActionFiche;
  const ParamExeName, ParamAction: string;
  ParamTob: tob = nil;
  const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True): boolean;
begin
  with TPFicheJob.Create(application) do
  begin
    ExeName := ParamExeName;
    action := ParamAction;
    JobId := ParamJobId;
    AvecAction := ParamAvecAction;
    FTypeJob := TypeJob;
    FNomJob := NomJob;
    FLibreJob := LibreJob;

    // XP 10-10-2005 FQ 11930
    FLibelleDuJob := LibelleDuJob;
    FEnabledDataSys := EnabledDataSys;

    // XP 04-07-2005 : Par défaut, par de création en boucle
    FVisibleInsert := False;

    // XP 16-02-2005 : Il faut bien la stocker quelque part !!!
    TobParams := PAramTob;
    result := ShowModal = mrOk;
    Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.FormCreate(Sender: TObject);
begin
  JobId := -1;
  AvecAction := taConsult;
  tobjob := tob.create('STKJOBS', nil, -1);

  tobperiode := nil;

  // XP 04-07-2005 : Par défaut, par de création en boucle
  FVisibleInsert := False;

  Frame := TFramePeriodicite.CreateParented(PanelFrame.Handle);
  Frame.Parent := PanelFrame;
  Frame.OnChange := Changement;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.FormShow(Sender: TObject);
begin
  // XP 21-06-2005 : Nouveau champ
  if V_PGI.SAV then
  begin
    SKJ_TYPEJOB.Visible := True;
    SKJ_NOMJOB.Visible := True;
    SKJ_LIBREJOB.Visible := True;

    SKJ_TYPEJOB.Enabled := False;
    SKJ_NOMJOB.Visible := False;
    SKJ_LIBREJOB.Visible := False;
  end;

  // XP 04-07-2005 : Traitement identique pour BInsert et BSupprime !!!
  BInsert.Enabled := Self.FVisibleInsert;

  if AvecAction <> taCreat then
  begin
    // Lecture du job
    ChargeJob();

    // XP 05-07-2005 : Utile uniquement en modification, Affichage du job
    AfficheJob();

    // XP 07.02.2007
    BPurge.Enabled := True;
  end
  else
    BInsertClick(BInsert);

  // XP 21.04.2006 FQ 12503 DEBUT
  TradMini.AglTraduireFrame(Frame);
  // XP 21.04.2006 FQ 12503 FIN
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  //
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.FormDestroy(Sender: TObject);
begin
  FreeAndNil(tobjob);

  if assigned(tobperiode) then FreeAndNil(tobperiode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.FormCanResize(Sender: TObject; var NewWidth, NewHeight: Integer; var Resize: Boolean);
begin
  resize := false;
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = vk_f10) and (bvalide.enabled) then
  begin
    key := 0;
    BValideClick(bvalide);
  end
  else if (key = vk_escape) then
  begin
    key := 0;
    BFermeClick(bferme);
  end
  else if (key = vk_f7) and (bsupprime.enabled) then
  begin
    key := 0;
    BSupprimeClick(bsupprime);
  end
  else if (key = vk_f8) and (binsert.enabled) then
  begin
    key := 0;
    BInsertClick(binsert);
  end;
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.BValideClick(Sender: TObject);
var
  datetmp: tdatetime;
  MyParams: TOB;
begin
  // enregistrement
  if ControleFiche() then
  begin
    // récupération des données de l'écran
    tobjob.getecran(self);

    tobjob.putvalue('SKJ_LIBELLEAUTO', TobPeriode.GetValue('LIBAUTO'));
    datetmp := strtodatetime(TobPeriode.GetValue('DATEDEBUT'));
    tobjob.putvalue('SKJ_DATENEXTEXEC', datetmp);

    // XP 21-06-2005 : Mise à jour des champs auto
    if TobJob.FieldExists('SKJ_TYPEJOB') then
    begin
      TobJob.PutValue('SKJ_TYPEJOB', FTypeJob);
      TobJob.PutValue('SKJ_NOMJOB', FNomJob);
      TobJob.PutValue('SKJ_LIBREJOB', FLibreJob);
    end;

    // XP 22-10-2004 : Bizarre, mais la séquence n'était pas reprise !!!
    TobJob.putvalue('SKJ_SEQUENCE', JobSequenceToString(TobPeriode.getvalue('sequence')));

    // mise à jour de la tob période
    tobjob.PutValue('SKJ_DATA', AglConvertTobToXml(tobperiode));
  
    // enregistrement de le fiche
    if AvecAction = taCreat then
    begin
      // Si pas de TOB renseignée par l'application création
      if assigned(TobParams) then
        MyParams := TobParams
      else
        MyParams := tob.create('tob_params', nil, -1);

      // added by XP le 25-03-2003
      PrepareInput(MyParams);
      MyParams.addchampsupvaleur('JobMail', SKJ_EMAIL.Text);
      MyParams.addchampsupvaleur('Action', SKJACTION.Text);

      try
        if not JobSubmit(TobJob, SKJ_EXENAME.Text, MyParams) then
        begin
          HShowMessage('0;' + Caption + ';Une erreur est survenue lors de l''enregistrement de la tâche. Recommencer l''opération.;E;O;O;O', '', '');
          Exit;
        end
        else
        begin
          PGIInfo(TraduireMemoire('La tâche n° ') + IntToStr(tobjob.getvalue('SKJ_JOBID')) + TraduireMemoire(' a été programmée avec succès.'), '');
        end;
      finally
        if assigned(MyParams) and (MyParams <> TobParams) then
          FreeAndNil(MyParams);
      end;
    end
    else
    begin
      if assigned(TobParams) then
        MyParams := TobParams
      else
        MyParams := AglConvertXmlToTob(tobjob.detail[0].getvalue('SKD_DATA'));

      try
        // added by XP le 25-03-2003
        PrepareInput(MyParams);
        MyParams.addchampsupvaleur('JobMail', SKJ_EMAIL.Text);
        MyParams.addchampsupvaleur('Action', SKJACTION.Text);

        tobjob.detail[0].putvalue('SKD_DATA', AglConvertTobToXml(MyParams));
        BeginTrans();
        try
          tobjob.UpdateDB();
          CommitTrans();
        except
          on E: Exception do
          begin
            RollBack();
            HShowMessage('0;?Caption?;Erreur (%%) en modification de la tâche n°$$ ;E;O;O;O', E.Message, IntToStr(TobJob.GetInteger('SKJ_JOBID')));
            exit;
          end;
        end;
      finally
        if assigned(MyParams) and (MyParams <> TobParams) then
          FreeAndNil(MyParams);
      end;
    end;

    // Changement de l'action
    AvecAction := taModif;

    // affichage du job
    AfficheJob();

    // XP 04-07-2005 : FQ 11636 Si Multi création
    if not FVisibleInsert then
      ModalResult := mrOk;

    // XP 07.02.2007
    BPurge.Enabled := True;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.BFermeClick(Sender: TObject);
begin
  // En cours de création
  if AvecAction = taCreat then
  begin
    // Demande de confirmation
    if HShowMessage('0;' + Caption + ';Confirmez-vous l''annulation de la saisie en cours ?;Q;YN;N;N', '', '') = mrNo then Exit;
  end;

  Close;

  // XP 04-07-2005
  ModalResult := mrCancel;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... : 08-11-2004
Description .. : Suppression activée
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.BsupprimeClick(Sender: TObject);
begin
  if HShowMessage('0;' + Caption + ';Confirmez-vous la suppression de la tâche n°%% ?;Q;YN;N;N', InttoStr(JobId), '') = mrYes then
  begin
    if not uTaskJob.JobDel(JobId) then
    begin
      HShowMessage('0;' + Caption + ';La suppression n''a pas pu avoir lieu.Opération annulée.;E;O;O;O', '', '');
    end
    else
      Close();
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.BinsertClick(Sender: TObject);
begin
  // XP 07.02.2007
  BPurge.Enabled := False;

  // modification de l'action de la fiche
  AvecAction := taCreat;

  // Nettoyage de la tob
  tobjob.ClearDetail;

  // Vidage de la tob périodicite
  if assigned(tobperiode) then FreeAndNil(tobperiode);

  // Creation de la tob periodicite
  tobperiode := tob.create('tob_data', nil, -1);

  // Initialization de la tob
  JobInit(tobjob);

  // Affichage
  AfficheJob();
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.ChangeField(Sender: TObject);
begin
  // activation du bouton BValide
  if AvecAction <> taConsult then BValide.enabled := true;
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TPFicheJob.ControleFiche(): boolean;
begin
  Result := false;
  if TRIM(SKJ_LIBELLE.Text) = '' then
  begin
    HSHowMessage('0;' + Caption + ';Le libellé est obligatoire !!;E;O;O;O', '', '');
    if SKJ_LIBELLE.CanFocus() then
      SKJ_LIBELLE.SetFocus;
  end
  else if Frame.Ok(TobPeriode) then
    result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... : 21/03/2003
Description .. : Cette fonction charge dans une tob réelle le job et les lignes
Suite ........ : détails qui lui sont rattachées.
Suite ........ : Pour l'instant 1 seule ligne détail sera pris en compte pour la
Suite ........ : saisie !!
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.ChargeJob();
var
  q: tquery;
begin
  // vidage du job précédent
  tobjob.ClearDetail;

  // chargement du job et de son détail
  q := opensql('select * from stkjobs where skj_jobid=' + inttostr(jobid), true);
  if not q.eof then
  begin
    if tobjob.selectdb('', q) then
    begin
      ferme(q);
      // Alimentation de la tobperiode
      tobperiode := AglConvertXmlToTob(tobjob.getvalue('SKJ_DATA'));

      // recherche du détail
      q := opensql('select * from stkdetails where skd_jobid=' + inttostr(jobid), true);
      if not q.eof then tobjob.LoadDetailDb('STKDETAILS', '', '', q, false);
      ferme(q);

      // XP 16-02-2005 : C'est pas ma TOB !!
      // if assigned(tobparams) then FreeAndNil(TobParams);

      // XP 21-06-2005 : Alimentation des données libre
      if TobJob.FieldExists('SKJ_TYPEJOB') then
      begin
        FTypeJob := TobJob.GetString('SKJ_TYPEJOB');
        FNomJob := TobJob.GetString('SKJ_NOMJOB');
        FLibreJob := TobJob.GetString('SKJ_LIBREJOB');
      end;
    end
    else
      ferme(q);
  end
  else
    ferme(q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 21/03/2003
Modifié le ... : 21/03/2003
Description .. : Cette procédure gère l'affichage des champs (venant de 2
Suite ........ : tables) et tient compte du type d'accès à la fiche.
Suite ........ :
Suite ........ : Si le concept ccProgTask n'est pas actif, consultation !!
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.AfficheJob();
var
  MyParams: TOB;
begin
  BInsert.Enabled := AvecAction = taModif;
  BSupprime.Enabled := AvecAction = taModif;

  // si en cours de création
  if AvecAction = taCreat then
  begin
    GroupIdentification.enabled := true;

    // XP 18-10-2005 FQ 11969 alimentation des champs de l'écran
    TobJob.PutEcran(self);

    // XP 10-10-2005 FQ 11930 Déplacement car uniquement valable en création
    if FLibelleDuJob <> '' then
    begin
      SKJ_LIBELLE.Text := FLibelleDuJob;
      SKJ_LIBELLE.Enabled := False;
    end;

    // dans tous les cas déplacement car uniquement valable en création
    if action <> '' then
    begin
      SKJACTION.Text := Action;
      SKJACTION.Enabled := False;
    end;

    if ExeName <> '' then
    begin
      SKJ_EXENAME.Text := ExeName;
      SKJ_EXENAME.Enabled := False;
    end;

    // positionnement par défaut
    if SKJ_LIBELLE.CanFocus then
      SKJ_LIBELLE.SetFocus;

    // effacement du n° de job
    PanelJob.Caption := '';

    // Affichage de la frame
    Frame.Init(tobperiode);

    // adresse émail des préférences
    if v_pgi.SMTPFrom <> '' then SKJ_EMAIL.Text := v_pgi.SMTPFrom;

    // XP 18-10-2005 FQ 11969
    PanelJob.Visible := False;
  end
  else if AvecAction = taModif then
  begin
    GroupIdentification.enabled := true;

    // alimentation des champs de l'écran
    TobJob.PutEcran(self);

    // XP 05-07-2005
    try
      // Le nom de l'exécutable
      if tobjob.detail.Count > 0 then
      begin
        skj_exename.text := tobjob.detail[0].getvalue('SKD_EXENAME');
        // XP 05-07-2005
        MyParams := AglConvertXmlToTob(tobjob.detail[0].getvalue('SKD_DATA'));
        if assigned(MyParams) and MyParams.fieldexists('Action') then
          skjaction.text := MyParams.getvalue('Action')
        else
          skjaction.text := '';
      end
      else
        MyParams := nil;

      // positionnement par défaut
      if SKJ_LIBELLE.CanFocus() then
        SKJ_LIBELLE.SetFocus;

      // XP 18-10-2005 FQ 11969
      PanelJob.Visible := True;

      // affichage du n° de job
      PanelJob.Caption := 'N° ' + IntToStr(tobjob.getvalue('SKJ_JOBID'));

      // affichage de la frame
      Frame.Init(tobperiode);

      // par défaut
      BValide.enabled := false;
    finally
      if assigned(MyParams) then
        FreeAndNil(MyParams);
    end;
  end
  else
  begin
    GroupIdentification.enabled := false;
  end;

  // XP 10-10-2005 FQ 11930
  if (SKJ_EMAIL.Text <> '') and (SKJ_EXENAME.Text <> '') and (SKJACTION.Text <> '') then
  begin
    BAgrandir.Visible := True;
    BReduire.Visible := False;
    BReduire.Click();
  end
  else
  begin
    BAgrandir.Visible := False;
    BReduire.Visible := True;
    BAgrandir.Click();
  end;

  // XP 10-10-2005 FQ 11930
  if SKJ_EMAIL.Enabled then
    SKJ_EMAIL.Enabled := FEnabledDataSys;
  if SKJ_EXENAME.Enabled then
    SKJ_EXENAME.Enabled := FEnabledDataSys;
  if SKJACTION.Enabled then
    SKJACTION.Enabled := FEnabledDataSys;

  // XP 18-10-2005 FQ 11969
  if not FEnabledDataSys then
    BReduireClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 22/10/2004
Modifié le ... :   /  /
Description .. : Contrôle de cohénrece des dates
Mots clefs ... :
*****************************************************************}
procedure TPFicheJob.Changement(sender: tobject);
begin
  // activation du bouton BValide
  if AvecAction <> taConsult then BValide.enabled := true;
end;

procedure TPFicheJob.BAgrandirClick(Sender: TObject);
begin
  GroupeCaracteristiques.Visible := True;
  BAgrandir.Visible := False;
  BReduire.Visible := True;
end;

procedure TPFicheJob.BReduireClick(Sender: TObject);
begin
  GroupeCaracteristiques.Visible := False;
  BAgrandir.Visible := True;
  BReduire.Visible := False;
end;

procedure TPFicheJob.BPurgeClick(Sender: TObject);
begin
  if Mess.Execute(0, '', '') = mrYes then
  begin
    BeginTrans();
    try
      ExecuteSql('DELETE FROM STKLOGS WHERE SKL_JOBID=' + IntToStr(TobJob.GetInteger('SKJ_JOBID')));

      ExecuteSql('UPDATE STKJOBS SET SKJ_ENCOURS="-" WHERE SKJ_JOBID=' + IntToStr(TobJob.GetInteger('SKJ_JOBID')));

      TobJob.PutValue('SKJ_ENCOURS', '-');

      COmmitTrans();
    except
      ROllBack();
    end;
  end;
end;

end.

