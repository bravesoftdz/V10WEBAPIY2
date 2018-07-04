unit PGFicheJobMultiDossier;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Hent1, Hctrls, uTaskJob, hmsgbox, ExtCtrls, HPanel, HTB97, uTob, PGFicheJob,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  Grids;


function FicheJobMultiDossier(const ParamJobId: integer;
  const ParamAvecAction: TActionFiche;
  const ParamExeName, ParamAction: string;
  ParamTob: tob = nil;
  const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True; TempsExecParSalarie : Double = 0): boolean;

function FicheJobMultiDossierGetLastID(var JobID : Integer; const ParamJobId: integer;
  const ParamAvecAction: TActionFiche;
  const ParamExeName, ParamAction: string;
  ParamTob: tob = nil;
  const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True; TempsExecParSalarie : Double = 0; UniqueID : Boolean = True): boolean;

type
  TFicheJobMultiDossier = class(TPFicheJob)
  private
    { Déclarations privées }
    boLoading : Boolean;
    GBMultiDossier : THGroupBox;
    GridMultiDossier : THGrid;
    FTempsExecParSalarieEnSecondes: Double;
    TobDossiers : Tob;
    procedure SetTempsExecParSalarieEnSecondes(const Value: Double);
    Procedure GridMultiDossierCreate();
  public
    { Déclarations publiques }
    SaveFormCanResize : TCanResizeEvent;
    property TempsExecParSalarieEnSecondes : Double read FTempsExecParSalarieEnSecondes write SetTempsExecParSalarieEnSecondes;
    procedure MyFormCanResize(Sender: TObject; var NewWidth, NewHeight: Integer; var Resize: Boolean);
    procedure OnGridMultiDossierCellEnter(Sender: TObject);
    constructor Create(AOwner: TComponent); override;
    Destructor Destroy; override;
  end;

implementation

uses
  PgOutils2, DateUtils;


function FicheJobMultiDossierGetLastID(var JobID : Integer; const ParamJobId: integer;
  const ParamAvecAction: TActionFiche;
  const ParamExeName, ParamAction: string;
  ParamTob: tob = nil;
  const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True; TempsExecParSalarie : Double = 0; UniqueID : Boolean = True): boolean;
var
  FicheJobMultiDossier : TFicheJobMultiDossier;
  TobJob_skj, TobJob_skd: Tob; 
  ID : Integer;
  IndexRow, IndexDossier : Integer;
  ListeDossiers, ListeDossiersTpsExec : TStringList;
  stInsertSQL_skj, stXML_skj, stTempInsertSQL_skj : String;
  stInsertSQL_skd, stXML_skd, stTempInsertSQL_skd : String;
  stTempXML_skj, stTempXML_skd : String;
  TobRefJob_skj, TobRefJob_skd : Tob;
  NewID : integer;
  QFindID, QNbDos : TQuery;
  TobDataJob, TobDataJobDetail : Tob;
  SommeTpsExec : Integer;
  boMultiDossier : Boolean;
begin
  boMultiDossier := False;
  QNbDos:=OpenSql('SELECT COUNT(*) AS NB FROM DOSSIER',TRUE);
  if not QNbDos.eof then
      boMultiDossier := (QNbDos.Fields[0].asInteger > 1);
  Ferme(QNbDos);
  if boMultiDossier then { Gestion Multi-Dossier }
  begin
    FicheJobMultiDossier := TFicheJobMultiDossier.Create(Application);
    with FicheJobMultiDossier do
    begin
      { Simulation d'une partie de la fonction AglFicheJob }
      ExeName := ParamExeName;
      action := ParamAction;
      JobId := ParamJobId;
      AvecAction := ParamAvecAction;
      FTypeJob := TypeJob;
      FNomJob := NomJob;
      FLibreJob := LibreJob;
      FLibelleDuJob := LibelleDuJob;
      FEnabledDataSys := EnabledDataSys;
      FVisibleInsert := False;
      TobParams := PAramTob;
      TempsExecParSalarieEnSecondes := TempsExecParSalarie;
      result := ShowModal = mrOk;
      if not result then exit;
      { On récupère l'ID du job enregistré }
      ID := FicheJobMultiDossier.TobJob.GetInteger('SKJ_JOBID');
      { On récupère la liste des dossiers pour lesquels on veut activé le job }
      ListeDossiers := TStringList.Create;
      ListeDossiersTpsExec := TStringList.Create;
      for IndexRow := 1 to GridMultiDossier.RowCount -1 do
      begin
        if GridMultiDossier.CellValues[0, IndexRow] = 'X' then
        begin
          ListeDossiers.Add(GridMultiDossier.CellValues[2, IndexRow]);
          ListeDossiersTpsExec.Add(GridMultiDossier.CellValues[3, IndexRow]);
        end;
      end;
      Free;
    end;
    { On recherche le job enregistré et on le duplique sur l'ensemble des dossiers sélectionnés }
    TobRefJob_skj := Tob.Create('liste des jobs du dossier - Référence', nil, -1);
    TobRefJob_skj.LoadDetailDBFromSQL('stkjobs','select * from stkjobs where skj_jobid = "'+IntToStr(ID)+'"');

    TobRefJob_skd := Tob.Create('liste des jobs-details du dossier - Référence', nil, -1);
    TobRefJob_skd.LoadDetailDBFromSQL('STKDETAILS','SELECT * FROM STKDETAILS where SKD_JOBID = "'+IntToStr(ID)+'"');

    TobJob_skj := Tob.Create('liste des jobs du dossier', nil, -1);
    TobJob_skd := Tob.Create('liste des jobs-details du dossier', nil, -1);
    { On récupère les memos (ils ne sont pas gérés par le MakeInsertSql) }
    stXML_skj := TobRefJob_skj.detail[0].GetString('SKJ_DATA');
    stXML_skd := TobRefJob_skd.detail[0].GetString('SKD_DATA');
    { Sur le dossier en cours on supprime le job (on le rajoutera ensuite si nécéssaire) }
    JobDel(ID);
    NewID := 0;
    { Si on choisit d'avoir un ID unique, on parcours tous les dossiers pour trouver
      l'id utilisé le plus grand, puis on ajoute 1; }
    if UniqueID then
    begin
      for IndexDossier := 0 to ListeDossiers.Count -1 do
      begin
        QFindID:=OpenSql('SELECT MAX(SKJ_JOBID)+1 AS ID FROM '+ListeDossiers.Strings[IndexDossier]+'.dbo.stkjobs',TRUE);
        if (not QFindID.eof) and (QFindID.Fields[0].asInteger > NewID) then
          NewID := QFindID.Fields[0].asInteger;
        Ferme(QFindID);
      end;
    end;
    { On parcours les dossier et on ajoute le job (il faut trouvé le premier ID non utilisé par dossier) }
    SommeTpsExec := 0;
    for IndexDossier := 0 to ListeDossiers.Count -1 do
    begin
      stTempXML_skj := stXML_skj;
      stTempXML_skd := stXML_skd;
      TobJob_skj.Dupliquer(TobRefJob_skj, True, True);
      TobJob_skd.Dupliquer(TobRefJob_skd, True, True);
      BEGINTRANS;
      try
        if not UniqueID then
        begin
          NewID := 0;
          QFindID:=OpenSql('SELECT MAX(SKJ_JOBID)+1 AS ID FROM '+ListeDossiers.Strings[IndexDossier]+'.dbo.stkjobs',TRUE);
          if not QFindID.eof then
            NewID := QFindID.Fields[0].asInteger;
          Ferme(QFindID);
        end;
        if IndexDossier >= 1 then
        begin
          SommeTpsExec := SommeTpsExec + StrToInt(ListeDossiersTpsExec.Strings[IndexDossier - 1]);
          { Changement de la date de prochaine exécution dans le job }
          TobJob_skj.Detail[0].PutValue('SKJ_DATENEXTEXEC', IncMinute(TobJob_skj.Detail[0].GetDateTime('SKJ_DATENEXTEXEC'),SommeTpsExec));
          { Changement de la première date d'exécution dans les données du job }
          TobDataJob := AglConvertXmlToTob(stTempXML_skj);
          { On passe un string et pas un dateTime pour rester dans le même format date que l'original }
          TobDataJob.PutValue('DATEDEBUT', DateTimeToStr(IncMinute(TobDataJob.GetDateTime('DATEDEBUT'), SommeTpsExec)));
          stTempXML_skj := AglConvertTobToXml(TobDataJob);
          FreeAndNil(TobDataJob);
        end;
        { Changement du dossier dans les données du job-détail }
        TobDataJobDetail := AglConvertXmlToTob(stTempXML_skd);
        TobDataJobDetail.PutValue('DOSSIER', ListeDossiers.Strings[IndexDossier]);
        stTempXML_skd := AglConvertTobToXml(TobDataJobDetail);
        FreeAndNil(TobDataJobDetail);
        { On ajoute les caractère pour passer la traduction Décla ("" -> ") et la traduction SQL (\" -> ") }
        stTempXML_skj := StringReplace(stTempXML_skj, '"', '\""', [rfReplaceAll, rfIgnoreCase]);
        stTempXML_skd := StringReplace(stTempXML_skd, '"', '\""', [rfReplaceAll, rfIgnoreCase]);
        { On récupère la requête du job et on ajoute les données du mémo }
        TobJob_skj.Detail[0].PutValue('SKJ_JOBID',NewID);
        stTempInsertSQL_skj := TobJob_skj.Detail[0].MakeInsertSql;
        stTempInsertSQL_skj := StringReplace(stTempInsertSQL_skj,'INTO STKJOBS','INTO '+ListeDossiers.Strings[IndexDossier]+'.dbo.STKJOBS',[RfIgnoreCase]);
        stInsertSQL_skj := READTOKENPipe(stTempInsertSQL_skj, '(');
        stInsertSQL_skj := stInsertSQL_skj + '(SKJ_DATA, ';
        stInsertSQL_skj := stInsertSQL_skj + READTOKENPipe(stTempInsertSQL_skj, '(');
        stInsertSQL_skj := stInsertSQL_skj + '("'+stTempXML_skj+'", ';
        stInsertSQL_skj := stInsertSQL_skj + stTempInsertSQL_skj;
        ExecuteSQL(stInsertSQL_skj);
        { On récupère la requête du job-détail et on ajoute les données du mémo }
        TobJob_skd.Detail[0].PutValue('SKD_JOBID',NewID);
        stTempInsertSQL_skd := TobJob_skd.Detail[0].MakeInsertSql;
        stTempInsertSQL_skd := StringReplace(stTempInsertSQL_skd,'INTO STKDETAILS','INTO '+ListeDossiers.Strings[IndexDossier]+'.dbo.STKDETAILS',[RfIgnoreCase]);
        stInsertSQL_skd := READTOKENPipe(stTempInsertSQL_skd, '(');
        stInsertSQL_skd := stInsertSQL_skd + '(SKD_DATA, ';
        stInsertSQL_skd := stInsertSQL_skd + READTOKENPipe(stTempInsertSQL_skd, '(');
        stInsertSQL_skd := stInsertSQL_skd + '("'+stTempXML_skd+'", ';
        stInsertSQL_skd := stInsertSQL_skd + stTempInsertSQL_skd;
        ExecuteSQL(stInsertSQL_skd);
        JobID := NewID;
        COMMITTRANS;
      except
        ROLLBACK;
        PGIError('L''enregistrement de la tâche sur le dossier '+ListeDossiers.Strings[IndexDossier]+' a échoué.');
      end;
    end;
    FreeAndNil(TobJob_skj);
    FreeAndNil(TobJob_skd);
    FreeAndNil(TobRefJob_skj);
    FreeAndNil(TobRefJob_skd);
    FreeAndNil(ListeDossiersTpsExec);
    FreeAndNil(ListeDossiers);
  end else begin { Gestion non multi-dossier }
    result := AglFicheJob(ParamJobId, ParamAvecAction, ParamExeName, ParamAction, ParamTob,
                TypeJob, NomJob, LibreJob, LibelleDuJob, EnabledDataSys);
  end;
end;


Function FicheJobMultiDossier(const ParamJobId: integer;
  const ParamAvecAction: TActionFiche;
  const ParamExeName, ParamAction: string;
  ParamTob: tob = nil;
  const TypeJob: string = '';
  const NomJob: string = '';
  const LibreJob: string = '';
  const LibelleDuJob: string = '';
  const EnabledDataSys: boolean = True; TempsExecParSalarie : Double = 0): boolean;
var
  JobID : Integer;
begin
  result := FicheJobMultiDossierGetLastID(JobID, ParamJobId, ParamAvecAction,
              ParamExeName, ParamAction, ParamTob, TypeJob, NomJob, LibreJob,
              LibelleDuJob, EnabledDataSys, TempsExecParSalarie, False);
end;

{ TFicheJobMultiDossier }

constructor TFicheJobMultiDossier.Create(AOwner: TComponent);
var
  TobFindDossier : Tob;
  IndexDossier : Integer;
begin
  inherited;
  TobDossiers := Tob.Create('Liste des dossiers', nil, -1);
  TobDossiers.LoadDetailFromSQL('SELECT "-" as DOS_ACTIVE, DOS_NODOSSIER, DOS_LIBELLE, DOS_NOMBASE, "" as TPS from DOSSIER');
  TobFindDossier := TobDossiers.FindFirst(['DOS_NODOSSIER'],[V_PGI.NoDossier], False);
  { Gestion spécifique au multi-dossier }
  boLoading := True;
  OnCanResize := MyFormCanResize;
  Height := Height + 250;
  GBMultiDossier := THGroupBox.Create(self);
  GBMultiDossier.Parent := PanelHaut;
  GBMultiDossier.Name := 'GBMULTIDOSSIER';
  GBMultiDossier.Align := alTop;
  GBMultiDossier.Height := 250;
  GBMultiDossier.Caption := 'Liste des dossiers sur lesquels exécuter la tâche';
  for IndexDossier := TobDossiers.Detail.Count -1 downto 0 do
  begin
    { Si on est sur le dossier courant, on le sélectionne par défaut. }
    if TobFindDossier.GetString('DOS_NOMBASE') = TobDossiers.Detail[indexDossier].GetString('DOS_NOMBASE') then
      TobDossiers.Detail[indexDossier].PutValue('DOS_ACTIVE','X');
  end;
  { Remplissage de la grille avec les noms des dossiers }
  GridMultiDossierCreate;
  GridMultiDossier.Col := 1;
  TobDossiers.PutGridDetail(GridMultiDossier, False, False, 'DOS_ACTIVE;DOS_LIBELLE;DOS_NOMBASE;TPS');
  boLoading := False;
  { On met les groupBox dans le bon ordre (elles sont toutes alignées à alTop) }
  GroupIdentification.Visible := False;
  GroupeCaracteristiques.Visible := False;
  GroupeCaracteristiques.Visible := True;
  GroupIdentification.Visible := True;
end;

destructor TFicheJobMultiDossier.Destroy;
begin
  FreeAndNil(TobDossiers);
  inherited;
end;

procedure TFicheJobMultiDossier.GridMultiDossierCreate;
begin
  GridMultiDossier := THGrid.Create(self);
  GridMultiDossier.Options := [goColSizing, goColMoving, goEditing, goTabs, goHorzLine];
  GridMultiDossier.Name := 'GRIDMULTIDOSSIER';
  GridMultiDossier.Parent := GBMultiDossier;
  GridMultiDossier.Align := alClient;
  GridMultiDossier.DefaultRowHeight := 16;
  GridMultiDossier.ColCount := 4;
  GridMultiDossier.FixedCols := 0;
  GridMultiDossier.RowCount := 2;
  GridMultiDossier.FixedRows := 1;
  GridMultiDossier.CellValues[0,0] := 'Activé';
  GridMultiDossier.CellValues[1,0] := 'Nom du dossier';
  GridMultiDossier.CellValues[2,0] := 'Nom de la base';
  GridMultiDossier.CellValues[3,0] := 'Temps estimé (en minutes)';
  GridMultiDossier.ColNames[0] := 'Activé';
  GridMultiDossier.ColEditables[0] := False;
  GridMultiDossier.ColAligns[0] := taCenter;
  GridMultiDossier.ColWidths[0] := 60;
  GridMultiDossier.ColTypes[0] := 'B';
  GridMultiDossier.ColFormats[0] := IntToStr(Ord(csCheckBox));
  GridMultiDossier.ColNames[1] := 'Nom du dossier';
  GridMultiDossier.ColEditables[1] := False;
  GridMultiDossier.ColTypes[1] := 'C';
  GridMultiDossier.ColAligns[1] := taLeftJustify;
  GridMultiDossier.ColWidths[1] := 200;
  GridMultiDossier.ColNames[2] := 'Nom de la base';
  GridMultiDossier.ColEditables[2] := False;
  GridMultiDossier.ColTypes[2] := 'C';
  GridMultiDossier.ColAligns[2] := taLeftJustify;
  GridMultiDossier.ColWidths[2] := 200;
  GridMultiDossier.ColNames[3] := 'Temps d''exéction estimé (en minutes)';
  GridMultiDossier.ColEditables[3] := True;
  GridMultiDossier.ColTypes[3] := 'C';
  GridMultiDossier.ColAligns[3] := taLeftJustify;
  GridMultiDossier.ColWidths[3] := 110;
  { Affectation du gestionnaire de cliques sur la grille }
  GridMultiDossier.OnClick := OnGridMultiDossierCellEnter;
end;

procedure TFicheJobMultiDossier.MyFormCanResize(Sender: TObject;
  var NewWidth, NewHeight: Integer; var Resize: Boolean);
begin
  resize := boLoading;
end;

procedure TFicheJobMultiDossier.OnGridMultiDossierCellEnter(
  Sender: TObject);
begin
  { Clique sur une cellule de la 1ere colonne (à part sur la ligne de titre de colonnes) }
  if (GridMultiDossier.Col = 0) and (GridMultiDossier.Row <> 0) then
  begin
    if GridMultiDossier.CellValues[0,GridMultiDossier.Row] = '-' then
      GridMultiDossier.CellValues[0,GridMultiDossier.Row] := 'X'
    else
      GridMultiDossier.CellValues[0,GridMultiDossier.Row] := '-';
  end;
end;

procedure TFicheJobMultiDossier.SetTempsExecParSalarieEnSecondes(
  const Value: Double);
var
  DebExer, FinExer : TDateTime;
  DateDeb, DateFin: TDateTime;
  MoisE, AnneeE, ComboExer : String;
  IndexDossier : Integer;
  TempTob, TobEstimation : Tob;
begin
  FTempsExecParSalarieEnSecondes := Value;
  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
  DateDeb := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01);
  DateFin := FINDEMOIS(EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01));
  TobEstimation := Tob.Create('Estimation du temps d''execution', nil, -1);
  for IndexDossier := TobDossiers.Detail.Count -1 downto 0 do
  begin
    { Recherche du nombre salariés présents dans le mois en cours (dans le dossier) pour estimer le temps d'execution }
    try
      TobEstimation.LoadDetailFromSQL('Select Count(*) as NBSAL from '
        +TobDossiers.Detail[indexDossier].GetString('DOS_NOMBASE')+'.dbo.SALARIES WHERE '
        +'((PSA_DATESORTIE >="' + UsDateTime(DateDeb)+'") OR (PSA_DATESORTIE IS NULL) '
        +'OR (PSA_DATESORTIE <= "' + UsDateTime(iDate1900) + '")) AND (PSA_DATEENTREE <="' + UsDateTime(DateFin) + '")');
      TobDossiers.Detail[indexDossier].PutValue('TPS', Trunc(TobEstimation.Detail[0].GetInteger('NBSAL') * TempsExecParSalarieEnSecondes / 60 + 5));
    except
      TempTob := TobDossiers.Detail[indexDossier];
      TempTob.ChangeParent(nil, -1);
      FreeAndNil(TempTob);
    end;
  end;
  FreeAndNil(TobEstimation);
  { Mise à jour de la grille }
  GridMultiDossier.Free;
  GridMultiDossierCreate;
  TobDossiers.PutGridDetail(GridMultiDossier, False, False, 'DOS_ACTIVE;DOS_LIBELLE;DOS_NOMBASE;TPS');
  GridMultiDossier.Col := 1;
end;

end.
