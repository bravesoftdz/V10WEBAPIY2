unit FicheLanceImport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Mask, UImportDP, Strutils, Sql7util, MenuOLG, HEnt1,
  {$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ELSE}
  {$ENDIF}
  utiltrans, HTB97, Hctrls, ShellAPI, Registry;

type
  TFLanceImport = class(TForm)
    opendlg: TOpenDialog;
    savedlg: TSaveDialog;
    GrpGeneral: TGroupBox;
    LCheminTra: TLabel;
    LCheminBackup: TLabel;
    ChkSauvegarderAvant: TCheckBox;
    BExporter: TButton;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    ChkSupprAvantImport: TCheckBox;
    FCheminTra: THCritMaskEdit;
    FCheminBackup: THCritMaskEdit;
    ChkGenererLog: TCheckBox;
    ChkEcraserEnreg: TCheckBox;
    ChkStopSurErreur: TCheckBox;
    procedure ChkSauvegarderAvantClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BExporterClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BValiderClick(Sender: TObject);
    procedure FCheminTraElipsisClick(Sender: TObject);
    procedure FCheminBackupElipsisClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure ChkEcraserEnregClick(Sender: TObject);
    procedure ChkStopSurErreurClick(Sender: TObject);
  private
    { Déclarations privées }
    IsAutomate : Boolean;
    procedure deleteTables;
    function Sauvegarde (var import : TImportDP) : integer;
  public
    { Déclarations publiques }
  end;

var
  FLanceImport: TFLanceImport;

implementation

{$R *.DFM}

uses hmsgbox;

//Option de sauvegarde préalable
procedure TFLanceImport.ChkSauvegarderAvantClick(Sender: TObject);
begin
  if (ChkSauvegarderAvant.State = cbChecked) or IsAutomate then
    begin
    LCheminBackup.enabled := true;
    FCheminBackup.enabled := true;
//    button4.enabled := true;
    end
  else
    begin
    if PGIAsk('Etes-vous sûr de ne pas vouloir sauvegarder les tables actuelles ?','Sauvegarde')=mrYes then
      begin
      LCheminBackup.enabled := false;
      FCheminBackup.enabled := false;
//      button4.enabled := false;
      end
    else
      ChkSauvegarderAvant.State := cbChecked;
    end;
end;

procedure TFLanceImport.deleteTables;
var
  table, s_sql : string;
  rsql : TQuery;
begin
//delete toutes les tables concernées avant
  s_sql := '"CAE","T","YTC","ANN","ANB","ANL","C","DOS","DAP","DOG","DCI","DCL","DCV","DEC","DFI","DPM","DOR","DPP","DSO","DSC","DTC","DT1","DTP","ARS","GA","MR","R","MP","ANP","RPR","AFO","GRP","LDO","JUR","ROP","RD2","RPJ","RDQ","RPE","RDV","RPT","RAC","RD1","RAI"'
          +',"RD6"';
  //s_sql := '"CAE","T","YTC","ANN","ANB","ANL","US","C","DOS","DAP","DOG","DCI","DCL","DCV","DEC","DFI","DPM","DOR","DPP","DSO","DSC","DTC","DT1","DTP","ARS","GA","MR","R","MP","ANP","RPR","AFO","JUR","ROP","RD2","RPJ","RDQ","RPE","RDV","RAC","RD1"';
  //rsql := OpenSql('SELECT DT_NOMTABLE FROM DETABLES WHERE DT_PREFIXE IN ("CAE","T","YTC","ANN","ANB","ANL","US","C","DOS","DAP","DOG","DCI","DCL","DCV","DEC","DFI","DPM","DOR","DPP","DSO","DSC","DTC","DT1","DTP","ARS","GA","MR","R","MP","ANP","RPR","AFO","JUR","ROP","RD2","RPJ","RDQ","RPE","RDV","RAC","RD1")', true);
  rsql := OpenSql('SELECT DT_NOMTABLE FROM DETABLES WHERE DT_PREFIXE IN ('+ s_sql + ')', true);
  while not rsql.eof do
    begin
    table := rsql.FindField('DT_NOMTABLE').AsString;
    executeSql('DELETE FROM '+table);
    rsql.Next;
    end;
  Rsql.Free;
//et les tablettes, mais uniquement pour les types concernés par l'import
  executeSql('DELETE FROM CHOIXCOD WHERE CC_TYPE IN ("DRA","DMS","DOC","DCR","DWI","DMT","DNT","DRI","DRT","DTB","DTM","UCO","AN1","AN2","AN3","YEW","VOT","LGU","GZC","FN1","FN2","FN3","TX1","TX2","TX3","TX4"'
      +',"TX5","GCT","GCA","JUR","GCO","GOR","TRC","SCC","RTV","TAR","FON","NVR","INC","MEX","FVS","SRV","LPA","PRO","GDM","CIV","LIP","TRE","TAS","AP1","AP2","AP3","ONB"'
      +',"ZLA","ZLC","ZLT","RLZ","ROO","ROF","RTP","RMP","RMS")');
  executeSql('DELETE FROM CHOIXEXT WHERE YX_TYPE IN ("DCG","DCI","DLC","DC2","ANP","DAU","LB1","LB2","LB3","LB4","LB5","LB6","LB7","LB8","LB9","LBA","LA1","LA2","LA3","LA4","LA5","LA6","LA7","LA8","LA9"'
      +',"LAA","LC1","LC2","LC3","LT1","LT2","LT3","LT4","LT5","LT6","LT7","LT8","LT9","LTA","LR1","LR2","LR3","LR4","LR5","LR6","LR7","LR8","LR9","LRA","EL1","EL2","EL3"'
      +',"EL4","EL5","EL6","EL7","EL8","EL9","ELA","LF1","LF2","LF3","LF4","LF5","LF6","LF7","LF8","LF9","LFA","OR1","OR2","OR3","OR4","OR5","RR1","RR2","RR3","RR4","RR5")');
end;

procedure TFLanceImport.FormCreate(Sender: TObject);
var
  param, valeur : string;
  i : integer;
  KeySuppTables, KeyEcraserEnreg : boolean;
begin
  FCheminBackup.text := 'C:\Backup.tra';
  GrpGeneral.Caption := 'Format d''échange : v'+version;
  IsAutomate := False;
  KeySuppTables := False;
  KeyEcraserEnreg := False;

  // Modif Boulet Proof :
  // si clé spécifique présente dans le registre, on désactive les options "dangereuses" SuppTables et EcraserEnreg
  // pour info GetFromRegistry ne fonctionne pas

  with TRegistry.Create do
  try
    RootKey := HKEY_LOCAL_MACHINE;
    if OpenKey('Software\' + Apalatys + '\' + NomHalley, False) then
    begin
      if ValueExists('DisableSuppTables') then KeySuppTables   :=ReadBool ('DisableSuppTables');
      if ValueExists('DisableEcraserEnreg') then KeyEcraserEnreg :=ReadBool ('DisableEcraserEnreg');
    end;
  finally
    Free;
  end;

  if KeySuppTables then ChkSupprAvantImport.Visible := False;
  if KeyEcraserEnreg then ChkEcraserEnreg.Visible := False;

  // Affectation des champs si programme lancé en ligne de commande
  If paramcount <> 0 then
    begin
    // Valeurs par défaut
    IsAutomate := True;
    ChkSupprAvantImport.Checked := False;
    ChkGenererLog.Checked := True;
    ChkSauvegarderAvant.Checked := False;
    ChkEcraserEnreg.Checked := False;
    ChkStopSurErreur.Checked := False;
    // Récup des paramètres précisés
    For i := 1 to paramcount do
      begin
      param := copy (ParamStr(i), 1, Pos('=',ParamStr(i)));
      valeur := RightStr(ParamStr(i), (Length(ParamStr(i))-Length(param)));
      if param = '/FICHIERTRADP=' then FCheminTra.Text := valeur;
      if param = '/SUPPTABLES='   then ChkSupprAvantImport.Checked := (UpperCase(valeur) = 'TRUE'); //StrToBool(valeur);
      if param = '/GENERERLOG='   then ChkGenererLog.Checked       := (UpperCase(valeur) = 'TRUE'); //StrToBool(valeur);
      if param = '/SAVETABLES='   then ChkSauvegarderAvant.Checked := (UpperCase(valeur) = 'TRUE'); //StrToBool(valeur);
      if param = '/ECRASERENREG=' then ChkEcraserEnreg.Checked     := (UpperCase(valeur) = 'TRUE'); //StrToBool(valeur);
      if param = '/STOPERREUR='   then ChkStopSurErreur.Checked    := (UpperCase(valeur) = 'TRUE'); //StrToBool(valeur);
      end;
    end;
    modalresult := mrOK;
end;

//Export Simple
procedure TFLanceImport.BExporterClick(Sender: TObject);
var
  res : integer;
  import : TImportDP;
begin
  import := TImportDP.Create;
  res := Sauvegarde(import);
  if res = 0 then
    begin
    PgiInfo('Export Réussi !');
    modalresult := mrOk;
    end;
  FreeAndNil(import);
  exit;
end;

//Génération d'un TRA
function TFLanceImport.Sauvegarde (var import : TImportDP) : integer;
var res : integer;
begin
  res := 0;
  if (upperCase(ExtractFileExt(FCheminBackup.text)) = '.TRA') then
    // *** Rajouter ici les nouveaux types décrits dans la fonction __initExtract de UImportDP
    res := import.generateFile(['CAE','T','YTC','ANN','ANB','ANL','US','C','DOS','DAP','DOG','DCI','DCL','DCV','DEC','DFI','DPM','DOR','DPP','DSO','DSC','DTC','DT1','DTP','CC','YX','ARS','GA','PY','ET','MR','R','MDP','ANP','RPR','AFO','JUR','ROP','RD2','RPJ','RDQ','RPE','RDV','RPT','RAC','RD1','RAI','RD6'], FCheminBackup.text)
  else
    begin
    PgiInfo('Fichier de destination incorrect.');
    FCheminBackup.Text := 'C:\Backup.tra';
    modalResult := mrAbort;
    result := -1;
    exit;
    end;
  result := res;
  if res <> 0 then
    begin
    PgiInfo('Erreur lors de la sauvegarde, Export impossible.');
    modalResult := mrAbort;
    exit;
    end;
end;

procedure TFLanceImport.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  //Import.Free;
  // On force la fermeture de l'application si programme lancé en ligne de commande
  if IsAutomate then
    begin
    modalresult := MrOK;
    FMenuG.FermeSoc;
    FMenuG.Quitter;
    end;
end;

//Import
procedure TFLanceImport.BValiderClick(Sender: TObject);
var
  res : integer;
  import : TImportDP;
  SInformation : String;
  SErreur : String;
begin
  res := 0;
  Import := TImportDP.Create;

  //***** CONTROLES *****
  //VERIF QUE LE FICHIER EXISTE.
  if Not FileExists(FCheminTra.text) then
    begin
    SErreur := 'Le fichier est introuvable !';
    Import.GenererFichierLog(SErreur);
    if Not IsAutomate then pgiInfo(SErreur);
    FCheminTra.text := '*.tra';
    modalResult := mrAbort;
    exit;
    end;
  //VERIF QUE LE FICHIER EST UN TRA
  if (Not (upperCase(ExtractFileExt(FCheminTra.text)) = '.TRA')) or (copy(FCheminTra.text,1,1)='*') then
    begin
    SErreur := 'Le fichier n''est pas au format TRA !';
    Import.GenererFichierLog(SErreur);
    if Not IsAutomate then pgiInfo(SErreur);
    FCheminTra.text := '*.tra';
    modalResult := mrAbort;
    exit;
    end;
  //Verif de la version
  If getVersion(FCheminTra.text) <> version then
    begin
    SErreur := 'Le fichier n''est pas de la bonne version du format d''échange.';
    Import.GenererFichierLog(SErreur);
    if Not IsAutomate then pgiInfo(SErreur);
    FCheminTra.text := '*.tra';
    modalResult := mrAbort;
    exit;
    end;
  //Verif fichier d'import différent de fichier d'export
  If (upperCase(FCheminTra.text) = upperCase(FCheminBackup.text)) and ChkSauvegarderAvant.Checked then
    begin
    SErreur := 'Vous ne pouvez sauvegarder dans le fichier qui sera importé !';
    Import.GenererFichierLog(SErreur);
    if Not IsAutomate then PgiInfo(SErreur);
    FCheminTra.text := '*.tra';
    FCheminBackup.text := 'C:\Backup.tra';
    modalResult := mrAbort;
    exit;
    end;
  //Demande confirmation
  If Not IsAutomate and Not (Pgiask('Vous allez remplacer toutes les données du Dossier Permanent / Bureau Expert.'+#13#10+' Confirmez-vous ?') = mrYes) then
    begin
    modalResult := mrAbort;
    exit;
    end;

  //sauvegarde
  if (ChkSauvegarderAvant.State = cbChecked) then
    res := sauvegarde(import);//Procédure de sauvegarde : génération du fichier
  if res <> 0 then//si échec de la sauvegarde, on ne continue pas
    begin
    FreeAndNil(import);
    exit;
    end;

  //suppression des tables si l'option est activée
  if ChkSupprAvantImport.State = cbChecked then
    deleteTables;

  Import.ChkFichierLog:=ChkGenererLog.Checked;

  //lancement import, aucun message d'information si programme exécuté en ligne de commande
  res := import.LanceImport(FCheminTra.text, IsAutomate, (ChkEcraserEnreg.State = cbChecked), (ChkStopSurErreur.State = cbChecked));
  if Not IsAutomate then
    begin
    if res = 0 then
      begin
      SInformation:='Le fichier '+ExtractFileName (FCheminTra.text)+' a été importé avec succés.';
      if (Import.FichierLogGen) then
        begin
        SInformation:=SInformation+#13+#10
        +'Cependant, des enregistrements comportent des erreurs et n''ont pu être insérés dans la base.'+#13+#10
        +'Voulez-vous visualiser le fichier LOG qui a été généré ?';
        if (PgiAsk (SInformation,'Import DP')=mrYes) then ShellExecute (0,Pchar ('open'),pchar ('NotePad.exe'),pchar (import.CheminFichierLog),nil,SW_RESTORE);
        end
      else
        PgiInfo(SInformation,'Import DP');
      end
    else
      PgiInfo('Echec de l''import, vérifiez si le fichier est correct.');
    end;

  modalresult := mrOK;
  FreeAndNil(import);

end;

procedure TFLanceImport.FCheminTraElipsisClick(Sender: TObject);
begin
  opendlg.Execute;
  FCheminTra.text := opendlg.Filename;
end;

procedure TFLanceImport.FCheminBackupElipsisClick(Sender: TObject);
begin
  savedlg.Execute;
  FCheminBackup.text := savedlg.Filename;
end;

procedure TFLanceImport.FormActivate(Sender: TObject);
var action : TCloseAction;
begin
   // Lancement de l'import automatique si programme exécuté en ligne de commande
   If IsAutomate then
    begin
    BValiderClick(Self);
    action := caFree;
    FormClose(Self, action);
    end;
end;

procedure TFLanceImport.ChkEcraserEnregClick(Sender: TObject);
begin
  if Not IsAutomate then
  begin
    if (ChkEcraserEnreg.State = cbChecked) then
      PGIInfo('En activant cette option, vous allez écraser les enregistrements existants dans la base par ceux importés depuis le fichier .TRA !')
      //ChkEcraserEnreg.Hint := 'En désactivant cette option, vous allez conserver les enregistrements existants dans la base et importer seulement les champs remplis dans le fichier .TRA !'
    else
      PGIInfo('En désactivant cette option, vous allez conserver les enregistrements existants dans la base et importer seulement les champs remplis dans le fichier .TRA !');
      //ChkEcraserEnreg.Hint := 'En activant cette option, vous allez écraser les enregistrements existants dans la base par ceux importés depuis le fichier .TRA !';
  end;
end;

procedure TFLanceImport.ChkStopSurErreurClick(Sender: TObject);
begin
  if Not IsAutomate then
  begin
    if (ChkStopSurErreur.State = cbChecked) then
      PGIInfo('En activant cette option, si une erreur est détectée, l''import sera arrêté et aucune donnée ne sera intégrée !')
      //ChkStopSurErreur.Hint := 'En désactivant cette option, si une erreur est détectée, l''enregistrement ne sera pas intégré mais l''import continuera avec les enregistrements suivants !'
    else
      PGIInfo('En désactivant cette option, si une erreur est détectée, l''enregistrement ne sera pas intégré mais l''import continuera avec les enregistrements suivants !');
      //ChkStopSurErreur.Hint := 'En activant cette option, si une erreur est détectée, l''import sera arrêté et aucune donnée ne sera intégrée !';
  end;
end;

end.
