{***********UNITE*************************************************
Auteur  ...... : JSI
Créé le ...... : 06/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BACKUPSOCIETE ()
           ... : Sauvegarde multi-sociétés
Mots clefs ... : TOF;BACKUPSOCIETE
*****************************************************************}
unit BackupRestore_TOF;

interface

uses StdCtrls,
  Controls,
  Classes, Windows,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  mul,FE_Main,
  MajTable,
{$ELSE}
  eMul,
  Maineagl, UtileAGL,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls, HTB97,
  HEnt1, HSysMenu,
  HMsgBox,
  UTOF, UTOB, UTOZ;

Type T_CopieInfo = (ciOk, ciError, ciExist, ciZip);

Type T_ParamBackUpRestore = record
          BackupNotRestore : boolean;
          Directory : string;
          DirectoryTemp : string;
          Zip : boolean;
          ForceLaRecopie : boolean;
          DeleteAfterRestore : boolean;
          User : string;
          Password : string;
        end;

type
  TOF_BACKUPRESTORE = class (TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument (S: string) ; override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;

  private
    Stop : Boolean;
    bBackupNotRestore : boolean;
    Recap,ListSoc : TListBox;
    GSoc : THGrid;
    HMTrad: THSystemMenu;
    sServerName : string;

    { controles }
    procedure InitialiseLesCtrl;
    procedure bStopClick(Sender: TObject);
    procedure LookupDirectoryDrive(Sender: TObject);

    { registry }
    procedure GetParamFromRegistry(PrefixeCle : string);
    procedure SetParamToRegistry(PrefixeCle : string);

    { société }
    procedure ChargeLesSocietes;
    function  TraiteLaSociete (Societe : string ; Params : T_ParamBackUpRestore) : Boolean;

    function  GetBackupFile(Societe : string ; Params : T_ParamBackUpRestore ; WithoutZip : boolean = false) : string;
    function  GetInfoFromCegidPGIIni(Societe,QuelleInfo : string) : string;

    { sauvegarde & restauration }
    function  BackupDB(Societe : string  ; Params : T_ParamBackUpRestore)  : T_CopieInfo;
    function  RestoreDB(Societe : string ; Params : T_ParamBackUpRestore)  : T_CopieInfo;
    function  ExecSQLAdm(Societe, BackupFile : string ; Params : T_ParamBackUpRestore): boolean;

    { message }
    procedure AfficheMessage(Msg : string ; Avertissement : boolean = false);
  end;

function YYLanceFiche_BackupSociete (Nat, Cod: string; Range, Lequel, Argument: string) : string;

implementation

uses filectrl, galSystem, Sql7Util, inifiles, Registry;

const icoAvert   = '#ICO#57';
      icoZip     = '#ICO#98';
      icoSave    = '#ICO#52';
      icoRestore = '#ICO#53';
      icoMain    = '#ICO#49';


function YYLanceFiche_BackupSociete (Nat, Cod: string; Range, Lequel, Argument: string) : string;
begin
  AGLLanceFiche (Nat, Cod, Range, Lequel, Argument) ;
end;

procedure TOF_BACKUPRESTORE.OnArgument (S: string) ;

  procedure InitFormWithArgument;
  begin
    if bBackupNotRestore then
    begin
      Ecran.Caption := TraduireMemoire('Sauvegarde multi-sociétés');
      SetControlVisible('CHKDELETE',false);
    end else
    begin
      Ecran.Caption := TraduireMemoire('Restauration multi-sociétés');
      SetControlVisible('CHKFORCESAUVE',false);
    end;
    UpdateCaption(Ecran);
  end;
  function GetPrefixe : string;
  begin
    if bBackupNotRestore then
           result := 'BackupSociete'
      else result := 'RestoreSociete';
  end;

  function GetServerName : string;
  var Qs : TQuery;
      iPos : integer;
  begin
    try
      Qs := OpenSQL('SELECT CAST (@@SERVERNAME AS VARCHAR) AS MONSERVEUR',true);
      if not Qs.Eof then result := Qs.FindField('MONSERVEUR').AsString //ExtractFileDrive(Qs.FindField('MONSERVEUR').AsString)
      else result := '';
    finally
      Ferme(Qs);
    end;
    iPos := pos('\',result);
    if iPos > 0 then result := copy(result,1,iPos-1);
end;

  function GetLocalMachineName : String;
  var Buffer: array[0..1023] of Char;
      lp: DWORD;
  begin
    Result:='';
    lp := SizeOf(Buffer);
    GetComputerName(Buffer,lp);
    SetString(Result,Buffer,lp);
    Result:=UpperCase(Result);
  end;

begin
  inherited;
  bBackupNotRestore := (pos('BACKUP',S) > 0);

  { affectation des controles }
  InitialiseLesCtrl;
  InitFormWithArgument;

  { chargement des sociétés }
  ChargeLesSocietes;

  { lecture de la registry }
  GetParamFromRegistry(GetPrefixe);

  { recherche du nom de la machine hébergeant le serveur sql de la base de connection }
  sServerName := GetServerName;
  if sServerName = GetLocalMachineName then
     sServerName := '';
end;

procedure TOF_BACKUPRESTORE.OnNew;
begin
  inherited;
end;

procedure TOF_BACKUPRESTORE.OnDelete;
begin
  inherited;
end;

procedure TOF_BACKUPRESTORE.OnUpdate;
var Societe,PathSQLAdm : String;
    iInd : Integer;
    DatDebut, DatFin: TDateTime;
    Params : T_ParamBackUpRestore;

  procedure SetParams;
  begin
    FillChar(Params,Sizeof(Params),#0);
    Params.BackupNotRestore := bBackupNotRestore;
    Params.Zip := (GetCheckBoxState  ('CHKCOMPRESSION') = cbChecked);
    Params.ForceLaRecopie := (GetCheckBoxState  ('CHKFORCESAUVE') = cbChecked);
    Params.DeleteAfterRestore := (GetCheckBoxState  ('CHKDELETE') = cbChecked);
    Params.Directory := GetControlText ('DIRECTORYBACKUP') ;
    Params.DirectoryTemp := GetControlText ('DIRECTORYBACKUP') + '\Temp001';
    Params.User := GetControlText ('USER') ;
    Params.Password := GetControlText ('PASSWORD') ;
  end;

  function GereStop : boolean;
  begin
    Application.ProcessMessages;
    if Stop then AfficheMessage('Traitement interrompu par l''utilisateur');
    result := Stop;
  end;

  function GetDirectoryDrive(DirectoryPath : string) : string;
  var st : string;
      iPos : integer;
  begin
    result := '';
    if copy(DirectoryPath,0,2) = '\\' then  { @ réseau }
    begin
      st := StringReplace(DirectoryPath,'\\','',[rfReplaceAll]);
      iPos := pos('\',st);
      if iPos > 0 then
      begin
        result:= UpperCase(copy(st,0,iPos-1));
      end;
    end;
  end;

begin
  inherited;
  Stop := False;
  if GSoc.NbSelected = 0 then
  begin
    PGIBox('Aucune société sélectionnée');
    exit;
  end;

  { Test de validité du répertoire de sauvegarde }
  if not DirectoryExists(GetControlText ('DIRECTORYBACKUP')) then
    begin
    PGIBox('Le chemin de sortie est incorrect');
    exit;
  end;

  if GetDirectoryDrive(ExpandUNCFileName(GetControlText ('DIRECTORYBACKUP'))) <> UpperCase(sServerName) then
  begin
    PGIBox('La sauvegarde ne peut se faire que sur un disque local');
    exit;
  end;

  PathSQLAdm := IncludeTrailingBackslash(ExtractFileDir(Application.ExeName)) + 'SQLAdm.exe';
  if not FileExists(PathSQLAdm) then
  begin
    PGIBox(Format('Fichier %s non trouvé',[PathSQLAdm]));
    exit;
  end;

  if PGIAsk ('Confirmez vous  le traitement ?') = mrYes then
  begin
    { Enregistre les paramètres}
    SetParams;

   { Création du répertoire de travail }
    CreateDir(Params.DirectoryTemp);

    { Attente heure déclenchement }
    DatDebut := StrToDateTime(GetControlText('DATDEBUT') + ' ' + GetControlText('HRDEBUT'));
    DatFin   := StrToDateTime(GetControlText('DATFIN')   + ' ' + GetControlText('HRFIN'));
    if GetCheckBoxState('CHKDEBUT') = cbUnchecked then
    begin
      if NowH < DatDebut then AfficheMessage('Attente heure de déclenchement ' + DateTimeToStr(DatDebut));
      While NowH < DatDebut do
      begin
        if GereStop then exit;
      end;
    end;
    SourisSablier;
    
    { sauvegarde/restauration des sociétés }
    for iInd := 1 to GSoc.RowCount -1 do
    begin
      if not GSoc.IsSelected(iInd) then continue;
      Societe := GSoc.Cells[1,iInd];
      GSoc.Cells[0,iInd] := icoMain;
      Application.ProcessMessages;
      if NowH >= DatFin then
      begin
        AfficheMessage('Fin de période de traitement ' + DateTimeToStr(DatDebut));
        break;
      end
      else
        TraiteLaSociete(Societe, Params);
      GSoc.Cells[0,iInd] := '';
      if GereStop Then break;
    end;
    RemoveDir(Params.DirectoryTemp);
    GSoc.ClearSelected;
    SourisNormale;
  end;
end;


procedure TOF_BACKUPRESTORE.OnLoad;
begin
  inherited;
end;

procedure TOF_BACKUPRESTORE.OnClose;

  function GetPrefixe : string;
  begin
    if bBackupNotRestore then
           result := 'BackupSociete'
      else result := 'RestoreSociete';
  end;
begin
  inherited;
  { enregistrement des clés de registre }
  SetParamToRegistry(GetPrefixe);
end;

procedure TOF_BACKUPRESTORE.OnDisplay () ;
begin
  inherited;
end;

procedure TOF_BACKUPRESTORE.OnCancel () ;
begin
  inherited;
end;

procedure TOF_BACKUPRESTORE.GetParamFromRegistry(PrefixeCle : string);
var stKey : string;
    Reg   : TRegistry;

  procedure SetCle (ControlName, Cle, ValeurDefaut: string) ;
  begin
    stKey := Trim(GetSynRegKey (Cle, '', True));
    if (stKey = '') or (stKey = ':') then
      stKey := ValeurDefaut;
    SetControlText (ControlName, stKey) ;
  end;

begin
  SetCle ('CHKDEBUT',       PrefixeCle + '_ChkDebut', 'X') ;
  SetCle ('HRDEBUT',        PrefixeCle + '_HrDebut', '20:00') ;
  SetCle ('HRFIN',          PrefixeCle + '_HrFin', '08:00') ;
  SetCle ('DIRECTORYBACKUP',PrefixeCle + '_DirectoryBackup', '') ;//'C:\PGI01') ;
  SetCle ('CHKCOMPRESSION', PrefixeCle + '_ChkCompression', '-') ;
  SetCle ('CHKFORCESAUVE',  PrefixeCle + '_ChkForceSauv', 'X') ;
  SetCle ('CHKDELETE',      PrefixeCle + '_ChkDelete', '-') ;
  SetCle ('USER',           PrefixeCle + '_User', 'sa') ;

  { mot de passe CEGID si vide et jamais renseigné }
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.KeyExists('\Software\Microsoft\Cegid\PGIMAJVER\' + PrefixeCle + '_Password') then
         SetCle ('PASSWORD',       PrefixeCle + '_Password', '')
    else SetCle ('PASSWORD',       PrefixeCle + '_Password', 'CEGID');
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
end;

procedure TOF_BACKUPRESTORE.SetParamToRegistry(PrefixeCle : string);
begin
  SaveSynRegKey (PrefixeCle + '_ChkDebut', GetControlText ('CHKDEBUT') , True) ;
  SaveSynRegKey (PrefixeCle + '_HrDebut', GetControlText ('HRDEBUT') , True) ;
  SaveSynRegKey (PrefixeCle + '_HrFin', GetControlText ('HRFIN') , True) ;
  SaveSynRegKey (PrefixeCle + '_DirectoryBackup', GetControlText ('DIRECTORYBACKUP') , True) ;
  SaveSynRegKey (PrefixeCle + '_ChkCompression', GetControlText ('CHKCOMPRESSION') , True) ;
  SaveSynRegKey (PrefixeCle + '_ChkForceSauv', GetControlText ('CHKFORCESAUVE') , True) ;
  SaveSynRegKey (PrefixeCle + '_ChkDelete', GetControlText ('CHKDELETE') , True) ;
  SaveSynRegKey (PrefixeCle + '_User', GetControlText ('USER') , True) ;
  SaveSynRegKey (PrefixeCle + '_Password', GetControlText ('PASSWORD') , True) ;
end;

procedure TOF_BACKUPRESTORE.InitialiseLesCtrl;
begin
  ListSoc := TListBox (GetControl ('LBSOCIETE')) ;
  GSoc := THGrid (GetControl ('GSOC')) ;
  Recap := TListBox (GetControl ('LBRECAP')) ;
  GSoc.Cells [1, 0] := TraduireMemoire ('Société') ;
  GSoc.Cells [2, 0] := TraduireMemoire ('Etat') ;

  TToolbarButton97 (GetControl ('BSTOP') ) .OnClick := bStopClick;

  SetControlText ('DATDEBUT', DateToStr (nowH) ) ;
  SetControlText ('DATFIN', DateToStr (nowH + 1) ) ;

  THEdit(GetControl ('DIRECTORYBACKUP')).OnElipsisClick := LookupDirectoryDrive;
end;

procedure TOF_BACKUPRESTORE.bStopClick(Sender: TObject);
begin
  Stop := True;
end;

procedure TOF_BACKUPRESTORE.LookupDirectoryDrive(Sender: TObject);
var Directory : string;

  function GetNetworkPath() : string;
  begin
    if sServerName <> '' then result := '\\'+ sServerName
    else result := sServerName;
  end;

begin
  Directory := GetControlText('DIRECTORYBACKUP');
  if SelectDirectory(TraduireMemoire('Répertoire de sauvegarde'),GetNetworkPath,Directory) then
     SetControlText('DIRECTORYBACKUP',Directory);
end;

procedure TOF_BACKUPRESTORE.ChargeLesSocietes;
var iInd : integer;
    Societe,CurrentServer : string;
    TobSoc, TobS : tob;
begin
  CurrentServer := GetInfoFromCegidPGIIni(V_PGI.CurrentAlias,'Server');
  ChargeDossier (ListSoc.Items, True) ;
  TobSoc := TOB.Create('',nil,-1);
  try
    for iInd := 0 to ListSoc.Items.Count - 1 do
    begin
      Societe := ListSoc.Items [iInd];
      if (CurrentServer = GetInfoFromCegidPGIIni(Societe,'Server'))
          and (pos('ORACLE', GetInfoFromCegidPGIIni(Societe,'Driver'))<=0) then
      begin
        TobS := TOB.Create('',TobSoc,-1);
        TobS.AddChampSupValeur('SOCIETE',Societe);
      end;
    end;
    TobSoc.PutGridDetail(GSoc,False,False,';SOCIETE',True);
  finally
    TobSoc.Free;
  end;
  HMTrad.ResizeGridColumns (GSoc) ;
end;

function TOF_BACKUPRESTORE.TraiteLaSociete(Societe : string ; Params : T_ParamBackUpRestore) : Boolean;
var
  sSoc : string;
  CopieInfo : T_CopieInfo;

  procedure SetIconeToGrid();
  var iInd : integer;
  begin
    for iInd := 1 to GSoc.RowCount - 1 do
    begin
      if GSoc.Cells [1, iInd] = Societe then
      begin
        case CopieInfo of
          ciOk    : if Params.Zip then GSoc.Cells [2, iInd] := icoZip
                    else
                    begin
                      if Params.BackupNotRestore then GSoc.Cells [2, iInd] := icoSave
                      else GSoc.Cells [2, iInd] := icoRestore;
                    end;
          ciError : GSoc.Cells [2, iInd] := icoAvert;
          ciExist : GSoc.Cells [2, iInd] := icoAvert;
          ciZip   : GSoc.Cells [2, iInd] := icoAvert;
        end;
        break;
      end;
    end;
  end;

begin
  Result := False;
  AfficheMessage ('Connexion société ' + Societe);

  { déconnecte si la base à copier est celle sur laquelle on se trouve }
  if Societe = V_PGI.DbName then
  begin
    if ExisteSQL('select US_UTILISATEUR from UTILISAT where US_PRESENT="X" and US_UTILISATEUR<>"' + V_PGI.User + '"') then
    begin
      AfficheMessage('D''autres utilisateurs sont connectés, impossible de copier la société' + Societe);
      exit;
    end;
    sSoc := V_PGI.CurrentAlias;
    DeconnecteHalley ;
  end;

  { copie (et zip) }
  if Params.backUpNotRestore then
  begin
    CopieInfo := BackupDB(Societe,Params);
    case CopieInfo of
      ciOk    : AfficheMessage ('Sauvegarde effectuée');   //Affiche ('copie ' + ChemSrc) ;    Affiche (' vers ' + ChemDest + '...') ;
      ciError : AfficheMessage ('Impossible de sauvegarder la société ' + Societe,true) ;
      ciExist : AfficheMessage (format('Un fichier %s existe déjà',[GetBackupFile(Societe,Params)]),true) ;
      ciZip   : AfficheMessage ('Echec. La sauvegarde ne peut être compressée',true);
    end;
  end else
  begin
    CopieInfo := RestoreDB(Societe,Params);
    case CopieInfo of
      ciOk    : AfficheMessage ('Restauration terminée');   //Affiche ('copie ' + ChemSrc) ;    Affiche (' vers ' + ChemDest + '...') ;
      ciError : AfficheMessage ('Impossible de restaurer la société ' + Societe,true) ;
      ciExist : AfficheMessage (format('Le fichier %s n''existe pas',[GetBackupFile(Societe,Params)]),true) ;
      ciZip   : AfficheMessage ('Echec. Le fichier ne peut être décompressé',true);
    end;
  end;

  SetIconeToGrid();

  { reconnecte }
  if Societe = V_PGI.DBName then
  begin
    V_PGI.OkOuvert := ConnecteHalley(sSoc, false, nil, nil, nil, nil);
  end;

  { vérifie l'existence du fichier ? à voir... }

  Result := True;
end;

function TOF_BACKUPRESTORE.BackupDB(Societe : string ; Params : T_ParamBackUpRestore)  : T_CopieInfo;
var BackupFile : string;
    Zip: toz;
begin
  result := ciOk;
  BackupFile := GetBackupFile(Societe,Params);

  { vérifie l'existence du fichier }
  if FileExists (BackupFile) then
  begin
    if Params.ForceLaRecopie then DeleteFile (BackupFile)
    else
    begin
      result := ciExist;
      exit;
    end;
  end;

  { sauvegarde }
  if ExecSQLADM(Societe, GetBackupFile(Societe,Params,true),Params) then
  begin
  { zippe }
    if Params.Zip then
    begin
      AfficheMessage ('Compression du fichier');
      result := ciZip;
      try
        Zip := TOZ.Create;
        Zip.NiveauDeCompression := 9;
        if Zip.OpenZipFile (BackupFile, moCreate) then
        begin
          if Zip.OpenSession (osAdd) then
          begin
            Zip.ProcessFile(GetBackupFile(Societe,Params,true));
            Zip.CloseSession;
            DeleteFile(GetBackupFile(Societe,Params,true));
            result := ciOk;
          end;
        end else Zip.CancelSession;
      finally
        FreeAndNil(Zip);
      end;
    end;
  end else
    result := ciError;
end;

function TOF_BACKUPRESTORE.RestoreDB(Societe : string ; Params : T_ParamBackUpRestore)  : T_CopieInfo;
var BackupFile : string;
    Zip: toz;
begin
  result := ciOk;
  BackupFile := GetBackupFile(Societe,Params);

  { vérifie l'existence du fichier }
  if not FileExists (BackupFile) then
  begin
    result := ciExist;
    exit;
  end;

  { dézippe }
  if Params.Zip then
  begin
    AfficheMessage ('Décompression du fichier');
    result := ciZip;
    try
      Zip := TOZ.Create;
      Zip.NiveauDeCompression := 9;
      if Zip.OpenZipFile (BackupFile, moOpen) then
      begin
        if Zip.OpenSession (osExt) then
          if Zip.SetDirOut (Params.DirectoryTemp) then
          begin
            Zip.CloseSession;
            result := ciOk;
          end;
      end else Zip.CancelSession;
    finally
      FreeAndNil(Zip);
    end;
  end;

  { restaure }
  if (result = ciOk) then
    if not ExecSQLADM(Societe, GetBackupFile(Societe,Params,true),Params) then result := ciError;

  { supprime les fichiers }
  if Params.Zip and FileExists (GetBackupFile(Societe,Params,true)) then
     DeleteFile (GetBackupFile(Societe,Params,true));
  if (result = ciOk) and FileExists (BackupFile) then
  begin
    if Params.DeleteAfterRestore then DeleteFile (BackupFile)
    else
    begin
      result := ciExist;
      exit;
    end;
  end;
end;

function TOF_BACKUPRESTORE.ExecSQLAdm(Societe, BackupFile : string ; Params : T_ParamBackUpRestore): boolean;
var sCmdLine: string;
  iHwnd: HWND;
  dwExitCode: DWORD;
begin
  Result := False;
     sCmdLine :=
      IncludeTrailingBackslash(ExtractFileDir(Application.ExeName))
      + 'SQLAdm /USER=' + Params.User + ' /PASSWORD=' + Params.Password
      + ' /SERVER="' + Societe + '"'
      + ' /FILE="' + BackupFile + '"'
      + ' /DATABASE=' + GetInfoFromCegidPGIIni(Societe,'DataBase')
      + ' /VERBOSE'
      + ' /INI=' + HalSocIni;
    if Params.BackupNotRestore then
      sCmdLine := sCmdLine + ' /BACKUP'
    else
      sCmdLine := sCmdLine + ' /RESTORE';

  dwExitCode := 0;
  iHwnd := FileExec(sCmdLine, True, True, False);
  if (iHwnd <> 0) then
  begin
    GetExitCodeProcess(iHwnd, dwExitCode); //contrôle du code retour
    Result := (dwExitCode = 0); //si erreur dans le nom de fichier, le code retour vaut quand même 0 (SQLAdm v1.0.191.1)
  end
  else
    AfficheMessage('Impossible de lancer la commande de sauvegarde/restauration externe',true);
end;

function TOF_BACKUPRESTORE.GetBackupFile(Societe : string ; Params : T_ParamBackUpRestore ; WithoutZip : boolean = false) : string;
begin
  { pour les opérations de compression/décompression, on utilise un répertoire temporaire }
  if (Params.Zip and not WithoutZip) then result := Params.Directory + '\' + Societe + '.zip'
  else if (Params.Zip and WithoutZip) then result := Params.DirectoryTemp + '\' + Societe + '.bak'
  else result := Params.Directory + '\' + Societe + '.bak';
end;

procedure TOF_BACKUPRESTORE.AfficheMessage(msg: String ; Avertissement : boolean = false);
begin
  if Avertissement then Recap.Items.Add('***  ' + TraduireMemoire(msg) + '  ***')
  else Recap.Items.Add(TraduireMemoire(msg));
  Recap.ItemIndex := Recap.Items.Count-1;
  Recap.Refresh;
  Application.ProcessMessages;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JSI
Créé le ...... : 14/09/2005
Modifié le ... :   /  /    
Description .. : Décode le fichier CEGIDPGI.ini
Mots clefs ... : 
*****************************************************************}
function TOF_BACKUPRESTORE.GetInfoFromCegidPGIIni(Societe,QuelleInfo : string) : string;
var iHalSocIni : TIniFile;
begin
  iHalSocIni := TIniFile.Create(GetHalSocIni);
  try
    result := iHalSocIni.ReadString(Societe, QuelleInfo, EmptyStr);
  finally
    iHalSocIni.Free;
  end;
end;

initialization
  registerclasses ([TOF_BACKUPRESTORE] ) ;
end.
