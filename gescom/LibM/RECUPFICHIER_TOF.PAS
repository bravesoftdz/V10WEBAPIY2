{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 24/01/2002
Modifié le ... : 24/01/2002
Description .. : Source TOF de la FICHE : RECUPFICHIER ()
Suite ........ :
Suite ........ : Récupération des fichiers de MAJ qui se trouvent sur le
Suite ........ : serveur par le biais du serveur eAGL pour les mettre dans
Suite ........ : les boutiques
Mots clefs ... : TOF;RECUPFICHIER
*****************************************************************}
unit RECUPFICHIER_TOF;

interface

uses StdCtrls, Controls, Classes, Windows,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  db, dbtables, Fe_Main,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
  PARAMSITEDIST_TOF, Ras, Vierge, HTB97, LicUtil, M3FP,
  uHttp, uHttpCS, ParamSoc, ed_tools, FileCtrl, UTOB;

procedure LanceMiniFTP;

type
  TOF_RECUPFICHIER = class(TOF)
  private
    StModeConnexion, StServerAdd, RasEntry, RasUser, RasPass, RasTel: string;
    IsConnected: boolean;
    hRas: THandle;
    DossierLocal, DossierDistant: string;
    GDistant, GLocal: THGrid;
    aServer: THttpServer;
    function LanceConnexion(BDemarrer: boolean): boolean;
    procedure RefreshDossiers(QuelDossier: integer);
    procedure TransfererFichiers;
    procedure TransfereUnFichier(FileName, FileDateModif: string);
    procedure ToutSelectionner(Action: string);
    procedure MAJCheminDossier(QuelChemin: integer);
    procedure AfficheContenuDossier;
    procedure AfficheDossierLocal;
    function OctetToKo(Octets: Integer): string;
  public
    procedure OnArgument(S: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  end;

implementation

procedure LanceMiniFTP;
begin
  AGLLanceFiche('MBO', 'RECUPFICHIER', '', '', '');
end;

procedure TOF_RECUPFICHIER.MAJCheminDossier(QuelChemin: integer);
var StTemp: string;
begin
  case QuelChemin of
    1:
      begin
        StTemp := GetControlText('EDIRECTORYDISTANT');
        if (StTemp <> '') and (StTemp <> DossierDistant) then
        begin
          DossierDistant := IncludeTrailingBackslash(StTemp);
          SetParamSoc('SO_RFDOSSIERDISTANT', DossierDistant);
        end;
      end;
    2:
      begin
        StTemp := GetControlText('EDIRECTORYLOCAL');
        if (StTemp <> '') and (StTemp <> DossierLocal) then
        begin
          DossierLocal := IncludeTrailingBackslash(StTemp);
          SetParamSoc('SO_RFDOSSIERLOCAL', DossierLocal);
        end;
      end;
  end;
end;

procedure TOF_RECUPFICHIER.TransfereUnFichier(FileName, FileDateModif: string);
var Fichier: TFileStream;
  ContentType, StFichier, ErrFileName: string;
  Ligne: integer;
  OkCopie: boolean;
  TheDate, TheTime, TheDateModif: TDateTime;
begin
  OkCopie := True;
  for Ligne := 1 to GLocal.RowCount - 1 do
  begin
    if GLocal.Cells[1, Ligne] = FileName then
    begin
      if GLocal.Cells[3, Ligne] = FileDateModif then
        if PGIAsk(TraduireMemoire('Le fichier ' + FileName +
          ' existe déjà en local avec la même date de modif, voulez-vous le remplacer ?'),
          Ecran.Caption) = MrNo then OkCopie := False;
      Break;
    end;
  end;
  if not OkCopie then exit;
  ErrFileName := '';
  StFichier := aServer.RequestFile(DossierDistant + FileName, ContentType);
  if ContentType = 'text/html' then ErrFileName := '_error.htm';
  Fichier := TFileStream.Create(DossierLocal + FileName + ErrFileName, fmCreate);
  Fichier.Write(StFichier[1], SizeOf(StFichier[1]) * Length(StFichier));

  TheDate := EncodeDate(StrToInt(Copy(FileDateModif, 7, 4)),
    StrToInt(Copy(FileDateModif, 4, 2)), StrToInt(Copy(FileDateModif, 1, 2)));
  TheTime := EncodeTime(StrToInt(Copy(FileDateModif, 12, 2)),
    StrToInt(Copy(FileDateModif, 15, 2)), 0, 0);
  TheDateModif := TheDate + TheTime;
  FileSetDate(Fichier.Handle, DateTimeToFileDate(TheDateModif));

  Fichier.Free;
end;

procedure TOF_RECUPFICHIER.TransfererFichiers;
var MsgServeur, NomFichier: string;
  i, NbFichiers: integer;
begin
  NbFichiers := GDistant.NBSelected;
  if NbFichiers = 0 then
  begin
    PGIError(TraduireMemoire('Aucun fichier à transférer.'), Ecran.Caption);
    exit;
  end;

  if not IsConnected then
    if not LanceConnexion(True) then exit;

  aServer := ConnectHttpServer(StServerAdd, MsgServeur);
  if MsgServeur <> '' then
  begin
    PGIInfo(TraduireMemoire(MsgServeur), ecran.caption);
    exit;
  end;

  if (aServer <> nil) and (aServer.IsConnected) then
  begin
    InitMoveProgressForm(Ecran, Ecran.Caption,
      TraduireMemoire('Copie des fichiers sélectionnés.'),
      NbFichiers, True, True);
    try
      for i := 1 to GDistant.RowCount - 1 do
      begin
        if GDistant.IsSelected(i) then
        begin
          NomFichier := GDistant.Cells[1, i];
          if not MoveCurProgressForm(TraduireMemoire('Copie du fichier : ' + NomFichier)) then Break
          else TransfereUnFichier(NomFichier, GDistant.Cells[3, i]);
        end;
      end;
    finally
      FiniMoveProgressForm;
      DeconnectHttpServer(aServer);
    end;
    RefreshDossiers(2);
    ToutSelectionner('Aucun');
  end;
end;

procedure TOF_RECUPFICHIER.ToutSelectionner(Action: string);
begin
  if Action = 'Tous' then GDistant.AllSelected := True
  else
    if Action = 'Aucun' then GDistant.AllSelected := False
  else GDistant.AllSelected := not GDistant.AllSelected;
  if GDistant.AllSelected then SetControlProperty('BSELECTALL', 'Hint', 'Tout déselectionner')
  else SetControlProperty('BSELECTALL', 'Hint', 'Tout sélectionner');
end;

procedure TOF_RECUPFICHIER.RefreshDossiers(QuelDossier: integer);
var MsgServeur: string;
begin
  case QuelDossier of
    1:
      begin
        if not IsConnected then
          if not LanceConnexion(True) then exit;
        if DossierDistant = '' then Exit;

        aServer := ConnectHttpServer(StServerAdd, MsgServeur);
        if MsgServeur <> '' then
        begin
          PGIInfo(TraduireMemoire(MsgServeur), ecran.caption);
          exit;
        end;

        if (aServer <> nil) and (aServer.IsConnected) then
        begin
          try
            AfficheContenuDossier;
          finally DeconnectHttpServer(aServer);
          end;
        end;
      end;
    2:
      begin
        if DossierLocal = '' then Exit;
        if not DirectoryExists(DossierLocal) then
          if not CreateDir(DossierLocal) then
          begin
            raise Exception.Create(TraduireMemoire('Impossible de créer ' + DossierLocal));
            Exit;
          end;

        AfficheDossierLocal;
      end;
  end;
end;

procedure TOF_RECUPFICHIER.AfficheDossierLocal;
var TOBDossier, TOBFichier: TOB;
  Sr: TSearchRec;
  Ext: string;
begin
  GLocal.VidePile(False);
  TOBDossier := TOB.Create('', nil, -1);
  try
    if FindFirst(DossierLocal + '*.*', 0, sr) = 0 then //faAnyFile
    begin
      repeat
        Ext := Uppercase(ExtractFileExt(sr.Name));
        if Ext <> '' then
        begin
          TOBFichier := TOB.Create('', TOBDossier, -1);
          with TOBFichier do
          begin
            AddChampSup('NOM', False);
            PutValue('NOM', Sr.Name);
            AddChampSup('TAILLE', False);
            PutValue('TAILLE', OctetToKo(Sr.Size));
            AddChampSup('MODIF', False);
            PutValue('MODIF', FormatDateTime('dd/mm/yyyy t', FileDateToDateTime(Sr.Time)));
          end;
        end;
      until FindNext(Sr) <> 0;
      TOBDossier.PutGridDetail(GLocal, False, False, 'NOM;TAILLE;MODIF');
      FindClose(Sr);
    end;
  finally
    TOBDossier.Free;
  end;
end;

function TOF_RECUPFICHIER.OctetToKo(Octets: Integer): string;
var Ko: integer;
begin
  Ko := Octets div 1024;
  if Ko = 0 then Result := IntToStr(Octets) else Result := IntToStr(Ko) + 'K';
end;

procedure TOF_RECUPFICHIER.AfficheContenuDossier;
var TOBDossier, TOBFichier: TOB;
  ContentType, St: string;
  ListeFichiers: TStringList;
  i: integer;
  StTaille, StNomFichier, StModif, Ext: string;
begin
  GDistant.VidePile(False);
  ListeFichiers := TStringList.Create;
  TOBDossier := TOB.Create('', nil, -1);
  try
    ListeFichiers.Text := aServer.RequestFile(DossierDistant, ContentType);
    for i := 0 to ListeFichiers.Count - 1 do
    begin
      St := ListeFichiers.Strings[i];
      if Pos('Error 404', St) > 0 then
      begin
        PGIError(TraduireMemoire('Ce répertoire n''existe pas sur le central.'), Ecran.Caption);
        exit;
      end;
      StTaille := Trim(Copy(St, 19, 8));
      StNomFichier := Trim(Copy(St, 28, Length(St) - 27));
      Ext := Uppercase(ExtractFileExt(StNomFichier));
      if (StTaille <> 'Folder') and (Ext <> '.GZ') and (Ext <> '.HTM')
        and (Ext <> '.HTML') and (Ext <> '') then
      begin
        StModif := Copy(St, 2, 16);
        TOBFichier := TOB.Create('', TOBDossier, -1);
        with TOBFichier do
        begin
          AddChampSup('NOM', False);
          PutValue('NOM', StNomFichier);
          AddChampSup('TAILLE', False);
          PutValue('TAILLE', StTaille);
          AddChampSup('MODIF', False);
          PutValue('MODIF', StModif);
        end;
      end;
    end;
    TOBDossier.PutGridDetail(GDistant, False, False, 'NOM;TAILLE;MODIF');
  finally
    ListeFichiers.Free;
    TOBDossier.Free;
  end;
end;

function TOF_RECUPFICHIER.LanceConnexion(BDemarrer: boolean): boolean;
var BConnexion: boolean;
  Reponse: DWORD;
  StDossier, StUser, StPassword: string;
begin
  result := False;
  SourisSablier;
  if (StModeConnexion = '') then
  begin
    ChargeParamSiteDistant(StServerAdd, StDossier, StUser, StPassword, StModeConnexion,
      RasEntry, RasUser, RasPass, RasTel);
    if StModeConnexion = '' then
    begin
      PGIError(TraduireMemoire('Votre mode de connexion au site distant n''est pas correctement paramètré !'), ecran.caption);
      SourisNormale;
      exit;
    end;
  end;
  if StModeConnexion = '002' then
  begin
    BConnexion := AglGetActiveRAS(RasEntry, hRas);
    if BDemarrer then
    begin
      if not BConnexion then
      begin
        Reponse := AglInitRasDial(RasEntry, RasTel, RasUser, RasPass, hRas, nil);
        if (RasEntry = '') or (Reponse <> 0) then
        begin
          PGIError(TraduireMemoire('Votre accès distant n''est pas correctement paramètré !'), ecran.caption);
          SourisNormale;
          exit;
        end;
      end;
      IsConnected := True;
    end
    else
    begin
      if BConnexion then
      begin
        Reponse := AglRasHangup(hRas);
        if Reponse > 600 then
        begin
          PGIInfo(TraduireMemoire('La connexion n''a pas été fermée correctement, veuillez la fermer manuellement.'), ecran.caption);
          SourisNormale;
          exit;
        end;
      end;
      IsConnected := False;
    end;
  end
  else IsConnected := BDemarrer;
  result := True;
  SetControlVisible('BCONNECT', not BDemarrer);
  SetControlVisible('BDECONNECT', BDemarrer);
  SourisNormale;
end;

procedure TOF_RECUPFICHIER.OnUpdate;
begin
  inherited;
  TransfererFichiers;
end;

procedure TOF_RECUPFICHIER.OnArgument(S: string);
begin
  inherited;
  GDistant := THGrid(GetControl('GDISTANT'));
  GLocal := THGrid(GetControl('GLOCAL'));
  GDistant.Cells[1, 0] := 'Nom';
  GDistant.Cells[2, 0] := 'Taille';
  GDistant.Cells[3, 0] := 'Modifié le';
  GLocal.Cells[1, 0] := 'Nom';
  GLocal.Cells[2, 0] := 'Taille';
  GLocal.Cells[3, 0] := 'Modifié le';
  GDistant.ColAligns[2] := taRightJustify;
  GLocal.ColAligns[2] := taRightJustify;
  GDistant.ColAligns[3] := taCenter;
  GLocal.ColAligns[3] := taCenter;

  DossierDistant := IncludeTrailingBackslash(GetParamSoc('SO_RFDOSSIERDISTANT'));
  SetControlText('EDIRECTORYDISTANT', DossierDistant);
  DossierLocal := IncludeTrailingBackslash(GetParamSoc('SO_RFDOSSIERLOCAL'));
  SetControlText('EDIRECTORYLOCAL', DossierLocal);
end;

procedure TOF_RECUPFICHIER.OnClose;
begin
  inherited;
  LanceConnexion(False);
end;

///////////////////////////////////////////////////////////////////////////////

procedure TOFConnectSiteDist(Parms: array of variant; nb: integer);
var F: TFVierge;
  Latof: TOF;
begin
  F := TFVierge(Integer(Parms[0]));
  if (F is TFVierge) then Latof := TFVierge(F).Latof else exit;
  if (Latof is TOF_RECUPFICHIER) then
  begin
    if UpperCase(string(Parms[1])) = 'TRUE' then TOF_RECUPFICHIER(Latof).LanceConnexion(True)
    else TOF_RECUPFICHIER(Latof).LanceConnexion(False);
  end;
end;

procedure TOFParamSiteDist(Parms: array of variant; nb: integer);
var F: TFVierge;
  Latof: TOF;
  LesParametres, UnParam, StParam, ValParam: string;
  x: integer;
begin
  LesParametres := '';
  F := TFVierge(Integer(Parms[0]));
  if (F is TFVierge) then Latof := TFVierge(F).Latof else exit;
  if (Latof is TOF_RECUPFICHIER) then
  begin
    LesParametres := AGLLanceFiche('MBO', 'PARAMSITEDIST', '', '', '');
    TToolbarButton97(TOF_RECUPFICHIER(Latof).GetControl('BPARAMCONNECT')).Down := False;
    while LesParametres <> '' do
    begin
      UnParam := READTOKENST(LesParametres);
      if UnParam <> '' then
      begin
        x := pos('=', UnParam);
        if x <> 0 then
        begin
          StParam := copy(UnParam, 1, x - 1);
          ValParam := copy(UnParam, x + 1, length(UnParam));
          with TOF_RECUPFICHIER(Latof) do
          begin
            if StParam = 'SO_SDSERVERADDRESS' then StServerAdd := ValParam
            else
              if StParam = 'SO_SDMODECONNEXION' then StModeConnexion := ValParam
            else
              if StParam = 'SO_SDPROFIL' then RasEntry := ValParam
            else
              if StParam = 'PROFILUSER' then RasUser := ValParam
            else
              if StParam = 'PROFILPASS' then RasPass := DeCryptageSt(ValParam)
            else
              if StParam = 'PROFILTEL' then RasTel := ValParam
              ;
          end;
        end;
      end;
    end;
  end;
end;

procedure TOFRefreshDossiers(Parms: array of variant; nb: integer);
var F: TFVierge;
  Latof: TOF;
begin
  F := TFVierge(Integer(Parms[0]));
  if (F is TFVierge) then Latof := TFVierge(F).Latof else exit;
  if (Latof is TOF_RECUPFICHIER) then
  begin
    TOF_RECUPFICHIER(Latof).RefreshDossiers(1);
    TOF_RECUPFICHIER(Latof).RefreshDossiers(2);
  end;
end;

procedure TOFTransfererFichiers(Parms: array of variant; nb: integer);
var F: TFVierge;
  Latof: TOF;
begin
  F := TFVierge(Integer(Parms[0]));
  if (F is TFVierge) then Latof := TFVierge(F).Latof else exit;
  if (Latof is TOF_RECUPFICHIER) then TOF_RECUPFICHIER(Latof).TransfererFichiers;
end;

procedure TOFToutSelectionner(Parms: array of variant; nb: integer);
var F: TFVierge;
  Latof: TOF;
begin
  F := TFVierge(Integer(Parms[0]));
  if (F is TFVierge) then Latof := TFVierge(F).Latof else exit;
  if (Latof is TOF_RECUPFICHIER) then TOF_RECUPFICHIER(Latof).ToutSelectionner('');
end;

procedure TOFMAJCheminDossier(Parms: array of variant; nb: integer);
var F: TFVierge;
  Latof: TOF;
begin
  F := TFVierge(Integer(Parms[0]));
  if (F is TFVierge) then Latof := TFVierge(F).Latof else exit;
  if (Latof is TOF_RECUPFICHIER) then
    TOF_RECUPFICHIER(Latof).MAJCheminDossier(Integer(Parms[1]));
end;

initialization
  registerclasses([TOF_RECUPFICHIER]);
  RegisterAglProc('ConnectSiteDist', True, 1, TOFConnectSiteDist);
  RegisterAglProc('ParamSiteDist', True, 0, TOFParamSiteDist);
  RegisterAglProc('RefreshDossiers', True, 0, TOFRefreshDossiers);
  RegisterAglProc('TransfererFichiers', True, 0, TOFTransfererFichiers);
  RegisterAglProc('ToutSelectionner', True, 0, TOFToutSelectionner);
  RegisterAglProc('MAJCheminDossier', True, 1, TOFMAJCheminDossier);
end.
