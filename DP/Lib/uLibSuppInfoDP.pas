unit uLibSuppInfoDP;

interface

uses
  Classes, Hent1, uDbxDataSet, hctrls, Utob, eSession,
  SysUtils;

type
  TDeleteDoss = Class(TObject)
    GuidPer    : String;
    NoDossier  : String;
    ErrorMess  : String;
    ForceSuppr : Boolean;
  private
    ListChpDPSTD : TStringList;
    function  SuppressionBasePhysique: Boolean;
    function  DeleteMessage(sMsgGuid: String): Boolean;
    function  DeleteDansTableMessage(sMsgGuid : String): Boolean;
    procedure SupprimeDocumentGed(SDocGUID: string);
    function  SupprimeDocument(SDocGUID : string): Boolean;
    procedure SuppressionEnregDPSTD;
    procedure ChargeListChpDPSTD(lst: TStringList);
  public
    function    SuppressionDossier: Boolean;
    constructor Create;
    destructor  Destroy; override;
  end;

implementation
uses GalSystem;

function TDeleteDoss.SuppressionBasePhysique: Boolean;
var
  StDatabase : String;
  supprok : Boolean;
begin
  Result := FALSE;
  LogEAGL('SuppressionBasePhysique...');

  if (NoDossier = '') then
  begin
    ErrorMess := 'Le dossier n''a pas d''informations de connexions suffisantes.'+#13#10
                +'La base n''existe pas ou ne sera pas supprimée...';
    exit;
  end;

  if (V_PGI.Driver=dbMsSQL) or (V_PGI.Driver=dbMSSQL2005) then
  begin
    // nom de la DB
    StDatabase := 'DB'+NoDossier;
    supprok := False;
    try
      supprok := SupprimeBaseSql(StDatabase);
    except
      ErrorMess := 'Une erreur s''est produite à l''exécution de la fonction "SupprimeBaseSql"';
    end;

    if not supprok then
    begin
      ErrorMess := 'La base de données '+StDatabase+' n''a pas pu être supprimée.';
      exit;
    end;
  end
  // ####DRIVER NON SUPPORTE
  else
  begin
    // normalement, cas déjà éliminé lors de la création...
    ErrorMess := 'Le driver "'+DriverToSt(V_PGI.Driver, True)+'" de la connexion en cours n''est pas supporté dans un environnement multi-dossier.';
    exit;
  end;

  Result := True;
  LogEAGL('SuppressionBasePhysique... OK');
end;
////////////////////////////////////////////////////////////////////////////////
function TDeleteDoss.DeleteMessage(sMsgGuid : String): Boolean;
var
  sUserMail : String;
  iMailId   : Integer;
  Q         : TQuery;
begin
  Result := False;
  LogEAGL('DeleteMessage...');

  if sMsgGuid = '' then
    exit;

  // Clé pour recherche YMAILS associé
  sUserMail := '';
  iMailId   := 0;
  Q := OpenSQL('SELECT YMS_USERMAIL, YMS_MAILID FROM YMESSAGES WHERE YMS_MSGGUID="'+sMsgGuid+'"', True);
  if Not Q.Eof then
  begin
    sUserMail := Q.FindField('YMS_USERMAIL').AsString;
    iMailId   := Q.FindField('YMS_MAILID').AsInteger;
  end;
  Ferme(Q);

  Result := DeleteDansTableMessage(sMsgGuid);
  LogEAGL('DeleteMessage... OK');
  //if (Result) and (iMailId <> 0) then
  //  UserMailDelete(iMailId, sUserMail); Déclaration UmailBox.pas
end;
////////////////////////////////////////////////////////////////////////////////
function TDeleteDoss.DeleteDansTableMessage(sMsgGuid: String): Boolean;
var
  tMsg, tResult: Tob;
begin
  LogEAGL('DeleteDansTableMessage...');
  //$$$ JP 06/12/05 - warning delphi -> Result := False;
  tMsg := Tob.Create('_YMESSAGES_', nil, -1);
  tResult := Tob.Create('_RESULT_', nil, -1);

  try
    tMsg.LoadDetailDBFromSQL('YMESSAGES', 'SELECT * FROM YMESSAGES WHERE (YMS_MSGGUID="'+sMsgGuid+'")');

    if tResult.Detail.Count > 0 then
    begin
//      FLastErrorMsg := tResult.Detail[1].GetValue('ERROR');
      Result := False;
    end
    else
      Result := True;

  except
//    FLastErrorMsg := 'Erreur dans le processus';
    Result := False;
  end;

  tMsg.Free;
  tResult.Free;
  LogEAGL('DeleteDansTableMessage... OK');
end;
////////////////////////////////////////////////////////////////////////////////
procedure TDeleteDoss.SupprimeDocumentGed(SDocGUID: string);
begin
  LogEAGL('SupprimeDocumentGed...');
  // pour l'instant pas besoin de tom sur DPDOCUMENT (niveau le plus haut)
  ExecuteSQL('DELETE DPDOCUMENT WHERE DPD_DOCGUID="' + SDocGUID + '"' );

  // suppression dans YDOCUMENTS, YDOCFILES, YFILES, YFILEPARTS :
  SupprimeDocument(SDocGUID);
  LogEAGL('SupprimeDocumentGed... OK');
end;
////////////////////////////////////////////////////////////////////////////////
function TDeleteDoss.SupprimeDocument(SDocGUID : string): Boolean;
var
  tDocs, tResult: Tob;
begin
  LogEAGL('SupprimeDocument...');
  tDocs := Tob.Create('_YDOCUMENTS_', nil, -1);
  tResult := Tob.Create('_RESULT_', nil, -1);

  try
    tDocs.LoadDetailDBFromSQL('YDOCUMENTS', 'SELECT * FROM YDOCUMENTS WHERE (YDO_DOCGUID = "' + SDocGUID + '")');
    tDocs.DeleteDB;

    if tResult.Detail.Count > 0 then
    begin
      Result := False;
    end
    else
      Result := True;

  except
//    FLastErrorMsg := 'Erreur dans le processus';
    Result := False;
  end;

  tDocs.Free;
  tResult.Free;
  LogEAGL('SupprimeDocument... OK');
end;
////////////////////////////////////////////////////////////////////////////////
procedure TDeleteDoss.ChargeListChpDPSTD(lst: TStringList);
// charge la liste des chps des tables communes (DP ou STD)
// de type XX_NODOSSIER ou XX_GUIDPER (donc liés aux dossiers)
var Q: TQuery;
begin
  LogEAGL('ChargeListChpDPSTD...');
  if lst = nil then
    exit;

  // liste des champs dépendant d'un no de dossier, ou d'un code personne
  Q := OpenSQL('SELECT DH_NOMCHAMP FROM DETABLES INNER JOIN DECHAMPS ON DT_PREFIXE=DH_PREFIXE '
              + 'WHERE (DT_COMMUN="D" OR DT_COMMUN="S") '
                + 'AND (DH_NOMCHAMP LIKE "%[_]NODOSSIER" OR DH_NOMCHAMP LIKE "%[_]GUIDPERDOS" '
                  + 'OR DH_NOMCHAMP LIKE "%[_]GUIDPER" OR DH_NOMCHAMP LIKE "%[_]GUIDLIA")', True, -1, '', True);
  while not Q.Eof do
  begin
    lst.Add(Q.FindField('DH_NOMCHAMP').AsString);
    Q.Next;
  end;
  Ferme(Q);
  LogEAGL('ChargeListChpDPSTD... OK');
end;
////////////////////////////////////////////////////////////////////////////////
procedure TDeleteDoss.SuppressionEnregDPSTD;
var
  i: Integer;
  nomtbl, prefix, sql : String;
begin
  LogEAGL('SuppressionEnregDPSTD...');
  for i:=1 to High(V_PGI.HDETables[LookupCurrentSession.SocNum]) do
  begin
    if V_PGI.DeTables[i].TypeDP = tdNone then
      Continue;

    nomtbl := V_PGI.DeTables[i].Nom;

    if (nomtbl='TRFDOSSIER') or (nomtbl='ANNUAIRE') or (nomtbl='DOSSAPPLI') then
      Continue;

    prefix := V_PGI.DeTables[i].Prefixe;
    Sql    := '';

    if ListChpDPSTD.IndexOf(prefix+'_NODOSSIER') > -1 then
    begin
      if NoDossier <> '' then
      begin
        sql := 'DELETE FROM '+nomtbl+' WHERE '+prefix+'_NODOSSIER="'+NoDossier+'"';
        if nomtbl = 'JUEVENEMENT' then
          sql := sql +' AND JEV_CODEDOS="&#@"';
      end;
    end
    else if ListChpDPSTD.IndexOf(prefix+'_GUIDPERDOS') > -1 then
    begin
      sql := 'DELETE FROM '+nomtbl+' WHERE '+prefix+'_GUIDPERDOS="'+GuidPer+'"';
      if nomtbl = 'ANNULIEN' then
        sql := sql + ' AND ANL_CODEDOS="&#@"'
      else if nomtbl = 'JURIDIQUE' then
        sql := sql + ' AND JUR_CODEDOS="&#@"';
    end;

    // TRAITEMENT
    if sql <> '' then
      ExecuteSQL(sql);
  end;

  if NoDossier <> '' then
    ExecuteSQL('DELETE FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+NoDossier+'"');

  LogEAGL('SuppressionEnregDPSTD... OK');
end;
////////////////////////////////////////////////////////////////////////////////
function TDeleteDoss.SuppressionDossier: Boolean;
var
  supprok: Boolean;
  Q : TQuery;
  lstGuid  :TStringList;
  i : Integer;
begin
  LogEAGL('SuppressionDossier...');
  // Par défaut, dossier non supprimé
  Result := False;

  // 1. TESTS SYSTEMES/FONCTIONNELS
  if ExisteSQL ('SELECT 1 FROM YMESSAGES WHERE YMS_NODOSSIER="' + NoDossier + '" AND YMS_TRAITE="-"') then
  begin
    ErrorMess := 'Le dossier ' + NoDossier + ' comporte un ou plusieurs messages non clôturés.'+#13#10+'Suppression refusée.';
    exit;
  end
  else
  begin
    // 2. SUPPRESSION BASE PHYSIQUE
    supprok := not DBExists ('DB' + NoDossier);
    if supprok = FALSE then
    begin
      try
        supprok := SuppressionBasePhysique;
      except
        ErrorMess := 'Erreur pendant la suppression de la base du dossier ' + NoDossier;
      end;
    end;

    // Suppression du log résiduel
    // ExecuteMSSQL7('xp_cmdshell ''del '+VH_Doss.PathLdf+'\'+VH_Doss.DBSocName+'.ldf''');

    // 3. SUPPRESSION ELEMENTS LIES A LA GED
    if supprok = TRUE then
    begin
      // On détruit dans les tables les plus hautes, les toms font la suppression en cascade
      // MESSAGES
      // $$$ JP 01/08/06: même gestion pour suppr' des messages
      lstGuid := TStringList.Create;
      Q := OpenSQL('SELECT YMS_MSGGUID FROM YMESSAGES WHERE YMS_NODOSSIER="'+NoDossier+'"', True, -1, '', True);
      While not Q.Eof do
      begin
        lstGuid.Add (Trim (Q.FindField ('YMS_MSGGUID').AsString));
        Q.Next;
      end;
      Ferme(Q);

      for i := 0 to lstGuid.Count-1 do
      begin
        if lstGuid [i] <> '' then
          DeleteMessage(lstGuid [i]);
      end;
      lstGuid.Clear;

      // DOCUMENTS
      Q := OpenSQL('SELECT DPD_DOCGUID FROM DPDOCUMENT WHERE DPD_NODOSSIER="'+NoDossier+'"', True, -1, '', True);
      While not Q.Eof do
      begin
        lstGuid.Add (Q.FindField('DPD_DOCGUID').AsString);
        Q.Next;
      end;
      Ferme(Q);

      for i := 0 to lstGuid.Count-1 do
      begin
        SupprimeDocumentGed(lstGuid [i]);
      end;
      lstGuid.Free;
    end;

    // 5. SUPPRESSION ENREG DP/DOSSIER
    if supprok = TRUE then
    begin
      ListChpDPSTD := TStringList.Create;

      // charge liste de chps en dehors des transactions car select sur tables systèmes
      ChargeListChpDPSTD (ListChpDPSTD);
      if Transactions (SuppressionEnregDPSTD, 3) = oeOK then //if Transactions(SuppressionEnregDPSTD,1) = oeOK then
      begin
        Result := True; // suppression ok

        //RemoveInDir (VH_Doss.PathDos, True, True);
        //ExecuteMSSQL7('xp_cmdshell ''rmdir '+VH_Doss.PathLdf+'''');
      end;
    end;
  end;
  LogEAGL('SuppressionDossier... OK');
end;
////////////////////////////////////////////////////////////////////////////////
constructor TDeleteDoss.Create;
begin
  ListChpDPSTD := Nil;
end;
////////////////////////////////////////////////////////////////////////////////
destructor TDeleteDoss.Destroy;
begin
  if ListChpDPSTD<>Nil then ListChpDPSTD.Free;
  inherited;
end;
////////////////////////////////////////////////////////////////////////////////
end.
