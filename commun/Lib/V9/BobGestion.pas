{ NOUVELLE GESTION DES BOBS
  MIS EN PLACE PBASSET LE 01/02/2006
  certaines fonctions se trouvaient dans galoutil
}
unit BobGestion;

interface

uses
  Sysutils,
  Dialogs,
  Forms,
  classes,
  Controls,
  ubob,
  hent1,
  HCtrls,
  HMsgBox,
{$IFNDEF EAGLSERVER}
  CBPPath, //
{$ENDIF}
{$IFDEF eAGLClient}
  // uLanceProcess,
{$IFNDEF ENTREPRISE}
  IntegreBobUtil, // dans commun/lib
{$ENDIF}
{$ENDIF}
//  FileCtrl,
  windows,
  // MB : uBobUtiles a rajouter de COMMUN\LIB
  uBobUtiles,
  // MB :  galPatience a rajouter de CorrectionsV800\CegidPgi\Lib
  galPatience,
  cbptrace,
  extctrls;

function BOB_IMPORT_PCL(CodeProduit: string; TypeBOZ: boolean = False): Integer;
function BOB_IMPORT_PCL2(CodeProduit: string; sLibelle: string; TypeBOZ: boolean = False; NumBasePourMajForce: string = ''): Integer;

function BOB_IMPORT_PCL_STD(CodesProduits: string; sLibelle: string; GestionBOBCommun: Boolean = False; ShowWaitForm : Boolean = True; Showmessage : boolean = true ): Integer;
function BOB_IMPORT_PCL_STD_MANU(FiLeName: string): Integer;

function BOB_IMPORT_PCL_STD_MANU_EXISTS(FiLeName: string): boolean;
// function GetWindowsDrive: string; //PB LE 02/10/2007 : fonction temporaire car non portée en WA

//function BOB_IMPORT_PCL_STD(CodeProduit: string; sLibelle: string; RenameAfter : Boolean=False): Integer;
function BOB_GET_CHEMIN: string;
procedure BOB_IMPORT_STD_MENU(ShowWaitForm : Boolean = True);

function BOB_IMPORT_PCL_DB(CodesProduits: string; sLibelle: string; DBName : String ; GestionBOBCommun: Boolean = False; ShowWaitForm : Boolean = True; Showmessage : boolean = true ; MessageSupplementaire : String = '' ): Integer;
// MB : A porter demain en CWAS ....

{$IFNDEF EAGLCLIENT}
// pourrait marcher dans les deux modes ...
procedure BOB_DemandeSuppression(sBob: string);
// Marche qu'en 2/3
procedure BOB_PCL_GARDE(CodeProduit, NumVersionDuBob: string);
procedure BOB_RENAME_BOB(CodeProduit: string; TypeBOZ: boolean = false);
function  BOB_Alerte(CodeProduit: string; NomProduit: string = ''; nb: integer = 3; poids: integer = 0; GestionBOBCommun: Boolean = true): integer;
{$ELSE}
function  CWAS_EXECSCRIPTSQL(nomscript, nomlog, nombase: String; FP : TFPatience; bPurgeDesCaches: Boolean=False): boolean;
{$ENDIF}


////////////////// IMPLEMENTATION ///////////////////
implementation
uses
{$IFNDEF EAGLCLIENT}
  db,
  uDbxDataSet,
{$ENDIF EAGLCLIENT}
  uTob;

{$IFDEF eAGLClient}

function IntegreBOBAuto(CodesProduits: string; DBName : String ; GestionBOBCommun: Boolean ; FP : TFPatience ; ShowWaitForm : boolean = true ; Showmessage : Boolean = True): Integer;
{$IFNDEF ENTREPRISE}
var
  TobResult, TobParam: Tob;
  Count: Integer;
{$ENDIF ENTREPRISE}
begin
  Result := -3;  //FP 07/05/2010 Suppression des avertissements
{$IFNDEF ENTREPRISE}
  SourisSablier;
  //Result := -3;
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  if ShowWaitForm then
      if Assigned(FP) then
        if not FP.Visible then
           FP.Show ;

  // GesCom := '-' ;
  // if GestionBOBCommun then Gescom := 'X' ;
  if DBName = '' then
     TobResult := MyIntegreBob.INTEGREAUTO(CodesProduits, GestionBOBCommun)
  else
     TobResult := MyIntegreBob.INTEGREAUTODB(CodesProduits, GestionBOBCommun,DBName);
  //TobResult := LanceProcessServer('cgiIntegreBOB', 'INTEGREAUTO', CodesProduits + ';' + GesCom, TobParam, True);

  if ShowWaitForm then
    if Assigned(FP) then
      if FP.Visible then
         FP.Hide ;

  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    Count := 0;
    if TobResult.FieldExists('COUNT') then
    begin
      Count := TobResult.GetInteger('COUNT');
      Result := 0;
    end;
    if V_PGI.SAV
      and TobResult.FieldExists('RESULTS') then
      if Showmessage then
         PGIInfo(TobResult.GetString('RESULTS'));
    if TobResult.FieldExists('RESULTS') then
    begin
       Trace.TraceInformation('BOB',TobResult.GetString('RESULTS'));
    end;
    if TobResult.FieldExists('MAJMENU')
      and TobResult.GetBoolean('MAJMENU') then
      Result := 1;
    if Count > 0 then // Je vais vider les caches ...
    begin
      // Purge des cache server
      AvertirCacheServer('*');
         // Purge des cache client : sont dans {TEMP}\PGI\{Nom de section de la DB0}
      RemoveInDir(TCBPPath.GetCegidUserTempPath + 'PGI\' + V_PGI.CurrentAlias, False, False);
      if Showmessage then
         PgiInfo('Une mise à jour a été passée sur votre environnement.'+#10#13+'Veuillez relancer le produit pour la prendre en compte.');
      // Application.Terminate ;
    end;
    // Libération mémoire
    TobResult.Free;
  end
  else
  begin
    if Showmessage then
      if PGIAsk('Impossible d''intégrer les BOBs. Problème lors de l''appel au processus serveur !' + #13#10 + 'Souhaitez vous continuer ?', 'Erreur') = mrNo then
         Halt;
  end;
  if Assigned(TobParam) then
    TobParam.Free;
  SourisNormale;
{$ENDIF ENTREPRISE}
end;

function IntegreBOBOnDemand(BobName: string): Integer;
{$IFNDEF ENTREPRISE}
var
  TobResult, TobParam: Tob;
{$ENDIF ENTREPRISE}
begin
  Result := -3;
{$IFNDEF ENTREPRISE}
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  TobResult := MyIntegreBob.INTEGREMANU(BobName);
  //TobResult := LanceProcessServer('cgiIntegreBOB', 'INTEGREMANU', BobName, TobParam, True);

  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    if TobResult.FieldExists('COUNT') then
      Result := 0;
    if V_PGI.SAV
      and TobResult.FieldExists('RESULTS') then
      PGIInfo(TobResult.GetString('RESULTS'));
    if TobResult.FieldExists('RESULTS') then
       Trace.TraceInformation('BOBOD',TobResult.GetString('RESULTS'));
    if TobResult.FieldExists('MAJMENU')
      and TobResult.GetBoolean('MAJMENU') then
      Result := 1;

    // Libération mémoire
    TobResult.Free;
  end;
  if Assigned(TobParam) then
    TobParam.Free;
{$ENDIF ENTREPRISE}
end;

{$ENDIF}


function BOB_IMPORT_PCL_STD_MANU_EXISTS(FiLeName: string): boolean;
{$IFNDEF ENTREPRISE}
{$IFDEF EAGLCLIENT}
var
  TobResult, TobParam: Tob;
{$ENDIF}
{$ENDIF ENTREPRISE}
begin
  Result := false;  //FP 07/05/2010 Suppression des avertissements
{$IFNDEF ENTREPRISE}
  Filename := ExtractFileName(FiLeName);
{$IFNDEF EAGLCLIENT}
  Result := FileExists(BOB_GET_CHEMIN + 'DDE\' + FiLeName);
{$ELSE}
  //Result := false;
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  TobResult := MyIntegreBob.BOBMANUEXISTS(FiLeName);
//  TobResult := LanceProcessServer('cgiIntegreBOB', 'BOBMANUEXISTS', FiLeName, TobParam, True);

  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    if TobResult.FieldExists('EXISTS') then
       Result := TobResult.GetBoolean('EXISTS');
    // Libération mémoire
    TobResult.Free;
  end;
  if Assigned(TobParam) then
    TobParam.Free;
{$ENDIF}
{$ENDIF ENTREPRISE}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : P. BASSET
Créé le ...... : ??/??/????
Modifié le ... : 10/02/2006
Description .. : IMPORTATION DES BOBS
Suite ........ : MODIF PB
Suite ........ : le case : case TestAGLIntegreBob(Chemin + sFileBOB) of
Suite ........ : devient
Suite ........ : case AGLIntegreBob(Chemin + sFileBOB) of
Suite ........ : PB le 10/02/06
Suite ........ : ex fonction PCL_IMPORT_BOB
Suite ........ :
Mots clefs ... :
*****************************************************************}

function BOB_IMPORT_PCL(CodeProduit: string; TypeBOZ: boolean): Integer;
begin
  Result := BOB_IMPORT_PCL_STD(CodeProduit, 'BOB');
end;

function BOB_IMPORT_PCL2(CodeProduit: string; sLibelle: string; TypeBOZ: boolean = False; NumBasePourMajForce: string = ''): Integer; //LMO20070102
begin
  Result := BOB_IMPORT_PCL_STD(CodeProduit, 'BOB');
end;


{***********A.G.L.***********************************************
Auteur  ...... : P. BASSET
Créé le ...... : 10/02/2006
Modifié le ... :   /  /
Description .. : ex fonction BOB_GET_CHEMIN
Suite ........ : retourne le chemin des bobs
Mots clefs ... :
*****************************************************************}

function BOB_GET_CHEMIN: string;
begin
  //PB LE 27/07/2007
  //PB LE 02/10/2007
  // MB LE 22/11 Ne pas utiliser en CWAS !
  // Utiliser plutot BOB_IMPORT_PCL_STD_MANU_EXISTS pour savoir, sur toute plateforme si une bob existe.
{$IFDEF eAGLClient}
  result := ''; // GetWindowsDrive + 'CWS\BOB\';
{$ELSE}
  Result := IncludeTrailingPathDelimiter(TCBPPath.GetCegidDistriBob);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : P.basset
Créé le ...... : 05/01/2005
Modifié le ... : 02/05/2006
Description .. : suppression des bobs qui sont
Suite ........ : différents de code produit et
Suite ........ : numéro de version
Suite ........ : el clair on ne garde que codeproduit+NumVersion
Suite ........ : ex fonction : BOB_PCL_GARDE
Suite ........ : PB:140406 on supprime également les _BOB
Suite ........ : PB:020506 on cherche les "*"
Mots clefs ... :
*****************************************************************}

procedure BOB_PCL_GARDE(CodeProduit, NumVersionDuBob: string);
var
  sFileBOB, Chemin: string;
  SearchRec: TSearchRec;
  ret: integer;
  position: integer;
begin
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version BOB 999
  // - extension .BOB
  // - exemple CCS50582F0001.BOB
  Chemin := BOB_GET_CHEMIN + CodeProduit + '\'; //EX : C:\PGI00\BOB\2035\
  ret := FindFirst(Chemin + '*' + CodeProduit + '*.BO*', faAnyFile, SearchRec);
  while ret = 0 do
  begin
    sFileBOB := SearchRec.Name; //RECUPERE NOM DU BOB  {25610900A001.BOB}
    if Copy(sFileBob, 1, 1) = '_' then
      position := 6
    else
      position := 5;
    if Copy(sFileBOB, Position, 4) <> NumVersionDuBob then
      DeleteFile(PChar(Chemin + '\' + SFileBOB));
    ret := FindNext(SearchRec);
  end;
  Sysutils.FindClose(SearchRec);
end;

{***********A.G.L.***********************************************
Auteur  ...... : PASCAL BASSET
Créé le ...... : 26/05/2005
Modifié le ... : 10/02/2006
Description .. : Fonction incompatibe avec celle de galoutil
Suite ........ : ex fonction BOB_RENAME_BOB
Mots clefs ... :
*****************************************************************}

procedure BOB_RENAME_BOB(CodeProduit: string; TypeBOZ: boolean);
var sFile: string;
  Chemin: string;
  SearchRec: TSearchRec;
  ret: integer;
begin
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version BOB 999
  // - extension .BOB
  // - exemple CCS50582F0001.BOB
  Chemin := BOB_GET_CHEMIN + CodeProduit + '\'; //EX : C:\PGI00\BOB\2035\
  if TypeBOZ then
    ret := FindFirst(Chemin + CodeProduit + '*.BOZ', faAnyFile, SearchRec)
  else
    ret := FindFirst(Chemin + CodeProduit + '*.BOB', faAnyFile, SearchRec);
  while ret = 0 do
  begin
    //RECUPERE NOM DU BOB
    sFile := SearchRec.Name;
    if Copy(sFile, 9, 1) = 'D' then
    begin
      if FileExists(PChar(Chemin + '\' + '_' + SFile)) then DeleteFile(PChar(Chemin + '\' + '_' + SFile));
      RenameFile(PChar(Chemin + '\' + SFile), PChar(Chemin + '\' + '_' + SFile));
    end;
    if Copy(sFile, 9, 1) = 'F' then
    begin
      if FileExists(PChar(Chemin + '\' + '_' + SFile)) then DeleteFile(PChar(Chemin + '\' + '_' + SFile));
      RenameFile(PChar(Chemin + '\' + SFile), PChar(Chemin + '\' + '_' + SFile));
    end;
    ret := FindNext(SearchRec);
  end;
  sysutils.FindClose(SearchRec);
end;

{***********A.G.L.***********************************************
Auteur  ...... : PASCAL BASSET
Créé le ...... : 10/02/2006
Modifié le ... :   /  /
Description .. : supprime les enregistrements dans ymybob
Mots clefs ... :
*****************************************************************}

procedure BOB_DemandeSuppression(sBob: string);
var QuelBOB: string;
begin
  if PGIASK('Voulez-vous supprimer des BOBS de type ' + sBob + ' ?') = mryes then
  begin
    QuelBOB := '0' + IntToStr(V_PGI.NumVersionBase + 1);
    QuelBOB := InputBox('N° de BOB', '', QuelBOB);
    if PGIASK('Suppression des bobs ' + sBob + QuelBOB + '%') = mryes then
      if PGIASK('Confirmez de nouveau la suppression des bobs ' + sBob + QuelBOB + '%') = mryes then
        ExecuteSQL('DELETE FROM YMYBOBS WHERE YB_BOBNAME LIKE "' + sBob + QuelBOB + '%"');
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 20/03/2006
Modifié le ... : 16/05/2006
Description .. : Regarde le nombre de bobs ou le poids des fichiers pour
Suite ........ : alerter l'utilisateur que le parametrage va etre mis en place
Suite ........ : et risque de prendre un peu de temps.
Suite ........ :
Suite ........ : si poids renseigne, c'est le poids des fichiers qui compte,
Suite ........ : sinon, c'est leur nombre.
Suite ........ : PB:le 16/05/06, passe en fonction, retourne le nombre de
Suite ........ : bob
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT}

function BOB_Alerte(CodeProduit: string; NomProduit: string = ''; nb: integer = 3; poids: integer = 0; GestionBOBCommun: Boolean = true): integer;
var
  Chemin: string;
  SearchRec: TSearchRec;
  ret, i, p: integer;
  parle: boolean;
  BobList: TStringList;
begin
//  result := 0;
  i := 0;
  p := 0;
  Chemin := BOB_GET_CHEMIN + CodeProduit + '\'; //EX : C:\PGI00\BOB\2035\
  ret := FindFirst(Chemin + CodeProduit + '*.BO*', faAnyFile, SearchRec);
  BobList := ChargeListeBOB(GestionBOBCommun);
  while ret = 0 do
  begin
    if not BobIsAlreadyIntegred(BOBList, SearchRec.Name, GestionBOBCommun) then // BOB non intégrée
    begin
      Inc(i);
      Inc(p, SearchRec.Size);
    end;
    ret := FindNext(SearchRec);
  end;
  Sysutils.FindClose(SearchRec);


  if poids <> 0 then
    parle := (p / 1000000) > poids
  else
    parle := i > nb;


  if parle then
    PGIInfo('Mise en place du paramétrage ' + NomProduit + '.' + #10 + #13
      + ' Cette opération va prendre quelques minutes.' + #10 + #13
      + ' Veuillez cliquer sur "Ok" pour continuer.');
  result := i;
end;

{$ELSE}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : RPO
Créé le ...... : 13/02/2008
Modifié le ... :   /  /
Description .. : Permet d'exécuter un script SQL figurant dans le
Suite ........ : répertoire \DAT\SOCREF du serveur de production
Suite ........ : (serveur SQL).
Suite ........ : Le log demandé sera généré au même endroit.
Mots clefs ... :
*****************************************************************}
function CWAS_EXECSCRIPTSQL(nomscript, nomlog, nombase: String; FP : TFPatience; bPurgeDesCaches: Boolean=False): boolean;
{$IFNDEF ENTREPRISE}
var
  TobResult, TobParam: Tob;
{$ENDIF ENTREPRISE}
begin
  Result := False;
{$IFNDEF ENTREPRISE}
  if (nomscript='') or (nomlog='') then exit;
  SourisSablier;
  TobParam := TOB.Create('le param', nil, -1);
  if not FP.Visible then
     FP.Show ;
  TobResult := MyIntegreBob.EXECSCRIPTSQL(nomscript+';'+nomlog+';'+nombase);
  // LanceProcessServer('cgiIntegreBOB', 'EXECSCRIPTSQL', nomscript+';'+nomlog+';'+nombase, TobParam, True);
  if FP.Visible then
     FP.Hide ;
  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    if V_PGI.SAV
      and TobResult.FieldExists('RESULTS') then
      PGIInfo(TobResult.GetString('RESULTS'));
    if TobResult.FieldExists('RESULTS') then
    begin
       Trace.TraceInformation('EXECSCRIPTSQL',TobResult.GetString('RESULTS'));
       Result := True;
    end;
    if bPurgeDesCaches then
    begin
      // Purge des cache server
      AvertirCacheServer('*');
      // Purge des cache client : sont dans {TEMP}\PGI\{Nom de section de la DB0}
      RemoveInDir(TCBPPath.GetCegidUserTempPath + 'PGI\' + V_PGI.CurrentAlias, False, False);
      // MD 11/06/08 - seul un script lancé sur la base en cours nécessite un redémarrage de l'application
      //if V_PGI.DbName = nombase then
      //  PgiInfo('Une mise à jour a été passée sur votre environnement.'+#10#13+'Veuillez relancer le produit pour la prendre en compte.');
      // Application.Terminate ;
    end;
    // Libération mémoire
    TobResult.Free;
  end
  else
  begin
    PGIInfo('Impossible d''exécuter le script '+nomscript+'. Problème lors de l''appel au processus serveur !');
  end;
  if Assigned(TobParam) then
    TobParam.Free;
  SourisNormale;
{$ENDIF ENTREPRISE}
end;
{$ENDIF EAGLCLIENT}



function BOB_IMPORT_PCL_STD_MANU(FiLeName: string): Integer;
begin
{$IFNDEF EAGLCLIENT}
  Result := AglIntegreBob(FiLeName, FALSE, FALSE);
{$ELSE}
  Result := IntegreBOBOnDemand(ExtractFileName(FiLeName));
{$ENDIF}
end;

{$IFNDEF EAGLCLIENT}

// MB : Même source que le CGI.

function _2TIERS_BOB_IMPORT_PCL_STD(BobList : TstringList ; FP : TFPatience ; CodesProduits: TStringList ; Libelles : TStringList ; GestionBOBCommun: Boolean = False): Integer;
var
  Chemin: string;
  ResultTob : Tob ;
  i{, Count}: integer;
begin
  Result := 0;
  try
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version 999
  // - extension .BOB
  // - exemple CCS50582F001.BOB

  // CODE RETOUR DE LA FONCTION
  Chemin := BOB_GET_CHEMIN ;

  ResultTob := TOB.Create('le param', nil, -1);

  if GetCurDb = GetDefaultDb then // bob dans la racine
  begin
     FP.SetTitre('Mises à jour Spécifiques') ;
     TraiteDir(BobList, Chemin,ResultTob, GestionBOBCommun, FP.lAide );
  end;
  for i:=0 to CodesProduits.Count-1 do
  begin
     FP.SetTitre('Mises à jour ' + Libelles.Strings[i]) ;
     TraiteDir(BobList, Chemin + CodesProduits.Strings[i] + '\' ,ResultTob, GestionBOBCommun, FP.lAide);
  end;


  if FP.Visible then
     FP.Hide ;

  //Count := 0;
  // TobResult bien renseignée ?
  if ResultTob.FieldExists('RESULT') then
  begin
    //Count := ResultTob.GetInteger('RESULT');
    // #### Dommage : on n'exploite pas ce count !
    Result := 0;
  end;
  if V_PGI.SAV
    and ResultTob.FieldExists('RESULTS') then
    PGIInfo(ResultTob.GetString('RESULTS'));
  if ResultTob.FieldExists('RESULTS') then
    Trace.TraceInformation('BOB',ResultTob.GetString('RESULTS'));

  if ResultTob.FieldExists('MAJMENU')
    and ResultTob.GetBoolean('MAJMENU') then
    Result := 1;

  // Libération mémoire
  FreeAndNil(ResultTob) ;

  except
    on e: exception do
       Showmessage('erreur : '+e.Message);
  end;

end;


{$ENDIF}

{$IFDEF EAGLCLIENT}

function _CWAS_BOB_IMPORT_PCL_STD(CodeProduit: string; GestionBOBCommun: Boolean ; FP : TFPatience ; ShowWaitForm : Boolean = True; Showmessage : Boolean = true ): Integer;
begin
  Result := IntegreBOBAuto(CodeProduit, '', GestionBOBCommun,FP,ShowWaitForm, Showmessage);
end;

function _CWAS_BOB_IMPORT_PCL_DB(CodeProduit: string; DBName : String ; GestionBOBCommun: Boolean ; FP : TFPatience ; ShowWaitForm : Boolean = True; Showmessage : Boolean = true ): Integer;
begin
  Result := IntegreBOBAuto(CodeProduit,DBName, GestionBOBCommun,FP,ShowWaitForm, Showmessage);
end;

{$ENDIF}

function BOB_IMPORT_PCL_DB(CodesProduits: string; sLibelle: string; DBName : String ; GestionBOBCommun: Boolean = False; ShowWaitForm : Boolean = True; Showmessage : boolean = true ; MessageSupplementaire : String = '' ): Integer;
var
  FP : TFPatience ;
{$IFNDEF EAGLCLIENT}
  BobList: TStringList;
  LCode: TstringList;
  LLibelle: TstringList;
  i : integer;
{$ENDIF}
begin
  FP := FenetrePatience('Gestionnaire des BOBs',aoCentreBas, False,ShowWaitForm);
  if ShowWaitForm then
  begin
    if MessageSupplementaire <> '' then
      FP.lCreation.Caption := MessageSupplementaire
    else
      FP.lcreation.visible := false ;
    FP.StartK2000 ;
  end;

{$IFNDEF EAGLCLIENT}

  BobList := ChargeListeBOB(GestionBOBCommun);

  LCode := TstringList.Create;
  LCode.Delimiter := ';';
  LCode.DelimitedText := CodesProduits;

  LLibelle := TstringList.Create;
  LLibelle.Delimiter := ';';
  LLibelle.DelimitedText := sLibelle;
  Result := 0 ;
  try
     ExecuteSQL('@@USE '+DBName) ;
  except
     Result := -1 ;
  end;
  if Result = 0 then
  begin
     V_PGI.DBName := DBName;
     V_PGI.Nodossier := copy(DBName,3,8);
     Result := _2TIERS_BOB_IMPORT_PCL_STD(BobList,FP,LCode,LLibelle, GestionBOBCommun);
  end;
  try
     ExecuteSQL('@@USE '+GetDefaultDb) ;
  except
  end;

{  for i := 0 to LCode.Count - 1 do
  begin
    _2TIERS_BOB_IMPORT_PCL_STD(BobList,FP.lAide,LCode.Strings[i], GestionBOBCommun);
  end;}

  LCode.Free;
  LLibelle.Free;

  if BobList <> nil then
  begin
    for i:=0 to BobList.Count-1 do
    begin
      if BobList.Objects[i] <> nil then
        BobList.Objects[i].Free ;
    end;
    BobList.free;
  end;
{$ELSE}
  FP.Affiche('Vérification de la présence de mises à jour et installation si nécessaire.') ;
  Result := _CWAS_BOB_IMPORT_PCL_DB(CodesProduits, DBName, GestionBOBCommun,FP,ShowWaitForm,Showmessage);
{$ENDIF}

  if ShowWaitForm then
     FP.StopK2000 ;
  FP.Free ;
end;



function BOB_IMPORT_PCL_STD(CodesProduits: string; sLibelle: string; GestionBOBCommun: Boolean = False; ShowWaitForm : Boolean = True; Showmessage : boolean = true  ): Integer;
var
  FP : TFPatience ;
{$IFNDEF EAGLCLIENT}
  BobList: TStringList;
  LCode: TstringList;
  LLibelle: TstringList;
  i : integer;
{$ENDIF}
begin
  FP := FenetrePatience('Gestionnaire des BOBs',aoCentreBas, False,ShowWaitForm);
  if ShowWaitForm then
  begin
    FP.lcreation.visible := false ;
    FP.StartK2000 ;
  end;

{$IFNDEF EAGLCLIENT}

  BobList := ChargeListeBOB(GestionBOBCommun);

  LCode := TstringList.Create;
  LCode.Delimiter := ';';
  LCode.DelimitedText := CodesProduits;

  LLibelle := TstringList.Create;
  LLibelle.Delimiter := ';';
  LLibelle.DelimitedText := sLibelle;

  Result := _2TIERS_BOB_IMPORT_PCL_STD(BobList,FP,LCode,LLibelle, GestionBOBCommun);

{  for i := 0 to LCode.Count - 1 do
  begin
    _2TIERS_BOB_IMPORT_PCL_STD(BobList,FP.lAide,LCode.Strings[i], GestionBOBCommun);
  end;}

  LCode.Free;
  LLibelle.Free;

  if BobList <> nil then
  begin
    for i:=0 to BobList.Count-1 do
    begin
      if BobList.Objects[i] <> nil then
        BobList.Objects[i].Free ;
    end;
    BobList.free;
  end;
{$ELSE}
  FP.Affiche('Vérification de la présence de mises à jour et installation si nécessaire.') ;
  Result := _CWAS_BOB_IMPORT_PCL_STD(CodesProduits, GestionBOBCommun,FP,ShowWaitForm,Showmessage);
{$ENDIF}

  if ShowWaitForm then
     FP.StopK2000 ;
  FP.Free ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : G HARLEZ
Créé le ...... : 07/11/2008
Modifié le ... :   /  /
Description .. : FQ 12099. Màj des menus dans la base standard.
Suite ........ :
Mots clefs ... : STANDARD;MENU;BOB
*****************************************************************}
procedure BOB_IMPORT_STD_MENU(ShowWaitForm : Boolean = True);
var
  FP : TFPatience ;
{$IFNDEF ENTREPRISE}
  TobResult : TOB;
  {$IFNDEF EAGLCLIENT}
  NameBob,PathBob, DB0 : string;
  i : integer;
  TBob,TbP : TOB;
  PathDirBob : string;
  {$ELSE}
  TobParam : TOB ;
  {$ENDIF}
{$ENDIF ENTREPRISE}
begin
  // Création d'une fenêtre de patience
  FP := FenetrePatience('Gestionnaire des BOBs',aoCentreBas, False, ShowWaitForm);
  if ShowWaitForm then
  begin
    FP.lcreation.visible := false ;
    FP.StartK2000 ;
  end;
  FP.Affiche('Intégration des BOBs Menu sur la base des standards') ;
{$IFNDEF ENTREPRISE}
  {$IFDEF EAGLCLIENT}
  // en CWAS, on passe par l'utilitaire cgiIntegreBOB (..CWS\SCRIPTS), On recherche les BOBs
  // à partir de l'emplacement GetCWSInstallDir (retourne C:\CWS\BOB)
  TobParam := TOB.Create('Le param',nil,-1);
  try
    TobResult := MyIntegreBob.CWSBOBMENU ;
    // TobResult := LanceProcessServer('cgiIntegreBOB', 'CWSBOBMENU', '', TobParam, True);

    // TobResult bien renseignée ?
    if Assigned(TobResult) then
    begin
      if V_PGI.SAV
        and TobResult.FieldExists('RESULTS') then
        PGIInfo(TobResult.GetString('RESULTS'));
      if TobResult.FieldExists('RESULTS') then
         Trace.TraceInformation('BOBMENU',TobResult.GetString('RESULTS'));

      // Libération mémoire
      TobResult.Free;
    end;
  finally
    if Assigned(TobParam) then
      TobParam.Free;
  end;
  {$ELSE}
  //////////// PHASE 1 : Recherche des fichiers BOBs de type Menu ////////////
  // en 2/3, on recherche les BOBs à partir de l'emplacement  TcbpPath.GetCegidDistriBob
  // (retourne C:\Documents and Settings\All Users\Application Data\CEGID\Cegid Expert\BOB)
  PathDirBob := IncludeTrailingPathDelimiter(TcbpPath.GetCegidDistriBob);
  TobResult := GetListeBobMenu(PathDirBob);

  // Interrompt le traitement si aucun menu n'est à intégrer.
  if (TobResult.FieldExists('RESULT')) and (TobResult.GetString('RESULT') = 'NOK') then
    exit;

  //////////// PHASE 2 : Lecture de la table YMYBOBS ////////////
  TBob := TOB.Create('YMYBOBS_STD',nil,-1);
  try
    // Chargement des enreg de la table. L'existence de la base DB000STD est contrôlée en amont.
    TBob.LoadDetailFromSQL('@@SELECT YB_BOBNAME FROM DB000STD.DBO.YMYBOBS');

    // Init de la base DB0
    DB0 := V_PGI.DefaultSectionDBName;
    if DB0 = '' then
      DB0 := V_PGI.DbName;

    // On bascule sur le dossier STD.
    ExecuteSQL('@@USE DB000STD');

    // Lecture de la tob TobResult
    for i:=0 to TobResult.Detail.Count-1 do
    begin
      NameBob := TobResult.Detail[i].GetString('BOBNAME');
      PathBob := TobResult.Detail[i].GetString('BOBPATH');
      // Recherche de la BOB dans la Tob
      TbP := TBob.FindFirst(['YB_BOBNAME'],[NameBob],FALSE);
      // Si le bob n'est pas référencé dans la table, on le traite.
      if not assigned(TbP) then
      begin
        FP.Affiche('Intégration du BOB Menu '+NameBob) ;
        AglIntegreBob(PathBob,FALSE,FALSE);
      end;

      Application.ProcessMessages;
    end;
  finally
    // Ne pas oublier de revenir sur la DB0
    ExecuteSQL('@@USE '+DB0);
    TBob.Free;
    TobResult.Free;
  end;
  {$ENDIF}
{$ENDIF ENTREPRISE}

  if ShowWaitForm then
    FP.StopK2000 ;
  FP.Free ;
end;

 {  ZoomSuivi: TForm;
  Timer1 : TTimer ;
  TH: THLABEL;}

  {ZoomSuivi := TForm.Create(nil);
  ZoomSuivi.Caption := 'Mise à jour des BOBs' ;
  ZoomSuivi.BorderIcons := [biSystemMenu, biMinimize, biMaximize];
  ZoomSuivi.BorderStyle := bsSizeable;
  ZoomSuivi.Height := 70;
  ZoomSuivi.Width := 400;
  ZoomSuivi.Position := poScreenCenter;
  ZoomSuivi.BorderIcons := [];
  TH := THLABEL.Create(ZoomSuivi);
  TH.Parent := ZoomSuivi;
  TH.Caption := '';

  Timer1 := TTimer.Create(ZoomSuivi);
  Timer1.Enabled := False ;
  Timer1.OnTimer :=


  TH.Top := 10;
  TH.left := 10;
  ZoomSuivi.KeyPreview := TRUE;
  if ShowWaitForm then
  begin
     ZoomSuivi.Show;
     ZoomSuivi.FormStyle := fsStayOnTop;
  end;
  }


end.

