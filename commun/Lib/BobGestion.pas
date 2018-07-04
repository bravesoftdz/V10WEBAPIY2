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
  uLanceProcess,
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

function BOB_IMPORT_PCL_STD(CodesProduits: string; sLibelle: string; GestionBOBCommun: Boolean = False; ShowWaitForm : Boolean = True ): Integer;
function BOB_IMPORT_PCL_STD_MANU(FiLeName: string): Integer;

function BOB_IMPORT_PCL_STD_MANU_EXISTS(FiLeName: string): boolean;
// function GetWindowsDrive: string; //PB LE 02/10/2007 : fonction temporaire car non portée en WA

//function BOB_IMPORT_PCL_STD(CodeProduit: string; sLibelle: string; RenameAfter : Boolean=False): Integer;
function BOB_GET_CHEMIN: string;

// MB : A porter demain en CWAS ....

{$IFNDEF EAGLCLIENT}
// pourrait marcher dans les deux modes ...
procedure BOB_DemandeSuppression(sBob: string);
// Marche qu'en 2/3
procedure BOB_PCL_GARDE(CodeProduit, NumVersionDuBob: string);
procedure BOB_RENAME_BOB(CodeProduit: string; TypeBOZ: boolean = false);
function BOB_Alerte(CodeProduit: string; NomProduit: string = ''; nb: integer = 3; poids: integer = 0; GestionBOBCommun: Boolean = true): integer;
{$ELSE}
procedure CWAS_EXECSCRIPTSQL(nomscript, nomlog: String; FP : TFPatience; bPurgeDesCaches: Boolean=False);
{$ENDIF}
procedure AfficheBox(resultats : String);


////////////////// IMPLEMENTATION ///////////////////
implementation
uses
{$IFDEF EAGLCLIENT}

{$ELSE}
  db,
{$IFNDEF DBXPRESS}
  dbtables,
{$ELSE}
  uDbxDataSet,
{$ENDIF}
{$ENDIF EAGLSERVER}
	UtilsRapport,
  uTob;


procedure AfficheBox(resultats : String);
var Rapport : TGestionRapport;
  	HH : string;
    MM : string;
begin
	Rapport := TGestionRapport.create (Application.MainForm);
  Rapport.InitRapport;
  Rapport.Close := True;
  Rapport.Titre := 'Resultat intégration des BOBS';
  Rapport.Sauve := false;
  Rapport.Affiche := True;
  HH := resultats;
  repeat
		MM := READTOKENST(HH);
    if MM <> '' then
    begin
      Rapport.SauveLigMemo (MM);
    end;
  until MM = '';
  Rapport.AfficheRapport;
end;

{$IFDEF eAGLClient}

function IntegreBOBAuto(CodesProduits: string; GestionBOBCommun: Boolean ; FP : TFPatience): Integer;
var
  TobResult, TobParam: Tob;
  GesCom: string;
  Count: Integer;
begin
  SourisSablier;
  Result := -3;
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  GesCom := '-';
  if not FP.Visible then
     FP.Show ;
  if GestionBOBCommun then GesCom := 'X';
  TobResult := LanceProcessServer('cgiIntegreBOB', 'INTEGREAUTO', CodesProduits + ';' + GesCom, TobParam, True);

  if FP.Visible then
     FP.Hide ;

  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    Count := 0;
    if TobResult.FieldExists('RESULT') then
    begin
      Count := TobResult.GetInteger('RESULT');
      Result := 0;
    end;
    if V_PGI.SAV
      and TobResult.FieldExists('RESULTS') then
    begin
      PGIBox(TobResult.GetString('RESULTS'));
    end;
    if TobResult.FieldExists('RESULTS') then
       Trace.TraceInformation('BOB',TobResult.GetString('RESULTS'));
    if TobResult.FieldExists('MAJMENU')
      and TobResult.GetBoolean('MAJMENU') then
      Result := 1;
    if Count > 0 then // Je vais vider les caches ...
    begin
      // Purge des cache server
      AvertirCacheServer('*');
         // Purge des cache client : sont dans {TEMP}\PGI\{Nom de section de la DB0}
      RemoveInDir(TCBPPath.GetCegidUserTempPath + 'PGI\' + V_PGI.CurrentAlias, False, False);
      PgiBox('Une mise à jour a été passée sur votre environnement.'+#10#13+'Veuillez relancer le produit pour la prendre en compte.');
      // Application.Terminate ;
    end;
    // Libération mémoire
    TobResult.Free;
  end
  else
  begin
    if PGIAsk('Impossible d''intégrer les BOBs. Problème lors de l''appel au processus serveur !' + #13#10 + 'Souhaitez vous continuer ?', 'Erreur') = mrNo then
      Halt;
  end;
  if Assigned(TobParam) then
    TobParam.Free;
  SourisNormale;
end;

function IntegreBOBOnDemand(BobName: string): Integer;
var
  TobResult, TobParam: Tob;
begin
  Result := -3;
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  TobResult := LanceProcessServer('cgiIntegreBOB', 'INTEGREMANU', BobName, TobParam, True);

  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    if TobResult.FieldExists('RESULT') then
      Result := 0;
    if V_PGI.SAV
      and TobResult.FieldExists('RESULTS') then
      PGIBox(TobResult.GetString('RESULTS'));
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
end;

{$ENDIF}


function BOB_IMPORT_PCL_STD_MANU_EXISTS(FiLeName: string): boolean;
{$IFDEF EAGLCLIENT}
var
  TobResult, TobParam: Tob;
{$ENDIF}
begin
  Filename := ExtractFileName(FiLeName);
{$IFNDEF EAGLCLIENT}
  Result := FileExists(BOB_GET_CHEMIN + 'DDE\' + FiLeName);
{$ELSE}
  Result := false;
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  TobResult := LanceProcessServer('cgiIntegreBOB', 'BOBMANUEXISTS', FiLeName, TobParam, True);

  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    if TobResult.FieldExists('RESULT') then
       Result := TobResult.GetBoolean('RESULT');
    // Libération mémoire
    TobResult.Free;
  end;
  if Assigned(TobParam) then
    TobParam.Free;
{$ENDIF}
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
  if PGIASK('Voulez vous supprimer des BOBS de type ' + sBob + ' ?') = mryes then
  begin
    QuelBOB := '0' + IntToStr(V_PGI.NumVersionBase + 1);
    QuelBOB := InputBox('N° de BOB', '', QuelBOB);
    if PGIASK('Supression des bobs ' + sBob + QuelBOB + '%') = mryes then
      if PGIASK('Confirmez de nouveau la Supression des bobs ' + sBob + QuelBOB + '%') = mryes then
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
    PGIInfo('Mise en place du paramètrage ' + NomProduit + '.' + #10 + #13
      + ' Cette opération va prendre quelques minutes.' + #10 + #13
      + ' Veuillez cliquer sur "Ok" pour continuer.');
  result := i;
end;

{$ELSE}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : RPO
Créé le ...... : 13/02/2008
Modifié le ... :   /  /
Description .. : Permet d'éxécuter un script SQL figurant dans le
Suite ........ : répertoire \DAT\SOCREF du serveur de production
Suite ........ : (serveur SQL).
Suite ........ : Le log demandé sera généré au même endroit.
Mots clefs ... :
*****************************************************************}
procedure CWAS_EXECSCRIPTSQL(nomscript, nomlog: String; FP : TFPatience; bPurgeDesCaches: Boolean=False);
var
  TobResult, TobParam: Tob;
begin
  if (nomscript='') or (nomlog='') then exit;
  SourisSablier;
  TobParam := TOB.Create('le param', nil, -1);
  if not FP.Visible then
     FP.Show ;
  TobResult := LanceProcessServer('cgiIntegreBOB', 'EXECSCRIPTSQL', nomscript+';'+nomlog, TobParam, True);
  if FP.Visible then
     FP.Hide ;
  // TobResult bien renseignée ?
  if Assigned(TobResult) then
  begin
    if V_PGI.SAV
      and TobResult.FieldExists('RESULTS') then
      PGIBox(TobResult.GetString('RESULTS'));
    if TobResult.FieldExists('RESULTS') then
       Trace.TraceInformation('EXECSCRIPTSQL',TobResult.GetString('RESULTS'));
    if bPurgeDesCaches then
    begin
      // Purge des cache server
      AvertirCacheServer('*');
      // Purge des cache client : sont dans {TEMP}\PGI\{Nom de section de la DB0}
      RemoveInDir(TCBPPath.GetCegidUserTempPath + 'PGI\' + V_PGI.CurrentAlias, False, False);
      PgiBox('Une mise à jour a été passée sur votre environnement.'+#10#13+'Veuillez relancer le produit pour la prendre en compte.');
      // Application.Terminate ;
    end;
    // Libération mémoire
    TobResult.Free;
  end
  else
  begin
    PGIInfo('Impossible d''éxécuter le script '+nomscript+'. Problème lors de l''appel au processus serveur !');
  end;
  if Assigned(TobParam) then
    TobParam.Free;
  SourisNormale;
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
  i, Count: integer;
begin
  try
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version 999
  // - extension .BOB
  // - exemple CCS50582F001.BOB

  // CODE RETOUR DE LA FONCTION
  Result := 0;
  Chemin := BOB_GET_CHEMIN ;

  ResultTob := TOB.Create('le param', nil, -1);

  if V_PGI.NoDossier = '000000' then // bob dans la racine
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

  Count := 0;
  // TobResult bien renseignée ?
  if ResultTob.FieldExists('RESULT') then
  begin
    Count := ResultTob.GetInteger('RESULT');
    // #### Dommage : on n'exploite pas ce count !
    if count > 0 then result := 1 else Result := 0;
  end;
  if V_PGI.SAV
    and ResultTob.FieldExists('RESULTS') then
    	AfficheBox (ResultTob.GetString('RESULTS'));

//    PgiBox(ResultTob.GetString('RESULTS'));
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

function _CWAS_BOB_IMPORT_PCL_STD(CodeProduit: string; GestionBOBCommun: Boolean ; FP : TFPatience ): Integer;
begin
  Result := IntegreBOBAuto(CodeProduit, GestionBOBCommun,FP);
end;
{$ENDIF}


function BOB_IMPORT_PCL_STD(CodesProduits: string; sLibelle: string; GestionBOBCommun: Boolean = False; ShowWaitForm : Boolean = True ): Integer;
var
  BobList: TStringList;
  FP : TFPatience ;
{$IFNDEF EAGLCLIENT}
  LCode: TstringList;
  LLibelle: TstringList;
{$ENDIF}
begin
  FP := FenetrePatience(sLibelle,aoCentreBas, False,ShowWaitForm);
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
  BobList.Free ;

{$ELSE}
  FP.Affiche('Vérification de la présence de mises à jour et installation si nécessaire.') ;
  Result := _CWAS_BOB_IMPORT_PCL_STD(CodesProduits, GestionBOBCommun,FP);
{$ENDIF}

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

