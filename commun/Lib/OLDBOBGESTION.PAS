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
  extctrls;

function BOB_IMPORT_PCL(CodeProduit: string; TypeBOZ: boolean = False): Integer;
function BOB_IMPORT_PCL2(CodeProduit: string; sLibelle: string; TypeBOZ: boolean = False; NumBasePourMajForce: string = ''): Integer;

function BOB_IMPORT_PCL_STD(CodeProduit: string; sLibelle: string; NumBasePourMajForce: string = ''): Integer;
function BOB_IMPORT_PCL_STD_MANU(FiLeName : String ): Integer;

function  BOB_GET_CHEMIN: string;
procedure BOB_PCL_GARDE(CodeProduit, NumVersionDuBob: string);
procedure BOB_RENAME_BOB(CodeProduit: string; TypeBOZ: boolean = false);
procedure BOB_DemandeSuppression(sBob: string);
function  BOB_Alerte(CodeProduit: string; NomProduit: string = ''; nb: integer = 3; poids: integer = 0): integer;

function tmpIntegreBob(const FileBobName: string; CreationSeulement: boolean = False;
  SiTrouveJeSors: boolean = false; CallBack:
  TTImportObjectCallBack = nil; controleNoVersion: boolean = true): Integer; //LMO20070108

implementation
uses
{$IFDEF EAGLCLIENT}

{$ELSE}
  db,
{$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
{$ENDIF EAGLSERVER}
  uTob;

{$IFDEF eAGLClient}

Function IntegreBOBAuto (CodeProduit : String = ''; NumBasePourMajForce: string = '') : Integer ;
var
  TobResult,TobParam : Tob ;
begin
  Result := -3 ;
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  TobResult := LanceProcessServer('cgiIntegreBOB', 'INTEGREAUTO', V_PGI.Nodossier+';'+CodeProduit+';'+NumBasePourMajForce , TobParam, True ) ;

  // TobResult bien renseignée ?
  if Assigned(TobResult)
  and TobResult.FieldExists('RESULT') then
     Result := 0 ;
  if V_PGI.SAV
  and Assigned(TobResult)
  and TobResult.FieldExists('RESULTS') then
     PGIBox(TobResult.GetString('RESULTS'));
  if Assigned(TobResult)
  and TobResult.FieldExists('MAJMENU')
  and TobResult.GetBoolean('MAJMENU') then
     Result := 1 ;

  // Libération mémoire
  if Assigned(TobResult) then
    TobResult.Free ;
  if Assigned(TobParam) then
    TobParam.Free ;
end;

Function IntegreBOBOnDemand (BobName : String) : Integer ;
var
  TobResult,TobParam : Tob ;
begin
  Result := -3 ;
  TobParam := TOB.Create('le param', nil, -1);
  // Traitement
  TobResult := LanceProcessServer('cgiIntegreBOB', 'INTEGREMANU', BobName , TobParam, True ) ;

  // TobResult bien renseignée ?
  if Assigned(TobResult)
  and TobResult.FieldExists('RESULT') then
     Result := 0 ;
  if V_PGI.SAV
  and Assigned(TobResult)
  and TobResult.FieldExists('RESULTS') then
     PGIBox(TobResult.GetString('RESULTS'));
  if Assigned(TobResult)
  and TobResult.FieldExists('MAJMENU')
  and TobResult.GetBoolean('MAJMENU') then
     Result := 1 ;

  // Libération mémoire
  if Assigned(TobResult) then
    TobResult.Free ;
  if Assigned(TobParam) then
    TobParam.Free ;
end;

{$ENDIF}



{***********Fonction de choix d'intégration du bob***************
Auteur : M. BRUN
Le : 25/09/2007
}


function CanIntegreBOB(S : TSearchRec; Forced : Boolean) : boolean ;
var
  sext : string ;
begin
   Result := False ;
   if (S.attr and faDirectory) = faDirectory then exit ;
   if (copy(S.Name,1,1) = '_') and not Forced then exit ;
   sext := ExtractFileExt(S.Name) ;
   Result := SameText(sext,'.bob') or SameText(sext,'.boz')  ;
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
var sFileBOB: string;
  Chemin: string;
  SearchRec: TSearchRec;
  NumVersion: integer;
  ret: integer;
  iCase_l: integer;
begin
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version 999
  // - extension .BOB
  // - exemple CCS50582F001.BOB
  //
  // CODE RETOUR DE LA FONCTION
  Result := 0;
  NumVersion := 0;
  Chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\'; //EX: C:\PGI00\BOB\2035\
  if V_PGI.SAV then
  begin
    PGIINFO('dans BOB_IMPORT_PCL, répertoire : ' + Chemin);
    if typeboz
      then PGIINFO('typeboz à true')
    else PGIINFO('typeboz à false');
  end;

  if TypeBOZ then
    ret := FindFirst(Chemin + CodeProduit + '*.BOZ', faAnyFile, SearchRec)
  else
    ret := FindFirst(Chemin + CodeProduit + '*.BOB', faAnyFile, SearchRec);
//      Result := 0;
//  Chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\';  //EX : C:\PGI00\BOB\2035\
//  ret := FindFirst(Chemin+CodeProduit+'*.BOB', faAnyFile, SearchRec);
  while ret = 0 do
  begin
       //RECUPERE NOM DU BOB
    sFileBOB := SearchRec.Name;
    if V_PGI.SAV then PGIINFO('Analyse du BOB : ' + sFileBOB);
       //RECUPERE NUM VERSION
    try
      NumVersion := ValeurI(Copy(sFileBOB, 5, 4));
    except
      Result := -6;
      exit;
    end;

       // MD 08/11/04 if NumVersion > V_PGI.NumVersionBase then
       // => remplacé par NumVersionSoc sinon risque pour un vieil exe
       // de remonter des bobs périmés dans une base récente !!
    if NumVersion > V_PGI.NumVersionSoc then
    begin
      iCase_l := AglIntegreBob(Chemin + sFileBOB, FALSE, TRUE);
      case iCase_l of
//            case TestAGLIntegreBob(Chemin + sFileBOB) of
        0: // OK
          begin
            if V_PGI.SAV then Pgiinfo('Intégration de : ' + sFileBOB, TitreHalley); //Resultif not LIA_JOURNAL_EVENEMENT(sTempo) then Result := -1;
            if copy(sFileBob, 9, 1) = 'M' then Result := 1; //SI BOB AVEC MENU, ON REND 1 POUR SORTIR DE L'APPLICATION
          end;
        1: if V_PGI.SAV then Pgiinfo('Intégration déjà effectuée :' + sFileBOB, TitreHalley); // Intégration déjà effectuée
        -1: // Erreur d'écriture dans la table YMYBOBS
          begin
            if V_PGI.SAV then PGIInfo('Erreur d''écriture dans la table YMYBOBS :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
        -2: // Erreur d'intégration dans la fonction AglImportBob
          begin
            if V_PGI.SAV then PGIInfo('Erreur d''intégration dans la fonction AglImportBob :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
        -3: //Erreur de lecture du fichier BOB.
          begin
            if V_PGI.SAV then PGIInfo('Erreur de lecture du fichier BOB :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
        -4: // Erreur inconnue.
          begin
            if V_PGI.SAV then PGIInfo('Erreur inconnue :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
      else
        if V_PGI.SAV then PGIInfo('Erreur inconnue :[' + IntToStr(iCase_l) + '],' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');

      end;

    end;
    ret := FindNext(SearchRec);
  end;
  sysutils.FindClose(SearchRec);
end;

function BOB_IMPORT_PCL2(CodeProduit: string; sLibelle: string; TypeBOZ: boolean = False; NumBasePourMajForce: string = ''): Integer; //LMO20070102
var sFileBOB: string;
  Chemin: string;
  SearchRec: TSearchRec;
  NumVersion: integer;
  ret: integer;
  iCase_l: integer;
  ZoomSuivi: TForm;
  TH: THLABEL;
  st, prefixe, tipe : string; //LMO20061129
  i: integer;

    function FullVersion(ver:string) :string;
    var st : string ;
    begin
      st:=copy('0000', 1, 4-length(ver)) ;
      result:=st + ver ;
    end ;

begin
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version 999
  // - extension .BOB
  // - exemple CCS50582F001.BOB

  // CODE RETOUR DE LA FONCTION

  Result := 0;
  NumVersion := 0;

  if (NumBasePourMajForce <> '') then //+LM20070702
  begin
    chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\' + CodeProduit + FullVersion(NumBasePourMajForce) + 'F*.BO*';
    i := SysUtils.FindFirst(chemin, faAnyFile, SearchRec);
    SysUtils.FindClose ( SearchRec ) ;
    if (i =0) then //=fichier trouvé
    begin
      if (PgiAsk('Intégrer ' + CodeProduit + '?') <> mrYes) then
        exit //LMO20070102
    end
    else
      exit;
  end; //-LM20070702


  ZoomSuivi := TForm.Create(nil);
  ZoomSuivi.Caption := 'Mise à jour des ' + sLibelle;
  ZoomSuivi.BorderIcons := [biSystemMenu, biMinimize, biMaximize];
  ZoomSuivi.BorderStyle := bsSizeable;
  ZoomSuivi.Height := 70;
  ZoomSuivi.Width := 400;
  ZoomSuivi.Position := poScreenCenter;
  ZoomSuivi.BorderIcons := [];
  TH := THLABEL.Create(ZoomSuivi);
  TH.Parent := ZoomSuivi;
  TH.Caption := '';
  TH.Top := 10;
  TH.left := 10;
//          ZoomSuivi.WindowState := wsMinimized	;
  ZoomSuivi.KeyPreview := TRUE;
//          TravauxPlanning.Parent  := ZoomSuivi;
//          PanelTravaux.Caption := 'Affichage des travaux dans une fenêtre agrandie';
  ZoomSuivi.Show;
  ZoomSuivi.FormStyle := fsStayOnTop;
//          pgiinfo('test apres show');

  // Fin de la fenêtre maximiser: le planning "revient" dans la fenêtre de lancement
//          TravauxPlanning.Parent := PanelTravaux;

  Chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\'; //EX: C:\PGI00\BOB\2035\
  (*if V_PGI.SAV then
  begin
    PGIINFO('dans BOB_IMPORT_PCL, répertoire : ' + Chemin);
    //if typeboz then PGIINFO('typeboz à true') else PGIINFO('typeboz à false');
  end;*)

  //+LMO20070102
  //Si on souhaite forcer l'intégration des bobs, on scanne la présence de fichier avec ou sans préfixe
  if (NumBasePourMajForce <> '')
    and (((FindFirst(Chemin + '_' + CodeProduit + '*.BOZ', faAnyFile, SearchRec) = 0) and TypeBOZ)
    or ((FindFirst(Chemin + '_' + CodeProduit + '*.BOB', faAnyFile, SearchRec) = 0) and (not TypeBOZ))) then
    prefixe := '_'
  else
    prefixe := '';
  sysutils.FindClose(SearchRec);
  //-LMO20070102

  if TypeBOZ then
    ret := FindFirst(Chemin + prefixe + CodeProduit + '*.BOZ', faAnyFile, SearchRec) //LMO20070102
  else
    ret := FindFirst(Chemin + prefixe + CodeProduit + '*.BOB', faAnyFile, SearchRec); //LMO20070102

  while ret = 0 do
  begin
      //RECUPERE NOM DU BOB
    sFileBOB := SearchRec.Name;
    //if V_PGI.SAV then PGIINFO('Analyse du BOB : ' + sFileBOB);
      //RECUPERE NUM VERSION
    try
      NumVersion := ValeurI(Copy(sFileBOB, 5 + length(prefixe), 4)); //LMO20070102 NumVersion := ValeurI(Copy(sFileBOB, 5, 4));
    except
      Result := -6;
      break;
    end;

       // MD 08/11/04 if NumVersion > V_PGI.NumVersionBase then
       // => remplacé par NumVersionSoc sinon risque pour un vieil exe
       // de remonter des bobs périmés dans une base récente !!
    if (NumVersion > V_PGI.NumVersionSoc)
      or (valeurI(NumBasePourMajForce) = NumVersion) then //LMO20070102
    begin
      //+LMO20070102
      tipe := copy(sFileBob, 9 + length(prefixe), 1);
      if tipe = 'M' then TH.Caption := 'Mise à jour Menu n° ' + sFileBOB
      else if tipe = 'D' then TH.Caption := 'Mise à jour Données n° ' + sFileBOB
      else if tipe = 'F' then TH.Caption := 'Mise à jour Fiche n° ' + sFileBOB
      else if tipe = 'A' then TH.Caption := 'Mise à jour Tablette n° ' + sFileBOB;

      if NumBasePourMajForce <> '' then
      begin // si on veut réimporter => on supprime ce qui existe en base
        st := UpperCase(ExtractFileName(sFileBOB));
        i := pos('.', st);
        if i > 0 then delete(st, i, 50);
        if st[1] = '_' then st := copy(st, 2, 50);
        if ExecuteSql('delete from YMYBOBS where YB_BOBNAME="' + st + '"') = 0 then
          TH.Caption := 'Problème dans la suppression de ' + st;
      end;

      try
        if (NumBasePourMajForce <> '') and (sFileBob[1] = '_') then
        begin
          sFileBob := copy(sFileBob, 2, 50);
          renameFile(Chemin + '_' + sFileBob, Chemin + sFileBob);
        end;
        iCase_l := tmpIntegreBob(Chemin + sFileBOB, FALSE, (NumBasePourMajForce = ''), nil, false); //LMO en attendant le ctl version par l'agl iCase_l := AglIntegreBob(Chemin + sFileBOB, FALSE, (NumBasePourMajForce=''));

      except
        iCase_l := -2;
      end;
      //-LMO20070102

      case iCase_l of
        0: // OK
          begin
            if V_PGI.SAV then Pgiinfo('Intégration de : ' + sFileBOB, TitreHalley); //Resultif not LIA_JOURNAL_EVENEMENT(sTempo) then Result := -1;
            if copy(sFileBob, 9, 1) = 'M' then Result := 1; //SI BOB AVEC MENU, ON REND 1 POUR SORTIR DE L'APPLICATION
          end;
        1: if V_PGI.SAV then Pgiinfo('Intégration déjà effectuée :' + sFileBOB, TitreHalley); // Intégration déjà effectuée
        -1: // Erreur d'écriture dans la table YMYBOBS
          begin
            if V_PGI.SAV then PGIInfo('Erreur d''écriture dans la table YMYBOBS :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
        -2: // Erreur d'intégration dans la fonction AglImportBob
          begin
            if V_PGI.SAV then PGIInfo('Erreur d''intégration dans la fonction AglImportBob :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
        -3: //Erreur de lecture du fichier BOB.
          begin
            if V_PGI.SAV then PGIInfo('Erreur de lecture du fichier BOB :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
        -4: // Erreur inconnue.
          begin
            if V_PGI.SAV then PGIInfo('Erreur inconnue :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
          end;
      else
        if V_PGI.SAV then PGIInfo('Erreur inconnue :[' + IntToStr(iCase_l) + '],' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');

      end;

    end;
    ret := FindNext(SearchRec);
  end;
  sysutils.FindClose(SearchRec);
  ZoomSuivi.Free;

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
  Result :=TCBPPath.GetCegidDistriBob ;
{
  showmessage('result1='+Result);

  Result := UpperCase(ExtractFilePath(Application.ExeName));
  // retire \APP
  if (Length(Result) > 5) and (UpperCase(Copy(Result, Length(Result) - 4, 5)) = '\APP\') then
    Result := Copy(Result, 1, Length(Result) - 4);
  SupprimeLastAntislash(Result);
  // rajoute \BOB
  Result := Result + '\BOB';
  if (not DirectoryExists(Result)) and (DirectoryExists('C:\PGI00\BOB')) then
    Result := 'C:\PGI00\BOB';
  showmessage('result2='+Result);
}
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
  Chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\'; //EX : C:\PGI00\BOB\2035\
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
  Chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\'; //EX : C:\PGI00\BOB\2035\
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

function BOB_Alerte(CodeProduit: string; NomProduit: string = ''; nb: integer = 3; poids: integer = 0): integer;
var
  Chemin: string;
  SearchRec: TSearchRec;
  ret, i, p: integer;
  parle: boolean;
begin
//  result := 0;
  i := 0;
  p := 0;


  Chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\'; //EX : C:\PGI00\BOB\2035\
  ret := FindFirst(Chemin + CodeProduit + '*.BO*', faAnyFile, SearchRec);
  while ret = 0 do
  begin
    i := i + 1;

    p := p + SearchRec.Size;

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


function tmpIntegreBob(const FileBobName: string; CreationSeulement: boolean =
  False; SiTrouveJeSors: boolean = false; CallBack:
  TTImportObjectCallBack = nil; controleNoVersion: boolean = true): Integer; //LMO20070108
{$IFNDEF HVCL}
var
  BobName: string;
  MdlName: string;
  I: Integer;
  Q: TQuery;
  FindIt: Boolean;
  sMdl: string;
{$IFDEF aucasou}
  // XP 18-04-2005 : WC
  L: TStrings;
{$ENDIF aucasou}
  T: TOB;
  V: Integer;
  { Ajout XP du 19-07-2002 }
  Libelle: string;
  { Ajout XP du 15-01-2003 }
  Bool: boolean;
  datecreation: tdatetime;

  Stream: TMemoryStream;

  // XP 31-05-2005
  SocRefDepart, SocRefFinale: integer;
  BobType: TTypeBob;
  OkSuite: boolean;
{$IFNDEF EAGL} // XP 22.02.2006
  SaveEnabledDeshare: boolean;
{$ENDIF EAGL} // XP 22.02.2006
  VersionSoc: integer; // XP 22.02.2006
{$IFDEF EAGLSERVER} // XP 22.02.2006
  LaSession: TiSession;
{$ENDIF EAGLSERVER}
  // XP 15.09.2006 FQ 12889
  boolFaitParGep: boolean;
{$ENDIF HVCL}
begin
  Result := 0;
{$IFNDEF HVCL}

{$IFDEF EAGLSERVER}
  LaSession := LookUpCurrentSession;
  VersionSoc := LaSession.NumVersionSoc;
{$ELSE}
  VersionSoc := V_PGI.NumVersionSoc;
{$ENDIF}

{$IFNDEF EAGL} // XP 22.02.2006
  SaveEnabledDeshare := V_PGI.enableDEShare;
{$ENDIF EAGL}

  T := nil;
  SetLength(sMdl, 0);

  // Est-ce que le fichier existe ?
  if not FileExists(FileBobName) then
  begin
    result := -5;
    exit;
  end;

  try
    try
      { Extraction du nom du fichier sans extension }
      BobName := UpperCase(ExtractFileName(FileBobName));
      I := Pos('.', BobName);
      if I > 0 then
        Delete(BobName, I, 128);

      { Est-ce que cette BOB est déjà intégrée dans YMYBOBS ? }
      Q := OpenSql('SELECT * FROM YMYBOBS WHERE YB_BOBNAME="' + BobName + '"',
        True, 1);
      if not q.eof then
      begin
        datecreation := q.findfield('YB_BOBDATECREAT').asdatetime;
      end
      else
        datecreation := now();
      FindIt := not Q.Eof;
      Ferme(Q);

      // added by XP le 02-04-2003
      if SiTrouveJeSors and FindIt then
      begin
        result := 1;
        exit;
      end;

      { La BOB n'existe pas }
      SocRefDepart := 0; // MajTable
      SocRefFinale := 0;
      boolFaitParGep := False; // Par défaut
      BobType := tbNormal;
      try
        // Récupération de la version du fichier BOB et plus
        T := TOB.Create('Maman', nil, -1);
        if AglReadContentBob(FileBobName, T, True) then
        begin
          V := T.GetValue('BOBVERSION');
          Libelle := T.GetValue('BOBNAME');

          // XP 15.09.2006 FQ 12889
          if T.FieldExists('BOBCREATEUR') and (UpperCase(T.GetString('BOBCREATEUR')) = 'CGEP') then
            boolFaitParGep := True;

          // XP 31-05-2005 : Récupération des éléments des nouvelles bob
          if T.FieldExists('BOBSOCREFDEPART') then
          begin
            SocRefDepart := T.GetInteger('BOBSOCREFDEPART');
            SocRefFinale := T.GetInteger('BOBSOCREFFINALE');
            BobType := TTypeBob(T.GetInteger('BOBTYPE'));
            // XP 05-09-2005 : Pour palier pb de non mise à jour du type de bob
            case BobType of
              tbMajStruct: ;
              tbNormalControl: ;
            else
              BobType := tbNormal;
            end;
          end;
        end
        else
        begin
          V := 0;
          libelle := '';
        end;
        FreeAndNil(T);
      except
        on E: Exception do
        begin
          Result := -3;
          raise;
        end;
      end;

      // XP 31-05-2005 : Contrôle pour les BOB type NormalControl et MajStruct
      // XP 15.09.2006 Il faut autoriser quand même les BOB du GEP car ne concerne que des
      // données Z
      if boolFaitParGep then
        OkSuite := True
      else if BobType <> tbNormal then
      begin
        // XP 06.03.2006 On test maintenant si <=
        // OkSuite := (VersionSoc >= SocRefDepart) and (VersionSoc < SocrefFinale);
        if controleNoVersion or (BobType = tbMajStruct) then //+LMO20070108
          OkSuite := (VersionSoc >= SocRefDepart) and (VersionSoc <= SocrefFinale)
        else
          OkSuite := true; //-LMO20070108

        if not OkSuite then
          Result := -99;
      end
      else
{$IFDEF DECLA}
        OkSuite := True;
{$ELSE}
{$IFDEF CGEP}
        OkSuite := True;
{$ELSE}
      begin
        OkSuite := False; // XP 06.03.2006 Les BOB sans contrôles sont interdites
        Result := -100;
      end;
{$ENDIF}
{$ENDIF}

      // XP 08-06-2005 : Blocage. Exclusif pour PgiMajVer.exe
      if OkSuite and (BobType = tbMajStruct) then
      begin
        OkSuite := (LowerCase(ExtractFileName(ParamStr(0))) = 'pgimajver.exe')
          and (V_PGI.IsPgiMajVer);
        if not OkSuite then
          Result := -98;
      end;

      // Dépend, du type de BOB et du n° de version de la socref
      if OkSuite then
      begin
        // Dans le cas d'une mise à jour de structure, il faut désactiver la redirection DESHARE
        if BobType = tbMajStruct then
        begin
{$IFNDEF EAGLSERVER} // XP 22.02.2006
          V_PGI.enableDEShare := False;
{$IFNDEF EAGL}
{$IFNDEF BASEEXT}
{$ENDIF}
{$ENDIF}
{$ENDIF EAGLSERVER} // XP 22.02.2006
        end;

        try
          { Récupération des informations de la BOB }
          if AglImportBob(FileBobName, False, False, CallBack,
            CreationSeulement, False, nil) then // XP 18.10.2006 FQ 12989
          begin
            // XP 12-08-2005
            // Pour permettre l'utilisation des paramsoc par les applications dans MajApres
{$IFNDEF EAGLSERVER} // XP 22.02.2006 Pas la peine en eaglserver, il faut supprimer les caches...
            try
              if TOB(V_PGI.TOBSOC) <> nil then
              begin
                TOB(V_PGI.TOBSOC).Free;
                V_PGI.TOBSOC := nil;
              end
            except
            end;
{$ENDIF EAGLSERVER} // XP 22.02.2006 Pas la peine en eaglserver, il faut supprimer les caches...

{$IFNDEF EAGL} // XP 22.02.2006 {$IFNDEF EAGLCLIENT}
            // XP 09-08-2005 : On force le nouveau n° de société
            if BobType = tbMajStruct then
            begin
              //AglSetVersionSoc(SocrefFinale);

              { Pour forcer le chargement du dictionnaire }
              V_PGI.DEDejaCharge := False;

              { Dictionnaire }
              ChargeTablePrefixe(False, False);
            end;
{$ENDIF}

            { Est-ce que le fichier modèle est présent dans le même répertoire ? }
            MdlName := ChangeFileExt(FileBobName, '.MDL');
            if FileExists(MdlName) then
            begin
              Stream := TMemoryStream.Create();
              try
                Stream.LoadFromFile(MdlName);
                Stream.Seek(0, 0);
                SetLength(sMdl, Stream.Size);
                Stream.Read(pchar(sMdl)^, Stream.Size);
              finally
                FreeAndNil(Stream);
              end;
            end;
{$IFDEF aucasou}
            if FileExists(MdlName) then
            begin
              L := TStringList.Create;
              try
                L.LoadFromFile(MdlName);
                sMdl := L.Text;
              finally
                freeandnil(l);
              end;
            end
            else
              sMdl := '';
{$ENDIF aucasou}
            try
              T := TOB.Create('YMYBOBS', nil, -1);
              T.PutValue('YB_BOBNAME', BobName);
              T.PutValue('YB_BOBLIBELLE', Libelle);
              T.PutValue('YB_BOBVERSION', V);
              T.PutValue('YB_BOBDATECREAT', datecreation);
              T.PutValue('YB_BOBDATEMODIF', Now());
              if copy(BobName, 1, 1) = 'Z' then
                T.PutValue('YB_DOMAINE', 'Z')
              else
                T.PutValue('YB_DOMAINE', '');
              T.PutValue('YB_BOBMODELE', sMdl);
              if FindIt then
                bool := T.UpdateDB()
              else
                bool := T.InsertDb(nil);
              if bool then
              begin
                Result := 0;
              end
              else
              begin
                Result := -1;
              end;
            except
              on E: Exception do
                Result := -1;
            end;
          end
          else { erreur lors de l'import }
            result := -2;
        finally
{$IFNDEF EAGL} // XP 22.02.2006
          V_PGI.enableDEShare := SaveEnabledDeshare;
{$ENDIF EAGLSERVER}
        end;
      end;
    except
      on E: Exception do
      begin
        { Erreur inconnue }
        if Result = 0 then
          Result := -4;
      end;
    end;
  finally
    if assigned(T) then
      FreeAndNil(T);
    SetLength(sMdl, 0);
  end;
{$ENDIF HVCL}
end;


function BOB_IMPORT_PCL_STD_MANU(FiLeName : String ): Integer;
begin
  {$IFNDEF EAGLCLIENT}
     Result := AglIntegreBob(FiLeName,FALSE, TRUE) ;
  {$ELSE}
     Result := IntegreBOBOnDemand(ExtractFileName(FiLeName)) ;
  {$ENDIF}
end;

function BOB_IMPORT_PCL_STD(CodeProduit: string; sLibelle: string; NumBasePourMajForce: string = ''): Integer; //LMO20070102
var sFileBOB: string;
  Chemin: string;
  SearchRec: TSearchRec;
  NumVersion: integer;
  ret: integer;
  iCase_l: integer;
  ZoomSuivi: TForm;
  TH: THLABEL;
  st, tipe : string; //LMO20061129
  i: integer;

    function FullVersion(ver:string) :string;
    var st : string ;
    begin
      st:=copy('0000', 1, 4-length(ver)) ;
      result:=st + ver ;
    end ;

begin
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version 999
  // - extension .BOB
  // - exemple CCS50582F001.BOB

  // CODE RETOUR DE LA FONCTION
  Result := 0;
  NumVersion := 0;

  {$IFNDEF eAGLClient}
  if (NumBasePourMajForce <> '') then //+LM20070702
  begin
    chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\' + CodeProduit + FullVersion(NumBasePourMajForce) + 'F*.BO*';
    i := SysUtils.FindFirst(chemin, faAnyFile, SearchRec);
    SysUtils.FindClose ( SearchRec ) ;
    if (i =0) then //=fichier trouvé
    begin
      if (PgiAsk('Intégrer ' + CodeProduit + '?') <> mrYes) then
        exit //LMO20070102
    end
    else
      exit;
  end; //-LM20070702
  {$ENDIF}



  ZoomSuivi := TForm.Create(nil);
  ZoomSuivi.Caption := 'Mise à jour des ' + sLibelle;
  ZoomSuivi.BorderIcons := [biSystemMenu, biMinimize, biMaximize];
  ZoomSuivi.BorderStyle := bsSizeable;
  ZoomSuivi.Height := 70;
  ZoomSuivi.Width := 400;
  ZoomSuivi.Position := poScreenCenter;
  ZoomSuivi.BorderIcons := [];
  TH := THLABEL.Create(ZoomSuivi);
  TH.Parent := ZoomSuivi;
  TH.Caption := '';
  {$IFDEF eAGLClient}
  TH.Caption := 'Intégration des BOB. Veuillez patienter.';
  {$ENDIF}

  TH.Top := 10;
  TH.left := 10;
//          ZoomSuivi.WindowState := wsMinimized	;
  ZoomSuivi.KeyPreview := TRUE;
//          TravauxPlanning.Parent  := ZoomSuivi;
//          PanelTravaux.Caption := 'Affichage des travaux dans une fenêtre agrandie';
  ZoomSuivi.Show;
  ZoomSuivi.FormStyle := fsStayOnTop;
//          pgiinfo('test apres show');

  // Fin de la fenêtre maximiser: le planning "revient" dans la fenêtre de lancement
//          TravauxPlanning.Parent := PanelTravaux;
  {$IFNDEF eAGLClient}

  Chemin := BOB_GET_CHEMIN + '\' + CodeProduit + '\'; //EX: C:\PGI00\BOB\2035\
  (*if V_PGI.SAV then
  begin
    PGIINFO('dans BOB_IMPORT_PCL, répertoire : ' + Chemin);
    //if typeboz then PGIINFO('typeboz à true') else PGIINFO('typeboz à false');
  end;*)

  //+LMO20070102
  //Si on souhaite forcer l'intégration des bobs, on scanne la présence de fichier avec ou sans préfixe
{  if (NumBasePourMajForce <> '')
    and (((FindFirst(Chemin + '_' + CodeProduit + '*.BOZ', faAnyFile, SearchRec) = 0) and TypeBOZ)
    or ((FindFirst(Chemin + '_' + CodeProduit + '*.BOB', faAnyFile, SearchRec) = 0) and (not TypeBOZ))) then
    prefixe := '_'
  else
    prefixe := '';
  sysutils.FindClose(SearchRec); }
  //-LMO20070102
  // MB : Ne sert plus : Fonciton CanIntegreBOB remplace.

  ret := FindFirst(Chemin + '*.*', faAnyFile, SearchRec) ;

  while ret = 0 do
  begin
      //RECUPERE NOM DU BOB
    if CanIntegreBob(SearchRec,NumBasePourMajForce <> '') then
    begin
        sFileBOB := SearchRec.Name;
        //if V_PGI.SAV then PGIINFO('Analyse du BOB : ' + sFileBOB);
          //RECUPERE NUM VERSION
        tipe := 'F' ;
        try
          if copy(sFileBOB,1,1) = '_' then
          begin
             NumVersion := ValeurI(Copy(sFileBOB, 6, 4));
             tipe := copy(sFileBob, 10, 1);
          end
          else
          begin
             NumVersion := ValeurI(Copy(sFileBOB, 5, 4));
             tipe := copy(sFileBob, 9 , 1);
          end;
        except
          NumVersion := -1;
        end;

           // MD 08/11/04 if NumVersion > V_PGI.NumVersionBase then
           // => remplacé par NumVersionSoc sinon risque pour un vieil exe
           // de remonter des bobs périmés dans une base récente !!
        if (NumVersion > V_PGI.NumVersionSoc)
          or (valeurI(NumBasePourMajForce) = NumVersion) then //LMO20070102
        begin
          //+LMO20070102
          if tipe = 'M' then TH.Caption := 'Mise à jour Menu n° ' + sFileBOB
          else if tipe = 'D' then TH.Caption := 'Mise à jour Données n° ' + sFileBOB
          else if tipe = 'F' then TH.Caption := 'Mise à jour Fiche n° ' + sFileBOB
          else if tipe = 'A' then TH.Caption := 'Mise à jour Tablette n° ' + sFileBOB;

          if NumBasePourMajForce <> '' then
          begin // si on veut réimporter => on supprime ce qui existe en base
            st := UpperCase(ExtractFileName(sFileBOB));
            i := pos('.', st);
            if i > 0 then delete(st, i, 50);
            if st[1] = '_' then st := copy(st, 2, 50);
            if ExecuteSql('delete from YMYBOBS where YB_BOBNAME="' + st + '"') = 0 then
              TH.Caption := 'Problème dans la suppression de ' + st;
          end;

          try
            if (NumBasePourMajForce <> '') and (sFileBob[1] = '_') then
            begin
              sFileBob := copy(sFileBob, 2, 50);
              renameFile(Chemin + '_' + sFileBob, Chemin + sFileBob);
            end;
            iCase_l := tmpIntegreBob(Chemin + sFileBOB, FALSE, (NumBasePourMajForce = ''), nil, false); //LMO en attendant le ctl version par l'agl iCase_l := AglIntegreBob(Chemin + sFileBOB, FALSE, (NumBasePourMajForce=''));

          except
            iCase_l := -2;
          end;
          //-LMO20070102

          case iCase_l of
            0: // OK
              begin
                if V_PGI.SAV then Pgiinfo('Intégration de : ' + sFileBOB, TitreHalley); //Resultif not LIA_JOURNAL_EVENEMENT(sTempo) then Result := -1;
                if copy(sFileBob, 9, 1) = 'M' then Result := 1; //SI BOB AVEC MENU, ON REND 1 POUR SORTIR DE L'APPLICATION
              end;
            1: if V_PGI.SAV then Pgiinfo('Intégration déjà effectuée :' + sFileBOB, TitreHalley); // Intégration déjà effectuée
            -1: // Erreur d'écriture dans la table YMYBOBS
              begin
                if V_PGI.SAV then PGIInfo('Erreur d''écriture dans la table YMYBOBS :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
              end;
            -2: // Erreur d'intégration dans la fonction AglImportBob
              begin
                if V_PGI.SAV then PGIInfo('Erreur d''intégration dans la fonction AglImportBob :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
              end;
            -3: //Erreur de lecture du fichier BOB.
              begin
                if V_PGI.SAV then PGIInfo('Erreur de lecture du fichier BOB :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
              end;
            -4: // Erreur inconnue.
              begin
                if V_PGI.SAV then PGIInfo('Erreur inconnue :' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');
              end;
          else
            if V_PGI.SAV then PGIInfo('Erreur inconnue :[' + IntToStr(iCase_l) + '],' + Chemin + sFileBOB, 'PCL_IMPORT_BOB');

          end;

        end;
    end;
    ret := FindNext(SearchRec);
  end;
  sysutils.FindClose(SearchRec);
  {$ELSE}
  // Gestion en EAGL
  IntegreBOBAuto(CodeProduit,NumBasePourMajForce);
  {$ENDIF}

  ZoomSuivi.Free;

end;

end.

