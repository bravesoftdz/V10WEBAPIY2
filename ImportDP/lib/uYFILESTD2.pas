{***********UNITE*************************************************
Auteur  ...... : CHRISTOPHE VACHEZ
Créé le ...... : 29/06/2005
Modifié le ... :   /  /
Description .. : Fonction de gestion des fichiers en base
Mots clefs ... :
*****************************************************************}
unit uYFILESTD2;

interface

uses
  windows,
  sysutils,
  hctrls,
  utob,
  dialogs,
  // FileCtrl,
  {$IFNDEF eAGLClient}
  db,
  {$IFNDEF DBXPRESS}dbtables,
  {$ELSE}uDbxDataSet,
  {$ENDIF}
  {$ENDIF eAGLClient}
  uGedN;

// fonction Extract
function AGL_YFILESTD_EXTRACT(var FilePath: string; const CODEPRODUIT, NOM: string;
  const CRIT1: string = ''; const CRIT2: string = ''; const CRIT3: string = ''; const CRIT4: string = ''; const CRIT5: string = '';
  const TestNum: boolean = false; const LANGUE: string = 'FRA'; const PREDEFINI: string = 'CEG'; const Dossier: string = ''): integer;

// fonction Import
// XP 28-11-2005 FQ 12082
function AGL_YFILESTD_IMPORT(const FilePath, CODEPRODUIT, NOM, EXTFILE: string;
  const CRIT1: string = ''; const CRIT2: string = ''; const CRIT3: string = ''; const CRIT4: string = ''; const CRIT5: string = '';
  const B1: string = '-'; const B2: string = '-'; const B3: string = '-'; const B4: string = '-'; const B5: string = '-';
  const LANGUE: string = 'FRA'; const PREDEFINI: string = 'CEG'; const LIBELLE: string = 'Ressource'; const Dossier: string = ''): integer;

// fonction Recuperation message erreur
function AGL_YFILESTD_GET_ERR(CODE: integer): string;

// fonctions utilisees dans Import/Extract :
{$IFNDEF OLDGED}
function AGL_YFILESTD_EXTRACTGED(const FilePath: string; const ID: string): integer;
{$ELSE}
function AGL_YFILESTD_EXTRACTGED(const FilePath: string; ID: integer): integer;
{$ENDIF}
function AGL_YFILESTD_GET_NOMVERSION(const NOM: string; const NUM: integer): string;
function AGL_YFILESTD_GET_PATH(const CODEPRODUIT, NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI: string; const Dossier: string = ''): string;
function AGL_YFILESTD_GET_TOB(const CODEPRODUIT, NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI: string; const DOSSIER: string = ''): TOB;

// XP 28-11-2005 FQ 12084
function Agl_YFileStd_Delete(const CodeProduit, Nom, Langue, Predefini, Crit1, NoDossier: string): Integer;

implementation

uses Hent1;

const
  {$IFNDEF OLDGED}
  VALEUR_VIDE = '';
  {$ELSE}
  VALEUR_VIDE = -1;
  {$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 28/11/2005
Modifié le ... :   /  /
Description .. : Permet la suppression d'un fichier de la base
Mots clefs ... :
*****************************************************************}
function Agl_YFileStd_Delete(const CodeProduit, Nom, Langue, Predefini, Crit1, NoDossier: string): Integer;
var
  Q: Tquery;
  MaGed: TMaGed;
  Erreur: integer;
begin
  Erreur := 0;
  try
  // XP 17.07.2007 FQ 14115 YFS_FILEGUID en lieu et place de YFS_FILEID !!!!
    Q := OpenSql('SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="' + Trim(CodeProduit) + '" AND YFS_NOM="' + Trim(Nom) + '" AND YFS_LANGUE="' + Trim(Langue) + '" AND YFS_PREDEFINI="' + Trim(Predefini) + '" AND YFS_CRIT1="' + Trim(Crit1) + '" AND YFS_NODOSSIER="' + Trim(NoDossier) + '"', True);
    if not Q.Eof and (Q.FindField('YFS_FILEGUID').AsString <> VALEUR_VIDE) then
    begin
      BeginTrans();
      try
      // initialisation de la GED
        MaGed := uGedN.TMaGed.Init();
        if Assigned(MaGed) then
        begin
          try
            MaGed.Erase(Q.FindField('YFS_FILEGUID').AsString);
          finally
          // Destruction de la GED
            FreeAndNil(MaGed);
          end;
        end;

      // Suppression de la GED
        if ExecuteSql('DELETE FROM YFILESTD WHERE YFS_CODEPRODUIT="' + Trim(CodeProduit) + '" AND YFS_NOM="' + Trim(Nom) + '" AND YFS_LANGUE="' + Trim(Langue) + '" AND YFS_PREDEFINI="' + Trim(Predefini) + '" AND YFS_CRIT1="' + Trim(Crit1) + '" AND YFS_NODOSSIER="' + Trim(NoDossier) + '"') <> 1 then
        begin
          Erreur := -3;
          Raise Exception.Create ('Erreur suppression GED') ;
        end;

        CommitTrans();
      except
        on E: Exception do
        begin
          RollBack();
          if Erreur = 0 then
            Erreur := -2;
        end;
      end;
    end
    else
      Erreur := -1;
  finally
    Result := Erreur;
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 27/06/2005
Modifié le ... :   /  /
Description .. : Verifie si un enregistrement existe deja
Suite ........ : avec la  meme cle, et recupere la TOB associee a
Suite ........ : l'enregistrement
Mots clefs ... :
*****************************************************************}
function AGL_YFILESTD_GET_TOB(const CODEPRODUIT, NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI: string; const DOSSIER: string = ''): TOB;
var
  sSQL: string;
  Q: TQuery;
  TT, TRes: TOB;
begin
  Result := nil; // valeur par defaut

  sSQL := 'SELECT * FROM YFILESTD WHERE '
    + 'YFS_CODEPRODUIT =  "' + CODEPRODUIT + '" AND '
    + 'YFS_NOM =  "' + UpperCase(NOM) + '" AND '
    + 'YFS_LANGUE =  "' + UpperCase(LANGUE) + '" AND '
    + 'YFS_PREDEFINI =  "' + UpperCase(PREDEFINI) + '" AND '
    + 'YFS_CRIT1 =  "' + CRIT1 + '" AND '
    + 'YFS_CRIT2 =  "' + CRIT2 + '" AND '
    + 'YFS_CRIT3 =  "' + CRIT3 + '" AND '
    + 'YFS_CRIT4 =  "' + CRIT4 + '" AND '
    + 'YFS_CRIT5 =  "' + CRIT5 + '" ';

  // XP 24.05.2007 FQ 13859
  if Dossier <> '' then
    sSql := sSql + 'AND YFS_NODOSSIER = "' + Dossier + '"';

  Q := OpenSQL(sSQL, true); // ouverture de la requete
  if not Q.Eof then // si un enregistrement a ete trouve
  begin
    TT := TOB.Create('_', nil, -1); // creation de la tob
    try
      TT.LoadDetailDB('YFILESTD', '', '', Q, False);
      if TT.Detail.Count = 1 then // la cle etant unique, on ne doit recuperer qu'un enregistrement, sinon, on renvoie nil
      begin
        TRes := TT.Detail[0];
        TRes.ChangeParent(nil, -1);
        Result := TRes;
      end;
    finally
      FreeAndNil(TT); // on fait un free pour liberer la TOB mere
    end;
  end;
  Ferme(Q); // fermeture du query
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 29/06/2005
Modifié le ... :   /  /
Description .. : Recuperation du Path du fichier a extraire en fonction des
Suite ........ : parametres
Mots clefs ... :
*****************************************************************}
function AGL_YFILESTD_GET_PATH(const CODEPRODUIT, NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI: string; const Dossier: string = ''): string;
var
  Path: array[0..255] of Char;
  St: string;
begin
  GetTempPath(255, Path); // recup du repertoire Temp -> voir si la methode est la bonne
  st := StrPas(Path) + 'PGI\STD\' + CODEPRODUIT + '\'; // debut de la construction de l'arborescence

  // PB 24.05.2007 FQ 13859
  if Dossier <> '' then
    st := st + Dossier + '\';

  if CRIT1 <> '' then
    st := st + CRIT1 + '\';
  if CRIT2 <> '' then
    st := st + CRIT2 + '\';
  if CRIT3 <> '' then
    st := st + CRIT3 + '\';
  if CRIT4 <> '' then
    st := st + CRIT4 + '\';
  if CRIT5 <> '' then
    st := st + CRIT5 + '\';
  st := st + UpperCase(LANGUE) + '\' + UpperCase(PREDEFINI) + '\'; // ajout de la langue et du predefini

  Result := st + NOM (*+ '.' + EXT*);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 27/06/2005
Modifié le ... : 29/06/2005
Description .. : Recupere le path du fichier extrait de la GED en fonction
Suite ........ : des criteres demandes.
Suite ........ : Gestion du NumVersion de fichier si TestNum est TRUE
Mots clefs ... :
*****************************************************************}
function AGL_YFILESTD_EXTRACT(var FilePath: string; const CODEPRODUIT, NOM: string;
  const CRIT1: string = ''; const CRIT2: string = ''; const CRIT3: string = ''; const CRIT4: string = ''; const CRIT5: string = '';
  const TestNum: boolean = false; const LANGUE: string = 'FRA'; const PREDEFINI: string = 'CEG'; const Dossier: string = ''): integer;
var
  TR: TOB;
  {$IFNDEF OLDGED}
  ID: string;
  {$ELSE}
  ID: integer;
  {$ENDIF}
  continue: boolean;
  MON_NOM: string;
begin
  if not TestNum then // si on ne teste pas le numVersion
  begin
    FilePath := AGL_YFILESTD_GET_PATH(CODEPRODUIT, NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI, DOssier); // recup du path
    Continue := not FileExists(FilePath); // continue l'extraction si le fichier n'existe pas
  end
  else
    Continue := TestNum; // continue l'extraction

  if Continue then
  begin
    TR := AGL_YFILESTD_GET_TOB(CODEPRODUIT, NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI, Dossier); // recuperation de l'enregistrement
    if Assigned(TR) then // s'il y a un enregistrement
    begin
      try
        if TestNum then // si on teste le numVersion
        begin
          MON_NOM := AGL_YFILESTD_GET_NOMVERSION(NOM, TR.GetInteger('YFS_NUMVERSION')); // recupere le nouveau nom du fichier  avec 0000NUMVERSION_NOM
          FilePath := AGL_YFILESTD_GET_PATH(CODEPRODUIT, MON_NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI, Dossier); // recup du path
          Continue := not FileExists(FilePath); // continue l'extraction si le fichier n'existe pas
        end;

        if Continue then
        begin
          {$IFNDEF OLDGED}
          ID := TR.GetString('YFS_FILEGUID'); // recupere ID
          {$ELSE}
          ID := TR.GetInteger('YFS_FILEID'); // recupere ID
          {$ENDIF}

          if ID <> VALEUR_VIDE then
            Result := AGL_YFILESTD_EXTRACTGED(FilePath, ID) // extraction depuis la GED
          else
            Result := 202; // aucune ressource associee a l'enregistrement
        end
        else
          Result := -1 // ok
      finally
        FreeAndNil(TR);
      end;
    end
    else
      Result := 201; // Aucune ressouce ne correspond aux criteres saisis
  end
  else
    Result := -1; //  ok
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 29/06/2005
Modifié le ... :   /  /
Description .. : Recuperation du nom de fichier avec la version
Suite ........ : Le nom de fichier est de la forme :
Suite ........ :
Suite ........ : xxxxx_NOM.EXT ou : xxxxx correspond au numero de
Suite ........ : version precede par des 0
Mots clefs ... :
*****************************************************************}
function AGL_YFILESTD_GET_NOMVERSION(const NOM: string; const NUM: integer): string;

  function RECUP_VERSION(NUM: integer; NLength: integer = 4): string; // recupere un codage alpha sur NLength caracteres
  var
    LeNum: string;
  begin
    LeNum := '0000000000' + intToStr(Num);
    Result := Copy(LeNum, Length(LeNum) - (NLength - 1), Length(LeNum));
  end;
begin
  Result := RECUP_VERSION(NUM, 4) + '_' + NOM; // codage du num sur 4 caratectere -> 9999+1 possibilites !
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 29/06/2005
Modifié le ... :   /  /
Description .. : Initialisation de la GED puis
Suite ........ : extraction du FILEID=ID
Suite ........ : dans le fichier FilePath
Mots clefs ... :
*****************************************************************}
{$IFNDEF OLDGED}
function AGL_YFILESTD_EXTRACTGED(const FilePath: string; const ID: string): integer;
{$ELSE}
function AGL_YFILESTD_EXTRACTGED(const FilePath: string; ID: integer): integer;
{$ENDIF}
var
  MaGed: TMaGed;
  Path: string;
begin
  if ID = VALEUR_VIDE then
    Result := 202
  else
  begin
    //if not Assigned(MaGED) then
    MaGed := uGedN.TMaGed.Init(); // initialisation de la GED
    if Assigned(MaGed) then
    begin
      try
        // Affectation des tables et champs NFILES et NFILESPARTS
        {$IFNDEF OLDGED}
        // XP 14-02-2006 Cagade !!!
        MaGed.AffecteFiles('NFILES', 'NFI_FILEGUID', 'NFI_FILENAME', 'NFI_FILESIZE', 'NFI_FILECOMPRESSED',
          'NFI_FILESTORAGE', 'NFI_FILEDATECREATE', 'NFI_FILEDATEMODIFY', 'NFI_CREATEUR', 'NFI_DATECREATION', 'NFI_FILEID'); // XP 27.02.2007 FQ 13538
        MaGed.AffecteFileParts('NFILEPARTS', 'NFS_FILEGUID', 'NFS_PARTID', 'NFS_DATA', 'NFS_FILEID');
        {$ELSE OLDGED}
        MaGed.AffecteFiles('NFILES', 'NFI_FILEID', 'NFI_FILENAME', 'NFI_FILESIZE', 'NFI_FILECOMPRESSED',
          'NFI_FILESTORAGE', 'NFI_FILEDATECREATE', 'NFI_FILEDATEMODIFY', 'NFI_CREATEUR', 'NFI_DATECREATION');
        MaGed.AffecteFileParts('NFILEPARTS', 'NFS_FILEID', 'NFS_PARTID', 'NFS_DATA');
        {$ENDIF OLDGED}

        Path := ExtractFileDir(FilePath); // Extraction du path pour forcer la creation de l'arboresence
        if Path <> '' then
          ForceDirectories(Path); // creation de l'arboresence

        if not MaGED.Extract(ID, FilePath) then
          Result := 203 // Erreur d'extraction de la GED
        else
          Result := -1;
      finally
        FreeAndNil(MaGed); // Destruction de la GED
      end;
    end
    else
      result := 204;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 24/06/2005
Modifié le ... : 24/06/2005
Description .. : Importation dans la Table YFILESTD
Suite ........ : importation du fichier en GED
Mots clefs ... :
*****************************************************************}
// XP 28-11-2005 FQ 12082
function AGL_YFILESTD_IMPORT(const FilePath, CODEPRODUIT, NOM, EXTFILE: string;
  const CRIT1: string = ''; const CRIT2: string = ''; const CRIT3: string = ''; const CRIT4: string = ''; const CRIT5: string = '';
  const B1: string = '-'; const B2: string = '-'; const B3: string = '-'; const B4: string = '-'; const B5: string = '-';
  const LANGUE: string = 'FRA'; const PREDEFINI: string = 'CEG'; const LIBELLE: string = 'Ressource'; const Dossier: string = ''): integer;
var
  MaGed: TMaGed;
  {$IFNDEF OLDGED}
  ID: string;
  NUMV: integer;
  {$ELSE}
  NUMV, ID: integer;
  {$ENDIF}
  TR: TOB;
begin
  ID := VALEUR_VIDE;
  Result := 0; // Erreur

  if not FileExists(FilePath) then
    Result := 100 // le fichier n'existe pas, donc pas d'import possible.
  else
  begin
    //if not Assigned(MaGED) then
    MaGed := uGedN.TMaGed.Init(); // initialisation de la GED
    try
      if not Assigned(MaGed) then
        Result := 1 // probleme de creation de la GED
      else
      begin
        {$IFNDEF OLDGED}
        MaGed.AffecteFiles('NFILES', 'NFI_FILEGUID', 'NFI_FILENAME', 'NFI_FILESIZE', 'NFI_FILECOMPRESSED',
          'NFI_FILESTORAGE', 'NFI_FILEDATECREATE', 'NFI_FILEDATEMODIFY', 'NFI_CREATEUR', 'NFI_DATECREATION', 'NFI_FILEID');
        MaGed.AffecteFileParts('NFILEPARTS', 'NFS_FILEGUID', 'NFS_PARTID', 'NFS_DATA', 'NFS_FILEID');
        {$ELSE OLDGED}
        MaGed.AffecteFiles('NFILES', 'NFI_FILEID', 'NFI_FILENAME', 'NFI_FILESIZE', 'NFI_FILECOMPRESSED',
          'NFI_FILESTORAGE', 'NFI_FILEDATECREATE', 'NFI_FILEDATEMODIFY', 'NFI_CREATEUR', 'NFI_DATECREATION');
        MaGed.AffecteFileParts('NFILEPARTS', 'NFS_FILEID', 'NFS_PARTID', 'NFS_DATA');
        {$ENDIF OLDGED}

        TR := AGL_YFILESTD_GET_TOB(CODEPRODUIT, NOM, CRIT1, CRIT2, CRIT3, CRIT4, CRIT5, LANGUE, PREDEFINI, Dossier);
        if Assigned(TR) then
        begin // Update
          with TR do
          begin
            PutValue('YFS_BCRIT1', B1);
            PutValue('YFS_BCRIT2', B2);
            PutValue('YFS_BCRIT3', B3);
            PutValue('YFS_BCRIT4', B4);
            PutValue('YFS_BCRIT5', B5);

            PutValue('YFS_EXTFILE', UpperCase(EXTFILE));
            PutValue('YFS_LIBELLE', LIBELLE);
            PutValue('YFS_COMFICHIER', '');
            PutValue('YFS_NUMBOB', 0);

            // XP 28-11-2005 FQ 12082
            PutValue('YFS_NODOSSIER', Dossier);

            {$IFNDEF OLDGED}
            ID := GetString('YFS_FILEGUID');
            {$ELSE}
            ID := GetInteger('YFS_FILEID');
            {$ENDIF}
            NUMV := GetInteger('YFS_NUMVERSION');

            if ID <> VALEUR_VIDE then
            begin
              if not MaGed.Replace(FilePath, ID) then
                Result := 103 // Erreur remplacement dans la GED
              else
                PutValue('YFS_NUMVERSION', NUMV + 1);
            end
            else
            begin
              if not MaGed.Import(FilePath, ID) then
                Result := 104 // Erreur d'import dans la GED
              else
              begin
                PutValue('YFS_NUMVERSION', NUMV + 1);
                {$IFNDEF OLDGED}
                PutValue('YFS_FILEGUID', ID);
                {$ELSE}
                PutValue('YFS_FILEID', ID);
                {$ENDIF}
              end;
            end;

            if ID <> VALEUR_VIDE then
            begin
              BeginTrans();
              try
                UpDateDb();
                CommitTrans();
                Result := -1; // Pas d'erreur !
              except
                on E: Exception do
                begin
                  RollBack();
                  result := 101; // Erreur lors de l'insert/ update de YFILESTD
                end;
              end;
              Free; // XP 27.02.2007 Perte de mémoire concernant TR
            end;
          end;
        end
        else
        begin // Insert
          TR := TOB.Create('YFILESTD', nil, -1);
          try
            with TR do
            begin
              PutValue('YFS_CODEPRODUIT', CODEPRODUIT);
              PutValue('YFS_NOM', UpperCase(NOM));
              PutValue('YFS_LANGUE', UpperCase(LANGUE));
              PutValue('YFS_PREDEFINI', UpperCase(PREDEFINI));
              PutValue('YFS_CRIT1', CRIT1);
              PutValue('YFS_CRIT2', CRIT2);
              PutValue('YFS_CRIT3', CRIT3);
              PutValue('YFS_CRIT4', CRIT4);
              PutValue('YFS_CRIT5', CRIT5);

              PutValue('YFS_BCRIT1', B1);
              PutValue('YFS_BCRIT2', B2);
              PutValue('YFS_BCRIT3', B3);
              PutValue('YFS_BCRIT4', B4);
              PutValue('YFS_BCRIT5', B5);

              PutValue('YFS_EXTFILE', UpperCase(EXTFILE));
              PutValue('YFS_LIBELLE', LIBELLE);
              PutValue('YFS_COMFICHIER', '');
              PutValue('YFS_NUMVERSION', 0);
              PutValue('YFS_NUMBOB', 0);

              // XP 19-01-2006 FQ 12219 DEBUT
              // Il ne faut pas tenir compte du N° de dossier quand on écrit avec PREDEFINI = STANDARD, mais
              // uniquement si egal CEG soit CEGID
              if UpperCase(PREDEFINI) = 'STD' then
                PutValue('YFS_NODOSSIER', '000000')
              else
                PutValue('YFS_NODOSSIER', Dossier);
              // XP 19-01-2006 FQ 12219 FIN

              // Si OK
              if MaGed.Import(FilePath, ID) and (ID <> VALEUR_VIDE) then
              begin
                {$IFNDEF OLDGED}
                PutValue('YFS_FILEGUID', ID);
                {$ELSE}
                PutValue('YFS_FILEID', ID);
                {$ENDIF}
                BeginTrans();
                try
                  InsertDB(nil);
                  CommitTrans();
                  Result := -1;
                except
                  on E: Exception do
                  begin
                    RollBack();
                    Result := 101 // Erreur lors de l'insert/ update de YFILESTD
                  end;
                end;
              end
              else
              begin
                Result := 104; // Erreur d'import dans la GED
              end;
            end;
          finally
            FreeAndNil(TR);
          end;
        end;
      end;
    finally
      if assigned(MaGed) then
        FreeAndNil(MaGed); // destruction de la GED
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 24/06/2005
Modifié le ... :   /  /
Description .. : recupere un message d'erreur en fonction du code
Mots clefs ... :
*****************************************************************}
function AGL_YFILESTD_GET_ERR(CODE: integer): string;
begin
  case CODE of
    -1: Result := TraduireMemoire('ok');
    1: Result := TraduireMemoire('Erreur lors de la création de la GED');
    100: Result := TraduireMemoire('Le fichier que vous tentez d''importer n''existe pas.');
    101: Result := TraduireMemoire('Erreur lors de la mise à jour dans la table YFILESTD.');
    103: Result := TraduireMemoire('Erreur lors du remplacement dans la GED.');
    104: Result := TraduireMemoire('Erreur lors de l''import dans la GED.');

    201: Result := TraduireMemoire('Aucune ressource ne correspond aux critères sélectionnés.');
    202: Result := TraduireMemoire('La ressource n''a pas été trouvée.');
    203: Result := TraduireMemoire('Erreur lors de l''extraction dans la GED.');
  else
    Result := TraduireMemoire('Erreur inconnue :') + ' ' + IntToStr(CODE);
  end;
end;

end.

