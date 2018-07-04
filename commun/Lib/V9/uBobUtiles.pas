unit uBobUtiles;

interface


uses
  classes,
  {$ifdef ver150}
  variants,
  {$endif}
  uBOB,
  HEnt1,
  hctrls,
  registry,
  //IniFiles,
  uTob,
  Db,
  {$IFNDEF EAGLCLIENT}
  uDbxDataSet,
  {$ENDIF}
  // cbpIniFiles,
  sysutils,
  Windows,
  {$IFDEF EAGLSERVER}
  eSession,
  {$ELSE}
  CbpTrace,
  {$ENDIF}
  Paramsoc,
  uWaBDSqlServer,
  galSystem;

Const
  BOBHORSCOMMUN = ['M','P','V'] ;
  BOBPARTOUT = ['A'] ;

function ChargeListeBOB (GestionBOBCommun : Boolean) : TStringList ;
function IsBobCommun (BobName : String) : Boolean ;
function IsBobPartout (BobName : String) : Boolean ;
function IsBobMenu (BobName : String) : Boolean ;
function GetCWSInstallDir : String ;
function BobIsAlreadyIntegred (BOBList : TStringlist ; BobName : String ; GestionBOBCommun : Boolean ) : Boolean ;
function GetListeBobMenu(PathDirBob : string) : TOB;
{$IFDEF KPMG}
procedure RendreConforme(FileName : String; var ErrMsg : String; socref : integer) ;
{$ENDIF KPMG}

{$IFDEF EAGLSERVER}
procedure TraiteDir(BOBList : TStringList ; Directory : String; ResultTob : TOB; GestionBOBCommun : Boolean ) ;
{$ELSE}
procedure TraiteDir(BOBList : TStringList ; Directory : String; ResultTob : TOB; GestionBOBCommun : Boolean ; BOBLabel : THLABEL ) ;
{$ENDIF}
function  TraiteFile (FileName : String ; var LogMessage : String ; SortSiTrouve : Boolean ; GestionBobCommun : Boolean ; DB : String = '' ) : Integer ;
procedure cgiDDWriteLn (S : String) ;
{$IFDEF EAGLSERVER}
function  TraiteScript (NomScript, NomLog, DbName : string) : TOB;
{$ENDIF}
function GetCurDb : String ;
function GetDefaultDb : String ;
function GetNumSocref : integer ;



implementation

procedure cgiDDWriteLn (S : String) ;
begin
{$IFDEF DEBUG}
  {$IFDEF EAGLSERVER}
  ddWriteLn(S);
  {$ENDIF}
{$ENDIF}
end;


const
   Default = 'DB000000' ;
   DefaultSoc = 850 ;

function GetCurDb : String ;
{$IFDEF EAGLSERVER}
var
  LaSession   :TISession;
{$ENDIF}
begin
   {$IFDEF EAGLSERVER}
   Result := Default ;
   LaSession  := LookupCurrentSession;
   if LaSession <> nil then
   begin
//     Result := 'DB'+LaSession.NoDossier ;
     Result := LaSession.DBName ;
   end
   else
      exit ;
   {$ELSE}
     Result := V_PGI.DBName ;
   {$ENDIF}
   if Result = '' then Result := Default ;
end;

function GetDefaultDb : String ;
{$IFDEF EAGLSERVER}
var
  LaSession   :TISession;
{$ENDIF}
begin
   {$IFDEF EAGLSERVER}
   Result := Default ;
   LaSession  := LookupCurrentSession;
   if LaSession <> nil then
   begin
     Result := LaSession.DefaultSectionDBName ;
   end
   else
      exit ;
   {$ELSE}
   Result := V_PGI.DefaultSectionDBName ;
   {$ENDIF}
   if Result = '' then Result := Default ;
end;

function GetNumSocref : integer ;
{$IFDEF EAGLSERVER}
var
  LaSession   :TISession;
{$ENDIF}
begin
   {$IFDEF EAGLSERVER}
   Result := DefaultSoc ;
   LaSession  := LookupCurrentSession;
   if LaSession <> nil then
   begin
     Result := LaSession.NumVersionSoc ;
   end
   else
      exit ;
   {$ELSE}
   Result := V_PGI.NumVersionSoc ;
   {$ENDIF}
   if Result = 0 then Result := DefaultSoc ;
end;

function BobIsAlreadyIntegred (BOBList : TStringlist ; BobName : String ; GestionBOBCommun : Boolean ) : Boolean ;
var
  sfile : String ;
  bPartout : Boolean ;
  CurDB, DefaultDB : String ;
begin
   sfile := Copy(BobName,1,pos('.',BobName)-1);
   CurDB := GetCurDb ;
   DefaultDB := GetDefaultDb ;

   if GestionBOBCommun
   and ( CurDb <> DefaultDB ) then
   begin
     bPartout := IsBobPartout(BobName);
     if not bPartout then
     begin
       if not IsBobCommun(BobName) then
       begin
          Result := BOBList.IndexOf(UpperCase(sfile)) >=0;
          cgiDDWriteLn('Recherche '+UpperCase(sfile)+'='+booltostr(Result,True));
       end
       else
       begin
          Result := BOBList.IndexOf('DB0.'+UpperCase(sfile)) >=0 ;
          cgiDDWriteLn('Recherche '+'DB0.'+UpperCase(sfile)+'='+booltostr(Result,True));
       end;
     end
     else
     begin
          Result := ( BOBList.IndexOf(UpperCase(sfile)) >=0 )
          and ( BOBList.IndexOf('DB0.'+UpperCase(sfile)) >=0 ) ;
          cgiDDWriteLn('Recherche Partout '+UpperCase(sfile)+'='+booltostr(Result,True));
     end;
   end
   else
       Result := BOBList.IndexOf(UpperCase(sfile)) >=0 ;
end;


function ChargeListeBOB (GestionBOBCommun : Boolean) : TStringList ;
var
   T : TOb ;
   i : integer ;
  CurDB, DefaultDB : String ;
begin
   Result := TStringList.Create ;
   CurDB := GetCurDb ;
   DefaultDB := GetDefaultDb ;

   T := TOB.Create('test',nil,0);
   try
   T.LoadDetailFromSQL('SELECT YB_BOBNAME FROM YMYBOBS');
   except
      on e:exception do
      begin
{$IFNDEF EAGLSERVER}
        Trace.TraceInformation('UBOBUTILES', 'Erreur LoadDetailFromSQL : '+e.message);
{$ELSE}
        ddwriteln('Erreur LoadDetailFromSQL : '+e.message);
{$ENDIF}
      end;
   end;
   try
     for i:=0 to T.Detail.Count-1 do
        Result.Add(UpperCase(T.Detail[i].GetString('YB_BOBNAME'))) ;
   except
      on e:exception do
      begin
{$IFNDEF EAGLSERVER}
         Trace.TraceInformation('UBOBUTILES', 'Erreur sur T.Detail : '+e.message);
{$ELSE}
         ddwriteln('Erreur sur T.Detail : '+e.message);
{$ENDIF}
      end;
   end;
   FreeAndNil(T);
   if GestionBOBCommun
   and ( CurDB <> DefaultDB ) then
   begin
       T := TOB.Create('test',nil,0);
       T.LoadDetailFromSQL('@@SELECT YB_BOBNAME FROM '+DefaultDB+'.dbo.YMYBOBS');
       for i:=0 to T.Detail.Count-1 do
          Result.Add('DB0.'+UpperCase(T.Detail[i].GetString('YB_BOBNAME'))) ;
       FreeAndNil(T);
   end;
end;

function IsBobCommun (BobName : String) : Boolean ;
var
   TypeBob : char ;
begin
   if copy(ExtractFileName(BobName), 9, 1) <> '' then
      TypeBob := copy(ExtractFileName(BobName), 9, 1)[1]
   else
      TypeBob := 'W' ;
   Result := not ( TypeBob in BOBHORSCOMMUN ) ;
end;

function IsBobPartout (BobName : String) : Boolean ;
var
   TypeBob : char ;
begin
   if copy(ExtractFileName(BobName), 9, 1) <> '' then
      TypeBob := copy(ExtractFileName(BobName), 9, 1)[1]
   else
      TypeBob := 'W' ;
   Result := ( TypeBob in BOBPARTOUT ) ;
end;

function IsBobMenu (BobName : String) : Boolean ;
var
   TypeBob : String ;
begin
   TypeBob := copy(ExtractFileName(BobName), 9, 1) ;
   Result := TypeBob = 'M' ;
end;


{function CanMoveBOB : Boolean ;
var
   CurSection : String ;
   HIni : THInifile ;
   L : hTStringList ;
   i, CountSections : Integer ;
begin
Result := True ;
try
    HIni := THInifile.create('CEGIDPGI.INI');
    cgiDDWriteLn('Utilisation du fichier ini : '+HIni.FileName);

    L := hTStringList.Create ;
    CountSections := 0 ;
    Hini.ReadSections(L);
    for i:=0 to L.Count-1 do
    begin
       CurSection := L.Strings[i] ;
       If not SameText(CurSection,'Reference')
       and not ( pos('##',CurSection) = 1 ) then
          Inc(CountSections);
    end;
    L.Free ;
    hini.Free ;
    Result := CountSections = 1 ;

    cgiDDWriteLn('Renomage des BOB : '+booltoStr(Result,True)+', '+inttostr(CountSections)+' dossiers.');
except
  on e:exception do
    cgiDDWriteLn('Erreur CanMoveBOB : '+e.message);
end;
end; }



function GetWindowsDrive: String;
const MaxWinPathLen = MAX_PATH + 1;
var I : LongWord;
begin
  SetLength(Result, MaxWinPathLen);
  I := GetWindowsDirectory(PChar(Result), MaxWinPathLen);
  if I > 0 then
    SetLength(Result, I)
  else
    Result := '';
  Result := IncludeTrailingPathDelimiter(ExtractFileDrive(Result)) ;
end;

function GetCWSInstallDir : String ;
var
   r : TRegistry ;
begin
   Result := GetWindowsDrive + 'CWS\' ;
   r := TRegistry.Create ;
   r.RootKey := HKEY_LOCAL_MACHINE ;
   if r.OpenKeyReadOnly('SOFTWARE\CEGID_RM\Pgiservice') then
   begin
      if r.ValueExists('Path') then
         Result := IncludeTrailingPathDelimiter(r.ReadString('Path')) ;
      r.CloseKey ;
   end;
   r.Free ;
end;

function TraiteFile (FileName : String ; var LogMessage : String ; SortSiTrouve : Boolean ; GestionBobCommun : Boolean ; DB : String = '' ) : Integer ;
var
  CurDB, DefaultDB : String ;
//  socref : integer ;
begin
  Result := 0 ;
  LogMessage := '' ;
//  D := DummyObject.Create ;
  try
      cgiDDWriteLn('Traitement du fichier : '+FileName);
      CurDB := GetCurDb ;
      if DB = '' then
        DefaultDB := GetDefaultDb
      else
        DefaultDB := DB;
//      socref := GetNumSocref ;
      cgiDDWriteLn('DBs : Default='+DefaultDB+', Current='+CurDB);
      if Result = 0 then
      begin
          if ( GestionBobCommun and IsBobCommun(FileName) and ( CurDB <> DefaultDB ) )
          OR ( (DB<>'') and IsBobMenu(FileName) )   then
          begin
             try
               ExecuteSQL('@@USE '+DefaultDB) ;
             except
               on e:Exception do
               begin
                  Result := -201 ;
                  LogMessage := 'Erreur sur utilisation de '+DefaultDB+' : '+ E.Message ;
               end;
             end;
          end;
          if Result = 0 then
            try
{$IFDEF KPMG}
//               rendreconforme(FileName,LogMessage,socref);
{$ENDIF KPMG}
               Result := AglIntegreBob(FileName, False, SortSiTrouve, nil,nil);//D.CallBack) ;
            except
               on e:exception do
               begin
                  Result := -999 ;
                  LogMessage := 'Except : '+E.Message ;
               end;
            end;
          if GestionBobCommun
          and IsBobCommun(FileName)
          and ( CurDB <> DefaultDB ) then
          begin
             try
             ExecuteSQL('@@USE '+CurDB) ;
             except
             on e: exception do
             begin
                  Result := -201 ;
                  LogMessage := 'Erreur sur utilisation de '+CurDB+' : '+ E.Message ;
             end;
             end;
          end;
      end;

      if Result in [0,1] then // OK on gére le cas du partout : en plus la DBX .
      begin
        if GestionBobCommun
        and IsBobPartout(FileName)
        and ( CurDB <> DefaultDB ) then
        begin
          try
{$IFDEF KPMG}
//             rendreconforme(FileName,LogMessage,socref);
{$ENDIF KPMG}
             Result := AglIntegreBob(FileName, False, SortSiTrouve, nil,nil);//D.CallBack) ;
          except
             on e:exception do
             begin
                Result := -999 ;
                LogMessage := 'Except : '+E.Message ;
             end;
          end;
        end;
      end;
      case Result of
        0: LogMessage := 'OK' ;
        1: LogMessage := 'Intégration déjà effectuée' ;
        -1: LogMessage := 'Erreur d''écriture dans la table YMYBOBS' ;
        -2: LogMessage := 'Erreur d''intégration dans la fonction AglImportBob' ;
        -3: LogMessage := 'Erreur de lecture du fichier BOB.' ;
        -4: LogMessage := 'Erreur inconnue.' ;
        -5: LogMessage := 'Fichier non trouvé.' ;
        -98: LogMessage := 'Cette BOB de Maj de structure n''est pas exécutée depuis PgiMajVer.exe' ;
        -99: LogMessage := 'Cette BOB n''est pas prévue pour cette version de SocRef' ;
        -100: LogMessage := 'Bob sans contrôle' ;
        -200 : LogMessage := 'Erreur sur recupération de la session.' ;
        -999 : ;
      else
        LogMessage := 'Erreur inconnue ('+IntToStr(Result)+') : '+LogMessage; ;
      end;
      //LogMessage := LogMessage + '('+D.Text+')' ;
  except
     on e:exception do
     begin
        Result := -998 ;
        LogMessage := ' Except : '+E.Message + '('+LogMessage+')' ;
     end;
  end;
//  D.Free ;
end;

{$IFDEF EAGLSERVER}
procedure TraiteDir(BOBList : TStringList ; Directory : String; ResultTob : TOB; GestionBOBCommun : Boolean ) ;
{$ELSE}
procedure TraiteDir(BOBList : TStringList ; Directory : String; ResultTob : TOB; GestionBOBCommun : Boolean ; BOBLabel : THLABEL ) ;
{$ENDIF}
const CRLF = #13#10 ;
var
  ResultString, STmp, BobFile : String ;
  sr : TSearchRec;
  iRes, iCount : Integer ;
  MAJMenu : Boolean ;
  function CanIntegre : boolean ;
  var
    sext : string ;
  begin
     Result := False ;
     if (sr.attr and faDirectory) = faDirectory then exit ;
     if (copy(sr.Name,1,1) = '_') then exit ;
     sext := ExtractFileExt(sr.Name) ;
     Result := SameText(sext,'.bob') or SameText(sext,'.boz')  ;
  end;
  procedure SetResults (S : String );
  begin
       if ResultTob.FieldExists('RESULTS') then
          ResultTob.SetString('RESULTS',ResultTob.GetString('RESULTS')+CRLF+S)
       else
          ResultTob.AddChampSupValeur('RESULTS',S);
  end;
  procedure SetResult (I : Integer );
  begin
     if ResultTob.FieldExists('COUNT') then
        ResultTob.SetInteger('COUNT',ResultTob.GetInteger('COUNT')+I)
     else
        ResultTob.AddChampSupValeur('COUNT',I);
  end;
  function GetResultB : Boolean ;
  begin
      Result := False ;
      if ResultTob.FieldExists('MAJMENU') then
         Result := ResultTob.GetBoolean('MAJMENU')
      else
         ResultTob.AddChampSupValeur('MAJMENU',False);
  end;

  procedure SetResultB (B : Boolean );
  begin
     if ResultTob.FieldExists('MAJMENU') then
        ResultTob.SetBoolean('MAJMENU',B)
     else
        ResultTob.AddChampSupValeur('MAJMENU',B);
  end;
begin
  ResultString := '' ;
  {$IFDEF EAGLSERVER}
  // LG - 13/02/2008 - la fct booltostr provoque une fuite memoire. Passage ds une directive de compile
  cgiDDWriteLn('Traitement du repertoire :' + Directory+ ' GestionBOBCommun : '+booltostr(GestionBOBCommun,True));
  {$ENDIF}
  iCount := 0 ;

  MajMenu := GetResultB ;

  if DirectoryExists(Directory) then
  begin
    if sysutils.FindFirst(Directory+'*.bo*',faAnyFile - faDirectory,sr)=0 then
    begin
      repeat
         if CanIntegre then
         begin
            if ResultString <> '' then
               ResultString := ResultString+CRLF;
            ResultString := ResultString+sr.Name + ' : ';
            BobFile := Sr.Name ;
            if not BobIsAlreadyIntegred(BOBList,BobFile,GestionBOBCommun) then
            begin
              // Traitement de la bob
              {$IFNDEF EAGLSERVER}
              BOBLabel.Caption := 'Intégration de : '+ BobFile ;
              BOBLabel.Refresh ;
              {$ENDIF}
              iRes := TraiteFile(Directory+BobFile,STmp,True,GestionBOBCommun);
              ResultString := ResultString+STmp ;
              if iRes = 0 then
              begin
                 if not MAJMenu then
                    MAJMenu := IsBobMenu(BobFile) ;
                 Inc(iCount) ;
              end;
            end
            else
               ResultString := ResultString + 'déja intégré.' ;
         end;
      until sysutils.FindNext(sr)<>0;
      sysutils.FindClose(sr);
      // Maj Tob
      SetResults(ResultString);
      SetResult(iCount);
    end
    else
    begin
      // Maj Tob
      SetResults(Directory+' : Aucun BOB Trouvé');
      SetResult(0);
    end;
  end
  else
  begin
      // Maj Tob
     SetResults('Chemin BOB introuvable : '+Directory);
     SetResult(0);
  end;
  SetResultB(MAJMenu);
end;

{$IFDEF EAGLSERVER}
function  TraiteScript (NomScript, NomLog, DbName : string) : TOB;
var
  WABD        :cWABDSQLServer;
  CheminReseau, CheminLocal, DatPath, DatPathLoc : String ;
  //dbname      :string;
  LigneCmd    :string;
begin
  {
  LaSession  := LookupCurrentSession;
  if LaSession = nil then
  begin
    cgiDDWriteLn('#### Session Nil');
    result.AddChampSupValeur('ERROR', 'EXECSCRIPTSQL : Session = Nil');
    exit;
  end;

    // Parsing du Paramètre
    L := TStringList.Create ;
    L.Delimiter := ';' ;
    L.DelimitedText := Param ;
  }
  result := TOB.Create('le result', nil, -1);
  try
    if ( NomScript = '' ) then
    begin
      cgiDDWriteLn('#### Aucun script demande');
      result.AddChampSupValeur('ERROR', 'EXECSCRIPTSQL : Aucun script demande');
      exit;
    end;

    DatPath := GetParamsocDpSecur('SO_ENVPATHDAT', '');
    if DatPath='' then
    begin
      cgiDDWriteLn('#### PathDat vide');
      result.AddChampSupValeur('ERROR', 'EXECSCRIPTSQL : PathDat vide');
      exit;
    end;

    DatPathLoc := GetParamsocDpSecur('SO_ENVPATHDATLOC', '');
    if DatPathLoc='' then
    begin
      cgiDDWriteLn('#### PathDatLoc vide');
      result.AddChampSupValeur('ERROR', 'EXECSCRIPTSQL : PathDatLoc vide');
      exit;
    end;

    //Chemin Reseau vu depuis le serveur CWAS
    CheminReseau := DatPath + '\SOCREF\';
    //CheminLocal vu par le serveur SQL
    CheminLocal := DatPathLoc + '\SOCREF\';
    if Not FileExists(CheminReseau+NomScript) then
    begin
      cgiDDWriteLn('#### Script '+CheminReseau+NomScript+' absent');
      result.AddChampSupValeur('ERROR', 'EXECSCRIPTSQL : Script '+CheminReseau+NomScript+' absent');
      exit;
    end;

    WABD := cWABDSQLServer.Create;
    if WABD = nil then
    begin
      cgiDDWriteLn('#### WABD = Nil');
      result.AddChampSupValeur('ERROR', 'EXECSCRIPTSQL : WABD = Nil');
      exit;
    end;

    result.AddChampSupValeur('RESULTS', 'Serveur:'+WABD.MasterServer+'; DB :'+DbName+'; User:'+WABD.MasterUser
                         +'; Password:'+WABD.MasterPass+'; Chemin script:'+CheminReseau+NomScript+'; Log:'+CheminReseau+NomLog);
    cgiDDWriteLn('Serveur:'+WABD.MasterServer+'; DB :'+dbname+'; User:'+WABD.MasterUser);
    LigneCmd := 'osql -S'+WABD.MasterServer //.LeServeurSql
             +' -d'+DbName
             +' -U'+WABD.MasterUser
             +' -P'+WABD.MasterPass
             +' -i"'+CheminLocal+NomScript+'"'
             +' -o"'+CheminLocal+NomLog+'"'
             +' -u'
             +' -n';
    try
      ExecCmdShell(LigneCmd);
      AvertirCacheServer('*');
    finally
      FreeAndNil(WABD);
    end;
  finally

  end;

end;
{$ENDIF}

// Alimente le tableau en fonction des fichiers BOBs de type M (Menu) trouvé à l'emplacement
// indiqué : PathDirBob.
function GetListeBobMenu(PathDirBob : string) : TOB;
var
  PathFileBob,NameBob : string;
  InfoDir,InfoFic : TSearchRec;
  TbDet : TOB;
begin
  Result := TOB.Create('GetListeBobMenu',nil,-1);
  TbDet := nil;

  if FindFirst(PathDirBob+'*.*',faAnyFile,InfoDir) = 0 Then
  begin
    repeat
      //if (InfoDir.Attr = faDirectory) and (InfoDir.Name <> '.') and (InfoDir.Name <> '..') then
      if ((InfoDir.Attr and faDirectory)<>0) and (InfoDir.Name <> '.') and (InfoDir.Name <> '..') then
      begin
        // Lecture des fichiers bob contenus dans le sous répertoire InfoDir.Name
        PathFileBob := IncludeTrailingPathDelimiter(PathDirBob+InfoDir.Name);
        if FindFirst(PathFileBob+'*.BOB',faAnyFile,InfoFic) = 0 Then
        begin
          repeat
            if IsBobMenu(InfoFic.Name) then
            begin
              // Nom du fichier sans extension
              NameBob := copy(InfoFic.Name,1,pos('.',InfoFic.Name)-1);
              TbDet := TOB.Create('Detail',Result,-1);
              TbDet.AddChampSupValeur('BOBNAME' , NameBob);
              TbDet.AddChampSupValeur('BOBPATH' , PathFileBob + InfoFic.Name);
            end;
          until FindNext(InfoFic) <> 0;
          SysUtils.FindClose(InfoFic);
        end;
      end;
    until FindNext(InfoDir) <> 0;
    SysUtils.FindClose(InfoDir);
  end;

  if not Assigned(TbDet) then
    Result.AddChampSupValeur('RESULT','NOK')
end;
type TBobConforme = (bcErr, bcWarn, bcOk) ;

function IsConforme (T : Tob ; FN : String ; var ErrMsg : String; socref : integer ) : TBobConforme ;
const
  BAD_BOB = 1 ;
  BAD_TYPE = 2 ;
  BAD_SOCREF = 4 ;
  BAD_LIBELLE = 8 ;
var
  control : set of byte ;
  function GetTobInfoString (Field : String; Default : String = '' ) : String ;
  begin
     Result := Default ;
     If T.FieldExists(Field) then
        Result := T.GetString(Field) ;
  end;
  function GetTobInfoInt (Field : String; Default : integer = 0 ) : integer ;
  begin
     Result := Default ;
     If T.FieldExists(Field) then
        Result := T.GetValue(Field) ;
  end;
  function DumpControl : TBobConforme ;
  begin
     Result := bcOk ;
     if control = [] then
       ErrMsg :=  'La BOB est conforme.'
     else
     begin
       ErrMsg := '' ;
       if BAD_BOB in control then
       begin
         ErrMsg := ErrMsg + #10 +  'Le Fichier n''est pas une BOB correcte.';
         Result := bcErr ;
       end;
       if BAD_TYPE in control then
       begin
         ErrMsg :=  ErrMsg +  #10 + 'Cette BOB n''est pas intégrable en automatique (BOB sans contrôle).';
         if Result = bcOk then Result := bcWarn ;
       end;
       if BAD_SOCREF in control then
       begin
         ErrMsg := ErrMsg +  #10 + 'Cette BOB n''est pas intégrable sur le numéro de SOCREF en cours ('+IntToStr(socref)+').';
         if Result = bcOk then Result := bcWarn ;
       end;
       if BAD_LIBELLE in control then
       begin
         ErrMsg := ErrMsg +  #10 + 'Attention : Cette BOB a été renommée. Ancien nom = '+GetTobInfoString('BOBNAME');
         if Result = bcOk then Result := bcWarn ;
       end;
     end;
  end;
begin
  control := [] ;
  if GetTobInfoString('BOBNAME') = '' then
     control := [BAD_BOB] ;

  if not ( BAD_BOB in control ) then
  begin
    if GetTobInfoInt('BOBTYPE') <> 1 then
       control := control + [BAD_TYPE] ;

    if ( GetTobInfoInt('BOBSOCREFDEPART') > socref )
    or ( GetTobInfoInt('BOBSOCREFFINALE') < socref )  then
       control := control + [BAD_SOCREF] ;

    if pos(GetTobInfoString('BOBNAME'),FN) = 0 then
       control := control + [BAD_LIBELLE] ;

  end;
  Result := DumpControl;
end;


{$IFDEF KPMG}
procedure RendreConforme(FileName : String;var ErrMsg : String; socref : integer) ;
var
  T : TOB ;
  Bname : String ;
  R : TBobConforme ;
begin
  ErrMsg := '' ;
  T := TOB.Create('la bob '+ExtractFileName(FileName) ,nil,-1);
  try
      if TOBLoadFromBinFile(FileName,nil,T,false) then
      begin
        R := IsConforme(T,FileName,ErrMsg,socref) ;
        if ( R  = bcWarn ) then
        begin
          cgiDDWriteLn('     Non Conforme !!! '+ErrMsg);
          T.SetInteger('BOBTYPE',1);
          T.SetInteger('BOBSOCREFDEPART',socref);
          T.SetInteger('BOBSOCREFFINALE',socref);
          T.SetInteger('BOBNUMSOCREF',socref);
          T.SetString('BOBCREATEUR',T.GetString('BOBCREATEUR')+'-OK');
          Bname := ExtractFileName(FileName) ;
          Bname := UpperCase(copy(Bname,1,pos('.',BName)-1));
          T.SetString('BOBNAME',Bname);
          T.SaveToBinFile(FileName,False,True,True,True);
        end ;
        if ( R = bcErr ) then
          cgiDDWriteLn('    '+ErrMsg);
      end
      else
        ErrMsg := '    Impossible de lire la BOB "'+FileName ;
  except
     on e:exception do
        ErrMsg := '    Except Rendre Conforme : '+e.Message ;
  end;
  T.Free ;
end;
{$ENDIF KPMG}



end.



