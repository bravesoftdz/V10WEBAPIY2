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
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  // cbpIniFiles,
  sysutils,
  Windows;

Const
  BOBHORSCOMMUN = ['M','P','V'] ;
  BOBPARTOUT = ['A'] ;

function ChargeListeBOB (GestionBOBCommun : Boolean) : TStringList ;
function IsBobCommun (BobName : String) : Boolean ;
function IsBobPartout (BobName : String) : Boolean ;
function IsBobMenu (BobName : String) : Boolean ;
function GetCWSInstallDir : String ;
function BobIsAlreadyIntegred (BOBList : TStringlist ; BobName : String ; GestionBOBCommun : Boolean ) : Boolean ;

{$IFDEF EAGLSERVER}
procedure TraiteDir(BOBList : TStringList ; Directory : String; ResultTob : TOB; GestionBOBCommun : Boolean ) ;
{$ELSE}
procedure TraiteDir(BOBList : TStringList ; Directory : String; ResultTob : TOB; GestionBOBCommun : Boolean ; BOBLabel : THLABEL ) ;
{$ENDIF}
function  TraiteFile (FileName : String ; var LogMessage : String ; SortSiTrouve : Boolean ; GestionBobCommun : Boolean ) : Integer ;
procedure cgiDDWriteLn (S : String) ;

implementation


procedure cgiDDWriteLn (S : String) ;
begin
  {$IFDEF DEBUG}
  {$IFDEF EAGLSERVER}
  WriteLn(S);
  {$ENDIF}
  {$ENDIF}
end;

function BobIsAlreadyIntegred (BOBList : TStringlist ; BobName : String ; GestionBOBCommun : Boolean ) : Boolean ;
var
   sfile : String ;
   bPartout : Boolean ;
begin
   sfile := Copy(BobName,1,pos('.',BobName)-1);
   if GestionBOBCommun
   and ( V_PGI.NoDossier <> '000000' ) then
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
begin
   Result := TStringList.Create ;
   T := TOB.Create('test',nil,0);
   T.LoadDetailFromSQL('SELECT YB_BOBNAME FROM YMYBOBS');
   for i:=0 to T.Detail.Count-1 do
      Result.Add(UpperCase(T.Detail[i].GetString('YB_BOBNAME'))) ;
   FreeAndNil(T);
   if GestionBOBCommun
   and ( V_PGI.NoDossier <> '000000' ) then
   begin
       T := TOB.Create('test',nil,0);
       T.LoadDetailFromSQL('@@SELECT YB_BOBNAME FROM DB000000.dbo.YMYBOBS');
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
   if r.OpenKeyReadOnly('SOFTWARE\CEGDI_RM\Pgiservice') then
   begin
      if r.ValueExists('Path') then
         Result := IncludeTrailingPathDelimiter(r.ReadString('Path')) ;
      r.CloseKey ;
   end;
   r.Free ;
end;

function TraiteFile (FileName : String ; var LogMessage : String ; SortSiTrouve : Boolean ; GestionBobCommun : Boolean ) : Integer ;
//var
//  D : DummyObject ;
begin
  LogMessage := '' ;
//  D := DummyObject.Create ;
  try
      cgiDDWriteLn('Traitement du fichier : '+FileName);
      if GestionBobCommun
      and IsBobCommun(FileName)
      and ( V_PGI.NoDossier <> '000000' )  then
      begin
         cgiDDWriteLn('Utilisation de la DB000000');
         ExecuteSQL('@@USE DB000000') ;
      end;
      try
         Result := AglIntegreBob(FileName, False, SortSiTrouve, nil,nil);//D.CallBack) ;
      except
         on e:exception do
         begin
            Result := -999 ;
            LogMessage := ' Except : '+E.Message ;
         end;
      end;
      if GestionBobCommun
      and IsBobCommun(FileName)
      and ( V_PGI.NoDossier <> '000000' ) then
      begin
         cgiDDWriteLn('Utilisation de la DB'+V_PGI.NoDossier);
         ExecuteSQL('@@USE DB'+V_PGI.NoDossier) ;
      end;
      if Result in [0,1] then // OK on gére le cas du partout : en plus la DBX .
      begin
        if GestionBobCommun
        and IsBobPartout(FileName)
        and ( V_PGI.NoDossier <> '000000' ) then
        begin
          try
             Result := AglIntegreBob(FileName, False, SortSiTrouve, nil,nil);//D.CallBack) ;
          except
             on e:exception do
             begin
                Result := -999 ;
                LogMessage := ' Except : '+E.Message ;
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
        -999 : ;
      else
        LogMessage := 'Erreur inconnue ('+IntToStr(Result)+')'; ;
      end;
      //LogMessage := LogMessage + '('+D.Text+')' ;
  except
     on e:exception do
     begin
        Result := -998 ;
        LogMessage := ' Except : '+E.Message ;
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
     if ResultTob.FieldExists('RESULT') then
        ResultTob.SetInteger('RESULT',ResultTob.GetInteger('RESULT')+I)
     else
        ResultTob.AddChampSupValeur('RESULT',I);
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
  cgiDDWriteLn('Traitement du repertoire :' + Directory+ ' GestionBOBCommun : '+booltostr(GestionBOBCommun,True));
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

{procedure ChargeSubDirs (Chemin : String ; Var List : TStringList);
var
   sd : TSearchRec;
begin
  List.Clear ;
  if sysutils.FindFirst(Chemin+'*.*',faDirectory,sd)=0 then
  begin
    repeat
       if ( sd.Name <> '.' )
       and ( sd.Name <> '..' )
       and ( sd.Name <> 'DBX' )
       and ( sd.Name <> 'DDE' ) then
          List.Add(sd.Name) ;
    until sysutils.FindNext(sd)<>0;
    sysutils.FindClose(sd);
  end;
end;
}

end.

