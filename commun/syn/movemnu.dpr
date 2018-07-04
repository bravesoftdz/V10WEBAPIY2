program movemnu;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  inifiles,
  windows,
  registry,
  ShlObj,
  StrUtils,
  Classes;

var
  sr : TSearchRec;
  s, sOldPath, sNewPath, sFile, sOldFile, sNewFile : string ;
  reg : TRegistry ;
  F1,F2,FTrace : TextFile ;
  FilePath: array[0..MAX_PATH] of Char;
  FileOK : boolean ;
  ts : TStringList ;
  i : integer ;
const
  CSIDL_PROGRAM_FILES = $0026;

begin
  { TODO -oUser -cConsole Main : placez le code ici }

// Initialisations
// Répertoires
sOldPath :='c:\pgi00\app\' ;
// lecture du new path
reg :=TRegistry.Create;
reg.RootKey:=HKEY_LOCAL_MACHINE;
reg.OpenKey('SOFTWARE\CEGID\CEGID Expert',False); //pas le créer s'il n'existe pas
sNewPath:=reg.ReadString('INSTALLDIR');
reg.Free;
if sNewPath='' then
   begin
   ShlObj.SHGetSpecialFolderPath(0, FilePath, CSIDL_PROGRAM_FILES, False) ;
   sNewPath:=FilePath+'\CEGID\Cegid Expert\APP\' ;
   end else
   begin
   if RightStr(sNewPath,1)<>'\' then sNewPath:=sNewPath+'\' ;
   if RightStr(sNewPath,4)<>'APP\' then sNewPath:=sNewPath+'APP\' ;
   end;

if not DirectoryExists(sNewPath) then
   ForceDirectories(sNewPath);

AssignFile(FTrace, sNewPath+'movemnu.log');
Rewrite(FTrace);

writeln(FTrace,'MoveMNU version 0.01 (c) Cegid 2007') ;
// Fichier d'exclusions
ts:=TStringList.Create ;
if paramcount>0 then
   begin
   if FileExists(paramstr(1)) then
      begin
      ts.LoadFromFile(paramstr(1));
      writeln(FTrace,'Fichiers d''exclusions : '+paramstr(1)) ;
      end;
   end ;

writeln(FTrace,'Répertoire cible : '+sNewPath) ;
writeln(FTrace,'');
writeln(FTrace,'Liste des .mnu pris en compte :');
writeln(FTrace,'-------------------------------');
// Lecture des mnu
if FindFirst(sOldPath+'*.mnu',faAnyFile,sr)=0 then
   begin
   repeat
   sFile:=sr.Name ;
   if uppercase(rightstr(sfile,7)) <> 'MNU_OLD' then
   begin
     sOldFile:=sOldPath+sFile ;
     sNewFile:=sNewPath+sFile ;

     // Test d'exclusion
     Delete(sFile,Length(sFile)-3,4);
     FileOK:=True ;
     if ts.count>0 then
        begin
        for i:=0 to ts.Count-1 do
           begin
           if Uppercase(sFile)= Uppercase(ts.Strings[i]) then
              begin
              FileOK:=False ;
              break ;
              end ;
           end ;
        end ;

  {
     //Test si présence de NOMHAL=Plaquettes PGI PubliFi
     AssignFile(F1, sOldFile);
     Reset(F1);
     FileOK:=False ;
     while not Eof(F1) do
        begin
        Readln(F1, s);
        if pos('PubliFi',s)>0 then
           begin
           FileOK:=True ;
           break ;
           end ;
        end;
     CloseFile(F1);
  }

     if FileOK then // prévoir une liste de conditions en ligne de commande
        begin
        // Traitement d'un fichier publifi.mnu
        //Copie du fichier avec modification de PATH
        AssignFile(F1, sOldFile);
        Reset(F1);
        AssignFile(F2, sNewFile);
        Rewrite(F2);
        while not Eof(F1) do
           begin
           Readln(F1, s);
           if s='PATH=' then s:=s+sOldPath ;
           Writeln(F2, s);
           end;
        CloseFile(F2);
        CloseFile(F1);
        RenameFile(sOldFile, uppercase(sOldFile+'_old'));
        writeln(sFile) ;
        end ;
     end
     else writeln('Attention ' + sfile);
   until (FindNext(sr)<>0) ;
   sysutils.FindClose(sr);
   ts.Free ;
   end ;
   CloseFile(FTrace);
end.
