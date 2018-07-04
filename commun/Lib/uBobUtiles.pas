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
uses MajHalleyUtil;

procedure  BeforeIntegreOneBob ;
begin
//
end;

procedure  BeforeIntegrationBob (BobFile : string);
var NamFic : string;
begin
  NamFic := ExtractFileName(BOBFile);
  if NamFic='XXCEGIDV9ED10.BOB' then
  begin
    ExecuteSql('DELETE FROM PUBLICOTIS WHERE PUO_PREDEFINI="CEG" AND PUO_TYPEAFFECT="MSA" AND PUO_RUBRIQUE="Z33" AND PUO_UTILSEGMENT="019"');
  end else if NamFic = 'XXCEGIDV9ED13P5-001.BOB' then
  begin
    ExecuteSQL('DELETE FROM MOTIFSORTIEPAY WHERE PMS_CODE IN ("70", "71")');
    ExecuteSQL('DELETE FROM DADSLEXIQUE WHERE PDL_DADSSEGMENT="S40.G10.05.013.003" OR '+
               'PDL_DADSSEGMENT="S40.G15.00.002" OR PDL_DADSSEGMENT="S40.G25.00.021" OR '+
               'PDL_DADSSEGMENT="S40.G30.04.033.001" OR PDL_DADSSEGMENT="S40.G30.11.001.001" OR '+
               'PDL_DADSSEGMENT="S40.G30.11.001.002" OR PDL_DADSSEGMENT="S40.G30.30.001.001" OR '+
               'PDL_DADSSEGMENT="S40.G30.30.002.001" OR PDL_DADSSEGMENT="S48.G55.00.001.001" OR '+
               'PDL_DADSSEGMENT="S60.G05.47.001.001" OR PDL_DADSSEGMENT="S60.G05.47.002.001" OR '+
               'PDL_DADSSEGMENT="S60.G05.47.003.001" OR PDL_DADSSEGMENT="S60.G05.47.004.001" OR '+
               'PDL_DADSSEGMENT="S60.G05.47.004.002" OR PDL_DADSSEGMENT="S60.G05.47.005.001" OR '+
               'PDL_DADSSEGMENT="S65.G55.10.001" OR PDL_DADSSEGMENT="S65.G55.10.002.001" OR '+
               'PDL_DADSSEGMENT="S65.G55.10.003.001" OR PDL_DADSSEGMENT="S70.G10.00.001.001" OR '+
               'PDL_DADSSEGMENT="S70.G10.00.001.002" OR PDL_DADSSEGMENT="S70.G10.00.002" OR '+
               'PDL_DADSSEGMENT="S70.G10.00.005"');
  end else if NamFic = 'XXCEGIDV9ED14-002.BOB' then
  begin
    SupprimeEtat('E', 'IK3', '154');
    SupprimeEtat('E', 'IK3', '155');
    SupprimeEtat('E', 'IK4', '154');
    SupprimeEtat('E', 'IK4', '155');
	end else if NamFic = 'XXCEGIDV9ED14P2-001' then
  begin
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "ILISTIMOVUETVTS1"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "ILISTIMOVUETVTS2"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "ILISTVTS1"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "IMOTVTSENLIST"');
	end else if NamFic = 'XXCEGIDV9ED16P2-001.BOB' then
  begin
    ExecuteSQL('DELETE FROM INSTITUTIONPAYE WHERE PIP_PREDEFINI = "CEG"');
    ExecuteSQL('DELETE FROM DADSLEXIQUE WHERE PDL_PREDEFINI = "CEG"');
    ExecuteSQL('DELETE FROM DSNAFFECTATION WHERE PNF_PREDEFINI = "CEG"');
    ExecuteSQL('DELETE FROM DSNDONNEES');
    ExecuteSQL('DELETE FROM DSNDECLARATIONS');
    ExecuteSQL('DELETE FROM DSNLEXIQUE');
    ExecuteSQL('DELETE FROM DSNFONCTION WHERE PFN_PREDEFINI = "CEG"');
    ExecuteSQL('DELETE FROM DSNFONCTIONLG WHERE PFL_PREDEFINI = "CEG"');
	end else if NamFic = 'UTIL0999F001.BOB' then
  begin
    ExecuteSQL('DELETE FROM MENU WHERE (MN_1=279) or (MN_1=0 AND MN_2=279)');
	end else if NamFic = 'XXCEGIDV9ED16P4-003.BOB' then
  begin
    ExecuteSQL('DELETE FROM PARAMSALARIE WHERE PPP_PREDEFINI = "CEG" AND PPP_PGTYPEINFOLS = "PCI"');
	end else if NamFic = 'BAT30999F650.BOB' then
  begin
    ExecuteSQL('DELETE FROM MENU WHERE MN_1=280');
	end else if NamFic = 'BAT30999F660.BOB' then
  begin
    ExecuteSQL('DELETE FROM BRGPDTABLESP');
    ExecuteSQL('DELETE FROM BRGPDTABLESL');
	end else if NamFic = 'BAT30999F661.BOB' then
  begin
    ExecuteSQL('DELETE FROM BRGPDTABLESL');
	end;
end;

procedure AfterIntegrationAllBobs;
begin
//
end;

procedure  AfterIntegrationBob (BobFile : string);
begin
  if ExtractFileName(BOBFile)='XXCEGIDV9ED11-4' then
  begin
    ExecuteSql ('UPDATE MENU SET MN_LIBELLE="DPAE" WHERE MN_1="42" AND MN_2="3" AND MN_3="2" AND MN_4="9"');
    ExecuteSql ('UPDATE MENU SET MN_ACCESGRP=(SELECT MN_ACCESGRP FROM MENU WHERE MN_1="49" AND MN_2="9" AND MN_3="7" AND MN_4="0"), MN_VERSIONDEV="X" WHERE MN_1="49" AND MN_2="9" AND MN_3="9" AND MN_4="0"');
    ExecuteSql ('UPDATE MENU SET MN_ACCESGRP=(SELECT MN_ACCESGRP FROM MENU WHERE MN_1="372" AND MN_2="2" AND MN_3="1" AND MN_4="0"), MN_VERSIONDEV="X" WHERE MN_1="372" AND MN_2="2" AND MN_3="3" AND MN_4="0"');
  end;
end;

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
  MAJMenu,First : Boolean ;
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
  First := True;
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
              if First then
              begin
              	BeforeIntegreOneBob;
                First := false;
              end;
              BeforeIntegrationBob (BobFile);
              iRes := TraiteFile(Directory+BobFile,STmp,True,GestionBOBCommun);
              ResultString := ResultString+STmp ;
              if iRes = 0 then
              begin
                 if not MAJMenu then
                    MAJMenu := IsBobMenu(BobFile) ;
                 Inc(iCount) ;
              end;
              AfterIntegrationBob (BobFile);
            end
            else
               ResultString := ResultString + 'déja intégré.' ;
         end;
      until sysutils.FindNext(sr)<>0;
      if iCount > 0 then
      begin
        AfterIntegrationAllBobs;
      end;
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

