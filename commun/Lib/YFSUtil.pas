{ FONCTION D'OPERATION SUR LA TABLE YFILESTD
  MIS EN PLACE PBASSET LE 10/02/2006
}
unit YFSUtil;

interface
uses
  Sysutils,
  Dialogs,
  Forms,
  classes,
  Controls,
  hent1,
  HCtrls,
  HDebug,
{$IFDEF eAGLClient}
  utileagl,
  eTablette,
  UHTTP,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF eAGLClient}
  uYFILESTD,
  ubob,
  utob,
  HMsgBox,
  FileCtrl,
  windows,
  extctrls;

function VerifNewYFILESTD(sDir, Sbob: string): boolean;
function LiasseGetTempPath(avecSTD: boolean = false): string;
procedure NLP_DeleteRepertoire(sDir: string);


implementation

{***********A.G.L.***********************************************
Auteur  ...... : Pascal BASSET
Créé le ...... : 04/08/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VerifNewYFILESTD(sDir, Sbob: string): boolean;
var
  St, sFile, sSQL, sDate, sBuffer: string;
  QDos: TQUERY;
  F1: TextFile;
begin
  Result := false;
  st := LiasseGetTempPath + 'PGI\STD\' + sDir + '\'; // debut de la construction de l'arborescence
  if V_PGI.Debug then
    Debug('st=' + st);
  sFile := st + 'YFILESTD.txt';
  if V_PGI.Debug then
    Debug('sfile=' + sfile);

  //Recherche de la dernière date de MAJ poiur les bobs
  sSQL := 'SELECT MAX(YB_BOBDATEMODIF) AS LADATE FROM YMYBOBS WHERE YB_BOBNAME LIKE "' + sBob + '%"';
  try
    QDos := OpenSQL(sSQL, true);
    if not QDos.Eof then
      sDate := QDos.findfield('LADATE').asString
    else
      sDate := '';
    Ferme(QDos);
  except
    PGIINFO(SSQL, 'Anomalie:VerifNewYFILESTD:' + sSQL);
    Ferme(QDos);
    exit;
  end;

  //si le fichier existe, on compare les dates
  if FileExists(sFile) then
  begin
    AssignFile(F1, sFile);
    Reset(F1);
    Readln(F1, sBuffer);
    //si même date
    if sBuffer = sDate then
      Result := true
    else
    begin
      CloseFile(F1);
      Erase(F1);
      AssignFile(F1, sFile);
      Rewrite(F1);
      Writeln(F1, sDate);
    end;
    CloseFile(F1);
  end
  else
  begin
    //Création du répertoire
    ForceDirectories(st); // creation de l'arboresence
    AssignFile(F1, sFile);
    Rewrite(F1);
    Writeln(F1, sDate);
    CloseFile(F1);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CV4
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. : Retourne le repertoire temporaire de l'utilisateur Windows
Suite ........ : Si le parametre avecSTD est TRUE, on ajoute le chemin
Suite ........ : PGI\STD
Mots clefs ... :
*****************************************************************}

function LiasseGetTempPath(avecSTD: boolean = false): string;
var
  Path: array[0..255] of Char;
begin
  GetTempPath(255, Path); // recup du repertoire Temp -> voir si la methode est la bonne
  Result := StrPas(Path);

  if AvecSTD then
    Result := Result + 'PGI\STD\';

  if V_PGI.Debug then
    Debug('LiasseGetTempPath=' + Result);
end;

{***********A.G.L.***********************************************
Auteur  ...... : PASCAL BASSET
Créé le ...... : 10/02/2006
Modifié le ... :   /  /
Description .. : cette fonction est mis la pour l'instant
Mots clefs ... :
*****************************************************************}

procedure NLP_DeleteRepertoire(sDir: string);
var
  iIndex: Integer;
  SearchRec: TSearchRec;
  sFileName: string;
begin
  sDir := sDir + '\*.*';
  iIndex := FindFirst(sDir, faAnyFile, SearchRec);
  while iIndex = 0 do
  begin
    sFileName := ExtractFileDir(sDir) + '\' + SearchRec.Name;
    if SearchRec.Attr = faDirectory then
    begin
      if (SearchRec.Name <> '') and
        (SearchRec.Name <> '.') and
        (SearchRec.Name <> '..') then
        NLP_DeleteRepertoire(sFileName);
    end else
    begin
      if SearchRec.Attr <> 0 then
        FileSetAttr(sFileName, 0);
      Sysutils.DeleteFile(sFileName);
    end;
    iIndex := FindNext(SearchRec);
  end;
  Sysutils.FindClose(SearchRec);
  RemoveDir(ExtractFileDir(sDir));
end;

end.

