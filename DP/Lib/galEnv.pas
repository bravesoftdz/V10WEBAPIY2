unit galEnv;
// Lecture de l'environnement (répertoires, mode de fonctionnement)

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HEnt1, HCtrls, Hmsgbox, IniFiles, FileCtrl, Registry, MajTable,
{$IFDEF EAGLCLIENT}
{$ELSE}
  Sql7Util, // pour TPGIConn
{$ENDIF}
  LicUtil;

{$IFDEF EAGLCLIENT}
// #### type déclaré dans Sql7Util
type
  TPGIConn = class(TObject)
    Societe  : String;
    Share    : String;
    Dir      : String;
    Driver   : String;
    Server   : String;
    Path     : String;
    ODBC     : String;
    Base     : String;
    Options  : String;
    User     : String;
    Password : String;
  end;
{$ENDIF}

function  GetModePcl(nomsoc: String): String;
function  GetEntreeEnv(nomsoc, section, entree: String): String;
function  GetDisquesLocaux(reploc: String; var repertoires: TStringList): Boolean;
{$IFDEF EAGLCLIENT}
{$ELSE}
procedure GetListSocietes(var LstSocietes : TList) ;
{$ENDIF}
// function  GetRepLocal : String; => utiliser TCBPPath.GetCegidDistri


////////////// IMPLEMENTATION //////////////
implementation

{***********A.G.L.***********************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 28/09/2000
Modifié le ... : 20/10/2000
Description .. : Renvoie : '1' si ModePcl,
Suite ........ :            '' ou '0' sinon
Suite ........ : Attention : ModePcl est dans un fichier cegidenv.ini
Suite ........ :  sur un partage (Share) dépendant de la soc choisie...
Suite ........ : Si soc vide, cherche la dern. soc ouverte (SocCommune=)
Suite ........ : Si vide aussi, prend la 1ère soc utilisable
Mots clefs ... : PGIENV;CEGIDENV;ENVIRONNEMENT
*****************************************************************}
function GetModePcl(nomsoc: String): String;
{var Buffer: array[0..1023] of Char;
    sIni: String;
    IniFile: TIniFile;
    i : integer;
    sWinPath, sShare, lasoc : String;
    SocList: TStrings; }
begin
  Result := '0';
  if FileExists(ExtractFilePath(Application.ExeName) + 'PgiLanceur.Dat') then Result := '1';

{ MD 30/06/05 obsolète :
  // ---- Répertoires
  GetWindowsDirectory(Buffer,1023);
  SetString(sWinPath, Buffer, StrLen(Buffer));

  // ---- Fichier ini
  sIni := sWinPath + '\CEGIDPGI.INI';
  if Not FileExists(sIni) then exit;
  IniFile := TIniFile.Create(sIni);

  // ---- Soc choisie
  lasoc := nomsoc;
  if nomsoc='' then lasoc := IniFile.ReadString('###DEFAULT###', 'SocCommune', '');
  // tjs pas de soc, on prend 1ère soc du ini
  if lasoc='' then
    begin
    SocList := TStringList.Create ;
    IniFile.ReadSections(SocList);
    // pour chaque société
    for i := 0 to SocList.count-1 do
      begin
      lasoc := SocList[i];
      // #### tester aussi GetSocRef
      if (lasoc<>'Reference') then
        begin
        sShare := IniFile.ReadString(lasoc, 'Share', '');
        if sShare <> '' then Break;
        end;
      end;
    SocList.free;
    end;
  if lasoc='' then begin IniFile.Free; exit; end;

  // ---- Emplacement (Share)
  sShare := IniFile.ReadString(lasoc, 'Share', '');
  IniFile.Free;

  // ---- Fichier d'environnement (cegidenv.ini)
  sIni := sShare+'\ENV\CEGIDENV.INI';
  if Not FileExists(sIni) then exit;
  IniFile := TIniFile.Create(sIni);

  // ---- Mode PCL
  Result := IniFile.ReadString('General', 'ModePCL', '');
  IniFile.Free;
}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 20/10/2000
Modifié le ... : 20/10/2000
Description .. : Lecture d'un paramètre dans le fichier d'environnement
Suite ........ : (cegidenv.ini) de la soc choisie.
Suite ........ : Si soc vide => cherche la dernière soc choisie
Suite ........ : Si vide aussi => cherche la première soc utilisable
Mots clefs ... : PGIENV;CEGIDENV;ENVIRONNEMENT
*****************************************************************}
function GetEntreeEnv(nomsoc, section, entree: String): String;
var Buffer: array[0..1023] of Char;
    sIni: String;
    IniFile: TIniFile;
    i : integer;
    sWinPath, sShare, lasoc : String;
    SocList: TStrings;
begin
  Result := '';
  // ---- Répertoires
  GetWindowsDirectory(Buffer,1023);
  SetString(sWinPath, Buffer, StrLen(Buffer));

  // ---- Fichier ini
  sIni := sWinPath + '\CEGIDPGI.INI';
  if Not FileExists(sIni) then exit;
  IniFile := TIniFile.Create(sIni);

  // ---- Soc choisie
  lasoc := nomsoc;
  if nomsoc='' then lasoc := IniFile.ReadString('###DEFAULT###', 'SocCommune', '');
  // tjs pas de soc, on prend 1ère soc du ini
  if lasoc='' then
    begin
    soclist := TStringList.Create ;
    IniFile.ReadSections(soclist);
    // pour chaque société
    for i := 0 to soclist.count-1 do
      begin
      lasoc := soclist[i];
      // #### tester aussi GetSocRef
      if (lasoc<>'Reference') then
        begin
        sShare := IniFile.ReadString(lasoc, 'Share', '');
        if sShare <> '' then Break;
        end;
      end;
    soclist.free;
    end;
  if lasoc='' then begin IniFile.Free; exit; end;

  // ---- Emplacement (Share)
  sShare := IniFile.ReadString(lasoc, 'Share', '');
  IniFile.Free;

  // ---- Fichier d'environnement (cegidenv.ini)
  sIni := sShare+'\ENV\CEGIDENV.INI';
  if Not FileExists(sIni) then exit;
  IniFile := TIniFile.Create(sIni);

  // ---- Lecture de l'entrée
  Result := IniFile.ReadString(section, entree, '');
  IniFile.Free;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 30/10/2000
Modifié le ... :   /  /    
Description .. : Récupère la liste des répertoires PGI
Suite ........ : de l'environnement local pour la
Suite ........ : création des bases dossiers en local.
Mots clefs ... : ENVIRONNEMENT;LOCAL;DISQUES;REPERTOIRES;
*****************************************************************}
function GetDisquesLocaux(reploc: String; var repertoires: TStringList): Boolean;
var  sIni: String;
     IniFile: TIniFile;
     ligne, volumeloc: String;
     i : integer;
begin
  Result := False;
  // ---- pas avec rép. réseau
  if (reploc='') or (Copy(reploc,1,2)='\\') then exit;
  sIni := reploc+'\ENV\CEGIDENV.INI';
  if not FileExists(sIni) then exit;

  IniFile := TIniFile.Create(sIni);
  if repertoires=Nil then repertoires := TStringList.Create;

  // ---- Liste des répertoires et partages
  // infos lues dans [répertoire]\env\cegidenv.ini (local ou serveur)
  // ;Diskx=Chemin_local;Partage_lanMan; (ou Chemin_local seul)
  ligne := '';
  i := 0;
  // pour chaque disque...
  repeat
    i := i+1;
    ligne := IniFile.ReadString('Disks', 'Disk'+IntToStr(i), '');
    // ajoute les lignes dans les TStringList
    if ligne <> '' then
      begin
      // rajout d'un point virgule final si nécessaire
      if Copy(ligne, Length(ligne), 1) <> ';' then ligne := ligne +';' ;
      volumeloc := ReadTokenSt(ligne);
      SupprimeLastAntislash(volumeloc);
      if Not DirectoryExists(volumeloc) then
        begin
        PGIInfo('Le répertoire '+volumeloc+ ' n''existe pas, il va être créé...', TitreHalley);
        ForceDirectories(volumeloc);
        if Not DirectoryExists(volumeloc) then
          begin
          // sortie si datas inaccessibles !
          PGIInfo('Le répertoire '+volumeloc+ ' n''a pas pu être créé.', TitreHalley);
          break;
          end;
        end;

      // remplit le TStringList
      repertoires.Add(volumeloc);
      end;
    // fin du if ligne <> ''
  until ligne = '';

  IniFile.Free;
  if repertoires.Count>0 then Result := True;
end;


{$IFDEF EAGLCLIENT}
{$ELSE}
{***********A.G.L.***********************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 30/10/2000
Modifié le ... :   /  /
Description .. : Récupère la liste des sociétés du INI
Mots clefs ... : ENVIRONNEMENT;INI;SOCIETES;
*****************************************************************}
procedure GetListSocietes(var LstSocietes : TList) ;
var sIni, socref : String;
    IniFile : TIniFile;
    soclist: TStringList;
    ObjConn: TPGIConn;
    i: Integer;
begin
  // ---- Lecture INI
  sIni := GetHalSocIni;
  if LstSocietes=Nil then LstSocietes := TList.Create;
  IniFile := TIniFile.Create (sIni) ;
  soclist := TStringList.Create ;
  IniFile.ReadSections(soclist);
  socref := UpperCase(GetSocRef);
  if soclist.count>0 then
    begin
    // pour chaque société
    for i := 0 to soclist.count-1 do
      begin
      // sauf [###DEFAULT###], [###DOSSIER###]...
      if (Copy(soclist[i],1,3) <> '###') then
        begin
        if UpperCase(soclist[i])= socref then Continue;
        objconn := TPGIConn.Create;
        objconn.Societe  := soclist[i];
        objconn.Share    := IniFile.ReadString(soclist[i], 'Share', '');
        objconn.Dir      := IniFile.ReadString(soclist[i], 'Dir', '');
        objconn.Driver   := IniFile.ReadString(soclist[i], 'Driver', '');
        objconn.Server   := IniFile.ReadString(soclist[i], 'Server', '');
        objconn.Path     := IniFile.ReadString(soclist[i], 'Path', '');
        objconn.ODBC     := IniFile.ReadString(soclist[i], 'ODBC', '');
        objconn.Base     := IniFile.ReadString(soclist[i], 'Database', '');
        objconn.Options  := IniFile.ReadString(soclist[i], 'Options', '');
        objconn.User     := IniFile.ReadString(soclist[i], 'User', '');
        if objconn.User<>'' then objconn.User := DecryptageSt(objconn.User);
        objconn.Password := IniFile.ReadString(soclist[i], 'Password', '');
        if objconn.Password<>'' then objconn.Password := DecryptageSt(objconn.Password);
        // ajout dans la liste
        LstSocietes.Add(ObjConn);
        end;
      end;
    end;
  soclist.Free;
  IniFile.Free;
end;
{$ENDIF}


end.


