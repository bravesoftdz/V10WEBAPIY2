{***********UNITE*************************************************
Auteur  ...... : MD
Créé le ...... : 26/11/2002
Modifié le ... :   /  /
Description .. : Regroupement variables globales DP
Mots clefs ... :
*****************************************************************}
unit EntDP;

interface

uses Windows, Classes, SysUtils, FileCtrl, HCtrls, Hent1, Forms, IniFiles, utob,
{$IFDEF EAGLCLIENT}
  UHttp,
{$ELSE}
{$IFNDEF DBXPRESS}dbtables {BDE}, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  Paramsoc, PGIAppli, HMsgBox
{$IFDEF BUREAU}
  , ctiAlerte
{$ENDIF}

{$IFDEF DP}
{$IFDEF EAGLCLIENT}
{$IFDEF BUREAU}
  , CegidPgiUtil // $$$ JP 28/12/05 - pour avoir un cWACegidPgi global à l'appli bureau
{$ENDIF}
{$ENDIF}
{$ENDIF}
  ;

{$IFDEF DP}
type
  LaVariableDP = class // ou Class(TObject)
  private
  public
    VersionMsSql: string; // Version de sql server (SQL7 / SQL2000 / ...)
       // ActiverChoixColl    : Boolean;  // Est-ce qu'on impose la collation
       // CollationDef        : String;   // Collation par défaut pour réparations SQL
    Group: Boolean; // True si séria du Bureau pour Entreprise/Groupe
    ModeStd: Boolean; // True si bureau de gestion des standards
    EcranParDefaut: string; // Ecran de travail par défaut du bureau
    SaisieActManager: Boolean; // Accès concept 26031
    GIActive: Boolean;
    GIAccesPlanfact: Boolean;  // Accès éditions ou cube du plan de facturation
    SeriaDP            : Boolean;
    SeriaGed           : Boolean;
    SeriaDocumentation : Boolean;
    SeriaMessagerie    : Boolean;
    SeriaLienDiode     : Boolean;
    SeriaKpmg          : Boolean;
    SeriaTDI           : Boolean;
    SeriaJeDeclare     : Boolean;
    SeriaExpertScan    : Boolean;
    ExeParam           : string; // Paramètres optionnels à passer à une appli
    ExeEncours         : Boolean; // Etat de l'appli éxécutée
    EnvUserOuvert      : Boolean; // Environnement user est chargé ?
    DuringLoadSynthese : Boolean; // Fiche synthese en cours de chargement
    NetExpertActif     : Boolean; // Est-ce que le cabinet gère Cegid S1 ASP
    EwsActif           : Boolean; // Est-ce que le cabinet gère la plate-forme EWS
    TranspEchanges     : Boolean; // Est-ce qu'on accède au transport à base d'échanges ComSx
    PathDatLoc         : string; // Chemin des standards cab (vu du serveur)
    ModeFonc           : string; // DB0, LIB, ...
    LeServeur          : string; // Nom du serveur où est la distrib :
                                 // - serveur de prod si share UNC ou lecteur réseau,
                                 // - nom de machine locale si share sur drive local
    LeServeurSql       : string; // Nom du serveur SQL (tient compte des instances nommées)
    LeMode             : string; // Mode de connexion : Local, Serveur, ou Web access
    CodeObligation     : string;
    DateObligation     : TDateTime; // Date d'affichage des dossiers pour une obligation
    TobObligation      : Tob;
    StrToutLeMonde     : String; // Chaine pour affectation des droits en SQL2005

       // $$$ JP 09/06/05: il faut la fiche d'alerte CTI (seule propriétaire de l'interface CTI)
{$IFDEF BUREAU}
    SeriaCti: boolean; // $$$ JP 13/10/05 - si module CTI sérialisé
    ctiAlerte: TFCtiAlerte;
{$ENDIF}

       // $$$ JP 04/01/06 - il faut ce ePlugIn en DP, pas seulement en mode BUREAU
{$IFDEF EAGLCLIENT}
{$IFDEF BUREAU}
    ePlugin: cWACegidPgi; // $$$ JP 28/12/05 - un cWACegidPgi global à l'appli bureau
{$ENDIF}
{$ENDIF}

{$IFDEF EAGLCLIENT}
    procedure RechargeApplisProposees;
    procedure LibereApplisProposees;
{$ENDIF}

  end;

var VH_DP: LaVariableDP;

procedure InitLaVariableDP;
procedure LibereLaVariableDP;
procedure ChargeParamsDP;
{$ENDIF DP}

{$IFDEF EAGLCLIENT}
{$ELSE}
function ExtractServerName(strpath: string): string;
function GetNomComputer: string;
function GetLeShare: string;
{$ENDIF}
function GetLeMode: string;
function GetModeFonc: string;


/////////// IMPLEMENTATION ///////////
implementation

{$IFDEF DP}

procedure InitLaVariableDP;
begin
  VH_DP := LaVariableDP.Create;
  VH_DP.VersionMsSql := '';
  // VH_DP.ActiverChoixColl   := False;
  // VH_DP.CollationDef       := 'French_CI_AS';
  VH_DP.Group := FileExists(ExtractFilePath(Application.ExeName) + 'hq.Dat'); //LM20070727
  VH_DP.ModeStd := False;
  VH_DP.EcranParDefaut := 'Liste des clients';
  VH_DP.SaisieActManager := True;
  VH_DP.GIActive := False;
  VH_DP.GIAccesPlanfact := False;
  VH_DP.SeriaDP := False;
  VH_DP.SeriaGed := False;
  VH_DP.SeriaDocumentation := False;
  VH_DP.SeriaMessagerie := False;
  VH_DP.SeriaLienDiode := False;
  VH_DP.SeriaKpmg := False;
  VH_DP.SeriaTDI := False;
  VH_DP.SeriaJeDeclare := False;
  VH_DP.SeriaExpertScan := False;
  VH_DP.ExeParam := '';
  VH_DP.ExeEncours := False;
  VH_DP.EnvUserOuvert := False;
  VH_DP.DuringLoadSynthese := False;
  VH_DP.NetExpertActif := False;
  VH_DP.EwsActif := False;
  VH_DP.TranspEchanges := False;
  VH_DP.PathDatLoc := '';
  VH_DP.ModeFonc := '';
  VH_DP.LeServeur := '';
  VH_DP.LeServeurSql := '';
  VH_DP.CodeObligation := '';
  VH_DP.DateObligation := iDate1900;
  VH_DP.TobObligation := Tob.Create('', nil, 0);
  VH_DP.StrToutLeMonde := '';

  // charge la liste des applis reconnues sur la machine (fichiers .mnu)
  if V_Applis = nil then
    V_Applis := TPGILesApplis.Create; // libéré dans le finalization de PGIAppli
end;


procedure LibereLaVariableDP;
begin
  if Assigned(VH_DP) then
  begin
    // $$$ JP 24/10/05 - libération TobObligation
    FreeAndNil(VH_DP.TobObligation);

    FreeAndNil(VH_DP); //VH_DP.Free; VH_DP:=Nil ;
  end;
end;


procedure ChargeParamsDP;
{$IFNDEF EAGLCLIENT}
var
  share: string;
{$ENDIF}
begin
  // Ces paramsoc sont à lire dans la base commune (##DP##.PARAMSOC)
  // donc utilisation de GetParamSocDp (cas de la GI qui compile avec la directive DP)

  // MD 02/06/06 : paramsoc obsolète
  // VH_DP.ActiverChoixColl := GetParamsocDpSecur('SO_MDACTIVERCHOIXCOLL', False);

  // s := GetParamsocDpSecur('SO_MDCOLLATIONDEF', '');
  // if s<>'' then VH_DP.CollationDef := RechDom('YYCOLLATIONDEF', s, False);

  // Ecran par défaut
  VH_DP.EcranParDefaut := GetParamsocDpSecur('SO_MDBUREAUPARDEFAUT', 'Liste des clients');

  // Gère Cegid S1 ASP
  VH_DP.NetExpertActif := (GetParamsocDpSecur('SO_NECWASURL', '') <> '');

{$IFDEF EWS}
  VH_DP.EwsActif := GetParamsocDpSecur('SO_NE_EWSACTIF', False);
{$ENDIF}

  VH_DP.TranspEchanges := GetParamsocDpSecur('SO_MDTRANSPECHANGES', False);

  VH_DP.PathDatLoc := GetParamsocDpSecur('SO_ENVPATHDATLOC', '');

  VH_DP.ModeFonc := GetModeFonc;

{$IFDEF EAGLCLIENT}
  VH_DP.LeServeur := ''; // non pertinent en eAGL, car pas d'accès direct
  VH_DP.LeServeurSql := '';
  // GH Mis en commentaire le 10/05/07 >>> VH_DP.LeMode := 'W';   // nouveau mode...
  if GetParamSocSecur('SO_MDCONFIGNOMADE', FALSE) = TRUE then
    VH_DP.LeMode := 'L' // Mode Nomade
  else
    VH_DP.LeMode := 'S'; // Mode Serveur

  VH_DP.RechargeApplisProposees; // complète liste d'applis
{$ELSE}
  share := GetLeShare;
  VH_DP.LeServeur := ExtractServerName(share);
{$IFDEF DBXADO}
  VH_DP.LeServeurSql := DBSOC.Params.Values['HostName'];
{$ELSE}
  GetFromRegistry(HKEY_LOCAL_MACHINE, 'SOFTWARE\ODBC\ODBC.INI\' + DBSOC.Params.Values['ODBC DSN'], 'Server', VH_DP.LeServeurSql);
{$ENDIF}
  VH_DP.LeMode := 'S'; // par défaut
  // en monoserveur TSE, ne pas se considérer comme local...
  // (même si le nom du serveur = nom de machine, on a un chemin réseau dans share !)
  if ((VH_DP.LeServeur = GetNomComputer) and (Copy(share, 1, 2) <> '\\'))
    or (VH_DP.LeServeur = '') then VH_DP.LeMode := 'L';
{$ENDIF EAGLCLIENT}

  if V_PGI.Driver=dbMSSQL2005 then
    begin
    //SL 14/11/2007 prise en compte des OS US
    if existeSQL('SELECT 1 FROM MASTER.DBO.SYSLOGINS WHERE NAME="AUTORITE NT\SYSTEM"') then
      VH_DP.StrToutLeMonde := 'tout le monde'
    else
      if existeSQL('SELECT 1 FROM MASTER.DBO.SYSLOGINS WHERE NAME="NT AUTHORITY\SYSTEM"') then
        VH_DP.StrToutLeMonde := 'everyone';
    end;

end;
{$ENDIF DP}


{$IFDEF EAGLCLIENT}
{$ELSE}

function ExtractServerName(strpath: string): string;
// extrait nom du serveur dans un chemin de la forme
//- \\serveur\partage\chemin
//- serveur:disque_local:\chemin (interbase)
//- sinon utilise GetUNCPathName (pour extraire unc si lecteur réseau)
//  => si local, retourne le nom de la machine local
var
  i, lg: Integer;
  tmp: string;
begin
  Result := '';
  lg := Length(strpath);
  // \\serveur\partage
  i := pos('\\', strpath);
  if (i = 1) then
  begin
    tmp := Copy(strpath, 3, lg - 2);
    // recherche le prochain "\" dans la chaine restante
    i := pos('\', tmp);
    if i > 1 then
    begin
      Result := Copy(tmp, 1, i - 1);
      exit;
    end;
  end;
  // essaie GetUNCPathName
  if DirectoryExists(ExtractFilePath(strpath)) then
  begin
    tmp := ExpandUNCFileName(strpath);
    // si récupère un nom réseau
    i := pos('\\', tmp);
    if (i = 1) then
    begin
      tmp := Copy(tmp, 3, lg - 2);
      // recherche le prochain "\" dans la chaine restante
      i := pos('\', tmp);
      if i > 1 then
      begin
        Result := Copy(tmp, 1, i - 1);
        exit;
      end;
    end;
    // sinon, rép. local => nom de machine locale
    Result := GetNomComputer;
    exit;
  end;
end;


function GetNomComputer: string;
// renvoie le nom de machine local, en majuscules
var
  Buffer: array[0..1023] of Char;
  lp: DWORD;
begin
  Result := '';
  lp := SizeOf(Buffer);
  GetComputerName(Buffer, lp);
  SetString(Result, Buffer, lp);
  Result := UpperCase(Result);
end;


function GetLeShare: string;
begin
  Result := V_PGI.Share;
end;
{$ENDIF EAGLCLIENT}


function GetLeMode: string;
{$IFNDEF EAGLCLIENT}
var
  share, serveur: string;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  Result := 'W';
{$ELSE}
  Result := 'S'; // par défaut
  share := GetLeShare;
  serveur := ExtractServerName(share);
  // en monoserveur TSE, ne pas se considérer comme local...
  // (même si le nom du serveur = nom de machine, on a un chemin réseau dans share !)
  if ((serveur = GetNomComputer) and (Copy(share, 1, 2) <> '\\'))
    or (serveur = '') then Result := 'L';
{$ENDIF}
end;


function GetModeFonc: string;
var
  Q: TQuery;
begin
  try
    Q := OpenSQL('SELECT SO_MODEFONC FROM SOCIETE', True, -1, '', True);
    Result := Q.FindField('SO_MODEFONC').AsString;
    Ferme(Q);
  except
    // PGIInfo('Le mode de fonctionnement par défaut n''est pas renseigné dans cette société.');
  end;
end;

{ LaVariableDP }


{$IFDEF DP}
{$IFDEF EAGLCLIENT}

procedure LaVariableDP.RechargeApplisProposees;
// - complète la liste des applis en ajoutant des applis non encore installées
//  (donc fictives !) lues dans CEGIDPGI.ASC
// - ces applis seront visibles sur le bureau CWAS grâce à des fichiers .ICO
//   et seront installées à la volée si l'utilisateur les active
// - 06/06/07 : rajout du n° Tag après le nom détaillé de l'appli
var
  sIni, st, nom: string;
  IniFile: TIniFile;
  SectionValues: TStringList;
  i: Integer;
  uneappli: TPGIAppli;
  ContentType, tmp: string;
begin
  LibereApplisProposees;
  sIni := ExtractFilePath(Application.ExeName) + 'CEGIDPGI.ASC';

  if not FileExists(sIni) then exit;

  // Lecture de CEGIDPGI.ASC
  IniFile := TIniFile.Create(sIni);
  SectionValues := TStringList.Create;
  IniFile.ReadSectionValues('BUREAUCWAS', SectionValues);

  // pour chaque appli
  for i := 0 to SectionValues.Count - 1 do
  begin
    st := SectionValues[i];
    nom := Trim(uppercase(ReadTokenPipe(st, '=')));

    // si .asc mal paramètré... (il faut .EXE)
    if (Length(nom) < 4) or (copy(nom, Length(nom) - 3, 4) <> '.EXE') then nom := nom + '.EXE';

    // élimine applis déjà connues dans le tableau des applis installées
    if V_Applis.Lappli(nom, False) <> nil then Continue;

    // élimine applis dont le kit n'est pas dispo côté serveur
    tmp := AppServer.RequestFile('/kits/' + Copy(nom, 1, Length(nom) - 4) + '.INF', ContentType);
    if (ContentType <> 'text/plain') or (tmp = '') then Continue;

    uneappli := TPGIAppli.Create;
    uneappli.Nom := nom;
    uneappli.Titre := ReadTokenSt(st); // Format : Libellé;NoTag
    uneappli.Serie := 'S5';
    uneappli.Ordre := Copy('000', 1, 3 - Length(IntToStr(i))) + IntToStr(i);
    uneappli.Pgi := True;
    uneappli.Visible := True;
    if st='' then
      uneappli.Tag := '78500' // ...ancienne valeur avant qu'on précise les tag, mais intérêt ?
    else
      uneappli.Tag := st;
    uneappli.Eagl := True;
    uneappli.Args := '###A INSTALLER###'; // seul repère !
    uneappli.Path := ExtractFilePath(Application.ExeName);
    V_Applis.Applis.Add(uneappli);
  end;
  SectionValues.Free;
  IniFile.Free;
end;

procedure LaVariableDP.LibereApplisProposees;
var
  i: Integer;
begin
  for i := V_Applis.Applis.Count - 1 downto 0 do
  begin
    if TPGIAppli(V_Applis.Applis[i]).Args = '###A INSTALLER###' then
    begin
      TPGIAppli(V_Applis.Applis[i]).Free;
      V_Applis.Applis.Delete(i);
    end;
  end;
end;
{$ENDIF EAGLCLIENT}
{$ENDIF DP}

end.
