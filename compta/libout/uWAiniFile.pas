{***********UNITE*************************************************
Auteur  ...... : Un inconnu
Créé le ...... : 05/07/2004
Modifié le ... :   /  /
Description .. : XP 05-07-2004 : Correction pour CBRS5
Mots clefs ... :
*****************************************************************}
unit uWAiniFile;

interface
uses windows,
  sysutils,
  hent1,
  classes,
  uWAINI,
  uWAInfoIni,
  utob,
  uWAEnvironnement
{$IFNDEF EAGLCLIENT}

{$IFDEF ODBCDAC}
  ,
  odbcconnection,
  odbctable,
  odbcquery,
  odbcdac
{$ELSE}
,
  uDbxDataSet
{$ENDIF}

{$ENDIF EAGLCLIENT}
  ;

type
  cWAINIFile = class (cWAINI)
  private
  public
    constructor create () ; override;
    destructor destroy; override;
    class function createInstance(): cWAINIFile ;

    {$IFNDEF EAGLCLIENT}
    function Chargedossier (VireRef, VireDP, VireModel: boolean) : HTStringList; override;
    function Ajoutedossier (WAInfoIni: cWAInfoIni) : boolean; override;
    function Supprimedossier (WAInfoIni: cWAInfoIni) : boolean; override;
    function Modifiedossier (WAInfoIni: cWAInfoIni) : boolean; override;
    function Existedossier (WAInfoIni: cWAInfoIni) : boolean; override;
    procedure ChargeDBParams (WAInfoIni: cWAInfoIni) ; override;
  public
    function GetListofEnvironnement() : HTStringList; override ;
    function GetEnvironnement(WAEnvironnement : cWAEnvironnement): boolean; override ;
    {$ENDIF EAGLCLIENT}
  end;

implementation

uses   cbpInifiles,
  licutil,
  licJFD,
  majtable
{$IFDEF EAGLSERVER}
  ,
  eSession,
  eIsapi,
  Forms
{$ELSE}
{$ENDIF EAGLSERVER}
  ;



constructor cWAINIFile.create () ;
begin
  inherited;
  ReaffecteNomTable (className, true) ;
  mode := 'FIC';
end;

destructor cWAINIFile.destroy;
begin
  //
  inherited;
end;

class function cWAINIFile.createInstance(): cWAINIFile ;
begin
result:=cWAINIFile.create() ;
end ;

{$IFNDEF EAGLCLIENT}
function cWAINIFile.Chargedossier (VireRef, VireDP, VireModel: boolean) : HTStringList;
var
  IniFile: THIniFile;
  i: Integer;
  sServer, sDriver, sRef: string;
{$IFDEF EAGLSERVER}
  sGroup: string;
{$ENDIF}
  // $$$ JP 18/01/08
  lITems: HTStrings ;
begin
  if getModeEagl then
  begin
    result := inherited Chargedossier (Vireref, VireDP, VireModel) ;
  end
  else
  begin
    // $$$ JP 18/01/08
    //litems := HTStringList.Create; // XP 06.11.2006 Remplacement de Result par lItems
//{$IFNDEF EAGLCLIENT}
//    IniFile := THIniFile.Create (NomIni) ;
//    IniFile.ReadSections (lItems) ;
//    result:=HTStringList.create() ;
//    result.Clear;
//    for i := 0 to lItems.Count - 1 do
//      result.Add(lItems[i]);
//    lITems.Free;
    Result := nil;
    IniFile := THIniFile.Create (NomIni);
    if IniFile <> nil then
    begin
      try
        lItems := HTStringList.Create;
        if lItems <> nil then
        begin
          try
            IniFile.ReadSections (lItems) ;
            Result := HTStringList.Create;
            Result.Clear;
            for i := 0 to lItems.Count - 1 do
              Result.Add (lItems [i]);
          finally
            FreeAndNil (lItems);
          end;
        end
        else
          exit;
    // $$$ JP 18/01/08: fin 1

        if VireRef then
        begin
          sRef := GetSocRef () ; //YCP une seule socref en ini
          i := result.IndexOf (sRef) ;
          if i >= 0 then result.Delete (i) ;
        end;
        i := 0;
        while i <= result.Count - 1 do
        begin
{$IFDEF EAGLSERVER}
          sGroup := IniFile.ReadString (result [i] , 'Group', '') ;
{$ENDIF EAGLSERVER}
          sServer := IniFile.ReadString (result [i] , 'Server', '') ;
          sDriver := IniFile.ReadString (result [i] , 'Driver', '') ;

          if (VireDP and (Copy (result [i] , 1, 3) = '###') )
            or ((Server <> '') and (Driver <> '') and ((sServer <> Server) or (SDriver <> Driver) ) )
{$IFDEF EAGLSERVER}
            or ((sGroup <> '') and (sGroup <> IDClient) )
{$ENDIF}
          then
            result.Delete (i)
          else
            inc (i) ;
        end;

        // trier les sociétes d'apres la section ###RESERVED###
{$IFDEF DECLA}
        if IniFile.SectionExists ('###RESERVED###') then
        begin
          for i := 0 to result.Count - 1 do
          begin
            if Copy (result [i] , 1, 14) = '###RESERVED###' then
              result.Delete (i)
            else
              result.Objects [i] := TObject (IniFile.ReadInteger ('###RESERVED###', result [i] , result.Count) ) ;
          end;
          QuickSort (result, 0, result.Count - 1) ;
        end;
{$ENDIF}
      // $$$ JP 18/01/08: suite
    //IniFile.Free;
      finally
        FreeAndNil (IniFile);
      end;
    end;
// $$$ JP 18/01/08 {$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIFile.Ajoutedossier (WAInfoIni: cWAInfoIni) : boolean;
begin
  if getModeEagl then
  begin
    result := inherited Ajoutedossier (WAInfoIni) ;
  end
  else
  begin
    result := true;
{$IFNDEF EAGLCLIENT}
    try
      with THIniFile.Create (NomIni) do
      begin
        WriteString (WAInfoIni.NomDeBase, 'Driver', WAInfoIni.Driver) ;
        WriteString (WAInfoIni.NomDeBase, 'Server', WAInfoIni.Server) ;
        WriteString (WAInfoIni.NomDeBase, 'Path', WAInfoIni.path) ;
        WriteString (WAInfoIni.NomDeBase, 'DataBase', WAInfoIni.Base) ;
        WriteString (WAInfoIni.NomDeBase, 'User', cryptageST (WAInfoIni.User) ) ;
        WriteString (WAInfoIni.NomDeBase, 'Password', cryptagest (WAInfoIni.Pass) ) ;
        WriteString (WAInfoIni.NomDeBase, 'ODBC', WAInfoIni.odbc) ;
        WriteString (WAInfoIni.NomDeBase, 'Options', WAInfoIni.options) ;
        WriteString (WAInfoIni.NomDeBase, 'LastUser', '') ;
        WriteString (WAInfoIni.NomDeBase, 'Groupe', WAInfoIni.Groupe) ;
        Free;
      end;
      // AVOIR XMG if Assigned(Inifile) then ReadSectionValues(BaseName,Inifile) Else UpdateFile ;
    except
      on E: Exception do
      begin
        result := false;
        ErrorMessage := E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIFile.Supprimedossier (WAInfoIni: cWAInfoIni) : boolean;
// $$$ JP 18/01/08
{$IFNDEF EAGLCLIENT}
var
  IniFile :THIniFile;
{$ENDIF}
begin
  if getModeEagl then
  begin
    result := inherited Supprimedossier (WAInfoIni) ;
  end
  else
  begin
    result := true;
{$IFNDEF EAGLCLIENT}
    // $$$ JP 18/01/08
    {try
      with THIniFile.Create (NomIni) do
      begin
        EraseSection (WAInfoIni.NomDeBase) ;
        //AVOIR XMG if Assigned(Inifile) then ReadSectionValues(BaseName,Inifile) Else UpdateFile ;
        UpdateFile;
        Free;
      end;
    except
      on E: Exception do
      begin
        result := false;
        ErrorMessage := E.Message;
      end;
    end;}
    IniFile := THIniFile.Create (NomIni);
    if IniFile <> nil then
      try
       try
        with IniFile do
        begin
          EraseSection (WAInfoIni.NomDeBase) ;
          //AVOIR XMG if Assigned(Inifile) then ReadSectionValues(BaseName,Inifile) Else UpdateFile ;
          UpdateFile;
        end;
       except
        on E: Exception do
        begin
          result := false;
          ErrorMessage := E.Message;
        end;
       end;
      finally
        FreeAndNil (IniFile);
      end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIFile.Modifiedossier (WAInfoIni: cWAInfoIni) : boolean;
begin
  if getModeEagl then
  begin
    result := inherited Modifiedossier (WAInfoIni) ;
  end
  else
  begin
    result := Ajoutedossier (WAInfoIni) ;
  end;
end;

function cWAINIFile.Existedossier (WAInfoIni: cWAInfoIni) : boolean;
// $$$ JP 18/01/08
var
  IniFile :THIniFile;
begin
  // $$$ JP 18/01/08
  {try
    with THIniFile.Create (NomIni) do
    begin
      result := SectionExists (WAInfoIni.NomDeBase) ;
      //AVOIR XMG if Assigned(Inifile) then ReadSectionValues(BaseName,Inifile) Else UpdateFile ;
      Free;
    end;
  except
    on E: Exception do
    begin
      result := false;
      ErrorMessage := E.Message;
    end;
  end;}
  IniFile := THIniFile.Create (NomIni);
  try
   try
    with IniFile do
      result := SectionExists (WAInfoIni.NomDeBase) ;
   except
    on E: Exception do
    begin
      result := false;
      ErrorMessage := E.Message;
    end;
   end;
  finally
    FreeAndNil (IniFile);
  end;
end;

procedure cWAINIFile.ChargeDBParams (WAInfoIni: cWAInfoIni) ;
var
  IniFile: THIniFile;
  FSect: HTStringList;
  EncryptUserPW: boolean;
begin
{$IFDEF eAGLServer}
  EncryptUserPW := true;
{$ELSE}
  // XP 05-07-2004 : Attention, dans le cas de CBRS5, nomini = Nom complet.
  // Il ne faut prendre que ExtractFileName
  EncryptUserPW := (Copy (ExtractFileName (nomIni) , 1, 1) <> '_') ;
{$ENDIF}

  IniFile := THIniFile.Create (nomIni) ;
  // $$$ JP 18/01/08
  try
    if WAInfoIni.NomDeBase = '' then
    begin
      FSect := HTStringList.Create;
      IniFile.ReadSections (FSect) ;
      WAInfoIni.NomDeBase := FSect [0] ;
      FSect.Free;
    end;
    WAInfoIni.Driver := IniFile.ReadString (WAInfoIni.NomDeBase, 'Driver', '') ;
    WAInfoIni.server := IniFile.ReadString (WAInfoIni.NomDeBase, 'Server', '') ;
    WAInfoIni.Path := IniFile.ReadString (WAInfoIni.NomDeBase, 'Path', '') ;
    WAInfoIni.ODBC := IniFile.ReadString (WAInfoIni.NomDeBase, 'ODBC', '') ;
    WAInfoIni.Base := IniFile.ReadString (WAInfoIni.NomDeBase, 'DataBase', '') ;
    WAInfoIni.User := IniFile.ReadString (WAInfoIni.NomDeBase, 'User', '') ;
    WAInfoIni.Pass := IniFile.ReadString (WAInfoIni.NomDeBase, 'PassWord', '') ;
    WAInfoIni.Options := IniFile.ReadString (WAInfoIni.NomDeBase, 'Options', '') ;
    WAInfoIni.Groupe := IniFile.ReadString (WAInfoIni.NomDeBase, 'Groupe', '') ;
    WAInfoIni.Share := IniFile.ReadString (WAInfoIni.NomDeBase, 'Share', '') ;
    //StDir:=IniFile.ReadString(StNom,'Dir','');
    if EncryptUserPW then
    begin
      WAInfoIni.User := DeCryptageSt (WAInfoIni.User) ;
      WAInfoIni.Pass := DeCryptageSt (WAInfoIni.Pass) ;
    end;
    WAInfoIni.Pass := DecryptJFD (WAInfoIni.Pass) ;

  // $$$ JP 18/01/08
  finally
    IniFile.Free;
  end;
end;

function cWAINIFile.GetListofEnvironnement(): HTStringList;
var
  IniFile: THIniFile;
  lItems: HTStringList ;
  i: integer ;
begin
  if ModeEAGL then
  begin
    result:=inherited GetListofEnvironnement ;
  end else
  begin
    // $$$ JP 18/01/08
//  lItems := HTStringList.Create; // XP 06.11.2006 Remplacement de Result par lItems
//  try
//    {$IFDEF EAGLSERVER}
//    IniFile := THIniFile.Create (GlobalModulePath+'\EnvCEGID.ini') ;
//    {$ELSE}
//    IniFile := THIniFile.Create ('EnvCEGID.ini') ;
//    {$ENDIF EAGLSERVER}
//    IniFile.ReadSectionValues('LISTENV',lItems) ;
//    result:=HTStringList.create() ;
//    result.Clear;
//    for i := 0 to lItems.Count - 1 do
//      result.Add(lItems[i]);
//    lITems.Free;
//    IniFile.Free;
//  except
//    on E: Exception do
//      begin
//      ErrorMessage := 'GetListofEnvironnement: ' + E.Message;
//      end;
//    end;
      Result := nil;
      {$IFDEF EAGLSERVER}
      IniFile := THIniFile.Create (GlobalModulePath+'\EnvCEGID.ini') ;
      {$ELSE}
      IniFile := THIniFile.Create ('EnvCEGID.ini') ;
      {$ENDIF EAGLSERVER}
      if IniFile <> nil then
      begin
        try
          lItems := HTStringList.Create; // XP 06.11.2006 Remplacement de Result par lItems
          if lItems <> nil then
          begin
            try
              try
                IniFile.ReadSectionValues ('LISTENV', lItems) ;
                Result := HTStringList.Create;
                Result.Clear;
                for i := 0 to lItems.Count - 1 do
                  Result.Add (lItems [i]);
              except
                on E: Exception do
                  ErrorMessage := 'GetListofEnvironnement: ' + E.Message;
              end;
            finally
              FreeAndNil (lItems);
            end;
          end;
        finally
          FreeAndNil (IniFile);
        end;
      end;
  end;
end;

function cWAINIFile.GetEnvironnement(WAEnvironnement : cWAEnvironnement): boolean;
var
  IniFile: THIniFile;
  tADMENV, T2: tob;
  tt: HTStringList ;
  name: string ;
begin
  if ModeEagl then
  begin
    result := inherited GetEnvironnement(WAEnvironnement) ;
  end
  else
  begin
    result:=false ;
    try
      TT:=HTStringList.Create ;

      // $$$ JP 18/01/08
      try
        {$IFDEF EAGLSERVER}
        IniFile := THIniFile.Create (GlobalModulePath+'\EnvCEGID.ini') ;
        {$ELSE}
        IniFile := THIniFile.Create ('EnvCEGID.ini') ;
        {$ENDIF EAGLSERVER}
        IniFile.ReadSectionValues('LISTCAB',TT) ;
        name:=TT.Values[IDHeberge] ;
        if name='' then name:=TT.Values['DEFAUT'] ;
      // $$$ JP 18/01/08
      finally
        TT.free ;
      end;

      tADMENV := TOB.Create('ADMENVIRONNEMENT_s', nil, -1);
      // $$$ JP 18/01/08
      try
        T2:=tob.create('ADMENVIRONNEMENT',tADMENV,-1) ;
        T2.AddChampSupValeur('AEN_IDENV',name);
        T2.AddChampSupValeur('AEN_LIBELLE',IniFile.ReadString(name,'LIBELLE','')) ;
        T2.AddChampSupValeur('AEN_VERSION',IniFile.ReadString(name,'VERSION','')) ;
        T2.AddChampSupValeur('AEN_BLOCNOTE',IniFile.ReadString(name,'BLOCNOTE','')) ;
        T2.AddChampSupValeur('AEN_NOMSOCREF',IniFile.ReadString(name,'NOMSOCREF','')) ;
        T2.AddChampSupValeur('AEN_PREFIXEBASE',IniFile.ReadString(name,'PREFIXEBASE','')) ;
        result:=WAEnvironnement.setvaluesFromTob(tADMENV) ;
      // $$$ JP 18/01/08
      finally
        tADMENV.free ;
        IniFile.free ;
      end;
    except
      on E: Exception do
      begin
        ErrorMessage := 'GetEnvironnement: ' + E.Message;
      end;
    end;
  end;
end ;
{$ENDIF EAGLCLIENT}


end.
