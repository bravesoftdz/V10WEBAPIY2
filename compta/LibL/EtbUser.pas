unit EtbUser;

interface

uses
  Windows, Messages, Forms, Classes, Graphics, Controls, Dialogs,
  StdCtrls, Buttons, ExtCtrls, HSysMenu, hmsgbox, Hctrls,uEntCommun,
  TntButtons;

function z_Teletransmission(NomFichier,RepDest,Modele,CodeBanque : PChar; bVerbose : integer) : integer ;
function z_GetDestinataires(Code, Libelle : HTStrings) : Boolean ;
//function z_GetCartes(Code, Libelle : TStrings; DestCode : string) : Boolean ;
function z_EclatementReleve(Nomfichier, RepDest : string) : Boolean ;
function ConvertDate(DateChar,Format : string ) : TDateTime ; overload
function ConvertDate(DateChar : string) : TDateTime ;  overload ; //XMG 24/02/2005

type
  TFEtbUser = class(TForm)
    LBBanque: TListBox;
    Panel1: TPanel;
    BAnnuler: THBitBtn;
    BValider: THBitBtn;
    BAide: THBitBtn;
    HMsgBox1: THMsgBox;
    HmTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure LBBanqueDblClick(Sender: TObject);
  private
    FName,FName2 : string ;
    FbBank : boolean ;
    FNameBank : string ;
    FBank : HTStringList ;
  public
    procedure SetFileName(Name : string) ;
    procedure SetFileName2(Name2 : string) ;
    procedure SetTypeBank(bBank : boolean) ;
    procedure SetNameBank(Bank : string) ;
  end;

implementation

uses
  SysUtils, IniFiles, HEnt1
  , Ent1, UtilPGI
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
     ;
      
{$R *.DFM}

const
   SECTION_BQ		='BANQUE';
   SECTION_MODELE	='MODELE';
   SECTION_PORT		='PORT';
   SECTION_MODEM	='MODEM';
   SECTION_CARTE	='CARTE';
   SECTION_TEL		='TELEPHONE';
   SECTION_TRACE	='TRACE';

   FILE_BQ		='BANQUE';
   FILE_MODELE		='MODELE';
   FILE_MODEM		='MODEM';
   FILE_CARTE		='CARTE';
   FILE_SCENARIO	='TRANS';
   FILE_GEN		='GENERAL';
   FILE_DEBUG		='DEBUG';

   FILENAME_DEBUG	='DEBUG.SYS';
   FILENAME_RECEPTION   ='FUGIER.ETB';

   MAX_ENTRY		=999;

   ETEBAC_PATH		='ETEBAC\';
   ETEBAC_EXT		='.INI';
   ETEBAC_PAR		='.PAR';
   ETEBAC_LOG		='.LOG';
   ETEBAC_SCN		='.SCN';
   ETEBAC_DAT		='.DAT';
   ETEBAC_SYS		='.SYS';

   ERR_LOGMISSING	=1000;
   ERR_DLLMISSING	=1001;
   ERR_PROCMISSING	=1002;
   ERR_USERCANCEL	=1003;
   ERR_BADCARTE         =1004;

   DLL_NAME             ='W32L341.DLL';
   DLL_DEBUGNAME        ='W32L341DBG.DLL';
   DLL_PROCNAME         ='w32trx';

   Y_BORD               =12;
   Y_ESPACE             =40;

var
   FBanque: TFEtbUser ;
   m_pszProfileName: string ;
   m_pszParfileName: string ;
   m_pszLogfileName: string ;
   m_pszGenfileName: string ;
   m_pszScnfileName: string ;
   m_pszDatfileName: string ;
   m_pszSysfileName: string ;

function GetStr(Name : string; bBank : boolean; Name2,Bank : string) : string ;
begin
FBanque.SetFileName(Name) ;
FBanque.SetTypeBank(bBank) ;
FBanque.SetFileName2(Name2) ;
FBanque.SetNameBank(Bank) ;
FBanque.FBank:=HTStringList.Create ;
if (FBanque.ShowModal=mrOk) and (FBanque.FBank.Count>0)
   then Result:=FBanque.FBank.Strings[FBanque.LBBanque.ItemIndex]
   else Result:='' ;
FBanque.FBank.Free ;
end ;

function GetBank(Name : string) : string ;
begin
Result:=GetStr(Name,TRUE,'','') ;
end;

function GetModele(Name,Name2,Bank : string) : string ;
begin
Result:=GetStr(Name,FALSE,Name2,Bank) ;
end;

procedure TFEtbUser.SetFileName(Name : string) ;
begin
FName:=Name ;
end;

procedure TFEtbUser.SetFileName2(Name2 : string) ;
begin
FName2:=Name2 ;
end;

procedure TFEtbUser.SetTypeBank(bBank : boolean) ;
begin
FbBank:=bBank ;
end;

procedure TFEtbUser.SetNameBank(Bank : string) ;
begin
FNameBank:=Bank ;
end;

procedure TFEtbUser.FormShow(Sender : TObject) ;
var Entry : string ;
    ic,lLigne : integer ;
    Profile, Profile2 : TIniFile ;
    Buffer, Buffer2 : string ;
    Modeles : HTStringList ;
begin
  lLigne:=0 ; Profile:=TIniFile.Create(FName) ;
  try
    LBBanque.Clear;
    if FbBank=TRUE then begin
      for ic:=0 to MAX_ENTRY do begin
        Entry:=Format('CLE%.3d', [ic]); Buffer:=Profile.ReadString(SECTION_BQ, Entry, '') ;
        if Buffer<>'' then begin
          FBank.Add(Buffer); Buffer2:=Profile.ReadString(Buffer, 'NOM', '') ;
          LBBanque.Items.Add(Buffer2) ;
          if lLigne<Canvas.TextWidth(Buffer2) then lLigne:=Canvas.TextWidth(Buffer2) ;
        end
        else break ;
      end ;
    end
    else begin
      Modeles:=HTStringList.Create;
      try
        for ic:=0 to MAX_ENTRY do begin
          Entry:=Format('E%.3d', [ic]); Buffer:=Profile.ReadString(SECTION_MODELE, Entry, '');
          if Buffer<>'' then begin
            Modeles.Add(Entry+' '+Buffer);
            if lLigne<Canvas.TextWidth(Modeles.Strings[ic]) then lLigne:=Canvas.TextWidth(Modeles.Strings[ic]);
          end
          else break;
        end;
        for ic:=0 to MAX_ENTRY do begin
          Entry:=Format('R%.3d', [ic]);
          Buffer:=Profile.ReadString(SECTION_MODELE, Entry, '');
          if (Buffer<>'') then begin
            Modeles.Add(Entry+' '+Buffer);
            if lLigne<Canvas.TextWidth(Modeles.Strings[ic]) then lLigne:=Canvas.TextWidth(Modeles.Strings[ic]);
          end
          else break;
        end;
        Profile2:=TIniFile.Create(FName2);
        try
          for ic:=0 to Modeles.Count-1 do begin
            Entry:=Copy(Modeles.Strings[ic], 1, 4);
            Buffer:=Profile2.ReadString(FNameBank, Entry, '');
            if (Buffer<>'') then begin
              FBank.Add(Entry);
              LBBanque.Items.Add(Modeles.Strings[ic]);
            end;
          end;
        finally
          Profile2.Free;
        end;
      finally
        Modeles.Free;
      end;
    end;
  finally;
    Profile.Free;
  end;
  if FbBank=FALSE then Caption:=HMsgBox1.Mess.Strings[0]
  else Caption:=HMsgBox1.Mess.Strings[1];
  if lLigne>244 then begin
    Width:=lLigne+20;
    BAide.Left:=Width-Y_BORD-Y_ESPACE;
    BAnnuler.Left:=BAide.Left-Y_ESPACE;
    BValider.Left:=BAnnuler.Left-Y_ESPACE;
  end;
  LBBanque.ItemIndex:=0;
end;

procedure TFEtbUser.LBBanqueDblClick(Sender: TObject);
begin
  ModalResult:=mrOK;
end;

procedure ZFile_GetProfilePath(const FileName: PChar);
var
  szBuff: string;
begin
  szBuff:=ExtractFilePath(Application.ExeName);
  m_pszProfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_EXT]);
end;

procedure ZFile_GetParamPath(const FileName: PChar);
var
  szBuff: string;
begin
  szBuff:=ExtractFilePath(Application.ExeName);
  m_pszParfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_PAR]);
end;

procedure ZFile_GetScenarioPath(const FileName: PChar; bEtebac: boolean);
var
  szBuff: string;
begin
  szBuff:=ExtractFilePath(Application.ExeName);
  if bEtebac=TRUE then
    m_pszScnfileName:=Format('etebac %s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_SCN])
  else
    m_pszScnfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_SCN])
end;

procedure ZFile_GetLogPath(const FileName: PChar);
var
  szBuff: string;
begin
  szBuff:=ExtractFilePath(Application.ExeName);
  m_pszLogfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_LOG])
end;

procedure ZFile_GetGenPath(const FileName: PChar);
var
  szBuff: string;
begin
  szBuff:=ExtractFilePath(Application.ExeName);
  m_pszGenfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_LOG])
end;

procedure ZFile_GetDatPath(const FileName: PChar);
var
  szBuff: string;
begin
  szBuff:=ExtractFilePath(Application.ExeName);
  m_pszDatfileName:=Format('%s%s%s', [szBuff, ETEBAC_PATH, FileName])
end;

procedure ZFile_GetDebugPath(const FileName: PChar);
var
  szBuff: string;
begin
  szBuff:=ExtractFilePath(Application.ExeName);
  m_pszSysfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_SYS])
end;

function ZFile_Copy(SourceName, DestName: string; bAppend: boolean; nbCar: integer): integer;
var
  inStream, outStream: file;
  nbRead: integer;
  NumRead, NumWritten, NumWritten2: integer;
  sBuffer: array[0..4096] of char;
begin
  AssignFile(inStream, SourceName);
  {$I-}
  Reset(inStream, 1);
  {$I+}
  if IOResult=0 then begin
    AssignFile(outStream, DestName);
    // TODO: test d'erreur sur Append
    if bAppend then begin
      {$I-}
      Reset(outStream, 1);
      {$I+}
      if IOResult<>0 then Rewrite(outStream, 1)
      else begin
        Seek(outStream, FileSize(outStream));
      end;
    end
    else Rewrite(outStream, 1);

    if nbCar<0 then
    begin
      while not Eof(inStream) do
      begin
        BlockRead(inStream, sBuffer, 1, NumRead);
        if (sBuffer[0]=#13) or (sBuffer[0]=#10) then continue;
        BlockWrite(outStream, sBuffer, NumRead, NumWritten);
      end;
    end
    else
    begin
      if nbCar>0 then
      begin
        BlockRead(inStream, sBuffer, nbCar+2, nbRead);
        if sBuffer[nbCar+1]=#10 then nbCar:=0;
        Reset(inStream,1);
      end;
      if nbCar=0 then
      begin
        while not Eof(inStream) do
        begin
          BlockRead(inStream, sBuffer, 1, NumRead);
          BlockWrite(outStream, sBuffer, NumRead, NumWritten);
        end;
      end
      else
      begin
        repeat
          BlockRead(inStream, sBuffer, nbCar, NumRead);
          BlockWrite(outStream, sBuffer, NumRead, NumWritten);
          BlockWrite(outStream, #13#10, 2, NumWritten2);
        until (NumRead = 0) or (NumWritten <> NumRead);
      end;
    end;
    CloseFile(outStream);
    CloseFile(inStream);
  end;
  Result:=0;
end;

function ZFile_WriteModemPar: integer;
var
  m_Com, m_Vitesse, m_Nbbit, m_Parite, m_Nbstop, l_Vitesse: Longint;
  c_Parite: AnsiChar;
  i_Cmd: integer;
  Stream: TextFile;
  buffer, m_Cmd: string;
  Profile: TIniFile;
begin
  // Modem
  ZFile_GetProfilePath(FILE_MODEM);
  ZFile_GetParamPath(FILE_MODEM);

  Profile:=TIniFile.Create(m_pszProfileName);
  try
    m_Com:=Profile.ReadInteger(SECTION_PORT, 'SERIE', 0);
    m_Vitesse:=Profile.ReadInteger(SECTION_PORT, 'VITESSE', 2);
    m_Nbbit:=Profile.ReadInteger(SECTION_PORT, 'BITS', 3);
    m_Parite:=Profile.ReadInteger(SECTION_PORT, 'PARITE', 0);
    m_Nbstop:=Profile.ReadInteger(SECTION_PORT, 'STOP', 0);

    case m_Parite of
        0: c_Parite:='N';
        1: c_Parite:='E';
        2: c_Parite:='O';
      else c_Parite:='N';
    end;
    case m_Vitesse of
        0: l_Vitesse:=300;
        1: l_Vitesse:=9600;
        2: l_Vitesse:=19200;
        3: l_Vitesse:=38400;
        4: l_Vitesse:=57600;
        5: l_Vitesse:=115200;
      else l_Vitesse:=19200;
    end;

    AssignFile(Stream, m_pszParfileName);
    Rewrite(Stream);

    // Paramètres du modem
    buffer:=Format('[PORTSER=com%d]', [m_Com+1]);
    Writeln(Stream, buffer);
    buffer:=Format('[VITESSE=%d]', [l_Vitesse]);
    Writeln(Stream, buffer);
    buffer:=Format('[NBBITS=%d]', [m_Nbbit+5]);
    Writeln(Stream, buffer);
    buffer:=Format('[PARITE=%s]', [c_Parite]);
    Writeln(Stream, buffer);
    buffer:=Format('[NBSTOP=%d]', [m_Nbstop+1]);
    Writeln(Stream, buffer);
    // Paramètres fixes
    buffer:='[LGMAXINF=4096]';
    Writeln(Stream, buffer);
    buffer:='[SEPARAT=059]';
    Writeln(Stream, buffer);
    buffer:='[SECURIT=O]';
    Writeln(Stream, buffer);
    buffer:='[NBREPET=002]';
    Writeln(Stream, buffer);
    buffer:='[TPRECEP=250]';
    Writeln(Stream, buffer);
    buffer:='[TPRECAR=020]';
    Writeln(Stream, buffer);
    buffer:='[DISQUE=]';
    Writeln(Stream, buffer);
    buffer:='[CONSOLE=]';
    Writeln(Stream, buffer);
    // Paramètres commandes du modem
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'CMDREC', 'ATI');
    buffer:=Format('[C_RECONN=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VOKREC', '288');
    buffer:=Format('[O_RECONN=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VECREC', 'ERR');
    buffer:=Format('[N_RECONN=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    i_Cmd:=Profile.ReadInteger(SECTION_MODEM, 'DATREC', 100);
    buffer:=Format('[T_RECONN=%d]', [i_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'CMDRAZ', 'ATZ');
    buffer:=Format('[C_RAZ=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VOKRAZ', 'OK');
    buffer:=Format('[O_RAZ=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VECRAZ', 'ERR');
    buffer:=Format('[N_RAZ=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    i_Cmd:=Profile.ReadInteger(SECTION_MODEM, 'DATRAZ', 100);
    buffer:=Format('[T_RAZ=%d]', [i_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'CMDINI', 'ATE1V1');
    buffer:=Format('[C_INIT=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VOKINI', 'OK');
    buffer:=Format('[O_INIT=%s]', [m_Cmd]);                 // PCS
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VECINI', 'ERR');
    buffer:=Format('[N_INIT=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    i_Cmd:=Profile.ReadInteger(SECTION_MODEM, 'DATINI', 100);
    buffer:=Format('[T_INIT=%d]', [i_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'CMDNUM', 'ATD');
    buffer:=Format('[C_APPEL=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VOKNUM', 'CONN');
    buffer:=Format('[O_APPEL=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VECNUM', 'ERR');
    buffer:=Format('[N_APPEL=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    i_Cmd:=Profile.ReadInteger(SECTION_MODEM, 'DATNUM', 1200);
    buffer:=Format('[T_APPEL=%d]', [i_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'CMDAUT', 'ATS0=2');
    buffer:=Format('[C_APP_ENT=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VOKAUT', 'OK');
    buffer:=Format('[O_APP_ENT=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VECAUT', 'ERR');
    buffer:=Format('[N_APP_ENT=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    i_Cmd:=Profile.ReadInteger(SECTION_MODEM, 'DATAUT', 100);
    buffer:=Format('[T_APP_ENT=%d]', [i_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VOKDET', 'CONNECT');
    buffer:=Format('[O_ATT_APP=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile.ReadString(SECTION_MODEM, 'VECDET', 'CARRIER');
    buffer:=Format('[N_ATT_APP=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    i_Cmd:=Profile.ReadInteger(SECTION_MODEM, 'DATDET', 500);
    buffer:=Format('[T_ATT_APP=%d]', [i_Cmd]);
    Writeln(Stream, buffer);
    Writeln(Stream, '');
    Flush(Stream);
    CloseFile(Stream);
  finally
    Profile.Free;
  end;
  Result:=0;
end;

function ZFile_WriteScenario(NomFichier, RepDest, Banque, Modele: PChar; bVerbose : Integer): integer;
var Stream: TextFile ; PathDest, buffer, z_Buff, m_Cmd, m_Cmd2, Entry : string ;
     i_Cmd: Longint ; Profile, Profile2: TIniFile ; i : Integer ; bBad : Boolean ;
begin
  Result := 0;
  ZFile_GetScenarioPath(FILE_SCENARIO, FALSE);
  AssignFile(Stream, m_pszScnfileName);
  Rewrite(Stream);
  // Protocole
  Writeln(Stream, '[PROTOCOL=ETEBAC]');

  ZFile_GetLogPath(FILE_SCENARIO);
  // Fichier LOG de la session
  buffer:=Format('[FICLOG=%s]', [m_pszLogfileName]);
  Writeln(Stream, buffer);
  // La fenêtre de status s'affiche chez le client
  if bVerbose>0 then Writeln(Stream, '[SIGNAL=O]') else WriteLn(Stream, '[SIGNAL=N]') ;
  // Empêche la DLL almacom d'ajouter &Q5 à la séquence de numérotation (Bug?)
  Writeln(Stream, '[MNP4=N]');
  // Nombre de retry
  Writeln(Stream, '[NBRETRY=1]');
  // Numero de téléphone
  ZFile_GetProfilePath(FILE_MODEM);
  Profile:=TIniFile.Create(m_pszProfileName);
  try
    if Profile.ReadInteger(SECTION_TEL, 'SYNC', 0)=1 then
      m_Cmd:=Profile.ReadString(SECTION_TEL, 'TSYNC', '')
    else
      m_Cmd:=Profile.ReadString(SECTION_TEL, 'TASYNC', '');
    m_Cmd2:=Profile.ReadString(SECTION_TEL, 'TELEXT', '');
    if m_Cmd2<>'' then
      buffer:=Format('[NUMTEL=%s%s]', [m_Cmd2, m_Cmd])
    else
      buffer:=Format('[NUMTEL=%s]', [m_Cmd]);
    Writeln(Stream, buffer);

    ZFile_GetProfilePath(FILE_BQ);
    Profile2:=TIniFile.Create(m_pszProfileName);
    try
    m_Cmd:=Profile2.ReadString(Banque, 'ADR', '');
    buffer:=Format('[NUMX25=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    m_Cmd:=Profile2.ReadString(Banque, 'EXT', '');
    buffer:=Format('[DONX25=%s]', [m_Cmd]);
    Writeln(Stream, buffer);
    i_Cmd:=Profile2.ReadInteger(Banque, 'PCV', 0);
    finally
      Profile2.Free;
    end;

    // Identification
    if i_Cmd>0 then
      begin
      Writeln(Stream, '[IDENT=O]'); Writeln(Stream, '');
      m_Cmd:=Profile.ReadString(SECTION_TEL, 'NUI', '');
      buffer:=Format('[NUI=%s]', [m_Cmd]);
      Writeln(Stream, buffer);
      end
    else
      Writeln(Stream, '[IDENT=N]');

    // Sens de la communication (E ou R)
    buffer:=Format('[SENS=%s]', [Modele[0]]);
    Writeln(Stream, buffer);
    // Nom du fichier DAT
    if Modele[0]='E' then
      begin
      ZFile_GetDatPath(NomFichier);
      PathDest:=Format('%s%s', [RepDest, NomFichier]);
      ZFile_Copy(PathDest, m_pszDatfileName, FALSE, -1);
      end;
    ZFile_GetDatPath(NomFichier);
    buffer:=Format('[FICDAT=%s]', [m_pszDatfileName]);
    Writeln(Stream, buffer);

    // lecture Carte d'appel
    ZFile_GetProfilePath(FILE_CARTE);
    Profile2:=TIniFile.Create(m_pszProfileName);
    try
      Entry:=Format('%-4.4s', [Modele]);
      m_Cmd:=Profile2.ReadString(Banque, Entry, '');
      if m_Cmd<>'' then
        begin
        z_Buff:=Copy(m_Cmd, 6, 4); bBad:=FALSE ;
        for i:=1 to Length(z_Buff) do
          if not (z_Buff[i] in ['0'..'9']) then begin bBad:=TRUE ; break ; end ;
        if bBad then i_Cmd:=0 else i_Cmd:=StrToInt(z_Buff) ;
        end
      else i_Cmd:=0;
    finally
      Profile2.Free;
    end;

    if i_Cmd=0 then begin CloseFile(stream) ; Result:=0 ; Exit ; end ;

    ZFile_GetDatPath(NomFichier);
    buffer:=Format('[LGART=%d]', [i_Cmd]);
    Writeln(Stream, buffer);
    Writeln(Stream, '[NBFINART=0]');
    Writeln(Stream, '[TRANSC=N]');
    Writeln(Stream, '[TAILCAR=80]');

    // Carte d'appel
    buffer:=Format('[CARTE=%80.80s]', [string((PChar(m_Cmd)+1))]);	// Ne pas prendre l'@
    Writeln(Stream, buffer);
    Writeln(Stream, '');
    CloseFile(stream);
  finally
    Profile.Free;
  end;
Result:=i_Cmd;	// Retourne le nombre de caractères dans la séquence ETEBAC3
end;

function AnalyseLog(bVerbose, bLog: integer): integer;
var
   stream: TextFile;
    s_out: TextFile;
   Buffer: string;
     s_zz: string;
    v_ret: integer;
   bPurge: integer;
      bOn: boolean;
  i_Jours: integer;
      nbj: word;
    j, jl: word;
    m, ml: word;
    a, al: word;
  Profile: TIniFile;
    FDate: TDateTime;
    NDate: TDateTime;
begin
  v_ret:=0;
  AssignFile(Stream, m_pszLogfileName);
  {$I-}
  Reset(Stream);
  {$I+}
  if IOResult<>0 then
  begin
    if bVerbose>0 then MessageAlerte(FBanque.HMsgBox1.Mess.Strings[2]);
    v_ret:=ERR_LOGMISSING;
  end
  else begin
    while not Eof(Stream) do
    begin
      Readln(Stream, Buffer);
      if (Buffer[1]='0') and (Buffer[2]='3') then
      begin
        if bVerbose>0 then MessageAlerte(TraduireMemoire(string(PChar(Buffer)+39)));
        s_zz:=Buffer[4]+Buffer[5]+Buffer[6]+Buffer[7];
	v_ret:=StrToInt(s_zz);
	break;
      end;
    end;
    CloseFile(stream);
    ZFile_GetGenPath(FILE_GEN);
    if bLog>0 then ZFile_Copy(m_pszLogfileName, m_pszGenfileName, TRUE, 0);
    DeleteFile(PChar(m_pszLogfileName));

    ZFile_GetProfilePath(FILE_MODEM);
    Profile:=TIniFile.Create(m_pszProfileName);
    try
      bPurge:=Profile.ReadInteger(SECTION_TRACE, 'EVDEL', 0);
      if bPurge>0 then
      begin
        i_Jours:=Profile.ReadInteger(SECTION_TRACE, 'JOURS', 0);
        DecodeDate(now, a, m, j);
        nbj:=MonthDays[IsLeapYear(a), m];
        if (j-i_Jours)<0 then begin j:=(nbj+j)-i_Jours; m:=m-1; if m<=0 then begin m:=12; a:=a-1; end; end else j:=j-i_Jours;
        if j<=0 then j:=1;
        nbj:=MonthDays[IsLeapYear(a), m];
        if j>nbj then j:=nbj;
        bOn:=FALSE;
        AssignFile(Stream, m_pszGenfileName);
        {$I-}
        Reset(Stream);
        {$I+}
        if IOResult=0 then
        begin
          AssignFile(s_out, m_pszLogfileName);
          {$I-}
          Append(s_out);
          {$I+}
          if IOResult<>0 then Rewrite(s_out);
          while not Eof(Stream) do
      	  begin
            Readln(Stream, Buffer);
	    if (Buffer[1]='0') and (Buffer[2]='1') then
	    begin
	      s_zz:=Buffer[40]+Buffer[41]; jl:=StrToInt(s_zz);
	      s_zz:=Buffer[43]+Buffer[44]; ml:=StrToInt(s_zz);
	      s_zz:=Buffer[46]+Buffer[47]; al:=StrToInt(s_zz);
              if al<50 then al:=al+2000 else al:=al+1900;
              FDate:=EncodeDate(al, ml, jl);
              NDate:=EncodeDate(a, m, j);
	      if FDate>=NDate then bOn:=TRUE else bOn:=FALSE;
            end;
	    // Copie
	    if bOn=TRUE then Writeln(s_out, Buffer);
          end;
        end;
        CloseFile(stream);
        CloseFile(s_out);
        DeleteFile(PChar(m_pszGenfileName));
        RenameFile(m_pszLogfileName, m_pszGenfileName);
      end;
    finally
      Profile.Free;
    end;
  end;
  Result:=v_ret;
end;

function z_GetDestinataires(Code, Libelle : HTStrings) : Boolean ;
var Profile : TIniFile ; ic : Integer ; Entry, Buffer, Buffer2 : string ;
begin
Result:=TRUE ;
ZFile_GetProfilePath(FILE_BQ);
Profile:=TIniFile.Create(m_pszProfileName) ;
for ic:=0 to MAX_ENTRY do
    begin
    Entry:=Format('CLE%.3d', [ic]);
    Buffer:=Profile.ReadString(SECTION_BQ, Entry, '') ;
    if Buffer<>'' then
       begin
       Code.Add(Buffer) ;
       Buffer2:=Profile.ReadString(Buffer, 'NOM', '') ;
       Libelle.Add(Buffer2) ;
       end else break ;
    end ;
Profile.Free ;
if Code.Count<=0 then Result:=FALSE ;
end ;

(*
function z_GetCartes(Code, Libelle : TStrings; DestCode : string) : Boolean ;
var Profile : TIniFile ; ic : Integer ; Buffer : string ;
begin
Result:=TRUE ;
ZFile_GetProfilePath(FILE_CARTE);
Profile:=TIniFile.Create(m_pszProfileName) ;
Profile.ReadSection(DestCode, Code) ;
Profile.Free ;
if Code.Count<=0 then begin Result:=FALSE ; Exit ; end ;
ZFile_GetProfilePath(FILE_MODELE);
Profile:=TIniFile.Create(m_pszProfileName) ;
for ic:=0 to Code.Count-1 do
  begin
  Buffer:=Profile.ReadString(SECTION_MODELE, Code[ic], '') ;
  Libelle.Add(Buffer) ;
  end ;
Profile.Free ;
end ;*)

{Christophe Ayel : 29/09/2003
Cette procédure retraite les fichiers afin d'obtenir en sortie
des enregistrements de 120 caractères avec un retour à la ligne.
Le fichier initial est sauvegardé dans le fichier ~tmpFichier.ext.
{---------------------------------------------------------------------------------------}
procedure FormateReleve120(NomFichier : string );
{---------------------------------------------------------------------------------------}
var FSource, FDest : TextFile;
    c : char;
    iCompteur : integer;
    stBuffer : string;
    stFichierSource : string;
begin
  { Sauvegarde du fichier initial }
  stFichierSource := ExtractFilePath(NomFichier)+'~tmp'+ExtractFileName(NomFichier);
  CopyFile( PChar(NomFichier), PChar(stFichierSource),False);
  { Ouverture du fichier source }
  AssignFile ( FSource, stFichierSource );
  Reset ( FSource );
  { Ouverture du fichier de destination }
  AssignFile ( FDest, NomFichier );
  ReWrite ( FDest);
  iCompteur := 0;
  { Traitement du fichier }
  while not Eof(FSource) do
  begin
    Read (FSource, c);
    if Ord(c) < 30 then continue;
    Inc( iCompteur, 1 );
    stBuffer := stBuffer + c;
    if iCompteur = 120 then
    begin
      Writeln ( FDest, stBuffer);
      iCompteur := 0;
      stBuffer := '';
    end;
  end;
  if stBuffer <> '' then Writeln ( FDest, stBuffer);
  { Fermeture des fichiers }
  CloseFile ( FSource );
  CloseFile ( FDest );
end;

function z_EclatementReleve(NomFichier, RepDest : string) : Boolean;
var FSource, FDest: TextFile ; StReleve, NomDest : string ; bOuvert : Boolean ;
    CurrDate : TDateTime ; Year, Month, Day, Hour, Min, Sec, Msec, MsecTmp : Word ;
    NbrEnreg : integer ;
begin
bOuvert:=FALSE ; Result:=TRUE; StReleve:='' ;
{$IFNDEF EAGLCLIENT}
FormateReleve120( NomFichier );
{$ENDIF}
AssignFile(FSource, Nomfichier);
{$I-} Reset(FSource); {$I+}
if IOResult<>0 then begin Result:=FALSE ; Exit ; end ;
NbrEnreg:=0 ;
Readln(FSource, StReleve) ;
while not EOF(FSource) do
  begin
  if ((VH^.PaysLocalisation=CodeISOES)  and (Copy(StReleve,1,2)<>'88')) then
     inc(NbrEnreg) ;
  if ((VH^.PaysLocalisation=CodeISOES)  and (Copy(StReleve,1,2)='11')) or
     ((VH^.PaysLocalisation<>CodeISOES) and (Copy(StReleve,1,2)='01')) then //XVI 24/02/2005
    begin
    if bOuvert then begin Result:=FALSE ; CloseFile(FDest) ; end ;
    CurrDate:=Date ;
    DecodeDate(CurrDate, Year, Month, Day) ;
    DecodeTime(Time, Hour, Min, Sec, Msec) ;
    MSecTmp:=MSec ; while MSecTmp=MSec do DecodeTime(Time, Hour, Min, Sec, Msec) ;
    if RepDest[Length(RepDest)-1]<>'\' then RepDest:=RepDest+'\' ;
    NomDest:=Format('%s%s%.2d%.2d%.2d%.2d%.2d%.3d.DAT', [RepDest, FormatDateTime('yy', CurrDate), Month, Day, Hour, Min, Sec, MSec]) ;
    AssignFile(FDest, NomDest) ;
    ReWrite(FDest) ;
    bOuvert:=TRUE ;
    end ;
    if bOuvert then begin
       if ((VH^.PaysLocalisation=CodeISOES)  and (Copy(StReleve,1,2)<>'88')) or
          (VH^.PaysLocalisation<>CodeISOES)                                   then //XVI 24/02/2005
          Writeln(FDest, StReleve) ;
      if ((VH^.PaysLocalisation=CodeISOES)  and (Copy(StReleve,1,2)='33')) or
         ((VH^.PaysLocalisation<>CodeISOES) and (Copy(StReleve,1,2)='07'))   then Begin //XVI 24/02/2005
         CloseFile(Fdest) ;
         bOuvert:=FALSE ;
      end ;
    end ;
  Readln(FSource, StReleve) ;
  end ;
if ((VH^.PaysLocalisation=CodeISOES)  and ((Copy(StReleve,1,2)<>'88')) or (NbrEnreg<>Valeuri(Copy(StReleve,21,6)))) or
   ((VH^.PaysLocalisation<>CodeISOES) and (Copy(StReleve,1,2)<>'07')) then
   Result:=FALSE
else
  if (VH^.PaysLocalisation<>CodeISOES) then
     Writeln(FDest, StReleve) ;  //XVI 24/02/2005
if bOuvert then  CloseFile(FDest) ;
CloseFile(FSource) ;
if (VH^.PaysLocalisation=CodeISOES) and (not Result) then
   DeleteFile(NomDest) ; //XVI 24/02/2005
end;

function z_Teletransmission(NomFichier, RepDest, Modele, CodeBanque: PChar; bVerbose: integer):integer;
type
  TwProc=function(hwndParent: HWnd; KeepSignalisationWindow: integer;
                 parettd: PChar; mode: PChar; commande: PChar): integer; stdcall;
var Ident: THandle ; wProc: TwProc ; Profile: TIniFile ; nbCar: LongInt ;
    bDebug, bLog, v_ret: Integer ; PathDest, NameDest: array[0..500] of Char ;
    BankSel, ModeleSel, TmpName, FileName: string ; bCopy : Boolean ;
begin
  Result := 0;
  v_ret:=0 ;  bCopy:=TRUE ;

  if NomFichier[0]=#0 then FileName:=FILENAME_RECEPTION else FileName:=string(NomFichier) ;

  FBanque:=TFEtbUser.Create(nil);
  try

  m_pszParfileName:=ExtractFilePath(Application.ExeName);
  m_pszProfileName:=Format('%s%s', [m_pszParfileName, ETEBAC_PATH]);
  StrCopy(PathDest, RepDest);
  if RepDest[StrLen(RepDest)-1]<>'\' then StrCat(PathDest, '\') ;
  if CompareText(m_pszProfileName, PathDest)=0 then bCopy:=FALSE ;
  {$I-}
  MkDir(m_pszProfileName);
  {$I+}
  if IOResult>=0 then
  begin
    ZFile_GetProfilePath(FILE_BQ);
    if FileExists(m_pszProfileName)=FALSE then
      MessageAlerte(FBanque.HMsgBox1.Mess.Strings[3]+Chr(10)+Chr(13)+FBanque.HMsgBox1.Mess.Strings[5])
    else begin
      if CodeBanque=nil then
        BankSel:=GetBank(m_pszProfileName)
      else BankSel:=string(CodeBanque);

      if BankSel<>'' then begin

        ZFile_GetProfilePath(FILE_CARTE);
        if FileExists(m_pszProfileName)=FALSE then
          MessageAlerte(FBanque.HMsgBox1.Mess.Strings[4]+Chr(10)+Chr(13)+FBanque.HMsgBox1.Mess.Strings[5])
        else begin
          if Modele=nil then begin
	        ZFile_GetProfilePath(FILE_MODELE);
            TmpName:=m_pszProfileName;
	        ZFile_GetProfilePath(FILE_CARTE);
            ModeleSel:=GetModele(TmpName, m_pszProfileName, BankSel)
          end
          else ModeleSel:=string(Modele);

          if ModeleSel<>'' then begin
            // Préparation du scénario
            ZFile_WriteModemPar;
            nbCar:=ZFile_WriteScenario(PChar(FileName), RepDest, PChar(BankSel), PChar(ModeleSel), bVerbose);
            if nbCar=0 then begin Result:=ERR_BADCARTE ; Exit ; end ;
            // Charge la DLL.
            ZFile_GetProfilePath(FILE_MODEM);
            Profile:=TIniFile.Create(m_pszProfileName);
            try
              bDebug:=Profile.ReadInteger(SECTION_TRACE, 'DEBUG', 0);
              bLog:=Profile.ReadInteger(SECTION_TRACE, 'CONNECT', 0);
            finally
              Profile.Free;
            end;
            if bDebug=0 then Ident:=LoadLibrary(DLL_NAME) else Ident:=LoadLibrary(DLL_DEBUGNAME);
            if (Ident<>0) then begin
              @wProc:=GetProcAddress(Ident, DLL_PROCNAME);
              if @wProc<>nil then begin
                ZFile_GetScenarioPath(FILE_SCENARIO, TRUE);
{$IFNDEF PFUGIER}
                v_ret:=wProc(Application.Handle, 1, PChar(m_pszParfileName), 'C', PChar(m_pszScnfileName));
{$ENDIF}
                FreeLibrary(Ident);
                // Test du .log
                if v_ret=0 then begin
{$IFNDEF PFUGIER}
                  v_ret:=AnalyseLog(bVerbose, bLog);
{$ENDIF}
                  if v_ret=0 then begin
                    if (ModeleSel[1]='R') then
                      begin
                      ZFile_GetDatPath(PChar(FileName));
                      StrCopy(NameDest, PathDest) ;
                      StrCat(NameDest, PChar(FileName));
                      if bCopy then v_ret:=ZFile_Copy(m_pszDatfileName, string(NameDest), FALSE, nbCar) ;
                      if v_ret=0 then
                        begin
                        z_EclatementReleve(NameDest, PathDest) ;
                        DeleteFile(NameDest) ;
                        end ;
                      end ;
                  end;
                end;
                ZFile_GetDatPath(PChar(FileName));
                DeleteFile(PChar(m_pszDatfileName));
                if bDebug>0 then begin
                  ZFile_GetDebugPath(FILE_DEBUG);
                  DeleteFile(PChar(m_pszSysfileName));
                  ZFile_Copy(FILENAME_DEBUG, m_pszSysfileName, FALSE, 0);
                  DeleteFile(FILENAME_DEBUG);
                end;
              end
              else begin

                if bVerbose>0 then MessageAlerte(FBanque.HMsgBox1.Mess.Strings[6]);
                v_ret:=ERR_PROCMISSING;
              end;
            end
            else begin
              if bVerbose>0 then
                if bDebug=0 then MessageAlerte(FBanque.HMsgBox1.Mess.Strings[7])
                            else MessageAlerte(FBanque.HMsgBox1.Mess.Strings[8]) ;
              v_ret:=ERR_DLLMISSING;
            end;
          end;
        end;
      end;
    end;
  end;
  finally
    FBanque.Free;
  end;
  Result:=v_ret;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD+XMG
Créé le ...... : 15/05/2000
Modifié le ... : 29/10/2003
Description .. : Convertit une date au format TDATETIME avec gestion an
Suite ........ : 2000
Suite ........ : on specifique le format d'entreé de la date.
Mots clefs ... : DATE;CONVERTION;
*****************************************************************}
function ConvertDate(DateChar,Format : string ) : TDateTime ; overload
var
  Cent,Year,Month,Day  : word ;
  pCent,pYear,pmonth,pDay : byte ;
begin
   Format:=uppercase(Format) ;
   pCent:=pos('AAAA',Format) ;
   if pCent>0 then
      Format:=copy(Format,1,pCent-1)+'CC'+Copy(Format,pcent+2,length(Format)) ;
   pYear:=pos('AA',Format) ;
   pMonth:=pos('MM',Format) ;
   pDay:=pos('DD',Format) ;
   Cent:=0 ; Year:=0 ; Month:=0 ; Day:=0 ;
   if (pyear>0) and (pmonth>0) and (pDay>0) then begin
      if pCent>0 then
         Cent:=Valeuri(copy(DateChar,pCent,2)) ;
      Year:=StrToInt(copy(DateChar,pYear,2)) ;
      if pcent>0 then
         Year:=Year+Cent*100
      else begin
         if (Year>90) then
            Year:=Year+1900
         else
            Year := Year + 2000;
      End ;
      Month:=StrToInt(copy(DateChar,pMonth,2));
      Day:=StrToInt(copy(DateChar,pDay,2));
   End ;
   Result:= EncodeDate(Year,Month,Day);
end;

function ConvertDate(DateChar : string) : TDateTime ; overload ; 
var Year,Month,Day : word ;
begin
Year  := StrToInt(copy(DateChar,5,2));
if ( Year > 90 )  then Year := Year + 1900 else Year := Year + 2000;
Month:=StrToInt(copy(DateChar,3,2));
Day:=StrToInt(copy(DateChar,1,2));
Result:= EncodeDate(Year,Month,Day);
end;

end.
