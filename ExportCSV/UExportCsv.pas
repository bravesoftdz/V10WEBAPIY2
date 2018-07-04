unit UExportCsv;

interface

uses
  Windows, Messages,
  SysUtils,
  Classes, Graphics,
  Controls, Forms,
  Dialogs, StdCtrls,
  HENT1,
  HCtrls,
  inifiles,
  Hqry,
  DB, ADODB, Variants,
  CBPPath,
  ULog,
  IdFTP,
  IdAntiFreeze,
  IdFTPCommon
  ;
const
  ADDRSERVER = '94.23.179.245';

type
  TForm1 = class(TForm)
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    fDataConnect : string;
    FIniFile: TIniFile;
    TheLog : string;
    //
    fIpServ : string;
    fServer : string;
    fDataBase : string;
    fSeparator : string;
    fExportName : string;
    flastDate : TDateTime;
    fnewDate : TdateTime;
    fidFTP : TIdFTP;
    fuser,fPassword : string;
    fIdAntiFreeze1: TIdAntiFreeze;
    TheFileExp,TheFileSav : String;
    //
    procedure TraiteSection (NomSection : String);
    procedure ExporteData;
    function TraduitStringToDate(DateStr: string): TdateTime;
    function DecodeBlocNote(const RTF: string; ReplaceLineFeedWithSpace,DoTrimLeft: Boolean; TrailAfter: integer): string;
    function EnleveForbidenChar (TheChaine : string) : string;
    function ConstitueMessageErr (TheMessage : string) : string;
    function ConstitueMessageOk: string;
    procedure SauvegardeFichier;
    procedure DeleteFichier;
  public
    { Déclarations publiques }
  end;

var
  Form1: TForm1;

implementation

uses IdBaseComponent, DateUtils;

{$R *.dfm}

{ TForm1 }


function TForm1.ConstitueMessageErr(TheMessage: string): string;
  function nettoie (UnMessage : string) : string;
  begin
    result := StringReplace(UnMessage,#$D#$A,'',[rfReplaceAll, rfIgnoreCase]);
  end;
begin
  result := DatetimeToStr(NowH)+ ' Envoie du fichier '+fExportName+' sur '+fIpServ+' ERREUR --> ('+nettoie(TheMessage)+')' ;
end;


function TForm1.ConstitueMessageOk : string;
begin
  result := DatetimeToStr(NowH)+ ' Envoie du fichier '+fExportName+' sur '+fIpServ+'  OK' ;
end;


function TForm1.DecodeBlocNote(const RTF: string;ReplaceLineFeedWithSpace,DoTrimLeft:Boolean;TrailAfter:integer): string;
var
  n,i,x:Integer;
  GetCode:Boolean;
  Code:String;
  ThisChar,LastChar:Char;
  Group:Integer;
  Skip:Boolean;

  procedure ProcessCode;
  begin
    if ThisChar in ['\',' ',#13,#10] then
    begin
      if (Code='\par') or (Code='\line') or (Code='\pard') then
        begin
          if ReplaceLineFeedWithSpace
          then Result:=Result+' '
          else Result:=Result+#9;
          GetCode:=FALSE;
          skip:=TRUE;
        end;
      if Code='\tab' then
      begin
        Result:=Result+' ';
        GetCode:=FALSE;
        skip:=TRUE;
      end;
    end;
    if ((Length(Code)=4) and ((Code[1]+code[2])='\''')) then
    begin
      Result:=Result+Chr(strtoint('$'+code[3]+code[4]));
      GetCode:=FALSE;
      skip:=TRUE;
    end;
  end;

begin
  try
    // length is always shorter
    n:=Length(RTF);
    Result:='';
 
    // are we getting a code?
    GetCode:=FALSE;
    // indent level of curlies
    Group:=0;
    // allow for escaped characters
    LastChar:=#0;
    // parse each character
    x:=1;
    while x<=n do
      begin
        Skip:=FALSE;
        ThisChar:=RTF[x];
        case ThisChar of
          '{':if LastChar<>'\'
              then begin // Curly start
                     inc(Group);
                     skip:=TRUE;
                   end
              else GetCode:=FALSE;
          '}':if LastChar<>'\'
              then begin // Curly end
                     dec(Group);
                     skip:=TRUE;
                   end
              else GetCode:=FALSE;
          '\':if LastChar<>'\'
              then begin // Code start
                     if GetCode then ProcessCode;
                     Code:='';
                     GetCode:=TRUE;
                   end
              else GetCode:=FALSE;
          ' ':begin
                if GetCode then
                begin
                  ProcessCode;
                  GetCode:=FALSE;
                  skip:=TRUE;
                end;
            end;
 
          #10:
            begin
              if GetCode then
                begin
                  ProcessCode;
                  GetCode:=FALSE;
                  skip:=TRUE;
                end;
              if Group>0 then
              begin
                skip:=TRUE;
              end;
            end;
          #13:
            begin
              if GetCode then
                begin
                  ProcessCode;
                  GetCode:=FALSE;
                  skip:=TRUE;
                end;
              if Group>0 then
              begin
                skip:=TRUE;
              end;
            end;
        end;
        if not GetCode then
          begin
            if (not Skip) and (Group <= 1) then
            Result:=Result+ThisChar;
          end
        else begin
            Code:=Code+ThisChar;
            ProcessCode;
          end;
        LastChar:=ThisChar;
        Inc(x,1);
      end;
 
    // remove trailing spaces and cr/lf
    n:=Length(Result);
    while ((n>0) and (Result[n]<' ')) do
      dec(n);
    if n>0 then
      SetLength(Result,n);
 
    if DoTrimLeft then
    begin
      i := Length(Result);
      n := 1;
      while (n <= i) and (Result[n] <= ' ') do Inc(n);
      Result := Copy(Result, n, Maxint);
    end;
    result:=TrimLeft(Result);

    if TrailAfter>0 then
    begin
      if Length(Result)>TrailAfter then
      begin
        SetLength(Result,TrailAfter);
        result:=result+'...';
      end;
    end;
  except
    on e:exception do
  end;
end;

procedure TForm1.DeleteFichier;
begin
  DeleteFile(TheFileExp);
end;

function TForm1.EnleveForbidenChar(TheChaine: string): string;
begin
  result := StringReplace (TheChaine,';',',',[rfReplaceAll, rfIgnoreCase]);
end;

procedure TForm1.ExporteData;
var Conn : string;
    SQl : String;
    QQ: TADOQuery;
    UneLigne : WideString;
    myFile : TextFile;
begin
  //
  AssignFile(myFile, TheFileExp);
  ReWrite(myFile);
  fnewDate := NowH;
  Conn := 'Provider=SQLOLEDB;Password=ADMIN;User ID=ADMIN;Initial Catalog='+fDatabase+';Data Source='+fServer+';';
  Sql := 'SELECT GP_DATEPIECE,'+
         'GP_AFFAIRE2,'+
         'GP_NUMERO,'+
         'T_AUXILIAIRE,'+
         'T_LIBELLE,'+
         'T_PRENOM,'+
         'T_BLOCNOTE,'+
         'P1.GPA_LIBELLE,'+
         'P1.GPA_ADRESSE1,'+
         'P1.GPA_ADRESSE2,'+
         'P1.GPA_VILLE,'+
         'GP_REFINTERNE,'+
         'GP_DATELIBREPIECE2,'+
         'GP_REFEXTERNE,'+
         'GP_LIBREPIECE1,'+
         'AFF_REFEXTERNE '+
         'FROM PIECE '+
         'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIRE AND AFF_TIERS=GP_TIERS '+
         'LEFT JOIN TIERS ON T_TIERS=GP_TIERS '+
         'LEFT JOIN PIECEADRESSE P1 ON P1.GPA_NATUREPIECEG=GP_NATUREPIECEG AND P1.GPA_SOUCHE=GP_SOUCHE AND P1.GPA_NUMERO=GP_NUMERO AND P1.GPA_INDICEG=GP_INDICEG AND P1.GPA_NUMLIGNE=0 AND P1.GPA_TYPEPIECEADR=GP_NUMADRESSELIVR '+
         'WHERE GP_NATUREPIECEG=''DBT'' AND GP_LIBREPIECE1=''ENC''';
	QQ := TADOQuery.Create(nil);
  QQ.ConnectionString := Conn;
  QQ.SQL.clear;
  QQ.SQL.Add(Sql);
  TRY
    QQ.Open;
    if not QQ.Eof then
    begin
      UneLigne := 'Date Devis'+fSeparator+
                  'N° Affaire'+fSeparator+
                  'N° Devis'+fSeparator+
                  'Code Client'+fSeparator+
                  'Nom Client'+fSeparator+
                  'Chargé Client'+fSeparator+
                  'Bloc Note'+fSeparator+
                  'Nom Locataire'+fSeparator+
                  'Adresse 1'+fSeparator+
                  'Adresse 2'+fSeparator+
                  'Ville'+fSeparator+
                  'N° Cde Client'+fSeparator+
                  'Date Livraison'+fSeparator+
                  'Intervenant'+fSeparator+
                  'Statut DSL'+fSeparator+
                  'Lot';
      Writeln(myFile,uneLigne);
      QQ.First;
      while not QQ.Eof do
      begin
        UneLigne := QQ.findfield('GP_DATEPIECE').AsString+fSeparator+
        						QQ.findfield('GP_AFFAIRE2').AsString+fSeparator+
        						QQ.findfield('GP_NUMERO').AsString+fSeparator+
                    QQ.findfield('T_AUXILIAIRE').AsString+fSeparator+
        						EnleveForbidenChar(QQ.findfield('T_LIBELLE').AsString)+fSeparator+
                    QQ.findfield('T_PRENOM').AsString+fSeparator+
                    DecodeBlocNote(QQ.findfield('T_BLOCNOTE').AsString,false,true,0)+fSeparator+
                    QQ.findfield('GPA_LIBELLE').AsString+fSeparator+
                    QQ.findfield('GPA_ADRESSE1').AsString+fSeparator+
                    QQ.findfield('GPA_ADRESSE2').AsString+fSeparator+
                    QQ.findfield('GPA_VILLE').AsString+fSeparator+
                    QQ.findfield('GP_REFINTERNE').AsString+fSeparator+
                    QQ.findfield('GP_DATELIBREPIECE2').AsString+fSeparator+
                    QQ.findfield('GP_REFEXTERNE').AsString+fSeparator+
                    QQ.findfield('GP_LIBREPIECE1').AsString+fSeparator+
                    QQ.findfield('AFF_REFEXTERNE').AsString;
        Writeln(myFile,uneLigne);
        QQ.Next;
      end;
    end;
  	QQ.Close;
  finally
    CloseFile(MyFile);
    QQ.Free;
  end;

end;

procedure TForm1.FormCreate(Sender: TObject);
var fini : string;
begin
  fDataConnect := GetFromRegistry(HKEY_LOCAL_MACHINE, 'SOFTWARE\MICROSOFT\WINDOWS\CurrentVersion\Explorer\Shell Folders', 'Common AppData', fDataConnect, True);
  fini := IncludeTrailingBackslash(fDataConnect)+'LSE\CONFIG\EXPORT.INI';
  TheLog := IncludeTrailingBackslash(fDataConnect)+'LSE\LOG';
  //
  FIniFile := TIniFile.Create (fini);
end;

procedure TForm1.FormShow(Sender: TObject);
var Sections : TStringList;
    indice : integer;
begin
  Sections := TStringList.Create;
  FiniFile.ReadSections (Sections);
  for Indice := 0 to Sections.Count -1 do
  begin
    TraiteSection(Uppercase(Sections.strings[Indice]));
  end;
  Sections.Clear;
  Sections.free;
  close;
end;

procedure TForm1.SauvegardeFichier;
var StSt,Str : string;
    IIP : integer;
    YY,MM,DD,HH,MN,SS,MS : word;
begin
  StSt := ExtractFileName(fExportName);
  IIP := Pos('.',StSt);
  DecodeDateTime ( fnewDate,YY,MM,DD,HH,MN,SS,MS);
  Str := Format('%s%.02d%.02d%.02d%.02d%.02d%.02d',[copy(StSt,1,IIP-1),DD,MM,YY,HH,MN,SS])+'.sav';
  TheFileSav := IncludeTrailingBackslash(fDataConnect)+'LSE\DONE\'+Str;
  MoveFile(PansiChar(TheFileExp),PAnsiChar(TheFileSav));
end;

function TForm1.TraduitStringToDate (DateStr : string) : TdateTime;
begin
  if not Hent1.IsValidDate(DateStr) then result := DEBUTANNEE(NowH) 
                              else result := StrToDateTime(DateStr);
end;

procedure TForm1.TraiteSection(NomSection: String);
var LesChamps : TStringList;
    Indice : integer;
    Ligne,Champ,parametres : string;
    LeNom : string;
begin
  fIpServ := ADDRSERVER;
  TheFileSav := '';
  LeNom := UpperCase (NomSection);
  lesChamps := TStringList.Create;
  TRY
    //
    FiniFile.ReadSectionValues (LeNom,LesChamps);
    if lesChamps.Count > 0 then
    begin
      // Ajout d'un descripteur de fichier
      for Indice :=0 to lesChamps.Count -1 do
      begin
        // Recup des infos
        Ligne := lesChamps.Strings [Indice];
        Champ := Copy(Ligne,1,Pos('=',Ligne)-1);
        Parametres := Copy(Ligne,Pos('=',Ligne)+1,Length(Ligne));
        // desciption du fichier
        if UpperCase(Champ) = 'SERVER' then fServer:= Parametres
        else if UpperCase(Champ) = 'IP' then fIpServ := Parametres
        else if UpperCase(Champ) = 'DATABASE' then fDataBase:= Parametres
        else if UpperCase(Champ) = 'SEPARATOR' then fSeparator:= Parametres
        else if UpperCase(Champ) = 'NAME' then fExportName:= Parametres
        else if UpperCase(Champ) = 'USER' then fUser:= Parametres
        else if UpperCase(Champ) = 'PASSWORD' then fPassword:= Parametres
        else if UpperCase(Champ) = 'LASTDATE' then flastDate := TraduitStringToDate(Parametres);
      end;
      if (fServer = '') or (fDataBase='') or (fSeparator='') or (fExportName='') then exit;
      TheFileExp := IncludeTrailingBackslash(fDataConnect)+'LSE\PROCESSING\'+fExportName;
      ExporteData;
      //
      fIdAntiFreeze1 := TIdAntiFreeze.Create(self);
      //
      fidFTP := TidFtp.Create(self);
      fIdFTP.Host := fIpServ;
      fidFTP.Port := 21;
      fIDFTP.Username := fUser;
      fidFTP.Password := fPassword;
      fidFTP.Passive := true;
      fIdFTP.TransferType := ftBinary;
      fIdFTP.TransferTimeout := 10000;
      fidFTP.ReadTimeout := 10000;
      //
      TRY
        TRY
          fidFTP.connect;
          fIdFTP.Put(TheFileExp,ExtractFileName(TheFileExp));
        FINALLY
          fIDFTP.Disconnect;
          fIdFTP.Free;
          fIdAntiFreeze1.free;
        END;
        fIniFile.WriteString(LeNom,'LASTDATE',DateTimeToStr(fnewDate)) ;
        SauvegardeFichier;
        ecritLog ( TheLog,ConstitueMessageOk, 'EnvoieDSL.log')
      EXCEPT
        on E: Exception do
        begin
          ecritLog ( TheLog,ConstitueMessageErr(E.Message), 'EnvoieDSL.log');
          deleteFichier;
        end;
      END;
    end;
    //
  FINALLY
    LesChamps.Clear;
    LesChamps.free;
  END;
end;



end.
