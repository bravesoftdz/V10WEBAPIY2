unit EskerInterface;

interface

uses
{$IFDEF EAGLCLIENT}
      Maineagl,
{$ELSE}
      Fe_Main,
{$ENDIF}
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,Types,
  StdCtrls, WSEskerUtil, ExtCtrls, ComCtrls, Grids, UTob, HMsgBox, HCtrls, HEnt1, LicUtil,
  DateUtils,ParamSoc, SubmissionService;

function LoginEsker : Boolean;
procedure EnvoiEsker (Quoi : integer;TobInfosEnvoi : TOB;Document : string;TobAdr,TobCont : Tob);
function SendEsker(Quoi: integer;TobInfos: Tob;Document : string;TobAdr : Tob): SubmissionResult;
procedure LogoutEsker;
procedure QueryEsker (Identifiant : SubmissionResult);

const
     TexteMessage : Array[1..3] of String = (
          {1} 'Problème de connexion au service FlyDoc'
          {2} ,'Impossible de créer les bindings'
          {3} ,'Le document envoyé est en échec : motif %s'
           );

implementation

uses InvokeRegistry, SessionService, QueryService, math, shellapi
;

var
    CurrSessionService : SessionServiceSoap;
    bindings  : BindingResult;
    authenticatedSession : LoginResult;
    SubmitService : SubmissionServiceSoap;

function CreateValue(AttributeName, AttributeValue: string) : Var_;
begin
  result := Var_.Create();
  result.attribute := AttributeName;
  result.simpleValue := AttributeValue;
  result.type_ := TYPE_STRING;
end;    

function FIleToByteArray( const FileName : string ) : TByteDynArray;
const BLOCK_SIZE=1024;
var BytesRead, BytesToWrite, Count : integer;
 F : FIle of Byte;
 pTemp : Pointer;
begin
 AssignFile( F, FileName );
 Reset(F);
try
 Count := FileSize( F );
 SetLength(Result, Count );
 pTemp := @Result[0];
 BytesRead := BLOCK_SIZE;
 while (BytesRead = BLOCK_SIZE ) do
 begin
  BytesToWrite := Min(Count, BLOCK_SIZE);
  BlockRead(F, pTemp^, BytesToWrite , BytesRead );
   pTemp := Pointer(LongInt(pTemp) + BLOCK_SIZE);
  Count := Count-BytesRead;
 end;
finally
  CloseFile( F );
 end;
end;

procedure ByteArrayToFIle(    const ByteArray : TByteDynArray;
const FileName : string );
var Count : integer;
 F : FIle of Byte;
 pTemp : Pointer;
begin
 AssignFile( F, FileName );
 Rewrite(F);
 try
  Count := Length( ByteArray );
  pTemp := @ByteArray[0];
  BlockWrite(F, pTemp^, Count );
 finally
  CloseFile( F );
 end;
end;

function ReadFile(filename: string) : WSFile;
begin
  result := WSFile.Create;
  result.name_ := filename;
  result.mode := MODE_INLINED;
  result.content := FIleToByteArray(filename);
end;

procedure WriteFile(filename: string ; content: TByteDynArray);
begin
  ByteArrayToFile( content, filename );
end;

function Login : integer;
var InfosMP,Compte,MotDePasse,stConnexion,Enregistrer : string;
begin
  CurrSessionService := Nil;
  authenticatedSession := Nil;
  bindings := Nil;
  Result := -1;
  try
    CurrSessionService := GetSessionServiceSoap(false, '');  //connexion service
    if Assigned(CurrSessionService) then
    begin
      InfosMP := GetSynRegKey('RTCONNEXIONESKER','',TRUE);
      if InfosMP <> '' then
        begin
        Compte := ReadTokenSt (InfosMP);
        MotDePasse := ReadTokenSt (InfosMP);
        MotDePasse := DeCryptageSt (MotDePasse);
        end
      else
        begin
        InfosMP := AglLanceFiche ('RT','RTSAISIEMPESKER','','','');
        if InfosMP <> '' then
          begin
            Compte := ReadTokenSt (InfosMP);
            MotDePasse := ReadTokenSt (InfosMP);
            Enregistrer := ReadTokenSt (InfosMP);
          end;
        end;
      if Compte <> '' then
        begin
          try
            authenticatedSession := CurrSessionService.Login(Compte, MotDePasse);
            if Assigned(authenticatedSession) then
            begin
              result := 1;
              if ENREGISTRER = '1' then
                begin
                  stConnexion := Compte + ';' + CryptageSt(MotDePasse);
                  SaveSynRegKey('RTCONNEXIONESKER',stConnexion,TRUE);
                end;
            end;
          except
            on E: Exception do
            PGIError('Erreur de login : ' + E.Message, '');
          end;
        end
      else result := 0;
    end;
  except
    on E: Exception do
    PGIError('Erreur de connexion au service : ' + E.Message, '');
  end;
end;

procedure Logout;
var  shv : SessionHeaderValue;
begin
 	shv := SessionHeaderValue.Create;
  shv.SessionID :=  authenticatedSession.sessionID;
  (CurrSessionService as ISoapHeaders).Send(TSoapHeader(shv));
  try
    CurrSessionService.Logout;
  except
  end;
  FreeAndNil(shv);
end;

procedure GetBindings;
begin
  if Assigned(authenticatedSession) then
  begin
    try
      bindings := CurrSessionService.GetBindings('');
      if assigned(bindings) then
      begin
      end;
    except
      on E: Exception do
      PGIError('Erreur de bindings : ' + E.Message, '');
    end;
  end;
end;

function SendEsker(Quoi: integer;TobInfos: Tob;Document : string;TobAdr : Tob): SubmissionResult;
var
//  resource: WSFile;
  headers: ISOAPHeaders;
  shv : SessionHeaderValue;
  sr: SubmissionResult;
  t: SubmissionTransport;
  vs: TVars;
  ats: TAttachments;
  at: TAttachment;
  ChpAdresse : string;
  stDateEnvoi : string;
begin
  Result := Nil;
  SubmitService := Nil;
  try
  SubmitService := GetSubmissionServiceSoap(false, bindings.submissionServiceLocation);
    if assigned(SubmitService) then
    begin
      headers := (CurrSessionService as ISOAPHeaders);
      try
        shv := SessionHeaderValue.create();
        shv.SessionID :=  authenticatedSession.sessionID;
        headers.Send(TSoapHeader(shv));

        headers := (SubmitService as ISOAPHeaders);
        headers.Send(TSoapHeader(shv));
        try
          t := SubmissionTransport.Create();
          at := TAttachment.Create();
  //        at.inputFormat := uppercase(ExtractFileExt('c:\mailing.doc'));
  //        at.sourceAttachment := ReadFile('c:\mailing.doc');//resource
          if Document <> '' then
          begin
            at.inputFormat := uppercase(ExtractFileExt(Document));
            at.sourceAttachment := ReadFile(Document);//resource
          end;
          case Quoi of
            0:  //MailOnDemand  (Courrier)
                begin
                t.transportName := 'MODEsker';
                // Specifies MailOnDemand variables (see documentation for their meanings)
                t.nVars := 14;
                SetLength(vs,t.nVars);
                vs[0] := CreateValue('Subject', TobInfos.GetString('NOMENVOI'));
                vs[1] := CreateValue('FromName', TobInfos.GetString('NOMEMETTEUR'));
                vs[2] := CreateValue('FromCompany', TobInfos.GetString('NOMSOCIETE'));
                ChpAdresse := TobAdr.GetString('ADR_LIBELLE') +#13#10+ TobAdr.GetString('C_PRENOM') + ' ' + TobAdr.GetString('C_NOM') ;
                ChpAdresse := ChpAdresse +#13#10+ TobAdr.GetString('ADR_ADRESSE1') +#13#10+ TobAdr.GetString('ADR_ADRESSE2') ;
                ChpAdresse := ChpAdresse +#13#10+ TobAdr.GetString('ADR_ADRESSE3') +#13#10+ TobAdr.GetString('ADR_CODEPOSTAL') + ' ' + TobAdr.GetString('ADR_VILLE');
                ChpAdresse := ChpAdresse +#13#10+ RechDom('TTPAYS',TobAdr.GetString('ADR_PAYS'),False) ;
                vs[3] := CreateValue('ToBlockAddress', ChpAdresse);
                if TobInfos.GetString('COULEUR') = 'X' then vs[4] := CreateValue('Color', 'Y')
                else vs[4] := CreateValue('Color', 'N');
                if TobInfos.GetString('PORTEADRESSE') = 'X' then vs[5] := CreateValue('Cover', 'Y')
                else vs[5] := CreateValue('Cover', 'N');
                if TobInfos.GetString('RECTOVERSO') = 'X' then vs[6] := CreateValue('BothSided', 'Y')
                else vs[6] := CreateValue('BothSided', 'N');
                if TobInfos.GetString('TAILLEENV') = '001' then vs[7] := CreateValue('Envelop', 'C4')
                else if TobInfos.GetString('TAILLEENV') = '003' then vs[7] := CreateValue('Envelop', 'C6DW')
                else vs[7] := CreateValue('Envelop', 'C6');
                vs[8] := CreateValue('MaxSheets', '5');
                if TobInfos.GetString('AFFRANCHISSEMENT') = '002' then vs[9] := CreateValue('StampType', 'E')
                else if TobInfos.GetString('AFFRANCHISSEMENT') = '003' then vs[9] := CreateValue('StampType', 'En')
                else if TobInfos.GetString('AFFRANCHISSEMENT') = '004' then vs[9] := CreateValue('StampType', 'PI')
                else vs[9] := CreateValue('StampType', 'U');
                vs[10] := CreateValue('MaxRetry', '3');
                stDateEnvoi := FormatDateTime ('yyyy-mm-dd hh:mm:ss',strToDate(TobInfos.GetString('DATEENVOI')));
                vs[11] := CreateValue('DeferredDateTime', stDateEnvoi);
                if GetParamSocSecur('SO_NPAI',True) = True then vs[12] := CreateValue('AskReturnedMail', 'Y')
                else vs[12] := CreateValue('AskReturnedMail', 'N');
                if TobInfos.GetString('VALIDATION') = 'Y' then vs[13] := CreateValue('NeedValidation', '1')
                else vs[13] := CreateValue('NeedValidation', '0');
                t.nAttachments := 1;
                SetLength(ats,t.nAttachments);
                ats[0] := at;
                end;
            1: //fax
                begin
                t.transportName := 'Fax';
                t.nVars := 10;
                SetLength(vs,t.nVars);
                vs[0] := CreateValue('Subject',TobInfos.GetString('OBJET'));
                vs[1] := CreateValue('FaxNumber','+33' + TobAdr.GetString('FAXCONTACT'));
                vs[2] := CreateValue('Message',TobInfos.GetString('TEXTESAISI'));
                vs[3] := CreateValue('FromName',TobInfos.GetString('NOMEMETTEUR'));
                vs[4] := CreateValue('FromCompany',TobInfos.GetString('NOMSOCIETE'));
                vs[5] := CreateValue('FromFax','+33'+ TobInfos.GetString('FAXSOCIETE'));
                vs[6] := CreateValue('ToName',TobAdr.GetString('NOMCONTACT'));
                vs[7] := CreateValue('ToCompany',TobAdr.GetString('NOMSOCCONTACT'));
                stDateEnvoi := FormatDateTime ('yyyy-mm-dd hh:mm:ss',strToDate(TobInfos.GetString('DATEENVOI')));
                vs[8] := CreateValue('DeferredDateTime', stDateEnvoi);
                if TobInfos.GetString('VALIDATION') = 'Y' then vs[9] := CreateValue('NeedValidation', '1')
                else vs[9] := CreateValue('NeedValidation', '0');
                //vs[8] = CreateValue('CoverTemplate','Default.rtf');
                t.nAttachments := 1;
                SetLength(ats,t.nAttachments);
                ats[0] := at;
                end;
            2: //email
                begin
                t.transportName := 'Mail';
                // Specifies smtp variables (see documentation for their meanings)
                t.nVars := 11;
                SetLength(vs,t.nVars);
                vs[0] := CreateValue('Subject', TobInfos.GetString('OBJET'));
                vs[1] := CreateValue('EmailAddress', TobAdr.GetString('EMAILCONTACT'));
                vs[2] := CreateValue('Message', TobInfos.GetString('TEXTESAISI'));
                vs[3] := CreateValue('FromName', TobInfos.GetString('NOMEMETTEUR'));
                vs[4] := CreateValue('FromCompany', TobInfos.GetString('NOMSOCIETE'));
                vs[5] := CreateValue('FromAddress', TobInfos.GetString('EMAILEMETTEUR'));
                vs[6] := CreateValue('ToName', TobAdr.GetString('NOMCONTACT'));
                vs[7] := CreateValue('ToCompany', TobAdr.GetString('NOMSOCCONTACT'));
                if TobInfos.GetString('IMPORTANCE') = '0' then vs[8] := CreateValue('EmailImportance', 'L')
                else if TobInfos.GetString('IMPORTANCE') = '2' then vs[8] := CreateValue('EmailImportance', 'H')
                else vs[8] := CreateValue('EmailImportance', 'N');
                stDateEnvoi := FormatDateTime ('yyyy-mm-dd hh:mm:ss',strToDate(TobInfos.GetString('DATEENVOI')));
                vs[9] := CreateValue('DeferredDateTime', stDateEnvoi);
                if TobInfos.GetString('VALIDATION') = 'Y' then vs[10] := CreateValue('NeedValidation', '1')
                else vs[10] := CreateValue('NeedValidation', '0');
                // Specify a text attachment to append to the email.
                // The attachment content is inlined in the transport description
                // We also request the attachment to be converted to pdf format and named 'YourInformations.pdf'
                // before being sent
                t.nAttachments := 1;
                SetLength(ats,t.nAttachments);
                ats[0] := at;
                end;
            3: //sms
                begin
                t.transportName := 'SMS';
                t.nVars := 5;
                SetLength(vs,t.nVars);
                // Specifies SMS variables (see documentation for their meanings)
  {              vs[0] := CreateValue('Subject', 'Sample SMS - Test Epz');
                vs[1] := CreateValue('FromName', 'John DOE');
                vs[2] := CreateValue('SMSNumber', '+33685567403');
                vs[3] := CreateValue('Message', 'Everything''s OK');  }
                vs[1] := CreateValue('FromName', TobInfos.GetString('NOMEMETTEUR'));
                vs[2] := CreateValue('SMSNumber', '+33'+TobInfos.GetString('TEL'));
                vs[3] := CreateValue('Message', TobInfos.GetString('TEXTESAISI'));
                if TobInfos.GetString('VALIDATION') = 'Y' then vs[4] := CreateValue('NeedValidation', '1')
                else vs[4] := CreateValue('NeedValidation', '0');
                t.nAttachments := 0;
  (*
                t.nAttachments := 1;
                SetLength(ats,t.nAttachments);
                ats[0] := at;
  *)
                end;
            end;
          t.vars := vs;
          t.attachments := ats;

          try
            sr := SubmitService.SubmitTransport(t);
            Result := sr;
          except
            on e: ERemotableException do
              PGIError('Erreur SubmitTransport : ' + E.Message, '');
          end;
          FreeAndNil(t);
        except
//          on e: ERemotableException do
          on e: Exception do
            PGIError('Erreur SubmitTransport : ' + E.Message, '');
        end;
        FreeAndNil(shv);
      except
//        on e: ERemotableException do
        on e: Exception do
          PGIError('Erreur Session (send) : ' + E.Message, '');
      end;
    end;
  except
    on E: Exception do
    PGIError('Erreur de SubmissionService : ' + E.Message, '');
  end;
end;

function LoginEsker : Boolean;
var Resultat : integer;
begin
  Result := False;
  Resultat := Login;
  if Resultat = -1 then
  begin
    PGIBox (TexteMessage[1],'Attention');
    Exit;
  end;
  if Resultat = 0 then Exit;
  GetBindings;
  if Not assigned(bindings) then
  begin
    PGIBox (TexteMessage[2],'Attention');
    Exit;
  end;
  Result := True;
end;

procedure EnvoiEsker (Quoi : integer;TobInfosEnvoi : TOB;Document : string;TobAdr,TobCont : Tob);
var Resultat : integer;
begin
  Resultat := Login;
  if Resultat = -1 then
  begin
    PGIBox (TexteMessage[1],'Attention');
    Exit;
  end;
  if Resultat = 0 then Exit;
  GetBindings;
  if Not assigned(bindings) then
  begin
    PGIBox (TexteMessage[2],'Attention');
    Exit;
  end;
//  Send(Quoi,TobInfosEnvoi,Document,TobAdr,TobCont);
  Logout;
end;

procedure LogoutEsker;
begin
  Logout;
end;

procedure QueryEsker (Identifiant : SubmissionResult);
var headers: ISOAPHeaders;
    shv : SessionHeaderValue;
    QueryService : QueryServiceSoap;
    request: QueryRequest;
    qresult: QueryResult;
    qAttachments : Attachments2;
    iVar, i, j: integer;
    qstate : integer;
    qstatus {,qdate} : string;
    buf: TByteDynArray;
    Path: array[0..255] of Char;
    sTempFileName : string;
begin
  try
    queryService := GetQueryServiceSoap(false, bindings.queryServiceLocation);
    headers := (CurrSessionService as ISOAPHeaders);
    try
      SourisSablier;
      shv := SessionHeaderValue.create();
      shv.SessionID :=  authenticatedSession.sessionID;

      headers := (queryService as ISOAPHeaders);
      headers.Send(TSoapHeader(shv));

      request := QueryRequest.Create;
      request.attributes := 'State,ShortStatus,CompletionDateTime';
      request.filter := 'msn=' + IntToStr(Identifiant.transportID);
      request.nItems := 1;

      qstate := 0;
      While true  do
      begin
        try
        // Ask the Application Server
        qresult := queryService.QueryFirst(request);
        except
        qresult := Nil;
        // traiter SoapException
        end ;
        if (qresult <> Nil) and (qresult.nTransports  = 1) then //trouvé
        begin
          // Hopefully, we found it
          // Parse the returned variables
          for iVar :=0 to  qresult.transports[0].nVars-1 do
          begin
            if(LowerCase(qresult.transports[0].vars[iVar].attribute) = 'state' ) then
            qstate := Valeuri(qresult.transports[0].vars[iVar].simpleValue)
            else if (LowerCase(qresult.transports[0].vars[iVar].attribute) = 'longstatus' ) then
            qstatus := qresult.transports[0].vars[iVar].simpleValue;
{            else if (LowerCase(qresult.transports[0].vars[iVar].attribute) = 'completiondatetime' ) then
            qdate := qresult.transports[0].vars[iVar].simpleValue;  }

          end;

          if( qstate >= 70 ) then
          break;

        end
        else // pas trouvé echec de la query… resoumettre ?????
        begin
  //      Console.WriteLine("Error !! MailOnDemand not found in database");
          PGIError('Identifiant non trouvé dans la base' , '');
          qstate := 0;
          break;
        end;
  //        Console.WriteLine("MailOnDemand pending...");
        Sleep(5000);
        headers.Send(TSoapHeader(shv));
      end ; // Fin While true

      if (qstate >= 70) AND (qstate <= 100 ) then
        begin
        try
          headers.Send(TSoapHeader(shv));
          qAttachments := queryService.QueryAttachments(Identifiant.transportID, FILTER_CONVERTED, MODE_ON_SERVER) ;
          for i:=0 to qAttachments.nAttachments-1 do
          begin
            for j:=0 to qAttachments.attachments[i].nConvertedAttachments-1 do
            begin

              headers.Send(TSoapHeader(shv));
              buf := queryService.DownloadFile(qAttachments.attachments[i].convertedAttachments[j]);

              FillChar(Path, Sizeof(Path), 0);
              GetTempPath(255, Path);
              sTempFileName := string(Path)+'\'+qAttachments.attachments[i].convertedAttachments[j].name_;
              WriteFile(string(Path)+'\'+qAttachments.attachments[i].convertedAttachments[j].name_,buf);
    //          ShowGedFileViewer(sTempFileName, False, 'Pré-Visualisation du document', True, False, True, True);  ne marche pas pr les .TIF
              SourisNormale ;
              ShellExecute (0, pchar ('open'), pchar (sTempFileName), nil, nil, SW_RESTORE);
            end;
          end;
        except
          on E: Exception do
          begin
            SourisNormale ;
            PGIError('Récupération des informations : ' + E.Message, '');
          end;
        end;
        end
      else
        begin
        SourisNormale ;
        PGIBox(Format(TexteMessage[3], [qstatus]),'Prévisualisation');
        end;
       ;
    finally
      FreeAndNil(request);
      FreeAndNil(shv);
    end;
  except
    on E: Exception do
    PGIError('Erreur QueryService : ' + E.Message, '');
  end;
end;

end.
