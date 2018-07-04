unit uWA deprecated;

interface

uses classes,
  uHTTP,
  utob,
  hctrls,
  stdctrls;

type
  cWA = class(tob)
  private
    xxActionMessage: string;
    xxErrorMessage: string;
    xxParle: boolean;
    xxdbss_dll: string;
    procedure setActionMessage(valeur: string);
    procedure setErrorMessage(valeur: string);
    procedure setModeEagl2(valeur: boolean);
    function getModeEagl2: boolean;
    function getaServer2: THttpServer;
    procedure setAServer2(valeur: THttpServer);
  public
    OnCallBack: TNotifyEvent;
    constructor create(); reintroduce; virtual;
    destructor destroy; override;
    function Request(Verb, Complement: string; QueryTob: TOb; StackId1, StackId2: string): TOB;
    procedure setModeEagl(valeur: boolean); //YCP 15/05/07 
    function getModeEagl: boolean; //YCP 15/05/07 
    function getaServer: THttpServer; //YCP 15/05/07 
    procedure setAServer(valeur: THttpServer); //YCP 15/05/07 
    class procedure TraceExecution(st: string);
    class procedure AjouteUneTob(tobDestination, tobSource: TOB; forceNom: string = '');
    class function RecupereUneTob(tobSource: TOB; nom: string): TOB;
    procedure setRetour(var ResponseTOB: TOB; result: boolean; T1: tob; TT: HTStringList; st: string); overload;
    function GetRetour(ResponseTOB: TOB): boolean; overload;
    function GetRetour(ResponseTOB: TOB; T1: tob): boolean; overload;
    function GetRetour(ResponseTOB: TOB; TT: HTStringList): boolean; overload;
    function GetRetour(ResponseTOB: TOB; var st: string): boolean; overload;
    function GetRetour(ResponseTOB: TOB; T1: tob; TT: HTStringList; var st: string): boolean; overload;
    class function si(test: boolean; alors, sinon: string): string;
{$IFDEF EAGLSERVER}
    class procedure MessagesAuClient(Action, fonction, st: string);
{$ENDIF EAGLSERVER}
  published
    property dbss_dll: string read xxdbss_dll write xxdbss_dll;
    property parle: boolean read xxparle write xxparle;
    property ActionMessage: string read xxActionMessage write setActionMessage;
    property ErrorMessage: string read xxErrorMessage write setErrorMessage;
    property ModeEAGL: boolean read getModeEagl2 write setModeEagl2;
    property aServer: THttpServer read getaServer2 write setAServer2;
  private
  xxaServer: THttpServer ; //YCP 15/05/07 
  MSDEModeEAGL: boolean ; //YCP 15/05/07 
  end;

implementation

uses
  Windows,
  forms,
  sysutils,
  uHTTPCS,
  hmsgBox,
  hent1,
  majtable,
  cbpTrace
{$IFDEF EAGLSERVER}
  ,
  esession,
  ehttp
//{$ELSE}
//  ,
//  hdebug
{$ENDIF}

  ;

//var  MSDEModeEAGL: boolean = false;  xxaServer: THttpServer; //YCP 15/05/07 

  {***********A.G.L.***********************************************
  Auteur  ...... : Yann-Cyril PELUD
  Créé le ...... : 10/05/2004
  Modifié le ... :   /  /
  Description .. : Constructeur de l'objet
  Mots clefs ... : COWA
  *****************************************************************}
constructor cWA.create();
begin
  inherited create(className, nil, -1);
  dbss_dll := 'dbss';
  parle := true;
  ModeEAGL:=false ;
  xxaServer:=nil ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Destructeur de l'objet
Mots clefs ... : COWA
*****************************************************************}
destructor cWA.destroy;
begin
  inherited
    //
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Permet de concatener des chaines conditionnées.
Mots clefs ... : cWA
*****************************************************************}
class function cWA.si(test: boolean; alors, sinon: string): string;
begin
  if test then
    result := alors
  else
    result := sinon;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Affiche la trace des actions soit dans le debugger soit dans
Suite ........ : la fenêtre d'execution du serveur
Mots clefs ... : cWA
*****************************************************************}
class procedure cWA.TraceExecution(st: string);
begin
  if (FileExists(ExtractFilePath(Application.ExeName) + '\trace.txt')) then
//{$IFDEF EAGLSERVER}
//    ddWriteLN(st);
//{$ELSE}
    Trace.TraceInformation ('DEBUG', '(--' + st);
    //Debug('(--' + st);
//{$ENDIF EAGLSERVER}
end;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Renseigne le serveur Web Access sur lequel seront
Suite ........ : executées les requêtes.
Mots clefs ... : cWA
*****************************************************************}
procedure cWA.setAServer2(valeur: THttpServer);
begin
  setAserver(valeur);
end;

procedure cWA.setAServer(valeur: THttpServer);
begin
  (*YCP 27/09/05 doit être fait par le créateur if (xxAServer<>nil) and (xxAServer.IsConnected) then
    begin
    {$IFDEF EAGLCLIENT}
    if (xxAServer<>appserver) then
    {$ENDIF EAGLCLIENT}
      begin
      DeconnecteAGLServer(xxAServer);
      DeconnectHttpServer(xxAServer);
      xxAServer:=nil ;
      end ;
    end;*)
  xxaServer := valeur;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Récupère le serveur sur lequel les requêtes seront
Suite ........ : executées.
Mots clefs ... : cWA
*****************************************************************}
function cWA.getaServer2: THttpServer;
begin
  result := getAServer();
end;

function cWA.getaServer: THttpServer;
begin
{$IFDEF EAGLCLIENT}
  if xxaServer = nil then
    result := appserver
  else
    result := xxaServer;
{$ELSE}
  result := xxaServer;
{$ENDIF EAGLCLIENT}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Renseigne le message d'information.
Mots clefs ... : cWA
*****************************************************************}
procedure cWA.setActionMessage(valeur: string);
begin
  if (valeur <> '') and parle then
    PgiInfo(valeur);
  TraceExecution('*** ACTIONMESSAGE' + valeur);
  xxActionMessage := valeur;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Renseigne le message d'erreur
Mots clefs ... : cWA
*****************************************************************}
procedure cWA.setErrorMessage(valeur: string);
begin
  if (valeur <> '') and parle then
    PgiBox(valeur);
  TraceExecution('*** ERRORMESSAGE *** : ' + valeur);
  xxErrorMessage := valeur;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Indique si l'objet doit se comporte en mode Client Web
Suite ........ : Access. Toujours vrai pour EAGLCLIENT, peut être vrai en
Suite ........ : mode 2 tier pour du Web service
Mots clefs ... : cWA
*****************************************************************}
procedure cWA.setModeEagl2(valeur: boolean);
begin
  setModeEagl(valeur)
end;

procedure cWA.setModeEagl(valeur: boolean);
begin
{$IFDEF EAGLCLIENT}
  MSDEModeEAGL := true;
{$ELSE}
  MSDEModeEAGL := valeur;
{$ENDIF EAGLCLIENT}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Indique si l'objet se comporte en mode Client Web
Suite ........ : Access. Toujours vrai pour EAGLCLIENT, peut être vrai en
Suite ........ : mode 2 tier pour du Web service
Mots clefs ... : cWA
*****************************************************************}
function cWA.getModeEagl2: boolean;
begin
  result := getModeEagl();
end;

function cWA.getModeEagl: boolean;
begin
{$IFDEF EAGLCLIENT}
  result := true;
{$ELSE}
  result := MSDEModeEAGL;
{$ENDIF EAGLCLIENT}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 17/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
class procedure cWA.AjouteUneTob(tobDestination, tobSource: TOB; forceNom: string = '');
var
  tmpTOB: tob;
begin
  TmpTOB := tobDestination.FindFirst(['NOM'], [tobSource.NomTable], true);
  if tmpTob = nil then
  begin
    TmpTOB := tob.Create('', tobDestination, -1);
    if (forcenom <> '') then
      TmpTOB.AddChampSupValeur('NOM', forceNom)
    else
      TmpTOB.AddChampSupValeur('NOM', tobSource.NomTable);
  end;
  tobSource.ChangeParent(TmpTOB, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 17/03/2006
Modifié le ... :   /  /
Description .. : Permet de récupérer une tob attaché avec AjouteUneTob
Mots clefs ... :
*****************************************************************}
class function cWA.RecupereUneTob(tobSource: TOB; nom: string): TOB;
var
  i: integer;
begin
  result := nil;
  for i := 0 to tobSource.Detail.Count - 1 do
  begin
    if (tobSource.Detail[i].detail.count > 0)
      and ((nom = '') or (tobSource.Detail[i].getValue('NOM') = nom)) then
    begin
      result := tobSource.Detail[i];
      tobSource.Detail[i].ChangeParent(nil, -1);
      break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Renseigne les éléments de la TOB de retour T1 en fonction
Suite ........ : du resultat.
Suite ........ : Renseigne les champs MESSAGE ou ERROR avec les
Suite ........ : propriétés ActionMessage et ErrorMessage de l'objet.
Mots clefs ... : cWA
*****************************************************************}
procedure cWA.setRetour(var ResponseTOB: TOB; result: boolean; T1: tob; TT: HTStringList; st: string);
var
  i: integer;
  TmpTOB: TOB;
begin
  if result then
  begin
    ResponseTOB.AddChampSupValeur('RESULT', 'TRUE');
    ResponseTOB.AddChampSupValeur('ERROR', '');
    ResponseTOB.AddChampSupValeur('MESSAGE', ActionMessage);

    if T1 <> nil then
    begin
      while T1.Detail.Count > 0 do
        AjouteUneTob(ResponseTOB, T1.Detail[0]);
    end;

    if tt <> nil then
    begin
      for i := 0 to tt.Count - 1 do
      begin
        TmpTOB := tob.Create('TSTRING', nil, -1);
        TmpTOB.AddChampSupValeur('STRING', tt.Strings[i]);
        AjouteUneTob(ResponseTOB, tmpTOB);
      end;
    end;

    TmpTOB := tob.Create('SSTRING', nil, -1);
    TmpTOB.AddChampSupValeur('STRING', st);
    AjouteUneTob(ResponseTOB, tmpTOB);
  end
  else
  begin
    ResponseTOB.AddChampSupValeur('RESULT', 'FALSE');
    ResponseTOB.AddChampSupValeur('ERROR', ErrorMessage);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Récupère les valeurs de retour de la tob réponse d'une
Suite ........ : requete. Positionne les méthdes ActionMessage ou
Suite ........ : ErrorMessage en fonction du résultat. Le cas échéant peu
Suite ........ : afficher ou non le résultat dans une fenèter de message.
Mots clefs ... : cWA
*****************************************************************}
function cWA.GetRetour(ResponseTOB: TOB): boolean;
var
  st: string;
begin
  result := GetRetour(ResponseTOB, nil, nil, st);
end;

function cWA.GetRetour(ResponseTOB: TOB; T1: tob): boolean;
var
  st: string;
begin
  result := GetRetour(ResponseTOB, T1, nil, st);
end;

function cWA.GetRetour(ResponseTOB: TOB; TT: HTStringList): boolean;
var
  st: string;
begin
  result := GetRetour(ResponseTOB, nil, TT, st);
end;

function cWA.GetRetour(ResponseTOB: TOB; var st: string): boolean;
begin
  result := GetRetour(ResponseTOB, nil, nil, st);
end;

function cWA.GetRetour(ResponseTOB: TOB; T1: tob; TT: HTStringList; var st: string): boolean;
var
  i: integer;
  tmpTob: tob;
begin
  result := false;
  if ResponseTOB <> nil then
  begin
    if ResponseTOB.FieldExists('RESULT') then
      result := ResponseTOB.GetValue('RESULT');
    if result then
    begin
      if ResponseTOB.FieldExists('MESSAGE') then
        ActionMessage := ResponseTOB.GetValue('MESSAGE');

      //récupération du HTStringList
      if (tt <> nil) then
      begin
        tmpTob := RecupereuneTob(ResponseTOB, 'TSTRING');
        if (tmpTob <> nil) then
        begin
          for i := 0 to tmpTob.Detail.Count - 1 do
            tt.Add(tmpTob.Detail[i].GetvaLue('STRING'));
          tmpTob.free;
        end;
      end;

      //récupération de la chaine
      tmpTob := RecupereUneTob(ResponseTOB, 'SSTRING');
      if (tmpTob <> nil) then
      begin
        for i := 0 to tmpTob.Detail.Count - 1 do
          st := tmpTob.Detail[i].GetvaLue('STRING');
        tmpTob.free;
      end;

      //récupération de la TOB
      if (T1 <> nil) then
      begin
        tmpTob := RecupereUneTob(ResponseTOB, '');
        if (tmpTob <> nil) then
        begin
          T1.Dupliquer(tmpTob, true, true);
          tmpTob.free;
        end;
      end;

    end
    else
    begin
      if ResponseTOB.FieldExists('ERROR') then
        ErrorMessage := ResponseTOB.GetValue('ERROR');
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 10/05/2004
Modifié le ... :   /  /
Description .. : Execute une requete sur le serveur de l'objet
Mots clefs ... : cWA
*****************************************************************}
function cWA.Request(Verb, Complement: string; QueryTob: TOb; StackId1, StackId2: string): TOB;
var
  it: TOb;
  sa, sf, St: string;
begin
{$IFDEF EAGLSERVER}
//FQ 15973: Suppression des dispatchEx
//  LookUpCurrentSession.HandlePlugin(Verb, Complement, QueryTob, result, ResponseStream);
  LookUpCurrentSession.InvokePlugin(Verb, Complement, QueryTob, result);
{$ELSE}
  result := getaServer().Request(Verb, Complement, QueryTob, StackId1, StackId2);
{$ENDIF}

  if result = nil then
    //St := 'Erreurrrrr'
    St := 'Votre connexion Web Access a été interrompue, vous devez relancer le traitement. #13#10'
      + 'Si le problème persiste, contactez votre administrateur Web Access.'
  else
    st := result.GetValue('ERROR');
  if St <> '' then
    exit;
  if result.FieldExists('MessageCallBack') then
    St := result.GetValue('MessageCallBack')
  else
    St := '';
  //>>>>>>>>>>>>>>>>>>>>
  while (st <> '') do
  begin
    sa := result.GetValue('Action');
    sf := result.GetValue('Function');
    //  msg := result.GetValue('Param');
    if Assigned(OnCallBack) then
      OnCallBack(result);
    //YCP 29/04/05 result.Free;

    iT := TOB.Create('In', nil, -1);
    iT.AddChampSupValeur('Action', sa);
    result.ChangeParent(it, -1); //YCP 29/04/05

{$IFDEF EAGLSERVER}
//FQ 15973: Suppression des dispatchEx
//    LookUpCurrentSession.HandlePlugin('CallBack', sa, iT, result, ResponseStream);
    LookUpCurrentSession.InvokePlugin('CallBack', sa, iT, result);
{$ELSE}
    if getaServer().IsConnected = TRUE then // $$$ JFD/JP 09/08/07 FQ 13340: à compléter: quoi donner à Result exactement??? //YCP 11/09/2007 appserver n'est pas renseigné
       result := getaServer().Request('CallBack', sa, iT, '', '')
    else
       result := nil;
{$ENDIF}

    //Faire le onCallBack
    iT.Free;
    if result = nil then
      //St := 'Erreurrrrr'
      St := 'Votre connexion Web Access a été interrompue, vous devez relancer le traitement. #13#10'
        + 'Si le problème persiste, contactez votre administrateur Web Access.'
    else
      st := result.GetValue('ERROR');
    if St <> '' then
      break;
    if result.FieldExists('MessageCallBack') then
      St := result.GetValue('MessageCallBack')
    else
      St := '';
  end;
end;

{$IFDEF EAGLSERVER}
class procedure cWA.MessagesAuClient(Action, fonction, st: string);
var
  LaSession: tISession;
  Req: TOB;
begin
  laSession := LookUpCurrentSession;
  if not Assigned(LaSession) then
    exit;

  Req := TOB.Create('le callback', nil, -1);
  Req.AddChampSupValeur('ERROR', '');
  Req.AddChampSupValeur('Action', Action);
  Req.AddChampSupValeur('Function', fonction);
  Req.AddChampSupValeur('Param', st);
  Req.AddChampSupValeur('MessageCallBack', 'OK');
  laSession.BroadCastClientNotification(Req, [cnmSelfReply]);
  Req.Free;

  laSession.RunSubServer(fonction);
end;
{$ENDIF EAGLSERVER}

end.

