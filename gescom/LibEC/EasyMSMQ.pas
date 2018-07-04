unit EasyMSMQ;

(*******************************************************************************
                            EXEMPLES D'UTILISATION

Envoi de deux tobs T1 et T2 :

   with TMSMQueue.Create('SERVEUR\fileenvoi') do
   try
     EnqueueMsg('TOB 1', T1.SaveToBuffer(false, false, false));
     EnqueueMsg('TOB 2', T2.SaveToBuffer(false, false, false));
     if SendQueue then ClearQueue;

   finally
     Free;
   end;


Réception d'une série de messages :

   var i : integer;

   with TMSMQueue.Create('SERVEUR\filereception') do
   try
     if ReceiveQueue then
      for i := 0 to MsgQueue.Count-1 do
       if MsgQueue[i].Subject = 'A traiter'
          then Traitement(MsgQueue[i].Body);

   finally
     Free;
   end;

*******************************************************************************)

interface
uses Classes, SysUtils, MSMQ_TLB;

const MSMQDebug : boolean = false;
const MSMQMaxSize : integer = 4000000; // Taille maxi des messages envoyés à MSMQ

type
    TMsg = class
    public
      Subject : string;
      Body : String;
    end;

    TMsgList = class(TList)
    private
      function GetMsg(index : integer) : TMsg;
      function FindFreeSubject(S : String) : String;

    public
      function Ajout(Sujet, Corps : string) : TMsg;
      procedure Suppr(index : integer);
      procedure Clear; override;
      destructor destroy; override;

      function FindMsg(Subject : String) : TMsg;
      function FindMsgIndex(Subject : String) : integer;

      procedure SplitMessages; // Pour les msg de taille > 4Mo
      procedure UnSplitMessages;

      property Msg[index : integer] : TMsg read GetMsg;

    end;

    TOpenQueueEvent = procedure(Sender : TObject; Lecture, Exclusif : boolean) of object;
    TEndTransEvent = procedure(Sender : TObject; Aborted : boolean) of object;
    TMsgEvent = procedure(Sender : TObject; Msg : TMsg) of object;

    TMSMQueue = class
    private
      FOnLookupQueue : TNotifyEvent;
      FOnOpenQueue : TOpenQueueEvent;
      FOnCloseQueue : TNotifyEvent;
      FOnBeginTransaction : TNotifyEvent;
      FOnEndTransaction : TEndTransEvent;
      FOnSendMessage : TMsgEvent;
      FOnMessageSent : TMsgEvent;
      FOnGetNextMessage : TNotifyEvent;
      FOnMessageReceived : TMsgEvent;

      FInfoQueue : IMSMQQueueInfo;
      FDisp : IMSMQTransactionDispenser;
      FMsgOut : IMSMQMessage;
      FMsgQueue : TMsgList;
      FExclusif : boolean;
      FEncrypted : boolean;
      FFirst : boolean;
      FExcept : Exception;

      function LookupQueue(Nom : String) : IMSMQQueueInfo;
      function OuvrirQueue(Lecture, Exclusif : boolean) : IMSMQQueue;
      procedure FermerQueue(LaQueue : IMSMQQueue);
      function LireMessage(LaQueue : IMSMQQueue; Msg : TMsg; var Trans : IMSMQTransaction; OnlyPeek : boolean = false) : boolean;
      procedure EcrireMessage(LaQueue : IMSMQQueue; XAct : IMSMQTransaction; Msg : TMsg);
      function PreparerTransaction : IMSMQTransaction;
      procedure FinirTransaction(XAct : IMSMQTransaction; Abort : boolean = false);

    public
      ErrorMsg : String;

      constructor Create(NomQueue : String); overload;
      constructor Create(InfoQueue : IMSMQQueueInfo); overload;
      destructor Destroy; override;

      procedure CreateQueue;
      procedure DeleteQueue;

      function Send(MList : TMsgList) : boolean;
      function Receive(MList : TMsgList; Transactionnel : boolean = false) : boolean;

      procedure ClearQueue;
      procedure EnqueueMsg(Sujet, Corps : String);
      function SendQueue : boolean;
      function ReceiveQueue : boolean;

    published
      property MsgQueue : TMsgList read FMsgQueue;
      property Exclusif : boolean read FExclusif write FExclusif;
      property Encrypted : boolean read FEncrypted write FEncrypted;

      property OnLookupQueue : TNotifyEvent read FOnLookupQueue write FOnLookupQueue;
      property OnOpenQueue : TOpenQueueEvent read FOnOpenQueue write FOnOpenQueue;
      property OnCloseQueue : TNotifyEvent read FOnCloseQueue write FOnCloseQueue;
      property OnBeginTransaction : TNotifyEvent read FOnBeginTransaction write FOnBeginTransaction;
      property OnEndTransaction : TEndTransEvent read FOnEndTransaction write FOnEndTransaction;
      property OnSendMessage : TMsgEvent read FOnSendMessage write FOnSendMessage;
      property OnMessageSent : TMsgEvent read FOnMessageSent write FOnMessageSent;
      property OnGetNextMessage : TNotifyEvent read FOnGetNextMessage write FOnGetNextMessage;
      property OnMessageReceived : TMsgEvent read FOnMessageReceived write FOnMessageReceived;

    end;


procedure WideCompress(var S : WideString);
procedure WideExpand(var S : WideString);


implementation
uses ComObj, Dialogs;

// Divise la taille d'une WideString par 2 en plaçant la moitié des caractères dans
//  l'octet de poids fort de l'autre moitié (car les caractères sont sur 16 bits)
procedure WideCompress(var S : WideString);
var tmp : AnsiString;
    i : integer;
begin
if (length(S) mod 2) = 1 then S := S + #0;
tmp := S;
SetLength(S, Length(S) div 2);
for i := 1 to Length(S) do
 S[i] := WideChar((Ord(tmp[i*2]) shl 8)+Ord(tmp[(i*2)-1]));
end;

// Opération inverse pour récupérer la WideString d'origine
procedure WideExpand(var S : WideString);
var CStr : WideString;
    EStr : AnsiString;
    i : integer;
begin
CStr := S;
SetLength(EStr, Length(S) * 2);
for i := 1 to Length(CStr) do
 begin
 EStr[i*2] := Char(Hi(Ord(CStr[i])));
 EStr[(i*2)-1] := Char(Lo(Ord(CStr[i])));
 end;
S := EStr;
end;


function TMsgList.GetMsg(index : integer) : TMsg;
begin
result := Items[index];
end;

// Ajoute un message à la liste et retourne le message créé
function TMsgList.Ajout(Sujet, Corps : string) : TMsg;
var M : TMsg;
begin
M := TMsg.Create;
M.Subject := Sujet;
M.Body := Corps;
Add(M);
result := M;
end;

// Supprime le message d'index index. Les messages suivants sont décalés
procedure TMsgList.Suppr(index : integer);
begin
Msg[index].Free;
Delete(index);
end;

// Supprime tous les messages
procedure TMsgList.Clear;
var i : integer;
begin
for i := Count-1 downto 0 do Suppr(i);
inherited;
end;

destructor TMsgList.destroy;
begin
Clear;
inherited;
end;

// Recherche un message en fonction de son sujet et retourne le premier trouvé (ou nil)
function TMsgList.FindMsg(Subject : String) : TMsg;
var i : integer;
begin
i := FindMsgIndex(Subject);
if i = -1 then result := nil
          else result := Msg[i];
end;

// Recherche un message en fonction de son sujet et retourne l'index du premier trouvé (ou -1)
function TMsgList.FindMsgIndex(Subject : String) : integer;
var i : integer;
begin
result := -1;
for i := 0 to Count-1 do
 if Msg[i].Subject = Subject then
   begin
   result := i;
   break;
   end;
end;

// Ajoute une extention de la forme .000 au sujet, de manière à ce que le sujet
//  obtenu n'existe pas déjà dans la liste
function TMsgList.FindFreeSubject(S : String) : String;
var i : integer;
begin
i := 0;
repeat
  result := S+Format('.%.3d', [i]);
  inc(i);
until FindMsg(result) = nil;
end;



// Découpe des message dépassant la taille maxi en plusieurs messages. Le message
//  d'origine est remplacé par un message contenant le sujet du message d'origine
//  ainsi que les sujets des messages contenant les morceaux du message d'origine.
procedure TMsgList.SplitMessages;
var i : integer;
    s,b, Unique : string;
    SL : TStringList;
begin
for i := Count-1 downto 0 do
  if Length(Msg[i].Body) > MSMQMaxSize then
    begin
    SL := TStringList.Create;
    SL.Add(Msg[i].Subject);
    b := Msg[i].Body;

    Unique := '.' + IntToStr(Trunc(Now)) + FormatDateTime('hhnnsszzz', Now);

    repeat
      s := FindFreeSubject(Msg[i].Subject + Unique);

      SL.Add(s);
      Ajout(s, Copy(b, 1, MSMQMaxSize));
      System.Delete(b, 1, MSMQMaxSize);
    until b = '';

    Msg[i].Subject := '[Splitted]';
    Msg[i].Body := SL.Text;

    SL.Free;
    end;
end;

// Recolle les morceaux de messages découpés afin de retrouver les messages d'origine.
procedure TMsgList.UnSplitMessages;
var j,idx : integer;
    b : string;
    SL : TStringList;
    M : TMsg;
begin
if Count = 0 then exit;
repeat
  M := FindMsg('[Splitted]');
  if M <> nil then
    begin
    SL := TStringList.Create;
    SL.Text := M.Body;
    b := '';

    for j := 1 to SL.Count-1 do
      begin
      idx := FindMsgIndex(SL.Strings[j]);
      b := b + Msg[idx].Body;
      Suppr(idx);
      end;

    M.Subject := SL.Strings[0];
    M.Body := b;

    SL.Free;
    end;
until M = nil;
end;

// Si NomQueue ne contient pas d'antislash, le constructeur fait un LookUp sur le
//  nom passé en paramètre. Sinon il crée directement l'objet MSMQQueueInfo
constructor TMSMQueue.Create(NomQueue : String);
begin
inherited Create;
try
  if Pos('\', NomQueue) = 0 then FInfoQueue := LookupQueue(NomQueue)
    else begin
         FInfoQueue := CreateOleObject('MSMQ.MSMQQueueInfo') as IMSMQQueueInfo;
         FInfoQueue.PathName := NomQueue;
         end;
  FDisp := CreateOleObject('MSMQ.MSMQTransactionDispenser') as IMSMQTransactionDispenser;//CreateCOMObject(CLASS_MSMQTransactionDispenser) as IMSMQTransactionDispenser;
  FMsgOut := CreateOleObject('MSMQ.MSMQMessage') as IMSMQMessage;//CreateCOMObject(CLASS_MSMQMessage) as IMSMQMessage;
  FMsgQueue := TMsgList.Create;
  Exclusif := false;
  FExcept := nil;
except
  on E : Exception do FExcept := E;
end;
end;

constructor TMSMQueue.Create(InfoQueue : IMSMQQueueInfo);
begin
inherited Create;
try
  FInfoQueue := InfoQueue;
  FDisp := CreateOleObject('MSMQ.MSMQTransactionDispenser') as IMSMQTransactionDispenser;//CreateCOMObject(CLASS_MSMQTransactionDispenser) as IMSMQTransactionDispenser;
  FMsgOut := CreateOleObject('MSMQ.MSMQMessage') as IMSMQMessage;//CreateCOMObject(CLASS_MSMQMessage) as IMSMQMessage;
  FMsgQueue := TMsgList.Create;
  Exclusif := false;
  FExcept := nil;
except
  on E : Exception do FExcept := E;
end;
end;

destructor TMSMQueue.Destroy;
begin
FMsgQueue.Free;

FInfoQueue := nil;
FDisp := nil;
FMsgOut := nil;

inherited;
end;

// Crée la queue sur le serveur MSMQ (transactionnelle)
procedure TMSMQueue.CreateQueue;
var isTrans,isRead : OLEVariant;
begin
if FExcept <> nil then raise FExcept else
begin
  isTrans := true; isRead := true;
  FInfoQueue.Create(isTrans, isRead);
end;
end;

// Détruit la queue sur le serveur
procedure TMSMQueue.DeleteQueue;
begin
  if FExcept <> nil then raise FExcept else
  FInfoQueue.Delete;
end;

// Recherche une queue en fonction de son label
function TMSMQueue.LookupQueue(Nom : String) : IMSMQQueueInfo;
var Q : IMSMQQuery;
    QIs : IMSMQQueueInfos;
    n : OleVariant;
begin
  if Assigned(FOnLookupQueue) then FOnLookupQueue(self);

  Q := CreateOleObject('MSMQ.MSMQQuery') as IMSMQQuery;
  n := Nom;
  QIs := Q.LookupQueue(EmptyParam, EmptyParam, n, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  if QIs = nil then raise Exception.Create('Impossible d''effectuer la recherche de file');
  QIs.ReSet;
  result := QIs.Next;
  if result = nil then raise Exception.Create('Impossible de trouver la file "'+Nom+'"');
end;

// Ouvre la queue
function TMSMQueue.OuvrirQueue(Lecture, Exclusif : boolean) : IMSMQQueue;
var Access, ShareMode : integer;
begin
if Assigned(FOnOpenQueue) then FOnOpenQueue(self, Lecture, Exclusif);

if Exclusif then ShareMode := MQ_DENY_RECEIVE_SHARE
            else ShareMode := MQ_DENY_NONE;
if Lecture then Access := MQ_RECEIVE_ACCESS
           else Access := MQ_SEND_ACCESS;
result := FInfoQueue.Open(Access, ShareMode);
if result.IsOpen = 0 then result := nil;
end;

// Ferme la queue
procedure TMSMQueue.FermerQueue(LaQueue : IMSMQQueue);
begin
if Assigned(FOnCloseQueue) then FOnCloseQueue(self);

LaQueue.Close;
end;

// Récupere le message suivant dans la queue
function TMSMQueue.LireMessage(LaQueue : IMSMQQueue; Msg : TMsg; var Trans : IMSMQTransaction; OnlyPeek : boolean = false) : boolean;
var tr, wdq, wb, rt : OLEVariant;
    S : WideString;
    FMsgIn : IMSMQMessage;
begin
result := false;
if Assigned(FOnGetNextMessage) then FOnGetNextMessage(self);

if Trans = nil then tr := MQ_MTS_TRANSACTION
               else tr := Trans;

wdq := False; wb := True; rt := 0;

if OnlyPeek then
  begin
  if FFirst then FMsgIn := LaQueue.PeekCurrent(wdq, wb, rt)
            else FMsgIn := LaQueue.PeekNext(wdq, wb, rt);
  FFirst := false;
  end
  else FMsgIn := LaQueue.Receive(tr, wdq, wb, rt);

if FMsgIn <> nil then
  begin
  Msg.Subject := FMsgIn.Label_;

  S := WideString(FMsgIn.Body);
  WideExpand(S);
  Msg.Body := S;

  FMsgIn := nil;

  if Assigned(FOnMessageReceived) then FOnMessageReceived(self, Msg);
  result := true;
  end;
end;

// Envoi un message dans la queue
procedure TMSMQueue.EcrireMessage(LaQueue : IMSMQQueue; XAct : IMSMQTransaction; Msg : TMsg);
var tr : OLEVariant;
    S : WideString;
begin
if Assigned(FOnSendMessage) then FOnSendMessage(self, Msg);
tr := XAct;
FMsgOut.Label_ := Msg.Subject;
S := Msg.Body; WideCompress(S);
FMsgOut.Body := S;
FMsgOut.Delivery := MQMSG_DELIVERY_RECOVERABLE;
FMsgOut.MaxTimeToReachQueue := -1;
FMsgOut.MaxTimeToReceive := -1;
if FEncrypted then
begin
  FMsgOut.PrivLevel := MQMSG_PRIV_LEVEL_BODY;
  FMsgOut.EncryptAlgorithm := MQMSG_CALG_RC4;
end
 else FMsgOut.PrivLevel := MQMSG_PRIV_LEVEL_NONE;
FMsgOut.Send(LaQueue, tr);
if Assigned(FOnMessageSent) then FOnMessageSent(self, Msg);
end;

// BeginTrans
function TMSMQueue.PreparerTransaction : IMSMQTransaction;
begin
if Assigned(FOnBeginTransaction) then FOnBeginTransaction(self);

result := FDisp.BeginTransaction;
end;

// Abort / Commit
procedure TMSMQueue.FinirTransaction(XAct : IMSMQTransaction; Abort : boolean = false);
var fret, grfTC, grFrm, fAsn : OLEVariant;
begin
if Assigned(FOnEndTransaction) then FOnEndTransaction(self, Abort);

fret := false; grfTC := False; grFrm := false; fAsn := true;
if Abort then XAct.Abort(fRet, fAsn)
         else XAct.Commit(fRet, grfTC, grFrm);
end;

// Envoie tous les messages d'une TMsgList
function TMSMQueue.Send(MList : TMsgList) : boolean;
var ZeQueue : IMSMQQueue;
    XAct : IMSMQTransaction;
    i : integer;
begin
  if FExcept <> nil then raise FExcept;
result := false;

MList.SplitMessages;

try
  ZeQueue := OuvrirQueue(false, FExclusif);
  if ZeQueue = nil then exit;

  XAct := PreparerTransaction;
  for i := 0 to MList.Count-1 do
    EcrireMessage(ZeQueue, XAct, MList.Msg[i]);
  FinirTransaction(XAct);

  FermerQueue(ZeQueue);
  result := true;
  ErrorMsg := '';
except
  on E : Exception do ErrorMsg := E.Message;
end;
end;

// Reçoit tous les messages de la queue dans une TMsgList
function TMSMQueue.Receive(MList : TMsgList; Transactionnel : boolean = false) : boolean;
var ZeQueue : IMSMQQueue;
    XAct : IMSMQTransaction;
begin
  if FExcept <> nil then raise FExcept;
result := false;

try
  ZeQueue := OuvrirQueue(true, FExclusif);
  if ZeQueue = nil then exit;

  if Transactionnel then XAct := PreparerTransaction
                    else XAct := nil;

  FFirst := true;
  MList.Clear;
  repeat until LireMessage(ZeQueue, MList.Ajout('', ''), XAct, MSMQDebug) = false;
  MList.Suppr(MList.FindMsgIndex(''));
  //until FMsgIn = nil;

  if Transactionnel then FinirTransaction(XAct);

  FermerQueue(ZeQueue);

  MList.UnSplitMessages;

  result := true;
  ErrorMsg := '';
except
  on E : Exception do ErrorMsg := E.Message;
end;
end;

// Efface la TMsgList interne
procedure TMSMQueue.ClearQueue;
begin
FMsgQueue.Clear;
end;

// Ajoute un message à la TMsgList interne
procedure TMSMQueue.EnqueueMsg(Sujet, Corps : String);
begin
FMsgQueue.Ajout(Sujet, Corps);
end;

// Envoie les messages de la TMsgList interne
function TMSMQueue.SendQueue : boolean;
begin
result := Send(FMsgQueue);
end;

// Reçoit les messages dans la TMsgList interne
function TMSMQueue.ReceiveQueue : boolean;
begin
result := Receive(FMsgQueue, false);
end;


end.
