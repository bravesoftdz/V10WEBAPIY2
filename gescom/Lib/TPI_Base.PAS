Unit TPI_Base ;

Interface

Uses Classes,MC_Comm,MC_Timer,sysutils,HCtrls,HEnt1,forms;

Type TStartComTPIEvent   = procedure( var Cancel : Boolean )                 of object;
Type TReceiveComTPIEvent = procedure( Data : String ; var Cancel : Boolean ) of object;
Type TStopComTPIEvent    = procedure( Code : Integer ; Libelle : String )    of object;

///////////////////////////////////////////////////////////////////////////////////////
//  TTPI_Base :
///////////////////////////////////////////////////////////////////////////////////////
Type
  TTPI_Base = Class ( TComponent)
     Private
      FTPI            : String ;
      FRecu           : String ;
      FArret          : Boolean ;
      FStartCom       : TStartComTPIEvent ;   // Evenement sur le démarrage de la communication avec un TPI
      FReceiveCom     : TReceiveComTPIEvent ; // Evenement sur la réception de trame du TPI
      FStopCom        : TStopComTPIEvent ;    // Evenement sur le démarrage de la communication avec un TPI
      procedure LirePort (Sender: TObject; DataPtr: pointer; DataSize: integer ) ;
     Protected
      FComm         : TMC_CommportDriver ;
      Procedure SetPort      ( AValue : TComPortNumber) ;   Virtual ;
      Function  GetPort      : TComPortNumber ;             Virtual ;
      Procedure SetBaudRate  ( AValue : TComPortBaudRate) ; Virtual ;
      Function  GetBaudRate  : TComPortBaudRate ;           Virtual ;
      Procedure SetDataBits  ( AValue : TComPortDataBits) ; Virtual ;
      Function  GetDataBits  : TComPortDataBits ;           Virtual ;
      Procedure SetStopBits  ( AValue : TComPortStopBits) ; Virtual ;
      Function  GetStopBits  : TComPortStopBits;            Virtual ;
      Procedure SetParity    ( AValue : TComPortParity) ;   Virtual ;
      Function  GetParity    : TComPortParity ;             Virtual ;
      Procedure SetHWProtocol ( AValue : TComPortHwHandshaking) ; Virtual ;
      Function  GetHWProtocol : TComPortHwHandshaking ;     Virtual ;
      Procedure SetSWProtocol ( AValue : TComPortSwHandshaking) ; Virtual ;
      Function  GetSWProtocol : TComPortSwHandshaking ;           Virtual ;
      Procedure SetArret     ( AValue : Boolean) ;          Virtual ;
      Function  GetArret     : Boolean ;                    Virtual ;

      Function  Sendmessage ( Mess : String ) : Boolean ;
      Function  ReceiveMess : String ;
      Procedure VideBuffer ;
      Procedure VideNbCaractBuffer (NbCar : Integer) ;
     public
      Constructor Create ( AOwner : TComponent ) ; overload ; override ;
      Constructor Create ( AOwner : TComponent ; ATPI : String ) ; reintroduce ; overload ; virtual ;
      Destructor  Destroy  ; Override ;

      Function ChargePortEtParams(Aport, AParams : String ) : Integer ; virtual ;
      Function Transaction( RemiseAttente : Boolean )       : Integer ; virtual ; Abstract ;

      Property TPI        : String                read FTPI ;
      Property Port       : TComPortNumber        Read GetPort       write SetPort ;
      property Speed      : TComPortBaudRate      read GetBaudRate   write SetBaudRate ;
      property DataBits   : TComPortDataBits      read GetDataBits   write SetDataBits ;
      property StopBits   : TComPortStopBits      read GetStopBits   write SetStopBits ;
      property Parity     : TComPortParity        read GetParity     write SetParity ;
      Property HWProtocol : TComPortHwHandshaking read GetHWProtocol write SetHWProtocol ;
      Property SWProtocol : TComPortSwHandshaking read GetSWProtocol write SetSWProtocol ;
      Property Arret      : Boolean               read GetArret      write SetArret ;
      // Evenements de la communication
      Property OnStartCom     : TStartComTPIEvent   read FStartCom    write FStartCom ;
      Property OnReceiveCom   : TReceiveComTPIEvent read FReceiveCom  write FReceiveCom ;
      Property OnStopCom      : TStopComTPIEvent    read FStopCom     write FStopCom ;
     End ;

///////////////////////////////////////////////////////////////////////////////////////
//  TTPI_Timer :
///////////////////////////////////////////////////////////////////////////////////////
Type
  TTPI_Timer = Class (TComponent)
     private
      FTimer   : TMC_Timer ;
      FExceded : Boolean ;
      Procedure IsTimer (Sender : TObject ) ;
     protected
      Procedure SetMax (AValue : Longword ) ;
      function  GetMax : Longword  ;
     public
      Constructor Create (AOwner : TComponent) ; Override ;
      Destructor  Destroy ; Override ;
      Procedure   Initialise ;

      Property MaxInterval : Longword read GetMax write SetMax ;
      Property Exceded     : Boolean  read FExceded ;
     End ;

///////////////////////////////////////////////////////////////////////////////////////
//  Paramètres du protocole de communication
///////////////////////////////////////////////////////////////////////////////////////
Const Max_NbCon = 3 ;      // Nombre maximun de tentative de connexion
      Max_NbRec = 3 ;      // Nombre maximum de lecture d'une trame

Type
  TEtatRecepteur = (R0, R1, R2, R3, R4) ;

///////////////////////////////////////////////////////////////////////////////////////
//  TTPISYMBOL :
///////////////////////////////////////////////////////////////////////////////////////
Type
  TTPISYMBOL = Class (TTPI_BASE)
     private
      ERecep    : TEtatRecepteur ;  // Etat de la communication
      NbCon     : Integer ;         // Nombre de tentatives de connexion
      NbRec     : Integer ;         // Nombre de tentative de réception d'une trame
      TMOMsg    : TTPI_Timer ;      // Timeout entre chaque trame
      TMOCar    : TTPI_Timer ;      // Timeout entre chaque caractère d'une trame
      Function    GereReception  : Integer ;
     public
      Constructor Create ( AOwner : TComponent ; ATPI : String ) ; Override ;
      Destructor  Destroy ; Override ;
      Function    ChargePortetParams(APort, AParams : String ) : Integer ; override ;
      Function    Transaction( RemiseAttente : Boolean )       : Integer ; override ;
     End ;

///////////////////////////////////////////////////////////////////////////////////////
//  TTPITELXON :
///////////////////////////////////////////////////////////////////////////////////////
Type
  TTPITELXON = Class (TTPI_BASE)
     private
      ERecep    : TEtatRecepteur ;  // Etat de la communication
      NbCon     : Integer ;         // Nombre de tentatives de connexion
      NbRec     : Integer ;         // Nombre de tentative de réception d'une trame
      TMOMsg    : TTPI_Timer ;      // Timeout entre chaque trame
      TMOCar    : TTPI_Timer ;      // Timeout entre chaque caractère d'une trame
      Function    GereReception  : Integer ;
     public
      Constructor Create ( AOwner : TComponent ; ATPI : String ) ; Override ;
      Destructor  Destroy ; Override ;
      Function    ChargePortetParams(APort, AParams : String ) : Integer ; override ;
      Function    Transaction( RemiseAttente : Boolean )       : Integer ; override ;
     End ;

///////////////////////////////////////////////////////////////////////////////////////
//  TTPI :
///////////////////////////////////////////////////////////////////////////////////////
Type
  TTPI = Class (TComponent)
     Private
      FTPI          : TTPI_Base ;
     Protected
      Procedure SetPort           ( AValue : TComPortNumber) ;
      Function  GetPort           : TComPortNumber ;
      Procedure SetBaudRate       ( AValue : TComPortBaudRate) ;
      Function  GetBaudRate       : TComPortBaudRate ;
      Procedure SetDataBits       ( AValue : TComPortDataBits) ;
      Function  GetDataBits       : TComPortDataBits ;
      Procedure SetStopBits       ( AValue : TComPortStopBits) ;
      Function  GetStopBits       : TComPortStopBits;
      Procedure SetParity         ( AValue : TComPortParity) ;
      Function  GetParity         : TComPortParity ;
      Procedure SetArret          ( AValue : Boolean) ;
      Function  GetArret          : Boolean ;
      Procedure SetStartCom       ( AValue : TStartComTPIEvent ) ;
      Function  GetStartCom       : TStartComTPIEvent ;
      Procedure SetStopCom        ( AValue : TStopComTPIEvent ) ;
      Function  GetStopCom        : TStopComTPIEvent ;
      Procedure SetReceiveCom     ( AValue : TReceiveComTPIEvent ) ;
      Function  GetReceiveCom     : TReceiveComTPIEvent ;
     Public
      Constructor Create (AOwner : TComponent) ; overload ; Override ;
      Constructor Create (AOwner : TComponent ; ATPI : String ) ; reintroduce ; Overload ; Virtual ;
      Destructor  Destroy ; Override ;

      Function ChargePortetParams(APort, AParams : String ) : Integer ;
      Function Transaction( RemiseAttente : Boolean )       : Integer ;

      Property Port         : TComPortNumber     Read GetPort     write SetPort     ;
      property Speed        : TComPortBaudRate   read GetBaudRate write SetBaudRate ;
      property DataBits     : TComPortDataBits   read GetDataBits write SetDataBits ;
      property StopBits     : TComPortStopBits   read GetStopBits write SetStopBits ;
      property Parity       : TComPortParity     read GetParity   write SetParity   ;
      Property Arret        : Boolean            read GetArret    write SetArret ;

      // Evenements de la communication
      Property OnStartCom     : TStartComTPIEvent   read GetStartCom     write SetStartCom ;
      Property OnStopCom      : TStopComTPIEvent    read GetStopCom      write SetStopCom ;
      Property OnReceiveCom   : TReceiveComTPIEvent read GetReceiveCom   write SetReceiveCom ;
     End ;

Function TPI_MsgErrDefaut ( Code : Integer ) : String ;

Implementation

Const STX = $02 ;
      ETX = $03 ;
      EOT = $04 ;
      ENQ = $05 ;
      ACK = $06 ;
      HT  = $09 ;
      VT  = $0B ;
      NAK = $15 ;
      Chr_STX : String = chr(STX) ;
      Chr_ETX : String = chr(ETX) ;
      Chr_EOT : String = chr(EOT) ;
      Chr_ENQ : String = chr(ENQ) ;
      Chr_ACK : String = chr(ACK) ;
      Chr_NAK : String = chr(NAK) ;
      Chr_BTQ = chr(ACK) ;
      Chr_QTE = chr(HT)  ;
      Msg_ACK = '%0' ;
      Msg_NAK = '%1' ;


{==============================================================================================}
{=================================== MESSAGES D'ERREUR ========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  TPI_MsgErrDefaut : donne le libellé correspondant à un code erreur
///////////////////////////////////////////////////////////////////////////////////////
Function TPI_MsgErrDefaut ( Code : Integer ) : String ;
BEGIN
Result := '' ;
case Code of
  1001 : Result:='Le code du TPI (%s) est inconnu.' ;
  1002 : Result:='Le TPI n''est pas défini.' ;
  1003 : Result:='Il faut utiliser le create en indiquant le code du TPI.' ;
  1004 : Result:='Aucun port de communication n''est paramétré.' ;
  1005 : Result:='Erreur d''émission.' ;
  1006 : Result:='La connexion a échouée.' ;
  1007 : Result:='La communication a été interrompue par le programme' ;
  1008 : Result:='La communication a été interrompue par l''utilisateur' ;
  1009 : Result:='Erreur de réception.' ;
  1010 : Result:='Erreur sur timeout.' ;

  else Result:='Erreur n°'+IntToStr(Code) ;
  END ;
END ;

{==============================================================================================}
{==================================== OBJET TTPI_BASE =========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  Create :
///////////////////////////////////////////////////////////////////////////////////////
Constructor TTPI_Base.Create ( AOwner : TComponent ) ;
BEGIN
Raise EComponentError.Create(TPI_MsgErrDefaut(1003)) ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Create :
///////////////////////////////////////////////////////////////////////////////////////
Constructor TTPI_Base.Create ( AOwner : TComponent ; ATPI : String ) ;
BEGIN
Inherited Create(AOwner) ;
FTPI := ATPI ;
FComm := TMC_CommportDriver.Create(Self) ;
FComm.OnReceiveData := LirePort ;
Port := pnNone ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Destroy :
///////////////////////////////////////////////////////////////////////////////////////
Destructor TTPI_Base.Destroy ;
BEGIN
FComm.Free ;
Inherited Destroy ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  ChargePortEtParams : définition du port et des paramètres de communication
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.ChargePortEtParams(Aport, AParams : String ) : Integer ;
Var Stg      : String ;
    Ind, Val : Integer ;
BEGIN
Result := 0 ;
// Port de communication
APort := Uppercase(Trim(APort)) ;
if (APort = '') or (APort = 'AUC') then
   BEGIN
   Result := 1004 ;
   Exit ;
   END ;
Port := TComPortNumber(Pos(';'+APort+';',';AUC;FTX;CO1;CO2;CO3;CO4;LP1;LP2;') div 4) ;
if Port in [pnCOM1,pnCOM2,pnCOM3,pnCOM4] then
   BEGIN
   ind := 0 ;
   while Aparams <> '' do
     BEGIN
     Inc(Ind) ;
     Stg := ReadTokenPipe(AParams, ';') ;
     if Stg <> '' then
        BEGIN
        Case Ind of
          1 : Speed    := TComPortBaudRate(Valeuri(Stg)-1) ;
          2 : Parity   := TComPortParity(Valeuri(Stg)-1) ;
          3 : DataBits := TComPortDataBits(Valeuri(Stg)-1) ;
          4 : StopBits := TComPortStopBits(Valeuri(Stg)-1) ;
          5 : BEGIN
              Val := Valeuri(Stg)-1 ;
              if Val < 2 then HWProtocol := TComPortHwHandshaking(Val)
                         else SWProtocol := TComPortSwHandshaking(Val-2) ;
              END ;
          END ;
        END ;
     END ;
   END ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  LirePort :
///////////////////////////////////////////////////////////////////////////////////////
procedure TTPI_Base.LirePort (Sender: TObject; DataPtr: pointer; DataSize: integer ) ;
Var Stg : String ;
    Ptr : PChar ;
BEGIN
Stg := '' ;
Ptr := DataPtr ;
While DataSize > 0 do
  BEGIN
  Stg := Stg + Ptr^ ;
  Inc(Ptr) ;
  Dec(DataSize) ;
  END ;
FRecu := FRecu + Stg ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetPort :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetPort ( AValue : TComPortNumber) ;
BEGIN
FComm.ComPort := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetPort :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetPort : TComPortNumber ;
BEGIN
Result := FComm.ComPort ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetBaudRate :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetBaudRate ( AValue : TComPortBaudRate) ;
BEGIN
FComm.Speed := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetBaudRate :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetBaudRate : TComPortBaudRate ;
BEGIN
Result := FComm.Speed;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetDataBits :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetDataBits ( AValue : TComPortDataBits) ;
BEGIN
FComm.DataBits := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetDataBits :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetDataBits : TComPortDataBits ;
BEGIN
Result := FComm.DataBits ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetStopBits :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetStopBits ( AValue : TComPortStopBits) ;
BEGIN
FComm.StopBits := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetStopBits :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetStopBits : TComPortStopBits;
BEGIN
Result := FComm.StopBits ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetParity :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetParity ( AValue : TComPortParity) ;
BEGIN
FComm.Parity := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetParity :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetParity : TComPortParity ;
BEGIN
Result := FComm.Parity ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetHWProtocol :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetHWProtocol ( AValue : TComPortHwHandshaking) ;
BEGIN
FComm.HwHandshaking := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetHWProtocol :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetHWProtocol : TComPortHwHandshaking ;
BEGIN
Result := FComm.HwHandshaking ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetSWProtocol :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetSWProtocol ( AValue : TComPortSWHandshaking) ;
BEGIN
FComm.SWHandshaking := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetSWProtocol :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetSWProtocol : TComPortSWHandshaking ;
BEGIN
Result := FComm.SWHandshaking ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetArret :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.SetArret ( AValue : Boolean) ;
BEGIN
FArret := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetArret :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.GetArret : Boolean ;
BEGIN
Result := FArret ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SendMessage :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.SendMessage (Mess : String ) : Boolean ;
BEGIN
Result := FComm.SendString(Mess) ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  ReceiveMess :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI_Base.ReceiveMess : String ;
Var Ind : Integer ;
BEGIN
Result := FRecu ;
Ind := Length(Result) ;
while Ind > 0 do
  BEGIN
  if not(ord(Result[Ind]) in [STX,ETX,EOT,ENQ,ACK,HT,VT,NAK,10,13,32..127]) then Delete(Result, Ind, 1) ;
  Dec(Ind) ;
  END ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  VideBuffer :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.VideBuffer ;
BEGIN
FRecu := '' ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  VideNbCaractBuffer :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Base.VideNbCaractBuffer (NbCar : Integer);
BEGIN
if Length(FRecu)>NbCar
   then FRecu := Copy (FRecu,NbCar+1,(Length(FRecu))-NbCar)
   else FRecu := '' ;
END ;
{==============================================================================================}
{====================================== OBJET TTPI ============================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  Create :
///////////////////////////////////////////////////////////////////////////////////////
Constructor TTPI.Create (AOwner : TComponent) ;
BEGIN
Raise EComponentError.Create(TPI_MsgErrDefaut(1003)) ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Create :
///////////////////////////////////////////////////////////////////////////////////////
Constructor TTPI.Create (AOwner : TComponent ;  ATPI : String ) ;
BEGIN
Inherited Create (Aowner) ;
FTPI := Nil ;
// Dispatching en fonction du type de TPI
if Pos(';'+ATPI+';', ';001;002;') > 0 then FTPI := TTPISYMBOL.Create(Self, ATPI) else
 if Pos(';'+ATPI+';', ';003;') > 0 then FTPI := TTPITELXON.Create(Self, ATPI) ;
if not Assigned(FTPI) then Raise EComponentError.CreateFmt(TPI_MsgErrDefaut(1001), [ATPI]) ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Destroy :
///////////////////////////////////////////////////////////////////////////////////////
Destructor  TTPI.Destroy ;
BEGIN
if Assigned(FTPI) then FTPI.Free ;
Inherited Destroy ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetPort :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetPort ( AValue : TComPortNumber) ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.Port := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetPort :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetPort : TComPortNumber ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.Port ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetBaudRate :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetBaudRate ( AValue : TComPortBaudRate) ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.Speed := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetBaudRate :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetBaudRate : TComPortBaudRate ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.Speed;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetDataBits :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetDataBits ( AValue : TComPortDataBits) ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.DataBits := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetDataBits :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetDataBits : TComPortDataBits ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.DataBits ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetStopBits :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetStopBits ( AValue : TComPortStopBits) ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.StopBits := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetStopBits :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetStopBits : TComPortStopBits;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.StopBits ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetParity :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetParity ( AValue : TComPortParity) ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.Parity := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetParity :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetParity : TComPortParity ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.Parity ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetArret :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetArret ( AValue : Boolean) ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.FArret := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetArret :
///////////////////////////////////////////////////////////////////////////////////////
Function  TTPI.GetArret : Boolean ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.FArret ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetStartCom :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetStartCom ( AValue : TStartComTPIEvent) ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.FStartCom := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetStartCom :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetStartCom : TStartComTPIEvent ;
BEGIN
if not assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.FStartCom ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetStopCom :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetStopCom ( AValue : TStopComTPIEvent) ;
BEGIN
if not assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.FStopCom := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetStopCom :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetStopCom : TStopComTPIEvent ;
BEGIN
if not assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.FStopCom ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetReceiveCom :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI.SetReceiveCom ( AValue : TReceiveComTPIEvent) ;
BEGIN
if not assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
FTPI.FReceiveCom := AValue ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetReceiveCom :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.GetReceiveCom : TReceiveComTPIEvent ;
BEGIN
if not assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.FReceiveCom ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  ChargePortEtParams :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.ChargePortEtParams(APort, AParams : String ) : Integer ;
BEGIN
if Not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.ChargePortetParams(APort, AParams) ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Transaction :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPI.Transaction( RemiseAttente : Boolean ) : Integer ;
BEGIN
if not Assigned(FTPI) then Raise EComponentError.Create(TPI_MsgErrDefaut(1002)) ;
Result := FTPI.Transaction(RemiseAttente) ;
END ;

{==============================================================================================}
{=================================== OBJET TTPI_Timer =========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  Create :
///////////////////////////////////////////////////////////////////////////////////////
Constructor TTPI_Timer.Create (AOwner : TComponent) ;
BEGIN
Inherited Create (AOwner) ;
FTimer := TMC_Timer.Create(Self) ;
//FTimer.TimerPriority := tpHighest ;
FTimer.OnTimer := IsTimer ;
FTimer.Enabled := FALSE ;
MaxInterval := 0 ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Destroy :
///////////////////////////////////////////////////////////////////////////////////////
Destructor TTPI_Timer.Destroy ;
BEGIN
FTimer.Free ;
Inherited Destroy ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  IsTimer :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Timer.IsTimer (Sender : TObject ) ;
BEGIN
FExceded := TRUE ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  SetMax :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Timer.SetMax (AValue : Longword ) ;
Var Roule : Boolean ;
BEGIN
Roule := FTimer.Enabled ;
if AValue <> FTimer.Delay then FTimer.Delay := AValue ;
if (Roule) and (FTimer.delay > 0) then Initialise ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  GetMax :
///////////////////////////////////////////////////////////////////////////////////////
Function  TTPI_Timer.GetMax : Longword  ;
BEGIN
Result := FTimer.delay ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Initialise :
///////////////////////////////////////////////////////////////////////////////////////
Procedure TTPI_Timer.Initialise ;
BEGIN
FExceded := FALSE ;
if FTimer.Enabled then FTimer.Enabled := FALSE ;
FTimer.Enabled := (FTimer.delay > 0) ;
END ;

{==============================================================================================}
{==================================== OBJET TTPITELXON =========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  Create :
///////////////////////////////////////////////////////////////////////////////////////
Constructor TTPITELXON.Create ( AOwner : TComponent ; ATPI : String ) ;
BEGIN
Inherited Create(Aowner,ATPI) ;
// valeur par defaut
FComm.Speed := br1200 ;
FComm.DataBits := db7BITS ;
FComm.StopBits := sb1BITS ;
FComm.Parity := ptEVEN ;
// Valeurs des nombres de tentatives et des Timeout
ERecep := R0 ;
NbCon := 0 ;
NbRec := 0 ;
TMOMsg := TTPI_Timer.Create(Self) ; TMOMsg.MaxInterval := 1000 ; TMOMsg.Tag := 1 ;      // 1 seconde
TMOCar := TTPI_Timer.Create(Self) ; TMOCar.Tag := 3 ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Destroy :
///////////////////////////////////////////////////////////////////////////////////////
Destructor TTPITELXON.Destroy ;
BEGIN
TMOCar.Free ;
TMOMsg.Free ;
Inherited Destroy ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  ChargePortEtParams :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPITELXON.ChargePortEtParams(APort, AParams : String ) : Integer ;
BEGIN
Result := Inherited ChargePortEtParams(APort, AParams) ;
End ;
///////////////////////////////////////////////////////////////////////////////////////
//  GereReception :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPITELXON.GereReception : Integer ;
Var Trame, Rslt, Reponse  : String ;
    TimeExceded, Cancel   : Boolean ;
    pETX                  : Integer ;
    pBTQ, pQTE            : Integer ;
    NbCarTraites          : Integer ;
BEGIN
NbCon := 0 ;
repeat
   Result := 0 ;
   Trame := '' ;
   Cancel := False ;
   ERecep := R1 ;
   NbRec := 0 ;
   Inc(NbCon) ;
   // R1 = Attente de connexion jusqu'à réception d'un ENQ
   repeat
     Rslt := ReceiveMess ;
     if Length(Rslt) > 1 then Rslt := Rslt[1] ;
     Application.ProcessMessages ;
   until (Rslt = Chr_ENQ) or (GetArret) ;
   // Traitement de la cause de sortie de l'attente
   if Rslt = Chr_ENQ then
      BEGIN
      // Connexion établie
      if Assigned(OnStartCom) then OnStartCom(Cancel) ;
      if Cancel then
         BEGIN
         Result := 1007 ;
         Exit ;
         END ;
      ERecep := R2 ;
      // VideBuffer ;
      VideNbCaractBuffer(1) ;  // Supprime le premier caractère du buffer uniquement
      // Accusé de réception
      if not SendMessage(Chr_ACK) then Result := 1005 else
         // R2 = Réception des trames
         repeat
            ERecep := R2 ;
            if Result = 0 then
               BEGIN
               TMOMsg.Initialise ;             // T1 = timeout de 1 seconde
               NbCarTraites:=0 ;
               repeat
                 Rslt := '' ;
                 Trame := ReceiveMess ;
                 // En cas de réception de caractères, on considère que le 1er caractère est un STX
                 if Trame <> '' then Rslt := Chr_STX ;
                 // Recherche du caractère de déconnexion
                 if (Pos(Chr_EOT, Trame) > 0) then Rslt := Chr_EOT ;
                 Application.ProcessMessages ;
                 TimeExceded := TMOMsg.Exceded ;
               until (TimeExceded) or (Rslt = Chr_STX) or (Rslt = Chr_EOT) ;
               // Traitement des caractères reçus
               if Rslt = Chr_STX then
                  BEGIN
                  // Traitement de la trame reçue - recherche du caractère de fin de trame
                  pETX := Pos(Chr_ETX, Trame) ;
                  if pETX > 0 then
                     BEGIN
                     // suppression des éventuels caractères derrière la fin de trame
                     Delete(Trame, pETX+1, Length(Trame)-pETX) ;
                     END else
                     BEGIN
                     // R3 = on lit la suite de la trame
                     ERecep := R3 ;
                     TMOCar.Initialise ;         // TD = timeout de 1, 3 ou 6 secondes selon la vitesse
                     TimeExceded := FALSE ;
                     while (pETX <> Length(Trame)) and (Not TimeExceded) do
                       BEGIN
                       Trame := ReceiveMess ;
                       while (Pos(Chr_EOT, Trame) > 0) do Delete(Trame, pos(Chr_EOT, Trame), 1) ;
                       while (Pos(Chr_ENQ, Trame) > 0) do Delete(Trame, pos(Chr_ENQ, Trame), 1) ;
                       while (Pos(Chr_STX, Trame) > 0) do Delete(Trame, pos(Chr_STX, Trame), 1) ;
                       Application.ProcessMessages ;
                       pETX := Pos(Chr_ETX, Trame) ;
                       if pETX > 0 then Delete(Trame, pETX+1, Length(Trame)-pETX) ;
                       TimeExceded := TMOCar.Exceded ;
                       END ;
                     END ;
                  if (not TimeExceded) and ((Rslt=Chr_EOT) or ((Rslt=Chr_STX) and (Pos(Chr_ETX, Trame)=Length(Trame)))) then
                     BEGIN
                     if Rslt=Chr_STX then
                        BEGIN
                        // suppression du fin de trame
                        pETX := Pos(Chr_ETX, Trame) ;
                        if pETX > 0 then Delete(Trame, pETX, Length(Trame)-pETX+1) ;
                        NbCarTraites:=pETX ;
                        END ;
                     // Acquittement de la trame reçu
                     Reponse := Chr_ACK ;
                     NbRec := 0 ;
                     if Assigned(OnReceiveCom) then
                        BEGIN
                        // Remplacement du caractère d'identification du code Boutique
                        pBTQ := Pos('B', Trame) ;
                        if pBTQ > 0 then Trame[pBtq]:=Chr_BTQ ;
                        // Remplacement du caractère d'identification d'une Quantité
                        pQTE := Pos('Q', Trame) ;
                        if pQTE > 0 then Trame[pQTE]:=Chr_QTE ;
                        OnReceiveCom(Trame, Cancel) ;
                        END ;
                     if Cancel then
                        BEGIN
                        Result := 1007 ;
                        Reponse := Chr_EOT ;
                        ERecep := R4 ;
                        END ;
                     END else
                     BEGIN
                     // La trame reçue est incorrecte
                     Reponse := Chr_NAK ;
                     Inc(NbRec) ;
                     END ;
                  END else
                  BEGIN
                  Inc(NbRec) ;
                  if Rslt = Chr_EOT then
                     BEGIN
                     ERecep := R4 ;
                     Reponse := Chr_ACK ;
                     END else Reponse := Chr_NAK ;
                  END ;
               if NbCarTraites>0
                  then VideNbCaractBuffer(NbCarTraites)  // Supprime les premiers caractères du buffer uniquement
                  else VideBuffer ;
               if not SendMessage(Reponse) then Result := 1005 ;
               END ;
         until (ERecep = R4) or (NbRec >= Max_NbRec) or (Result <> 0) or (GetArret) ; // Fin de R2
         if (ERecep <> R4) and (NbRec >= Max_NbRec) and (Result = 0) then
            BEGIN
            if GetArret then Result := 1008 else if Length(Trame) = 0 then Result := 1010 else Result := 1009 ;
            END ;
      END else
      BEGIN
      if GetArret then Result := 1008 else Result := 1006 ;
      END ;
until (ERecep = R4) or (NbCon >= Max_NbCon) or (Result <> 0) or (GetArret) ; // Fin de R1
VideBuffer ;
if (ERecep in [R2,R3,R4]) and (Assigned(OnStopCom)) then OnStopCom(Result, TPI_MsgErrDefaut(Result)) ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Transaction :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPITELXON.Transaction( RemiseAttente : Boolean ) : Integer ;
BEGIN
if FComm.ComPort <> pnNone then
   BEGIN
   Repeat
     FComm.Connect ;
     // Définition des caractères de contrôle
     Chr_STX := chr(STX) ;
     Chr_ETX := 'S' ;
     Chr_EOT := 'F' ;
     Chr_ENQ := 'A' ;
     Chr_ACK := '%0' ;
     Chr_NAK := '%1' ;
     // Définition du timeout entre chaque caractère reçu en fonction de la vitesse de transmission
     TMOCar.MaxInterval := 3000 ; // 3 secondes
     Case Speed of
       br1200 : TMOCar.MaxInterval := 6000 ; // 6 secondes
       br9600 : TMOCar.MaxInterval := 1000 ; // 1 seconde ;
       END ;
     NbCon := 0 ;
     ERecep := R0 ;
     Result := GereReception ;
     ERecep := R0 ;
     FComm.Disconnect ;
   Until (not RemiseAttente) or (GetArret) ;
   END else Result := 1004 ;
END ;

{==============================================================================================}
{==================================== OBJET TTPISYMBOL =========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  Create :
///////////////////////////////////////////////////////////////////////////////////////
Constructor TTPISYMBOL.Create ( AOwner : TComponent ; ATPI : String ) ;
BEGIN
Inherited Create(Aowner,ATPI) ;
// valeur par defaut
FComm.Speed := br9600 ;
FComm.DataBits := db8BITS ;
FComm.StopBits := sb1BITS ;
FComm.Parity := ptNone ;
// Valeurs des nombres de tentatives et des Timeout
ERecep := R0 ;
NbCon := 0 ;
NbRec := 0 ;
TMOMsg := TTPI_Timer.Create(Self) ; TMOMsg.MaxInterval := 1000 ; TMOMsg.Tag := 1 ;      // 1 seconde
TMOCar := TTPI_Timer.Create(Self) ; TMOCar.Tag := 3 ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Destroy :
///////////////////////////////////////////////////////////////////////////////////////
Destructor TTPISYMBOL.Destroy ;
BEGIN
TMOCar.Free ;
TMOMsg.Free ;
Inherited Destroy ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  ChargePortEtParams :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPISYMBOL.ChargePortEtParams(APort, AParams : String ) : Integer ;
BEGIN
Result := Inherited ChargePortEtParams(APort, AParams) ;
End ;
///////////////////////////////////////////////////////////////////////////////////////
//  GereReception :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPISYMBOL.GereReception : Integer ;
Var Trame, Rslt, Reponse  : String ;
    TimeExceded, Cancel   : Boolean ;
    pETX                  : Integer ;
BEGIN
NbCon := 0 ;
repeat
   Result := 0 ;
   Trame := '' ;
   Cancel := False ;
   ERecep := R1 ;
   NbRec := 0 ;
   Inc(NbCon) ;
   // R1 = Attente de connexion jusqu'à réception d'un ENQ
   repeat
     Rslt := ReceiveMess ;
     if Length(Rslt) > 1 then Rslt := Rslt[1] ;
     Application.ProcessMessages ;
   until (Rslt = Chr_ENQ) or (GetArret) ;
   // Traitement de la cause de sortie de l'attente
   if Rslt = Chr_ENQ then
      BEGIN
      // Connexion établie
      if Assigned(OnStartCom) then OnStartCom(Cancel) ;
      if Cancel then
         BEGIN
         Result := 1007 ;
         Exit ;
         END ;
      ERecep := R2 ;
      VideBuffer ;
      // Accusé de réception
      if not SendMessage(Chr_ACK) then Result := 1005 else
         // R2 = Réception des trames
         repeat
            ERecep := R2 ;
            if Result = 0 then
               BEGIN
               TMOMsg.Initialise ;             // T1 = timeout de 1 seconde
               repeat
                 Rslt := '' ;
                 Trame := ReceiveMess ;
                 // En cas de réception de caractères, on considère que le 1er caractère est un STX
                 if Trame <> '' then Rslt := Chr_STX ;
                 // Recherche du caractère de déconnexion
                 if (Pos(Chr_EOT, Trame) > 0) then Rslt := Chr_EOT ;
                 Application.ProcessMessages ;
                 TimeExceded := TMOMsg.Exceded ;
               until (TimeExceded) or (Rslt = Chr_STX) or (Rslt = Chr_EOT) ;
               // Traitement des caractères reçus
               if Rslt = Chr_STX then
                  BEGIN
                  // Traitement de la trame reçue - recherche du caractère de fin de trame
                  pETX := Pos(Chr_ETX, Trame) ;
                  if pETX > 0 then
                     BEGIN
                     // suppression des éventuels caractères derrière la fin de trame
                     Delete(Trame, pETX+1, Length(Trame)-pETX) ;
                     END else
                     BEGIN
                     // R3 = on lit la suite de la trame
                     ERecep := R3 ;
                     TMOCar.Initialise ;         // TD = timeout de 1, 3 ou 6 secondes selon la vitesse
                     TimeExceded := FALSE ;
                     while (pETX <> Length(Trame)) and (Not TimeExceded) do
                       BEGIN
                       Trame := ReceiveMess ;
                       while (Pos(Chr_EOT, Trame) > 0) do Delete(Trame, pos(Chr_EOT, Trame), 1) ;
                       while (Pos(Chr_ENQ, Trame) > 0) do Delete(Trame, pos(Chr_ENQ, Trame), 1) ;
                       while (Pos(Chr_STX, Trame) > 0) do Delete(Trame, pos(Chr_STX, Trame), 1) ;
                       Application.ProcessMessages ;
                       pETX := Pos(Chr_ETX, Trame) ;
                       if pETX > 0 then Delete(Trame, pETX+1, Length(Trame)-pETX) ;
                       TimeExceded := TMOCar.Exceded ;
                       END ;
                     END ;
                  if (not TimeExceded) and ((Rslt=Chr_EOT) or ((Rslt=Chr_STX) and (Pos(Chr_ETX, Trame)=Length(Trame)))) then
                     BEGIN
                     if Rslt=Chr_STX then
                        BEGIN
                        // suppression du fin de trame
                        pETX := Pos(Chr_ETX, Trame) ;
                        if pETX > 0 then Delete(Trame, pETX, Length(Trame)-pETX+1) ;
                        END ;
                     // Acquittement de la trame reçu
                     Reponse := Chr_ACK ;
                     NbRec := 0 ;
                     if Assigned(OnReceiveCom) then OnReceiveCom(Trame, Cancel) ;
                     if Cancel then
                        BEGIN
                        Result := 1007 ;
                        Reponse := Chr_EOT ;
                        ERecep := R4 ;
                        END ;
                     END else
                     BEGIN
                     // La trame reçue est incorrecte
                     Reponse := Chr_NAK ;
                     Inc(NbRec) ;
                     END ;
                  END else
                  BEGIN
                  Inc(NbRec) ;
                  if Rslt = Chr_EOT then
                     BEGIN
                     ERecep := R4 ;
                     Reponse := Chr_ACK ;
                     END else Reponse := Chr_NAK ;
                  END ;
               VideBuffer ;
               if not SendMessage(Reponse) then Result := 1005 ;
               END ;
         until (ERecep = R4) or (NbRec >= Max_NbRec) or (Result <> 0) or (GetArret) ; // Fin de R2
         if (ERecep <> R4) and (NbRec >= Max_NbRec) and (Result = 0) then
            BEGIN
            if GetArret then Result := 1008 else if Length(Trame) = 0 then Result := 1010 else Result := 1009 ;
            END ;
      END else
      BEGIN
      if GetArret then Result := 1008 else Result := 1006 ;
      END ;
until (ERecep = R4) or (NbCon >= Max_NbCon) or (Result <> 0) or (GetArret) ; // Fin de R1
VideBuffer ;
if (ERecep in [R2,R3,R4]) and (Assigned(OnStopCom)) then OnStopCom(Result, TPI_MsgErrDefaut(Result)) ;
END ;
///////////////////////////////////////////////////////////////////////////////////////
//  Transaction :
///////////////////////////////////////////////////////////////////////////////////////
Function TTPISYMBOL.Transaction( RemiseAttente : Boolean ) : Integer ;
BEGIN
if FComm.ComPort <> pnNone then
   BEGIN
   Repeat
     FComm.Connect ;
     // Définition des caractères de contrôle
     Chr_STX := chr(STX) ;
     Chr_ETX := chr(ETX) ;
     Chr_EOT := chr(EOT) ;
     Chr_ENQ := chr(ENQ) ;
     Chr_ACK := '%0' ;
     Chr_NAK := '%1' ;
     // Définition du timeout entre chaque caractère reçu en fonction de la vitesse de transmission
     TMOCar.MaxInterval := 3000 ; // 3 secondes
     Case Speed of
       br1200 : TMOCar.MaxInterval := 6000 ; // 6 secondes
       br9600 : TMOCar.MaxInterval := 1000 ; // 1 seconde ;
       END ;
     NbCon := 0 ;
     ERecep := R0 ;
     Result := GereReception ;
     ERecep := R0 ;
     FComm.Disconnect ;
   Until (not RemiseAttente) or (GetArret) ;
   END else Result := 1004 ;
END ;

End.

