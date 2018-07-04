unit lp_print;

interface

Uses MC_Erreur, LP_Base, CLasses, MC_Comm, Controls ;

type

  TLP_TesteEtatEvent  = Function ( Fois : Integer ; Var Err : TMC_Err ) : TImprimErreurs of Object ;
  TLP_BooleanEvent    = Function ( var err : TMC_Err ) : Boolean of Object ;
  TMC_ErreurEvent     = Function ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs of Object ;
  TLP_Translate       = Function ( Texte : String ) : String of Object ;

  TLPPrinterInterne = Class ( TComponent)
     Private
      FImprimante     : String ;
      FLargeur        : Integer ;
      FSortie         : TMC_CommportDriver ;
      FRecu           : String ;
      FLastError      : TImprimErreurs ;
      FLasTMC_Error   : TMC_Err ; 
      FCountfois      : integer ;
      FCountErrors    : Integer ;
      FSequenceLP     : TabSequenceLP ;
      FLastLigne      : Integer ;
      FStP            : TLPSqcesLigne ;
      Finitialise     : Boolean ;
      FOnInitialise   : TLP_BooleanEvent ;
      FOnTesteEtat    : TLP_TesteEtatEvent ;
      FOnErreur       : TMC_ErreurEvent ;
      FOnTranslate    : TLP_Translate ;
      InTest          : Integer ;
      FTranslate      : Boolean ;
      FTesteEtat      : Boolean ;
      LeDelayLong     : Integer ;
      LeDelayCourt    : Integer ;
      LeDelayConflitA : Integer ; //XMG 10/07/03
      FirstLigne      : Boolean ;
      FFBuffer        : String ;
      FInTestMateriel : Boolean ;
      LongMaxBuffer   : Integer ;
      ForceFlush      : Boolean ;
      Function TesterEtat : TImprimErreurs ;
      procedure LirePort (Sender: TObject; DataPtr: pointer; DataSize: integer ) ;
      Function  LireResultat ( Combien,Tipe : Integer ; var V : variant ) : Boolean ;
      Function  WriteInterne( St : String ; OkCRLF,Flush : Boolean ) : Boolean ;
      Function  FermePort : Boolean ;
      Function  OuvrirPort : Boolean ;
      Function  LanceMessage (Err : TMC_Err ; isQuestion : Boolean ) : Integer ;
      Function  GetIsPrinter  : Boolean ;
      Procedure SetLargeur    ( Value : Integer ) ;
      Function  GetSqce       (Indice : TLP_SqcesPossibles ; Activer : Boolean ) : String ;
      Function  GetSqceExtra  (Indice : TLP_SqcesPossibles ) : String ;
      Function  GetSqceGrp    (Groupe : Integer ) : String ;
      Function  GetGroupe     (Indice : TLP_SqcesPossibles ) : Integer ;
      Procedure SetLastLigne  ( Value : Integer ) ;
      Function  GetSqceInUse  (Indice : TLP_SqcesPossibles) : Boolean ;
      Procedure SetSqceInUse  (Indice : TLP_SqcesPossibles ; Value : Boolean ) ;
      Function  GetLastErrorW  : LongWord ;
      Procedure SetAttendre   ( AValue : Boolean) ;
      Procedure SetTesteEtat  ( AValue : Boolean) ;
      Function  GetAttendre   : Boolean ;
      //XMG 30/03/04 début
      Function  GetNomFichierSortie : String ;
      Procedure SetNomFichierSortie ( AValue : String ) ;
      //XMG 30/03/04 fin
     protected

      Procedure SetPort       ( AValue : TComPortNumber) ;   Virtual ;
      Function  GetPort       : TComPortNumber ;             Virtual ;
      Procedure SetBaudRate   ( AValue : TComPortBaudRate) ; Virtual ;
      Function  GetBaudRate   : TComPortBaudRate ;           Virtual ;
      Procedure SetDataBits   ( AValue : TComPortDataBits) ; Virtual ;
      Function  GetDataBits   : TComPortDataBits ;           Virtual ;
      Procedure SetStopBits   ( AValue : TComPortStopBits) ; Virtual ;
      Function  GetStopBits   : TComPortStopBits;            Virtual ;
      Procedure SetParity     ( AValue : TComPortParity) ;   Virtual ;
      Function  GetParity     : TComPortParity ;             Virtual ;
      Function  WriteLnLP     (St : String ; Flush : Boolean = FALSE ) : Boolean ; virtual ;
      Function  WriteLP       (St : String ; Flush : Boolean = FALSE ) : Boolean ; virtual ;
      Procedure SetHWProtocol ( AValue : TComPortHwHandshaking) ; Virtual ;
      Function  GetHWProtocol : TComPortHwHandshaking ;           Virtual ;
      Procedure SetSWProtocol ( AValue : TComPortSwHandshaking) ; Virtual ;
      Function  GetSWProtocol : TComPortSwHandshaking ;           Virtual ;
      Procedure ChargeSequences ;                                 virtual ;
      Function  CodeCodeBarres( Masque : String) : String ;       virtual ;
     Public
      Constructor Create ( AOwner : TComponent) ; Overload ; Override ;
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; reintroduce; overload ; virtual ;
      Destructor  destroy ; override ;
      Procedure   InitLastligne ;
      Procedure   InitNewLigne ;
      Function    InitImprimante ( Var Err : TMC_Err ) : Boolean ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; virtual ;
      Function    OuvrirImprimante( Var Err : TMC_Err ) : Boolean ;
      Function    FermeImprimante ( Var Err : TMC_Err ) : Boolean ;
      Procedure   FormatPolice ( zn : TLP_ZoneImp) ; //Seq : TLP_SetSqcesPossibles ; Colonne, Longeur : Integer ; Extra : String ) ;
      //Procedure   FormatPolice ( Seq : TLP_SetSqcesPossibles ; Colonne, Longeur : Integer ; Extra : String ) ;
      Procedure   AttentResponse ;
      Function    EcrireLigne(St : String ; Ligne : Integer ) : Boolean ;
      Function    isConnected : Boolean ;
      Function    SautPage : Boolean ;
      Function    FaireTranslate (St : String ) : String ;
      Procedure   Demarre ;

      Property Imprimante       : String  read FImprimante ;
      Property Port             : TComPortNumber     Read GetPort          write SetPort     ;
      property Speed            : TComPortBaudRate   read GetBaudRate      write SetBaudRate ;
      property DataBits         : TComPortDataBits   read GetDataBits      write SetDataBits ;
      property StopBits         : TComPortStopBits   read GetStopBits      write SetStopBits ;
      property Parity           : TComPortParity     read GetParity        write SetParity   ;
      Property HWProtocol       : TComPortHwHandshaking read GetHWProtocol write SetHWProtocol ;
      Property SWProtocol       : TComPortSwHandshaking read GetSWProtocol write SetSWProtocol ;
      Property IsPrinter        : Boolean            Read GetIsPrinter ;
      Property TesteEtat        : Boolean            read FTesteEtat       write SetTesteEtat ;
      Property InTestMateriel   : Boolean            read FInTestMateriel  write FInTestMateriel ;

      Property Sqce             [Indice : TLP_SqcesPossibles ; Activer : Boolean ]  : String  read GetSqce ;
      Property SqceGrp          [Groupe : Integer ]  : String  read GetSqceGrp ;
      Property SqceExtra        [Indice : TLP_SqcesPossibles ]  : String  read GetSqceExtra ;
      Property Groupe           [Indice : TLP_SqcesPossibles ]  : Integer read GetGroupe ;
      Property SqceInUse        [Indice : TLP_SqcesPossibles]  : Boolean read GetSqceInUse write SetSqceInUse ;
      Property LastLigne        : Integer Read FLastLigne write SetLastLigne ;
      Property LastError        : TImprimErreurs read FLastError ;
      Property LasTMC_Error     : TMC_Err read FLasTMC_Error ;
      Property LastErrorW       : Longword read GetLastErrorW ;
      property Largeur          : Integer read FLargeur write SetLargeur ;
      Property Initialise       : Boolean read FInitialise write Finitialise ;
      property AttendreResponse : boolean read GetAttendre write SetAttendre  ;
      Property Translate        : Boolean read FTranslate  write FTranslate ;

      Property NomFichierSortie : String Read GetNomFichierSortie Write SetNomFichierSortie ; //XMG 30/03/04

      Property OnInitialise : TLP_BooleanEvent   read FOnInitialise write FOnInitialise ;
      Property OnTesteEtat  : TLP_TesteEtatEvent read FOnTesteEtat  write FOnTesteEtat ;
      Property OnErreur     : TMC_ErreurEvent    read FOnErreur     write FOnErreur ;
      Property OnTranslate  : TLP_Translate      read FOnTranslate  write FOnTranslate ;
     End ;

type
  TLP_EPSONTM210 = Class (TLPPrinterInterne)
     Protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    InitPrinter ( var err : TMC_Err ) : Boolean ;
      Function    TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    TranslateASCII ( Texte : String) : String ;
      End ;
type
  TLP_EPSONTM675 = Class (TLPPrinterInterne)
     Protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    InitPrinter ( var err : TMC_Err ) : Boolean ;
      Function    TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    TranslateASCII ( Texte : String) : String ;
     End ;
type                                             
  TLP_EPSONTM950 = Class (TLPPrinterInterne)
     protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    InitPrinter ( var err : TMC_Err ) : Boolean ;
      Function    TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    TranslateASCII ( Texte : String) : String ;
     End ;
type
  TLP_EPSONTM5000 = Class (TLPPrinterInterne)
     protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    InitPrinter ( var err : TMC_Err ) : Boolean ;
      Function    TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    TranslateASCII ( Texte : String) : String ;
     end ;
type
  TLP_EPSONTM88II = Class (TLPPrinterInterne)
     protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    InitPrinter ( var err : TMC_Err ) : Boolean ;
      Function    TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    TranslateASCII ( Texte : String) : String ;
      End ;
type TLP_STARTSP700 = Class(TLP_EPSONTM88II) ;
type
  TLP_EPSONTM300 = Class (TLPPrinterInterne)
     protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    InitPrinter ( var err : TMC_Err ) : Boolean ;
      Function    TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    TranslateASCII ( Texte : String) : String ;
      End ;
type
  TLP_EPSONTM6000 = Class (TLPPrinterInterne)
     protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    InitPrinter ( var err : TMC_Err ) : Boolean ;
      Function    TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
      Function    TranslateASCII ( Texte : String) : String ;
      End ;

type
  TLP_EXPORTERFICHIER= Class (TLPPrinterInterne)
     Protected
      Procedure   ChargeSequences ; Override ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; Override ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ; override ;
      Function    TranslateASCII ( Texte : String) : String ;
      End ;

implementation

uses Math, Sysutils, HCtrls, Hent1, MC_Lib, HDebug,
{$IFNDEF EXPORTASCII}
     MC_Admin,BlqMatPV,
{$ENDIF EXPORTASCII} //XMG 30/03/04
     HMsgBox, Graphics ;

Const MaxFFbuffer = 800 ; 

//////////////////////////////////////////////////////////////////////////////////////////////////
Function Insere(Car : String ; var Des : String ; ou : integer) : String ;
Var i : Integer ;
Begin
if ou>length(des) then For i:=length(des)+1 to ou-1 do des:=des+' ' ;
Insert(Car,Des,ou) ;
Result:=Des ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPPrinterInterne.Create ( AOwner : TComponent) ;
Begin
Raise EComponentError.Create(MC_MsgErrDefaut(1062)) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPPrinterInterne.Create ( AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner) ;
FImprimante:=AImprimante ;
FSortie:=TMC_CommportDriver.Create(Self) ;
FSortie.AttendreResponse:=FALSE ;
FSortie.OnReceiveData:=LirePort ;
Port:=pnNone ;
FRecu:='' ;
FLastError:=ieOk ;
fillchar(FLasTMC_Error,sizeof(FLasTMC_Error),#0) ;
FCountErrors:=-1 ;
FCountFois:=-1 ;
FStp:=TLPSqcesLigne.Create ;
Largeur:=2000 ; //initialise les arrays de la ligne
ChargeSequences ;
InTest:=0 ;
Translate:=FALSE ;
TesteEtat:=FALSE ;
LeDelayLong:=0 ;
LeDelayCourt:=0 ;
LeDelayConflitA:=25 ; //XMG 10/07/03
ForceFlush:=TRUE ;
LongMaxBuffer:=MaxFFBuffer ;
FirstLigne:=TRUE ;
InTestMateriel:=FALSE ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

destructor TLPPrinterInterne.Destroy ;
Begin
FSortie.Free ;
FStp.Free ;
Inherited destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetLargeur ( Value : Integer ) ;
Begin
If Value=Largeur then exit ;
Value:=MinIntValue([2000,Value]) ;
FLargeur:=Value ;
FStp.Largeur:=Largeur ;
InitNewLigne ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetPort ( AValue : TComPortNumber) ;
Begin
FSortie.ComPort:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetPort : TComPortNumber ;
Begin
Result:=FSortie.ComPort ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetBaudRate ( AValue : TComPortBaudRate) ;
Begin
FSortie.Speed:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetBaudRate : TComPortBaudRate ;
Begin
Result:=FSortie.Speed;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetDataBits ( AValue : TComPortDataBits) ;
Begin
FSortie.DataBits:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetDataBits : TComPortDataBits ;
Begin
Result:=FSortie.DataBits ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetStopBits ( AValue : TComPortStopBits) ;
Begin
FSortie.StopBits:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetStopBits : TComPortStopBits;
Begin
Result:=FSortie.StopBits ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetParity ( AValue : TComPortParity) ;
Begin
FSortie.Parity:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetParity : TComPortParity ;
Begin
Result:=FSortie.Parity ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetIsPrinter : Boolean ;
BEgin
Result:=FSortie.IsPrinter ;
End ; 
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
var st  : String ;
    i,j,a : Integer ;
Begin
Result:=FALSE ;
APort:=Uppercase(Trim(APort)) ;
if (APort='') or (APort='AUC') then Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1009) ; Exit ; end ;
{$IFNDEF EXPORTASCII}
if (Imprimante<>LP_CodeImprimanteExportFichier) and (not UsesTPV(Imprimante,mcPrinter)) then Begin Err.code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1010) ; exit ; end ;
{$ENDIF EXPORTASCII} //XMG 02/04/04
port:=TComPortNumber(Pos(';'+APort+';',';AUC;FTX;CO1;CO2;CO3;CO4;LP1;LP2;') div 4) ;
if Port = pnFile then NomFichierSortie:=ReadTokenPipe(AParams,';') else //XMG 30/03/04
if port in [pnCOM1,pnCOM2,pnCOM3,pnCOM4] then
   Begin
   i:=0 ;
   while Aparams<>'' do
     Begin
     inc(i) ;
     St:=ReadTokenPipe(AParams,';') ;
     if St<>'' then
        Begin
        Case i of
          1 : Speed:=TComPortBaudRate(Valeuri(St)-1) ;
          2 : Parity:=TComPortParity(Valeuri(St)-1) ;
          3 : DataBits:=TComPortDataBits(Valeuri(St)-1) ;
          4 : StopBits:=TComPortStopBits(Valeuri(St)-1) ;
          5 : begin
              j:=Valeuri(St)-1 ;
              if j<2 then HWProtocol:=TComPortHwHandshaking(j)
                 else SWProtocol:=TComPortSwHandshaking(j-2) ;
              End ;
          6 : TesteEtat:=TrueFalseBo(St) ;
          End ;
        End ;
     End ;
   End ;
A:=GetSynRegKey('LeDelayLong'+Classname, LeDelayLong, True) ; if A>-1 then LeDelayLong:=a ;
if V_PGI.Debug then DEBUG('LeDelayLong'+Classname+' = '+IntToStr(A)) ;
A:=GetSynRegKey('LeDelayCourt'+Classname, LeDelayCourt, True) ; if A>-1 then LeDelayCourt:=a ;
if V_PGI.Debug then DEBUG('LeDelayCourt'+Classname+' = '+IntToStr(A)) ;
A:=GetSynRegKey('LongBuffer'+Classname, LongMaxBuffer, True) ; if A>-1 then LongMaxBuffer:=a ;
if V_PGI.Debug then DEBUG('LongBuffer'+Classname+' = '+IntToStr(A)) ;
a:=GetSynRegKey('ForceFlush'+Classname, a, True) ; if (a in [0,1]) and (a<>ord(ForceFlush)) then ForceFlush:=(a=1) ;
if V_PGI.Debug then DEBUG('ForceFlush'+Classname+' = '+IntToStr(A)) ;
//XMG 10/07/03 début
A:=GetSynRegKey('ConflitAfficheur'+Classname, LeDelayConflitA, True) ; if A>-1 then LeDelayConflitA:=a ;
if V_PGI.Debug then DEBUG('ConflitAfficheur'+Classname+' = '+IntToStr(A)) ;
//XMG 10/07/03 fin
Result:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.FermePort : Boolean ;
Begin
Result:=FSortie.Disconnect ;
if Not result then FLastError:=ieError ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.OuvrirPort : Boolean ;
Begin
Result:=FALSE ;
if Imprimante='' then Exit ;
Result:=FSortie.Connect ;
if Not result then FLastError:=ieError ;
FirstLigne:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.isConnected : Boolean ;
Begin
Result:=FSortie.Connected ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetSqce (Indice : TLP_SqcesPossibles ; Activer : Boolean ) : String ;
Begin
result:='' ;
if (Indice in [MinLP_SqceCtrl..MaxLP_SqceCtrl]) and (Not Activer) then Activer:=TRUE ;
If FSequenceLP[Indice].Groupe=0 then Result:=FSequenceLP[Indice].Sqce[Activer] ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetSqceExtra (Indice : TLP_SqcesPossibles ) : String ;
Begin
result:=FSequenceLP[Indice].Extra ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinterInterne.GetSqceGrp (Groupe : Integer ) : String ;
Var i     : TLP_SqcesPossibles ;
    Val,c : Integer ;
Begin
Result:='' ;
if (Groupe>0) then
   Begin
   Val:=0 ; c:=-1 ;
   For i:=MinLP_SqceUser to MaxLP_SqceUser do
       if (FSequenceLP[i].Groupe=Groupe) then
          Begin
          Val:=Val+(FSequenceLP[i].Valeur*Ord(FSequenceLP[i].InUse)) ;
          if c=-1 then c:=ord(i) ;
          End ;
   Result:=Format(FSequenceLP[TLP_SqcesPossibles(C)].Sqce[TRUE],[chr(Val)]) ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinterInterne.GetGroupe (Indice : TLP_SqcesPossibles ) : Integer ;
Begin
Result:=FSequenceLP[Indice].Groupe ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinterInterne.GetSqceInUse (Indice : TLP_SqcesPossibles) : Boolean ;
Begin
if Indice in [MinLP_SqceUser..MaxLP_SqceUser] then Result:=FSequenceLP[Indice].InUse
   else Result:=FALSE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetSqceInUse (Indice : TLP_SqcesPossibles ; Value : Boolean ) ;
Begin
if Value=SqceInUse[Indice] then exit ;
FSequenceLP[Indice].Inuse:=Value ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetLastErrorW : LongWord ;
Begin
Result:=FSortie.LastError ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.InitLastligne ;
Begin
FLastLigne:=-1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetLastLigne ( Value : Integer ) ;
Begin
if Value<=LastLIgne then exit ;
FLastLigne:=Value ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.OuvrirImprimante ( Var Err : TMC_Err ) : Boolean ;
Begin
Result:=OuvrirPort ;
delay(100) ;
if (Not Result) then
   Begin
   if (Port<>pnFile) then Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1011) ; end
      else Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1012) ; end
   End ;
If (Result) and (Initialise) then Result:=InitImprimante(Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.FermeImprimante( Var Err : TMC_Err ) : Boolean ;
Begin
if Length(FFbuffer)>0 then WriteInterne('',FALSE,TRUE) ;
delay(100) ;
Result:=FermePort ;
if (not Result) then
   Begin
   if (Port<>pnFile) then Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1013) ; End
      else Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1014) ; End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPPrinterInterne.LirePort (Sender: TObject; DataPtr: pointer; DataSize: integer ) ;
Var St : String ;
    P  : PChar ;
Begin
St:='' ;   P:=DataPtr;
While DataSize>0 do
  Begin
  St:=St+P^ ;
  Inc(P) ;
  Dec(DataSize) ;
  End ;
FRecu:=FRecu+St ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.FormatPolice ( zn : TLP_ZoneImp) ; //Seq : TLP_SetSqcesPossibles ; Colonne, Longeur : Integer ; Extra : String ) ;
Var CC,Max : Integer ;
Begin
if (Zn.Colonne>=0) and (Zn.Colonne<=Largeur) then
   Begin
   if Zn.Police=[spCodeBarre] then
      Begin //Gestion Standard "EPSON" pour les codearresimages
      FStP[Zn.Colonne-1]:=Zn.Police ;
      FStP.Valeurs[Zn.Colonne-1]:=inttostr(Zn.NbrLignes)+';'+CodeCodeBarres(zn.Extra)+';'+inttostr(zn.taille)+';' ;  //NbrLignes+';'+Code du codeBArre+';'+Taille des données
      End ;
   if Zn.Police=[spBitMap] then
      Begin //Gestion Standard "EPSON" pour les images
      FStP[Zn.Colonne-1]:=Zn.Police ;
      FStP.Valeurs[Zn.Colonne-1]:=inttostr(Zn.Taille) ;
      End else //Gestion Standard "EPSON" pour les Polices
   if ZN.Taille>0 then
      Begin
      Max:=MinIntValue([Largeur-Zn.Colonne,Zn.Taille]) ;
      for cc:=0 to Max-1 do FStP[ZN.Colonne+cc-1]:=Zn.Police ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.AttentResponse ;
Begin
if V_PGI.Debug then DEBUG('Attente reponse de '+IntToStr(LeDelayLong)) ;
if LeDelayLong>0 then Delay(LeDelayLong) ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetAttendre   ( AValue : Boolean) ;
Begin
FSortie.AttendreResponse:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetTesteEtat( AValue : Boolean) ;
Begin
FTesteEtat:=AValue ;
AttendreResponse:=TesteEtat ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinterInterne.GetAttendre   : Boolean ;
Begin
Result:=FSortie.AttendreResponse ;
End ;
///////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 30/03/04 début
Function TLPPrinterInterne.GetNomFichierSortie : String ;
Begin
Result:=FSortie.NomFichierSortie ;
End ;
///////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetNomFichierSortie ( AValue : String ) ;
Begin
FSortie.NomFichierSortie:=AValue ; 
End ;
//XMG 30/03/04 fin
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetHWProtocol ( AValue : TComPortHwHandshaking) ;
Begin
FSortie.HwHandshaking:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetHWProtocol : TComPortHwHandshaking ;
Begin
Result:=FSortie.HwHandshaking ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.SetSWProtocol ( AValue : TComPortSWHandshaking) ;
Begin
FSortie.SWHandshaking:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.GetSWProtocol : TComPortSWHandshaking ;
Begin
Result:=FSortie.SWHandshaking ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.ChargeSequences ; 
Begin
Fillchar(FSequenceLP,sizeof(FSequenceLP),#0) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.CodeCodeBarres( Masque : String) : String ;
Var Cbarre : String ;
Begin
Result:='' ;
if stricomp(LPStCBarre,PChar(Copy(Masque,1,length(LPStCBarre))))=0 then
   Begin
   CBArre:=copy(MAsque,length(LPStCBarre)+1,length(Masque)) ;
   if stricomp('(UPC-A)',pchar(CBarre))=0   then Result:='65' else
   if stricomp('(UPC-E)',pchar(CBarre))=0   then Result:='66' else
   if stricomp('(EAN13)',pchar(CBarre))=0   then Result:='67' else
   if stricomp('(EAN8)',pchar(CBarre))=0    then Result:='68' else
   if stricomp('(CODE39)',pchar(CBarre))=0  then Result:='69' else
   if stricomp('(ITF)',pchar(CBarre))=0     then Result:='70' else
   if stricomp('(CODABAR)',pchar(CBarre))=0 then Result:='71' else
   if stricomp('(CODE93)',pchar(CBarre))=0  then Result:='72' else
   if stricomp('(CODE128)',pchar(CBarre))=0 then Result:='73' else
      ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.LireResultat ( Combien,Tipe : Integer ; var V : variant ) : Boolean ;
Var VV : Variant ;
Begin
V:=#0 ;
Result:=FALSE ;
if FRecu<>'' then
   Begin
   if (length(FRecu)<Combien) then Combien:=Length(FRecu) ;
   VV:=Copy(FRecu,length(FRecu)-Combien+1,Combien) ; //On renvoie les derniers combien caractères
   V:=Varastype(VV,Tipe) ;
   FRecu:='' ;
   Result:=TRUE ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.WriteLnLP (St : String ; Flush : Boolean = FALSE  ) : Boolean ;
Begin
Result:=WriteInterne(St,TRUE,Flush) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.WriteLP (St : String ; Flush : Boolean = FALSE ) : Boolean ;
Begin
Result:=WriteInterne(St,FALSE,Flush) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.WriteInterne( St : String ; OkCRLF,Flush : Boolean ) : Boolean ;
Var EtatImp : TImprimErreurs ;
{$IFNDEF EXPORTASCII}
    Mat     : String ;
{$ENDIF EXPORTASCII} //XMG 30/03/04
Begin
flush:=(Flush) or (ForceFlush) ;
if Length(St)>Largeur then delete(St,Largeur,Length(St)) ;
if OkCRLF then
   Begin
   While (Length(St)>0) and (st[length(St)]=' ') and (FStp[Length(St)-2]=[]) do Delete(St,length(St),1) ;
   St:=ST+Sqce[spCR,TRUE]+Sqce[spLF,TRUE] ;
   End ;
if (not Flush) and (length(FFBuffer)+Length(st)<=LongMaxBuffer) then Begin FFBuffer:=FFBuffer+St ; Result:=TRUE ; End else
   Begin
   if Flush then FFbuffer:=FFBuffer+St ;
   if Length(FFBuffer)<=0 then Result:=TRUE else
      Begin
      FCountFois:=FCountFois*ord(InTest<>0) ;
      FFBuffer:=Sqce[spActivate,TRUE]+FFBuffer+Sqce[spDesactivate,TRUE] ;
      Result:=FSortie.SendString(FFBuffer) ;
      {$IFNDEF EXPORTASCII}
      if (Imprimante<>LP_CodeImprimanteExportFichier) and (V_MC.ConflitImprimante) then delay(LeDelayConflitA) ; //XMG 10/07/03
      {$ENDIF EXPORTASCII} //XMG 30/03/04
      repeat
        EtatImp:=TesterEtat
      until ((Result) and (EtatImp<>ieRecovery)) or (EtatImp in [ieError,ieErrorPaper]) ;  //Teste l'état après de imprimer la ligne
      if FirstLigne then FirstLigne:=FALSE ;
      Result:=(Result) and (EtatImp in [ieOk,ieWarningPaper]) and (LastError<>ieError) ;
      //XMG 30/03/04 début
      {$IFNDEF EXPORTASCII}
      if (Imprimante<>LP_CodeImprimanteExportFichier) and (LastError=ieError) then //XMG 02/04/04
          Begin
          mat:='' ;
          if InTestMateriel then mat:='PRT;'+Imprimante+';|' ;
          bloquematerielpv(mat)  ; //erreur mateirel
          End ;
      {$ENDIF EXPORTASCII}
      //XMG 30/03/04 fin
      End ;
   FFBuffer:='' ;
   if Not Flush then FFBuffer:=St ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinterInterne.LanceMessage (Err : TMC_Err ; isQuestion : Boolean ) : Integer ;
Begin
Result:=mrNo ;
if IsQuestion then Result:=PGIAsk(Err.Libelle,LPMessageCap) else
   Begin
   PGIBox(Err.Libelle,LPMessageCap) ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.TesterEtat : TImprimErreurs ;
Var Err    : TMC_Err ;
    SysErr : Boolean ;
Begin
SysErr:=(FSortie.LastError<>0) ;
Fillchar(err,sizeof(err),#0) ;
inc(InTest) ;
if Not SysErr then
   Begin
   Result:=IeOk ;
   if FSortie.ComPort<>pnFile then
      Begin
      if (Assigned(OnTesteEtat)) and (TesteEtat) then
         begin
         Result:=OnTesteEtat(1,Err) ;
         FLasTMC_Error:=Err ;
         end else if LeDelayCourt>0 then Delay(LeDelayCourt) ;
      End ;
   End else
   Begin
   Result:=ieError ;
   Err.Libelle:=Format(MC_MsgErrDefaut(1024),[Fsortie.LastError,SysErrorMessage(FSortie.LastError)]) ;
   End ;
inc(FCountErrors,ord(Result in [ieWarning,ieWarningPaper,ieRecovery,iePowerOff])) ;
inc(FCountFois,1) ;
FLastError:=Result ; //On garde le vrai erreur dans LastError et après s'il faut on bricole sur le Resultat
if ((LastError=ieRecovery) and (FCountErrors>=NbRecoverys)) or
   ((LastError=ieWarning)  and (FCountErrors>=NbWarnings))  or
   (FCountFois>=NbErreursEnsemble)                          then Result:=ieError ;
if (Result<>ieOk) and ((Result in [ieError,ieErrorPaper,ieCover,ieOffLine,iePowerOff]) or ((FCountErrors mod NbErrorsMessage)=0)) then
   Begin
   if (SysErr) then
      Begin
      if LanceMessage(Err,TRUE)=mrYes then Result:=ieRecovery ;
      End else
      Begin
      if (Err.code=0) then Err.Code:=1011 ;
      if (Err.code<>0) and (Trim(Err.Libelle)='') then Err.Libelle:=MC_MsgErrDefaut(Err.Code) ;
      if FCountFois>=NbErreursEnsemble then Err.Libelle:=Err.Libelle+MC_MsgErrDefaut(1058) ;
      if ((FirstLigne) or (LastError=iePowerOff)) or (LanceMessage(Err,(Result=ieError))=mrYes) then
         Begin
         Result:=ieError ;
         if ((FirstLigne) or (LastError=iePowerOff)) then FLastError:=Result ;  //cas special, utilisé pour detecte le mat.....
         End else
      if Assigned(OnErreur) then Result:=OnErreur(Result,Err) ;
      End ;
   End ;
FCountErrors:=FCountErrors*ord((Result<>ieOK) and (not (LastError in [ieError,ieErrorpaper,iePowerOff,ieCover,ieOffLine]))) ;
Dec(InTest) ; if InTest<0 then InTest:=0 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.InitImprimante ( Var Err : TMC_Err ) : Boolean ;
Var Ok : Boolean ;
Begin
result:=TRUE ;
if FSortie.ComPort=pnFile then exit ;
OK:=FSortie.Connected ;
if not Ok then Result:=OuvrirImprimante(Err) ;
if Result then
  Repeat
    if Assigned(OnInitialise) then Result:=OnInitialise(err) ; 
  until (Result) or ((Not Result) and (LastError<>ieErrorPaper)) ;
If (Not Ok) and (Result) then Result:=FermeImprimante(Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.InitNewLigne ;
Begin
FStp.Clear ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.SautPage : Boolean ;
Begin
writeLP(Sqce[spSautPage,TRUE]) ;
Result:=TRUE ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinterInterne.Demarre ;
Begin
FFBuffer:='' ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.FaireTranslate ( St : String ) : String ;
Begin
Result:=St ;
if Translate then
   Begin
   if Assigned(OnTranslate) then Result:=OnTranslate(St) else Result:=Ansi2Ascii(St) ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinterInterne.EcrireLigne(St : String ; Ligne : Integer ) : Boolean ;
Var Txt,ss,Gr,Vals : String ;
    c,Last,Haut    : Integer ;
    FF             : TLP_SetSqcesPossibles ;
    Ok,Force       : Boolean ;
    i              : TLP_SqcesPossibles ;
Begin
Result:=TRUE ;
Force:=(Trim(St)<>St) and (St<>'') ; //Petit probleme avec les Sequences d'echapement et la Trim()
txt:=' ' ;
Last:=LastLigne ;
while (last+1<ligne) and (Result) do
  begin
  Result:=(Result) and (WriteLnLP(Txt)) ;
  if trim(txt)<>'' then txt:=' ' ;
  inc(Last) ;
  End ;
if Result then
   Begin
   c:=0 ; Txt:='' ;
   while(c<length(St)) do
     begin
     inc(c) ;
     if c<=Largeur then
        Begin
        Gr:='' ;
        FF:=FStP[c-1] ;
        if FF=[spBitmap] then
           Begin
           ss:=Sqce[spBitmap,TRUE] ;
           ss:=Format(ss,[chr(Valeuri(FStP.valeurs[c-1]))]) ;
           Txt:=Txt+ss+Copy(St,c,valeuri(FStP.valeurs[c-1])) ;
           inc(c,Valeuri(FStP.valeurs[c-1])-1) ;
           End else
           Begin
           if FF=[spCodeBarre] then
              Begin
             {
             +#29+'h%s'   //Hauteur du codebarre (Nombre points * lignes)
             +#29+'H%s'   //Position u HRI
             +#29+'f0'    //%s' //Font à utiliser
             +#29+'w'+#1  //%s' //Longeur du module....
             +#29+'k%s'  //Code du codebarre plus valeur du CodeBarre à imprimer
             }
              ss:=Sqce[spCodeBarre,TRUE] ;
              Haut:=Valeuri(gtfs(SqceExtra[spCodeBarre],';',1))  ;
              Vals:=FStP.valeurs[c-1] ;
              ss:=Format(ss,[chr(Haut*valeuri(gtfs(Vals,';',1))),                          // Hauteur
                             chr(2*ord(valeuri(gtfs(Vals,';',1))>1)),                      // Position HRI
                             chr(valeuri(gtfs(Vals,';',2)))                                // Code du codebarre
                            +chr(valeuri(gtfs(Vals,';',3)))                                //  + longeur des données
                            +Copy(St,c,valeuri(gtfs(FStP.valeurs[c-1],';',3)))]) ;         //  + Valeur du codebarre à imprimer
              Txt:=Txt+ss ;
              inc(c,valeuri(gtfs(Vals,';',3))-1) ;
              End else
              Begin
              for i:=MinLP_SqceUser to MaxLP_SqceUser do
                Begin
                Ok:=(i in FF) ;
                if (Ok<>SqceInUse[i]) then
                   begin
                   SqceInUse[i]:=Ok ;
                   if (Groupe[i]=0) then Txt:=Txt+Sqce[i,Ok]
                      else if pos(';'+inttostr(Groupe[i])+';',';'+Gr)<=0 then Gr:=Gr+inttostr(Groupe[i])+';' ;
                   End ;
                End ;
              while trim(Gr)<>'' do
                 Begin
                 ss:=Readtokenst(Gr) ;
                 Txt:=Txt+SqceGrp[valeuri(ss)] ;
                 End ;
              Txt:=Txt+St[c] ;
              End ;
           End ;
        End ;
     End ;
   if (trim(Txt)<>'') or (Force) then Result:=WritelnLP(Txt) else result:=TRUE ;
   If Result then LastLigne:=Ligne ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLP_EPSONTM210.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
OnTesteEtat:=TestePrinter ;
OnInitialise:=InitPrinter ;
OnErreur:=IlyaErreur ;
Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM210.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
LongMaxBuffer:=60 ;
LeDelayLong:=200 ;
LeDelayCourt:=250 ;

Result:=Inherited ChargePortEtParams(APort,AParams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EPSONTM210.ChargeSequences ;
Begin
Inherited ;
FSequenceLP[spCR]             .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]             .Sqce[TRUE] :=#10 ;

FSequenceLP[spInitialise]     .Sqce[TRUE] :=#27+'@'+#27+'3'+#16+#27+'R'+#1+#27+'t'+#19 ;

FSequenceLP[spSautPage]       .Sqce[TRUE] :='' ; //On n'a pas de saut de page en rouleau

FSequenceLP[spStatus]         .Sqce[TRUE] :=#16+#4 ; //1,2,3,4

if Port in [pnLPT1,pnLPT2] then FSequenceLP[spSensorPaperOut].Sqce[TRUE]:=#27+'c3'+#15 ; //Seulement si interface parallel

FSequenceLP[spReStart]        .Sqce[TRUE] :=#16+#5 ; //0,2

FSequenceLP[spActivateInterligne]    .Sqce[TRUE] :=#27+'<'+#27+'3'+#24 ;
FSequenceLP[spDesactivateInterligne] .Sqce[TRUE] :=#27+'<'+#27+'3'+#16 ;    

FSequenceLP[spTotalCut]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VA'+#1 ;

FSequenceLP[spPartialCut]     .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VB'+#1 ;

FSequenceLP[spActivate]       .Sqce[TRUE] :=#27+'='+#1 ;

FSequenceLP[spDesactivate]    .Sqce[TRUE] :=#27+'='+#2 ;

FSequenceLP[sp17Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Groupe:=1 ;
FSequenceLP[sp17Cpi]          .Valeur:=32 ;

FSequenceLP[sp12Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Groupe:=1 ;
FSequenceLP[sp12Cpi]          .Valeur:=1 ;

FSequenceLP[sp10Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Groupe:=1 ;
FSequenceLP[sp10Cpi]          .Valeur:=0 ;

FSequenceLP[sp05Cpi]           .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Groupe:=1 ;
FSequenceLP[sp05Cpi]           .Valeur:=48 ;

FSequenceLP[spUnderLine]      .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spUnderLine]      .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[spUnderline]      .Groupe:=1 ;
FSequenceLP[spUnderline]      .Valeur:=128 ;

FSequenceLP[spBold]           .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spBold]           .Sqce[TRUE] :=#27+'!%s' ; 
FSequenceLP[spBold]           .Groupe:=1 ;
FSequenceLP[spBold]           .Valeur:=8 ;

FSequenceLP[spItalic]         .Sqce[FALSE]:=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#0 ;
FSequenceLP[spItalic]         .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#1 ;
FSequenceLP[spItalic]         .Groupe:=0 ;
FSequenceLP[spItalic]         .Valeur:=0 ;

FSequenceLP[spBitMap]         .Sqce[FALSE]:='' ;
FSequenceLP[spBitMap]         .Sqce[TRUE] :=#27+'*'+#1+'%s'+#0 ;
FSequenceLP[spBitMap]         .Groupe:=0 ;
FSequenceLP[spBitMap]         .Valeur:=0 ;

FSequenceLP[spPrintBmp]       .Sqce[TRUE] :='' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM210.InitPrinter ( var err : TMC_Err ) : Boolean ;
Var St : String ;
Begin
Result:=FALSE ;
St:=Sqce[spInitialise,TRUE] ; if St<>'' then Result:=WriteLP(St,TRUE) ;
if Result then
   Begin
   St:=Sqce[spSensorPaperOut,TRUE] ; if St<>'' then Result:=writeLP(St,TRUE) ;
   End ;
if not result then err:=LasTMC_Error ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLP_EPSONTM210.TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
Var tt   : String ;
    b    : Byte ;
    R    : String ;
    V    : Variant ;
Begin
Result:=ieOk ;
if Fois=1 then err.code:=0 ; err.Libelle:='' ;
tt:=Sqce[spActivate,TRUE]+Sqce[spstatus,TRUE]+chr(Fois)+Sqce[spDesactivate,TRUE] ;
if FSortie.SendString(tt) then
   Begin
   Attentresponse ;
   if lireResultat(1,VarString,V) then
      Begin
      R:=VString(V) ;
      if Length(R)=1 then b:=ord(R[1]) else B:=18 ;
      if (B<>(18+4*ord(Fois=1))) then
         Case Fois of
           1 : Begin //Status de l'imprimante
               If (B and 32)=32 then //L'imprimante essaie de se recuperer
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1015) ;
                  End else
               if (B and 8)=8 then //L'imprimante est Off-Line
                  Begin
                  Result:=TestePrinter(2,Err) ;
                  if Result=ieOk then //Si aucun erreur est rapporté au deuxième niveau alors c'est OFF-Line
                     Begin
                     Result:=ieOffLine ;
                     Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1016) ;
                     End ;
                  End ;
               End ; //3: ....
           2 : Begin //Status de l'imprimante quand elle est en Off-Line
               if (B and 64)=64 then //Error
                  Result:=TestePrinter(3,Err) else
               if (B and 32)=32 then //End of paper
                  Result:=TestePrinter(4,Err) {else
               if (B and 4)=4 then //Cover Open
                  Begin
                  Result:=ieCover ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1017) ; Exit ;
                  End ;  XMG Il n'est pas sûr que ce bit soit bien initialise (Voir documentation) }
               End ; //3: ....
           3 : Begin //Cause de l'erreur de l'imprimante
               If (B and 32)=32 then //
                  Begin
                  Result:=ieError ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1018) ;
                  End else
               If (B and 64)=64 then //
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1019) ;
                  End else
               if ((B and 4)=4) or ((B and 8)=8) then
                  Begin
                  Result:=ieWarning ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1020) ;
                  End ;
               End ; //3: ....
           4 : Begin //Causse de la fin du papier
               If (B and 96)=96 then
                  Begin
                  Result:=ieErrorPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1021) ;
                  End else
               If (B and 12)=12 then
                  Begin
                  Result:=ieWarningPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1022) ;
                  End ;
               End ; // 4 : ...
           End ;//Case
      End else//If LireResultat(....
      if not IsPrinter then
         Begin  // Si l'imprimante est serial et on n'a rien reçu, possiblement l'imprimante est eteinte
         Result:=iePowerOff ;
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1023) ; Exit ;
         End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM210.IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
var tt : String ;
Begin
Result:=Etat ;
if Result=iePowerOff then
   Begin
   Result:=ieRecovery ;
   End else
   Begin
   Err.Libelle:=MC_MsgErrDefaut(1025) ;
   if ((Result in [ieError,ieErrorPaper]) and
       (LanceMessage(Err,TRUE)=mryes)) then Begin FlastError:=ieErrorPaper ; Result:=LastError ; End else
   if (Result in [iewarning,ieRecovery,ieError,ieErrorPaper])  then //ieError et ieErrorPaper si on a respondu NON à la re-impresion
      Begin
      tt:='' ;
      if Result in [ieWarning,ieRecovery] then tt:=Sqce[spRestart,TRUE]+#2 else
         if Result=ieErrorPaper then tt:=Sqce[spRestart,TRUE]+#0 ;
      if tt<>'' then
         begin
         tt:=Sqce[spActivate,TRUE]+tt+Sqce[spDesactivate,TRUE] ;
         FSortie.SendString(tt) ;
         Result:=ieRecovery ;
         End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM210.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;    
  if i>0 then Result[i]:=#213 ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLP_EPSONTM675.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
OnTesteEtat:=TestePrinter ;
OnInitialise:=InitPrinter ;
OnErreur:=IlyaErreur ;
Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM675.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
LongMaxBuffer:=200 ;
LeDelayLong:=100 ;
LeDelayCourt:=LeDelayLong ; 

Result:=Inherited ChargePortEtParams(APort,Aparams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EPSONTM675.ChargeSequences ;
Begin
Inherited ; 
FSequenceLP[spCR]                    .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]                    .Sqce[TRUE] :=#10 ;

FSequenceLP[spInitialise]            .Sqce[TRUE] :=#27+'@'+#27+'3'+#16+#27+'R'+#1+#27+'t'+#19+#27+'f'+#0+#5 ;

FSequenceLP[spSautPage]              .Sqce[TRUE] :='' ; //On n'a pas de saut de page en rouleau

FSequenceLP[spStatus]                .Sqce[TRUE] :=#16+#4 ; //1,2,3,4,5,6

if Port in [pnLPT1,pnLPT2] then
   FSequenceLP[spSensorPaperOut]     .Sqce[TRUE]:=#27+'c3'+#175 ; //Seulement si interface parallel

FSequenceLP[spReStart]               .Sqce[TRUE] :=#16+#5 ; //0,2

FSequenceLP[spActivateInterligne]    .Sqce[TRUE] :=#27+'3'+#24 ;
FSequenceLP[spDesactivateInterligne] .Sqce[TRUE] :=#27+'3'+#16 ;    

FSequenceLP[spTotalCut]              .Sqce[TRUE] :=FSequenceLP[spDesactivateInterligne].Sqce[TRUE]+#27+'d'+#13 ;

FSequenceLP[spPartialCut]            .Sqce[TRUE] :=FSequenceLP[spDesactivateInterligne].Sqce[TRUE]+#27+'d'+#13 ; 

FSequenceLP[spActivate]              .Sqce[TRUE] :=#27+'='+#1 ;

FSequenceLP[spDesactivate]           .Sqce[TRUE] :='' ; //#27+'='+#2 ;

FSequenceLP[spActivateSlip]          .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'<'+#27+'c0'+#4 ;

FSequenceLP[spDesactivateSlip]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'c0'+#1 ;

FSequenceLP[spActivateValidation]    .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'<'+#27+'c0'+#8 ;

FSequenceLP[spDesactivateValidation] .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'c0'+#1 ;

FSequenceLP[spDebutCheque]           .Sqce[TRUE] :=FSequenceLP[spActivateSlip].Sqce[TRUE]+#29+'P'+#10+#10+#27+'L'+#27+'W'+#3+#0+#0+#0+'^'+#0+':'+#0+#27+'T1' ;

FSequenceLP[spFinCheque]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#12+FSequenceLP[spDesactivateSlip].Sqce[TRUE] ;

FSequenceLP[sp17Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp17Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp17Cpi]                 .Groupe:=1 ;
FSequenceLP[sp17Cpi]                 .Valeur:=32 ;

FSequenceLP[sp12Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp12Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp12Cpi]                 .Groupe:=1 ;
FSequenceLP[sp12Cpi]                 .Valeur:=1 ;

FSequenceLP[sp10Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp10Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp10Cpi]                 .Groupe:=1 ;
FSequenceLP[sp10Cpi]                 .Valeur:=0 ;

FSequenceLP[sp05Cpi]                  .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp05Cpi]                  .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp05Cpi]                  .Groupe:=1 ;
FSequenceLP[sp05Cpi]                  .Valeur:=48 ;

FSequenceLP[spUnderLine]             .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spUnderLine]             .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[spUnderline]             .Groupe:=1 ;
FSequenceLP[spUnderline]             .Valeur:=128 ;

FSequenceLP[spBold]                  .Sqce[FALSE]:=#27+'!%s' ;  //#27+'G'+#0 ;
FSequenceLP[spBold]                  .Sqce[TRUE] :=#27+'!%s' ; //=#27+'G'+#1 ;
FSequenceLP[spBold]                  .Groupe:=1 ;
FSequenceLP[spBold]                  .Valeur:=8 ;

FSequenceLP[spItalic]                .Sqce[FALSE]:=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#0 ;
FSequenceLP[spItalic]                .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#1 ;
FSequenceLP[spItalic]                .Groupe:=0 ;
FSequenceLP[spItalic]                .Valeur:=0 ;

FSequenceLP[spBitMap]                .Sqce[FALSE]:='' ;
FSequenceLP[spBitMap]                .Sqce[TRUE] :=#27+'*'+#1+'%s'+#0 ;
FSequenceLP[spBitMap]                .Groupe:=0 ;
FSequenceLP[spBitMap]                .Valeur:=0 ;

FSequenceLP[spPrintBmp]              .Sqce[TRUE] :=#28+'p'+#1+'0' ;

FSequenceLP[spCodeBarre]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]
                                                  +#29+'h%s'   //Hauteur cdu codebarre (Nombre points * lignes)
                                                  +#29+'H%s'   //Position u HRI
                                                  +#29+'f'+#1  //%s' //Font à utiliser //XMG 23/06/03
                                                  +#29+'w'+#1  //%s' //Longeur du module....
                                                  +#29+'k%s'  //Code du codebarre plus valeur du CodeBarre à imprimer
                                                  +FSequenceLP[spLF].Sqce[TRUE] ;
FSequenceLP[spCodeBarre]             .Extra:='9;' ;


End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM675.InitPrinter ( var err : TMC_Err ) : Boolean ;
Var St : String ;
Begin
Result:=FALSE ;
St:=Sqce[spInitialise,TRUE] ; if St<>'' then Result:=WriteLP(St,TRUE) ;
if Result then
   Begin
   St:=Sqce[spSensorPaperOut,TRUE] ; if St<>'' then Result:=writeLP(St,TRUE) ;
   End ;
if not result then err:=LasTMC_Error ;  
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLP_EPSONTM675.TestePrinter ( Fois : Integer ; Var Err : TMC_Err ) : TImprimErreurs ;
Var tt   : String ;
    b    : Byte ;
    R    : String ;
    V    : Variant ;
Begin
Result:=ieOk ;
if Fois=1 then err.code:=0 ; err.Libelle:='' ;
tt:=Sqce[spActivate,TRUE]+Sqce[spstatus,TRUE]+chr(Fois)+Sqce[spDesactivate,TRUE] ;
if FSortie.SendString(tt) then
   Begin
   Attentresponse ;
   if lireResultat(1,VarString,V) then
      Begin
      R:=VString(V) ;
      if Length(R)=1 then b:=ord(R[1]) else B:=18 ;
      if (B<>(18+4*ord(Fois=1))) then
         Case Fois of
           1 : Begin //Status de l'imprimante
            {   If (B and 32)=32 then //L'imprimante essaie de se recuperer
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1015) ;
                  End else  //Ce type d'erreur n'existe plus en TM-675}
               if (B and 8)=8 then //L'imprimante est Off-Line
                  Begin
                  Result:=TestePrinter(2,Err) ;
                  if Result=ieOk then //Si aucun erreur est rapporté au deuxième niveau alors c'est OFF-Line
                     Begin
                     Result:=ieOffLine ;
                     Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1016) ;
                     End ;
                  End ;
               End ; //3: ....
           2 : Begin //Status de l'imprimante quand elle est en Off-Line
               if (B and 64)=64 then //Error
                  Result:=TestePrinter(3,Err) else
               if (B and 32)=32 then //End of paper
                  Result:=TestePrinter(4,Err) else
               if (B and 4)=4 then //Cover Open
                  Begin
                  Result:=ieCover ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1017) ; Exit ;
                  End ;
               End ; //3: ....
           3 : Begin //Cause de l'erreur de l'imprimante         
               If (B and 32)=32 then //
                  Begin
                  Result:=ieError ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1018) ;
                  End else
               If (B and 64)=64 then //
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1019) ; 
                  End else
               if ((B and 4)=4) then //or ((B and 8)=8) then  il n'existe plus la valeur 8 pour la TM-675
                  Begin
                  Result:=ieWarning ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1020) ; 
                  End ;
               End ; //3: ....
           4 : Begin //Causse de la fin du papier
               If (B and 96)=96 then
                  Begin
                  Result:=ieErrorPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1021) ; 
                  End else
               If (B and 12)=12 then
                  Begin
                  Result:=ieWarningPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1022) ;
                  End ;
               End ; // 4 : ...
           End ;//Case
      End else//If LireResultat(....
      if not IsPrinter then
         Begin  // Si l'imprimante est serial et on n'a rien reçu, possiblement l'imprimante est eteinte
         Result:=iePowerOff ;
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1023) ; Exit ;
         End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM675.IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
var tt : String ;
Begin
Result:=Etat ;
if Result=iePowerOff then
   Begin
   Result:=ieRecovery ;
   End else
   Begin
   Err.Libelle:=MC_MsgErrDefaut(1025) ;
   if ((Result in [ieError,ieErrorPaper]) and
       (LanceMessage(Err,TRUE)=mryes)) then Begin FlastError:=ieErrorPaper ; Result:=LastError ; End else
   if (Result in [iewarning,ieRecovery,ieError,ieErrorPaper])  then //ieError et ieErrorPaper si on a respondu NON à la re-impresion
      Begin
      tt:='' ;
      if Result in [ieWarning,ieRecovery] then tt:=Sqce[spRestart,TRUE]+#1 else
         if Result=ieErrorPaper then tt:=Sqce[spRestart,TRUE]+#3 ;
      if tt<>'' then
         begin
         tt:=Sqce[spActivate,TRUE]+tt+Sqce[spDesactivate,TRUE] ;
         FSortie.SendString(tt) ;
         Result:=ieRecovery ;
         End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM675.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;     
  if i>0 then Result[i]:=#213 ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLP_EPSONTM950.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
OnTesteEtat:=TestePrinter ;
OnInitialise:=InitPrinter ;
OnErreur:=IlyaErreur ;
Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM950.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
LongMaxBuffer:=60 ;
LeDelayLong:=250 ;
LeDelayCourt:=270 ;

Result:=Inherited ChargePortEtParams(APort,Aparams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EPSONTM950.ChargeSequences ;
Begin
//// activation journal esc z #1   désactivation journal esc z #2      sélection receipt+journal esc c 0 #3
Inherited ; 
FSequenceLP[spCR]                    .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]                    .Sqce[TRUE] :=#10 ;

FSequenceLP[spInitialise]            .Sqce[TRUE] :=#27+'@'+#27+'3'+#16+#27+'<'+#27+'z'+#1+#27+'R'+#1+#27+'t'+#0+#27+'f'+#0+#5 ;   

FSequenceLP[spSautPage]              .Sqce[TRUE] :='' ; //On n'a pas de saut de page en rouleau

FSequenceLP[spStatus]                .Sqce[TRUE] :=#16+#4 ; //1,2,3,4,5,6  ?? 1<=n<=5 ??

if Port in [pnLPT1,pnLPT2] then
   FSequenceLP[spSensorPaperOut]     .Sqce[TRUE]:=#27+'c3'+#175 ; //Seulement si interface parallel

FSequenceLP[spReStart]               .Sqce[TRUE] :=#16+#5 ; //0,2    ?? 1<=n<=3 ??

FSequenceLP[spActivateInterligne]    .Sqce[TRUE] :=#27+'<'+#27+'3'+#24 ;
FSequenceLP[spDesactivateInterligne] .Sqce[TRUE] :=#27+'<'+#27+'3'+#16 ;

FSequenceLP[spTotalCut]              .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+FSequenceLP[spDesactivateInterligne].Sqce[TRUE]+#27+'d'+#14+#27+'i' ;

FSequenceLP[spPartialCut]            .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+FSequenceLP[spDesactivateInterligne].Sqce[TRUE]+#27+'d'+#14+#27+'m' ;    

FSequenceLP[spActivate]              .Sqce[TRUE] :=#27+'='+#1 ;

FSequenceLP[spDesactivate]           .Sqce[TRUE] :='' ; //#27+'='+#2 ;

FSequenceLP[spActivateSlip]          .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'<'+#27+'c0'+#4 ;

FSequenceLP[spDesactivateSlip]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'<'+#27+'c0'+#3 ;    

FSequenceLP[spDebutCheque]           .Sqce[TRUE] :=FSequenceLP[spActivateSlip].Sqce[TRUE] ;

FSequenceLP[spFinCheque]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#12+FSequenceLP[spDesactivateSlip].Sqce[TRUE] ;

FSequenceLP[sp17Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp17Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp17Cpi]                 .Groupe:=1 ;
FSequenceLP[sp17Cpi]                 .Valeur:=32 ;

FSequenceLP[sp12Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp12Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp12Cpi]                 .Groupe:=1 ;
FSequenceLP[sp12Cpi]                 .Valeur:=1 ;

FSequenceLP[sp10Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp10Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp10Cpi]                 .Groupe:=1 ;
FSequenceLP[sp10Cpi]                 .Valeur:=0 ;

FSequenceLP[sp05Cpi]                  .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp05Cpi]                  .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp05Cpi]                  .Groupe:=1 ;
FSequenceLP[sp05Cpi]                  .Valeur:=48 ;

FSequenceLP[spUnderLine]             .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spUnderLine]             .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[spUnderline]             .Groupe:=1 ;
FSequenceLP[spUnderline]             .Valeur:=128 ;

FSequenceLP[spBold]           .Sqce[FALSE]:=#27+'!%s' ;  //#27+'G'+#0 ;
FSequenceLP[spBold]           .Sqce[TRUE] :=#27+'!%s' ; //=#27+'G'+#1 ;
FSequenceLP[spBold]           .Groupe:=1 ;
FSequenceLP[spBold]           .Valeur:=8 ;

FSequenceLP[spItalic]                .Sqce[FALSE]:=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#0 ;
FSequenceLP[spItalic]                .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#1 ;
FSequenceLP[spItalic]                .Groupe:=0 ;
FSequenceLP[spItalic]                .Valeur:=0 ;

FSequenceLP[spBitMap]                .Sqce[FALSE]:='' ;
FSequenceLP[spBitMap]                .Sqce[TRUE] :=#27+'*'+#1+'%s'+#0 ;
FSequenceLP[spBitMap]                .Groupe:=0 ;
FSequenceLP[spBitMap]                .Valeur:=0 ;

FSequenceLP[spPrintBmp]              .Sqce[TRUE] :='' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM950.InitPrinter ( var err : TMC_Err ) : Boolean ;
Var St : String ;
Begin
Result:=FALSE ;
St:=Sqce[spInitialise,TRUE] ; if St<>'' then Result:=WriteLP(St,TRUE) ;
if Result then
   Begin
   St:=Sqce[spSensorPaperOut,TRUE] ; if St<>'' then Result:=writeLP(St,TRUE) ;
   End ;
if not result then err:=LasTMC_Error ;  
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLP_EPSONTM950.TestePrinter ( Fois : Integer ; Var Err : TMC_Err ) : TImprimErreurs ;
Var tt   : String ;
    b    : Byte ;
    R    : String ;
    V    : Variant ;
Begin
Result:=ieOk ;
if Fois=1 then err.code:=0 ; err.Libelle:='' ;
tt:=Sqce[spActivate,TRUE]+Sqce[spstatus,TRUE]+chr(Fois)+Sqce[spDesactivate,TRUE] ;
if FSortie.SendString(tt) then
   Begin
   Attentresponse ;
   if lireResultat(1,VarString,V) then
      Begin
      R:=VString(V) ;
      if Length(R)=1 then b:=ord(R[1]) else B:=18 ;
      if (B<>(18+4*ord(Fois=1))) then
         Case Fois of
           1 : Begin //Status de l'imprimante
            {   If (B and 32)=32 then //L'imprimante essaie de se recuperer
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1015) ;
                  End else  //Ce type d'erreur n'existe plus en TM-675}
               if (B and 8)=8 then //L'imprimante est Off-Line
                  Begin
                  Result:=TestePrinter(2,Err) ;
                  if Result=ieOk then //Si aucun erreur est rapporté au deuxième niveau alors c'est OFF-Line
                     Begin
                     Result:=ieOffLine ;
                     Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1016) ;
                     End ;
                  End ;
               End ; //3: ....
           2 : Begin //Status de l'imprimante quand elle est en Off-Line
               if (B and 64)=64 then //Error
                  Result:=TestePrinter(3,Err) else
               if (B and 32)=32 then //End of paper
                  Result:=TestePrinter(4,Err) else
               if (B and 4)=4 then //Cover Open
                  Begin
                  Result:=ieCover ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1017) ; Exit ;
                  End ;
               End ; //3: ....
           3 : Begin //Cause de l'erreur de l'imprimante         
               If (B and 32)=32 then //
                  Begin
                  Result:=ieError ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1018) ;
                  End else
               If (B and 64)=64 then //
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1019) ;
                  End else
               if ((B and 4)=4) then //or ((B and 8)=8) then  il n'existe plus la valeur 8 pour la TM-675
                  Begin
                  Result:=ieWarning ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1020) ; 
                  End ;
               End ; //3: ....
           4 : Begin //Causse de la fin du papier
               If (B and 96)=96 then
                  Begin
                  Result:=ieErrorPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1021) ; 
                  End else
               If (B and 12)=12 then
                  Begin
                  Result:=ieWarningPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1022) ;
                  End ;
               End ; // 4 : ...
           End ;//Case
      End else//If LireResultat(....
      if not IsPrinter then
         Begin  // Si l'imprimante est serial et on n'a rien reçu, possiblement l'imprimante est eteinte
         Result:=iePowerOff ;
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1023) ; Exit ;
         End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM950.IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
var tt : String ;
Begin
Result:=Etat ;
if Result=iePowerOff then
   Begin
   //if InitImprimante(Err) then
   Result:=ieRecovery ;
   End else
   Begin
   Err.Libelle:=MC_MsgErrDefaut(1025) ;
   if ((Result in [ieError,ieErrorPaper]) and
       (LanceMessage(Err,TRUE)=mryes)) then Begin FlastError:=ieErrorPaper ; Result:=LastError ; End else
   if (Result in [iewarning,ieRecovery,ieError,ieErrorPaper])  then //ieError et ieErrorPaper si on a respondu NON à la re-impresion
      Begin
      tt:='' ;
      if Result in [ieWarning,ieRecovery] then tt:=Sqce[spRestart,TRUE]+#1 else
         if Result=ieErrorPaper then tt:=Sqce[spRestart,TRUE]+#3 ;
      if tt<>'' then
         begin
         tt:=Sqce[spActivate,TRUE]+tt+Sqce[spDesactivate,TRUE] ;
         FSortie.SendString(tt) ;
         Result:=ieRecovery ;
         End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM950.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;
  if i>0 then Result[i]:=#238 ;                       
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLP_EPSONTM5000.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
OnTesteEtat:=TestePrinter ;
OnInitialise:=InitPrinter ;
OnErreur:=IlyaErreur ;
Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM5000.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
LongMaxBuffer:=800 ;
LeDelayLong:=12 ;
LeDelayCourt:=0 ;

Result:=Inherited ChargePortEtParams(APort,Aparams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EPSONTM5000.ChargeSequences ;
Begin
//// activation journal esc z #1   désactivation journal esc z #2      sélection receipt+journal esc c 0 #3
Inherited ; 
FSequenceLP[spCR]                    .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]                    .Sqce[TRUE] :=#10 ;

FSequenceLP[spInitialise]            .Sqce[TRUE] :=#27+'@'+#27+'3'+#16+#27+'R'+#1+#27+'t'+#0+#27+'f'+#0+#5 ;

FSequenceLP[spSautPage]              .Sqce[TRUE] :='' ; //On n'a pas de saut de page en rouleau

FSequenceLP[spStatus]                .Sqce[TRUE] :=#16+#4 ; //1,2,3,4,5,6  ?? 1<=n<=5 ??

if Port in [pnLPT1,pnLPT2] then
   FSequenceLP[spSensorPaperOut]     .Sqce[TRUE]:=#27+'c3'+#175 ; //Seulement si interface parallel

FSequenceLP[spReStart]               .Sqce[TRUE] :=#16+#5 ; //0,2    ?? 1<=n<=3 ??

FSequenceLP[spActivateInterligne]    .Sqce[TRUE] :=#27+'3'+#24 ;
FSequenceLP[spDesactivateInterligne] .Sqce[TRUE] :=#27+'3'+#16 ;   

FSequenceLP[spTotalCut]              .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VA'+#1 ;

FSequenceLP[spPartialCut]            .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VB'+#1 ;

FSequenceLP[spActivate]              .Sqce[TRUE] :=#27+'='+#1 ;

FSequenceLP[spDesactivate]           .Sqce[TRUE] :='' ; //#27+'='+#2 ;

FSequenceLP[spActivateSlip]          .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'<'+#27+'c0'+#4 ;

FSequenceLP[spDesactivateSlip]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'c0'+#3 ;

FSequenceLP[spDebutCheque]           .Sqce[TRUE] :=FSequenceLP[spActivateSlip].Sqce[TRUE] ;

FSequenceLP[spFinCheque]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#12+FSequenceLP[spDesactivateSlip].Sqce[TRUE] ;

FSequenceLP[sp17Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp17Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp17Cpi]                 .Groupe:=1 ;
FSequenceLP[sp17Cpi]                 .Valeur:=32 ;

FSequenceLP[sp12Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp12Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp12Cpi]                 .Groupe:=1 ;
FSequenceLP[sp12Cpi]                 .Valeur:=1 ;

FSequenceLP[sp10Cpi]                 .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp10Cpi]                 .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp10Cpi]                 .Groupe:=1 ;
FSequenceLP[sp10Cpi]                 .Valeur:=0 ;

FSequenceLP[sp05Cpi]                  .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp05Cpi]                  .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp05Cpi]                  .Groupe:=1 ;
FSequenceLP[sp05Cpi]                  .Valeur:=48 ;

FSequenceLP[spUnderLine]             .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spUnderLine]             .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[spUnderline]             .Groupe:=1 ;
FSequenceLP[spUnderline]             .Valeur:=128 ;

FSequenceLP[spBold]                  .Sqce[FALSE]:=#27+'!%s' ;  //#27+'G'+#0 ;
FSequenceLP[spBold]                  .Sqce[TRUE] :=#27+'!%s' ; //=#27+'G'+#1 ;
FSequenceLP[spBold]                  .Groupe:=1 ;
FSequenceLP[spBold]                  .Valeur:=8 ;

FSequenceLP[spItalic]                .Sqce[FALSE]:=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#0 ;
FSequenceLP[spItalic]                .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#1 ;
FSequenceLP[spItalic]                .Groupe:=0 ;
FSequenceLP[spItalic]                .Valeur:=0 ;

FSequenceLP[spBitMap]                .Sqce[FALSE]:='' ;
FSequenceLP[spBitMap]                .Sqce[TRUE] :=#27+'*'+#1+'%s'+#0 ;
FSequenceLP[spBitMap]                .Groupe:=0 ;
FSequenceLP[spBitMap]                .Valeur:=0 ;

FSequenceLP[spCodeBarre]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]
                                                  +#29+'h%s'   //Hauteur cdu codebarre (Nombre points * lignes)
                                                  +#29+'H%s'   //Position u HRI
                                                  +#29+'f'+#1  //%s' //Font à utiliser //XMG 23/06/03
                                                  +#29+'w'+#1  //%s' //Longeur du module....
                                                  +#29+'k%s'  //Code du codebarre plus valeur du CodeBarre à imprimer
                                                  +FSequenceLP[spLF].Sqce[TRUE] ;
FSequenceLP[spCodeBarre]             .Extra:='9;' ; //XMG 11/06/03

FSequenceLP[spPrintBmp]              .Sqce[TRUE] :=#28+'p'+#1+'0' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM5000.InitPrinter ( var err : TMC_Err ) : Boolean ;
Var St : String ;
Begin
Result:=FALSE ;
St:=Sqce[spInitialise,TRUE] ; if St<>'' then Result:=WriteLP(St,TRUE) ;
if Result then
   Begin
   St:=Sqce[spSensorPaperOut,TRUE] ; if St<>'' then Result:=writeLP(St,TRUE) ;
   End ;
if not result then err:=LasTMC_Error ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLP_EPSONTM5000.TestePrinter ( Fois : Integer ; Var Err : TMC_Err ) : TImprimErreurs ;
Var tt   : String ;
    b    : Byte ;
    R    : String ;
    V    : Variant ;
Begin
Result:=ieOk ;
if Fois=1 then err.code:=0 ; err.Libelle:='' ;
tt:=Sqce[spActivate,TRUE]+Sqce[spstatus,TRUE]+chr(Fois)+Sqce[spDesactivate,TRUE] ;
if FSortie.SendString(tt) then
   Begin
   Attentresponse ;
   if lireResultat(1,VarString,V) then
      Begin
      R:=VString(V) ;
      if Length(R)=1 then b:=ord(R[1]) else B:=18 ;
      if (B<>(18+4*ord(Fois=1))) then
         Case Fois of
           1 : Begin //Status de l'imprimante
            {   If (B and 32)=32 then //L'imprimante essaie de se recuperer
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1015) ;
                  End else  //Ce type d'erreur n'existe plus en TM-675}
               if (B and 8)=8 then //L'imprimante est Off-Line
                  Begin
                  Result:=TestePrinter(2,Err) ;
                  if Result=ieOk then //Si aucun erreur est rapporté au deuxième niveau alors c'est OFF-Line
                     Begin
                     Result:=ieOffLine ;
                     Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1016) ;
                     End ;
                  End ;
               End ; //3: ....
           2 : Begin //Status de l'imprimante quand elle est en Off-Line
               if (B and 64)=64 then //Error
                  Result:=TestePrinter(3,Err) else
               if (B and 32)=32 then //End of paper
                  Result:=TestePrinter(4,Err) else
               if (B and 4)=4 then //Cover Open
                  Begin
                  Result:=ieCover ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1017) ; Exit ;
                  End ;
               End ; //3: ....
           3 : Begin //Cause de l'erreur de l'imprimante
               If (B and 32)=32 then //
                  Begin
                  Result:=ieError ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1018) ;
                  End else
               If (B and 64)=64 then //
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1019) ; 
                  End else
               if ((B and 4)=4) then //or ((B and 8)=8) then  il n'existe plus la valeur 8 pour la TM-675
                  Begin
                  Result:=ieWarning ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1020) ; 
                  End ;
               End ; //3: ....
           4 : Begin //Causse de la fin du papier
               If (B and 96)=96 then
                  Begin
                  Result:=ieErrorPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1021) ; 
                  End else
               If (B and 12)=12 then
                  Begin
                  Result:=ieWarningPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1022) ;
                  End ;
               End ; // 4 : ...
           End ;//Case
      End else//If LireResultat(....
      if not IsPrinter then
         Begin  // Si l'imprimante est serial et on n'a rien reçu, possiblement l'imprimante est eteinte
         Result:=iePowerOff ;
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1023) ; Exit ;
         End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM5000.IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
var tt : String ;
Begin
Result:=Etat ;
if Result=iePowerOff then
   Begin
   Result:=ieRecovery ;
   End else
   Begin
   Err.Libelle:=MC_MsgErrDefaut(1025) ;
   if ((Result in [ieError,ieErrorPaper]) and
       (LanceMessage(Err,TRUE)=mryes)) then Begin FlastError:=ieErrorPaper ; Result:=LastError ; End else
   if (Result in [iewarning,ieRecovery,ieError,ieErrorPaper])  then //ieError et ieErrorPaper si on a respondu NON à la re-impresion
      Begin
      tt:='' ;
      if Result in [ieWarning,ieRecovery] then tt:=Sqce[spRestart,TRUE]+#1 else
         if Result=ieErrorPaper then tt:=Sqce[spRestart,TRUE]+#3 ;
      if tt<>'' then
         begin
         tt:=Sqce[spActivate,TRUE]+tt+Sqce[spDesactivate,TRUE] ;
         FSortie.SendString(tt) ;
         Result:=ieRecovery ;
         End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM5000.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;
  if i>0 then Result[i]:=#238 ;                        
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLP_EPSONTM88II.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
OnTesteEtat:=TestePrinter ;
OnInitialise:=InitPrinter ;
OnErreur:=IlyaErreur ;
Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM88II.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
LongMaxBuffer:=800 ;
LeDelayLong:=12 ;
LeDelayCourt:=0 ;

Result:=Inherited ChargePortEtParams(APort,AParams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EPSONTM88II.ChargeSequences ;
Begin
Inherited ;
FSequenceLP[spCR]             .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]             .Sqce[TRUE] :=#10 ;

FSequenceLP[spInitialise]     .Sqce[TRUE] :=#27+'@'+#27+'3'+#16+#27+'R'+#1+#27+'t'+#19 ;

FSequenceLP[spSautPage]       .Sqce[TRUE] :='' ; //On n'a pas de saut de page en rouleau

FSequenceLP[spStatus]         .Sqce[TRUE] :=#16+#4 ; //1,2,3,4

if Port in [pnLPT1,pnLPT2] then FSequenceLP[spSensorPaperOut].Sqce[TRUE]:=#27+'c3'+#15 ; //Seulement si interface parallel

FSequenceLP[spReStart]        .Sqce[TRUE] :=#16+#5 ; //0,2

FSequenceLP[spActivateInterligne]    .Sqce[TRUE] :=#27+'3'+#24 ;
FSequenceLP[spDesactivateInterligne] .Sqce[TRUE] :=#27+'3'+#16 ;

FSequenceLP[spTotalCut]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VA'+#1 ;

FSequenceLP[spPartialCut]     .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VB'+#1 ;

FSequenceLP[spActivate]       .Sqce[TRUE] :=#27+'='+#1 ;

FSequenceLP[spDesactivate]    .Sqce[TRUE] :=#27+'='+#2 ;

FSequenceLP[sp17Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Groupe:=1 ;
FSequenceLP[sp17Cpi]          .Valeur:=32 ;

FSequenceLP[sp12Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Groupe:=1 ;
FSequenceLP[sp12Cpi]          .Valeur:=1 ;

FSequenceLP[sp10Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Groupe:=1 ;
FSequenceLP[sp10Cpi]          .Valeur:=0 ;

FSequenceLP[sp05Cpi]           .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Groupe:=1 ;
FSequenceLP[sp05Cpi]           .Valeur:=48 ;

FSequenceLP[spUnderLine]      .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spUnderLine]      .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[spUnderline]      .Groupe:=1 ;
FSequenceLP[spUnderline]      .Valeur:=128 ;

FSequenceLP[spBold]           .Sqce[FALSE]:=#27+'!%s' ;  //#27+'G'+#0 ;
FSequenceLP[spBold]           .Sqce[TRUE] :=#27+'!%s' ; //=#27+'G'+#1 ;
FSequenceLP[spBold]           .Groupe:=1 ;
FSequenceLP[spBold]           .Valeur:=8 ;

FSequenceLP[spItalic]         .Sqce[FALSE]:='' ;  //non supported
FSequenceLP[spItalic]         .Sqce[TRUE] :='' ;
FSequenceLP[spItalic]         .Groupe:=0 ;
FSequenceLP[spItalic]         .Valeur:=0 ;

FSequenceLP[spBitMap]         .Sqce[FALSE]:='' ;
FSequenceLP[spBitMap]         .Sqce[TRUE] :=#27+'*'+#1+'%s'+#0 ;
FSequenceLP[spBitMap]         .Groupe:=0 ;
FSequenceLP[spBitMap]         .Valeur:=0 ;

FSequenceLP[spCodeBarre]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]
                                                  +#29+'h%s'   //Hauteur cdu codebarre (Nombre points * lignes)
                                                  +#29+'H%s'   //Position u HRI
                                                  +#29+'f'+#1  //%s' //Font à utiliser //XMG 23/06/03
                                                  +#29+'w'+#1  //%s' //Longeur du module....
                                                  +#29+'k%s'  //Code du codebarre plus valeur du CodeBarre à imprimer
                                                  +FSequenceLP[spLF].Sqce[TRUE] ;
FSequenceLP[spCodeBarre]             .Extra:='9;' ; //XMG 11/06/03

FSequenceLP[spPrintBmp]              .Sqce[TRUE] :=#28+'p'+#1+'0' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM88II.InitPrinter ( var err : TMC_Err ) : Boolean ;  
Var St : String ;
Begin
Result:=FALSE ;
St:=Sqce[spInitialise,TRUE] ; if St<>'' then Result:=WriteLP(St,TRUE) ;
if Result then
   Begin
   St:=Sqce[spSensorPaperOut,TRUE] ; if St<>'' then Result:=writeLP(St,TRUE) ;
   End ;
if not result then err:=LasTMC_Error ;  
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLP_EPSONTM88II.TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
Var tt   : String ;
    b    : Byte ;
    R    : String ;
    V    : Variant ;
Begin
Result:=ieOk ;
if Fois=1 then err.code:=0 ; err.Libelle:='' ;
tt:=Sqce[spActivate,TRUE]+Sqce[spstatus,TRUE]+chr(Fois)+Sqce[spDesactivate,TRUE] ;
if FSortie.SendString(tt) then
   Begin
   Attentresponse ;
   if lireResultat(1,VarString,V) then
      Begin
      R:=VString(V) ;
      if Length(R)=1 then b:=ord(R[1]) else B:=18 ;
      if (B<>(18+4*ord(Fois=1))) then
         Case Fois of
           1 : Begin //Status de l'imprimante
               If (B and 32)=32 then //L'imprimante essaie de se recuperer
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1015) ;
                  End else
               if (B and 8)=8 then //L'imprimante est Off-Line
                  Begin
                  Result:=TestePrinter(2,Err) ;
                  if Result=ieOk then //Si aucun erreur est rapporté au deuxième niveau alors c'est OFF-Line
                     Begin
                     Result:=ieOffLine ;
                     Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1016) ;
                     End ;
                  End ;
               End ; //3: ....
           2 : Begin //Status de l'imprimante quand elle est en Off-Line
               if (B and 64)=64 then //Error
                  Result:=TestePrinter(3,Err) else
               if (B and 32)=32 then //End of paper
                  Result:=TestePrinter(4,Err) {else
               if (B and 4)=4 then //Cover Open
                  Begin
                  Result:=ieCover ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1017) ; Exit ;
                  End ;  XMG Il n'est pas sûr que ce bit soit bien initialise (Voir documentation) }
               End ; //3: ....
           3 : Begin //Cause de l'erreur de l'imprimante
               If (B and 32)=32 then //
                  Begin
                  Result:=ieError ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1018) ;
                  End else
               If (B and 64)=64 then //
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1019) ;
                  End else
               if ((B and 4)=4) or ((B and 8)=8) then
                  Begin
                  Result:=ieWarning ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1020) ;
                  End ;
               End ; //3: ....
           4 : Begin //Causse de la fin du papier
               If (B and 96)=96 then
                  Begin
                  Result:=ieErrorPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1021) ;
                  End else
               If (B and 12)=12 then
                  Begin
                  Result:=ieWarningPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1022) ;
                  End ;
               End ; // 4 : ...
           End ;//Case
      End else//If LireResultat(....
      if not IsPrinter then
         Begin  // Si l'imprimante est serial et on n'a rien reçu, possiblement l'imprimante est eteinte
         Result:=iePowerOff ;
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1023) ; Exit ;
         End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM88II.IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
var tt : String ;
Begin
Result:=Etat ;
if Result=iePowerOff then
   Begin
   Result:=ieRecovery ;
   End else
   Begin
   Err.Libelle:=MC_MsgErrDefaut(1025) ;
   if ((Result in [ieError,ieErrorPaper]) and
       (LanceMessage(Err,TRUE)=mryes)) then Begin FlastError:=ieErrorPaper ; Result:=LastError ; End else
   if (Result in [iewarning,ieRecovery,ieError,ieErrorPaper])  then //ieError et ieErrorPaper si on a respondu NON à la re-impresion
      Begin
      tt:='' ;
      if Result in [ieWarning,ieRecovery] then tt:=Sqce[spRestart,TRUE]+#2 else
         if Result=ieErrorPaper then tt:=Sqce[spRestart,TRUE]+#0 ;
      if tt<>'' then
         begin
         tt:=Sqce[spActivate,TRUE]+tt+Sqce[spDesactivate,TRUE] ;
         FSortie.SendString(tt) ;
         Result:=ieRecovery ;
         End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM88II.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;     
  if i>0 then Result[i]:=#213 ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLP_EPSONTM300.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
OnTesteEtat:=TestePrinter ;
OnInitialise:=InitPrinter ;
OnErreur:=IlyaErreur ;
Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM300.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
LongMaxBuffer:=40 ;
LeDelayLong:=1200 ;
LeDelayCourt:=215 ;

Result:=Inherited ChargePortEtParams(APort,AParams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EPSONTM300.ChargeSequences ;
Begin
Inherited ; 
FSequenceLP[spCR]             .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]             .Sqce[TRUE] :=#10 ;

FSequenceLP[spInitialise]     .Sqce[TRUE] :=#27+'@'+#27+'3'+#16+#27+'<'+#27+'z'+#1+#27+'R'+#1+#27+'t'+#0+#27+'f'+#0+#5 ;

FSequenceLP[spSautPage]       .Sqce[TRUE] :='' ; //On n'a pas de saut de page en rouleau

FSequenceLP[spStatus]         .Sqce[TRUE] :=#27+'v' ; //1,2,3,4

if Port in [pnLPT1,pnLPT2] then FSequenceLP[spSensorPaperOut].Sqce[TRUE]:=#27+'c3'+#15 ; //Seulement si interface parallel

FSequenceLP[spReStart]        .Sqce[TRUE] :=#16+#5 ; //0,2

FSequenceLP[spActivateInterligne]    .Sqce[TRUE] :=#27+'3'+#24 ;
FSequenceLP[spDesactivateInterligne] .Sqce[TRUE] :=#27+'3'+#16 ;  

FSequenceLP[spTotalCut]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+FSequenceLP[spDesactivateInterligne].Sqce[TRUE]+#27+'d'+#13+#27+'i' ;

FSequenceLP[spPartialCut]     .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+FSequenceLP[spDesactivateInterligne].Sqce[TRUE]+#27+'d'+#13+#27+'m' ;  

FSequenceLP[spActivate]       .Sqce[TRUE] :=#27+'='+#1 ;

FSequenceLP[spDesactivate]    .Sqce[TRUE] :=#27+'='+#2 ;

FSequenceLP[sp17Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Groupe:=1 ;
FSequenceLP[sp17Cpi]          .Valeur:=32 ;

FSequenceLP[sp12Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Groupe:=1 ;
FSequenceLP[sp12Cpi]          .Valeur:=1 ;

FSequenceLP[sp10Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Groupe:=1 ;
FSequenceLP[sp10Cpi]          .Valeur:=0 ;

FSequenceLP[sp05Cpi]           .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Groupe:=1 ;
FSequenceLP[sp05Cpi]           .Valeur:=48 ;

FSequenceLP[spUnderLine]      .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spUnderLine]      .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[spUnderline]      .Groupe:=1 ;
FSequenceLP[spUnderline]      .Valeur:=128 ;

FSequenceLP[spBold]           .Sqce[FALSE]:=#27+'!%s' ;  //#27+'G'+#0 ;
FSequenceLP[spBold]           .Sqce[TRUE] :=#27+'!%s' ; //=#27+'G'+#1 ;
FSequenceLP[spBold]           .Groupe:=1 ;
FSequenceLP[spBold]           .Valeur:=8 ;

FSequenceLP[spItalic]         .Sqce[FALSE]:=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#0 ;
FSequenceLP[spItalic]         .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'r'+#1 ;
FSequenceLP[spItalic]         .Groupe:=0 ;
FSequenceLP[spItalic]         .Valeur:=0 ;

FSequenceLP[spBitMap]         .Sqce[FALSE]:='' ;
FSequenceLP[spBitMap]         .Sqce[TRUE] :=#27+'*'+#1+'%s'+#0 ;
FSequenceLP[spBitMap]         .Groupe:=0 ;
FSequenceLP[spBitMap]         .Valeur:=0 ;

FSequenceLP[spPrintBmp]              .Sqce[TRUE] :='' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM300.InitPrinter ( var err : TMC_Err ) : Boolean ;
Var St : String ;
Begin
Result:=FALSE ;
St:=Sqce[spInitialise,TRUE] ; if St<>'' then Result:=WriteLP(St) ;
if Result then
   Begin
   St:=Sqce[spSensorPaperOut,TRUE] ; if St<>'' then Result:=writeLP(St) ;
   End ;
if not result then err:=LasTMC_Error ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLP_EPSONTM300.TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
Var tt   : String ;
    b    : Byte ;
    R    : String ;
    V    : Variant ;
Begin
Result:=ieOk ;
if Fois=1 then err.code:=0 ; err.Libelle:='' ;
tt:=Sqce[spActivate,TRUE]+Sqce[spstatus,TRUE]+Sqce[spDesactivate,TRUE] ;
if FSortie.SendString(tt) then
   Begin
   Attentresponse ;
   if lireResultat(1,VarString,V) then
      Begin
      R:=VString(V) ;
      if Length(R)=1 then
         Begin //Cause : fin du papier
         B:=ord(R[1]) ;
         If (B and 4)=4 then
            Begin
            Result:=ieErrorPaper ;
            Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1021) ;
            End else
         If (B and 1)=1 then
            Begin
            Result:=ieWarningPaper ;
            Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1022) ;
            End else
            Begin
            TesteEtat := False ; // Uniquement pour la 1ère ligne d'impression
            End ;
         End ;                                                                 
      End else//If LireResultat(....
      if not IsPrinter then
         Begin  // Si l'imprimante est série et si on n'a rien reçu, l'imprimante est sans doute éteinte
         Result:=iePowerOff ;
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1023) ; Exit ;
         End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM300.IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
var tt : String ;
Begin
Result:=Etat ;
if Result=iePowerOff then
   Begin
   Result:=ieRecovery ;
   End else
   Begin
   Err.Libelle:=MC_MsgErrDefaut(1025) ;
   if ((Result in [ieError,ieErrorPaper]) and
       (LanceMessage(Err,TRUE)=mryes)) then Begin FlastError:=ieErrorPaper ; Result:=LastError ; End else
   if (Result in [iewarning,ieRecovery,ieError,ieErrorPaper])  then //ieError et ieErrorPaper si on a respondu NON à la re-impresion
      Begin
      tt:='' ;
      if Result in [ieWarning,ieRecovery] then tt:=Sqce[spRestart,TRUE]+#2 else
         if Result=ieErrorPaper then tt:=Sqce[spRestart,TRUE]+#0 ;
      if tt<>'' then
         begin
         tt:=Sqce[spActivate,TRUE]+tt+Sqce[spDesactivate,TRUE] ;
         FSortie.SendString(tt) ;
         Result:=ieRecovery ;
         End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM300.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;
  if i>0 then Result[i]:=#238 ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLP_EPSONTM6000.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
OnTesteEtat:=TestePrinter ;
OnInitialise:=InitPrinter ;
OnErreur:=IlyaErreur ;
Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM6000.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
LongMaxBuffer:=800 ;
LeDelayLong:=12 ;
LeDelayCourt:=0 ;

Result:=Inherited ChargePortEtParams(APort,AParams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EPSONTM6000.ChargeSequences ;
Begin
Inherited ; 
FSequenceLP[spCR]             .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]             .Sqce[TRUE] :=#10 ;

FSequenceLP[spInitialise]     .Sqce[TRUE] :=#27+'@'+#27+'3'+#16+#27+'R'+#1+#27+'t'+#19 ;

FSequenceLP[spSautPage]       .Sqce[TRUE] :='' ; //On n'a pas de saut de page en rouleau

FSequenceLP[spStatus]         .Sqce[TRUE] :=#16+#4 ; //1,2,3,4

if Port in [pnLPT1,pnLPT2] then FSequenceLP[spSensorPaperOut].Sqce[TRUE]:=#27+'c3'+#15 ; //Seulement si interface parallel

FSequenceLP[spReStart]        .Sqce[TRUE] :=#16+#5 ; //0,2

FSequenceLP[spActivateInterligne]    .Sqce[TRUE] :=#27+'3'+#24 ;
FSequenceLP[spDesactivateInterligne] .Sqce[TRUE] :=#27+'3'+#16 ;

FSequenceLP[spTotalCut]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VA'+#1 ;

FSequenceLP[spPartialCut]     .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#29+'VB'+#1 ;

FSequenceLP[spActivate]       .Sqce[TRUE] :=#27+'='+#1 ;

FSequenceLP[spDesactivate]    .Sqce[TRUE] :=#27+'='+#2 ;

FSequenceLP[spActivateSlip]          .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'<'+#27+'c0'+#4 ;

FSequenceLP[spDesactivateSlip]       .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'c0'+#1 ;

FSequenceLP[spActivateValidation]    .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'<'+#27+'c0'+#8 ;

FSequenceLP[spDesactivateValidation] .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#27+'c0'+#1 ;

FSequenceLP[spDebutCheque]           .Sqce[TRUE] :=FSequenceLP[spActivateSlip].Sqce[TRUE]+#27+'L'+#27+'W'+#0+#0+#0+#0+#28+#2+#74#3+#27+'T'#1 ;

FSequenceLP[spFinCheque]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]+#12+FSequenceLP[spDesactivateSlip].Sqce[TRUE] ;

FSequenceLP[sp17Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp17Cpi]          .Groupe:=1 ;
FSequenceLP[sp17Cpi]          .Valeur:=32 ;

FSequenceLP[sp12Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp12Cpi]          .Groupe:=1 ;
FSequenceLP[sp12Cpi]          .Valeur:=1 ;

FSequenceLP[sp10Cpi]          .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp10Cpi]          .Groupe:=1 ;
FSequenceLP[sp10Cpi]          .Valeur:=0 ;

FSequenceLP[sp05Cpi]           .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[sp05Cpi]           .Groupe:=1 ;
FSequenceLP[sp05Cpi]           .Valeur:=48 ;

FSequenceLP[spUnderLine]      .Sqce[FALSE]:=#27+'!%s' ;
FSequenceLP[spUnderLine]      .Sqce[TRUE] :=#27+'!%s' ;
FSequenceLP[spUnderline]      .Groupe:=1 ;
FSequenceLP[spUnderline]      .Valeur:=128 ;

FSequenceLP[spBold]           .Sqce[FALSE]:=#27+'!%s' ;  //#27+'G'+#0 ;
FSequenceLP[spBold]           .Sqce[TRUE] :=#27+'!%s' ; //=#27+'G'+#1 ;
FSequenceLP[spBold]           .Groupe:=1 ;
FSequenceLP[spBold]           .Valeur:=8 ;

FSequenceLP[spItalic]         .Sqce[FALSE]:='' ;  //non supported
FSequenceLP[spItalic]         .Sqce[TRUE] :='' ;
FSequenceLP[spItalic]         .Groupe:=0 ;
FSequenceLP[spItalic]         .Valeur:=0 ;

FSequenceLP[spBitMap]         .Sqce[FALSE]:='' ;
FSequenceLP[spBitMap]         .Sqce[TRUE] :=#27+'*'+#1+'%s'+#0 ;
FSequenceLP[spBitMap]         .Groupe:=0 ;
FSequenceLP[spBitMap]         .Valeur:=0 ;

FSequenceLP[spCodeBarre]             .Sqce[TRUE] :=FSequenceLP[spLF].Sqce[TRUE]
                                                  +#29+'h%s'   //Hauteur cdu codebarre (Nombre points * lignes)
                                                  +#29+'H%s'   //Position u HRI
                                                  +#29+'f'+#1  //%s' //Font à utiliser //XMG 23/06/03
                                                  +#29+'w'+#1  //%s' //Longeur du module....
                                                  +#29+'k%s'  //Code du codebarre plus valeur du CodeBarre à imprimer
                                                  +FSequenceLP[spLF].Sqce[TRUE] ;
FSequenceLP[spCodeBarre]             .Extra:='9;' ; //XMG 11/06/03

FSequenceLP[spPrintBmp]              .Sqce[TRUE] :=#28+'p'+#1+'0' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM6000.InitPrinter ( var err : TMC_Err ) : Boolean ;
Var St : String ;
Begin
Result:=FALSE ;
St:=Sqce[spInitialise,TRUE] ; if St<>'' then Result:=WriteLP(St,TRUE) ;
if Result then
   Begin
   St:=Sqce[spSensorPaperOut,TRUE] ; if St<>'' then Result:=writeLP(St,TRUE) ;
   End ;
if not result then err:=LasTMC_Error ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLP_EPSONTM6000.TestePrinter ( Fois : INteger ; Var Err : TMC_Err ) : TImprimErreurs ;
Var tt   : String ;
    b    : Byte ;
    R    : String ;
    V    : Variant ;
Begin
Result:=ieOk ;
if Fois=1 then err.code:=0 ; err.Libelle:='' ;
tt:=Sqce[spActivate,TRUE]+Sqce[spstatus,TRUE]+chr(Fois)+Sqce[spDesactivate,TRUE] ;
if FSortie.SendString(tt) then
   Begin
   Attentresponse ;
   if lireResultat(1,VarString,V) then
      Begin
      R:=VString(V) ;
      if Length(R)=1 then b:=ord(R[1]) else B:=18 ;
      if (B<>(18+4*ord(Fois=1))) then
         Case Fois of
           1 : Begin //Status de l'imprimante
               If (B and 32)=32 then //L'imprimante essaie de se recuperer
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1015) ;
                  End else
               if (B and 8)=8 then //L'imprimante est Off-Line
                  Begin
                  Result:=TestePrinter(2,Err) ;
                  if Result=ieOk then //Si aucun erreur est rapporté au deuxième niveau alors c'est OFF-Line
                     Begin
                     Result:=ieOffLine ;
                     Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1016) ;
                     End ;
                  End ;
               End ; //3: ....
           2 : Begin //Status de l'imprimante quand elle est en Off-Line
               if (B and 64)=64 then //Error
                  Result:=TestePrinter(3,Err) else
               if (B and 32)=32 then //End of paper
                  Result:=TestePrinter(4,Err) {else
               if (B and 4)=4 then //Cover Open
                  Begin
                  Result:=ieCover ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1017) ; Exit ;
                  End ;  XMG Il n'est pas sûr que ce bit soit bien initialise (Voir documentation) }
               End ; //3: ....
           3 : Begin //Cause de l'erreur de l'imprimante
               If (B and 32)=32 then //
                  Begin
                  Result:=ieError ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1018) ;
                  End else
               If (B and 64)=64 then //
                  Begin
                  Result:=ieRecovery ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1019) ;
                  End else
               if ((B and 4)=4) or ((B and 8)=8) then
                  Begin
                  Result:=ieWarning ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1020) ;
                  End ;
               End ; //3: ....
           4 : Begin //Causse de la fin du papier
               If (B and 96)=96 then
                  Begin
                  Result:=ieErrorPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1021) ;
                  End else
               If (B and 12)=12 then
                  Begin
                  Result:=ieWarningPaper ;
                  Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1022) ;
                  End ;
               End ; // 4 : ...
           End ;//Case
      End else//If LireResultat(....
      if not IsPrinter then
         Begin  // Si l'imprimante est serial et on n'a rien reçu, possiblement l'imprimante est eteinte
         Result:=iePowerOff ;
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1023) ; Exit ;
         End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM6000.IlyaErreur ( Etat : TImprimErreurs ; Var Err : TMC_Err ) : TImprimErreurs ;
var tt : String ;
Begin
Result:=Etat ;
if Result=iePowerOff then
   Begin
   Result:=ieRecovery ;
   End else
   Begin
   Err.Libelle:=MC_MsgErrDefaut(1025) ;
   if ((Result in [ieError,ieErrorPaper]) and
       (LanceMessage(Err,TRUE)=mryes)) then Begin FlastError:=ieErrorPaper ; Result:=LastError ; End else
   if (Result in [iewarning,ieRecovery,ieError,ieErrorPaper])  then //ieError et ieErrorPaper si on a respondu NON à la re-impresion
      Begin
      tt:='' ;
      if Result in [ieWarning,ieRecovery] then tt:=Sqce[spRestart,TRUE]+#2 else
         if Result=ieErrorPaper then tt:=Sqce[spRestart,TRUE]+#0 ;
      if tt<>'' then
         begin
         tt:=Sqce[spActivate,TRUE]+tt+Sqce[spDesactivate,TRUE] ;
         FSortie.SendString(tt) ;
         Result:=ieRecovery ;
         End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EPSONTM6000.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;
  if i>0 then Result[i]:=#213 ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 30/03/04 début
Constructor TLP_EXPORTERFICHIER.Create (AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner,AImprimante) ;
//Translate:=TRUE ;
OnTranslate:=TranslateASCII ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EXPORTERFICHIER.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
Result:=Inherited ChargePortEtParams(APort,AParams,Err) ;
LongMaxBuffer:=0 ;
LeDelayLong:=0 ;
LeDelayCourt:=0 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLP_EXPORTERFICHIER.ChargeSequences ;
Begin
Inherited ;
FSequenceLP[spCR]             .Sqce[TRUE] :=#13 ;

FSequenceLP[spLF]             .Sqce[TRUE] :=#10 ;
//La reste nulle.....
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLP_EXPORTERFICHIER.TranslateASCII ( Texte : String) : String ;
Var PosEuro : String ;
    i       : Integer ;
Begin
Result:='' ;
PosEuro:='' ;
repeat
  i:=pos('€',Texte) ;
  if i>0 then
     Begin
     PosEuro:=PosEuro+inttostr(i)+';' ;
     Texte[i]:=#32
     End ;
until i<=0 ;
Result:=Ansi2AscII(Texte) ;
while trim(PosEuro)<>'' do
  Begin
  i:=valeuri(readtokenSt(PosEuro)) ;
  if i>Length(Result) then Insere(' ',Result, i) ;
  if i>0 then Result[i]:=#213 ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 30/03/04 fin
end.
