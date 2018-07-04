unit TZBackup;

interface

uses Classes,
     SysUtils,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
      Forms,
      HEnt1,
      UTob,
      Hctrls ,
      ULibEcriture ,
      Syncobjs,  // pour TCriticalSection
      Windows ,
      TZ ;

type


 TZFThread = class(TThread)
  private
   PCriticalSection : TCriticalSection ;
   PStream          : TMemoryStream ;
   FTOBEcrCompl     : TOB ;
  public
    constructor Create ( const MyStream : TMemoryStream ; const vECrCompl : TOB ; MyCriticalSection : TCriticalSection ) ;
    procedure   Execute; override ;
  end ;


  TZFBackUp = class
  private
   FCriticalSection : TCriticalSection;
  function GetFini : boolean ;
  public
    Ecrs            : TOB ;
    EcrsCompl       : TOB ;

    constructor     Create ;
    destructor      Destroy ; override ;

    procedure       DelBadRow ;
    function        DelBackupFile : Boolean ;
    function        Backup : Boolean ;
    function        Restore : Boolean ; virtual ;
    function        FolioSauvegarde(var Par1, Par2, Par3 : string) : Boolean ;

    property        SauvegardePasFini : boolean read getFini  ;

  end ;


implementation

const
 _InMaxChamps = 39 ;
 _TRC : array[0.._InMaxChamps] of string =
 ( 'E_EXERCICE' , 'E_JOURNAL' , 'E_DATECOMPTABLE' , 'E_NUMEROPIECE' , 'E_NUMLIGNE' , 'E_GENERAL' ,
   'E_DEBIT', 'E_CREDIT', 'E_NATUREPIECE' ,'E_QUALIFPIECE','E_TYPEMVT', 'E_VALIDE','E_UTILISATEUR',
   'E_DATECREATION', 'E_DATEMODIF','E_SOCIETE' , 'E_ETABLISSEMENT' , 'E_TVAENCAISSEMENT' , 'E_LETTRAGEDEV' ,
   'E_DATEPAQUETMAX' , 'E_ECRANOUVEAU' ,'E_DATEPAQUETMIN' , 'E_DEVISE', 'E_DEBITDEV' , 'E_CREDITDEV' ,
   'E_TAUXDEV', 'E_COTATION' ,'E_MODESAISIE' , 'E_PERIODE' , 'E_SEMAINE' , 'E_IO' , 'E_PAQUETREVISION' ,
   'E_CONTROLE' , 'E_ETATLETTRAGE' , 'E_ENCAISSEMENT' , 'E_CONTROLETVA'   ,'E_CREERPAR' , 'E_EXPORTE' , 'E_CONFIDENTIEL', 'E_LIBELLE' ) ;

{***********UNITE*************************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/07/2005
Modifié le ... :   /  /
Description .. : - 21/07/2005 - LG - FB 14819 - on donne le chemin de l'exe
Suite ........ : ds le nom du fichier. ne fct pas qd on numerisait un doc
Mots clefs ... :
*****************************************************************}
function _GetFileName( vNom : string )  : string ;
begin
{$IFDEF EAGLCLIENT}
 result := V_PGI.CodeSociete+V_PGI.User+vNom+'.BAK' ;
{$ELSE}
 result := UpperCase(ExtractFilePath(Application.ExeName)) +V_PGI.CurrentAlias+V_PGI.User+vNom+'.txt' ;
{$ENDIF}
end ;


procedure _WriteBinStreamSt(var TheStream : TMemoryStream ; Value : String ; len : Integer) ;
begin
 if Len <=0 then exit ;
 while Length(Value) < len do  Value := Value + ' ' ;
 TheStream.Write(Value[1],len) ;
end ;

procedure _WriteBinStreamValue( TheStream : TMemoryStream ; Value : variant ) ;
var
 lSize            : integer ;
 lType            : integer ;
 lBufferString    : String ;
 lBufferDate      : TDateTime ;
 lBufferDouble    : Double ;
 lBufferInteger   : integer ;
begin

 case VarType(Value) of
  varNull,
  varEmpty,
  varError : begin
              lSize := 0 ;
              lType := varNull ;
              TheStream.Write(lSize , Sizeof(lSize)) ;
              TheStream.Write(lType , Sizeof(lSize)) ;
             end ;
   varDate : begin
              lSize       := SizeOf(TDateTime) ;
              lType       := varDate ;
              lBufferDate := VarAsType(Value,vardate) ;
              TheStream.Write(lSize ,Sizeof(lSize)) ;
              TheStream.Write(lType ,Sizeof(lType)) ;
              TheStream.Write(lBufferDate , lSize) ;
             end ;
   varInteger,
   varSmallint,
   varByte :   begin
                lSize           := sizeof(Integer) ;
                lType           := varInteger ;
                lBufferInteger  := VarAsType(Value,varInteger) ;
                TheStream.Write(lSize,Sizeof(lSize)) ;
                TheStream.Write(lType,Sizeof(lType)) ;
                TheStream.Write(lBufferInteger,lSize) ;
              end ;
   varSingle,
   varDouble,
   varCurrency : begin
                  lSize         := sizeOf(double) ;
                  lType         := varDouble ;
                  lBufferDouble := VarAsType(Value,varDouble) ;
                  TheStream.Write(lSize,Sizeof(lSize)) ;
                  TheStream.Write(lType,Sizeof(lType)) ;
                  TheStream.Write(lBufferDouble ,lSize) ;
                 end ;
   else          begin
                  lBufferString := VarAsType(Value,varString) ;
                  lSize         := length(lBufferString) ;
                  lType         := varString ;
                  TheStream.Write(lSize,Sizeof(lSize)) ;
                  TheStream.Write(lType,Sizeof(lType)) ;
                  _WriteBinStreamSt(TheStream,lBufferString,lSize) ;
                 end ;
   end ; // case

end ;

function _TOBEcrToStream( TheTOB : TOB ) : TMemoryStream ;
var
 i          : integer ;
 j          : integer ;
 lInIndex   : integer ;
 _RecIndex  : array[0.._InMaxChamps] of integer ;
begin

 result := nil ;

 lInIndex := TheTOB.Detail.Count ;
 if lInIndex = 0 then exit ;

 for i := low(_TRC) to high(_TRC) do
  _RecIndex[i] := ChampToNumDicho(PrefixeToNum ( 'E' ) , _TRC[i] ) ;

 result  := TMemoryStream.Create ;

 result.Write(BeginTOBComment,Sizeof(BeginTOBComment)) ;
 _WriteBinStreamValue(result,TheTOB.Detail[0].GetValue('E_DATECOMPTABLE') ) ;
 _WriteBinStreamValue(result,TheTOB.Detail[0].GetValue('E_JOURNAL') ) ;
 _WriteBinStreamValue(result,TheTOB.Detail[0].GetValue('E_NUMEROPIECE') ) ;
 result.Write(BeginTOBComment,Sizeof(BeginTOBComment)) ;

 for i := 0 to lInIndex - 1 do
  begin
   result.Write(BeginTOBValue,Sizeof(BeginTOBValue)) ;
   for j := low(_TRC) to high(_TRC) do
     _WriteBinStreamValue(result,TheTOB.Detail[i].GetValeur(_RecIndex[j]))
  end ; // for

end ;

procedure _ReadBinStreamSt( var TheStream : TMemoryStream ; var Value : String ; len : Integer ) ;
begin
 Value := '' ;
 if len <= 0 then exit ;
 SetLength(Value,len) ;
 TheStream.Read(Value[1],len) ;
end ;

function _ReadBinStreamValue( TheStream : TMemoryStream ) : variant ;
var
 lType            : integer ;
 lBufferDate      : TDateTime ;
 lSize            : integer ;
 lBufferInteger   : integer ;
 lBufferDouble    : double ;
 lBufferString    : string ;
begin

 TheStream.Read(lSize,Sizeof(lSize)) ;
 TheStream.Read(lType,Sizeof(lType)) ;

 case lType of
  varNull :    result := #0 ;
  varDate :    begin
                TheStream.Read( lBufferDate, lSize) ;
                result := lBufferDate ;
               end ;
  varInteger : begin
                TheStream.Read(lBufferInteger,lSize) ;
                result := lBufferInteger ;
               end ;
  varDouble  : begin
                TheStream.Read(lBufferDouble,lSize) ;
                result := lBufferDouble ;
               end ;
  else         begin
                _ReadBinStreamSt(TheStream,lBufferString,lSize ) ;
                result := lBufferString ;
               end ;
 end ; // case

end ;

procedure _ReadCommentValue( TheStream : TMemoryStream ; var E_DATECOMPTABLE, E_JOURNAL , E_NUMEROPIECE : string ) ;
var
 lBuffer : char ;
begin
 E_DATECOMPTABLE := _ReadBinStreamValue(TheStream) ;
 E_JOURNAL       := _ReadBinStreamValue(TheStream) ;
 E_NUMEROPIECE   := _ReadBinStreamValue(TheStream) ;
 TheStream.Read( lBuffer , 1 ) ; // on passe le commentaire de fin
end ;

procedure _TOBLoadCommentFromBinStream( TheStream : TMemoryStream ; var E_DATECOMPTABLE, E_JOURNAL , E_NUMEROPIECE : string ) ;
var
 lBuffer : char ;
begin
 TheStream.Read( lBuffer , 1 ) ;
 if lBuffer = BeginTOBComment then
  _ReadCommentValue(TheStream,E_DATECOMPTABLE,E_JOURNAL,E_NUMEROPIECE ) ;
end ;

procedure _TOBLoadFromBinStream( TheStream : TMemoryStream ; TheTOB : TOB ) ;
var
 j               : integer ;
 lBuffer         : char ;
 lTOB            : TOB ;
 E_DATECOMPTABLE : string ;
 E_JOURNAL       : string ;
 E_NUMEROPIECE   : string ;
begin

 while TheStream.Position < TheStream.Size do
  begin
   TheStream.Read( lBuffer , 1 ) ;
   if lBuffer = BeginTOBComment then
    _ReadCommentValue(TheStream,E_DATECOMPTABLE,E_JOURNAL,E_NUMEROPIECE)
     else
      if lBuffer = BeginTOBValue then
       begin
        lTOB := TZF.Create ( 'ECRITURE', TheTOB, -1 ) ;
        for j := low(_TRC) to high(_TRC) do
         lTOB.PutValue(_TRC[j],_ReadBinStreamValue(TheStream) ) ;
        end ;
  end ; // while

end ;


constructor TZFThread.Create( const MyStream : TMemoryStream ; const vEcrCompl : TOB ;  MyCriticalSection : TCriticalSection );
begin
 PStream             := MyStream ;
 FTOBEcrCompl        := vEcrCompl ;
 PCriticalSection    := MyCriticalSection ;
 FreeOnTerminate     := true ;
 Inherited Create(false) ;
end;

procedure TZFThread.Execute ;
begin

 PCriticalSection.Enter ;
 try

  PStream.SaveToFile(_GetFileName('ECR')) ;
  PStream.Free ;

  FTOBEcrCompl.SaveToFile(_GetFileName('ECR'),False,true,false) ;
  FTOBEcrCompl.Free ;

 finally
  PCriticalSection.Leave ;
 end ;

end;

constructor TZFBackUp.Create ;
begin
 Ecrs              := TOB.Create('', nil, -1) ;
 EcrsCompl         := TOB.Create('', nil, -1) ;
 FCriticalSection  := TCriticalSection.Create;
end ;

destructor TZFBackUp.Destroy ;
begin
 if Ecrs <> nil then Ecrs.Free ;
 if EcrsCompl <> nil then EcrsCompl.Free ;
 Ecrs         := nil ;
 EcrsCompl    := nil ;
 FCriticalSection.Free ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/09/2002
Modifié le ... : 07/10/2002
Description .. : -25/09/2002 - ajout du code societe
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TZFBackUp.DelBackupFile : Boolean ;
var
 FileName : string ;
begin

 Result   := false ; exit ;

 FileName := _GetFileName('ECR') ;

 if FileExists(FileName) then
  Result := SysUtils.DeleteFile(FileName) ;
 if not result then exit ;

 FileName := _GetFileName('ECRCOMPL') ;
 Result   := false ;

 if FileExists(FileName) then
  Result := SysUtils.DeleteFile(FileName) ;

end ;

function TZFBackUp.Backup : Boolean ;
var
 lStream : TMemoryStream ;
 lTOB    : TOB ;
begin
 lStream := _TOBEcrToStream(Ecrs) ;
 lTOB    := TOB.Create('',nil,-1) ;
 lTOB.Dupliquer(EcrsCompl,true,true) ;
 TZFThread.Create(lStream,lTOB,FCriticalSection) ;
 Result  := true ;
end ;


function TZFBackUp.Restore : Boolean ;
var
 FileName : string ;
 MyStream : TMemoryStream ;
begin

 FileName := _GetFileName('ECR') ;
 Result   := false ;

 if FileExists(FileName) then
  begin
   Ecrs.ClearDetail ;
   EcrsCompl.ClearDetail ;
   MyStream := TMemoryStream.Create ;
   MyStream.LoadFromFile(FileName) ;
   _TOBLoadFromBinStream(MyStream,Ecrs) ;
   result := true ;
 end ;
 
end ;

function TZFBackUp.FolioSauvegarde( var Par1, Par2, Par3 : string ) : Boolean ;
var
 FileName : string ;
 MyStream : TMemoryStream ;
begin

 result   := false ;

 FileName := _GetFileName('ECR') ;

 if FileExists(FileName) then
  begin
   MyStream := TMemoryStream.Create ;
   try
    MyStream.LoadFromFile(FileName) ;
    _TOBLoadCommentFromBinStream(Mystream,Par1,Par2,Par3) ;
    result := (trim(Par1)<>'') and (trim(Par2)<>'') and (trim(Par3)<>'') ;
   finally
    MyStream.Free ;
   end ;
  end ; // if

end ;

procedure TZFBackUp.DelBadRow ;
var z : TList ; i : Integer ;
begin
z:=TList.Create ;
for i:=0 to Ecrs.Detail.Count-1 do
    if Ecrs.Detail[i].GetValue('BADROW')='X' then z.Add(Ecrs.Detail[i]) ;
for i:=0 to z.Count-1 do
 begin
 // CFreeTOBCompl(TZF(z[i])) ;
  TZF(z[i]).Free ;
 end ;
z.Clear ; z.Free ;
end ;


function TZFBackUp.GetFini: boolean;
begin
 result := true ;
end;


end.

