unit AGLInitRT;
{==========================
Publication de fonctions et de procédures pour les scripts
===========================}


interface
uses HCtrls,Hent1 ,classes,
{$IFDEF EAGLCLIENT}
     UTob,
{$ELSE}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DB,
{$ENDIF}
     EntGC,EntRT ;

function  CreePieceRT ( parms: array of variant; nb: integer ) : variant ;

implementation

uses
  Windows,
  HMsgBox,
  M3FP,
  Controls,
  SysUtils,
  Forms,
  LookUp,
  ShellAPI,
  FactUtil
  ,facture
  ,UtilXlsBTP
  ;


/////////////////// THGrid: récupération d'une valeur ////////////
procedure AGLRTLookupListAction( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     CC :THEdit ;
begin
F:=TForm(Longint(Parms[0])) ;
CC:=THEdit(F.FindComponent(string(Parms[1]))) ;
LookupCombo(CC);
end;

function  CreePieceRT ( parms: array of variant; nb: integer ) : variant ;
begin
  result:='' ;
  if string(parms[2])<>'' then
  begin
	if CreerPieceGRC(string(parms[2]),string(parms[0]),strtoint(string(parms[1]))) then result:=VH_GC.GCLastRefPiece;
  end else
  begin
	if CreerPieceGRC('DBT',string(parms[0]),strtoint(string(parms[1]))) then result:=VH_GC.GCLastRefPiece;
  end;
end;

function AGLRTCreerPieceDevis( parms: array of variant; nb: integer ) : variant ;
begin
  Result := CreePieceRT (parms,nb);
end;

function AGLRTRechDomAction( parms: array of variant; nb: integer ):variant ;
var QQ:TQuery;
begin
Result:='';
QQ:=OpenSql('Select RAC_LIBELLE from ACTIONS where RAC_AUXILIAIRE="'+string(Parms[0])+'" AND RAC_NUMACTION='+string(Parms[1]),TRUE,-1,'',true);
if not QQ.EOF then Result:= QQ.fields[0].asstring;
ferme(QQ);
end;

function AGLRTRechDomContact( parms: array of variant; nb: integer ):variant ;
var QQ:TQuery;
begin
Result:='';
QQ:=OpenSql('Select C_NOM from CONTACT where C_TYPECONTACT="'+string(Parms[0])+'" AND C_AUXILIAIRE="'+string(Parms[1])+'" AND C_NUMEROCONTACT='+string(Parms[2]),TRUE,-1,'',true);
if not QQ.EOF then Result:= QQ.fields[0].asstring;
ferme(QQ);
end;

procedure AGLRTOuvrirDocWord( parms: array of variant; nb: integer ) ;
var stDocWord : String ;
begin
stDocWord:=string(parms[0]) ;
FileExecAndWait (GetWinwordPath+'WinWORD.exe "'+stDocWord+'"');
//ShellExecute( 0, PCHAR('open'),PChar(stDocWord), Nil,Nil,SW_RESTORE);
end;

procedure AGLRTDetruireDocWord( parms: array of variant; nb: integer ) ;
var stDocWord : String ;
begin
stDocWord:=string(parms[0]) ;
if FileExists(stDocWord) then
   if PGIAsk('Confirmez-vous la suppression de ce document ?',stDocWord)=mryes then DeleteFile(stDocWord ) else exit;
end;

function AglValeurI ( parms: array of variant; nb: integer ):variant ;
begin
  Result := Integer ( ValeurI ( string ( parms[0] ) ) ) ;
end;

//////////////////////////////////////////////////////////////////////////////
procedure initM3Gescom();
begin
 RegisterAglProc( 'RTLookupListAction', True,1,AglRTLookupListAction) ;
 RegisterAglFunc( 'RTRechDomAction', False,2,AglRTRechDomAction) ;
 RegisterAglFunc( 'RTRechDomContact', False,3,AGLRTRechDomContact) ;
 RegisterAglFunc( 'RTCreerPieceDevis', False,1,AGLRTCreerPieceDevis) ;
 RegisterAglProc( 'RTOuvrirDocWord', False,1,AGLRTOuvrirDocWord) ;
 RegisterAglProc( 'RTDetruireDocWord', False,1,AGLRTDetruireDocWord) ;

 RegisterAglFunc ( 'ValeurI', False, 1, AglValeurI ) ;
 end;

procedure RTAGLFonctionCTI (parms:array of variant; nb:integer);
var
   strTelephone   :string;
begin
{$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
  if string (Parms [0]) = 'MAKECALL' then
    begin
    strTelephone := Trim (string (Parms [1]));
    if strTelephone <> '' then
      if PgiAsk ('Appeler le ' + strTelephone + ' ?') = mrYes then
        begin
        VH_RT.ctiAlerte.Close;
        VH_RT.ctiAlerte.MakeCall (strTelephone);
        end;
    end;
{$ENDIF}
end;

initialization
  initM3Gescom();
  RegisterAglProc ( 'FonctionCTI', FALSE, 2, RTAGLFonctionCTI);
end.
