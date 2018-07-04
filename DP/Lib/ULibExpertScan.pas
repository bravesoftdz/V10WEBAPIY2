unit ULibExpertScan;

interface

uses
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
 Hctrls ,
 Classes ,
 uhttp,
 ExploitEWS_TLB,
 UTOB ;


type

 TRecErrorScan = record
   RC_Error    : integer;
   RC_Message  : string;
   RC_Methode  : string ;
   RC_Valeur   : string ;
 end;

 TErrorProcScan = procedure (sender : TObject; Error : TRecErrorScan ) of object;
 TNotifyProgressScan = procedure ( vStDossier,vStNature, vStMessage : string ) of object;
 TNotifyInfoScan = procedure ( vStMessage : string ) of object;


 TExpertScan = Class
  private
   FTOBUtilisat      : TOB ;
   FeWS              : IGeneral ;
   FStPath           : string ;
   FStUserSystem     : string ;
   FStPwSystem       : string ;
   FStDom            : string ;
   FStIp             : string ;
   FStsynchro        : string ;
   FStUserPGI        : string ;
   FStPwUserPGI      : string ;
   FHttpServer       : THttpServer ;
   FLastEwsError     : integer ;
   FScanExpert       : IScanExpert;
   FeWSSecurite      : string ;
   FeWSPme           : string ;
   FeWSNature        : string ;
   FBoeWSActif       : boolean ;
   FeWSVersion       : integer ;
   FCWASName         : string ;
   FLastError        : TRecErrorScan ;
   FOnError          : TErrorProcScan ;
   FOnProgress       : TNotifyProgressScan ;
   FOnInfo           : TNotifyInfoScan ;
   function  GenereXMLNature( vStDossier , vStNature : string ) : string ;
   procedure SeteWSSecurite( Value : string ) ;
   procedure SeteWSVersion( Value : integer ) ;
   procedure SetIP ( value : string ) ;
   procedure NotifyError( RC_Error : integer ; RC_Message : string ; RC_Methode : string = '' ) ;
   procedure SetOnError( Value : TErrorProcScan ) ;
   procedure SetOnProgress ( Value : TNotifyProgressScan ) ;
   procedure SetOnInfo ( Value : TNotifyInfoScan ) ;
  public
   constructor Create ;
   destructor  Destroy ; override ;

   procedure   VideDossier ;
   procedure   AjouteDossier ( vStDossier : string ) ;
   function    CreationArboDoss ( vStDossier : string ) : integer ;
   function    ScanNbFichierD( vStDossier : string ) : integer ;
   function    GenererXML : boolean ;
   function    eWSSynchro( vStDossier : string ) : integer ;
   function    eWSRecuperer( vDossier : string ) : integer ;
//   function    eWSRecuperer1( vDossier : string ) : integer ;
   function    eWSDebloquer( vDossier : string ) : boolean ;
   function    eWSEstBloque( vDossier : string ) : boolean ;
   function    eWSActiverScan( vStDossier : string ) : boolean ;

   property    OnError        : TErrorProcScan      read FOnError           write SetOnError ;
   property    OnProgress     : TNotifyProgressScan read FOnProgress        write SetOnProgress ;
   property    OnInfo         : TNotifyInfoScan     read FOnInfo            write SetOnInfo ;
   property    Path           : string              read FStPath            write FStPath ;
   property    UserSystem     : string              read FStUserSystem      write FStUserSystem ;
   property    PwSystem       : string              read fStPwSystem        write fStPwSystem ;
   property    Dom            : string              read FStDom             write FStDom ;
   property    Ip             : string              read FStIp              write SetIP ;
   property    eWSSecurite    : string              read FeWSSecurite       write SeteWSSecurite ;
   property    eWSActif       : boolean             read fBoeWSActif ;
   property    V_eWS          : IGeneral            read feWS               write feWS ;
   property    EwsVersion     : integer             read feWSVersion        write SeteWSVersion ;
   property    PwUserPGI      : string              read fStPwUserPGI       write fStPwUserPGI ;
   property    UserPGI        : string              read fStUserPGI         write fStUserPGI ;
   property    CWASName       : string              read fCWASName          write fCWASName ;
  end ;


procedure CSetParamScanEWS( vStNoDossier,Value : string ) ;
procedure CSetParamScanHP( vStNoDossier , Value : string ) ;
//function  CGenereXML( vStDossier : string ) : string ;


implementation

uses
 HEnt1 ,
 cbpPath ,
 Licutil ,
 XMLDoc ,
 XMLIntf ,
 uHttpCS ,
 SysUtils ;

const
 cStEnteteXML = '<?xml version="1.0" encoding="UTF-8"?><configuration version="8.10"><fileVersion>Release 1</fileVersion>' +
	              '<topLevelButton ID="73ef99fa-e13f-11db-83-14-08-00-20-0c-9a-66"><title>Scan Cegid Expert</title>' +
		            '<description>Cegid Expert Scan</description><iconUp>' +
			          '<iconUrl width="48" height="44">' ; //http://10.103.120.106/animation.gif</iconUrl>' +
		          //  '</iconUp><form><label>liste des utilisateurs</label>' ;


function TExpertScan.CreationArboDoss ( vStDossier : string ) : integer ;

 procedure _CreationRep( vStPath : string ) ;
 begin
  if not DirectoryExists(vStPath ) then
   begin
    if not CreateDir(vStPath ) then
     begin
      NotifyError(99 , 'Impossible de créer ' + vStPath );
      exit ;
     end
      else
       begin
        OnProgress('','','Création du répertoire : ' + vStPath ) ;
        inc(result) ;
       end ;
   end ;
 end ;

begin

  result  := 0 ;

  if fStPath = '' then
   begin
    NotifyError(99 , 'Répertoire de destination des scans non renseignés') ;
    exit ;
   end ;

  _CreationRep(fStPath + '\' + vStDossier ) ;
  _CreationRep(fStPath + '\' + vStDossier + '\OD' ) ;
  _CreationRep(fStPath + '\' + vStDossier + '\VTE' ) ;
  _CreationRep(fStPath + '\' + vStDossier + '\ACH' ) ;
  _CreationRep(fStPath + '\' + vStDossier + '\CAI' ) ;
  _CreationRep(fStPath + '\' + vStDossier + '\BQE' ) ;

end ;



procedure _CSetParamScan( vBoHP : boolean ; Value : string  ; vStNoDossier : string ) ;
var
 lStCle    : string ;
 lTOB      : TOB ;
 lStChamps : string ;
begin

 lStCle := 'YDS_TYPE="CES" and YDS_CODE="SCANEXPERT" and YDS_NODOSSIER="' + vStNoDossier + '" and YDS_PREDEFINI="DOS"' ;
 if vBoHP then
  lStChamps := 'YDS_LIBRE'
   else
    lStChamps := 'YDS_LIBELLE' ;

 if ExisteSQL('select YDS_TYPE from CHOIXDPSTD where ' + lStCle) then
  ExecuteSQL('update CHOIXDPSTD set ' + lStChamps + '="' + Value + '" where ' + lStCle)
   else
    begin
     lTOB := TOB.Create('CHOIXDPSTD',nil,-1) ;
     lTOB.PutValue('YDS_TYPE'       , 'CES');
     lTOB.PutValue('YDS_CODE'       , 'SCANEXPERT');
     lTOB.PutValue('YDS_NODOSSIER'  , vStNoDossier);
     lTOB.PutValue('YDS_PREDEFINI'  , 'DOS');
     lTOB.PutValue('YDS_LIBELLE'    , '-' );
     lTOB.PutValue('YDS_LIBRE'      , '-' );
     lTOB.PutValue(lStChamps        , Value );
     lTOB.InsertDB(nil) ;
     lTOB.Free ;
    end ;

end ;

procedure CSetParamScanHP( vStNoDossier,Value : string ) ;
begin
 _CSetParamScan(true,Value,vStNoDossier) ;
end ;

procedure CSetParamScanEWS( vStNoDossier,Value : string ) ;
begin
  _CSetParamScan(false,Value,vStNoDossier) ;
end ;


function TExpertScan.ScanNbFichierD( vStDossier : string ) : integer ;
var
 lStPath    : string ;

 procedure _Compte( vStNat : string ) ;
 var
  lSearchRec : TSearchRec ;
 begin

  if FindFirst(lStPath + '\' + vStNat + '\*.*' , faAnyFile, lSearchRec) = 0 then
   repeat
    if lSearchRec.Attr <> faDirectory then
     Inc(Result) ;
   until FindNext(lSearchRec) <> 0 ;

  FindClose(lSearchRec) ;

 end ;

begin

 result := 0 ;

 if fStPath = '' then exit ;

 lStPath := fStPath + '\' + vStDossier ; // + '\' + vStUtilisateur ;

 if lStPath = '' then exit ;

 _Compte('OD') ;
 _Compte('ACH') ;
 _Compte('VTE') ;
 _Compte('CAI') ;
 _Compte('BQE') ;


end ;


{function  CGenereXML( vStDossier : string ) : string ;
var
 lQ : TQuery ;
begin

  lQ := OpenSQL('select us_utilisateur,us_libelle from utilisat, liendonnees,liendosgrp ' +
                'where ldo_nodossier="' + vStDossier + '" ' +
                'and ldo_NOM = "GROUPECONF" ' +
                'and ldo_grpid=lnd_grpid ' +
                'and lnd_userid=us_utilisateur ' , true ) ;

  if lQ.Eof then
   begin
//    FLastError := 'Aucun utilisateur pour se dossier' ;
    exit ;
   end ;

  while not lQ.Eof do
   begin

   end ;


end ;}

{function _GenererXMLUtilisateur(  vStUtilisateur : string ) : string ;
var
 lStSQL : string ;
 lQ     : TQuery ;
begin

 lStSQL := 'SELECT DOS_NODOSSIER, dos_libelle '+
           'FROM DOSSIER, ANNUAIRE WHERE DOS_GUIDPER=ANN_GUIDPER AND ' +
           '( ' +
           'NOT EXISTS (SELECT 1 FROM GRPDONNEES, LIENDOSGRP ' +
           'WHERE GRP_NOM = LDO_NOM AND GRP_NOM = "GROUPECONF" ' +
           'AND GRP_ID = LDO_GRPID AND LDO_NODOSSIER=DOS_NODOSSIER) ' +
           'OR EXISTS (SELECT 1 FROM LIENDOSGRP, GRPDONNEES, LIENDONNEES ' +
           'WHERE GRP_NOM = LDO_NOM AND GRP_NOM = LND_NOM AND GRP_NOM = "GROUPECONF" ' +
           'AND GRP_ID = LDO_GRPID AND GRP_ID = LND_GRPID AND LND_USERID="' + vStUtilisateur + '" ' +
           'AND LDO_NODOSSIER=DOS_NODOSSIER) ' +
           ') ' +
           'AND DOS_NODOSSIER<>"000STD" ' +
           'ORDER BY DOS_NODOSSIER' ;

 lQ := OpenSQL(lStSQL,true) ;

 result := '<label>'+vStUtilisateur+'</label><menu>' ;

 while not lQ.Eof do
  begin
   result := result + _GenereXMLBase(lQ.FindField('DOS_NODOSSIER').AsString) ;
   lQ.Next ;
  end ; // while

 result := result + '</menu>' ;

end ; }

{ TExpertScan }

constructor TExpertScan.Create ;
begin
 //inherited ;
 FTOBUtilisat := TOB.Create('',nil,-1) ;
 fBoeWSActif  := false ;
end ;

destructor TExpertScan.Destroy ;
begin
 FTOBUtilisat.Free ;
 if fBoeWSActif then
  FScanExpert._Release ;
 inherited ;
end ;

procedure TExpertScan.VideDossier ;
begin
 FTOBUtilisat.ClearDetail ;
end ;

procedure TExpertScan.AjouteDossier( vStDossier : string ) ;
 var
 lQ       : TQuery ;
 lTOB     : TOB ;
 lTOBDoss : TOB ;
begin

 lq :=nil;
 try
 lQ := OpenSQL('select us_utilisateur,us_libelle from utilisat, liendonnees,liendosgrp ' +
                'where ldo_nodossier="' + vStDossier + '" ' +
                'and ldo_NOM = "GROUPECONF" ' +
                'and ldo_grpid=lnd_grpid ' +
                'and lnd_userid=us_utilisateur ' , true ) ;

 if lQ.Eof then
  begin
   lTOB := FTOBUtilisat.FindFirst(['UTILISATEUR'],[fStUserPGI], false) ;
   if lTOB = nil then
    begin
    lTOB := TOB.Create('', FTOBUtilisat , - 1 ) ;
    lTOB.AddChampSupValeur('UTILISATEUR',fStUserPGI);
    end ;
   lTOBDoss := TOB.Create('', lTOB , - 1 ) ;
   lTOBDoss.AddChampSupValeur('DOSSIER',vStDossier);
   exit ;
  end ;

 while not lQ.Eof do
  begin
   lTOB := FTOBUtilisat.FindFirst(['UTILISATEUR'],[lQ.FindField('us_utilisateur').asString], false) ;
   if lTOB = nil then
    begin
    lTOB := TOB.Create('', FTOBUtilisat , - 1 ) ;
     lTOB.AddChampSupValeur('UTILISATEUR',lQ.FindField('us_utilisateur').asString);
    end ;
   lTOBDoss := TOB.Create('', lTOB , - 1 ) ;
   lTOBDoss.AddChampSupValeur('DOSSIER',vStDossier);
   lQ.Next ;
  end ;
  finally
  ferme(lQ) ;
  end ;
end;

function TExpertScan.GenereXMLNature( vStDossier , vStNature : string ) : string ;
begin
 OnProgress('','','Création du xml pour : ' + vStDossier + ' : ' + vStNature ) ;
 result := '<element><submitButton><label>' + vStNature + '</label><filename>' + vStDossier + vstNature + '</filename>' +
           '<scan><target><colorMode>SEND_COLORMODE_MONO</colorMode><resolution>SEND_RESOLUTION_200x200</resolution>' +
           '<fileType>SEND_FILETYPE_JPEG</fileType></target></scan><destination><networkFolder>' +
           '<domain>' + fStDom + '</domain>' +
           '<username>' + fStUserSystem + '</username>' +
           '<password>' + fStPwSystem + '</password>' +
           '<path>' + fStIP + Copy(fStPath,4,length(fStPath)) + '/' + vStDossier + '/' + vStNature + '</path>' +
           '</networkFolder></destination></submitButton></element>' ;
end ;

function TExpertScan.GenererXML : boolean ;
var
 i,j       : integer ;
 lTOB      : TOB ;
 lTOBD     : TOB ;
 lXML      : WideString ;
 lfile     : TstringLIst ;
begin

 result := false ;
 lXML := cStEnteteXML ;
 lXML := lXML + 'http:' + fStIp + 'CegidExpertScan/cegid.gif</iconUrl></iconUp>' ;
 lXML := lXML + '<form><label>liste des utilisateurs</label><menu>' ;

 for i := 0 to FTOBUtilisat.Detail.Count -1 do
  begin
   lTOB := FTOBUtilisat.Detail[i] ;
   lXML := lXML + '<element><form><label>' + lTOB.GetString('UTILISATEUR') + '</label><menu>' ;
   for j := 0 to lTOB.Detail.Count - 1 do
    begin
     lTOBD := lTOB.Detail[j] ;
     lXML := lXML + '<element><form><label>' + lTOBD.GetString('DOSSIER') + '</label><actionList>' +
            GenereXMLNature(lTOBD.GetString('DOSSIER') , 'ACH') +
            GenereXMLNature(lTOBD.GetString('dossier') , 'VTE') +
            GenereXMLNature(lTOBD.GetString('dossier') , 'CAI') +
            GenereXMLNature(lTOBD.GetString('dossier') , 'BQE') +
            GenereXMLNature(lTOBD.GetString('dossier') , 'OD') ;
     lXML := lXML + '</actionList></form></element>' ;
    end ;
   lXML := lXML + '</menu></form></element>' ;
  end ;

 lXML := lXML + '</menu></form></topLevelButton></configuration>' ;

 lfile := TStringList.Create ;
 try
  if FileExists(TcbpPath.GetCegidUserTempPath + '\config.xml') then
   DeleteFile(TcbpPath.GetCegidUserTempPath + '\config.xml') ;
  lfile.Add(lXML) ;
  lfile.SaveToFile(TcbpPath.GetCegidUserTempPath + '\config.xml');
  OnInfo('Création du fichier de configuration : ' + TcbpPath.GetCegidUserTempPath + '\config.xml' ) ;
 finally
  lfile.free ;
 end ;

 if fCWASName = '' then
  begin
   NotifyError(999 , 'le serveur CWAS n''est pas renseigné ( Module Administration\Administration\Préférences\Cegid Expert Scan\Serveur CWAS') ;
   exit ;
  end ;
 FHttpServer:=ConnectHttpServer(fCWASName,FLastError.RC_Message);
 try
 if FHttpServer = nil then
  begin
   NotifyError(999 , 'le serveur CWAS ' + fCWASName + ' n''est pas démarré ' + #10#13 + FLastError.RC_Message ) ;
   exit ;
  end ; // if
 if not ConnecteAGLServer(FHttpServer, fStUserPGI, fStPwUserPGI , 'DB000000', 'cegidpgi.ini') then
  begin
   NotifyError(999 , 'Impossible de ce connecter au serveur CWAS ' + fCWASName + ' avec l''utilisateur ' + fStUserPGI + #10#13 + FLastError.RC_Message ) ;
   exit ;
  end ;
 FHttpServer.Upload('./CegidExpertScan/config.xml',TcbpPath.GetCegidUserTempPath + '\config.xml',FLastError.RC_Message) ;
 if FLastError.RC_Message <> '' then
  NotifyError(999 , 'Erreur durant le transfert' + #10#13 + FLastError.RC_Message )
   else
    FOnInfo('Envoi du fichier config.xml réussi ds le répertorie \wwwroot\CegidExpertScan') ;
 finally
 DeconnectHttpServer(FHttpServer) ;
 end ;
end;

procedure TExpertScan.SeteWSVersion( Value : integer ) ;
begin
 if Value < 2039 then
  begin
   fBoeWSActif := false ;
   NotifyError(99 , 'Le numéro de version du portail doit être égal à 2.40 ( version installée + ' + intToStr(FeWSVersion) ) ;
   exit ;
  end ;
 FeWSVersion := Value ;
end ;

procedure TExpertScan.SeteWSSecurite( Value : string ) ;
begin
 if trim(Value) = '' then
  begin
   NotifyError(99 , 'le numéro de portail est vide') ;
   exit ;
  end ;
 FeWSSecurite := Copy(Value,1,8) ;
 fBoeWSActif := true ;
 // pour test a enlever
// fBoeWSActif := false ;
 try
 FScanExpert := CoExploitEWS_ScanExpert.Create as IScanExpert;
 if FScanExpert = nil then
  begin
   NotifyError(99 , 'Impossible de créer l''objet eWS' ) ;
   fBoeWSActif := false ;
  end ;
 except
  on e:exception do
   begin
    fBoeWSActif := false ;
    NotifyError(99 , 'Impossible de créer l''objet eWS' + #10#13 + e.Message) ;
   end ;
 end ;
end ;

procedure TExpertScan.SetIP ( Value : string ) ;
var
 lCar : string ;
begin
 lCar := Copy(Value,1,2) ;
 if isNumeric(lCar) then
  fStIP := '//' + Value
   else
    if lCar = '\\' then
     fStIP := '//' + Copy(Value,3,length(value))
      else
       fStIp := Value ;

 lCar := Copy(fStIp,length(fStIp),1) ;
 if isNumeric(lCar) then
  fStIp := fStIP + '/'
   else
    if lCar = '\' then
     fStIp := Copy(fStIp,1,length(fStIp)-1) + '/'
      else
       fStIp := Value ;
end ;

function  TExpertScan.eWSEstBloque( vDossier : string ) : boolean ;
begin
 result := false ;
 if not fBoeWSActif then
  begin
   NotifyError(99 , 'Tentative de test de blocage de fichier avec les objets eWS non créés !') ;
   exit ;
  end ;
 if FeWSSecurite = '' then exit ;
 result := FScanExpert.EtatBloquage(FeWSSecurite,vDossier) = 1 ;
end ;

function TExpertScan.eWSDebloquer( vDossier : string ) : boolean;
begin
 result := false ;
 if not fBoeWSActif then
  begin
   NotifyError(99 , 'Tentative de déblocage de fichier avec les objets eWS non créés !') ;
   exit ;
  end ;
 if FeWSSecurite = '' then exit ;
 FLastEwsError := FScanExpert.DeBloquer(FeWSSecurite,vDossier) ;
 if FLastEwsError <> 0 then
  NotifyError(99 , 'Impossible de débloquer le dossier :' + vDossier)
   else
    result := true ;
end;

function TExpertScan.eWSRecuperer( vDossier : string ) : Integer;

 procedure _Notify ;
  begin
   if FLastEwsError <> 0 then
    NotifyError(FLastEwsError , 'Impossible de récupéré les fichiers :' + vDossier + #10#13 + FScanExpert.MessageErreur(FeWSSecurite) ) ;
  end ;

begin
 result := 0 ;
 if not fBoeWSActif then
  begin
   NotifyError(99 , 'Tentative de récupération de fichier avec les objets eWS non créés !') ;
   exit ;
  end ;
 if FeWSSecurite = '' then exit ;
 FLastEwsError := FScanExpert.Bloquer(FeWSSecurite,vDossier) ;
 if FLastEwsError <> 0 then
  begin
   NotifyError(FLastEwsError , 'Impossible de bloquer le dossier :' + vDossier + #10#13 + FScanExpert.MessageErreur(FeWSSecurite) ) ;
   exit ;
  end ;
 try
  OnProgress(vDossier,'Achat','Début de la récupération' ) ;
  FLastEwsError := FScanExpert.Recuperer(FeWSSecurite,vDossier,'Achat',fStPath+'\'+vDossier+'\ACH') ;
  _Notify ;
  OnProgress(vDossier,'Banque','Début de la récupération' ) ;
  FLastEwsError := FScanExpert.Recuperer(FeWSSecurite,vDossier,'Banque',fStPath+'\'+vDossier+'\BQE') ;
  _Notify ;
  OnProgress(vDossier,'Caisse','Début de la récupération' ) ;
  FLastEwsError := FScanExpert.Recuperer(FeWSSecurite,vDossier,'Caisse',fStPath+'\'+vDossier+'\CAI') ;
  _Notify ;
  OnProgress(vDossier,'OD','Début de la récupération' ) ;
  FLastEwsError := FScanExpert.Recuperer(FeWSSecurite,vDossier,'OD',fStPath+'\'+vDossier+'\OD') ;
  _Notify ;
  OnProgress(vDossier,'Vente','Début de la récupération' ) ;
  FLastEwsError := FScanExpert.Recuperer(FeWSSecurite,vDossier,'Vente',fStPath+'\'+vDossier+'\VTE') ;
  _Notify ;
  OnProgress(vDossier,'','Fin de la récupération' ) ;

 finally
  FLastEwsError := FScanExpert.DeBloquer(FeWSSecurite,vDossier) ;
 end ;

end;

{function TExpertScan.eWSRecuperer1( vDossier : string ) : Integer ;

begin
 OnProgress(vDossier,'Achat','Début de la récupération' ) ;
 result := FScanExpert.RecupererAsync(FeWSSecurite,vDossier,'Achat',fStPath+'\'+vDossier+'\ACH') ;
 while result > 0 do
  begin
   result := FScanExpert.AvancementAsync(FeWSSecurite) ;
   OnProgress(vDossier,'', 'Octet restant : ' + intToStr(result) ) ;
   sleep(1000) ;
  end ;
end ;}


function TExpertScan.eWSActiverScan( vStDossier : string ) : boolean ;
var
  WService, WIdClient, WPwdClient : WideString ;
  lInResult : integer ;
begin
 if not fBoeWSActif then
  begin
   NotifyError(99 , 'Tentative d''activation des fichiers avec les objets eWS non créés !') ;
   exit ;
  end ;
 lInResult := V_eWS.RetourneServiceClient(FeWSSecurite+'-01', UserPGI, DecryptageSt(PwUserPGI), vStDossier,
              'PDOCS', WService, WIdClient, WPwdClient) ;
 result := lInResult = 0 ;

end ;

function TExpertScan.eWSSynchro( vStDossier : string ) : integer ;
var
 lXML                  : string ;
 lStream               : Tstringstream;
 lDOM                  : IXmlDocument ;
 lNode,lRoot,lNodeNat  : IXMLNode ;
 i                     : integer ;
begin
 result := 0 ;
 if FeWSSecurite = '' then exit ;

 if not fBoeWSActif then
  begin
   NotifyError(99 , 'Tentative de récupération de fichier avec les objets eWS non créés !') ;
   exit ;
  end ;

 lXML := FScanExpert.ListeVolume(FeWSSecurite,vStDossier, '' ) ;
 if lXML= '' then
  begin
   NotifyError(99 , FScanExpert.MessageErreur(FeWSSecurite)) ;
   exit ;
  end ;

 lStream := TstringStream.Create(lXML);
 lDOM    := TXmlDocument.Create(nil) ;
 lDOM.LoadFromStream(lStream);
 lRoot   := lDOM.ChildNodes[0].ChildNodes[0] ;

 for i := 0 to lRoot.ChildNodes.Count - 1 do
  begin
   lNode := lRoot.ChildNodes[i] ;
   if lNode.NodeName = 'NATURE' then
    begin
     lNodeNat := lNode.ChildNodes[1] ;
     result := result + StrToInt(lNodeNat.NodeValue) ;
    end ;
  end ;

 lStream.Free ;
 FScanExpert._Release ;
 lDOM._Release ;

end ;

procedure TExpertScan.NotifyError( RC_Error : integer ; RC_Message : string ; RC_Methode : string = '') ;
begin
 FLastError.RC_Error   := RC_Error ;
 FLastError.RC_Message := RC_Message ;
 FLastError.RC_Methode := RC_Methode ;
 if assigned(FOnError) then FOnError(self,FLastError) ;
end;

procedure TExpertScan.SetOnError( Value : TErrorProcScan );
begin
 FOnError := Value ;
end;

procedure TExpertScan.SetOnProgress ( Value : TNotifyProgressScan ) ;
begin
 FOnProgress := Value ;
end ;


procedure TExpertScan.SetOnInfo ( Value : TNotifyInfoScan ) ;
begin
 FOnInfo := Value ;
end ;

end.
