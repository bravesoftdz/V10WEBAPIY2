unit UtilMessages;

interface

uses Forms,
     SysUtils,
     HCtrls,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     Classes,
     HEnt1,
     HMsgBox,
     UTob,
     HPanel,
     UTreeLinks,controls;

function  DecodeDestinataires(sDestinataires: String): String;
function  DecodeStrDate(StrDate : String): String;
function  DecodeStrDateDate (strDate:string):TDateTime; // $$$ JP 15/01/07: renvoyer un TDateTime
function  GetNomUser(coduser: String): String;
function  DedoublonneDestinataires(sDestinataires: String): String;

function  DeleteMessage(sMsgGuid: String): Boolean;
function  DeleteDansTableMessage(sMsgGuid : String): Boolean;

function  ExtractEmail(From: String): String;
function  MailsToMessages(F: TForm; TitreMsgError: String='';
  QuelUser: String=''; AvecDownload: Boolean=True): Boolean;
{$ifdef bureau}
procedure ShowDroitsMailsParUsers(PRien: THPanel; HDTDrawNode: TOnHDTDrawNode);
{$endif}
function  PremierUserAvecMessagerie: String;
function  IsRtf(Texte: String): Boolean;
procedure VerifDecodageMimeAddress(LstDest: TStringList; Address: String);
procedure VireGuillemets(var str: String);

////////////// IMPLEMENTATION //////////////
implementation

uses UMailBox, HDTLinks;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /
Description .. : Convertit la liste des codes des destinataires
Suite ........ : (utilisateurs PGI) en liste de noms détaillés
Suite ........ :
Suite ........ : (sauf pour les destinataires connus par leur emails :
Suite ........ : risque de lenteur)
Mots clefs ... :
*****************************************************************}
function DecodeDestinataires(sDestinataires: String): String;
var alldest, undest, libuser : String;
begin
  Result := '';
  alldest := sDestinataires;
  if alldest='' then exit;

  undest := ReadTokenSt(alldest);
  While undest<>'' do
    begin
    if Result<>'' then Result := Result+';';
    // Si ce n'est pas un email
    if Pos('@', undest)=0 then
      begin
      // recherche d'un user
      libuser := GetNomUser(undest);
      if libuser<>'' then
        Result := Result + libuser
      else
        Result := Result + undest;
      end
    else
      Result := Result + undest;
    // destinataire suivant
    undest := ReadTokenSt(alldest);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /    
Description .. : Tente de convertir au format 'aaaa DD/MM/YYYY
Suite ........ : HH:NN:SS' (jeudi 20/01/2005 15:12:17)
Suite ........ : les diverses dates que reçoit pop3, par exemple :
Suite ........ : Wed, 8 Sep 2004 20:47:23 +0300
Suite ........ : Fri, 24 Sep 2004 16:53:59 +0200
Suite ........ : Sun, 12 Dec 2004 20:11:58 +0100 (CE
Suite ........ : 13 Dec 2004 08:26:25 -0000
Mots clefs ... : CONVERSION;DATES;POP3
*****************************************************************}
function DecodeStrDate(StrDate : String): String;
var
    iPos : Integer;
begin
  Result := StrDate;
  // Tronque partie résiduelle
  iPos := Pos('+', Result);
  if iPos>20 then Result := Copy(Result, 1, iPos-1);
  iPos := Pos('-', Result);
  if iPos>20 then Result := Copy(Result, 1, iPos-1);

  // formatage
  Result := StringReplace(Result, ',',  ' ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, '.',  ' ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, '  ', ' ', [rfReplaceAll, rfIgnoreCase]);

  // jours
  Result := StringReplace(Result, 'mon ',       'lundi ',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'monday ',    'lundi ',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'tue ',       'mardi ',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'tuesday ',   'mardi ',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'wed ',       'mercredi ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'wednesday ', 'mercredi ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'thu ',       'jeudi ',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'thursday ',  'jeudi ',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'fri ',       'vendredi ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'friday ',    'vendredi ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'sat ',       'samedi ',   [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'saturday ',  'samedi ',   [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'sun ',       'dimanche ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'sunday ',    'dimanche ', [rfReplaceAll, rfIgnoreCase]);

  // mois
  Result := StringReplace(Result, ' jan ',       '/01/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' january ',   '/01/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' janvier ',   '/01/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' feb ',       '/02/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' february ',  '/02/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' fév ',       '/02/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' février ',   '/02/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' mar ',       '/03/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' march ',     '/03/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' mars ',      '/03/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' apr ',       '/04/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' april ',     '/04/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' avr ',       '/04/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' avril ',     '/04/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' may ',       '/05/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' mai ',       '/05/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' jun ',       '/06/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' june ',      '/06/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' juin ',      '/06/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' jul ',       '/07/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' july ',      '/07/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' juil ',      '/07/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' juillet ',   '/07/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' aug ',       '/08/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' august ',    '/08/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' aou ',       '/08/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' aout ',      '/08/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' août ',      '/08/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' sep ',       '/09/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' sept ',      '/09/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' september ', '/09/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' septembre ', '/09/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' oct ',       '/10/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' october ',   '/10/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' octobre ',   '/10/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' nov ',       '/11/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' november ',  '/11/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' dec ',       '/12/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' déc ',       '/12/',    [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' december ',  '/12/',    [rfReplaceAll, rfIgnoreCase]);
end;

// $$$ JP 15/01/07: renvoyer un TDateTime, nécessite quelque correction supplémentaire
function DecodeStrDateDate (strDate:string):TDateTime; overload; // $$$ JP 15/01/07: renvoyer un TDateTime
var
    iPos   :integer;
begin
     Result := iDate1900;

     // On enlève les virgule, points en trop
     strDate := StringReplace (strDate, ',',  '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, '.',  '', [rfReplaceAll, rfIgnoreCase]);
     strDate := Trim (strDate);

     // On enlève le nom des jours
     strDate := StringReplace (strDate, 'mon',        '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'monday',     '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'tue',        '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'tuesday ',   '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'wed ',       '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'wednesday ', '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'thu ',       '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'thursday ',  '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'fri ',       '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'friday ',    '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'sat ',       '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'saturday ',  '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'sun ',       '', [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, 'sunday ',    '', [rfReplaceAll, rfIgnoreCase]);
     strDate := Trim (strDate);

     // On remplace le nom des mois par leur
     strDate := StringReplace (strDate, ' jan ',       '/01/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' january ',   '/01/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' janvier ',   '/01/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' feb ',       '/02/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' february ',  '/02/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' fév ',       '/02/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' février ',   '/02/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' mar ',       '/03/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' march ',     '/03/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' mars ',      '/03/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' apr ',       '/04/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' april ',     '/04/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' avr ',       '/04/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' avril ',     '/04/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' may ',       '/05/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' mai ',       '/05/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' jun ',       '/06/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' june ',      '/06/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' juin ',      '/06/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' jul ',       '/07/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' july ',      '/07/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' juil ',      '/07/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' juillet ',   '/07/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' aug ',       '/08/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' august ',    '/08/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' aou ',       '/08/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' aout ',      '/08/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' août ',      '/08/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' sep ',       '/09/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' sept ',      '/09/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' september ', '/09/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' septembre ', '/09/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' oct ',       '/10/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' october ',   '/10/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' octobre ',   '/10/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' nov ',       '/11/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' november ',  '/11/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' dec ',       '/12/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' déc ',       '/12/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := StringReplace (strDate, ' december ',  '/12/',    [rfReplaceAll, rfIgnoreCase]);
     strDate := Trim (strDate);

     // On enlève ce qui suit l'heure ($$$todo: gérer le fuseau horaire, quand on saura comment ça fonctionne et quand la chaine sera non tronquée)
     iPos := Pos (':', strDate);
     if iPos > 0 then
     begin
          if (iPos<=Length (strDate)-5) and (strDate [iPos+3] = ':') then
              strDate := Copy (strDate, 1, iPos+5)
          else
              strDate := Copy (strDate, 1, iPos+2);
     end;
     strDate := Trim (strDate);

     // La date en TDateTime
     Result := StrToDateTime (strDate);
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /
Description .. : Renvoit le nom détaillé d'un utilisateur PGI à partir de son
Suite ........ : code
Mots clefs ... :
*****************************************************************}
function GetNomUser(coduser: String): String;
var Q: TQuery;
begin
  Result := '';
  if coduser='' then exit;
  Q := OpenSQL('SELECT US_LIBELLE FROM UTILISAT WHERE US_UTILISATEUR="'+ coduser + '"', True);
  if Not Q.Eof then Result := Q.FindField('US_LIBELLE').AsString;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /    
Description .. : Dans une liste de destinataires séparés par des ;
Suite ........ : retourne la même liste en éliminant les doublons
Mots clefs ... :
*****************************************************************}
function DedoublonneDestinataires(sDestinataires: String): String;
var alldest, undest : String;
    lesdest : TStringList;
begin
  Result := '';
  alldest := Trim(sDestinataires);
  if alldest='' then exit;

  lesdest := TStringList.Create;
  undest := ReadTokenSt(alldest);
  While undest<>'' do
    begin
    if lesdest.IndexOf(undest)=-1 then
      begin
      if Result<>'' then Result := Result+';';
      Result := Result + undest;
      lesdest.Add(undest);
      end;
    undest := ReadTokenSt(alldest);
    end;

  lesdest.Free;
end;

function DeleteDansTableMessage(sMsgGuid: String): Boolean;
var
  tMsg, tResult: Tob;
begin
  //$$$ JP 06/12/05 - warning delphi -> Result := False;
  tMsg := Tob.Create('_YMESSAGES_', nil, -1);
  tResult := Tob.Create('_RESULT_', nil, -1);

  try
    tMsg.LoadDetailDBFromSQL('YMESSAGES', 'SELECT * FROM YMESSAGES WHERE (YMS_MSGGUID="'+sMsgGuid+'")');
    tMsg.DeleteDBTom('YMESSAGES', False, tResult);

    if tResult.Detail.Count > 0 then
    begin
//      FLastErrorMsg := tResult.Detail[1].GetValue('ERROR');
      Result := False;
    end
    else
      Result := True;

  except
//    FLastErrorMsg := 'Erreur dans le processus';
    Result := False;
  end;

  tMsg.Free;
  tResult.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /
Description .. : Fonction de suppression d'un message dans la messagerie
Suite ........ : PGI
Mots clefs ... : MESSAGERIE;YMESSAGES
*****************************************************************}
function DeleteMessage(sMsgGuid : String): Boolean;
var sUserMail : String;
    iMailId   : Integer;
    Q         : TQuery;
begin
  Result := False;
  if sMsgGuid='' then exit;

  // Clé pour recherche YMAILS associé
  sUserMail := '';
  iMailId := 0;
  Q := OpenSQL('SELECT YMS_USERMAIL, YMS_MAILID FROM YMESSAGES WHERE YMS_MSGGUID="'+sMsgGuid+'"', True);
  if Not Q.Eof then
    begin
    sUserMail := Q.FindField('YMS_USERMAIL').AsString;
    iMailId := Q.FindField('YMS_MAILID').AsInteger;
    end;
  Ferme(Q);

  // Delete table principale YMESSAGES
  // (passe par le delete de la TOM ymessages, qui vérifie si suppression possible,
  // et effectue le ménage sur YMSGFILES/YMSGADDRESS)
  Result := DeleteDansTableMessage(sMsgGuid);

  // Delete YMAILS et dépendances
  // (passe par le delete de la TOM ymails, qui vérifie notamment qu'il n'y a pas d'autres
  //  messages rattachés à ce ymails avant de faire le ménage sur YMAILFILES,
  //  mais qui ne supprime pas dans YMAILADDRESS car c'est fait dans la fct UserMailDelete)
  if (Result) and (iMailId<>0) then
    UserMailDelete(iMailId, sUserMail);
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /
Description .. : Fonction qui tente de retrouver, dans une adresse email, la
Suite ........ : partie adresse réellement utilisable (voir aussi fonction
Suite ........ : DecodeMimeAddress de UMailBox)
Mots clefs ... :
*****************************************************************}
function ExtractEmail(From: String): String;
// Exemple : From=Service de Messagerie <postmaster@libertysurf.fr>
// Retournera : postmaster@libertysurf.fr
var tmp : String;
    iArrobase, iInferieur, iSuperieur : Integer;
begin
  Result := From;
  iArrobase := Pos('@', From);
  iInferieur := Pos('<', From);
  // si @ est trouvé après <
  if (iArrobase>0) and (iInferieur>0) and (iInferieur<iArrobase) then
    begin
    // démarre au début de l'adresse
    tmp := Copy(From, iInferieur+1, Length(From)-iInferieur);
    // élimine la fin
    iSuperieur := Pos('>', tmp);
    if iSuperieur > Pos('@', tmp) then
      begin
      tmp := Copy(tmp, 1, iSuperieur-1);
      Result := tmp;
      end;
    end;
  // FQ 11873 - élimine les guillemets
  VireGuillemets(Result);
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/05/2003
Modifié le ... :   /  /
Description .. : Récupération des nouveaux mails de YMAILS
Suite ........ : dans la messagerie/évènements PGI (YMESSAGES)
Mots clefs ... : DP;MAILS;YMAILS;MESSAGERIE;YMESSAGES
*****************************************************************}
function MailsToMessages(F: TForm; TitreMsgError: String='';
  QuelUser: String=''; AvecDownload: Boolean=True): Boolean;
var Q :TQuery;
    email, nodoss, SQL, SQLUser, UnUser, NomUser, Guidper, UserEmail : String;
    iLastMailId, iUser, tmpIndex : Integer;
    sNewMsgGUID : String;
    DataTypeLink : TYDataTypeLink;
    TobUser, TobMsg : TOB;
    TOBMail         : TOB; // $$$ JP 11/10/05 - remplace QMail, pour compatibilité en CWAS
    iMail           : integer;

    ymsDate : TDateTime;
    ymsLibfrom,ymsSujet,sAsk : string;
    booMaj : Boolean;
begin
  Result := True;
  // Distribution des mails d'un user précis
  if QuelUser<>'' then
    BEGIN
    SQLUser := 'SELECT US_UTILISATEUR, US_LIBELLE, US_EMAIL'
              +' FROM UTILISAT WHERE US_UTILISATEUR="'+QuelUser+'"'
    END
  // Distribution pour le user en cours
  else
    BEGIN
    // Au minimum récup des mails du user en cours ...
    SQLUser := 'SELECT US_UTILISATEUR, US_LIBELLE, US_EMAIL'
              +' FROM UTILISAT WHERE US_UTILISATEUR="'+V_PGI.User+'"';

    // ... mais + si il a des users dont il peut voir les mails
    DataTypeLink := V_DataTypeLinks.FirstMaster('YYUTILISATMASTER', tmpIndex);
    if DataTypeLink<>Nil then
      begin
      // Recherche les users pour lesquels le user en cours peut voir les mails
      SQL := GetTabletteSQL('YYUTILISATSLAVE', DataTypeLink.GetSql(V_PGI.User, False), True )
       + ' AND US_EMAIL<>""';
      // GetSQL donne les enfants : DataTypeLink.GetSql(V_PGI.User, False) renverra
      // ' EXISTS(SELECT * FROM YDATATYPETREES WHERE (YDT_CODEHDTLINK = "YYUSERMASTERSLAVE")
      //  AND (YDT_MCODE = "CEG") AND (YDATATYPETREES.YDT_SCODE = UTILISAT.US_UTILISATEUR))'
      if ExisteSQL(SQL) then
        // auquel on rajoute systématiquement le user en cours (*)
        SQLUser := 'SELECT US_UTILISATEUR, US_LIBELLE, US_EMAIL FROM UTILISAT '
         + 'WHERE (' + DataTypeLink.GetSql(V_PGI.User, False)
         + ' AND US_EMAIL<>"") OR US_UTILISATEUR="'+V_PGI.User+'"';
      end;
    END;

  TobUser := TOB.Create('_Les users pour relève des mails_', Nil, -1);
  TobUser.LoadDetailFromSQL(SQLUser);

  // Récup des mails du user + ceux dont il peut voir les mails
  // #### pb si plusieurs postes font cette opération pour un même user !
  // #### Risque de doublonner (plusieurs messages pour le même mail) ?
  For iUser:=0 to TobUser.Detail.Count-1 do
    // ---- BOUCLE POUR CHAQUE UTILISATEUR ----
    BEGIN
    UnUser := TobUser.Detail[iUser].GetValue('US_UTILISATEUR');
    NomUser := TobUser.Detail[iUser].GetValue('US_LIBELLE');
    UserEmail := TobUser.Detail[iUser].GetValue('US_EMAIL');
    // (*) on laisse une chance, sur le poste du user en cours, pour que
    // si l'email n'est pas renseigné dans UTILISAT, on puisse lire la registry
    if (UnUser=V_PGI.User) and (UserEmail='') then
      begin
      UserEmail := V_PGI.SmtpFrom;
      // mais uniquement si l'email trouvé n'est pas celui d'un autre user !
      if ExisteSQL('SELECT US_EMAIL FROM UTILISAT WHERE US_EMAIL="'+UserEmail+'"'
       + ' AND US_UTILISATEUR<>"'+V_PGI.User+'"') then Continue;
      end;

    // Mise à jour des orphelins pour ne pas fausser le calcul du max(yms_mailid)
    // (cas d'un transport, qui purgeait YMAILS locale sans faire cette màj)
    ExecuteSQL('UPDATE YMESSAGES SET YMS_MAILID=0, YMS_USERMAIL="" WHERE NOT EXISTS'
     + ' (SELECT 1 FROM YMAILS WHERE YMS_USERMAIL=YMA_UTILISATEUR AND YMS_MAILID=YMA_MAILID)'
     + ' AND YMS_MAILID<>0 AND YMS_USERDEST="'+UnUser+'" AND YMS_USERMAIL="'+UnUser+'"');

    // Dernier n° mail rapatrié pour ce user dans les messages
    iLastMailId := 0;
    Q := OpenSQL('SELECT MAX(YMS_MAILID) FROM YMESSAGES WHERE YMS_USERDEST="'+UnUser+'" AND YMS_USERMAIL="'+UnUser+'"', True);
    if Not Q.Eof then iLastMailId := Q.Fields[0].AsInteger;
    Ferme(Q);

    // Récup. des mails
    // #### en attendant un UsersMailDownload avec passage d'une liste de users
    if AvecDownload then
      // essaye quand même les autres users, mais on considère qu'on aura une erreur
      if not UserMailDownLoad(F, UnUser, TitreMsgError) then
        begin Result:=False; Continue; end;

    // Rq : YMA_UTILISATEUR est vide pour les mails non attribuables
    // $$$ JP 11/10/05 - TOB pour CWAS + filtrer sur les "non lus" uniquement
    TOBMail := TOB.Create ('YMAILS', nil, -1);
    TOBMail.LoadDetailDBFromSQL ('YMAILS', 'SELECT * FROM YMAILS WHERE YMA_UTILISATEUR="'+ UnUser+'"'
       +' AND YMA_MAILID>' + IntToStr (iLastMailId) + ' AND YMA_READED<>"X"');

    // Table à remplir (clé bidon pour ne pas tout charger)
    TobMsg := TOB.Create('YMESSAGES', Nil, -1);
    TobMsg.PutValue('YMS_MSGGUID', '');
    TobMsg.LoadDB;

    // $$$ JP 11/10/05 - TOBMail au lieu de QMail
    for iMail := 0 to TOBMail.Detail.Count-1 do //while not QMail.Eof do
    begin
      booMaj := TRUE;
      // Si le mail est déjà affecté à un message, on passe au suivant
      if ExisteSQL('SELECT YMS_MAILID FROM YMESSAGES WHERE YMS_USERDEST="'+UnUser+'"'
       +' AND YMS_USERMAIL="'+UnUser+'"'
       +' AND YMS_MAILID=' + IntToStr (TOBMail.Detail[iMail].GetInteger ('YMA_MAILID'))) then
        Continue;

      // Valeurs par défaut
      TobMsg.InitValeurs;

      TobMsg.SetString('YMS_MSGTYPE', 'MAI');
      TobMsg.SetString('YMS_USERDEST', UnUser);

      // Clé pour le lien avec YMAILS (nécessite usermail + mailid)
      TobMsg.SetString('YMS_USERMAIL', UnUser);

      TobMsg.SetString('YMS_LIBDEST', NomUser);
      // FQ 11873 - élimine les guillemets dans le sujet
      //TobMsg.SetString('YMS_SUJET', TOBMail.Detail[iMail].GetString ('YMA_SUBJECT'));
      //ymsSujet := TOBMail.Detail[iMail].GetString ('YMA_SUBJECT');
      ymsSujet := TOBMail.Detail[iMail].GetString ('YMA_SUBJECT');
      VireGuillemets(ymsSujet);
      TobMsg.SetString('YMS_SUJET', ymsSujet);

      // Contenu "Texte" du message
      TobMsg.PutValue('YMS_MEMO', TOBMail.Detail[iMail].GetValue ('YMA_MEMO'));

      // Pièces jointes
      if ExisteSQL('SELECT 1 FROM YMAILFILES WHERE YMF_UTILISATEUR="'+UnUser+'"'
      + ' AND YMF_MAILID=' + IntToStr (TOBMail.Detail[iMail].GetInteger ('YMA_MAILID'))
      + ' AND YMF_ATTACHED="X"') then
        TobMsg.SetString('YMS_ATTACHED', 'X');

      // Retraitement préalable sur YMA_FROM (un seul émetteur...)
      email := ExtractEmail (TOBMail.Detail[iMail].GetString ('YMA_FROM'));

      // Plus loin, si on sait décrire plus précisément l'émetteur, on le mettra dans YMS_LIBFROM
      TobMsg.SetString('YMS_LIBFROM', email);
      ymsLibfrom := email;
      // Rq : les mails multi-destinataires sont déjà dupliqués (par le serveur)
      // donc il y autant de YMA_MAILID que de destinataires
      // donc inutile d'éclater les destinataires dans plusieurs ymessages !

      // savoir si le user destinataire est en copie ou non
    { if ( Pos(UserEmail, TOBMail.Detail[iMail].GetString ('YMA_DESTCC')) > 0 )
      and ( Pos(UserEmail, TOBMail.Detail[iMail].GetString ('YMA_DEST')) = 0 ) then }
      if ( ExisteSQL('SELECT 1 FROM YMAILADDRESS WHERE YME_UTILISATEUR="'+UnUser+'"'+
       ' AND YME_MAILID='+TOBMail.Detail[iMail].GetString('YMA_MAILID')+' AND YME_EMAILADDRESS="'+UserEmail+'" AND YME_ADTYPE=3') )
      and ( Not ExisteSQL('SELECT 1 FROM YMAILADDRESS WHERE YME_UTILISATEUR="'+UnUser+'"'+
       ' AND YME_MAILID='+TOBMail.Detail[iMail].GetString('YMA_MAILID')+' AND YME_EMAILADDRESS="'+UserEmail+'" AND YME_ADTYPE=2') )
      then
       TobMsg.SetString('YMS_ENCOPIE', 'X');

      //--- Recherche d'un enreg. annuaire correspondant à l'email de l'émetteur
      Guidper := '';
      if email<>'' then
        begin
        EntreCote(email, False); // => sinon pb en cas de guillemets dans la zone
        Q := OpenSQL('SELECT ANN_GUIDPER, ANN_NOM1, ANN_TEL1 FROM ANNUAIRE WHERE ANN_EMAIL="'+email+'"', True);
        if Not Q.Eof then
          begin
          Guidper := Q.FindField('ANN_GUIDPER').AsString;
          TobMsg.SetString('YMS_LIBFROM', Q.FindField('ANN_NOM1').AsString);
          ymsLibfrom := Q.FindField('ANN_NOM1').AsString;
          TobMsg.SetString('YMS_TEL', Q.FindField('ANN_TEL1').AsString);
          end;
        Ferme(Q);
        // ou recherche d'un contact d'un enreg. annuaire correspondant
        if Guidper='' then
          begin
          (* mcd 12/2005  Q := OpenSQL('SELECT ANI_GUIDPER, ANI_CV, ANI_NOM, ANI_TEL1, ANN_NOM1 FROM ANNUINTERLOC, ANNUAIRE'
           +' WHERE ANI_EMAIL="'+email+'" AND ANI_GUIDPER=ANN_GUIDPER', True); *)
          Q := OpenSQL('SELECT C_GUIDPER, C_CIVILITE, C_NOM, C_TELEPHONE, ANN_NOM1 FROM CONTACT, ANNUAIRE'
           +' WHERE C_RVA="'+email+'" AND C_GUIDPER=ANN_GUIDPER', True);
          if Not Q.Eof then
            begin
            Guidper := Q.FindField('C_GUIDPER').AsString;   //mcd 12/2005
            // 70c = 35c + 17c + 17c
            TobMsg.SetString('YMS_LIBFROM', Q.FindField('ANN_NOM1').AsString + ' - '
               + Q.FindField('C_CIVILITE').AsString + ' ' + Q.FindField('C_NOM').AsString );   //mcd 12/2005
            TobMsg.SetString('YMS_TEL', Q.FindField('C_TELEPHONE').AsString);              //mcd 12/2005
            end;
          Ferme(Q);
          end;
        end;

      //--- Tentative d'affectation du mail directement à un dossier de travail
      nodoss := '';
      if Guidper<>'' then
        begin
        Q := OpenSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_GUIDPER="'+Guidper+'"', True);
        if Not Q.Eof then nodoss := Q.FindField('DOS_NODOSSIER').AsString;
        Ferme(Q);
        // pas directement de dossier, on cherche parmi les intervenants
        if nodoss='' then
          begin
          Q := OpenSQL('SELECT DOS_NODOSSIER FROM DOSSIER, ANNULIEN WHERE ANL_GUIDPERDOS=DOS_GUIDPER'
            + ' AND ANL_GUIDPER="'+Guidper+'"', True, -1, '', True);
          While Not Q.Eof do
            begin
            nodoss := Q.FindField('DOS_NODOSSIER').AsString;
            Q.Next;
            // si intervient sur plusieurs dossiers, on ne sait pas lequel choisir...
            if Not Q.Eof then nodoss := '';
            end;
          Ferme(Q);
          end;
        end;

      //--- Recherche d'un utilisateur interne correspondant à l'email
      if (Guidper='') and (email<>'') then
        begin
        // normalement récup déjà faite par UMailBox :
        TobMsg.SetString('YMS_USERFROM', TOBMail.Detail[iMail].GetString ('YMA_EXPEDITEUR'));
        if TOBMail.Detail[iMail].GetString ('YMA_EXPEDITEUR') = '' then
          begin
          Q := OpenSQL('SELECT US_UTILISATEUR, US_LIBELLE FROM UTILISAT WHERE US_EMAIL="'+email+'"', True);
          if Not Q.Eof then
            begin
            TobMsg.SetString('YMS_USERFROM', Q.FindField('US_UTILISATEUR').AsString);
            TobMsg.SetString('YMS_LIBFROM', Q.FindField('US_LIBELLE').AsString);
            ymsLibfrom := Q.FindField('US_LIBELLE').AsString;
            end;
          Ferme(Q);
          end;
        end;

      // YMA_EXPEDITEUR combo => renseigné si expéditeur interne... (#### PG ?)
      // YMA_DEST string => supposé déjà converti en destinataire dans YMA_UTILISATEUR

      // YMA_STRDATE => dure à convertir en vraie date, autant prendre YMA_DATECREATION
      // $$$ JP 16/01/07: désormais, on tente de convertir en vraie date (heure incorrecte: fuseau horaire pas encore géré)
      try
         TobMsg.SetDateTime ('YMS_DATE', DecodeStrDateDate (TOBMail.Detail [iMail].GetString ('YMA_STRDATE')));
         ymsDate := DecodeStrDateDate (TOBMail.Detail [iMail].GetString ('YMA_STRDATE'));
      except
            // $$$ JP 16/01/07: pour l'instant, pas trop confiance à ce qu'il y a dans YMA_STRDATE...
            TobMsg.SetDateTime('YMS_DATE', TOBMail.Detail[iMail].GetDateTime ('YMA_DATECREATION'));
            ymsDate := TOBMail.Detail[iMail].GetDateTime ('YMA_DATECREATION');
      end;
      TobMsg.SetString('YMS_STRDATE', TobMail.Detail[iMail].GetString('YMA_STRDATE'));

      // Inutile, car mails non lus (récup en cours !)
      // TobMsg.SetString('YMS_LU', QMail.FindField('YMA_READED').AsString);

      // Clés
      // #### faudrait-il garder un guid annuaire, au cas où on identifie (ou on écrit à)
      // une personne de l'annuaire, qui n'est ni un interlocuteur d'un dossier,
      // ni un intervenant sur un dossier ? (exple : annuaire préalimenté, autre site...)
      // TobMsg.SetInteger('YMS_GUIDPER', guidper);
      TobMsg.SetString('YMS_NODOSSIER', nodoss);
      TobMsg.SetInteger('YMS_MAILID', TOBMail.Detail[iMail].GetInteger ('YMA_MAILID'));

      sNewMsgGUID := AglGetGUID();
      TobMsg.SetString ('YMS_MSGGUID',sNewMsgGUID);

      // GUH. Demande KMPG 07/10/07 BGM-22
      // Messagerie : doublons - Alerte si le message a déjà été archivé
      SQL := 'SELECT 1 FROM YMESSAGES'+
            ' WHERE YMS_DATE = "'+USDATETIME(YMSDATE)+'"'+
            ' AND YMS_LIBFROM = "'+YMSLIBFROM+'"'+
            ' AND YMS_SUJET = "'+YMSSUJET+'"'+
            ' AND YMS_USERMAIL = "'+UNUSER+'"';

      YMSLIBFROM := TOBMail.Detail[iMail].GetString('YMA_FROM');
      YMSLIBFROM := Copy(YMSLIBFROM,1,pos('<',YMSLIBFROM)-2);

      sAsk :='Le message de "'+YMSLIBFROM+'" [Objet:'+YMSSUJET+'],'+#13#10+
             ' en date du '+ FormatDateTime('dd/mm/yyyy hh:mm:ss',YMSDATE)+
             ' a déjà été importé.'+#13#10+
             ' Souhaitez-vous l''importer une nouvelle fois ?';

      if (ExisteSQL(SQL)) and (PGIAsk(sAsk,f.Caption)=mrNo) then
        booMaj := FALSE;

      if booMaj then
      begin
        TobMsg.InsertDB(Nil);

        // Recopie des pièces jointes + contenu eml/html
        ExecuteSQL('INSERT YMSGFILES (YMG_MSGGUID, YMG_FILEGUID, YMG_ATTACHED) SELECT "'+sNewMsgGUID
         + '" , YMF_FILEGUID, YMF_ATTACHED FROM YMAILFILES WHERE YMF_UTILISATEUR="'+UnUser+'"'
         + ' AND YMF_MAILID=' + IntToStr (TOBMail.Detail[iMail].GetInteger ('YMA_MAILID')) );
        // MD 28/02/06 - Rajout du contenu eml/html => donc plus de + ' AND YMF_ATTACHED="X"');

        //------------------------------------------------------------
        //--- Alimentation de YMSGADDRESS à partir de YMAILADDRESS
        //------------------------------------------------------------
        // Recopie des destinataires
        ExecuteSQL('INSERT YMSGADDRESS (YMR_MSGGUID, YMR_ADTYPE, YMR_ADORDRE, YMR_EMAILADDRESS, YMR_EMAILNAME)'
         + ' SELECT "'+sNewMsgGUID+'", YME_ADTYPE, YME_ADORDRE, YME_EMAILADDRESS, YME_EMAILNAME'
         + ' FROM YMAILADDRESS WHERE YME_UTILISATEUR="'+UnUser+'"'
         + ' AND YME_MAILID=' + TOBMail.Detail[iMail].GetString ('YMA_MAILID'));

        // #### ICI ON POURRAIT DETRUIRE LE MAIL UNE FOIS TRAITE ####
        // $$$ JP 11/10/05 - En fait, on marque le mail comme "lu" par l'appli ("lu" veut dire "récupéré", si on veut)
        TOBMail.Detail[iMail].SetBoolean ('YMA_READED', TRUE);
        TOBMail.Detail[iMail].UpdateDB;
      end
      else
      begin
        // suppression des enregs des tables YMAILS / YMAILADDRESS / YMAILFILES.
        // si l'utilisateur ne veut pas réimporter le message.
        ExecuteSql('DELETE FROM YMAILADDRESS'+
                   ' WHERE YME_UTILISATEUR="'+UnUser+'"'+
                     ' AND YME_MAILID='+ IntToStr(TOBMail.Detail[iMail].GetInteger('YMA_MAILID')));

        ExecuteSql('DELETE FROM YMAILFILES'+
                   ' WHERE YMF_UTILISATEUR="'+UnUser+'"'+
                     ' AND YMF_MAILID='+ IntToStr(TOBMail.Detail[iMail].GetInteger('YMA_MAILID')));

        ExecuteSql('DELETE FROM YMAILS'+
                   ' WHERE YMA_UTILISATEUR="'+UnUser+'"'+
                     ' AND YMA_MAILID='+ IntToStr(TOBMail.Detail[iMail].GetInteger('YMA_MAILID')));
      end;

      // Mail suivant
    end;
    FreeAndNil(TobMsg);
    FreeAndNil(TOBMail);
    //Ferme(QMail);
    END;
    // ---- FIN BOUCLE POUR CHAQUE UTILISATEUR ----

  TobUser.Free;
end;


{$IFDEF BUREAU}
{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /
Description .. : Affiche la saisie des droits des utilisateurs sur les mails dans
Suite ........ : une hiérarchie
Mots clefs ... : HIERARCHIQUE;DROITS
*****************************************************************}
procedure ShowDroitsMailsParUsers(PRien: THPanel; HDTDrawNode: TOnHDTDrawNode);
var F: TFYTreeLinks;
begin
  // pour administrer une tablette hiérarchique désignée
  F := ShowTreeLinks(PRien, 'Gestion des mails par utilisateurs',
            'YYUSERMASTERSLAVE', HDTDrawNode);
  // Customisation
  F.TSDatas.Caption := 'Utilisateurs';
  F.Panel1.Caption := 'Droits de visualisation sur :';
  F.Panel2.Caption := 'Autres utilisateurs';
  F.SBAddDisp.Visible := False;  // Ajouter un élément disponible
  F.SBEditDisp.Visible := False; // Modifier un élément disponible
  F.SBUpChilds.Visible := False; // Ajouter des éléments disponibles à tout le niveau
  F.SBDownChilds.Visible := False; // Supprimer des éléments liés à tout le niveau
  F.btnMainNormalize.Visible := False; // Normaliser les liaisons hiérarchiques
end;
{$ENDIF BUREAU}


{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /    
Description .. : Retourne le code du premier utilisateur trouvé ayant des
Suite ........ : paramètres de messagerie définis.
Mots clefs ... : 
*****************************************************************}
function PremierUserAvecMessagerie: String;
var Q: TQuery;
begin
  Result := '';
  Q := OpenSQL('SELECT US_UTILISATEUR FROM UTILISAT WHERE US_EMAIL<>""'
   +' AND US_EMAILPOPSERVER<>""', True);
  if Not Q.Eof then Result := Q.FindField('US_UTILISATEUR').AsString;
  Ferme(Q);
end;

function IsRtf(Texte: String): Boolean;
begin
  Result := False;
  // cf GetRTFStringText dans RtfCounter
  if (length(Texte)>5)and(copy(Texte,1,5)='{\rtf')then Result := True;
end;


{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 15/03/2005
Modifié le ... :   /  /
Description .. : Complète la liste des destinataires si le
Suite ........ : decodemimeaddress a tronqué des infos pertinentes...
Suite ........ : (en attendant améliorations PGR)
Suite ........ : Il manquera de toutes façons les cas de plusieurs 
Suite ........ : adresses sans délimiteur < ou ", mais dont certaines sont
Suite ........ : correctes et d'autres bizarres.
Mots clefs ... : DECODAGE;MIMEADDRESS
*****************************************************************}
procedure VerifDecodageMimeAddress(LstDest: TStringList; Address: String);
var tmp: String;
    c : Char;
    i : Integer;
    collecte : Boolean;
    var TabAutorizedChar: set of char;

begin
  if (LstDest=Nil) or (Address='') {or (LstDest.Count>0) (*)} then exit;
  // (*) en fait on tente toujours d'améliorer le décodage, même si on a déjà
  // une adresse dans LstDest, car souvent on en a récupéré une seule au lieu de n
  // (bug de DecodeMimeAddress)

  // Caractères autorisés
  TabAutorizedChar := ['.', '@', '_', '-', '0'..'9', 'A'..'Z', 'a'..'z'];

  tmp := '';
  collecte := True;

  for i := 1 to Length(Address) do
    begin
    c := Address[i];

    // Séparateur : on publie la chaine collectée
    if (c=',') or (c=';') or (c='"') or (c='''') or (c='<') or (c='>') then
      begin
      tmp := Trim(tmp);
      // uniquement si c'est un email (@), et non encore publié
      if (tmp<>'') and (Pos('@', tmp)>0) and (LstDest.IndexOf(tmp)=-1) then LstDest.Add(tmp);
      // on reprend pour une nouvelle collecte
      tmp := '';
      collecte := True;
      end

    // Collecte stoppée si caractère hors des valeurs autorisées
    else if Not (c in TabAutorizedChar) then
      collecte := False

    // En cours de collecte
    else if collecte then
      tmp := tmp + c;

    end;

  // Résidus : on ne prend pas car c'est souvent une adresse incomplète et tronquée !
  {if collecte then
    begin
    tmp := Trim(tmp);
    if (tmp<>'') and (Pos('@', tmp)>0) and (LstDest.IndexOf(tmp)=-1) then LstDest.Add(tmp);
    end;
  }

end;

procedure VireGuillemets(var str: String);
var tmp : String;
    i : Integer;
begin
  tmp := '';
  for i := 1 to Length(str) do
    begin
    if Copy(str, i, 1)<> '"' then tmp := tmp + Copy(str, i, 1);
    end;
  str := tmp;
end;

end.

