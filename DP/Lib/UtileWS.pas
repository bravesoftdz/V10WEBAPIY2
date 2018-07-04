{***********UNITE*************************************************
Auteur  ...... : M. DESGOUTTE
Créé le ...... : 04/10/2004
Modifié le ... :   /  /
Description .. : Gestion de l'API eWS pour les appels
Suite ........ : à la publication documentaire
Suite ........ : 19/10/2004 : Compatibilité appels depuis dossier
Suite ........ :              (GetParamsocDP, ##DP## ...)
Suite ........ : 13/12/2004 : L'identifiant collaborateur passé à
Suite ........ :              eWs est le UserLogin au lieu du User
Mots clefs ... : EWS;GED;PUBLICATION
*****************************************************************}
unit UtileWS;

interface

uses Windows, Registry, SysUtils, ShellApi, ParamSoc, classes, menus,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}
     {$ifndef EAGLCLIENT}
     uDbxDataSet,
     {$ENDIF}
     {$ENDIF}
     Controls,
     HEnt1, HMsgBox, HCtrls, UTob, LicUtil,
     ExploitEWS_TLB, // API eWS
     UGedFiles, UtilGed;


function  EwsLibrairieOk(bMsgOnBadInstall : Boolean = TRUE): Boolean;
function  AgenteWsLibrairieOk:boolean;
procedure EwsLibereCom;

procedure EwsLanceBackOffice;
procedure EwsHistoriquePublications (NoDossier, Utilisateur: String);
function  EwsVerifOuCreeCollab      (TOBUser:TOB=nil; bMsgOnBadInstall:boolean=TRUE): boolean; //function  EwsVerifOuCreeCollab(bMsgOnBadInstall : Boolean = TRUE): Boolean;
function  EwsSupprimeCollab         (UserCode, UserLogin: String): Boolean;
function  EwsModifieCollab          (UserLogin, NewPassword, NewNom, NewEmail: String): Boolean;
function  EwsCreeClient             (NoDossier: String): Boolean;
function  EwsModifProfilClient      (NoDossier: String; var bHasChanged: Boolean; var NewUsrS1, NewPwdS1: String): Boolean;
procedure EwsConnect                (NoDossier:string); // $$$ JP 18/08/06
function  EwsModifieClient          (NoDossier, LibDossier, EmailDossier: String): Boolean;
function  EwsSupprimeClient         (NoDossier: String): Boolean;
function  EwsSelectionneNoeud       (var IdEws, LibEws: String): Boolean;
procedure EwsListeRegle             (const lstRegles:TStringList); // $$$ JP 17/08/06
//procedure EwsSelectionneRegle        (var IdEwsRegle, LibEwsRegle:string; pm:TPopupMenu; pmClick:TNotifyEvent);
function  EwsRetourneLibelleNoeud   (IdEws: String): String;
function  EwsRetourneServiceClient  (NoDossier, FamilleService: String ; var IdClient, PwdClient, IdService: String): Integer;
function  EwsPublieDocument         (TGN:TGedDPNode; bWithIhm:boolean): Boolean;
function  EwsPublieUnDocument       (SDocGUID: String; EwsId,EwsRegle:string; var Retour:string; bForceRepub:boolean=True; bWithIhm:boolean=FALSE): Boolean;
function  EwsFichePublicationMemo   : Boolean;
function  EwsVerificationSIO        (SIO, SIOPassword: String): Integer;
procedure EwsGetInfosDossier        (NoDossier: String; var LibDossier, EmailDossier, UsrS1, PwdS1: String);
procedure EwsSynchronize            (bAskConfirm:boolean=TRUE);

function  IsDossierEws              (NoDossier: String): Boolean;
function  GetNumeroSIO : Boolean ;

// $$$ JP 11/10/06
function  GetEwsPath:string;
function  GetEwsVersion:integer; // de la forme: n° version majeur*1000 + n° version mineur (dans le n° de version std: x.y.z.b, x=n° version majeur et y=n° version mineur)

// $$$ JP 13/12/06 - connaitre le mot de passe de l'utilisateur "ADM"
function  GetAdmPwd:string;

var V_eWS      :IGeneral;    // Objet com pour appel fonctions eWS
    V_eWSAgent :IAgenteWS;   // Objet com pour appel fonctions d'agent eWS

const
  //--- Concepts bureau (Applications hébergées)
  ccParamApplisHebergees       = 187200 ;
  ccActiverDossierAsp          = 187205 ;
  ccBusinessLineDepuisSynthese = 187210 ;
  ccEwsDepuisSynthese          = 187215 ;
  ccPubliEwsDepuisSynthese     = 187220 ;
  ccOutilsAssistanceAsp        = 187225 ;
  ccGedAssociationBrancheEws   = 187230 ;

/////////////// IMPLEMENTATION ///////////////
implementation

uses
    forms, activex, comobj, galSystem;

var
   sNumeroSIO   :string;
   sAdmPwd      :string; // stocke le mot de passe de "ADM"

//---------------------------------------------------
//--- #### issue de InitialiserMp de galTomDossier
//--- voir si on peut factoriser...
//---------------------------------------------------
function QuatreDerniersCarac (ChNoDossier : String) : String;
var Longueur : Integer;
begin
 Result:='';
 Longueur:=length (ChNoDossier);
 if (Longueur<=4) then
  Result:=ChNoDossier
 else
  Result:=copy (ChNoDossier,Longueur-3,4);
end;

function  EwsLibrairieOk(bMsgOnBadInstall : Boolean = TRUE): Boolean;
// Test la présence et l'accès aux objets de la librairie Com
begin
  if V_eWS=Nil then
  begin
      SourisSablier;
      try
      try
         // Interface pour appels objets COM eWS
         V_eWS := CoGeneral.Create as IGeneral;
      except
        if bMsgOnBadInstall then
           PGIInfo('eWS ne peut pas être appelé de votre poste, l''application eWS n''est pas ou est mal installée.');
        V_eWS := Nil;
      end;
      finally
        SourisNormale;
      end;
  end;

  Result := (V_eWS<>Nil);
end;

function AgenteWSLibrairieOk : boolean;
begin
  if V_eWSAgent=Nil then
  begin
       SourisSablier;
      try
      try
         // Interface pour appels objets COM eWS
         V_eWSAgent := CoAgenteWS.Create as IAgenteWS;
      except
           PGIInfo('eWS ne peut pas être appelé de votre poste, l''application eWS n''est pas ou est mal installée.');
           V_eWSAgent := Nil;
      end;
      finally
        SourisNormale;
      end;
  end;

  Result := (V_eWSAgent<>Nil);
end;

procedure EwsLibereCom;
begin
  V_eWS := nil;
  V_eWSAgent := nil;
end;

function GetEwsPath:string;
var
   Reg  :TRegistry;
begin
     // Par défaut, inconnu
     Result := '';

     // Lecture emplacement de la dll à vérifier
     Reg := TRegistry.Create;
     try
        Reg.RootKey := HKEY_CURRENT_USER;
        if Reg.OpenKey ('\Software\CCMX\CCMX eWS', FALSE) then
             Result := Reg.ReadString ('Chemin')
        else
        begin
             Reg.RootKey := HKEY_LOCAL_MACHINE;
             if Reg.OpenKey ('\Software\CCMX\CCMX eWS', FALSE) then
                Result := Reg.ReadString ('Chemin');
        end;
     finally
            Reg.Free;
     end;
end;

procedure EwsLanceBackOffice;
var
 // Reg : TRegistry;
    PathEws : String;
begin
     // $$$ JP 11/10/06: fonction GetEwsPath
     PathEws := GetEwsPath;

     {PathEws := '';
  // Pas moyen d'utiliser GetFromRegistry (de Hent1) : ça renvoie jamais rien !!
  Reg := TRegistry.Create;

  // $$$ JP 05/09/06: FQ 10025
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey ('\Software\CCMX\CCMX eWS', FALSE) then
      PathEws := Reg.ReadString ('Chemin')
  else
  begin
       Reg.RootKey := HKEY_LOCAL_MACHINE;
       if Reg.OpenKey ('\Software\CCMX\CCMX eWS', FALSE) then
          PathEws := Reg.ReadString ('Chemin');
  end;}

{  if V_PGI.WinNT then Reg.RootKey := HKEY_CURRENT_USER
  else Reg.RootKey := HKEY_LOCAL_MACHINE;
  if Reg.OpenKey('\Software\CCMX\CCMX eWS', False) then PathEws := Reg.ReadString('Chemin');
  Reg.CloseKey;
  if PathEws='' then
  begin
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\Software\CCMX\CCMX eWS', False) then PathEws := Reg.ReadString('Chemin');
    Reg.CloseKey;
  end;
  Reg.Free;}

  // Lancement du Back Office eWS
  if (PathEws<>'') and FileExists (PathEws+'\ews.exe') then
      ShellExecute( 0, PCHAR('open'), PChar(PathEws+'\ews.exe'), nil, PChar(PathEws), SW_RESTORE)
  else
      MessageAlerte('Le programme eWS n''a pas été trouvé sur votre configuration.');
end;


procedure EwsHistoriquePublications(NoDossier, Utilisateur : String) ;
// Historique des publications du collaborateur
// => si NoDossier='',  c'est tous clients confondus
// => si Utilisateur='', c'est tous collaborateurs confondus
begin
  if Not EwsVerifOuCreeCollab then exit;

  try
    // donc on ne précise pas le client_id
    V_eWS.FichePublications(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), True, NoDossier, Utilisateur) ;
  except
    on E:Exception do
      if V_PGI.SAV then MessageAlerte('Erreur appel FichePublications : ' + #13#10 + E.Message );
  end;

end;

function EwsVerifOuCreeCollab (TOBUser:TOB; bMsgOnBadInstall:boolean):boolean;
var
    Q : TQuery;
    SQL : String;
    strUser, strLogin, strPwd, strLibelle, strEmail  :string;
begin
    Result := False;

    if GetNumeroSIO=False then exit;

    if Not EwsLibrairieOk(bMsgOnBadInstall) then exit;

    // $$$ JP 13/12/06: si TOB fournie, on prends les valeurs dans cette TOB, sinon dans V_PGI
    if TOBUser <> nil then
    begin
         strUser    := TOBUser.GetString ('US_UTILISATEUR');
         strLogin   := TOBUser.GetString ('US_ABREGE');
         strPwd     := DecryptageSt (TOBUser.GetString ('US_PASSWORD'));
         strLibelle := TOBUser.GetString ('US_LIBELLE');
         strEmail   := TOBUser.GetString ('US_EMAIL');
    end
    else
    begin
         strUser    := V_PGI.User;
         strLogin   := V_PGI.UserLogin;
         strPwd     := DecryptageSt (V_PGI.Password);
         strLibelle := V_PGI.UserName;
         Q          := OpenSQL ('SELECT US_EMAIL FROM UTILISAT WHERE US_UTILISATEUR="' + strUser + '"', True);
         if not Q.Eof then
            strEmail := Q.Fields[0].AsString;
         Ferme(Q);
    end;

    // - tablette YYEWSUTILISAT : trace les collab déjà créés dans eWS
    // (il suffit que la ligne existe dans la tablette, qqs la valeur de CC_ABREGE) le test interne à PGI est plus économique que les appels COM !
    Result := ExisteSQL('SELECT 1 FROM ##DP##.CHOIXCOD WHERE CC_CODE="' + strUser + '" AND CC_TYPE="YEW"');

    // tente création dans eWS
    if Not Result then
    begin
         // $$$ JP 14/12/06: si c'est ADM qu'on tente de créer, il faut d'abord essayer de modifier son mdp, si pas possible, on tente de le créer
         if strLogin = 'ADM' then
            try
               if V_eWS.ModificationCollaborateur (sNumeroSIO, 'ADM', GetAdmPwd, 'ADM', 'ADM', strPwd, strLibelle, '', strEmail) = TRUE then
               begin
                    // On retient le mot de passe modifié
                    sAdmPwd := strPwd;
                    Result  := TRUE;
               end;
            except
               on E:Exception do
                  if V_PGI.SAV then PGIInfo('Erreur appel ModificationCollaborateur (''ADM''):' + #13#10' ' + E.Message);
            end;

         // Si toujours pas réussi (user=ADM), on tente de créer cet utilisateur
         if Result = FALSE then
         begin
              // La création nécessite un compte administrateur : ADM
              // Rq : pas de prénom dans la fiche utilisateur => on passe le US_ABREGE (login)
              // #### Demande à FB 7/10/04 pour que la création se fasse avec tous les droits
              // $$$ JP 13/12/06 - mot de passe ADM + utilisation strXXX au lieu de V_PGI.XXX
              try
                 Result := V_eWS.CreationCollaborateur (sNumeroSIO, 'ADM', GetAdmPwd {'cegid'}, strLogin, strLogin, strPwd, strLibelle, '', strEmail);

                 // si CreationCollaborateur renvoie False sur un user autre que ADM, ça peut être qu'il existait déjà, donc on continue quand même
                 if strLogin <> 'ADM' then
                    Result := TRUE;
              except
                 on E:Exception do
                    if V_PGI.SAV then PGIInfo('Erreur appel CreationCollaborateur (''' + strLogin + ''') :' + #13#10' ' + E.Message);
              end;
         end;

         // Mémorise dans PGI que "collaborateur eWS créé"
         if Result = TRUE then
         begin
              if V_PGI.DefaultSectionDBName = '' then // ou V_PGI.RunFromLanceur
                  SQL := 'INSERT INTO CHOIXCOD'
              else
                  SQL := 'INSERT INTO '+V_PGI.DefaultSectionDBName+'.dbo.CHOIXCOD';
              SQL := SQL + ' (CC_TYPE, CC_CODE, CC_ABREGE, CC_LIBELLE) VALUES ("YEW", "' + strUser + '", "' + strLogin + '", "' + strLibelle + '")';
              ExecuteSQL(SQL);
         end;
    end;
end;


function  EwsSupprimeCollab (UserCode, UserLogin:string):boolean;
var
    SQL : String;
begin
  Result := False;
  if GetNumeroSIO=False then exit;

  if Not EwsLibrairieOk then exit;

  // $$$ JP 13/12/06: si c'est ADM qui doit être supprimé, on ne supprime pas dans Ews, dangereux (il faudra le faire à la main dans ews)
  if UserLogin <> 'ADM' then
  begin
       try
          Result := V_eWS.SuppressionCollaborateur (sNumeroSIO, 'ADM', GetAdmPwd {'cegid'}, UserLogin);
       except
          on E:Exception do
             if V_PGI.SAV then MessageAlerte('Erreur appel SuppressionCollaborateur (''' + UserLogin + ''') :' + #13#10' ' + E.Message);
       end;
  end
  else
  begin
       // On reprend le mot de passe par défaut de l'utilisateur ADM spécifique Ews
       sAdmPwd := 'cegid';
       Result := TRUE;
  end;
{  else
  begin
       // $$$ JP 13/12/06: mot de passe ADM
       try
          if V_eWS.ModificationCollaborateur (sNumeroSIO, 'ADM', GetAdmPwd, 'ADM', 'ADM', 'cegid', 'ADM', '', '') = TRUE then
          begin
               sAdmPwd := 'cegid';
               Result  := TRUE;
          end;
       except
          on E:Exception do
             if V_PGI.SAV then MessageAlerte('Erreur appel ModificationCollaborateur (''ADM'') :' + #13#10' ' + E.Message);
       end;
  end;}

  // $$$ JP 14/12/06: suppression de l'utilisateur "ews" dans Pgi (si suppression ou modification ok dans ews)
  if Result = TRUE then
  begin
      if V_PGI.DefaultSectionDBName='' then // ou V_PGI.RunFromLanceur
          SQL := 'DELETE FROM CHOIXCOD'
      else
          SQL := 'DELETE FROM '+V_PGI.DefaultSectionDBName+'.dbo.CHOIXCOD';
      SQL := SQL + ' WHERE CC_TYPE="YEW" AND CC_CODE="'+UserCode+'"';
      ExecuteSQL(SQL);
  end;
end;


function  EwsModifieCollab(UserLogin, NewPassword, NewNom, NewEmail: String): Boolean;
begin
  Result := False;
  if GetNumeroSIO=False then exit;

  if Not EwsLibrairieOk then exit;

  try
    if V_eWS.ModificationCollaborateur(sNumeroSIO, 'ADM', GetAdmPwd {'cegid'}, UserLogin, UserLogin, NewPassword, NewNom, '', NewEmail) then
    begin
         Result := True;

         // $$$ JP 13/12/06: si utilisateur concerné est "ADM", il faut màj le mot de passe pour les accès ews
         if UserLogin = 'ADM' then
            sAdmPwd := NewPassword;
    end;
  except
       on E:Exception do
          if V_PGI.SAV then MessageAlerte('Erreur appel ModificationCollaborateur (''' + UserLogin + ''') :' + #13#10' ' + E.Message);
  end;
end;


function  EwsCreeClient(NoDossier : String) : Boolean;
var
    Q : TQuery;
    SGuidPer : String;
    LibDossier, EmailDossier, UsrS1, PwdS1 : String;
    WService, WIdClient, WPwdClient : WideString;
    ret : Integer;
begin
  Result := False;
  if Not EwsVerifOuCreeCollab then exit;

  // Déjà créé
  // (on teste le flag interne à PGI car les appels aux objets Com sont assez lourds !)
  if IsDossierEws(NoDossier) then
    begin
    Result := True;
    exit;
    end;

  // Récup infos du dossier
  SGuidPer := '';
  Q := OpenSQL('SELECT DOS_GUIDPER, DOS_LIBELLE FROM DOSSIER WHERE DOS_NODOSSIER="'+NoDossier+'"', True);
  if Not Q.Eof then
    begin
    SGuidPer := Q.FindField('DOS_GUIDPER').AsString;
    LibDossier := Q.FindField('DOS_LIBELLE').AsString;
    end;
  Ferme(Q);
  if SGuidPer='' then
    begin
    if V_PGI.SAV then PGIInfo('Le dossier '+NoDossier+' n''a pas de fiche annuaire.');
    exit;
    end;

  // Récup email dossier en cours
  Q := OpenSQL('SELECT ANN_EMAIL FROM ANNUAIRE WHERE ANN_GUIDPER="'+SGuidPer+'"', True);
  if Not Q.Eof then EmailDossier := Q.Fields[0].AsString;
  Ferme(Q);
  // Récup user / password éventuellement utilisés pour S1
  Q := OpenSQL('SELECT DOS_USRS1, DOS_PWDS1, DOS_EWSCREE FROM DOSSIER WHERE DOS_NODOSSIER="'+NoDossier+'"', True);
  if Not Q.Eof then
    begin
    UsrS1 := Q.FindField('DOS_USRS1').AsString;
    PwdS1 := Q.FindField('DOS_PWDS1').AsString;
    end;
  Ferme(Q);
  if UsrS1='' then
    begin
    UsrS1 := NoDossier;
    PwdS1 := QuatreDerniersCarac(UsrS1);
    // Maj des valeurs par défaut, même si après on abouti pas...
    ExecuteSQL('UPDATE DOSSIER SET DOS_USRS1="'+UsrS1+'", DOS_PWDS1="'+PwdS1+'" WHERE DOS_NODOSSIER="'+NoDossier+'"');
    end;

  // Création du client
  if PGIAsk('Vous allez activer eWS sur le dossier '+LibDossier+' ('+NoDossier+').'+#13#10
   + ' Confirmez-vous ?') = mrYes then
     begin
     try
       // on aurait pu passer un compte ADM/ccmx mais autant savoir QUI (quel collab) crèe ce client
       if V_eWS.FicheClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier, NoDossier, LibDossier, EmailDossier, UsrS1, PwdS1) then
         begin;
         // Récupération du user/mot de passe S1 éventuellement modifié par l'utilisateur
         // via une interrogation du service (#### car non renvoyés dans l'API FicheClient)
         // en passant une famille '' pour être sûr d'avoir retour ok
         ret := V_eWS.RetourneServiceClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier, '', WService, WIdClient, WPwdClient);
         // si échec de récupération avec une famille vide, essai avec la famille BL
         if ret=3 then
              ret := V_eWS.RetourneServiceClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier, 'S1', WService, WIdClient, WPwdClient);
         if ret=0 then
         begin
              UsrS1  := WIdClient;
              PwdS1  := WPwdClient;
         end
         else
         begin
              if V_PGI.SAV then PGIInfo('Echec de récupération du mot de passe (RetourneServiceClient='+IntToStr(ret)+')');
         end;

         // le contact cabinet du client sera le collaborateur qui active le client
         V_eWS.CreationContact(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), V_PGI.UserLogin, NoDossier);
         Result := True;
         ExecuteSQL('UPDATE DOSSIER SET DOS_EWSCREE="X", DOS_USRS1="'+UsrS1+'", DOS_PWDS1="'+PwdS1+'" WHERE DOS_NODOSSIER="'+NoDossier+'"');

         // $$$ JP 21/07/06: synchro immédiate
         EwsSynchronize (FALSE);
         end;
     except
       on E:Exception do
         if V_PGI.SAV then MessageAlerte('Erreur appel CreationClient : ' + #13#10 + E.Message );
       end;
     end;
end;


function EwsModifProfilClient(NoDossier: String; var bHasChanged: Boolean; var NewUsrS1, NewPwdS1: String): Boolean;
// MD 21/09/07 - Rajout param bHasChanged pour savoir si il faut propager une modif de mot de passe sur BL
var
    LibDossier, EmailDossier, OldUsrS1, OldPwdS1, TmpUsrS1, TmpPwdS1 : String;
    WService, WIdClient, WPwdClient : WideString;
    ret : Integer;
begin
  Result := False;
  if Not EwsVerifOuCreeCollab then exit;

  if Not IsDossierEws(NoDossier) then exit;

  EwsGetInfosDossier(NoDossier, LibDossier, EmailDossier, OldUsrS1, OldPwdS1);
  NewUsrS1 := OldUsrS1; NewPwdS1 := OldPwdS1;

  // Accès à la fiche du client pour modif éventuelle de son profil eWS et de son user/password
  try
    if V_eWS.FicheClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier, NoDossier, LibDossier, EmailDossier, TmpUsrS1, TmpPwdS1) then
    begin
      Result := True;
      // Récupération du user/mot de passe S1 éventuellement modifié par l'utilisateur
      // via une interrogation du service (#### car non renvoyés dans l'API FicheClient)
      // en passant une famille '' pour être sûr d'avoir retour ok
      ret := V_eWS.RetourneServiceClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier, '', WService, WIdClient, WPwdClient);
      // si échec de récupération avec une famille vide, essai avec la famille BL
      if ret=3 then
           ret := V_eWS.RetourneServiceClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier, 'S1', WService, WIdClient, WPwdClient);
      if ret=0 then
      begin
           NewUsrS1  := WIdClient;
           NewPwdS1  := WPwdClient;
           // Mise à jour
           if (NewUsrS1<>OldUsrS1) or (NewPwdS1<>OldPwdS1) then
           begin
                bHasChanged := True;
                ExecuteSQL('UPDATE DOSSIER SET DOS_USRS1="'+NewUsrS1+'", DOS_PWDS1="'+NewPwdS1+'" WHERE DOS_NODOSSIER="'+NoDossier+'"');
           end;
      end
      else
      begin
           if V_PGI.SAV then PGIInfo('Echec de récupération du mot de passe (RetourneServiceClient='+IntToStr(ret)+')');
      end;

      // $$$ JP 21/07/06: synchro immédiate
      EwsSynchronize (FALSE);

    end;
  except
    on E:Exception do
      if V_PGI.SAV then MessageAlerte('Erreur appel FicheClient : ' + #13#10 + E.Message );
  end;
end;

procedure EwsConnect (NoDossier:string);
begin
  if not EwsVerifOuCreeCollab then exit;

  try
     V_eWS.ConnexionEWS (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password), NoDossier);
   except
     on E:Exception do
       if V_PGI.SAV then MessageAlerte('Erreur appel ConnexionEWS : ' + #13#10 + E.Message );
   end;
   
end;

function  EwsModifieClient(NoDossier, LibDossier, EmailDossier : String): Boolean;
begin
  Result := False;
  if Not EwsVerifOuCreeCollab then exit;

   try
     if V_eWS.ModificationClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier, NoDossier, LibDossier, EmailDossier) then
        begin
        Result := True;
        // $$$ JP 21/07/06: synchro immédiate
        EwsSynchronize (FALSE);
        end;
   except
     on E:Exception do
       if V_PGI.SAV then MessageAlerte('Erreur appel ModificationClient : ' + #13#10 + E.Message );
   end;
end;


function  EwsSupprimeClient(NoDossier : String): Boolean;
begin
  Result := False;
  if Not EwsVerifOuCreeCollab then exit;

  try
    V_eWS.SuppressionClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier) ;
    Result := True;
    // $$$ JP 21/07/06: synchro immédiate
    EwsSynchronize (FALSE);
  except
    on E:Exception do
      if V_PGI.SAV then MessageAlerte('Erreur appel SuppressionClient : ' + #13#10 + E.Message );
  end;
end;


function EwsSelectionneNoeud(var IdEws, LibEws: String): Boolean;
var
    WIdEws : WideString;
    RetourIdEws : String;
begin
  Result := False;
  if Not EwsVerifOuCreeCollab then exit;

  WIdEws := IdEws;
  try
    // Retourne Id;Libellé
    V_eWS.FicheSelectionNoeudPublication(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), WIdEws);
    Result := True;
    RetourIdEws := WIdEws;
    // donc on découpe
    IdEws := ReadTokenSt(RetourIdEws);
    LibEws := ReadTokenSt(RetourIdEws);
  except
    on E:Exception do
      if V_PGI.SAV then PGIInfo('Erreur appel FicheSelectionNoeudPublication : ' + #13#10 + E.Message );
  end;
end;

// $$$ JP 17/08/06: récupération code profil (code règle) de publication
procedure EwsListeRegle (const lstRegles:TStringList); // $$$ JP 17/08/06
var
   pSafe        :PSafeArray;
   lNbDim       :integer;
   i            :integer;
   iLow, iHigh  :integer;
   pElt         :PWideChar;
   strValue     :string;
begin
  if Not EwsVerifOuCreeCollab then exit;

  try
     pSafe   := V_eWS.ListeProfilPublication (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.PassWord));
     if pSafe <> nil then
     begin
          lNbDim := SafeArrayGetDim (pSafe);
          if lNbDim = 1 then
          begin
               SafeArrayGetLBound (pSafe, 1, iLow);
               SafeArrayGetUBound (pSafe, 1, iHigh);
               for i := iLow to iHigh do
               begin
                    SafeArrayGetElement (pSafe, i, pElt);
                    strValue := WideCharToString (pElt);
                    lstRegles.Add (strValue);
               end;
          end;
     end;
  except
        on E:Exception do if V_PGI.SAV then PGIInfo ('Erreur appel ListeProfilPublication : ' + #13#10 + E.Message);
  end;
end;


function  EwsRetourneLibelleNoeud(IdEws : String): String;
var
    WIdEws, WLibEws : WideString;
begin
  Result := '';
  if Not EwsVerifOuCreeCollab then exit;

  WIdEws := IdEws;
  try
    // Retourne Libellé détaillé si option 3
    V_eWS.RetourneLibelleNoeudPublication(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password),
      WIdEws, 3, WLibEws);
    Result := WLibEws;
  except
    on E:Exception do
      if V_PGI.SAV then PGIInfo('Erreur appel RetourneLibelleNoeudPublication : ' + #13#10 + E.Message );
  end;
end;


function  EwsRetourneServiceClient(NoDossier, FamilleService : String ; var IdClient, PwdClient, IdService : String): Integer;
// Récupère le premier service eWS de la FamilleService disponible
var
    WService, WIdClient, WPwdClient : WideString;
begin
  IdService := '';
  Result := -1;
  if Not EwsVerifOuCreeCollab then exit;

  try
    Result := V_eWS.RetourneServiceClient(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password), NoDossier,
      FamilleService, WService, WIdClient, WPwdClient);

   if Result=0 then
      begin
      IdService := WService;
      IdClient  := WIdClient;
      PwdClient := WPwdClient;
      end
    else if Result=1 then
      begin
      if V_PGI.SAV then PGIInfo('Le client '+NoDossier+' n''a pas été activé sur eWS');
      end
    else if Result=2 then
      begin
      if V_PGI.SAV then PGIInfo('Aucun service instancié sur eWS');
      end
    else if Result=3 then
      begin
      if V_PGI.SAV then PGIInfo('Le client '+NoDossier+' n''accède pas au service sur eWS');
      end
    else if Result=4 then
      begin
      if V_PGI.SAV then PGIInfo('Erreur d''accès à la méthode RetourneServiceClient');
      end;

  except
    on E:Exception do
      if V_PGI.SAV then PGIInfo('Erreur appel RetourneServiceClient : ' + #13#10 + E.Message );
  end;

end;


function EwsPublieDocument (TGN:TGedDPNode; bWithIhm:boolean):boolean;
var
    sTempFileName, sTitre   :string;
    sEwsID, sEwsLib         :string;
    iPubStatus              :integer;
    bFunc229orSup           :boolean; // $$$ JP 11/10/06 indique si la fonction v2.29 (ou +) à fonctionné
begin
     Result := False;

     // Réservé aux documents affectés à un dossier avec fichier accessible
     if (TGN=Nil) or (TGN.TypeGed<>'DOC') or (TGN.SFileGUID='') or (TGN.NoDossier='') then exit;

     if Not EwsVerifOuCreeCollab then exit;

     if Not IsDossierEws(TGN.NoDossier) then exit;

     // Document déjà publié
     if TGN.EwsDatePub<>iDate1900 then
        if PGIAsk('Ce document a déjà été publié sur eWS le ' + DateToStr (TGN.EwsDatePub) + '.'#13#10' Confirmez-vous la publication ?')=mrNo then
           exit;

     // Que le document soit rattaché ou non à une branche eWS, choix de la branche
     sEwsId    := TGN.EwsId;
     EwsSelectionneNoeud(sEwsId, sEwsLib);
     if sEwsId='' then exit;

     // Extraction fichier temporaire
     sTempFileName := ExtraitDocument(TGN.SDocGUID, sTitre);
     if sTempFileName='' then exit;

     // $$$ JP 11/10/06: choix de la bonne fonction à utiliser, selon la version d'ews
     bFunc229orSup := GetEwsVersion >= 2029;

     // Publication du document
     if bWithIhm = FALSE then
     begin
          try
             if bFunc229orSup = TRUE then
                 iPubStatus := V_eWS.PublicationDocumentRetour (sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password),
                                                True,           // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                                sTempFileName,   // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                                '',              // CodeProfilPub = Code du profil de publication
                                                //                (par réf à la publication automatique, ou "centre de tri")
                                                V_PGI.UserLogin, // IdEmetteur = Identifiant de l'émetteur
                                                sTitre,          // Intitule = Libellé de la publication 100c
                                                DateToStr(TGN.DateDeb), // jj/mm/aaaa = Date de début de validité du document
                                                DateToStr(TGN.DateFin), // jj/mm/aaaa = Date de fin de validité du document
                                                TGN.Caption,     // Description = Mémo de description du document 2000c
                                                TGN.NoDossier,   // PGI_ClientID = Identifiant du client = client destinataire
                                                '',              // Destinataire = interlocuteur destinataire (clé eWS)
                                                True,            // Email = génération d'un mail de notification au destinataire
                                                False,           // AR = Acusé de réception lors de la lecture du document
                                                sEwsId)          // IdNoeud = Identifiant du répertoire de publication du document.
             else
                 if V_eWS.PublicationDocument (sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password),
                                                   False,           // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                                   sTempFileName,   // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                                   '',              // CodeProfilPub = Code du profil de publication
                                                                    //                (par réf à la publication automatique, ou "centre de tri")
                                                   V_PGI.UserLogin, // IdEmetteur = Identifiant de l'émetteur
                                                   sTitre,          // Intitule = Libellé de la publication 100c
                                                   DateToStr(TGN.DateDeb), // jj/mm/aaaa = Date de début de validité du document
                                                   DateToStr(TGN.DateFin), // jj/mm/aaaa = Date de fin de validité du document
                                                   TGN.Caption,     // Description = Mémo de description du document 2000c
                                                   TGN.NoDossier,   // PGI_ClientID = Identifiant du client = client destinataire
                                                   '',              // Destinataire = interlocuteur destinataire (clé eWS)
                                                   True,            // Email = génération d'un mail de notification au destinataire
                                                   False,           // AR = Acusé de réception lors de la lecture du document
                                                   sEwsId) = FALSE then     // IdNoeud = Identifiant du répertoire de publication du document.
                    iPubStatus := -1;
          except
                on E:Exception do
                begin
                     iPubStatus := -1;
                     if V_PGI.SAV then if bFunc229orSup = TRUE then PGIInfo ('Erreur PublicationDocumentRetour: '#13#10' ' + E.Message) else PGIInfo ('Erreur PublicationDocument: '#13#10' ' + E.Message);
                end;
          end;
     end
     else
     begin
          try
             if bFunc229orSup = TRUE then
                 iPubStatus := V_eWS.FichePublicationDocumentRetour (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password),
                                                   True,           // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                                   sTempFileName,   // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                                   '',              // CodeProfilPub = Code du profil de publication
                                                                    //                (par réf à la publication automatique, ou "centre de tri")
                                                   V_PGI.UserLogin, // IdEmetteur = Identifiant de l'émetteur
                                                   sTitre,          // Intitule = Libellé de la publication 100c
                                                   DateToStr(TGN.DateDeb), // jj/mm/aaaa = Date de début de validité du document
                                                   DateToStr(TGN.DateFin), // jj/mm/aaaa = Date de fin de validité du document
                                                   TGN.Caption,     // Description = Mémo de description du document 2000c
                                                   TGN.NoDossier,   // PGI_ClientID = Identifiant du client = client destinataire
                                                   FALSE,           // Modifier destinataire
                                                   '',              // Destinataire = interlocuteur destinataire (clé eWS)
                                                   True,            // Email = génération d'un mail de notification au destinataire
                                                   False,           // AR = Acusé de réception lors de la lecture du document
                                                   sEwsId,          // IdNoeud = Identifiant du répertoire de publication du document.
                                                   0,               // X:
                                                   0                // Y:
                                                   )
             else
                 if V_eWS.FichePublicationDocument (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password),
                                                False,           // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                                sTempFileName,   // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                                '',              // CodeProfilPub = Code du profil de publication
                                                //                (par réf à la publication automatique, ou "centre de tri")
                                                V_PGI.UserLogin, // IdEmetteur = Identifiant de l'émetteur
                                                sTitre,          // Intitule = Libellé de la publication 100c
                                                DateToStr(TGN.DateDeb), // jj/mm/aaaa = Date de début de validité du document
                                                DateToStr(TGN.DateFin), // jj/mm/aaaa = Date de fin de validité du document
                                                TGN.Caption,     // Description = Mémo de description du document 2000c
                                                TGN.NoDossier,   // PGI_ClientID = Identifiant du client = client destinataire
                                                FALSE,           // Modifier destinataire
                                                '',              // Destinataire = interlocuteur destinataire (clé eWS)
                                                True,            // Email = génération d'un mail de notification au destinataire
                                                False,           // AR = Acusé de réception lors de la lecture du document
                                                sEwsId) = FALSE then     // IdNoeud = Identifiant du répertoire de publication du document.
                    iPubStatus := -1;
          except
                on E:Exception do
                begin
                     iPubStatus := -1;
                     if V_PGI.SAV then if bFunc229orSup = TRUE then PGIInfo ('Erreur FichePublicationDocumentRetour: '#13#10' ' + E.Message) else PGIInfo ('Erreur FichePublicationDocument: '#13#10' ' + E.Message);
                end;
          end;
     end;

     // Informe que la publication du document est effectuée
     if iPubStatus = 0 then
     begin
          Result := TRUE;
          ExecuteSQL('UPDATE DPDOCUMENT SET DPD_EWSDATEPUBL="'+UsDateTime(Now)+'" WHERE DPD_DOCGUID="'+TGN.SDocGUID+'"');
     end
     else
     case iPubStatus of
          1:
            PgiInfo ('Publication du document dans Ews annulée');

          2:
            PgiInfo ('Publication du document dans Ews impossible, car le répertoire cible est inconnu');

          3:
            PgiInfo ('Publication du document dans Ews impossible, car la règle (profil) est inconnue');

          10:
            PgiInfo ('Publication du document dans Ews impossible, car il est incomplet');
     end;

     // Ménage
     DeleteFile (sTempFileName);
end;


function EwsPublieUnDocument (SDocGUID:string; EwsId,EwsRegle:string; var Retour: String; bForceRepub: Boolean=True; bWithIhm:boolean=FALSE): Boolean;
var
   sTempFileName         :string;
   TobDoc, T             :TOB;
   strDeb, strFin        :string;
   strLib, strDesc       :string;
   iPubStatus            :integer;
   bFunc229orSup         :boolean; // $$$ JP 11/10/06 indique si la fonction v2.29 (ou +) à fonctionné
begin
  Result := False;

  // $$$ JP 11/04/06 - soit id, soit règle
  if (EwsId <> '') and (EwsRegle <> '') then
     exit;

  // Recherche les infos du document
  TobDoc := TOB.Create('Le document', Nil, -1);
  // INNER car il y a toujours un document, même s'il n'a pas de fichiers

  TobDoc.LoadDetailFromSQL('SELECT DPDOCUMENT.*, YDOCUMENTS.*, YDOCFILES.*, YFILES.YFI_FILENAME'
   + ' FROM DPDOCUMENT INNER JOIN YDOCUMENTS ON DPD_DOCGUID=YDO_DOCGUID'
                   + ' LEFT JOIN YDOCFILES ON YDO_DOCGUID=YDF_DOCGUID'
                   + ' LEFT JOIN YFILES ON YDF_FILEGUID=YFI_FILEGUID'
   + ' WHERE DPD_DOCGUID="'+SDocGUID+'"'
   + ' AND YDF_FILEROLE="PAL"'      // uniquement le fichier Principal
   + ' ORDER BY YDO_DATECREATION DESC'); // #### Quel classement ?
  if TobDoc.Detail.Count=0 then
    begin
    Retour := 'Le document '+SDocGUID+' n''a pas pu être extrait.';
    TobDoc.Free;
    exit;
    end;

  T := TobDoc.Detail[0];
  // Réservé aux documents affectés à un dossier avec fichier accessible
  if ( T.GetString('YDF_DOCGUID')='' ) or ( T.GetValue('YDF_FILEGUID')='' ) then
    begin
    Retour := 'Le document '+T.GetValue('YDO_LIBELLEDOC')+' ('+SDocGUID+') n''a pas de fichier associé.';
    TobDoc.Free;
    exit;
    end;

  // Le document pointe vers un fichier disparu
  if T.GetString('YFI_FILENAME')='' then
    begin
    Retour := 'Le fichier du document '+T.GetValue('YDO_LIBELLEDOC')+' ('+SDocGUID+') est indisponible.';
    TobDoc.Free;
    exit;
    end;

  // Le document n''est pas associé à un dossier
  if T.GetValue('DPD_NODOSSIER')='' then
    begin
    Retour := 'Le document '+T.GetValue('YDO_LIBELLEDOC')+' ('+SDocGUID+') n''est pas associé à un dossier.';
    TobDoc.Free;
    exit;
    end;

  if Not EwsVerifOuCreeCollab then
    begin
    Retour := 'Impossible de synchroniser le collaborateur '+V_PGI.UserLogin+' sur eWS.';
    TobDoc.Free;
    exit;
    end;

  if Not IsDossierEws(T.GetValue('DPD_NODOSSIER')) then
    begin
    Retour := 'Le client n''est pas '+T.GetValue('DPD_NODOSSIER')+' n''est pas créé sur eWS.';
    TobDoc.Free;
    exit;
    end;

  // Document déjà publié
  if (Not bForceRepub) and (T.GetDateTime('DPD_EWSDATEPUBL')<>iDate1900) then
    begin
    Retour := 'Le document '+T.GetValue('YDO_LIBELLEDOC')+' ('+SDocGUID+') a déjà été publié sur eWS le '+DateToStr(T.GetDateTime('DPD_EWSDATEPUBL'))+'.';
    TobDoc.Free;
    exit;
    end;

  // Extraction fichier temporaire
  sTempFileName := ExtraitDocument(SDocGUID, strLib);
  if sTempFileName='' then
    begin
    Retour := 'Impossible d''extraire le '+T.GetValue('YDO_LIBELLEDOC')+' ('+SDocGUID+')';
    TobDoc.Free;
    exit;
    end;

     // Publication du document
     // $$$ JP 18/08/06: si règle, il ne faut pas spécifier les dates, le libellé ni la description
     if EwsRegle <> '' then
     begin
          strDeb  := '';
          strFin  := '';
          //strLib  := '';
          strDesc := '';
     end
     else
     begin
          strDeb  := DateToStr (T.GetDateTime ('YDO_DATEDEB'));
          strFin  := DateToStr (T.GetDateTime ('YDO_DATEFIN'));
          strDesc := T.GetValue ('YDO_LIBELLEDOC');
     end;

     // $$$ JP 11/10/06: choix de la bonne fonction à utiliser, selon la version d'ews
     bFunc229orSup := GetEwsVersion >= 2029;

     if bWithIhm = FALSE then
     begin
          try
             if bFunc229orSup = TRUE then
                 iPubStatus := V_eWS.PublicationDocumentRetour (sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password),
                                    True,                                         // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                    sTempFileName,                                 // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                    EwsRegle,                                      // CodeProfilPub = Code du profil de publication (par réf à la publication automatique, ou "centre de tri")
                                    V_PGI.UserLogin,                               // IdEmetteur = Identifiant de l'émetteur
                                    strLib,                                        // Intitule = Libellé de la publication 100c
                                    strDeb,                                        // jj/mm/aaaa = Date de début de validité du document
                                    strFin,                                        // jj/mm/aaaa = Date de fin de validité du document
                                    strDesc,                                       // Description = Mémo de description du document 2000c
                                    T.GetValue('DPD_NODOSSIER'),                   // PGI_ClientID = Identifiant du client = client destinataire
                                    '',                                            // Destinataire = interlocuteur destinataire (clé eWS)
                                    True,                                          // Email = génération d'un mail de notification au destinataire
                                    False,                                         // AR = Acusé de réception lors de la lecture du document
                                    EwsId)                                         // IdNoeud = Identifiant du répertoire de publication du document.
             else
                 if V_eWS.PublicationDocument (sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password),
                                    False,                                         // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                    sTempFileName,                                 // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                    EwsRegle,                                      // CodeProfilPub = Code du profil de publication (par réf à la publication automatique, ou "centre de tri")
                                    V_PGI.UserLogin,                               // IdEmetteur = Identifiant de l'émetteur
                                    strLib,                                        // Intitule = Libellé de la publication 100c
                                    strDeb,                                        // jj/mm/aaaa = Date de début de validité du document
                                    strFin,                                        // jj/mm/aaaa = Date de fin de validité du document
                                    strDesc,                                       // Description = Mémo de description du document 2000c
                                    T.GetValue('DPD_NODOSSIER'),                   // PGI_ClientID = Identifiant du client = client destinataire
                                    '',                                            // Destinataire = interlocuteur destinataire (clé eWS)
                                    True,                                          // Email = génération d'un mail de notification au destinataire
                                    False,                                         // AR = Acusé de réception lors de la lecture du document
                                    EwsId) = FALSE then                                    // IdNoeud = Identifiant du répertoire de publication du document.
                     iPubStatus := -1;
          except
                on E:Exception do
                begin
                     iPubStatus := -1;
                     if V_PGI.SAV then if bFunc229orSup = TRUE then PGIInfo ('Erreur PublicationDocumentRetour: '#13#10' ' + E.Message) else PGIInfo ('Erreur PublicationDocument: '#13#10' ' + E.Message);
                end;
          end;
     end
     else
     begin
          try
             if bFunc229orSup = TRUE then
                 iPubStatus := V_eWS.FichePublicationDocumentRetour (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password),
                                                  True,                                      // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                                  sTempFileName,                             // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                                  EwsRegle,                                  // CodeProfilPub = Code du profil de publication (par réf à la publication automatique, ou "centre de tri")
                                                  V_PGI.UserLogin,                           // IdEmetteur = Identifiant de l'émetteur
                                                  strLib,                                    // Intitule = Libellé de la publication 100c
                                                  strDeb,                                    // jj/mm/aaaa = Date de début de validité du document
                                                  strFin,                                    // jj/mm/aaaa = Date de fin de validité du document
                                                  strDesc,                                   // Description = Mémo de description du document 2000c
                                                  T.GetValue('DPD_NODOSSIER'),               // PGI_ClientID = Identifiant du client = client destinataire
                                                  FALSE,                                     // Modifier destinataire
                                                  '',                                        // Destinataire = interlocuteur destinataire (clé eWS)
                                                  True,                                      // Email = génération d'un mail de notification au destinataire
                                                  False,                                     // AR = Acusé de réception lors de la lecture du document
                                                  EwsId,                                     // IdNoeud = Identifiant du répertoire de publication du document.
                                                  0,                                         // X:
                                                  0                                          // Y:
                                                  )
             else
                 if V_eWS.FichePublicationDocument (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password),
                                            False,                                     // FlagIncomplet = Booléan permettant de traiter la publication sur les documents incomplets
                                            sTempFileName,                             // CheminFichier = Chaine indiquant le nom complet du fichier à publier
                                            EwsRegle,                                  // CodeProfilPub = Code du profil de publication (par réf à la publication automatique, ou "centre de tri")
                                            V_PGI.UserLogin,                           // IdEmetteur = Identifiant de l'émetteur
                                            strLib,                                    // Intitule = Libellé de la publication 100c
                                            strDeb,                                    // jj/mm/aaaa = Date de début de validité du document
                                            strFin,                                    // jj/mm/aaaa = Date de fin de validité du document
                                            strDesc,                                   // Description = Mémo de description du document 2000c
                                            T.GetValue('DPD_NODOSSIER'),               // PGI_ClientID = Identifiant du client = client destinataire
                                            FALSE,                                     // Modifier destinataire
                                            '',                                        // Destinataire = interlocuteur destinataire (clé eWS)
                                            True,                                      // Email = génération d'un mail de notification au destinataire
                                            False,                                     // AR = Acusé de réception lors de la lecture du document
                                            EwsId) = FALSE then                            // IdNoeud = Identifiant du répertoire de publication du document.
                    iPubStatus := -1;
          except
                on E:Exception do
                begin
                     iPubStatus := -1;
                     if V_PGI.SAV then if bFunc229orSup = TRUE then PGIInfo ('Erreur FichePublicationDocumentRetour: '#13#10' ' + E.Message) else PGIInfo ('Erreur FichePublicationDocument: '#13#10' ' + E.Message);
                end;
          end;
     end;

     // Informe que la publication du document est effectuée
     if iPubStatus = 0 then
     begin
          Result := TRUE;
          ExecuteSQL('UPDATE DPDOCUMENT SET DPD_EWSDATEPUBL="'+UsDateTime(Now)+'" WHERE DPD_DOCGUID="'+SDocGUID+'"');
     end
     else
     case iPubStatus of
          1:
            Retour := 'publication annulée par l''utilisateur';

          2:
            Retour := 'le répertoire cible est inconnu';

          3:
            Retour := 'la règle (profil) de publication est inconnue';

          10:
            Retour := 'le document est incomplet pour Ews';
     end;

  // Ménage
  DeleteFile(sTempFileName);
  TobDoc.Free;
end;


function  EwsFichePublicationMemo: Boolean;
begin
  Result := False;

  if Not EwsVerifOuCreeCollab then exit;

  // Publication du document
  try
    if V_eWS.FichePublicationMemo(sNumeroSIO, V_PGI.UserLogin, DecryptageSt(V_PGI.Password),
     '',              // Type de message : défilant ou fixe
     '',              // Chemin du fichier : inutile
     '',              // CodeProfilPub : Code du profil de publication
     V_PGI.UserLogin, // IdEmetteur = Identifiant de l'émetteur
     '',              // Intitule = Libellé de la publication 100c
     '',              // jj/mm/aaaa = Date de début de validité du mémo
     '',              // jj/mm/aaaa = Date de fin de validité du mémo
     '',              // Description = Mémo de description du document 2000c
     '',              // PGI_ClientID = Identifiant du client = client destinataire
     True,
     '',              // Destinataire = interlocuteur destinataire (clé eWS)
     False,           // Email = Adresse électronique d'accusé de réception
     False,           // AR = Acusé de réception lors de la lecture du mémo
     '')              // IdNoeud = Identifiant du noeud de publication du mémo.
      then Result := True;
  except
    on E:Exception do
      if V_PGI.SAV then PGIInfo('Erreur appel FichePublicationMemo : ' + #13#10 + E.Message );
  end;
end;


function  EwsVerificationSIO(SIO, SIOPassword : String) : Integer;
//    1 : Ok
//    0 : Mot de passe erroné
//   -1 : SIO inexistant
begin
  Result := -1;
  if Not EwsLibrairieOk then exit;

  // Vérification du SIO
  try
    Result := V_eWS.VerificationSIO(
      SIO,            // Identifiant du cabinet 50c
     'ADM',           // Identifiant du collaborateur dans eWS 50c
     GetAdmPwd {'cegid'},         // Mot de passe du collaborateur dans eWS 50c (désormais prise en compte du mot de passe si ADM existe dans pgi)
      SIO,            // Numéro SIO du cabinet (site) 9c
      SIOPassword);   // Mot de passe du SIO site (8 position hexa actuellement)
  except
    on E:Exception do
      if V_PGI.SAV then PGIInfo('Erreur appel VerificationSIO : ' + #13#10 + E.Message );
  end;
end;


procedure EwsGetInfosDossier(NoDossier : String; var LibDossier, EmailDossier, UsrS1, PwdS1 : String);
var
    Q : TQuery;
    SGuidPer : String;
begin
  // Récup infos du dossier
  SGuidPer := '';
  Q := OpenSQL('SELECT DOS_GUIDPER, DOS_LIBELLE FROM DOSSIER WHERE DOS_NODOSSIER="'+NoDossier+'"', True);
  if Not Q.Eof then
    begin
    SGuidPer := Q.FindField('DOS_GUIDPER').AsString;
    LibDossier := Q.FindField('DOS_LIBELLE').AsString;
    end;
  Ferme(Q);
  if SGuidPer='' then
    begin
    if V_PGI.SAV then PGIInfo('Le dossier '+NoDossier+' n''a pas de fiche annuaire.');
    exit;
    end;

  // Récup email dossier en cours
  Q := OpenSQL('SELECT ANN_EMAIL FROM ANNUAIRE WHERE ANN_GUIDPER="'+SGuidPer+'"', True);
  if Not Q.Eof then EmailDossier := Q.Fields[0].AsString;
  Ferme(Q);
  // Récup user / password éventuellement utilisés pour S1
  Q := OpenSQL('SELECT DOS_USRS1, DOS_PWDS1, DOS_EWSCREE FROM DOSSIER WHERE DOS_NODOSSIER="'+NoDossier+'"', True);
  if Not Q.Eof then
    begin
    UsrS1 := Q.FindField('DOS_USRS1').AsString;
    PwdS1 := Q.FindField('DOS_PWDS1').AsString;
    end;
  Ferme(Q);
end;


function  IsDossierEws(NoDossier : String) : Boolean;
begin
  Result := ExisteSQL('SELECT 1 FROM DOSSIER WHERE DOS_NODOSSIER="'+NoDossier+'" AND DOS_EWSCREE="X"');
end;


function  GetNumeroSIO : Boolean ;
begin
  Result := False;
  sNumeroSIO := GetParamsocDpSecur('SO_NECABIDENT', '');
  if sNumeroSIO='' then
    begin
    if V_PGI.SAV then PGIInfo('Erreur appel GetNumeroSIO : ' + #13#10 + 'Impossible de déterminer l''identifiant Cabinet.');
    end
  else
    Result := True;
end;

// $$$ JP 21/07/06: demande de confirmation pas toujours systématique
procedure EwsSynchronize (bAskConfirm:boolean=TRUE);
{var
   strEtape   :string;
   wCount     :word; }
begin
     // $$$ JP 21/07/06: demander confirmation si c'est autorisé
     if (bAskConfirm = FALSE) or (PgiAsk ('Confirmez-vous la synchronisation eWS?') = mrYes) then
        if AgenteWsLibrairieOk = TRUE then
           try
              SourisSablier;
              // MD 17/08/06 - Aucune raison d'attendre la fin de la synchro !
              V_eWSAgent.LancementSynchronisation (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password));
            { if V_eWSAgent.LancementSynchronisation (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password)) = TRUE then
              begin
                strEtape := LowerCase (v_eWSAgent.EtapeSynchronisation (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password)));
                wCount   := 0;
                while (strEtape <> '') and (strEtape <> 'manuelle') and (strEtape <> 'programmé') do
                begin
                     Inc (wCount);
                     if wCount > 40 then
                     begin
                          wCount := 0;
                          if PgiAsk ('La synchronisation eWS est encore en cours ("' + strEtape + '"'#10' Désirez-vous attendre de nouveau la fin de cette synchronisation ?'#10' (il est déconseillé de ne pas attendre, sauf en cas de blocage)') = mrNo then
                             exit;
                     end;
                     Application.ProcessMessages;
                     Delay (500);
                     strEtape := LowerCase (v_eWSAgent.EtapeSynchronisation (sNumeroSIO, V_PGI.UserLogin, DecryptageSt (V_PGI.Password)));
                end;
              end;   }
           finally
                  SourisNormale;
           end;
end;

function GetEwsVersion:integer; //IsEwsSuperieur229 : Boolean;
var
   PathEws         :string;
   strVersion      :string;
   iMajor, iMinor  :integer;
begin
     // Par défaut, pas de version connue
     Result := 0;

     // Version de la dll principale d'Ews dans le répertoire d'installation d'Ews
     PathEws := GetEwsPath;
     if PathEws <> '' then
     begin
          strVersion := GetFileVersion (pathEws + '\Composants\ExploitEws.dll');

          // On reformate pour n'avoir que la version majeure/mineure, pas les n° de build
          iMajor := StrToInt (ReadTokenPipe (strVersion, '.'));
          if strVersion <> '' then
              iMinor := StrToInt (ReadTokenPipe (strVersion, '.'))
          else
              iMinor := 0;

          Result := iMajor*1000 + iMinor;
     end;
end;

// $$$ JP 13/12/06 - connaitre le mot de passe de l'utilisateur "ADM"
function GetAdmPwd:string;
var
   TOBAdm  :TOB;
begin
     if sAdmPwd = '' then
     begin
          TOBAdm := TOB.Create ('adm pwd', nil, -1);
          try
             // S'il y a une synchro sur l'utilisateur ADM, on doit utiliser le mot de passe de cet utilisateur (connu dans Pgi)
             TOBAdm.LoadDetailFromSQL ('SELECT US_PASSWORD FROM UTILISAT RIGHT JOIN CHOIXCOD ON (CC_TYPE="YEW" AND CC_CODE=US_UTILISATEUR) WHERE US_UTILISATEUR="ADM"');
             if TOBAdm.Detail.Count > 0 then
                 sAdmPwd := DeCryptageSt (TOBAdm.Detail [0].GetString ('US_PASSWORD'))
             else
                 sAdmPwd := 'cegid';
          finally
                 TOBAdm.Free;
          end;
     end;

     // Mot de passe à utiliser pour l'utilisateur 'ADM'
     Result := sAdmPwd;
end;


initialization
   sNumeroSIO := '';
   sAdmPwd    := '';

end.
