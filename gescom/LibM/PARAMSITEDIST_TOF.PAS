{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 10/09/2001
Modifié le ... : 24/01/2002
Description .. : Source TOF de la TABLE : PARAMSITEDIST ()
Mots clefs ... : TOF;PARAMSITEDIST
*****************************************************************}
Unit PARAMSITEDIST_TOF ;

Interface                          

Uses UTOF,HCtrls,HMsgBox,HEnt1,LicUtil,ParamSoc,Vierge,Ras,
     Windows,SysUtils,Classes;

procedure ChargeParamSiteDistant(Var StServerAdd,StDossier,StUser,StPassword,
          StModeConnexion,StProfil,StProfilUser,StProfilPassword,StProfilTel :string);
procedure ProfilChange(StProfil : string; var StPhone,StUser,StPass : string);

Type
  TOF_PARAMSITEDIST = Class (TOF)
    private
    CanClose : boolean;
    iActive : integer;
    procedure SDMODECONNEXION_OnChange(Sender: TObject);
    procedure SDPROFIL_OnChange(Sender: TObject);
    public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_PARAMSITEDIST.OnUpdate ;
var StPassword,StModeConnexion,StProfilPass : string;
begin
  Inherited ;
CanClose:=False;
StModeConnexion:=THValComboBox(GetControl('SDMODECONNEXION')).Value;
if StModeConnexion='002' then
  begin
  if THValComboBox(GetControl('SDPROFIL')).Text='' then
    begin
    PGIInfo(TraduireMemoire('Un profil doit être renseigné.'),ecran.caption);
    SetFocusControl('SDPROFIL'); exit;
    end;
  StProfilPass:=CryptageSt(GetControlText('PASSWORD'));
  end;
if GetControlText('SDSERVERADDRESS')='' then
  begin
  PGIInfo(TraduireMemoire('L''adresse du serveur doit être renseignée.'),ecran.caption);
  SetFocusControl('SDSERVERADDRESS'); exit;
  end;
if GetControlText('SDDOSSIER')='' then
  begin
  PGIInfo(TraduireMemoire('La société doit être renseignée.'),ecran.caption);
  SetFocusControl('SDDOSSIER'); exit;
  end;
if GetControlText('SDUSERNAME')='' then
  begin
  PGIInfo(TraduireMemoire('Le nom d''utilisateur doit être renseigné.'),ecran.caption);
  SetFocusControl('SDUSERNAME'); exit;
  end;
SetParamSoc('SO_SDMODECONNEXION',StModeConnexion);
if iActive <= 0 then SetParamSoc('SO_SDPROFIL',THValComboBox(GetControl('SDPROFIL')).Text);
SetParamSoc('SO_SDSERVERADDRESS',GetControlText('SDSERVERADDRESS'));
SetParamSoc('SO_SDDOSSIER',GetControlText('SDDOSSIER'));
SetParamSoc('SO_SDUSERNAME',GetControlText('SDUSERNAME'));
StPassword:=CryptageSt(GetControlText('SDPASSWORD')); SetParamSoc('SO_SDPASSWORD',StPassword);
TFVierge(ecran).Retour:='SO_SDSERVERADDRESS='+GetControlText('SDSERVERADDRESS')+
                        ';SO_SDDOSSIER='+GetControlText('SDDOSSIER')+
                        ';SO_SDUSERNAME='+GetControlText('SDUSERNAME')+
                        ';SO_SDPASSWORD='+StPassword+
                        ';SO_SDMODECONNEXION='+StModeConnexion+
                        ';SO_SDPROFIL='+THValComboBox(GetControl('SDPROFIL')).Text+
                        ';PROFILUSER='+GetControlText('USERNAME')+
                        ';PROFILPASS='+StProfilPass+
                        ';PROFILTEL='+GetControlText('TEL');
CanClose:=True;
end ;

procedure TOF_PARAMSITEDIST.OnLoad ;
var StModeConnexion : string;
begin
  Inherited ;
  THValComboBox(GetControl('SDMODECONNEXION')).OnChange := SDMODECONNEXION_OnChange ;
  THValComboBox(GetControl('SDPROFIL')).OnChange := SDPROFIL_OnChange ;

  StModeConnexion:=GetParamSoc('SO_SDMODECONNEXION');
  SetControlVisible('P_RTC_ISDN',StModeConnexion='002');
  THValComboBox(GetControl('SDMODECONNEXION')).Value := StModeConnexion;
  if AglLoadRasApi32 then
    begin
    AglInitPhonebookCombo(THValComboBox(GetControl('SDPROFIL')).Items, iActive);
    THValComboBox(GetControl('SDPROFIL')).Values.Assign(THValComboBox(GetControl('SDPROFIL')).Items);
    if iActive > 0 then
      begin
      THValComboBox(GetControl('SDPROFIL')).ItemIndex := iActive;
      SetControlEnabled('SDMODECONNEXION',True); SetControlEnabled('SDPROFIL',True);
      end
    else THValComboBox(GetControl('SDPROFIL')).Value:=GetParamSoc('SO_SDPROFIL');
    end;
  SetControlText('SDSERVERADDRESS', GetParamSoc('SO_SDSERVERADDRESS'));
  SetControlText('SDUSERNAME', GetParamSoc('SO_SDUSERNAME'));
  SetControlText('SDPASSWORD', DeCryptageSt(GetParamSoc('SO_SDPASSWORD')));
  SetControlText('SDDOSSIER', GetParamSoc('SO_SDDOSSIER'));
end ;

procedure TOF_PARAMSITEDIST.OnArgument (S : String ) ;
begin
Inherited ;
CanClose:=True;
end ;

procedure TOF_PARAMSITEDIST.OnClose ;
begin
if CanClose then LastError:=0 else LastError:=-1;
CanClose:=True ;
end;

// Sélection d'un nouveau mode de connexion
// 001 : Réseau local/étendu mais direct ( routeur )
//       Pas de besoin de donnée supplémentaire
// 002 : Réseau Téléphonique Commuté
//       Besoin du profil de connexion, de l'utilisateur et du mot de passe
procedure TOF_PARAMSITEDIST.SDMODECONNEXION_OnChange(Sender: TObject);
var StModeConnexion:string;
begin
  StModeConnexion:=THValComboBox(GetControl('SDMODECONNEXION')).Value;
  if StModeConnexion = '001' then
    SetControlVisible('P_RTC_ISDN',False)
  else if StModeConnexion = '002' then
    begin
    SetControlVisible('P_RTC_ISDN',True);
    // Positionnement sur le choix du profil de connexion
    THValComboBox(GetControl('SDPROFIL')).SetFocus ;
    end;
end;

procedure TOF_PARAMSITEDIST.SDPROFIL_OnChange(Sender: TObject);
var StPhone,StUser,StPass : string;
begin
ProfilChange(THValComboBox(GetControl('SDPROFIL')).Text,StPhone,StUser,StPass);
SetControlText('USERNAME',StUser);
SetControlText('PASSWORD',StPass);
SetControlText('TEL',StPhone);
end;

///////////////////////////////////////////////////////////////////////////////

procedure ProfilChange(StProfil : string; var StPhone,StUser,StPass : string);
var RasEntry : TRASENTRY;
    rasdialparams : TRASDIALPARAMS;
    dwEntryInfoSize, dwError : Dword;
    fPassword : LongBOOL;
begin
  FillChar(RasEntry, sizeof(RasEntry), 0);
  RasEntry.dwSize := sizeof(RasEntry);
  dwEntryInfoSize := RasEntry.dwSize;

  dwError := RasGetEntryPropertiesA(
		nil,        // pointer to full path and filename of phone-book file
		PChar(StProfil),  // pointer to an entry name
		@RasEntry,  // buffer that receives entry information
		@dwEntryInfoSize,  // size, in bytes, of the lpRasEntry buffer
		nil,nil);
  StPhone := '';
  if (dwError = 0) then StPhone := RasEntry.szLocalPhoneNumber;
  FillChar(rasdialparams, sizeof(rasdialparams), 0);
  rasdialparams.dwSize := sizeof(rasdialparams);
  strcopy(rasdialparams.szEntryName , PChar(StProfil));
  dwError := RasGetEntryDialParamsA(
		nil,            // pointer to the full path and filename of the phone-book file
		@rasdialparams, // pointer to a structure that receives the connection parameters
		@fPassword);    // indicates whether the user's password was retrieved
  if (dwError = 0) then
    begin
    if (rasdialparams.szUserName[0] <> #0) then StUser := rasdialparams.szUserName;
    if (fPassword) then
      if (rasdialparams.szPassword[0] <> #0) then StPass := rasdialparams.szPassword;
    end;
end;

procedure ChargeParamSiteDistant(Var StServerAdd,StDossier,StUser,StPassword,
          StModeConnexion,StProfil,StProfilUser,StProfilPassword,StProfilTel :string);
begin
StServerAdd:=GetParamSoc('SO_SDSERVERADDRESS');
StDossier:=GetParamSoc('SO_SDDOSSIER');
StUser:=GetParamSoc('SO_SDUSERNAME');
StPassword:=GetParamSoc('SO_SDPASSWORD');
StModeConnexion:=GetParamSoc('SO_SDMODECONNEXION');
StProfil:=''; StProfilUser:=''; StProfilPassword:=''; StProfilTel:='';
if StModeConnexion='002' then
  begin
  StProfil:=GetParamSoc('SO_SDPROFIL');
  if AglLoadRasApi32 then ProfilChange(StProfil,StProfilTel,StProfilUser,StProfilPassword);
  end;
end;

Initialization
  registerclasses ( [ TOF_PARAMSITEDIST ] ) ;
end.
