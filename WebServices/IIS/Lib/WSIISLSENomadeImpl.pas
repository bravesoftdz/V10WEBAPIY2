{ Fichier d'implémentation invocable pour TWSIISLSENomade qui implémente IWSIISLSENomade }

unit WSIISLSENomadeImpl;

interface

uses InvokeRegistry, Types, XSBuiltIns, WSIISLSENomadeIntf, WSNomadeDecl,IniFiles,Registry,Windows;

type

  { TWSIISLSENomade }
  TWSIISLSENomade = class(TInvokableClass, IWSIISLSENomade)
  protected
    function IsValideConnect ( ParamIn : TWSUserParam) : WideString; stdcall;
    function GetAppels (identificateur : WideString;Depuis : Widestring) : WideString; stdcall;
    function GetAppel (identificateur : WideString; CodeAppel : Widestring) : WideString; stdcall;
  end;

implementation
uses Udefinitions, SysUtils,Ulog,USql;

{ TWSIISLSENomade }

function TWSIISLSENomade.GetAppel(identificateur,CodeAppel: Widestring): WideString;
var IniFile : TIniFile;
    EmplacementIni,NomIni : string;
    IDSto : string;
    Response : WideString;
begin
  //
  EmplacementIni := getCheminIni;
  NomIni := IncludeTrailingPathDelimiter(EmplacementIni)+'Definitions.ini';
	ecritLog (EmplacementIni,'------------------le '+ DateTimeToStr (Now)+'----------------');
	ecritLog (EmplacementIni,'---> ouverture fichier INI');
	ecritLog (EmplacementIni,'---> Reception IdMateriel :'+Identificateur);
  IniFile := TIniFile.Create(NomIni);
  IdSto := IniFile.ReadString(identificateur,'ID','');
  TRY
    if (Identificateur = IDSto) then
    begin
      Response  := GetDataAppel (CodeAppel,
                    			 			IniFile.ReadString(identificateur,'DATABASE',''),
                            		IniFile.ReadString(identificateur,'SERVER',''));
    end else
    begin
      Response := 'FUCKOFF';
			ecritLog (EmplacementIni,'---> -- ERREUR -- Materiel ou utilisateur non identifié');
    end;
  FINALLY
		ecritLog (EmplacementIni,'---------------------------------------------------------------------------');
  	Result := Response;
    FreeAndNil (IniFile);
  end;
end;

function TWSIISLSENomade.GetAppels(identificateur : WideString; Depuis : Widestring): WideString;
var IniFile : TIniFile;
    EmplacementIni,NomIni : string;
    IDSto : string;
    Response : WideString;
begin
  //
  EmplacementIni := getCheminIni;
  NomIni := IncludeTrailingPathDelimiter(EmplacementIni)+'Definitions.ini';
	ecritLog (EmplacementIni,'------------------le '+ DateTimeToStr (Now)+'----------------');
	ecritLog (EmplacementIni,'---> ouverture fichier INI');
	ecritLog (EmplacementIni,'---> Reception IdMateriel :'+Identificateur);
  IniFile := TIniFile.Create(NomIni);
  IdSto := IniFile.ReadString(identificateur,'ID','');
  TRY
    if (Identificateur = IDSto) then
    begin
      Response  := GetTopAppels (IniFile.ReadString(identificateur,'RESSOURCE',''),
                    						IniFile.ReadString(identificateur,'DATABASE',''),
                                IniFile.ReadString(identificateur,'SERVER',''),Depuis);
    end else
    begin
      Response := 'FUCKOFF';
			ecritLog (EmplacementIni,'---> -- ERREUR -- Materiel ou utilisateur non identifié');
    end;
  FINALLY
		ecritLog (EmplacementIni,'---------------------------------------------------------------------------');
  	Result := Response;
    FreeAndNil (IniFile);
  end;
end;

function TWSIISLSENomade. IsValideConnect ( ParamIn : TWSUserParam) : WideString;
var IniFile : TIniFile;
    EmplacementIni,NomIni : string;
    UserSto,passWordSto,StrResponse : string;
begin
  //
  StrResponse := '????????????';
  EmplacementIni := getCheminIni;
  NomIni := IncludeTrailingPathDelimiter(EmplacementIni)+'Definitions.ini';
	ecritLog (EmplacementIni,'------------------le '+ DateToStr(Now)+'----------------');
	ecritLog (EmplacementIni,'---> ouverture fichier INI');
	ecritLog (EmplacementIni,'---> Reception IdMateriel :'+ParamIn.UniqueID);
	ecritLog (EmplacementIni,'---> Reception username :'+ParamIn.User);
	ecritLog (EmplacementIni,'---> Reception PassWord :'+ParamIn.Password);
  IniFile := TIniFile.Create(NomIni);
  UserSto := IniFile.ReadString(ParamIn.UniqueID,'USER','');
  passWordSto := IniFile.ReadString(ParamIn.UniqueID,'PASSWORD','');
  TRY
    if (ParamIn.Password <> '') and (ParamIn.User <> '') and (ParamIn.User = UserSto) and (ParamIn.Password=passWordSto) then
    begin
			ecritLog (EmplacementIni,'---> Materiel et utilisateur identifié');
      StrResponse := ParamIn.UniqueID+SepChamp+userSto;
      StrResponse := StrResponse + SepChamp + IniFile.ReadString(ParamIn.UniqueID,'DATABASE','');
      StrResponse := StrResponse + SepChamp + IniFile.ReadString(ParamIn.UniqueID,'SERVER','');
      StrResponse := StrResponse + SepChamp+ IniFile.ReadString(ParamIn.UniqueID,'RESSOURCE','');
      StrResponse := StrResponse + sepchamp + '0';
    end else
    begin
			ecritLog (EmplacementIni,'---> -- ERREUR -- Materiel ou utilisateur non identifié');
    end;
  FINALLY
		ecritLog (EmplacementIni,'---------------------------------------------------------------------------');
  	Result := StrResponse;
    FreeAndNil (IniFile);
  end;
end;


initialization
  { Les classes invocables doivent être recensées }
  InvRegistry.RegisterInvokableClass(TWSIISLSENomade);

end.
 