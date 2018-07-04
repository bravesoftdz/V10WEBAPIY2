unit UdroitUtilisation;

interface

uses
  AglInit,
	hctrls,
{$IFDEF EAGLCLIENT}
     MenuOLX,Maineagl,
{$ELSE}
     MenuOLG,Fe_Main,
{$ENDIF}
	LicUtil,
  sysutils,
  HEnt1,
	windows,
  CBPPath,
  Classes,
  ParamSoc,
  forms,
  Cryptage,
  HmsgBox
  ;
function VerificationDroitsOk : boolean;
function getIdDroits : string;
procedure SetInfoActivation (Nomfichier : string);
procedure DefiniFichierDroits (Emplacement,id,Date : string);

implementation
uses DateUtils;

procedure ExtractInfo (Chaine : string; var Clef,Data : string);
var iposEq : Integer;
begin
  Clef := '';
  Data := '';
  iposEq := Pos('=',Chaine);
  if iposEq > 0 then
  begin
    Clef := Copy(Chaine,1,IposEq-1);
    Data := Copy(Chaine,IposEq+1,255);
  end;
end;
                                                       
function getIdDroits : string;
begin
  result := GetParamSocSecur('SO_IDSECURITE','');
end;

procedure DefiniFichierDroits (Emplacement,id,Date : string);
var NomFile : string;
    FFile : TextFile;
    TheText : string;
begin
  NomFile := IncludeTrailingBackslash(Emplacement)+ID+'.LIC';
  AssignFile(FFile, NomFile);
  ReWrite(FFile);
  //
  Writeln(FFile,'FICHIER DE DEFINITION DES DROITS');
  Writeln(FFile,'--------------------------------');
  Writeln(FFile);
  Writeln(FFile,'TOUTE MODIFICATION MANUELLE EST INTERDITE ET SERA DETECTEE');
  Writeln(FFile);
  Writeln(FFile,EnCrypte('IDENTIFIANT='+ID));
  Writeln(FFile,EnCrypte('TELLEMENETKILVONTCHERCHER.J''ENRIGOLEDEJA'));
  Writeln(FFile,EnCrypte('SITUTRADUISCABENJETEDONNELEDROITDECRAKERLERESTE'));
  Writeln(FFile,EnCrypte('INFOPROTECTION='+Date));
  //
  CloseFile(FFile);
  //
end;

procedure SetInfoActivation (Nomfichier : string);
var  FFile : TextFile;
    CryptedData,DataDecrypt,IDSECURITE,Clef,valeur,INFOPROTECTION : string;
    InfIdSecurite : string;
    infDatemax : string;
    ACTIVATION : boolean;
begin
  //
  ACTIVATION :=  GetParamSocSecur('SO_BTACTIVATE',false);
  IDSECURITE :=  GetParamSocSecur('SO_IDSECURITE','');
  INFOPROTECTION := GetParamSocSecur('SO_INFOPROTECTION','');
  if not ACTIVATION then
  BEGIN
    PgiInfo ('Pas d''activation nécessaire');
    Exit;
  end;
  // 
  AssignFile(FFile, Nomfichier);
  Reset(FFile);
  //
  while not Eof(FFile) do
  begin
    ReadLn(FFile, CryptedData);
    DataDecrypt := Decrypte(CryptedData);
    ExtractInfo (DataDecrypt,Clef,valeur);
    if Clef = 'IDENTIFIANT' then
    begin
      InfIdSecurite := valeur;
    end else if Clef = 'INFOPROTECTION' then
    begin
      if HENt1.IsValidDate(valeur) then infDatemax := valeur;
    end;
  end;
  //
  CloseFile(FFile);
  //
  if (InfIdSecurite = IDSECURITE) and (infDatemax<>'') then
  begin
    SetParamSoc('SO_INFOPROTECTION',EnCrypte(infDatemax));
  end else
  begin
    PgiInfo('Fichier de licence non conforme !!');
  end;

end;

function VerificationDroitsOk : boolean;
var Emplacement,FichierSecu : string;
    CryptedData,Datatext,IDSECURITE,Clef,valeur,INFOPROTECTION : string;
    //
    InfIdSecurite : string;
    infDatemax : TdateTime;
    ToDay : TDateTime;
    Delta : Integer;
    ACTIVATION : boolean;
begin
  ACTIVATION :=  GetParamSocSecur('SO_BTACTIVATE',false);
  IDSECURITE :=  GetParamSocSecur('SO_IDSECURITE','');
  INFOPROTECTION := GetParamSocSecur('SO_INFOPROTECTION','');
  //
  result := true; // seul cas ou c'est ok sans controle
  if not ACTIVATION then exit;
  //
  Result := false;
  ToDay := StrToDate(DateToStr(Now));
  // VOLEUUUUURRRR ---
  if IDSECURITE = '' then exit;
  IF INFOPROTECTION = '' then exit;
  // ------------------------
    Datatext := Decrypte(INFOPROTECTION);
    // -- On laisse 15 jours de répit --
    if HEnt1.IsValidDate(Datatext) then
    begin
      infDatemax := PlusDate(StrToDate(Datatext),15,'J');
    end;
    //
    if (ToDay <= InfDatemax) then
    begin
      Result := True;
      if DaysBetween(ToDay, InfDatemax) < 15 then
      begin
        Delta := DaysBetween(ToDay, InfDatemax);
        PGIInfo('ATTENTION : Vous avez dépassé la date d''utilisation des droits d''utilisation de ce produit.#13#10 Il vous reste '+InttoStr(Delta)+' Jours d''utilisation.#13#10 Veuillez prendre contact avec notre service commercial');
      end;
    end;
end;

end.
