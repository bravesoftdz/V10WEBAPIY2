{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/06/2002
Modifié le ... : 19/06/2002
Description .. : Fonctions de gestion de la TOX depuis le Front Office
Mots clefs ... : FO
*****************************************************************}
unit FOToxUtil;

interface

uses Classes, Windows, ShellAPI, Sysutils, FileCtrl,
  {$IFNDEF EAGLCLIENT}
  dbtables,
  {$ENDIF}
  Hctrls, Hent1, ParamSoc, uToxClasses;

///////////////////////////////////////////////////////////////////////////////////////
//  déclarations des fonctions et procédures.
///////////////////////////////////////////////////////////////////////////////////////
function FOGetToxIdApp: string;
function FOVerifSiteTox: Boolean;
procedure FOToxServeurMajHeure(Heure1, Heure2: string);
procedure FOToxServeurMajConf;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne l'identifiant de l'application pour le ToxServeur
Mots clefs ... : FO
*****************************************************************}

function FOGetToxIdApp: string;
var IdApplication, ToxHeure: string;
  WithWindowsInfo, WithDeconnect: Boolean;
begin
  IdApplication := GetSynRegKey('ToxIdApplication', '', True);
  if IdApplication = '' then
  begin
    IdApplication := 'MODES5';
    SaveSynRegKey('ToxIdApplication', IdApplication, True);
    WithWindowsInfo := False;
    SaveSynRegKey('ToxWithWindowsInfo', WithWindowsInfo, True);
    WithDeconnect := True;
    SaveSynRegKey('ToxWithDeconnect', WithDeconnect, True);
    ToxHeure := '09:00';
    SaveSynRegKey('ToxheureDebut', ToxHeure, True);
    ToxHeure := '08:00';
    SaveSynRegKey('ToxHeureFin', ToxHeure, True);
  end;
  Result := IdApplication;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/09/2001
Modifié le ... : 24/09/2001
Description .. : Indique si au moins un site local est défini dans la table des
Suite ........ : sites de la TOX.
Mots clefs ... : FO
*****************************************************************}

function FOVerifSiteTox: Boolean;
begin
  with TCollectionSites.Create(TCollectionSite, True) do
  begin
    Result := (LeSiteLocal <> nil);
    Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOToxServeurConf : indique le nom du fichier de configuration de l'application ToxServeur
///////////////////////////////////////////////////////////////////////////////////////

function FOToxServeurConf: string;
var Stg: string;
  Buffer: array[0..1028] of Char;
  Taille: Integer;
begin
  Taille := GetEnvironmentVariable(PChar('ProgramFiles'), Buffer, SizeOf(Buffer));
  if Taille > 0 then SetString(Stg, Buffer, Taille)
  else Stg := 'C:\Program Files';
  Stg := Stg + '\Cegid\TxSrv\Data';
  if DirectoryExists(Stg) then Result := Stg + '\TOX.DAT' else Result := '';
end;

///////////////////////////////////////////////////////////////////////////////////////
// FOToxServeurMajHeure : Enregistrement des heures d'appel du ToxServeur
///////////////////////////////////////////////////////////////////////////////////////

procedure FOToxServeurMajHeure(Heure1, Heure2: string);
///////////////////////////////////////////////////////////////////////////////////////
  procedure MajHeure(Lst: TStrings; Rang, NoAppel: Integer; Heure: string; LigneOk: Boolean; var MajFile: Boolean);
  var Ligne, St1, St2, St3: string;
  begin
    if LigneOk then
    begin
      Ligne := Lst[Rang];
      St1 := ReadTokenST(Ligne);
      St2 := ReadTokenST(Ligne);
      St3 := ReadTokenST(Ligne);
    end else
    begin
      St1 := 'H' + IntToStr(NoAppel - 1) + '=' + 'Appel no ' + IntToStr(NoAppel);
      St2 := '';
      St3 := '2';
    end;
    if St2 <> Heure then
    begin
      if LigneOk then Lst.Delete(Rang);
      if Heure <> '' then
      begin
        Ligne := St1 + ';' + Heure + ';' + St3;
        Lst.Insert(Rang, Ligne);
      end;
      MajFile := True;
    end;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
var FileName, Stg: string;
  MajFile, LigneOk: Boolean;
  Lst: TStrings;
  Ind, Rang: Integer;
begin
  MajFile := False;
  if not FOVerifSiteTox then Exit;
  FileName := FOToxServeurConf;
  if FileName = '' then Exit;
  Lst := TStringList.Create;
  try
    if FileExists(FileName) then
    begin
      Lst.LoadFromFile(FileName);
      // Recherche de la méthode de scrutation
      Ind := Lst.IndexOfName('MethodeScrutation');
      if Ind >= 0 then
      begin
        Stg := '';
        Rang := Pos('=', Lst[Ind]);
        if Rang > 0 then Stg := Copy(Lst[Ind], (Rang + 1), (Length(Lst[Ind]) - Rang));
        // Vérifie si la méthode de scrutation est "A des heures précises"
        if Stg = '1' then
        begin
          // Recherche des heures
          Ind := Lst.IndexOf('[HEURES]');
          if Ind < 0 then Ind := Lst.Add('[HEURES]');
          Inc(Ind);
          LigneOk := False;
          if Ind < Lst.Count then
          begin
            Stg := Lst[Ind];
            if (Stg[1] <> '[') and (Stg[Length(Stg)] <> ']') then LigneOk := True;
          end;
          MajHeure(Lst, Ind, 1, Heure1, LigneOk, MajFile);
          Inc(Ind);
          LigneOk := False;
          if Ind < Lst.Count then
          begin
            Stg := Lst[Ind];
            if (Stg[1] <> '[') and (Stg[Length(Stg)] <> ']') then LigneOk := True;
          end;
          MajHeure(Lst, Ind, 2, Heure2, LigneOk, MajFile);
        end;
      end;
    end else with Lst do
      begin
        Add('[CONFIG]');
        Add('MethodeScrutation=1');
        Add('ToxStartAt=07:00');
        Add('ToxEndAt=06:59');
        Add('LaunchWhenApplicationStart=1');
        Add('LaunchWhenWindowsStart=1');
        Add('SearchTimeInMinute=30');
        Add('Langue=FRA');
        Add('ActiveLog=0');
        Add('[HEURES]');
        Add('H0=Appel no 1;' + Heure1 + ';2');
        Add('H1=Appel no 2;' + Heure2 + ';2');
        MajFile := True;
      end;
    if MajFile then Lst.SaveToFile(FileName);
  finally
    Lst.free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// FOToxServeurMajConf : Enregistrement de la configuration du ToxServeur de la caisse de référence
///////////////////////////////////////////////////////////////////////////////////////

procedure FOToxServeurMajConf;
var Caisse, Heure1, Heure2: string;
  QQ: TQuery;
begin
  if not FOVerifSiteTox then Exit;
  Caisse := Trim(GetParamSoc('SO_GCFOCAISREFTOX'));
  if Caisse = '' then Exit;
  QQ := OpenSQL('select GPK_TOXAPPEL1,GPK_TOXAPPEL2 from PARCAISSE where GPK_FERME="-" and GPK_CAISSE="' + Caisse + '"', TRUE);
  if not QQ.EOF then
  begin
    Heure1 := QQ.FindField('GPK_TOXAPPEL1').AsString;
    Heure2 := QQ.FindField('GPK_TOXAPPEL2').AsString;
    FOToxServeurMajHeure(Heure1, Heure2);
  end;
  Ferme(QQ);
end;

end.
