{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/11/2002
Modifié le ... : 27/11/2002
Description .. : Initialisation de la variable V_PGI
Mots clefs ... : FO
*****************************************************************}
unit InitFO;

interface

implementation

uses
  Sysutils, Forms,
  EntPGI,
  Hent1;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/11/2002
Modifié le ... : 30/09/2003
Description .. : Fonction d'initialisation de la variable V_PGI
Mots clefs ... : FO
*****************************************************************}

procedure InitLaVariablePGI;
begin
  HalSocIni := 'CEGIDPGI.INI';
  V_PGI.PGIContexte := [ctxGescom, ctxMode, ctxFO];
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CMFOS5.HLP';
  if ctxMode in V_PGI.PGIContexte then
  begin
   {$IFDEF MODES3}
    TitreHalley := 'Front Office Mode S3' ;
   {$ELSE}
    TitreHalley := 'Front Office Mode S5' ;
   {$ENDIF}
  end else TitreHalley := 'Front Office';
  
  {$IFDEF MODES3}
  NomHalley := 'FOS3';
  V_PGI.LaSerie := S3;
  {$ELSE}
  NomHalley := 'FOS5';
  V_PGI.LaSerie := S5;
  {$ENDIF}
  V_PGI.PortailWeb:='http://utilisateurs.cegid.fr';

  V_PGI.NumVersionBase := 625;
  V_PGI.NumVersion := '5.1.0';
  V_PGI.NumBuild := '1.1';
  V_PGI.DateVersion := EncodeDate(2003, 12, 23);

  Apalatys := 'CEGID';
  Copyright := '© Copyright ' + Apalatys;
  V_PGI.OutLook := TRUE;
  V_PGI.OfficeMsg := TRUE;
  V_PGI.ToolsBarRight := TRUE;
  V_PGI.Date1date2 := FALSE;

  V_PGI.CegidAPalatys := FALSE;
  V_PGI.CegidBureau := TRUE;
  V_PGI.StandardSurDP := True;
  V_PGI.MajPredefini := False;
  V_PGI.QRPdf := True;

  ChargeXuelib;

  V_PGI.CodeProduit := '015';

  if IsCegidDeveloppeur then
    V_PGI.SAV := GetSynRegKey('SAV', V_PGI.SAV, True);
  {$IFDEF VRELEASE}
  V_PGI.SAV := False;
  {$ENDIF}
  V_PGI.VersionDemo := True;
  V_PGI.MenuCourant := 0;
  V_PGI.VersionReseau := True;
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  V_PGI.MenuCourant := 0;
  V_PGI.NumMenuPop := 27;
  V_PGI.RazForme := TRUE;
  V_PGI.BlockMAJStruct := True;
  V_PGI.VoirDateEntree := True;
  if V_PGI.ZoomForm then ZoomScrollBarOn;
end;

initialization
  ProcChargeV_PGI := InitLaVariablePGI;

finalization
  if V_PGI.ZoomForm then ZoomScrollBarOff;

  if IsCegidDeveloppeur then
    SaveSynRegKey('SAV', V_PGI.SAV, True);

end.
