{***********UNITE*************************************************
Auteur  ...... : D. CARRET
Créé le ...... : 29/11/2002
Modifié le ... : 29/11/2002
Description .. : Initialisation de la variable V_PGI
Mots clefs ... : PCP
*****************************************************************}
unit InitPCP;

interface

{$IFNDEF AGL550B}
procedure InitLaVariablePGI;
{$ENDIF}

implementation

uses
  Sysutils, Forms,
  {$IFDEF AGL550B}
  EntPGI,
  {$ENDIF}
  UtilNomade, Hent1,ParamSoc;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. CARRET
Créé le ...... : 29/11/2002
Modifié le ... : 29/11/2002
Description .. : Fonction d'initialisation de la variable V_PGI
Mots clefs ... : PCP
*****************************************************************}

procedure InitLaVariablePGI;
begin
  HalSocIni := 'CEGIDPGI.INI';

	V_PGI.PGIContexte:=[ctxGescom,ctxAffaire,ctxBTP,ctxGRC];
//  if GetEnvPop = '1' then V_PGI.PGIContexte := V_PGI.PGIContexte + [ctxMode];
  TitreHalley := 'Mobilité';
{$IFDEF BUSINESSPLACE}
  NomHalley := 'CBPOS5';
  V_PGI.LaSerie := S5;
	Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CBPOS5.HLP';
{$ELSE}
  NomHalley := 'CBPOS3';
  V_PGI.LaSerie := S3;
	Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CBPOS3.HLP';
{$ENDIF}

  V_PGI.NumVersionBase := 714;
  V_PGI.NumVersion     := '6.5.3';
  V_PGI.DateVersion    :=EncodeDate(2006,05,18) ;
  V_PGI.NumBuild       := '001.015';
  V_PGI.CodeProduit    := '032';

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

  {$IFNDEF AGL545}
//  V_PGI.Decla100 := TRUE;
  {$ENDIF}
  V_PGI.RazForme := TRUE;
  V_PGI.BlockMAJStruct := True;
  V_PGI.VoirDateEntree := True;
end;

{$IFDEF AGL550B}
initialization
  ProcChargeV_PGI := InitLaVariablePGI;
  {$ENDIF}

end.
