unit InitGA;

interface

procedure InitLaVariablePGI ;

implementation

uses sysutils,Forms,Hent1,entPGI, Ent1 ;

procedure InitLaVariablePGI ;
BEGIN
  {$IFDEF GIGI}
  NomHalley:='CGIS5' ;
  TitreHalley:='Gestion Interne PGI Expert' ;
  HalSocIni:='CEGIDPGI.ini' ;
  V_PGI.PGIContexte := [ctxGescom,ctxAffaire,ctxScot,ctxPcl];  // PL le 24/10/03 ajout de ctxPcl sur ordre de XP pour les points bleus dans les muls... 
  {$ELSE}
  {$IFDEF CCS3}
  NomHalley:='CGAS3' ;
  TitreHalley:='Gestion d''Affaires S3 ' ;
  {$ELSE}
  NomHalley:='CGAS5' ;
  TitreHalley:='Gestion d''Affaires S5' ;
  {$ENDIF}
  HalSocIni:='CEGIDPGI.ini' ;
  V_PGI.PGIContexte:=[ctxGescom,ctxAffaire,ctxTempo];
  {$ENDIF}

  V_PGI.LaSerie:=S5 ;
  {$IFDEF GIGI}
  V_PGI.SansMessagerie :=True;
  V_PGI.NumVersion:='5.0.1';
  V_PGI.NumBuild:='1.3';
  V_PGI.NumVersionBase:=620;      // Blocage en cas de version < ou message si supérieur
  V_PGI.DateVersion:=EncodeDate(2003,11,18);
  V_PGI.PortailWeb:='http://experts.cegid.fr/home.asp';
  Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CGIS5.HLP' ;
  V_PGI.TabletteHierarchiques := true; //mcd 02/04/2003 pour avoir cette otpion dans appli +GED
  ForceCascadeForms;  // fct générique pour forcer à mode cascade, si pas déjà fait par client et pas ecran 800*600
  {$ELSE}  // GA
  V_PGI.SansMessagerie :=False;
    {$IFDEF CCS3}
    V_PGI.NumVersion:='5.0.0';
    V_PGI.NumBuild:='24.1';
    V_PGI.NumVersionBase:=620;      // Blocage en cas de version < ou message si supérieur
    V_PGI.LaSerie:=S3 ;
    V_PGI.DateVersion:=EncodeDate(2003,11,20);
    {$ELSE}
    V_PGI.NumVersion:='5.0.0';
    V_PGI.NumBuild:='24.1';
    V_PGI.NumVersionBase:=620;      // Blocage en cas de version < ou message si supérieur
    V_PGI.DateVersion:=EncodeDate(2003,11,20);
    {$ENDIF}
  V_PGI.PortailWeb:='http://utilisateurs.cegid.fr';
  Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CGAS5.HLP' ;
  V_PGI.TabletteHierarchiques := False; //mcd 15/09/03 pas OK pour V5  , a initialiser à true si tablette hierarchique et/ou GED
  {$ENDIF}



  V_PGI.CegidBureau:= True;      // permet à l'application d'être pilotée par le lanceur
  V_PGI.VersionDev := False;

  V_PGI.MajPredefini := False;   // permet le déversement des standards (chp XX_PREDEFINI)

  Apalatys := 'CEGID';
  Copyright:='© Copyright ' + Apalatys;
  V_PGI.OutLook:=TRUE;
  V_PGI.OfficeMsg:=TRUE;
  V_PGI.ToolsBarRight:=TRUE;
  V_PGI.Date1date2 := False;
  V_PGI.CegidAPalatys:=FALSE;
  V_PGI.CegidBureau:=TRUE;
  V_PGI.StandardSurDP:=True;
  V_PGI.MajPredefini:=False;

  V_PGI.CegidUpdateServerParams:= 'www.update.cegid.fr'; //mcd 23/10/03 pour intilaiser telassistance
  ChargeXuelib;
  {$IFDEF GIGI}
  V_PGI.CodeProduit:='011';  // code produit GI pour WebAcces
  {$ELSE}
  V_PGI.CodeProduit:='009';  // code produit GI pour WebAcces
  {$ENDIF}

  {$ifdef DEBUG}
  V_PGI.SAV:=True ; // doir obligatoirement être après CHargeXuelib
  V_PGI.VersionDemo:=True ;
  {$ELSE}
  V_PGI.SAV:=false ;
  V_PGI.VersionDemo:=False ;
  {$ENDIF}
  {$ifdef MEMCHECK}
  V_PGI.SAV:=True ; // pour en devlpt etre en viosn SAV et voir les erreurs de champs
  {$ENDIF}

  V_PGI.MenuCourant:=0;
  V_PGI.VersionReseau:=True;     // permet de rendre certaines actions mono-utilisateur (fcts blocagemonoposte, existeblocage...)
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert:=FALSE;
  V_PGI.Halley:=TRUE;
  V_PGI.NumMenuPop:=27;
  V_PGI.RazForme:=TRUE;          // Reprise de l'ensemble des forms de la base
  V_PGI.BlockMajStruct := true;  // Bloque la maj de structure dans le prg (passer par pgimajver)

END;

Initialization
ProcChargeV_PGI :=  InitLaVariablePGI ;

end.
 