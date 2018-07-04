unit InitGC;

interface

implementation

uses sysutils,Forms,Hent1,entPGI,entGC ;

procedure InitLaVariablePGI ;
BEGIN
{$IFDEF TESTCONTRAT}
    HalSocIni:='TEST3.INI' ;
{$ELSE}
    HalSocIni:='CEGIDPGI.INI' ;
{$ENDIF}

{$IFDEF CCS3}
    TitreHalley:='Gestion Commerciale S3' ;
    NomHalley:='CGS3' ;
    V_PGI.LaSerie:=S3 ;
{$ELSE}
  {$IFDEF ECGS5DEMO}
    TitreHalley:='Gestion Commerciale S5' ;
    NomHalley:='ECGS5DEMO' ;
    V_PGI.LaSerie:=S5 ;
  {$ELSE}
     {$IFDEF TESTSIC}
     TitreHalley:='TEST Gestion Commerciale S5' ;
     NomHalley:='ECGS5SIC' ;
     V_PGI.LaSerie:=S5 ;
     {$ELSE}
     TitreHalley:='Gestion Commerciale S5' ;
     NomHalley:='CGS5' ;
     V_PGI.LaSerie:=S5 ;
     {$ENDIF}
  {$ENDIF}
{$ENDIF}

    Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CGS35.HLP' ;

    V_PGI.NumVersion:='5.0.0' ;
    V_PGI.PortailWeb:='http://utilisateurs.cegid.fr';
    V_PGI.NumVersionBase:=620; V_PGI.NumBuild:='19.2' ;
    V_PGI.DateVersion:=EncodeDate(2003,11,20) ;
    V_PGI.PGIContexte:=[ctxGescom,ctxGCAFF] ;


{$IFDEF MODE}
//    V_PGI.PGIContexte:=[ctxGescom, ctxMode, ctxGRC] ; Contexte GRC géré ds AfterProtecMode
    V_PGI.PGIContexte:=[ctxGescom, ctxMode] ;
    Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CMBOS5.HLP' ;

  {$IFDEF MODES3}
    TitreHalley:='Mode S3';
    NomHalley:='BOS3' ;
    V_PGI.LaSerie:=S3 ;
  {$ELSE}
    TitreHalley:='Mode S5';
    NomHalley:='BOS5' ;
  {$ENDIF}
    V_PGI.NumVersionBase:=620 ;
    V_PGI.NumVersion:='5.0.0' ;                  // A faire en changt version
    V_PGI.NumBuild:='8.1' ;
    V_PGI.DateVersion:=EncodeDate(2003,11,20);  // A faire en changt version
{$ENDIF}

Apalatys:='CEGID' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.ToolsBarRight:=TRUE ;
V_PGI.Date1date2:=FALSE;

V_PGI.CegidAPalatys:=FALSE ;
V_PGI.CegidBureau:=TRUE ;
V_PGI.StandardSurDP:=True ;
V_PGI.MajPredefini:=False ;
V_PGI.QRPdf:=True ;

ChargeXuelib ;

{$IFDEF MODE}
V_PGI.CodeProduit := '015' ;
{$ENDIF}

if FileExists(ExtractFilePath(Application.ExeName)+'INI_4B8D120B-1877-4E78-A817-331A23A35209.ini') then V_PGI.SAV:=True ;

V_PGI.VersionDemo:=True ;
V_PGI.MenuCourant:=0 ;
V_PGI.VersionReseau:=True ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.MenuCourant:=0 ;
V_PGI.NumMenuPop:=27 ;
V_PGI.RazForme:=TRUE;
V_PGI.BlockMAJStruct:=True ;
V_PGI.VoirDateEntree:=True;
END;

Initialization
ProcChargeV_PGI :=  InitLaVariablePGI ;
end.
