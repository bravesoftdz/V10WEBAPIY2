unit InitGCS3;

interface

implementation

uses sysutils,Forms,Hent1,entPGI ;

procedure InitLaVariablePGI ;
BEGIN
HalSocIni:='CEGIDPGI.INI' ;
TitreHalley:='Gestion Commerciale S3' ;
NomHalley:='CGS3' ;
Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CGS35.HLP' ;
V_PGI.NumVersion:='4.2.0' ; V_PGI.DateVersion:=EncodeDate(2002,12,12) ;
V_PGI.NumVersionBase:=595 ; V_PGI.NumBuild:='14' ;

V_PGI.PGIContexte:=[ctxGescom,ctxGRC] ;
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

{$IFDEF VRELEASE}
V_PGI.SAV:=False ;
{$ENDIF}
V_PGI.VersionDemo:=True ;
V_PGI.MenuCourant:=0 ;
V_PGI.VersionReseau:=True ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.MenuCourant:=0 ;
V_PGI.NumMenuPop:=27 ;
V_PGI.LaSerie:=S3 ; V_PGI.EdtMode:=2 ; 
{$IFNDEF AGL545}
V_PGI.Decla100:=TRUE;
{$ENDIF}
V_PGI.RazForme:=TRUE ;
V_PGI.BlockMAJStruct:=True ;
END;

Initialization
ProcChargeV_PGI :=  InitLaVariablePGI ;
end.
