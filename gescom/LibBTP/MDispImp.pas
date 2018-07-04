unit MDispImp;

interface

Uses HEnt1,MenuOLG,Forms,sysutils,HMsgBox, Classes, Controls, HPanel, UIUtil,ImgList,
     Tablette,FE_Main, Ent1, EntGC, EntRT,LicUtil,
     UtilPGI;

Procedure InitApplication ;

type
  TFMenuDisp = class(TDatamodule)
    end ;

Var FMenuDisp : TFMenuDisp ;

implementation

{$R *.DFM}

Uses AssistImportBtp;

function VerifSuperviseur : boolean ;
begin
Result:=True;
if not V_PGI.Superviseur then
   begin
   PGIBox('Vous n''avez pas les droits pour accéder à cette application','Attention') ;
   Result:=False ; FMenuG.ForceClose:=True ; FMenuG.Close ;
   end ;
end;


Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
BEGIN
   case Num of
     10 : VerifSuperviseur ;
     11 : ; //Après deconnection
     12 : ; //Avant connection ou seria
     13 : ;
     16 : ; // entre connexion et chargemodule
     100 : ; // executer depuis le lanceur

     34101 : BEGIN Assist_ImportBtp (V_PGI.PassWord = CryptageSt(DayPass(Date)));  RetourForce:=TRUE;  END ;        // Import negoce
     else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
BEGIN
   case Num of
     1 : ;
     end ;
END ;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
UpdateSeries ;
FMenuG.RenameItem(34101,TraduireMemoire ('Import BTP')) ;
FMenuG.RemoveItem(34102) ;
END ;

procedure InitApplication ;
BEGIN
if Not IsLibrary then ;
InitLaVariableHalley ;
{$IFDEF GCGC}
InitLaVariableGC ;
{$ENDIF}
{$IFDEF PAIEGRH}
InitLaVariablePaie ;
{$ENDIF}
{$IFDEF GRC}
InitLaVariableRT ;
{$ENDIF}
{$IFDEF BTP}
InitLaVariableRT ;
{$ENDIF}

FMenuG.OnMajAvant:=Nil;
FMenuG.OnMajApres:=Nil;
FMenuG.OnMajPendant:=Nil ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChargeMag:=ChargeMagHalley ;
FMenuG.OnChangeModule:=AfterChangeModule ;
FMenuG.SetModules([34],[92]) ;
FMenuG.bSeria.Visible:=False ;
V_PGI.DispatchTT:=DispatchTT ;
END ;


initialization
V_PGI.NumVersionBase:=0 ;    // pas de controle sur version base
Apalatys:='LSE' ; NomHalley:='Import Gamme II' ;
//TitreHalley:='GC Import   Base '+intToStr(V_PGI.NumVersionBase) ;
TitreHalley:='Import gamme II BTP' ;
{$IFDEF DEMOBTP}                          
HalSocIni:='BTPIMPORT.ini' ;
{$ELSE}
HalSocIni:='CEGIDPGI.ini' ;
{$ENDIF}
Copyright:='© Copyright ' + Apalatys ;
//{$IFDEF CCS3}
V_PGI.LaSerie:=S3 ;
//{$ELSE}
//V_PGI.LaSerie:=S5 ;
//{$ENDIF}

V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.ToolsBarRight:=TRUE ;
ChargeXuelib ;
V_PGI.AlterTable:=True ;
V_PGI.VersionDemo:=TRUE ;
V_PGI.VersionReseau:=true ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.MenuCourant:=0 ;
V_PGI.RazForme:=TRUE;
V_PGI.StandardSurDP:=False ;
V_PGI.MajPredefini:=True ;
V_PGI.CegidApalatys:=False ;
V_PGI.CegidBureau:=True ;
V_PGI.NumVersion:='9.1.0' ;
V_PGI.NumVersionBase:=998 ;
V_PGI.NumBuild:='000.001' ;
V_PGI.DateVersion:=EncodeDate(2012,07,17);
end.
