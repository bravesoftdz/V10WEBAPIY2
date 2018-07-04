unit MDispImp;

interface

Uses HEnt1,MenuOLG,Forms,sysutils,HMsgBox, Classes, Controls, HPanel, UIUtil,ImgList,
     Tablette,FE_Main,
     UtilPGI,EntPGI;

Procedure InitApplication ;

type
  TFMenuDisp = class(TDatamodule)
    end ;

Var FMenuDisp : TFMenuDisp ;

implementation

{$R *.DFM}

Uses Ent1
     ,AssistImportProsp,AssistImport;

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
     15: ;
     16: ;
     100 : ; // executer depuis le lanceur

     34101 : BEGIN Assist_Import;  RetourForce:=TRUE;  END ;        // Import negoce
//   60902 : BEGIN AssistRecupTempo; RetourForce := True; END ;     // Import Tempo
     34102 : BEGIN Assist_ImportProspect;  RetourForce:=TRUE;  END ;        // Import Prospect II
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
END ;

procedure InitApplication ;
BEGIN
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


Procedure InitialisationVPGI;
Begin
V_PGI.NumVersionBase:=0 ;    // pas de controle sur version base
Apalatys:='CEGID' ; NomHalley:='GCImport' ;
//TitreHalley:='GC Import   Base '+intToStr(V_PGI.NumVersionBase) ;
TitreHalley:='GC Import' ;
HalSocIni:='CEGIDPGI.ini' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.NumVersion:='5.0.0' ; V_PGI.NumBuild:='3.1'  ;
V_PGI.DateVersion:=EncodeDate(2002,11,19) ;
{$IFDEF CCS3}
V_PGI.LaSerie:=S3 ;
{$ELSE}
V_PGI.LaSerie:=S5 ;
{$ENDIF}

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
V_PGI.StandardSurDP:=True ;
V_PGI.MajPredefini:=True ;
V_PGI.CegidApalatys:=False ;
V_PGI.CegidBureau:=True ;
End;

initialization
ProcChargeV_PGI := InitialisationVPGI ;
end.

