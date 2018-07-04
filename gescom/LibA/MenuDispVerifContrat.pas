unit MenuDispVerifContrat;

interface
Uses HEnt1,MenuOLG,Forms,sysutils,HMsgBox, Classes, Controls, HPanel, UIUtil,ImgList,
     Tablette,FE_Main,EdtEtat,AglInitPlus,AGLInit,Ent1 ;

Procedure InitApplication ;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1: TImageList;
    end ;

Var FMenuDisp : TFMenuDisp ;

implementation

{$R *.DFM}
uses TrtVerifContrat;
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
BEGIN
   case Num of
     10 :  //Apres connection
         begin
         LanceTrtContrat; RetourForce := True;
         FMenuG.ForceClose := True;
         FMenuG.mnQuitClick(Nil);
         end;
     11 : ; //Après deconnection
     12 : ; //Avant connection ou seria
     13 : ;
     74801 : begin LanceTrtContrat; RetourForce := True; end;
     else HShowMessage('2;?caption?;Fonction non disponible : ;W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
BEGIN
   case Num of
     1 : ;
     end ;
END ;


procedure InitApplication ;
BEGIN
FMenuG.OnChargeMag:=ChargeMagHalley ;
//FMenuG.OnMajAvant:=MAJHalleyAvant ;
//FMenuG.OnMajApres:=MAJHalleyApres ;
//FMenuG.OnMajPendant:=MajHalleyPendant ;

FMenuG.OnDispatch:=Dispatch ;

FMenuG.OnMajAvant:=Nil ;
FMenuG.OnMajApres:=Nil ;
FMenuG.SetModules([140],[1]) ;
V_PGI.DispatchTT:=DispatchTT ;
END ;

initialization
Apalatys:='CEGID' ;
NomHalley:='S5' ;
TitreHalley:='Vérification des affaires' ;
HalSocIni:='cegidpgi.ini' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.NumVersion:='1.0' ; V_PGI.NumBuild:='1' ;
V_PGI.NumVersionBase:=561;
V_PGI.DateVersion:=EncodeDate(2001,09,04) ;
V_PGI.LaSerie:=S5 ;
V_PGI.PGIContexte:=[ctxGescom,ctxAffaire,ctxTempo];

V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.ToolsBarRight:=TRUE ;
ChargeXuelib ;
V_PGI.VersionDemo:=False ;
V_PGI.VersionReseau:=True ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=False ;
V_PGI.Halley:=TRUE ;
V_PGI.NiveauAccesConf:=0 ;
V_PGI.MenuCourant:=0 ;

V_PGI.DateVersion:=EncodeDate(2001,08,21);
V_PGI.Decla100:=TRUE;          // sinon pas de maj de structure des tables de version > 100
V_PGI.RazForme:=TRUE;          // Reprise de l'ensemble des forms de la base

V_PGI.CegidBureau:= True;      // permet à l'application d'être pilotée par le lanceur
V_PGI.VersionDev := False;
V_PGI.Date1date2 := False;
V_PGI.MajPredefini := False;   // permet le déversement des standards (chp XX_PREDEFINI)
V_PGI.BlockMajStruct := true;
end.
