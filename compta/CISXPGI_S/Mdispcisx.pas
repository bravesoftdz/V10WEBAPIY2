{***********UNITE*************************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Dans cette unité, on trouve LA FONCTION qui sera
Suite ........ : appelée par l'AGL pour lancer les options.
Suite ........ : Ainsi que la gestion de l'HyperZoom, de la gestion de
Suite ........ : l'arobase, ...
Suite ........ : C'est aussi dans cette unité que l'on défini le fichier ini
Suite ........ : utilisé, le nom de l'application, sa version, que l'on lance la
Suite ........ : sérialisation, les différentes possibilités d'action sur la mise à
Suite ........ : jour de structure, ...
Mots clefs ... : IMPORTANT;STRCTURE;MENU;SERIALISATION
*****************************************************************}
unit Mdispcisx;

interface
Uses
  Windows, Messages,
  HEnt1,
{$ifdef eAGLClient}
  MenuOLX,MaineAGL,eTablette, UtilEagl,
{$else}
  MenuOLG, FE_Main,
{$endif eAGLClient}
  Forms,sysutils,HMsgBox, Classes, Controls, HPanel,
  hctrls, MajTable,
  ExtCtrls, inifiles, UTOB, ImgList, LicUtil, UDMIMP, UExport;

Procedure InitApplication ;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1  : TImageList;
    Timer1      : TTimer;
    end ;

Var FMenuDisp : TFMenuDisp ;


implementation

{$R *.DFM}
uses
USaveRestore;

Procedure GoMenu(SuperUser : Boolean) ;
BEGIN
  FMenuG.SetFunctions([TraduireMemoire('Scripts'),
                       TraduireMemoire('Restauration'),
                       TraduireMemoire('Sauvegarde'),
                       TraduireMemoire('Génération simple')
                      ],
                      [45100, 45000, 45300, 45200],
                      [46, 40, 41, 45])
END ;

FUNCTION CHARGEMAGHALLEYGG : Boolean ;
Var SuperUser : Boolean ;
BEGIN
Result:=TRUE ;
SuperUser:=(V_PGI.PassWord=CryptageSt(DayPass(Date))) ;
GoMenu(SuperUser) ;
ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette fonction est appellée par l'AGL à chaque sélection
Suite ........ : d'une option de menu, en lui indiquant le TAG du menu en
Suite ........ : question. Ce qui déclenche l'action en question.
Suite ........ : L'AGL lance aussi cette fonction directement afin d'offrir à
Suite ........ : l'application la possibilité d'agir avant ou après la connexion,
Suite ........ : et avant ou après la déconnexion.
Suite ........ : Cette fonction prend aussi en paramètre retourForce et
Suite ........ : SortieHalley. Si RetourForce est à True, alors l'AGL
Suite ........ : remontera au niveau des modules, si SortieHalley est à
Suite ........ : True, alors ...
Mots clefs ... : MENU;OPTION;DISPATCH
*****************************************************************}


Procedure DispatchCom ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var
Numopt           : integer;
Stdpath,Dospath  : string;
BEGIN
     Numopt := Num;
   case Numopt of
     16    : ;
     10    :
     begin
           if V_PGI.StdPath[length (V_PGI.StdPath)] <> '\' then  Stdpath := V_PGI.StdPath+'\CISX'
           else Stdpath := V_PGI.StdPath+'CISX';
           if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then  Stdpath := V_PGI.DosPath+'\CISX'
           else Dospath := V_PGI.DosPath+'CISX';

          InitDB (Stdpath, Dospath);
     end;
     11    : ;
     12    : ;
     13    : V_PGI.SAV:=FALSE ; //Avant deconnection
     15    : ;
     100   : ;
     45000 : SaveRestore(False, PRien);
     45300 : SaveRestore(TRUE, PRien);
     45100 : AglLanceFiche('CP','CPMULSCRIPT','', '', '')  ;
     45200 : ExportDonneesCISX(PRien);

  end ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette procédure permet d'initiliaser certaines référence de
Suite ........ : fonction, les modules des menus gérés par l'application, ...
Suite ........ :
Suite ........ : Cette procédure est appelée directement dans le source du
Suite ........ : projet.
Mots clefs ... : INITILISATION
*****************************************************************}

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
V_PGI.VersionDemo:=FALSE ;
END ;


procedure InitApplication ;
begin
     { Référence à la fonction PRINCIPALE qui permet de lancer les actions en fonction de la
       sélection d'une option des menus.
       }
     FMenuG.OnDispatch        :=DispatchCom ;

     { Référence à une fonction qui est lancée après la connexion à la base et le chargement du dictionnaire }
     FMenuG.OnChargeMag       :=ChargeMagHalleyGG ;

     { Référence à une fonction qui est lancée avant la mise à jour de structure }
     FMenuG.OnMajAvant        :=Nil ;

     { Référence à une fonction qui est lancée pendant la mise à jour de structure }
     FMenuG.OnMajpendant       :=Nil ;
     FMenuG.Timer1.Enabled     := FALSE;

     { Référence à une fonction qui est lancée après la mise à jour de structure }
     FMenuG.OnMajApres          :=Nil ;
     FMenuG.OnChangeModule      :=nil ;
     FMenuG.OnAfterProtec       :=AfterProtec ;

     V_PGI.NumVersionBase       :=833 ;
     V_PGI.NumVersion           :='1.5' ;
     V_PGI.NumBuild             :='0' ;
     V_PGI.DateVersion          :=EncodeDate(2006,06,22) ;
     HalSocIni                  :='CEGIDPGI.INI' ;
     NomHalley                  := 'CISXPGI';
     TitreHalley                := 'Importateur PGI' ;
     V_PGI.LaSerie              := S5;
     Application.HelpFile       := '';
     Application.Title          := TitreHalley ;

END ;

Procedure InitLaVariablePGI;
Begin
Apalatys            :='CEGID' ;
Copyright           :='© Copyright ' + Apalatys ;

HalSocIni           :='CEGIDPGI.INI' ;
V_PGI.OutLook       :=FALSE ;
V_PGI.VersionDemo   :=FALSE ;
V_PGI.MenuCourant   :=0 ;
V_PGI.VersionDEV    :=TRUE ;
V_PGI.ImpMatrix     := True ;
V_PGI.OKOuvert      :=FALSE ;
V_PGI.Halley        :=TRUE ;
V_PGI.OfficeMsg     :=True ;
V_PGI.NumMenuPop    :=27 ;
V_PGI.DispatchTT    :=nil ;
V_PGI.ParamSocLast  :=False ;
V_PGI.RAZForme      :=TRUE ;
V_PGI.NoModuleButtons:=False ;
V_PGI.PGIContexte    :=[ctxCompta] ;
V_PGI.BlockMAJStruct :=True ;
V_PGI.EuroCertifiee  :=False ;

V_PGI.SAV            :=False ;
V_PGI.VersionReseau  :=True ;
V_PGI.CegidAPalatys  :=FALSE ;
V_PGI.CegidBureau    :=TRUE ;
V_PGI.StandardSurDP  :=True ;
V_PGI.MajPredefini   :=False ;
V_PGI.MultiUserLogin :=False ;

V_PGI.OfficeMsg      :=True ;
V_PGI.NoModuleButtons:=FALSE ;
V_PGI.NbColModuleButtons:=1 ;
V_PGI.LaSerie           := S5;
ChargeXuelib ;
end;

Initialization
  InitLaVariablePGI ;

end.
