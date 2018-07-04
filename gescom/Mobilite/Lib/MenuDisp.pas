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
unit MenuDisp;
{$IFDEF HVCL}sdsdsd{$ENDIF}

interface

uses Windows,HEnt1, LicUtil, 
  {$IFDEF eAGLClient}
    utileagl,
    MenuOLX,
    MaineAGL,
    eTablette,
  {$ELSE}
    MenuOLG,
    uedtcomp,
    Tablette,
    FE_Main,
    EdtEtat,
  {$ENDIF eAGLClient}
    Forms,
    sysutils,
    HMsgBox,
    Classes,
    Controls,
    HPanel,
    Hctrls,
    UIUtil,
    ImgList,
    AglInitPlus,
    AGLInit,
    EntGC,
    UTOF_VideInside,
    inifiles,
    Messages;


procedure InitApplication;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1: TImageList;
  end;

var
  FMenuDisp: TFMenuDisp;
	GCCodeDomaine : string;
  GCCodesSeria : array[1..9] of string;
  GCTitresSeria : array[1..9] of Hstring;
  VersionInterne : boolean;


implementation
uses Ent1;

{$R *.DFM}

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
procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
begin

  case Num of
    10: begin
    			SaisirHorsInside('BTAPPELPARRESS');
          Fmenug.Quitter;
          Exit;
    		end;

    11: ; //Après deconnection
    12: ; //Avant connection ou seria
    13: ; //Avant deconnection
    15: ; //Avant formshow
    16 : ;
    324001 :;
    324002 :;
  else HShowMessage('2;?caption?;Fonction non disponible : ;W;O;O;O;', TitreHalley, IntToStr(Num));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette fonction est appelée par l'AGL a chaque fois qu'un
Suite ........ : utilisateur click sur un bouton de paramètrage d'un combo.
Suite ........ : Le paramètre NUM est la valeur qui est affectée dans la
Suite ........ : zone Tag des tablettes.
Mots clefs ... : TABLETTE;COMBO
*****************************************************************}
procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT, Range: string);
begin
  case Num of
    80: AglLanceFiche('XXX', 'YYYYY', '', LeQuel, '');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : cette procédure est appelée par l'AGL à chaque fois que
Suite ........ : dans un état, il rencontre une formule, un champ, ...
Suite ........ : contenant @FONCTIONX([P1],[P2],...).
Suite ........ : L'AGL passe en paramètre à cette fonction, le nom de la
Suite ........ : fonction dans "sf" (en l'occurence "FONCTIONX" dans
Suite ........ : l'exemple si-dessus) et les paramètres séparés par des
Suite ........ : virgules dans "sp".
Mots clefs ... : @;FORMULE;FONCTION;ETAT
*****************************************************************}
function CalcArobaseEtat(sf, sp: string): variant;
begin
  //
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette procédure est appelée par ZoomEdtEtat, et elle
Suite ........ : permet de lancer les fiches/paramtablette/... en fonction du
Suite ........ : code hyperzoom et de LeQuel.
Mots clefs ... : ZOOM;HYPERZOOM
*****************************************************************}
procedure ZoomEdt(Qui: integer; Quoi: string);
begin
  //
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette procédure est appelée par l'AGL quand l'utilisateur
Suite ........ : double-click sur un aperçu d'un état configuré pour la
Suite ........ : gestion de l'HyperZoom.
Suite ........ : Le paramètre "Quoi" contient le code du zoom ( valeur
Suite ........ : contenue dans la tablette TTZOOMEDT ), ainsi que la zone
Suite ........ : permettant de se positionner sur l'élément en question. Les
Suite ........ : deux valeurs étant séparées par un point-virgule.
Mots clefs ... : ZOOM;HYPERZOOM
*****************************************************************}
procedure ZoomEdtEtat(Quoi: string);
begin
  //
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette function est appelée à chaque fois que l'AGL
Suite ........ : rencontre une condition de la forme
Suite ........ : WHERE XX_YYYYYY="VH^.CHAMPX"
Suite ........ : dans les conditions des tablettes.
Mots clefs ... : TABLETTE;CONDITION;VH;COMPTA
*****************************************************************}
function GetVS1(S: string): string;
begin
  //
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette fonction permet à l'application d'utiliser des
Suite ........ : "raccourcies" pour déterminer les dates dans les écrans de
Suite ........ : saisie.
Suite ........ : Ainsi, si l'utilisateur saisie D, l'AGL va appeler cette fonction
Suite ........ : qui va lui retourner ( dans l'exemple ), la date de
Suite ........ : début de l'exercice  courant, ...
Mots clefs ... : DATE;RACCOURIE;SAISIE
*****************************************************************}
function GetDate(S: string): TDateTime;
begin
  //
end;


Procedure AfterProtec ( sAcces : String ) ;
BEGIN
  V_PGI.nolock := true;
  V_PGI.CADEnabled := false;
  If V_PGI.NoProtec then V_PGI.VersionDemo:=TRUE
                    else V_PGI.VersionDemo:=False;
  if (sAcces[1]='X') then   VH_GC.SeriaMobilite:= TRUE
                     else   VH_GC.SeriaMobilite:= FALSE;

  if not VH_GC.SeriaMobilite then V_PGI.VersionDemo:=TRUE;
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
procedure InitApplication;
var IniApplication : Tinifile;
    Exerepert : string;
begin
  VersionInterne := false;
  Exerepert := ExtractFilePath (Application.ExeName); // repertoire de l'application

  IniApplication := TiniFile.create (IncludeTrailingBackslash (Exerepert)+'GAMME-PRODUIT.LSE');
  TRY
  	VersionInterne := IniApplication.ReadString('Application','KIKONEST','')='CNOUS';
  FINALLY
  	IniApplication.Free;
  end;

  { Référence à la fonction de gestion de l'hyperzoom }
  ProcZoomEdt := ZoomEdtEtat;

  { Référence à la fonction de gestion @ dans les états }
  ProcCalcEdt := CalcArobaseEtat;

  FMenuG.OnDispatch := Dispatch;

  { Référence à une fonction qui est lancée après la connexion à la base et le chargement du dictionnaire }
  FMenuG.OnChargeMag := ChargeMagHalley ;

  { Référence à une fonction qui est lancée avant la mise à jour de structure }
  FMenuG.OnMajAvant := nil;

  { Référence à une fonction qui est lancée pendant la mise à jour de structure }
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  FMenuG.OnMajpendant := nil;
  {$ENDIF}

  { Référence à une fonction qui est lancée après la mise à jour de structure }
  FMenuG.OnMajApres := nil;
  V_PGI.NoProtec := false;

  { Renseigne les n° de modules ( menu ) que l'application doit gérer }
  FMenuG.SetModules([0], [0]);
  FMenuG.OutLook.Visible := false;
  FMenuG.OutLook.Visible := false;

  V_PGI.LaSerie:=S5 ;
	if VersionInterne then V_PGI.VersionDemo:=false;
  if not VersionInterne then
  begin
    GCCodeDomaine := '14906010'; //
    GCCodesSeria [1] :='14907090'; GCTitresSeria [1] := 'Mes interventions';
    FMenuG.OnAfterProtec:=AfterProtec ;
    FMenuG.SetSeria(GCCodeDomaine,GCCodesSeria,GCTitresSeria) ;
  end;


  { Référence à une fonction qui permet de lancer des actions dans le cas ou l'on autorise
  le paramètrage d'un combo dans une fiche de saisie.
  Il faut aussi dans ce cas renseigner le n° de tag au niveau de la tablette correspondante. }
  V_PGI.DispatchTT := DispatchTT;
end;

initialization
  { Ce nom apparaît en haut à gauche dans la fenêtre Inside }
  Apalatys := 'LSE';

  { Ce nom apparaît dans le caption de l'application, c'est aussi le clef dans la BDR pour le
    stockage des informations préférence de l'application. }
  NomHalley := 'Mobilité';

  { Ce nom apparaît en bas à gauche de l'application }
  TitreHalley := 'Mes Interventions';

  { Précise le nom du fichier ini utilisé par l'application. Normalement CEGIDPGI.INI }
  HalSocIni := 'CEGIDPGI.ini';

  { Ce nom apparaît en bas à gauche de l'application }
  Copyright := '© Copyright L.S.E';

V_PGI.PGIContexte:=[ctxGescom,ctxAffaire,ctxBTP,ctxGRC];
V_PGI.StandardSurDp := False;  // indispensable pour l'aiguillage entre le DP et les bases dossiers (gamme Expert), voir SQL025
V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.ToolsBarRight:=TRUE ;
ChargeXuelib ;
V_PGI.VersionDemo:=True ;
V_PGI.MenuCourant:=0 ;
V_PGI.VersionReseau:=True ;
V_PGI.NumVersion:='9.1.0' ;
V_PGI.NumVersionBase:=998 ;
V_PGI.NumBuild:='000.067';
V_PGI.CodeProduit:='033' ;
V_PGI.DateVersion:=EncodeDate(2014,06,17);
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.BlockMAJStruct:=TRUE ;
V_PGI.PortailWeb:='http://www.lse.fr/' ;
V_PGI.CegidApalatys:=FALSE ;
V_PGI.QRMultiThread:=TRUE;
V_PGI.MenuCourant:=0 ;
V_PGI.RazForme:=TRUE;
V_PGI.NumMenuPop:=27 ;
V_PGI.CegidBureau:= True;
V_PGI.LookUpLocate:= True;   // utile pour les ellipsis_click sur THEDIT
V_PGI.QRPDFOptions:=4;        // pour tronquer les libellés en impression
V_PGI.QRPdf:=True;           // Mode PDF Par défaut pour totalisations situations


end.


