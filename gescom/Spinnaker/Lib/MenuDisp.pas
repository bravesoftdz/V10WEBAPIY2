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

uses Windows,HEnt1, Ent1, EntPGI,LicUtil, ED_TOOLS,
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
    MajTable,
    Messages,
    UTraiteSaisies;


procedure InitApplication;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1: TImageList;
  end;

var
  FMenuDisp: TFMenuDisp;

  procedure LanceMenuAuto;

implementation

{$R *.DFM}

uses MenuSpec;        //Pour CGEP

procedure LanceMenuAuto;
var
Nom, St, Value    : string;
i                 : integer;
Connect           : Boolean;
begin
  InitMoveProgressForm(Application.MainForm,Application.Title, 'Traitement en cours, veuillez patienter...', 200, False, true) ;

      V_PGI.Halley:=TRUE ;
      V_PGI.NumVersion:='9.0.0' ;
      V_PGI.Versiondev:=FALSE ;
      VH^.ModeSilencieux := True;
      V_PGI.MultiUserLogin := TRUE;
      Connect := FALSE;

  V_PGI.UserLogin  := 'LSE';
  V_PGI.PassWord   := CryptageSt('ESL');
  V_PGI.DateEntree := Date();
  V_PGI.CurrentAlias := 'VERREMETAL';

      for i:=1 to ParamCount do
      begin
        St:=ParamStr(i) ;
        Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
        Value:=UpperCase(Trim(St)) ;
        //Paramètres de connexion
        if Nom='/USER'     then V_PGI.UserLogin :=Value;
        if Nom='/PASSWORD' then V_PGI.PassWord:=Value;
        if Nom='/DATE'     then BEGIN  V_PGI.DateEntree:=StrToDate(Value) ; END ;
        if Nom='/BASE'     then BEGIN  V_PGI.CurrentAlias:=Value ; END ;
       end;

      if (not Connect) then
      begin
           Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL,NIL);
      end;

//      if connect then LanceRecupChantiers;

      Application.ProcessMessages;
       if Connect then begin
{$IFNDEF EAGLCLIENT}
          Logout ;
          DeconnecteHalley ;
{$ENDIF}
       end
       else
       begin
                  // Fermeture de l'application
                  Application.ProcessMessages;
                  if FMenuG <> nil then
                  begin
                       FMenuG.ForceClose := True ;
                       PostMessage(FmenuG.Handle,WM_CLOSE,0,0) ;
                       FMenuG.Close ;
                  end;
                  VH^.STOPRSP:=TRUE ;
                  Exit;
       end;
       VH^.STOPRSP:=TRUE ;
       SourisNormale ;

  FiniMoveProgressForm ;

end;

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
    10: ;//Apres connection
    11: ; //Après deconnection
    12: ; //Avant connection ou seria
    13: ; //Avant deconnection
    15: ; //Avant formshow
    16 :;
    324001 :
        LanceRecupDocument ('DBT');
    324002 :
        LanceRecupDocument ('CF');
    324003 :
        LanceRecupDocument ('CONSOS');
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
  {$IFDEF EXEMPLE_UTILISATION}
  { Valeur par défaut si 'sf' est inconnue }
  Result := '';

  if sf = 'FONCTION1' then
  begin
    { Pour cette fonction, il doit y avoir 2 paramètres }
    Param1 := ReadTokenPipe(Sp, ',');
    Param2 := ValeurI(ReadTokenPipe(Sp, ','));

    Result := AppCalculValeur(Param1, Param2 * 100);
  end
  else if sf = 'FONCTION2' then
  begin
  end;
  {$ENDIF}
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
  {$IFDEF EXEMPLE_UTILISATION}
  case Qui of
    80: AglLanceFiche('XXX', 'XFICHEX1', '', Quoi, 'ACTION=CONSULTATION');
    81: AglLanceFiche('XXX', 'XFICHEX2', '', Quoi, 'ACTION=CONSULTATION');
  end;
  Screen.Cursor := crDefault;
  {$ENDIF}
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
var
  i, Qui: integer;
begin

  i := Pos(';', Quoi);

  { Qui est le code de l'HyperZoom }
  Qui := StrToInt(Copy(Quoi, 1, i - 1));

  { Quoi est le code de l'élément }
  Quoi := Copy(Quoi, i + 1, Length(Quoi) - i);

    if (Qui>=900) and (Qui<=999) and vmDisp_ZoomEdt( Qui,Quoi) then Exit ;   //Pour CGEP

  ZoomEdt(Qui, Quoi);

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
  {$IFDEF EXEMPLE_UTILISATION}
  { Valeur par défaut si 'S' est inconnue }
  Result := '';

  if S = 'VH^.CHAMP3' then Result := CProcCalculChampVH(S);
  {$ENDIF}
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
  {$IFDEF EXEMPLE_UTILISATION}
var
  Year, Month, Day: WORD;
  {$ENDIF}
begin
  {$IFDEF EXEMPLE_UTILISATION}
  { Valeur par défaut si 'S' est inconnue }
  Result := Now;

  if S = 'D' then { Date de Début d'exercice }
  begin
    Result := ComptaDateDebutExercice;
  end
  else if S = 'F' then { Date de fin d'exercice }
  begin
    Result := ComptaFinDebutExercice;
  end
  else if S = 'G' then { Date de début du mois en cours }
  begin
    DecodeDate(Date, Year, Month, Day);
    Result := EncodeDate(Year, Month, 1);
  end;
  {$ELSE}
  result := Now ;
  {$ENDIF}
end;

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
begin
  { Référence à la fonction de gestion de l'hyperzoom }
  ProcZoomEdt := ZoomEdtEtat;

  { Référence à la fonction de gestion @ dans les états }
  ProcCalcEdt := CalcArobaseEtat;

  { Référence à la fonction qui permet dans les tablettes de faire l'interprétation
    d'un condition de la forme WHERE G_GENERAUX="VH^.CHAMP3" ... , pour l'instant
    utilisé que dans la comptabilité. }
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  ProcGetVH := GetVS1;
  {$ENDIF}

  { Référence à une fonction qui permet d'avoir des dates supplémentaires particulières
    Début d'exercice, exercice suivant, ... }
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  ProcGetDate := GetDate;
  {$ENDIF}

  { Référence à la fonction PRINCIPALE qui permet de lancer les actions en fonction de la
    sélection d'une option des menus.
    }
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

  { Renseigne les n° de modules ( menu ) que l'application doit gérer }
  FMenuG.SetModules([324], [99]);

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
  NomHalley := 'SPINNAKER';

  { Ce nom apparaît en bas à gauche de l'application }
  TitreHalley := 'Spécifique SPINNAKER';

  { Précise le nom du fichier ini utilisé par l'application. Normalement CEGIDPGI.INI }
  HalSocIni := 'cegidpgi.ini';

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
V_PGI.NumBuild:='001.001';
V_PGI.CodeProduit:='033' ;
V_PGI.DateVersion:=EncodeDate(2016,03,09);
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


