{***********UNITE*************************************************
Auteur  ...... : ROHAULT Régis
Créé le ...... : 14/02/2003
Modifié le ... : 06/08/2004
Description .. : Source TOM de la TABLE : JOURNAL (JOURNAL)
Suite ........ : GCO - 05/03/2004
Suite ........ : Ajout de l'événement OnChangingPages du TPageControl
Suite ........ : CA - 06/08/2004 : ajout du avertirmultitable pour indiquer
Suite ........ : qu'une modification a été faite dans la table JOURNAL
Mots clefs ... : TOM;JOURNAL
*****************************************************************}

unit CPJOURNAL_TOM;

//================================================================================
// Interface
//================================================================================
interface

uses
  StdCtrls,
  Controls,
  Classes,
  Graphics, // clWindow, clBtnFace, clRed
  Windows, // VK_ ,
  Buttons, // TBitBtn
  Menus, // TPopupMenu
  ExtCtrls, // TPanel
  HCtrls, // ReadTokenSt, ReadTokenI
{$IFNDEF CCADM}
  ImpPrefG, // FicheImportPlanRef
{$ENDIF CCADM}
  LookUp, // LookUpValueExist
  HMsgBox, // ExJaiLeDroitConcept, THMsgBox
  ULibWindows, //IIF
  ULibAnalytique, //RecherchePremDerAxeVentil
  SAISUTIL, //caseNatJal
  HCompte, //ExisteH
  HRichOle, // THRichEditOle
  MajTable,
{$IFDEF VER150} {D7}
  Variants,
{$ENDIF}
{$IFDEF EAGLCLIENT}
  eFiche, // TTFiche
  MaineAGL, // AGLLanceFiche
  UtileAGL, // TNavigateBtn
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  edtREtat,{JP 28/10/05 : LanceEtat}
  Fiche,
  FichList,
  Fe_main,
(*JP 28/10/05 : Branchement de l'état AGL
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
  QRPlanJo,
{$ENDIF}
{$ENDIF}
*)

{$ENDIF}
  utilpgi,
  HTB97, // TToolBarButton97,
  forms,
  sysutils,
  ComCtrls,
  HEnt1, ent1,
  AglInit,       // ActionToString
  ParamSoc, // GetParamsoc
{$IFDEF COMPTA}
  CUMMENS,
{$ENDIF}
  UTOM,
  UTob  ,UentCommun
  ;

//==================================================
// Externe
//==================================================

procedure FicheJournal(Q: TQuery; Axe, Compte: string; Comment: TActionFiche; QuellePage: Integer);

function EstDansAnalytiq(St: string): Boolean;
function EstDansEcriture(St: string): Boolean;
function EstDansSociete(St: string): Boolean;
function EstDansSouche(St: string): Boolean;
function EstDansPiece(St: string): Boolean; // Modif Fiche 10856 SBO
function EstDansParamPiece(St: string): Boolean; // Modif Fiche 10856 SBO
function EstDansParamPieceCompl(St: string): Boolean; // Modif Fiche 10856 SBO

//==================================================
// Definition de class
//==================================================
type
  TOM_JOURNAL = class(TOM)
  public
  private
    // MsgBox
    MsgBox: THMsgBox;
    // popup Menu
    POPZ: TPopUpMenu;
    // PageControl
    Pages: TPageControl;
    // parametre
    Lequel, LeCompte, LesModif: string;
    Mode: TActionFiche;
    AvecMvt: Boolean;
    fAvertirTable: boolean;
    premier_axe: integer;
    //SG6 23/12/2004 dsns le mode croisaxe, numéro du premier axe ventilé
  public
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private

    FTobCumuls              : TOB ;                    // Contient les infos de cumuls locales en mode multisoc
    FAGLCloseQuery          : TCloseQueryEvent ;

    procedure InitFolio;
    procedure InitMsgBox;
    procedure FormateLesMontants;
    procedure FormateLesDates;
    procedure ActivationControl;

    procedure FormCloseQuery             (Sender: TObject; var CanClose: Boolean);

    // evenement
    procedure BImprimerClick(Sender: TObject);
    procedure J_EFFETClick(Sender: TObject);
    // menu Zoom
    procedure BMenuZoomEnter(Sender: TObject);
    procedure BZecrimvtClick(Sender: TObject);
    procedure BCumulClick(Sender: TObject);
    procedure BJALDIVClick(Sender: TObject);
    procedure BJALCPTClick(Sender: TObject);
    procedure BJALPERClick(Sender: TObject);
    procedure BJALCENTClick(Sender: TObject);
    // on change
    procedure J_NATUREJALChange(Sender: TObject);
    procedure J_JOURNALOnChange(Sender: TObject);
    procedure J_MODESAISIEChange(Sender: TObject);
    procedure OnChangingPages(Sender: TObject; var AllowChange: Boolean);
    // on exit
    procedure J_JOURNALExit(Sender: TObject);
    procedure CptRegulExit(Sender: TObject);
    // nouvel enreg
    procedure OnClickBInsert(Sender: TObject);
    procedure OnMouseDownJ_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

    function EnregOK: boolean;
    function CodExist: boolean;
    function EstJalAnal: boolean;

    procedure AutoriseOuNonLesChamps;
    procedure IsMouvemente;
    procedure ClearChampsSurEnregOk;
    function CrtlCpteDeRegul: Boolean;
    function ExisteLeCompte(LeWhere: string): Boolean;
    function FaitLeWhere(St: string): string;
    function ContrePartieValide: Boolean;
    function ContrePartieDevise: Boolean;
    procedure DevalideLexoSurCreatJal;

    function DansUnCpte(Cpte1, Cpte2: TObject): Boolean;
    procedure MsgCpteRegul(Sender: TOBject; i: Byte);
    function CaractereValide(Sender: TObject): Boolean;
    function ChaineCpteValide(Sender: TObject): Boolean;
    function CpteDeRegulValide(Sender: TObject): Boolean;

    procedure AfficheCumuls ;
    function  GetChampsMS( vStChamp : String ) : double ;   // Retourne le cumul demandé soit issu de la table commune, soit de la table locale en multisoc
    procedure ChargeCumulsMultiSoc ;                     // Renseigne FTobCumuls avec les valeurs locales en mode multisoc

  end;

  //================================================================================
  // Implementation
  //================================================================================
implementation

uses
{$IFDEF MODENT1}
  CPProcGen,
  CPTypeCons,
  CPProcMetier,
{$ENDIF MODENT1}

{$IFDEF eAGLCLIENT}
  MenuOLX
{$ELSE}
  MenuOLG
{$ENDIF eAGLCLIENT}
{$IFNDEF IMP}
{$IFNDEF PGIIMMO}
{$IFNDEF CCADM}
  , ZECRIMVT_TOF
{ FQ 20370 BVE 02.10.07 }
  ,CPMULANA_TOF // MultiCritereAnaZoom
  ,critedt // ClassCritEdt
{ END FQ 20370 }
  //Zecrimvt;
{$ENDIF CCADM}
{$ENDIF}
{$ENDIF}
  ,UFonctionsCBP
  ;

//==================================================
// Definition des Constante
//==================================================
// les tag de pages ;)
const
  PGeneral: integer = 0;
  PCOMPLEMENT: integer = 1;
  PInf: integer = 2;
  PReg: integer = 3;
  PTreso: integer = 4;
  PSAISIE: integer = 5;

  //==================================================
  // fonctions hors class
  //==================================================
  {***********A.G.L.***********************************************
  Auteur  ...... : BPY
  Créé le ...... : 31/03/2003
  Modifié le ... :   /  /
  Description .. :
  Mots clefs ... :
  *****************************************************************}

function EstDansAnalytiq(St: string): Boolean;
begin
  result := Presence('ANALYTIQ', 'Y_JOURNAL', St);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function EstDansEcriture(St: string): Boolean;
begin
  result := Presence('ECRITURE', 'E_JOURNAL', St);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function EstDansSociete(St: string): Boolean;
var
  StLoc: string;
begin
  result := true;
  StLoc := GetParamsocSecur('SO_JALOUVRE','');
  if (StLoc = St) then
    exit;
  StLoc := GetParamsocSecur('SO_JALFERME','');
  if (StLoc = St) then
    exit;
  result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function EstDansSouche(St: string): Boolean;
begin
  result := Presence('SOUCHE', 'SH_JOURNAL', St);
end;

function EstDansPiece(St: string): Boolean; // Modif Fiche 10856 SBO
begin
  if not EstBasePclAllegee then
    Result := Presence('PIECE', 'GP_JALCOMPTABLE', St)
  else Result := False;
end;

function EstDansParamPiece(St: string): Boolean; // Modif Fiche 10856 SBO
begin
  if not EstBasePclAllegee then
    Result := Presence('PARPIECE', 'GPP_JOURNALCPTA', St)
  else Result := False;
end;

function EstDansParamPieceCompl(St: string): Boolean; // Modif Fiche 10856 SBO
begin
  if not EstBasePclAllegee then
    Result := Presence('PARPIECECOMPL', 'GPC_JOURNALCPTA', St)
  else Result := False;
end;

procedure FicheJournal(Q: TQuery; Axe, Compte: string; Comment: TActionFiche; QuellePage: Integer) ;
var lStAction : String ;
begin
  // Lancement de la fiche selon le contexte de creation modification ou consultation
  lStAction := ActionToString( Comment ) ;
  case Comment of
    // creation
    taCreat, taCreatEnSerie, taCreatOne:
      begin
        if (not ExJaiLeDroitConcept(TConcept(ccJalCreat), true)) then
          exit;
        if lStAction='' then
          lStAction := 'ACTION=CREATION' ;
        AGLLanceFiche('CP', 'CPJOURNAL', '', compte, lStAction + ';' + Compte );
      end;
    // modification
    taModif, taModifEnSerie:
      begin
        if (not ExJaiLeDroitConcept(TConcept(ccJalModif), true)) then
          exit;
        if lStAction='' then
          lStAction := 'ACTION=MODIFICATION' ;
        AGLLanceFiche('CP', 'CPJOURNAL', '', compte, lStAction + ';' + Compte );
      end;
    // Consultation
    taConsult: AGLLanceFiche('CP', 'CPJOURNAL', '', compte,'ACTION=CONSULTATION;' + Compte );
  end;
  if (Comment in [taModif, taModifEnSerie]) then
    _Bloqueur('nrBatch', false);
end;

//==================================================
// Evenements par default de la TOM
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... : 23/02/2005
Description .. : - LG - 23/02/2005 - l'onglet des info comp de saisie est
Suite ........ : visible pour tout les type de journaux
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnNewRecord;
begin
  inherited;

  (*//EN ATTENTE
      if ((not V_PGI.MonoFiche) and (Lequel <> '') and (not (Mode in [taCreat..taCreatOne]))) then
      begin
          if (not DS.Locate('J_JOURNAL',Lequel,[])) then
          begin
              MessageAlerte(MsgBox.Mess[18]);
              PostMessage(Handle,WM_CLOSE,0,0);
              exit;
          end;
      end;

      Case Mode of
          taConsult :
              begin
                  Pages.Pages[PTreso].TabVisible := ((GetField('J_NATUREJAL') = 'BQE') and (ctxPCL in V_PGI.PGIContexte));
                  Pages.Pages[PSaisie].TabVisible := ((GetField('J_MODESAISIE') = 'BOR') or (GetField('J_MODESAISIE') = 'LIB'));
              end;
          taCreat..taCreatOne : SetField('J_JOURNAL',Lequel);
      end;
  *)

  AvecMvt := false;

  SetField('J_NATUREJAL', 'OD');
  SetField('J_MODESAISIE', '-');
  SetField('J_CHOIXDATE', 'DATEOPERATION');
  SetField('J_INCREF', '-');
  SetField('J_EQAUTO', '-');
 // SetField('J_NATCOMPL', '-');
  SetField('J_INCNUM', '-');
  SetField('J_DATECREATION', V_PGI.DateEntree);

  Pages.Pages[PReg].TabVisible := False;
  Pages.Pages[PTreso].TabVisible := False;

  ActivationControl;
  //J_NATUREJALChange( nil );

  // Affichage des date
  SetControlText('JDATECREATION', FormatDateTime('dd mmm yyyy', Now));
  SetControlText('JDATEMODIF', FormatDateTime('dd mmm yyyy', Now));
  SetControlText('JDATEOUVERTURE', FormatDateTime('dd mmm yyyy', iDate1900));
  SetControlText('JDATEFERMETURE', FormatDateTime('dd mmm yyyy', iDate1900));

  //    Pages.ActivePageIndex := PGeneral;
  //    SetFocusControl('J_JOURNAL');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnDeleteRecord;
begin
  inherited;
  LastError := 0;

  if (MsgBox.Execute(1, '', '') <> mrYes) then
    LastError := 1
  else
    FAvertirTable := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnUpdateRecord;
begin
  inherited;
  LastError := 0;

  if (not enregOK) then
    LastError := 1
  else
   DevalideLexoSurCreatJal;
   
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnAfterUpdateRecord;
begin
  inherited;
  fAvertirTable := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnLoadRecord;
begin
  inherited;

  (*//EN ATTENTE
      if ((not V_PGI.MonoFiche) and (Lequel <> '') and (not (Mode in [taCreat..taCreatOne]))) then
      begin
          if (not DS.Locate('J_JOURNAL',Lequel,[])) then
          begin
              MessageAlerte(MsgBox.Mess[18]);
              PostMessage(Handle,WM_CLOSE,0,0);
              exit;
          end;
      end;

      Case Mode of
          taConsult :
              begin
                  Pages.Pages[PTreso].TabVisible := ((GetField('J_NATUREJAL') = 'BQE') and (ctxPCL in V_PGI.PGIContexte));
                  Pages.Pages[PSaisie].TabVisible := ((GetField('J_MODESAISIE') = 'BOR') or (GetField('J_MODESAISIE') = 'LIB'));
              end;
          taCreat..taCreatOne : SetField('J_JOURNAL',Lequel);
      end;
  *)

  if (LeCompte = GetField('J_JOURNAL')) then
    exit;
  LeCompte := GetField('J_JOURNAL');

  ChargeCumulsMultiSoc ;
  AfficheCumuls ;

  // Affichage des date
  if (not VarIsNull(GetField('J_DATECREATION'))) then
    SetControlText('JDATECREATION', FormatDateTime('dd mmm yyyy',
      GetField('J_DATECREATION')))
  else
    SetControlText('JDATECREATION', FormatDateTime('dd mmm yyyy', Now));
  if (not VarIsNull(GetField('J_DATEMODIF'))) then
    SetControlText('JDATEMODIF', FormatDateTime('dd mmm yyyy',
      GetField('J_DATEMODIF')))
  else
    SetControlText('JDATEMODIF', FormatDateTime('dd mmm yyyy', Now));
  if (not VarIsNull(GetField('J_DATEOUVERTURE'))) then
    SetControlText('JDATEOUVERTURE', FormatDateTime('dd mmm yyyy',
      GetField('J_DATEOUVERTURE')))
  else
    SetControlText('JDATEOUVERTURE', FormatDateTime('dd mmm yyyy', idate1900));
  if (not VarIsNull(GetField('J_DATEFERMETURE'))) then
    SetControlText('JDATEFERMETURE', FormatDateTime('dd mmm yyyy',
      GetField('J_DATEFERMETURE')))
  else
    SetControlText('JDATEFERMETURE', FormatDateTime('dd mmm yyyy', iDate1900));

  if (Mode = taConsult) then
  begin
    if (GetField('J_NATUREJAL') = 'ODA') or (GetField('J_NATUREJAL') = 'ANA')
      then
      SetControlProperty('J_COMPTEURNORMAL', 'DataType', 'ttSoucheComptaODA')
    else
      SetControlProperty('J_COMPTEURNORMAL', 'DataType', 'ttSoucheCompta');
    exit;
  end;

  IsMouvemente;
  AutoriseOuNonLesChamps;

  // ici l'AGL set les zone de l'ecran avec les variable du DataSet !
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnChangeField(F: TField);
begin
  inherited;

  SetControlEnabled('BMenuZoom', (not (DS.State in [dsInsert])));
  SetControlEnabled('BImprimer', (not (DS.State in [dsInsert])));

  // set de l'abregé
  if (GetControlText('J_ABREGE') = '') then
    SetControlText('J_ABREGE', copy(GetControlText('J_LIBELLE'), 1, 17));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnArgument(S: string);
var
  S1, S2: string;
  i: integer;
begin
  inherited;

  fAvertirTable := False;

  // Récupère les arguments
  S1 := UpperCase(S);

  // mode pour l'eAGL
  S2 := ReadTokenSt(S1);
  if TFFiche(Ecran).TypeAction = taCreatEnSerie then
    TFFiche(Ecran).MonoFiche := False ;

  // Compte
  S2 := ReadTokenSt(S1);
  Lequel := S2;

  // Mode
  Mode := TFFiche(Ecran).TypeAction ;

  // LesModif
  S2 := ReadTokenSt(S1);
  LesModif := S2;


  // Création des contrôles
  MsgBox := THMsgBox.create(FMenuG);
  InitMsgBox;

  //RRO on supprime le bouton DELETE
  TToolBarButton97(GetControl('BDELETE', true)).Visible := False;

  // Evénements des contrôles
  TToolBarButton97(GetControl('BINSERT', true)).OnClick := OnClickBInsert;
  TToolBarButton97(GetControl('BIMPRIMER', true)).OnClick := BImprimerClick;
  TCheckBox(GetControl('J_EFFET', true, )).OnClick := J_EFFETClick;
  // menuZoom
  TToolBarButton97(GetControl('BMenuZoom', true)).OnMouseEnter :=
    BMenuZoomEnter;
  TbitBtn(GetControl('BZECRIMVT', true)).OnClick := BZecrimvtClick;
  TbitBtn(GetControl('BCumul', true)).OnClick := BCumulClick;
  TbitBtn(GetControl('BJALDIV', true)).OnClick := BJALDIVClick;
  TbitBtn(GetControl('BJALCPT', true)).OnClick := BJALCPTClick;
  TbitBtn(GetControl('BJALPER', true)).OnClick := BJALPERClick;
  TbitBtn(GetControl('BJALCENT', true)).OnClick := BJALCENTClick;

  // on change
  THValCombobox(GetControl('J_NATUREJAL', true)).OnChange := J_NATUREJALChange;
  THValCombobox(GetControl('J_MODESAISIE', true)).OnChange :=
    J_MODESAISIEChange;
  THEdit(GetControl('J_JOURNAL', true)).OnChange := J_JOURNALOnChange;

  // on exit
  THEdit(GetControl('J_JOURNAL', true)).OnExit := J_JOURNALExit;

  THEdit(GetControl('J_CPTEREGULDEBIT1', true)).OnExit := CptRegulExit;
  THEdit(GetControl('J_CPTEREGULDEBIT2', true)).OnExit := CptRegulExit;
  THEdit(GetControl('J_CPTEREGULDEBIT3', true)).OnExit := CptRegulExit;
  THEdit(GetControl('J_CPTEREGULCREDIT1', true)).OnExit := CptRegulExit;
  THEdit(GetControl('J_CPTEREGULCREDIT2', true)).OnExit := CptRegulExit;
  THEdit(GetControl('J_CPTEREGULCREDIT3', true)).OnExit := CptRegulExit;

  // TPageControl
  Pages := TPageControl(GetControl('PAGES', true));
  Pages.OnChanging := OnChangingPages;

  // PopupMenu
  POPZ := TPopUpMenu(GetControl('POPZ', true));

{$IFDEF EAGLCLIENT}
  THRichEditOle(GetControl('J_BLOCNOTE', True)).OnMouseDown := OnMouseDownJ_BlocNote;
{$ELSE}
  THDBRichEditOle(GetControl('J_BLOCNOTE', True)).OnMouseDown := OnMouseDownJ_BlocNote;
{$ENDIF}

  // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
  FAGLCloseQuery := TFFiche(Ecran).OnCloseQuery ;
  TFFiche(Ecran).OnCloseQuery := FormCloseQuery ;

  // enable de certaion control
  if (SaisieFolioLancee) then
  begin
    SetControlEnabled('J_COMPTEINTERDIT', false);
    SetControlEnabled('J_COMPTEAUTOMAT', false);
  end;

  //SG6 22/12/2004 Gestion du mode croisaxe

  if VH^.AnaCroisaxe then
  begin
    for i := 1 to 5 do
    begin
      SetControlVisible('TJ_AXE' + IntToStr(i), True);
      SetControlVisible('J_AXE' + IntToStr(i), True);
    end;
    SetControlVisible('J_AXE', False);
    SetControlVisible('TJ_AXE', False);         {FP FQ15836}

    premier_axe := RecherchePremDerAxeVentil.premier_axe;

  end
  else
  begin
    for i := 1 to 5 do
    begin
      SetControlVisible('TJ_AXE' + IntToStr(i), False);
      SetControlVisible('J_AXE' + IntToStr(i), False);
    end;
    SetControlVisible('J_AXE', True);
  end;

  (*
      if not ((ctxPCL in V_PGI.PGIContexte) and OkSynchro) then
      begin
          SetControlPropertry('J_MODESAISIE','Plus,'and CO_CODE<>"LIB"');
          SetControlPropertry('J_MODESAISIE','Exhaustif',exPlus);
      end;
  {$IFDEF CCS3}
      SetControlVisible('J_AXE',false);
      SetControlVisible('TJ_AXE',false);
      SetControlVisible('J_TYPECONTREPARTIE',false);
      SetControlVisible('TJ_TYPECONTREPARTIE',false);
  {$ENDIF}
  *)
      // formatage des champs numerique
  FormateLesMontants;

  // Formatage des champs date
  FormateLesDates;
{$IFDEF COMPTA}
  (*
  CA - 12/05/2004 - Pour éviter de tirer toute la comptabilité, le menu
  d'accès au cumuls des comptes mensuels est rendu invisible lorsque la
  fiche section est appelé depuis un module non complié avec la directive COMPTA
  *)
{$ELSE}
  SetControlProperty('BCUMUL', 'Tag', 0);
  // Le tag est différent de -(POPUP) donc le menu ne s'affiche pas.
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnClose;
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnCancelRecord;
begin
  inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{*****************************************************************}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnClickBInsert(Sender: TObject);
begin
  if (ExJaiLeDroitConcept(TConcept(ccJalCreat), true)) then
    TFFiche(Ecran).BInsertClick(Sender);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_Journal.BImprimerClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}

  // Gestion du multisoc : soldes des comptes repris de la table cumuls via la vue TIERSMS
  if EstTablePartagee('JOURNAL') then
    V_PGI.EnableTableToView := True ;

  { FQ 19534 BVE 12.04.07
  LanceEtat( 'E', 'CST', 'JOU', true, true, true, nil,
             'J_JOURNAL = "' + GetControlText('J_JOURNAL') + '"',
             Ecran.Caption, False, 0, '' ) ;
    END FQ 19534 }     
  LanceEtat( 'E', 'CST', 'JOU', true, false, true, nil,
             'J_JOURNAL = "' + GetControlText('J_JOURNAL') + '"',
             Ecran.Caption, False, 0, '' ) ;

  (* JP 28/10/05 : Branchement de l'état agl en 2/3
  {$IFDEF EAGLCLIENT}
    // A FAIRE
  {$ELSE}
    PlanJournal(GetField('J_JOURNAL'), true);
  {$ENDIF}
   *)

  // Gestion du multisoc : désactivation de la vue TIERSMS
  V_PGI.EnableTableToView := False ;

  {$ENDIF}
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_Journal.J_EFFETClick(Sender: TObject);
var
  enable: boolean;
  Checke: boolean;
  NatJal: string;
begin
  fAvertirTable := True;

  // recup des valeur
  NatJal := GetControlText('J_NATUREJAL');
  Checke := (GetCheckBoxState('J_EFFET') = cbChecked);
  enable := ((NatJal = 'BQE') or (NatJal = 'CAI') or (Checke));

  // set des control
  SetControlEnabled('J_CONTREPARTIE', enable);
  SetControlEnabled('TJ_CONTREPARTIE', enable);
  SetControlProperty('J_CONTREPARTIE', 'Color', IIF(enable, ClWindow,
    ClBtnFace));

  SetControlEnabled('J_TYPECONTREPARTIE', enable);
  SetControlEnabled('TJ_TYPECONTREPARTIE', enable);
  SetControlProperty('J_TYPECONTREPARTIE', 'Color', IIF(enable, ClWindow,
    ClBtnFace));

  // set de la zoomtable
  if (checke) then
    SetControlProperty('J_CONTREPARTIE', 'DataType', 'tzGEffet')
  else if (NatJal = 'BQE') then
    SetControlProperty('J_CONTREPARTIE', 'DataType', 'tzGbanque')
  else if (NatJal = 'CAI') then
    SetControlProperty('J_CONTREPARTIE', 'DataType', 'tzGcaisse');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.J_MODESAISIEChange(Sender: TObject);
begin
  InitFolio;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.J_JOURNALExit(Sender: TObject);
begin
  if (DS.State = dsInsert) then
    SetControlText('J_JOURNAL', Trim(GetControlText('J_JOURNAL')));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.CptRegulExit(Sender: TObject);
begin
  if (Mode <> taConsult) then
    LookUpValueExist(TControl(Sender));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.J_JOURNALOnChange(Sender: TObject);
begin
  ActivationControl;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.J_NATUREJALChange(Sender: TObject);
begin
  if (CtxPcl in V_Pgi.PgiContexte) and (Trim(GetControlText('J_JOURNAL')) = '')
    then
    exit;
  AutoriseOuNonLesChamps;
end;

{ Procédures de zoom sur les états à partir de la fiche compte }

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.BMenuZoomEnter(Sender: TObject);
begin
  PopZoom97(TToolBarButton97(GetControl('BMenuZoom', true)), POPZ);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.BZecrimvtClick(Sender: TObject);     
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCADM}
var
  ACritEdt : ClassCritEdt;
{$ENDIF CCADM}
{$ENDIF}
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCADM}
  if (GetControlText('J_NATUREJAL') = 'ODA') or (GetControlText('J_NATUREJAL') = 'ANA') then
  begin
{ FQ 20370 BVE 02.10.07 }
    ACritEdt := ClassCritEdt.Create;
    try
      ACritEdt.CritEdt.LibreCodes1 := GetControlText('J_JOURNAL') ;// Journal
      ACritEdt.CritEdt.sCpt1 := ''; // Compte
      ACritEdt.CritEdt.Cpt1 := ''; // Section
      ACritEdt.CritEdt.Bal.Axe := GetControlText('J_AXE'); // Axe
      ACritEdt.CritEdt.Exo.Code := EXRF(VH^.Entree.Code); //Exercice
      if VH^.Precedent.Code <> '' then
         ACritEdt.CritEdt.Date1 := VH^.Precedent.Deb // Date Comptable debut
      else
         ACritEdt.CritEdt.Date1 := VH^.Encours.Deb ;// Date Comptable debut
      ACritEdt.CritEdt.Date2 := V_PGI.DateEntree ;// Date Comptable fin
      TheData := ACritEdt;
      MultiCritereAnaZoom(taConsult,ACritEdt.CritEdt);
    finally
      ACritEdt.Free;
      TheData := nil;
    end;
  end
{ END FQ 20370 }
  else
    ZoomEcritureMvt(GetControlText('J_JOURNAL'), fbJal, 'MULMMVTS');
{$ENDIF CCADM}    
{$ENDIF}
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.BCumulClick(Sender: TObject);
begin
{$IFDEF COMPTA}
  CumulCpteMensuel(fbJal, GetControlText('J_JOURNAL'),
    GetControlText('J_LIBELLE'), VH^.Entree);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.BJALDIVClick(Sender: TObject);
(*
var
    Crit : TCritJour;
    D1,D2 : TdateTime;
*)
begin
  (*
      Fillchar(Crit,SizeOf(Crit),#0);
      D1 := V_PGI.Encours.Deb;
      D2 := V_PGI.Encours.Fin;
      if (V_PGI.Entree.Code = V_PGI.Suivant.Code) then
      begin

          D1 := V_PGI.Suivant.Deb;
          D2 := V_PGI.Suivant.Fin;
      end;
      Crit.Date1 := D1;
      Crit.Date2 := D2;
      Crit.DateDeb := Crit.Date1;
      Crit.DateFin := Crit.Date2;

      InitCritJour(Crit);

      Crit.Code1 := GetControlText('J_JOURNAL');
      Crit.Code2 := Crit.Code1;

      JalDivisioZoom(Crit);
  *)
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.BJALCPTClick(Sender: TObject);
(*
var
    Crit : TCritJour;
    D1,D2 : TdateTime;
*)
begin
  (*
      Fillchar(Crit,SizeOf(Crit),#0);
      D1 := V_PGI.Encours.Deb;
      D2 := V_PGI.Encours.Fin;
      if (V_PGI.Entree.Code = V_PGI.Suivant.Code) then
      begin
          D1 := V_PGI.Suivant.Deb;
          D2 := V_PGI.Suivant.Fin;
      end;
      Crit.Date1 := D1;
      Crit.Date2 := D2;
      Crit.DateDeb := Crit.Date1;
      Crit.DateFin := Crit.Date2;

      InitCritJour(Crit);

      Crit.Code1 := GetControlText('J_JOURNAL');
      Crit.Code2 := Crit.Code1;

      JalCpteGeZoom(Crit);
  *)
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.BJALPERClick(Sender: TObject);
(*
var
    Crit : TCritJourGene;
    D1,D2 : TdateTime;
*)
begin
  (*
      Fillchar(Crit,SizeOf(Crit),#0);
      D1 := V_PGI.Encours.Deb;
      D2 := V_PGI.Encours.Fin;
      if (V_PGI.Entree.Code = V_PGI.Suivant.Code) then
      begin
          D1 := V_PGI.Suivant.Deb;
          D2 := V_PGI.Suivant.Fin;
      end;
      Crit.Date1 := D1;
      Crit.Date2 := D2;
      Crit.DateDeb := Crit.Date1;
      Crit.DateFin := Crit.Date2;

      InitCritJourGene(Crit);

      Crit.Code1 := GetControlText('J_JOURNAL');
      Crit.Code2 := Crit.Code1;

      JalPeriodeZoom(Crit);
  *)
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.BJALCENTClick(Sender: TObject);
(*
var
    Crit : TCritJourGene;
    D1,D2 : TdateTime;
*)
begin
  (*
      Fillchar(Crit,SizeOf(Crit),#0);
      D1 := V_PGI.Encours.Deb;
      D2 := V_PGI.Encours.Fin;
      if (V_PGI.Entree.Code = V_PGI.Suivant.Code) then
      begin
          D1 := V_PGI.Suivant.Deb;
          D2 := V_PGI.Suivant.Fin;
      end;
      Crit.Date1 := D1;
      Crit.Date2 := D2;
      Crit.DateDeb := Crit.Date1;
      Crit.DateFin := Crit.Date2;

      InitCritJourGene(Crit);

      Crit.Code1 := GetControlText('J_Journal');
      Crit.Code2 := Crit.Code1;

      JalCentralZoom(Crit);
  *)
end;

{ Fin }

//==================================================
// Autres Evenements Bis
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.ChaineCpteValide(Sender: TObject): Boolean;
var
  St, StCode, Stn: string;
  LeWhere: string;
begin
  result := true;

  if (TCustomEdit(Sender).Text = '') then
    exit;

  St := UpperCase(TCustomEdit(Sender).Text);
  if (St[Length(St)] <> ';') then
    St := St + ';';
  Stn := TCustomEdit(Sender).Name;

  while (St <> '') do
  begin
    StCode := ReadTokenSt(St);
    LeWhere := FaitLeWhere(StCode);

    if (not ExisteLeCompte(LeWhere)) then
    begin
      if (Stn = 'J_COMPTEINTERDIT') or (Stn = 'J_COMPTEAUTOMAT') or (Stn =
        'J_CONTREPARTIE') then
        Pages.ActivePageIndex := PComplement
      else
        Pages.ActivePageIndex := PReg;

      TCustomEdit(Sender).SetFocus;
      MsgBox.Execute(6, '', '');
      result := false;
      exit;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.CaractereValide(Sender: TObject): Boolean;
var
  i: Byte;
  St, Stn: string;
begin
  result := true;

  if (TCustomEdit(Sender).Text = '') then
    exit;

  St := UpperCase(TCustomEdit(Sender).Text);
  Stn := TCustomEdit(Sender).Name;

  for i := 1 to Length(St) do
  begin
    if not (St[i] in ['0'..'9', 'A'..'Z', ';', ':', VH^.Cpta[fbGene].Cb]) then
    begin
      if (Stn = 'J_COMPTEINTERDIT') or (Stn = 'J_COMPTEAUTOMAT') or (Stn =
        'J_CONTREPARTIE') then
        Pages.ActivePageIndex := PComplement
      else
        Pages.ActivePageIndex := PReg;

      TCustomEdit(Sender).SetFocus;
      MsgBox.Execute(5, '', '');
      result := false;
      exit;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.DansUnCpte(Cpte1, Cpte2: TObject): Boolean;
var
  Q1, Q2: TQuery;
  MemoCpte2, St, St1: string;
  S1, S2, LeWhere1, LeWhere2: string;
begin
  result := false;

  if (TCustomEdit(Cpte1).Text = '') then
    exit;
  if (TCustomEdit(Cpte2).Text = '') then
    exit;

  S1 := TCustomEdit(Cpte1).Text;
  if (S1[Length(S1)] <> ';') then
    S1 := S1 + ';';
  S2 := TCustomEdit(Cpte2).Text;
  if (S2[Length(S2)] <> ';') then
    S2 := S2 + ';';

  MemoCpte2 := S2;

  while (S1 <> '') do
  begin
    St := ReadTokenSt(S1);
    LeWhere1 := FaitLeWhere(St);
    Q1 := OpenSQL('SELECT G_GENERAL FROM GENERAUX ' + LeWhere1, true);
    while (not Q1.EOF) do
    begin
      S2 := MemoCpte2;
      while (S2 <> '') do
      begin
        St1 := ReadTokenSt(S2);
        LeWhere2 := FaitLeWhere(St1);
        Q2 := OpenSQL('SELECT G_GENERAL FROM GENERAUX ' + LeWhere2, true);
        while (not Q2.Eof) do
        begin
          if (Q1.Fields[0].AsString = Q2.Fields[0].AsString) then
          begin
            result := true;
            Q1.Close;
            Q2.Close;
            exit;
          end;
          Q2.Next;
        end;
      end;
      Q1.Next;
      Ferme(Q2);
    end;
    Ferme(Q1);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.CpteDeRegulValide(Sender: TObject): Boolean;
begin
  result := true;

  if (THEdit(Sender).Text = '') then
    exit;
  if (not LookUpValueExist(THEdit(Sender))) then
    result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_Journal.MsgCpteRegul(Sender: TOBject; i: Byte);
begin
  Pages.ActivePageIndex := PReg;
  THEdit(Sender).SetFocus;
  MsgBox.Execute(i, '', '');
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.DevalideLexoSurCreatJal;
var
  i: Byte;
  St: string;
begin
  St := '';
  for i := 1 to 24 do
    St := St + '-';
  ExecuteSql('UPDATE EXERCICE SET EX_VALIDEE="' + St +
    '" WHERE EX_ETATCPTA="OUV"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_Journal.ClearChampsSurEnregOk;
var
  NatJal: string;
  i: Byte;
begin
  if DS.State = dsBrowse then
    exit;

  Natjal := GetControlText('J_NATUREJAL');
  if (NatJal <> 'REG') then
  begin
    for i := 1 to 3 do
    begin
      DS.FindField('J_CPTEREGULDEBIT' + IntToStr(i)).AsString := '';
      DS.FindField('J_CPTEREGULCREDIT' + IntToStr(i)).AsString := '';
    end;
  end;

  if (NatJal <> 'ODA') and (NatJal <> 'ANA') then
    SetField('J_AXE', '');
  if (NatJal = 'ODA') or (NatJal = 'ANA') then
  begin
    SetField('J_MULTIDEVISE', '-');
    SetField('J_COMPTEINTERDIT', '');
    SetField('J_COMPTEAUTOMAT', '');
    SetField('J_COMPTEURSIMUL', '');
    SetField('J_CONTREPARTIE', '');
    SetField('J_TYPECONTREPARTIE', '');
    exit;
  end;

  if (NatJal = 'ACH') or (NatJal = 'ECC') or (NatJal = 'EXT') or ((NatJal = 'OD')
    and (not (GetCheckBoxState('J_EFFET') = cbChecked)) or (NatJal = 'REG') or
    (NatJal = 'VTE')) then
  begin
    SetField('J_CONTREPARTIE', '');
    SetField('J_TYPECONTREPARTIE', '');
    exit;
  end;

  if (NatJal = 'ANO') then
  begin
    SetField('J_COMPTEINTERDIT', '');
    SetField('J_COMPTEAUTOMAT', '');
    SetField('J_COMPTEURSIMUL', '');
    SetField('J_CONTREPARTIE', '');
    SetField('J_TYPECONTREPARTIE', '');
    SetField('J_CENTRALISABLE', '-');
    exit;
  end;

  if ((NatJal = 'BQE') or (NatJal = 'CAI')) then
  begin
    SetField('J_COMPTEAUTOMAT', '');
    exit;
  end;

  if (NatJal = 'CLO') then
  begin
    SetField('J_CENTRALISABLE', '-');
    SetField('J_COMPTEINTERDIT', '');
    SetField('J_COMPTEAUTOMAT', '');
    SetField('J_COMPTEURSIMUL', '');
    SetField('J_CONTREPARTIE', '');
    SetField('J_TYPECONTREPARTIE', '');
    SetField('J_MULTIDEVISE', '-');
    exit;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.ExisteLeCompte(LeWhere: string): Boolean;
begin
  result := ExisteSql('SELECT G_GENERAL FROM GENERAUX ' + LeWhere);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.FaitLeWhere(St: string): string;
var
  S1, S2: string;
begin
  if (Pos(':', St) > 0) then
  begin
    S1 := Copy(St, 1, Pos(':', St) - 1);
    S2 := Copy(St, Pos(':', St) + 1, Length(St));
    S2 := BourreLaDonc(S2, fbGene);
    result := 'WHERE G_GENERAL>="' + S1 + '%" AND G_GENERAL<="' + S2 + '"';
  end
  else
  begin
    if (Length(St) < VH^.Cpta[fbGene].Lg) then
      result := 'WHERE G_GENERAL LIKE "' + St + '%"'
    else
      result := 'WHERE G_GENERAL="' + St + '"';
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.EstJalAnal: Boolean;
begin
  result := true;

  if (GetControlText('J_NATUREJAL') = 'ODA') or (GetControlText('J_NATUREJAL') =
    'ANA') then
  begin
    if (GetControlText('J_AXE') = '') then
    begin
      Pages.ActivePageIndex := PGeneral;
      SetFocusControl('J_AXE');
      MsgBox.Execute(19, '', '');
      result := false;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.CrtlCpteDeRegul: Boolean;
var
  i: Byte;
begin
  result := false;

  for i := 1 to 3 do
  begin
    if (not CpteDeRegulValide(THEdit(GetControl('J_CPTEREGULDEBIT' +
      IntToStr(i), true)))) then
    begin
      MsgCpteRegul(THEdit(GetControl('J_CPTEREGULDEBIT' + IntToStr(i), true)),
        21);
      exit;
    end;

    if (DansUnCpte(THEdit(GetControl('J_COMPTEAUTOMAT', true)),
      THEdit(GetControl('J_CPTEREGULDEBIT' + IntToStr(i), true)))) then
    begin
      MsgCpteRegul(THEdit(GetControl('J_CPTEREGULDEBIT' + IntToStr(i), true)),
        22);
      exit;
    end;

    if (DansUnCpte(THEdit(GetControl('J_COMPTEINTERDIT', true)),
      THEdit(GetControl('J_CPTEREGULDEBIT' + IntToStr(i), true)))) then
    begin
      MsgCpteRegul(THEdit(GetControl('J_CPTEREGULDEBIT' + IntToStr(i), true)),
        23);
      exit;
    end;

    if (not CpteDeRegulValide(THEdit(GetControl('J_CPTEREGULCREDIT' +
      IntToStr(i), true)))) then
    begin
      MsgCpteRegul(THEdit(GetControl('J_CPTEREGULCREDIT' + IntToStr(i), true)),
        21);
      exit;
    end;

    if (DansUnCpte(THEdit(GetControl('J_COMPTEAUTOMAT', true)),
      THEdit(GetControl('J_CPTEREGULCREDIT' + IntToStr(i), true)))) then
    begin
      MsgCpteRegul(THEdit(GetControl('J_CPTEREGULCREDIT' + IntToStr(i), true)),
        22);
      exit;
    end;

    if (DansUnCpte(THEdit(GetControl('J_COMPTEINTERDIT', true)),
      THEdit(GetControl('J_CPTEREGULCREDIT' + IntToStr(i), true)))) then
    begin
      MsgCpteRegul(THEdit(GetControl('J_CPTEREGULCREDIT' + IntToStr(i), true)),
        23);
      exit;
    end;
  end;

  result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.EnregOK: boolean;
var
  NatJal: string;
  QSouche: TQuery;

begin
  result := false;

  if (not (DS.State in [dsInsert, dsEdit])) then
    exit;

  NatJal := GetControlText('J_NATUREJAL');
  (*
      if (not (ctxPCL in V_PGI.PGIContexte)) then
      begin
          If ((QJal.State in [dsInsert]) or (GetControlEnabled('J_MODESAISIE') and (QJal.State in [dsEdit]))) and (GetControlText('J_MODESAISIE') = 'LIB') Then
          begin
              Pages.ActivePageIndex := PGeneral;
              SetFocusControl('J_MODESAISIE');
              MsgBox.execute(27,'','');
              exit;
          end;
      end;
  *)
  if DS.State in [dsInsert, dsEdit] then
  begin
    if (GetControlText('J_Journal') = '') then
    begin
      Pages.ActivePageIndex := PGeneral;
      SetFocusControl('J_JOURNAL');
      MsgBox.execute(2, '', '');
      exit;
    end;

    if (GetControlText('J_LIBELLE') = '') then
    begin
      Pages.ActivePageIndex := PGEneral;
      SetFocusControl('J_LIBELLE');
      MsgBox.execute(3, '', '');
      exit;
    end;

    if (not CaractereValide(GetControl('J_COMPTEINTERDIT', true))) then
      exit;
    if (not CaractereValide(GetControl('J_COMPTEAUTOMAT', true))) then
      exit;
    if (not CaractereValide(GetControl('J_CONTREPARTIE', true))) then
      exit;
    if (not ChaineCpteValide(GetControl('J_COMPTEINTERDIT', true))) then
      exit;
    if (not ChaineCpteValide(GetControl('J_COMPTEAUTOMAT', true))) then
      exit;
    if (not ChaineCpteValide(GetControl('J_CONTREPARTIE', true))) then
      exit;

    if (NatJal = 'REG') then
    begin
      if (not CaractereValide(GetControl('J_CPTEREGULDEBIT1', true))) then
        exit;
      if (not CaractereValide(GetControl('J_CPTEREGULCREDIT1', true))) then
        exit;
      if (not CaractereValide(GetControl('J_CPTEREGULDEBIT2', true))) then
        exit;
      if (not CaractereValide(GetControl('J_CPTEREGULCREDIT2', true))) then
        exit;
      if (not CaractereValide(GetControl('J_CPTEREGULDEBIT3', true))) then
        exit;
      if (not CaractereValide(GetControl('J_CPTEREGULCREDIT3', true))) then
        exit;
      if (not CrtlCpteDeRegul) then
        exit;
    end;

    if (((NatJal = 'BQE') or (NatJal = 'CAI')) and ((not ContrePartieValide) or
      (not ContrePartieDevise))) then
      exit;

    if ((NatJal = 'ACH') or (NatJal = 'ECC') or (NatJal = 'EXT') or (NatJal =
      'REG') or (NatJal = 'VTE') or (NatJal = 'OD')) then
    begin
      if (DansUnCpte(GetControl('J_COMPTEAUTOMAT', true),
        GetControl('J_COMPTEINTERDIT', true))) then
      begin
        Pages.ActivePageIndex := PComplement;
        SetFocusControl('J_COMPTEAUTOMAT');
        MsgBox.Execute(8, '', '');
        exit;
      end;
    end;

(*    if (NatJal = 'ECC') then
      SetControlChecked('J_MULTIDEVISE', true);
*)

    if ((NatJal = 'ODA') or (NatJal = 'ANA')) then
      if (not EstJalAnal) then
        exit;
    //Gestion du compteur Normal
    if (GetControlText('J_COMPTEURNORMAL') = '') then
    begin
      if (ctxPCL in V_PGI.PGIContexte) then
      begin
        if ((NatJal = 'ODA') or (NatJal = 'ANA')) then
        begin
          if
            ExisteSQL('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" and SH_SOUCHE="ANA"') then
            // SetControlText('J_COMPTEURNORMAL', 'ANA')
            { FQ 16111 - CA - 22/06/2005 : la mise à jour d'un champ c'est SetField et non SetControlText ... }
            SetField('J_COMPTEURNORMAL', 'ANA')
          else
          begin
            Pages.ActivePageIndex := PComplement;
            SetFocusControl('J_COMPTEURNORMAL');
            MsgBox.Execute(26, '', '');
            exit;
          end;
        end
        else
        begin
          if ((NatJal = 'ANO') or (NatJal = 'CLO')) then
          begin
            if
              (ExisteSQL('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" and SH_SOUCHE="'
              + NatJal + '"')) then
              // SetControlText('J_COMPTEURNORMAL', NatJal)
              { FQ 16111 - CA - 22/06/2005 : la mise à jour d'un champ c'est SetField et non SetControlText ... }
              SetField('J_COMPTEURNORMAL', NatJal)
            else
            begin
              Pages.ActivePageIndex := PComplement;
              SetFocusControl('J_COMPTEURNORMAL');
              MsgBox.Execute(26, '', '');
              exit;
            end;
          end
          else
          begin
            if
              (ExisteSQL('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" and SH_SOUCHE="CPT"')) then
              // SetControlText('J_COMPTEURNORMAL', 'CPT')
              { FQ 16111 - CA - 22/06/2005 : la mise à jour d'un champ c'est SetField et non SetControlText ... }
              SetField ('J_COMPTEURNORMAL', 'CPT')
            else
            begin
              Pages.ActivePageIndex := PComplement;
              SetFocusControl('J_COMPTEURNORMAL');
              MsgBox.Execute(26, '', '');
              exit;
            end;
          end;
        end;
      end
      else // si non  PCL
      begin
        if VH^.CPIFDEFCEGID then
        begin
          QSouche := OpenSQL('SELECT * FROM SOUCHE WHERE SH_SOUCHE="' +
            GetControlText('J_JOURNAL') + '" and SH_TYPE="CPT"', false);
          if (QSouche.Eof) then
          begin
            QSouche.Insert;
            InitNew(QSouche);

            QSouche.FindField('SH_TYPE').AsString := 'CPT';
            QSouche.FindField('SH_SOUCHE').AsString :=
              GetControlText('J_JOURNAL');
            QSouche.FindField('SH_LIBELLE').AsString := 'Souche ' +
              GetControlText('J_JOURNAL');
            QSouche.FindField('SH_ABREGE').AsString :=
              QSouche.FindField('SH_LIBELLE').AsString;
            QSouche.FindField('SH_NUMDEPART').AsInteger := 1;
            QSouche.FindField('SH_NUMDEPARTS').AsInteger := 1;
            QSouche.FindField('SH_NUMDEPARTP').AsInteger := 1;
            QSouche.FindField('SH_SOUCHEEXO').AsString := '-';
            QSouche.FindField('SH_DATEDEBUT').AsDateTime := EncodeDate(1900, 01,
              01);
            QSouche.FindField('SH_DATEFIN').AsDateTime := EncodeDate(2099, 12,
              31);
            if ((NatJal = 'ODA') or (NatJal = 'ANA')) then
              QSouche.FindField('SH_ANALYTIQUE').AsString := 'X'
            else
              QSouche.FindField('SH_ANALYTIQUE').AsString := '-';
            QSouche.FindField('SH_SIMULATION').AsString := '-';

            QSouche.Post;
            Ferme(QSouche);

            AvertirTable('TTSOUCHECOMPTA');
            THValComboBox(GetControl('J_COMPTEURNORMAL', true)).Reload;
            SetControlText('J_COMPTEURNORMAL', GetControlText('J_JOURNAL'));
          end
          else
          begin
            Ferme(QSouche);
            Pages.ActivePageIndex := PComplement;
            SetFocusControl('J_COMPTEURNORMAL');
            MsgBox.Execute(26, '', '');
            exit;
          end;
        end
        else
        begin
          Pages.ActivePageIndex := PComplement;
          SetFocusControl('J_COMPTEURNORMAL');
          MsgBox.Execute(26, '', '');
          Exit;
        end;
      end;
    end;
  end;
  //Gestion du compteur de simu.
  //RR 22/03/2005. FQ 15528
  //Suppression de la particularité "Cegid".
  //Compteur de simulation désormais obligatoire pour tous.
  //Incluant l'affectation auto d'un compteur SIM , s'il existe.
  if not (ctxPCL in V_PGI.PGIContexte) then
  begin
    //Mode PIECE UNIQUEMENT !
    if (GetControlText('J_COMPTEURSIMUL') = '')
      and (GetControlText('J_MODESAISIE') = '-')
      and (NatJal <> 'ODA')
      and (NatJal <> 'ANA')
      and (NatJal <> 'ANO')
      and (NatJal <> 'CLO') then
    begin
      //On affecte par défaut SIM
      if ExisteSQL('SELECT SH_SOUCHE,SH_TYPE FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="SIM"') then
        SetControlText('J_COMPTEURSIMUL', 'SIM')
      else
        //Sinon, message d'erreur afin que l'utilisateur en sélectionne un.
      begin
        Pages.ActivePageIndex := PComplement;
        SetFocusControl('J_COMPTEURSIMUL');
        MsgBox.Execute(26, '', '');
      end;
    end; //if GetControlText
  end //if not (ctxPCL...)
  else
  begin
    { AJOUT ME LE 18/04/2001 }
    if (GetControlText('J_COMPTEURSIMUL') = '')
      and (ctxPCL in V_PGI.PGIContexte)
      and (NatJal <> 'ODA')
      and (NatJal <> 'ANA')
      and (NatJal <> 'ANO')
      and (NatJal <> 'CLO') then
      if ExisteSQL('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="SIM"') then
        SetControlText('J_COMPTEURSIMUL', 'SIM');
  end;

  if (DS.state in [dsInsert]) then
  begin
    if (CodExist) then
    begin
      Pages.ActivePageIndex := PGeneral;
      SetFocusControl('J_JOURNAL');
      MsgBox.Execute(4, '', '');
      exit;
    end;
  end;

  ClearChampsSurEnregOk;

  result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.ContrePartieDevise: Boolean;
var
  QLoc: TQuery;
  StDev: string;
begin
  result := true;

  if (GetCheckBoxState('J_EFFET') = cbchecked) then
    exit;
  if (GetCheckBoxState('J_MULTIDEVISE') = cbchecked) then
    exit;
  if (GetControlText('J_CONTREPARTIE') = '') then
    exit;

  QLoc := OpenSql('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL="' +
    GetControlText('J_CONTREPARTIE') + '"', true);
  if (QLoc.Eof) then
  begin
    Ferme(QLoc);
    exit;
  end;

  StDev := QLoc.Fields[0].AsString;
  Ferme(QLoc);

  if ((StDev = '') or (StDev = V_PGI.DevisePivot)) then
    exit;

  MsgBox.Execute(24, '', '');
  Pages.ActivePageIndex := PGeneral;
  SetFocusControl('J_MULTIDEVISE');

  result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.CodExist: Boolean;
begin
  result := Presence('JOURNAL', 'J_JOURNAL', GetControlText('J_JOURNAL'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TOM_Journal.ContrePartieValide: Boolean;
begin
  result := false;

  if (GetControlText('J_CONTREPARTIE') = '') then
  begin
    Pages.ActivePageIndex := PComplement;
    SetFocusControl('J_CONTREPARTIE');
    MsgBox.Execute(9, '', '');
    exit;
  end;

  if (GetControlText('J_TYPECONTREPARTIE') = '') then
  begin
    Pages.ActivePageIndex := PComplement;
    SetFocusControl('J_TYPECONTREPARTIE');
    MsgBox.Execute(10, '', '');
    exit;
  end;

  if (not CpteDeRegulValide(GetControl('J_CONTREPARTIE', true))) then
  begin
    Pages.ActivePageIndex := PComplement;
    SetFocusControl('J_CONTREPARTIE');
    MsgBox.Execute(25, '', '');
    exit;
  end;

  if (DansUnCpte(GetControl('J_COMPTEINTERDIT', true),
    GetControl('J_CONTREPARTIE', true))) then
  begin
    Pages.ActivePageIndex := PComplement;
    SetFocusControl('J_CONTREPARTIE');
    MsgBox.Execute(20, '', '');
    exit;
  end;

  result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... : 14/06/2005
Description .. : - LG - 23/02/2005 - on rends inactif les info propre a la
Suite ........ : saisie bordereau
Suite ........ : - LG - 14/06/2005 - ajout du param tresolibre
Mots clefs ... : 
*****************************************************************}

procedure TOM_JOURNAL.InitFolio;
var
  enable: boolean;
begin


  if (GetControlText('J_MODESAISIE') = '-') then
  begin
    enable := ((GetControlText('J_NATUREJAL') = 'BQE') or
      (GetControlText('J_NATUREJAL') = 'CAI') or (GetCheckBoxState('J_EFFET') =
      cbchecked));
    SetControlEnabled('J_TYPECONTREPARTIE', enable);
    SetControlEnabled('TJ_TYPECONTREPARTIE', enable);
    SetControlEnabled('INFOBOR', false);
  end
  else
  begin
    // GCO - 13/09/2005 - FQ 15899
    //THValComboBox(GetControl('J_TYPECONTREPARTIE', true)).ItemIndex :=
    // THValComboBox(GetControl('J_TYPECONTREPARTIE', True)).Values.IndexOf('MAN');

    //SetControlEnabled('J_TYPECONTREPARTIE', false);
    //SetControlEnabled('TJ_TYPECONTREPARTIE', false);
    enable := ((GetControlText('J_NATUREJAL') = 'BQE') or
      (GetControlText('J_NATUREJAL') = 'CAI') or (GetCheckBoxState('J_EFFET') =
      cbchecked));
    SetControlEnabled('J_TYPECONTREPARTIE', enable);
    SetControlEnabled('TJ_TYPECONTREPARTIE', enable);

    SetControlEnabled('INFOBOR', true);
  end;

  SetControlEnabled('J_TRESOLIBRE', GetControlText('J_MODESAISIE') = 'LIB' );
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... : 23/02/2005
Description .. : - LG - 23/02/2005 - l'onglet des info comp de saisie est
Suite ........ : visible pour tout les type de journaux
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.AutoriseOuNonLesChamps;
var
  enable: boolean;
  NatJal: string;
  i: integer;
begin
  NatJal := GetControlText('J_NATUREJAL');

  // AXE

  enable := ((NatJal = 'ODA') or (NatJal = 'ANA'));

  //SG6 Gestion du mode Croisaxe 22/12/2004
  if VH^.AnaCroisaxe then
  begin
    for i := 1 to 5 do
    begin
      if enable then
        SetControlChecked('J_AXE' + IntToStr(i), GetParamSocSecur('SO_VENTILA' +
          IntToStr(i),False))
      else
        SetControlChecked('J_AXE' + IntToStr(i), False);
    end;
  end
  else
  begin
    SetControlEnabled('J_AXE', enable);
    SetControlEnabled('TJ_AXE', enable);
    SetControlProperty('J_AXE', 'Color', IIF(enable, ClWindow, ClBtnFace));
  end;

  //SG6 23/12/2004 En mode croisaxe on donne a la valeur J_AXE, le premier axe ventilée du mode croise axe
  if VH^.AnaCroisaxe then
  begin
    if premier_axe <> 0 then
      SetField('J_AXE', 'A' + IntToStr(premier_axe));
  end;

  // CENTRALISABLE
  enable := (not ((NatJal = 'ANO') or (NatJal = 'CLO')));
  SetControlEnabled('J_CENTRALISABLE', enable);

  // COMPTE INTERDIT
  enable := (not ((NatJal = 'ANO') or (NatJal = 'CLO') or (NatJal = 'ODA') or
    (NatJal = 'ANA')));
  SetControlEnabled('J_COMPTEINTERDIT', enable);
  SetControlEnabled('TJ_COMPTEINTERDIT', enable);
  SetControlProperty('J_COMPTEINTERDIT', 'Color', IIF(enable, ClWindow,
    ClBtnFace));

  // COMPTE AUTOMATIQUE
  enable := (not ((NatJal = 'ANO') or (NatJal = 'BQE') or (NatJal = 'CAI') or
    (NatJal = 'ODA') or (NatJal = 'CLO') or (NatJal = 'ANA')));
  SetControlEnabled('J_COMPTEAUTOMAT', enable);
  SetControlEnabled('TJ_COMPTEAUTOMAT', enable);
  SetControlProperty('J_COMPTEAUTOMAT', 'Color', IIF(enable, ClWindow,
    ClBtnFace));

  // COMPTEUR DE SIMULATION
  enable := ((NatJal <> 'ODA') and (NatJal <> 'ANA') and (NatJal <> 'ANO'));
  SetControlEnabled('J_COMPTEURSIMUL', enable);
  SetControlEnabled('TJ_COMPTEURSIMUL', enable);
  SetControlProperty('J_COMPTEURSIMUL', 'Color', IIF(enable, ClWindow,
    ClBtnFace));

  // CONTREPARTIE
  enable := ((NatJal = 'BQE') or (NatJal = 'CAI') or (GetCheckBoxState('J_EFFET')
    = cbChecked));
  SetControlEnabled('J_CONTREPARTIE', enable);
  SetControlEnabled('TJ_CONTREPARTIE', enable);
  SetControlProperty('J_CONTREPARTIE', 'Color', IIF(enable, ClWindow,
    ClBtnFace));

  // TYPE DE CONTREPARTIE
  enable := ((NatJal = 'BQE') or (NatJal = 'CAI') or (GetCheckBoxState('J_EFFET')
    = cbChecked));
  SetControlEnabled('J_TYPECONTREPARTIE', enable);
  SetControlEnabled('TJ_TYPECONTREPARTIE', enable);
  SetControlProperty('J_TYPECONTREPARTIE', 'Color', IIF(enable, ClWindow,
    ClBtnFace));
  if (NatJal = 'BQE') then
    SetControlProperty('J_CONTREPARTIE', 'DataType', 'TZGBanque');
  if (NatJal = 'CAI') then
    SetControlProperty('J_CONTREPARTIE', 'DataType', 'tzGcaisse');
  if (GetCheckBoxState('J_EFFET') = cbChecked) then
    SetControlProperty('J_CONTREPARTIE', 'DataType', 'tzGEffet');

  // NATURE DE JOURNAL
  SetControlEnabled('J_NATUREJAL', (not AvecMvt));

  // EFFET
  SetControlEnabled('J_EFFET', (NatJal = 'OD'));

  // MODE DE SAISIE
  enable := ((not AvecMvt) and (NatJal <> 'ANA') and (NatJal <> 'ODA') and
    (NatJal <> 'ANO') and (NatJal <> 'CLO') and (NatJal <> 'ECC'));
  SetControlEnabled('J_MODESAISIE', enable);
  SetControlProperty('J_MODESAISIE', 'Color', IIF(enable, ClWindow, ClBtnFace));
  if (DS.State = dsInsert) then
    if (not enable) then
    begin
      SetControlProperty('J_MODESAISIE', 'ItemIndex', 0);
      SetControlText('J_MODESAISIE', '-');
    end;

  // COMPTEUR NORMAL
  enable := not AvecMvt;
  if (NatJal <> 'CLO') and (NatJal <> 'ANO') then
  begin
    SetControlEnabled('J_COMPTEURNORMAL', enable);
    SetControlEnabled('TJ_COMPTEURNORMAL', enable);
    SetControlProperty('J_COMPTEURNORMAL', 'Color', IIF(enable, ClWindow,
      ClBtnFace));
  end;
  if ((NatJal = 'ODA') or (NatJal = 'ANA')) then
    SetControlProperty('J_COMPTEURNORMAL', 'DataType', 'ttSoucheComptaODA')
  else
    SetControlProperty('J_COMPTEURNORMAL', 'DataType', 'ttSoucheCompta');

  // MULTIDEVISE
  enable := (not (((GetCheckBoxState('J_MULTIDEVISE') = cbChecked) and (AvecMvt))
                                                                   // fiche 15681 ajout me 24-08-2005
    or (NatJal = 'CLO') or (NatJal = 'ODA') or (NatJal = 'ANA') or (NatJal = 'ECC')));
  SetControlEnabled('J_MULTIDEVISE', enable);
// fiche 15681 ajout me 24-08-2005
    if (NatJal = 'ECC') and (GetCheckBoxState('J_MULTIDEVISE') = cbChecked) then
      SetControlChecked('J_MULTIDEVISE', FALSE);

  InitFolio;

  // gestion de l'affichage des panels
  Pages.Pages[PReg].TabVisible := (NatJal = 'REG');
  Pages.Pages[PTreso].TabVisible := (NatJal = 'BQE') ;

  // Utiliser la bonne nature de pièce en fonction de la nature du journal
  case CaseNatJal(NatJal) of
    tzJVente: SetControlProperty('J_NATDEFAUT', 'DataType', 'ttNatPieceVente');
    tzJAchat: SetControlProperty('J_NATDEFAUT', 'DataType', 'ttNatPieceAchat');
    tzJBanque: SetControlProperty('J_NATDEFAUT', 'DataType',
        'ttNatPieceBanque');
    tzJEcartChange: SetControlProperty('J_NATDEFAUT', 'DataType',
        'ttNatPieceEcartChange');
    tzJOD: SetControlProperty('J_NATDEFAUT', 'DataType', 'ttNaturePiece');
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.IsMouvemente;
var
  NatJal: string;
begin
  if ((Mode = taConsult) or (DS.State = dsInsert)) then
  begin
    AvecMvt := false;
    exit;
  end;

  Natjal := GetControlText('J_NATUREJAL');
  if ((NatJal = 'ODA') or (NatJal = 'ANA')) then
    AvecMvt := ExisteSql('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL="' +
      GetControlText('J_JOURNAL') +
      '" AND EXISTS(SELECT Y_JOURNAL FROM ANALYTIQ WHERE Y_JOURNAL="' +
      GetControlText('J_JOURNAL') + '")')
  else
    AvecMvt := ExisteSql('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL="' +
      GetControlText('J_JOURNAL') +
      '" AND EXISTS(SELECT E_JOURNAL FROM ECRITURE WHERE E_JOURNAL="' +
      GetControlText('J_JOURNAL') + '")');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2002
Modifié le ... : 18/04/2002
Description .. : Active ou Désactive les contrôles de la fiche lors de
Suite ........ : la création d'un compte auxiliaire afin de forcer
Suite ........ : la saisie de  J_JOURNAL et J_NATUREJAL
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.ActivationControl;
var
  Enable: boolean;
begin
  //if (not (CtxPcl in V_Pgi.PGIContexte)) then exit;
  Enable := Trim(GetControlText('J_Journal')) <> '';

  SetControlEnabled('J_AXE', enable);
  SetControlEnabled('J_LIBELLE', enable);
  SetControlEnabled('J_ABREGE', enable);
  SetControlEnabled('J_MODESAISIE', enable);
  SetControlEnabled('J_MULTIDEVISE', enable);
  SetControlEnabled('J_CENTRALISABLE', enable);
  SetControlEnabled('J_EFFET', enable);

  SetControlProperty('J_AXE', 'Color', IIF(enable, ClWindow, ClBtnFace));
  SetControlProperty('J_LIBELLE', 'Color', IIF(enable, ClWindow, ClBtnFace));
  SetControlProperty('J_ABREGE', 'Color', IIF(enable, ClWindow, ClBtnFace));
  SetControlProperty('J_MODESAISIE', 'Color', IIF(enable, ClWindow, ClBtnFace));

  if (enable) then
    AutoriseOuNonLesChamps;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.FormateLesMontants;
begin
  SetControlProperty('J_DEBITDERNMVT', 'DisplayFormat', StrfMask(V_PGI.OkdecV,'', true));
  SetControlProperty('J_CREDITDERNMVT', 'DisplayFormat', StrfMask(V_PGI.OkdecV,'', true));
  SetControlProperty('J_TOTALDEBIT', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));
  SetControlProperty('J_TOTALCREDIT', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));
  SetControlProperty('J_TOTDEBP', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));
  SetControlProperty('J_TOTCREP', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));
  SetControlProperty('J_TOTDEBE', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));
  SetControlProperty('J_TOTCREE', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));
  SetControlProperty('J_TOTDEBS', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));
  SetControlProperty('J_TOTCRES', 'DisplayFormat', StrfMask(V_PGI.OkdecV, '', true));

  ChangeMask(THNumEdit(GetControl('JSOLCREP', true)), V_PGI.OkdecV, '');
  ChangeMask(THNumEdit(GetControl('JSOLCREE', true)), V_PGI.OkdecV, '');
  ChangeMask(THNumEdit(GetControl('JSOLCRES', true)), V_PGI.OkdecV, '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 23/06/2003
Modifié le ... :   /  /
Description .. : Formatage des dates
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.FormateLesDates;
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
  SetControlProperty('J_DATEDERNMVT', 'DisplayFormat', 'dd mmm yyy');
  SetControlProperty('J_DATECREATION', 'DisplayFormat', 'dd mmm yyy');
  SetControlProperty('J_DATEOUVERTURE', 'DisplayFormat', 'dd mmm yyy');
  SetControlProperty('J_DATEMODIF', 'DisplayFormat', 'dd mmm yyy');
  SetControlProperty('J_DATEFERMETURE', 'DisplayFormat', 'dd mmm yyy');
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.InitMsgBox;
begin
  MsgBox.Mess.Add('0;Journaux;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;');
  MsgBox.Mess.Add('1;Journaux;Confirmez-vous la suppression de l''enregistrement?;Q;YNC;N;C;');
  MsgBox.Mess.Add('2;Journaux;Vous devez renseigner un code.;W;O;O;O;');
  MsgBox.Mess.Add('3;Journaux;Vous devez renseigner un libellé.;W;O;O;O;');
  MsgBox.Mess.Add('4;Journaux;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
  MsgBox.Mess.Add('5;Journaux;Votre saisie est incorrecte : un caractère n''est pas conforme dans la chaîne.;W;O;O;O;');
  MsgBox.Mess.Add('6;Journaux;Vous devez renseigner un compte existant.;W;O;O;O;');
  MsgBox.Mess.Add('7;Journaux;Votre saisie est incorrecte : un caractère n''est pas conforme dans la chaîne.;W;O;O;O;');
  MsgBox.Mess.Add('8;Journaux;Vous ne pouvez pas déclarer un compte automatique et interdit à la fois;W;O;O;O;');
  MsgBox.Mess.Add('9;Journaux;Vous devez saisir un compte de contrepartie.;W;O;O;O;');
  MsgBox.Mess.Add('10;Journaux;Vous devez renseigner un type de contrepartie.;W;O;O;O;');
  MsgBox.Mess.Add('11;Journaux;Suppression impossible : ce journal est mouvementé par des écritures analytiques.;W;O;O;O;');
  MsgBox.Mess.Add('12;Journaux;Suppression impossible : ce journal est mouvementé par des écritures comptables.;W;O;O;O;');
  MsgBox.Mess.Add('13;Journaux;Suppression impossible : ce journal est mouvementé par des écritures d''immobilisation.;W;O;O;O;');
  MsgBox.Mess.Add('14;Journaux;Suppression impossible : ce journal est un modèle.;W;O;O;O;');
  MsgBox.Mess.Add('15;Journaux;Suppression impossible : ce journal est un journal d''ouverture ou de fermeture.;W;O;O;O;');
  MsgBox.Mess.Add('16;Journaux;Suppression impossible : ce journal est mouvementé par des écritures comptables issues de la gestion commerciale.;W;O;O;O;');
  MsgBox.Mess.Add('17;Erreur');
  MsgBox.Mess.Add('18;L''enregistrement est inaccessible');
  MsgBox.Mess.Add('19;Journaux;Ce journal est de nature analytique. Vous devez renseigner un axe analytique.;W;O;O;O;');
  MsgBox.Mess.Add('20;Journaux;Vous ne pouvez pas interdire le compte de contrepartie.;W;O;O;O;');
  MsgBox.Mess.Add('21;Journaux;Le compte saisi n''est pas valide pour cette régularisation.;W;O;O;O;');
  MsgBox.Mess.Add('22;Journaux;Ce compte est présent dans les comptes automatiques.;W;O;O;O;');
  MsgBox.Mess.Add('23;Journaux;Ce compte est présent dans les comptes interdits.;W;O;O;O;');
  MsgBox.Mess.Add('24;Journaux;Ce journal a un compte de contrepartie en devise. Vous devez renseigner la devise.;W;O;O;O;');
  MsgBox.Mess.Add('25;Journaux;Vous devez renseigner un compte de contrepartie existant.;W;O;O;O;');
  MsgBox.Mess.Add('26;Journaux;Vous devez renseigner un compteur;W;O;O;O;');
  MsgBox.Mess.Add('27;Journaux;Vous ne pouvez pas créer de jounal en "mode libre";W;O;O;O;');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2004
Modifié le ... : 08/03/2003
Description .. :
Suite ........ : GCO - 08/03/2003
Suite ........ : Remplacement GetField('J_JOURNAL') par GetControlText('J_JOURNAL')
Mots clefs ... :
*****************************************************************}

procedure TOM_JOURNAL.OnChangingPages(Sender: TObject; var AllowChange:
  Boolean);
begin
  if not (CtxPcl in V_Pgi.PGIContexte) then
    Exit;
  AllowChange := Trim(GetControlText('J_JOURNAL')) <> '';
  if not Allowchange then
  begin
    Pages.ActivePageIndex := PGeneral;
    SetFocusControl('J_JOURNAL');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/06/2005
Modifié le ... :   /  /    
Description .. : FQ 15897
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNAL.OnMouseDownJ_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DS.State = dsBrowse then
    DS.Edit;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOM_JOURNAL.AfficheCumuls;
var lStVal : String ;
begin

  // TOTAUX DU COMPTE
  // Maj des champs tampons pour affichage des montants d'infos
  lStVal := StrFMontant( GetChampsMS('J_TOTDEBP'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('J_TOTDEBP1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('J_TOTCREP'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('J_TOTCREP1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('J_TOTDEBE'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('J_TOTDEBE1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('J_TOTCREE'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('J_TOTCREE1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('J_TOTDEBS'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('J_TOTDEBS1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('J_TOTCRES'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('J_TOTCRES1', lStVal ) ;

  // Calcul des soldes
  AfficheLeSolde(THNumEdit(GetControl('JSOLCREP',true)), GetChampsMS('J_TOTDEBP'),    GetChampsMS('J_TOTCREP'));
  AfficheLeSolde(THNumEdit(GetControl('JSOLCREE',true)), GetChampsMS('J_TOTDEBE'),    GetChampsMS('J_TOTCREE'));
  AfficheLeSolde(THNumEdit(GetControl('JSOLCRES',true)), GetChampsMS('J_TOTDEBS'),    GetChampsMS('J_TOTCRES'));

end;

procedure TOM_JOURNAL.ChargeCumulsMultiSoc;
var lstReq  : String ;
    lQCumul : TQuery ;
begin

  if DS.State = dsInsert then Exit ;
  if not EstTablePartagee( 'JOURNAL' ) then Exit ;

  if Assigned( FTobCumuls ) then FreeAndNil( FTobCumuls ) ;

  lStReq := 'SELECT CU_DEBIT1 as J_TOTALDEBIT ,  CU_CREDIT1 as J_TOTALCREDIT ,'
                 + 'CU_DEBIT2 as J_TOTDEBP ,CU_CREDIT2 as J_TOTCREP ,'
                 + 'CU_DEBIT3 as J_TOTDEBE ,CU_CREDIT3 as J_TOTCREE ,'
                 + 'CU_DEBIT4 as J_TOTDEBS ,CU_CREDIT4 as J_TOTCRES '
             + ' FROM CUMULS WHERE CU_TYPE="' + fbToCumulType(fbJal) + '" '
                           + 'AND CU_COMPTE1="' + GetField('J_JOURNAL') + '"' ;

  lQCumul := OpenSQL( lStReq, True ) ;
  if not lQCumul.Eof then
    begin
    FTobCumuls := TOB.Create('_CUMULS', nil, -1 ) ;
    FTobCumuls.SelectDB('', lQCumul ) ;
    end ;
  Ferme(lQCumul);

end;

function TOM_JOURNAL.GetChampsMS(vStChamp: String): double;
begin

  result := 0 ;
  if DS.State = dsInsert then Exit ;

  if EstTablePartagee( 'JOURNAL' ) then
    begin
    if not Assigned( FTobCumuls ) then Exit ;
    result := FTobCumuls.GetDouble(  vStChamp ) ;
    end
  else result := Valeur( GetField( vStChamp ) ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 26/09/2007
Modifié le ... :   /  /    
Description .. : LG - 26/09/2007 - Fb 21503 - le avertir table n'etais plsu 
Suite ........ : lancer ds le formclose
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNAL.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

  if fAvertirTable then
    AvertirMultiTable('TTJOURNAL');

  // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
  if ( DS.State = dsInsert ) and
     ( GetControlText('J_JOURNAL') = '' ) then Exit ;

  if Assigned( FAGLCloseQuery ) then
    FAGLCloseQuery(Sender, Canclose ) ;

end;

initialization
  registerclasses([TOM_JOURNAL]);
end.

