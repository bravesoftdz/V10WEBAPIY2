{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : yTarifsAQui_TOF ()
Mots clefs ... : TOF;yTarifsAQui_TOF
*****************************************************************}
Unit TarifsAQui_TOF ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  HTB97,
  {$IFNDEF EAGLCLIENT}
    db,
    dbTables,
    Fe_Main,
    Fiche,
    Mul,
  {$ELSE}
    MainEagl,
    eFiche,
    eMul,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  HPanel,
  Menus,
  UTOF,
  wTOF,
  uTob;

Type
  TOF_tarifsAQui = Class (twTOF)
  procedure OnNew                    ; override ;
  procedure OnDelete                 ; override ;
  procedure OnUpdate                 ; override ;
  procedure OnLoad                   ; override ;
  procedure OnArgument (S : String ) ; override ;
  procedure OnClose                  ; override ;
  private
    sFonctionnalite            : string;
    sAction, sDroit, sAppel    : string;
    iDeIdentifiantYTA          : integer;
    sParametreYTQ              : string;

    lCommissionnement : boolean;

    BAccesSurQuoi      : tToolbarButton97;

    { Click sur les boutons }
    procedure FLISTE_OnDblClick(Sender: TOBject);
    procedure BINSERT_OnClick(Sender: TOBject);
    procedure BDUPLICATION_OnClick(Sender: TOBject);

    { Click sur le menu du bouton bAccesSurQuoi }
    procedure MnSaisieGrille_OnClick(Sender: TObject);
    procedure MnSaisieOnglet_OnClick(Sender: TObject);
    procedure MnSaisieMixte_OnClick(Sender: TObject);
    procedure BACCESSURQUOI_OnClick(Sender: TOBject);

  end ;

Implementation

uses
  wCommuns,
  tarifs,
  ParamSoc
  ;

Const
  sFournisseur = 'FOU';
  sClient      = 'CLI';

  sCommissionFournisseur  = '003';   //Fonctionnalité : commissionnement fournisseur
  sCommissionClient       = '004';   //Fonctionnalité : commissionnement client

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.OnArgument (S : String ) ;
var
  sNatureAuxi : string;
begin
	{ Init. pour la WTOF }
  FTableName := 'YTARIFSAQUI';
  FFiche     := 'YTARIFSAQUI_FIC';
  FLequel    := WMakeFieldString(fTableName, ';');
  Inherited ;

  { Nature de la fiche pour AGLLanceFiche }
  fNature := 'Y';

  { Argument }
  sFonctionnalite   := GetArgumentValue(S, 'YTA_FONCTIONNALITE');
  sNatureAuxi       := NatureAuxiliaire(sFonctionnalite);
  sAction           := GetArgumentValue(S, 'ACTION');
  sDroit            := GetArgumentValue(S, 'DROIT');
  sAppel            := GetArgumentValue(S, 'APPEL');
  sParametreYTQ     := iif((sAppel='ARTICLE'),';APPEL='+sAppel+';YTQ_ARTICLE='+GetArgumentValue(S,'YTQ_ARTICLE')+';YTQ_TARIFARTICLE='+GetArgumentValue(S,'YTQ_TARIFARTICLE'),'');

  Ecran.Caption     := RechDom('YFONCTIONNALITES',sFonctionnalite,False);
  UpdateCaption(Ecran);

  { Evénements }
  if (GetControl('FLISTE')        <> nil) then THGrid(GetControl('FLISTE' )).OnDblClick              := FLISTE_OnDblClick ;
  if (GetControl('BINSERT')       <> nil) then TToolBarButton97(GetControl('BINSERT')).OnClick       := BINSERT_OnClick ;
  if (GetControl('BDUPLICATION')  <> nil) then TToolBarButton97(GetControl('BDUPLICATION')).OnClick  := BDUPLICATION_OnClick ;

  { Menu associé au bouton BAccesSurQuoi }
  BAccesSurquoi := tToolbarButton97(GetControl('BACCESSURQUOI'));
  if (GetControl('BACCESSURQUOI') <> nil) then TToolBarButton97(GetControl('BACCESSURQUOI')).OnClick := BACCESSURQUOI_OnClick ;
	if (GetControl('MNSAISIEGRILLE')<> nil) then TMenuItem(GetControl('MNSAISIEGRILLE')).OnClick := MnSaisieGrille_OnClick;
	if (GetControl('MNSAISIEONGLET')<> nil) then TMenuItem(GetControl('MNSAISIEONGLET')).OnClick := MnSaisieOnglet_OnClick;
	if (GetControl('MNSAISIEMIXTE') <> nil) then TMenuItem(GetControl('MNSAISIEMIXTE') ).OnClick := MnSaisieMixte_OnClick;


  if (GetControl('YTA_TARIFSPECIAL' ) <> nil) then THEdit(GetControl('YTA_TARIFSPECIAL' )).Plus := 'YTP_FONCTIONNALITE="'+sFonctionnalite+'"';
  if (GetControl('YTA_TARIFSPECIAL_') <> nil) then THEdit(GetControl('YTA_TARIFSPECIAL_')).Plus := 'YTP_FONCTIONNALITE="'+sFonctionnalite+'"';

  if    (sNatureAuxi=sFournisseur) then
  begin
    if (GetControl('YTA_TIERS')      <> nil) then THEdit(GetControl('YTA_TIERS' )     ).DataType := 'GCTIERSFOURN';
    if (GetControl('YTA_TIERS_')     <> nil) then THEdit(GetControl('YTA_TIERS_')     ).DataType := 'GCTIERSFOURN';
    if (GetControl('YTA_TARIFTIERS' )<> nil) then THValComboBox(GetControl('YTA_TARIFTIERS' )).DataType := 'TTTARIFFOURNISSEUR';
    if (GetControl('YTA_TARIFTIERS_')<> nil) then THValComboBox(GetControl('YTA_TARIFTIERS_')).DataType := 'TTTARIFFOURNISSEUR';
  end
  else if (sNatureAuxi=sClient) then
  begin
    if (GetControl('YTA_TIERS' )     <> nil) then THEdit(GetControl('YTA_TIERS' )     ).DataType := 'GCTIERSCLI';
    if (GetControl('YTA_TIERS_')     <> nil) then THEdit(GetControl('YTA_TIERS_')     ).DataType := 'GCTIERSCLI';
    if (GetControl('YTA_TARIFTIERS' )<> nil) then THValComboBox(GetControl('YTA_TARIFTIERS' )).DataType := 'TTTARIFCLIENT';
    if (GetControl('YTA_TARIFTIERS_')<> nil) then THValComboBox(GetControl('YTA_TARIFTIERS_')).DataType := 'TTTARIFCLIENT';
  end;
  lCommissionnement := (sFonctionnalite=sCommissionClient) or (sFonctionnalite=sCommissionFournisseur);
  if (GetControl('PNCOMMISSIONNEMENT') <> nil) then THPanel(GetControl('PNCOMMISSIONNEMENT')).Visible   := lCommissionnement;

  { Variables utiles pour la duplication du système tarifaire}
  iDeIdentifiantYTA := 0; 
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.OnClose ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.OnNew ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.OnDelete ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.OnUpdate ;
var
  fMul : TFMul;
  iCpt    : integer;
begin
  inherited ;
  //Affichage des colonnes du mul
  fMul := TFMul(Ecran);

  //Si ecran de commissionnement on fait apparaitre dans la liste le commercial et le type de commercial
  {$IFNDEF EAGLCLIENT}
  for iCpt := 0 to fMul.FListe.Columns.Count-1 do
  begin
    if (fMul.FListe.Columns[iCpt].FieldName='YTA_COMMERCIAL') or (fMul.FListe.Columns[iCpt].FieldName='YTA_TYPECOMMERCIAL') then
        fMul.FListe.Columns[iCpt].Visible:=lCommissionnement;
  end;
  {$ELSE}
  for iCpt := 0 to fMul.FListe.ColCount-1 do
  begin
    // A faire voir DK
  end;
  {$ENDIF}
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.OnLoad ;
begin
  Inherited ;

//  if      (sAppel='MENU'   ) then SetControlText('XX_WHERE','((YTA_TIERS="") or ((YTA_TIERS<>"") and (T_NATUREAUXI="'+NatureAuxiliaire(sFonctionnalite)+'")))')
//  else if (sAppel='TIERS'  ) then SetControlText('XX_WHERE','(T_NATUREAUXI="'+NatureAuxiliaire(sFonctionnalite)+'")')
//  else
  if (sAppel='ARTICLE') then SetControlText('XX_WHERE','((YTA_TIERS="") and (YTA_TARIFTIERS=""))');
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.FLISTE_OnDblClick(Sender: TOBject);
begin
  FParamsLanceFiche := 'ACTION=MODIFICATION'+';DROIT='+sDroit+';YTA_FONCTIONNALITE=' + sFonctionnalite +';YTA_IDENTIFIANT=' + IntToStr(iDeIdentifiantYTA)+';METHODEDESAISIE=ONGLET'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ;
  inherited;
end;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.BINSERT_OnClick(Sender: TOBject);
begin
  FParamsLanceFiche := 'DROIT='+sDroit+';YTA_FONCTIONNALITE=' + sFonctionnalite +';YTA_IDENTIFIANT=' + IntToStr(iDeIdentifiantYTA)+';METHODEDESAISIE=ONGLET'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ;
  inherited;
  iDeIdentifiantYTA:=0; 
end;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.BDUPLICATION_OnClick(Sender: TOBject);
begin
  inherited;
  iDeIdentifiantYTA   := GetInteger('YTA_IDENTIFIANT');
  BINSERT_OnClick(Sender);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Sur Quoi en saisie en grille
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.MnSaisieGrille_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+sAction+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE=GRILLE'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Sur Quoi en saisie en onglet
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.MnSaisieOnglet_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+sAction+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE=ONGLET'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Sur Quoi en saisie mixte
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.MnSaisieMixte_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+sAction+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE=MIXTE'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Sur Quoi selon le paramètre
---------------------------------------------------------------------------------}
procedure TOF_tarifsAQui.BACCESSURQUOI_OnClick(Sender: TOBject);
var
  sAppelMethodeSaisie, sParamMethodeSaisie : string;
begin
  sParamMethodeSaisie := GetParamSoc('SO_TARIFSAISIE');
  if      (sParamMethodeSaisie = 'O') then sAppelMethodeSaisie := 'ONGLET'
  else if (sParamMethodeSaisie = 'G') then sAppelMethodeSaisie := 'GRILLE'
  else if (sParamMethodeSaisie = 'M') then sAppelMethodeSaisie := 'MIXTE'
  else                                     sAppelMethodeSaisie := 'MIXTE';

  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+sAction+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE='+sAppelMethodeSaisie+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

Initialization
  registerclasses ( [ TOF_tarifsAQui ] ) ;
end.

