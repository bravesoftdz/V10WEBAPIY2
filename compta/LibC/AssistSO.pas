{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. :
Suite ........ : GCO - 02/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Mots clefs ... : 
*****************************************************************}
unit AssistSO ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, ComCtrls, hmsgbox, StdCtrls, Hctrls, ExtCtrls, Mask,
  DBCtrls, DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Spin, HDB, Grids, DBGrids, HEnt1, {$IFDEF SPEC302} Societe,{$ENDIF} Ent1,
  FichComm, UTOMUTILISAT, USERGRP_TOM, Devise_TOM, Exercice_TOM, Tablette,
  hCompte, Menus, { BVE 03.05.07 Structur} CPSTRUCTURE_TOF,UTOTVentilType,
{$IFNDEF CCS3}
  CORRESP_TOM,ReparCle,
  //XMG 05/04/04 début
  //{$IFDEF ESP}
  CONTABON_TOM,
  (*{$ELSE}
  ContAbon, // A FAIRE
  {$ENDIF ESP} *)
  //XMG 05/04/04 fin
  Relance_TOM,
{$ENDIF}
  CPSection_TOM,Rupture,
  (*{$IFNDEF ESP}
  Tva,
{$ELSE}*)
  CPTVATPF_TOF,
//{$ENDIF ESP}//XMG 19/01/04
{$IFDEF V530}
     EdtDoc,
{$ELSE}
     EdtRDoc,
{$ENDIF}
  CodePost,SOUCHE_TOM,UtilSoc,Paramsoc,
  Refauto_Tom,  // ParamLib  BVE 25.06.07 FQ 20165
  // Scenario, BVE 21.05.07 FQ 19684
  SUIVCPTA_TOM,
{$IFDEF EAGLCLIENT}
  Rubrique_TOM,
{$ENDIF}
  CPJournal_TOM,
  CPGeneraux_TOM,
  CPAXE_TOM,     // remplace Axe,
  CPTiers_TOM,

  HSysMenu, HTB97, Hqry, UTOB, HPanel ;

type
  TFAssistSoc = class(TFAssist)
    Welcome: TTabSheet;
    HLabel1: THLabel;
    PlanRef: TTabSheet;
    SSociete: TDataSource;
    TSO_LIBELLE: THLabel;
    SO_LIBELLE: TEdit;
    HLabel2: THLabel;
    SO_CODE: TEdit;
    HLabel3: THLabel;
    SO_NUMPLANREF: TSpinEdit;
    HLabel4: THLabel;
    SPlanRef: TDataSource;
    QPlanRef: TQuery;
    bCreerPlanComptable: TToolbarButton97;
    GPlanRef: THDBGrid;
    ParamCPT: TTabSheet;
    Users: TTabSheet;
    HLabel5: THLabel;
    bUsers: TToolbarButton97;
    bUserGroups: TToolbarButton97;
    EtabExo: TTabSheet;
    HLabel7: THLabel;
    bEtab: TToolbarButton97;
    bExos: TToolbarButton97;
    bDevises: TToolbarButton97;
    bSoc1: TToolbarButton97;
    bSoc3: TToolbarButton97;
    bSoc4: TToolbarButton97;
    bSoc5: TToolbarButton97;
    HLabel6: THLabel;
    HLabel8: THLabel;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    StepMenu: TPopupMenu;
    ip1: TToolbarButton97;
    ip2: TToolbarButton97;
    FinEtape: TMenuItem;
    N1: TMenuItem;
    bUnites: TToolbarButton97;
    Reglements: TTabSheet;
    HLabel9: THLabel;
    GroupBox7: TGroupBox;
    bModePaiement: TToolbarButton97;
    bModeRegl: TToolbarButton97;
    Analytique: TTabSheet;
    HLabel10: THLabel;
    GroupBox8: TGroupBox;
    bAna1: TToolbarButton97;
    bAna2: TToolbarButton97;
    bAna3: TToolbarButton97;
    bAna4: TToolbarButton97;
    bAna5: TToolbarButton97;
    bAna6: TToolbarButton97;
    PopAxes: TPopupMenu;
    Axe1: TMenuItem;
    Axe2: TMenuItem;
    Axe3: TMenuItem;
    Axe4: TMenuItem;
    Axe5: TMenuItem;
    bStep: TToolbarButton97;
    nAna7: TToolbarButton97;
    PlanCpt: TTabSheet;
    HLabel11: THLabel;
    GroupBox9: TGroupBox;
    bCpt1: TToolbarButton97;
    bCpt2: TToolbarButton97;
    bCpt3: TToolbarButton97;
    TVA: TTabSheet;
    HLabel12: THLabel;
    GroupBox10: TGroupBox;
    bTaxes1: TToolbarButton97;
    bTaxes2: TToolbarButton97;
    bTaxes3: TToolbarButton97;
    Tiers: TTabSheet;
    HLabel13: THLabel;
    GroupBox11: TGroupBox;
    bTiers1: TToolbarButton97;
    bTiers2: TToolbarButton97;
    bTiers3: TToolbarButton97;
    bTiers4: TToolbarButton97;
    bTiers5: TToolbarButton97;
    bTiers6: TToolbarButton97;
    bTiers9: TToolbarButton97;
    Divers: TTabSheet;
    HLabel14: THLabel;
    GroupBox12: TGroupBox;
    bDivers1: TToolbarButton97;
    bDivers2: TToolbarButton97;
    bDivers3: TToolbarButton97;
    bDivers4: TToolbarButton97;
    bDivers5: TToolbarButton97;
    bSoc2: TToolbarButton97;
    bSoc6: TToolbarButton97;
    bSoc7: TToolbarButton97;
    bTiers10: TToolbarButton97;
    PopAxRupt: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    HDossierEuro: THLabel;
    DossierEuro: TCheckBox;
    procedure bSoc1Click(Sender: TObject);
    procedure bSoc2Click(Sender: TObject);
    procedure SO_NUMPLANREFChange(Sender: TObject);
    procedure GPlanRefDblClick(Sender: TObject);
    procedure bCreerPlanComptableClick(Sender: TObject);
    procedure bSoc3Click(Sender: TObject);
    procedure bSoc4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bUsersClick(Sender: TObject);
    procedure bUserGroupsClick(Sender: TObject);
    procedure bEtabClick(Sender: TObject);
    procedure bExosClick(Sender: TObject);
    procedure bDevisesClick(Sender: TObject);
    procedure bUnitesClick(Sender: TObject);
    procedure StepMenuClick(Sender: TObject);
    procedure PChange(Sender: TObject); Override ;
    procedure FinEtapeClick(Sender: TObject);
    procedure bModePaiementClick(Sender: TObject);
    procedure bModeReglClick(Sender: TObject);
    procedure bAna2Click(Sender: TObject);
    procedure bAna4Click(Sender: TObject);
    procedure bAna5Click(Sender: TObject);
    procedure bAna6Click(Sender: TObject);
    procedure AxesClick(Sender: TObject);
    procedure bCpt1Click(Sender: TObject);
    procedure bCpt2Click(Sender: TObject);
    procedure bCpt3Click(Sender: TObject);
    procedure bTaxes1Click(Sender: TObject);
    procedure bTaxes2Click(Sender: TObject);
    procedure bTaxes3Click(Sender: TObject);
    procedure bTiers1Click(Sender: TObject);
    procedure bTiers2Click(Sender: TObject);
    procedure bTiers3Click(Sender: TObject);
    procedure bTiers4Click(Sender: TObject);
    procedure bTiers5Click(Sender: TObject);
    procedure bTiers6Click(Sender: TObject);
    procedure bTiers9Click(Sender: TObject);
    procedure bDivers1Click(Sender: TObject);
    procedure bDivers2Click(Sender: TObject);
    procedure bDivers3Click(Sender: TObject);
    procedure bDivers4Click(Sender: TObject);
    procedure bDivers5Click(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    function  FirstPage : TTabSheet ; Override ;
    procedure bAna1Click(Sender: TObject);
    procedure bSoc5Click(Sender: TObject);
    procedure bSoc6Click(Sender: TObject);
    procedure bSoc7Click(Sender: TObject);
    procedure bTiers10Click(Sender: TObject);
    procedure AxesRuptClick(Sender: TObject);
  private
    Steps  : Array[1..20] of Boolean ;
    ChoixEuro,GeneCharge : boolean ;
    procedure CreerPlan ;
    procedure ChangeStepBitmap ;
    procedure ParamLaSoc(p : integer) ;
    procedure SauverLaSoc ( AvecEtapes : boolean ) ;
    procedure RestoreButton(b : TToolbarButton97) ;
    procedure VidePopAxes ;
    procedure SetCoord ;
//    Function  CodeDevisePris ( NewCode : String ) : boolean ;
    procedure EnregistreEuro ;
  public
  end;

procedure LanceAssistSO ;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  UtilPgi ;

procedure LanceAssistSO ;
Var X : TFAssistSoc ;
BEGIN
// GP le 02/09/98 if Not BlocageMonoPoste(True) then Exit ;
SourisSablier ;
X:=TFAssistSoc.Create(Application) ;
 Try
  X.ShowModal ;
 Finally
  X.Free ;
// GP le 02/09/98   DeblocageMonoPoste(True) ;
 End ;
SourisNormale ;
END ;

procedure TFAssistSoc.RestoreButton(b : TToolbarButton97) ;
BEGIN
//b.Restore ;
END ;

function TFAssistSoc.FirstPage : TTabSheet  ;
var i : Integer ;
BEGIN
//Result := Inherited FirstPage ;
result := Welcome ;
// positionnement sur la première étape non terminée
i := 1 ; while Steps[i] do Inc(i) ;
if i>1 then result := TTabSheet(P.Pages[i-1]) ;
END ;

procedure TFAssistSoc.SetCoord ;
BEGIN
{$IFNDEF SPEC302}
SO_CODE.Text:=GetParamSoc('SO_SOCIETE') ;
SO_LIBELLE.Text:=GetParamSoc('SO_LIBELLE') ;
{$ENDIF}
END ;

procedure TFAssistSoc.ParamLaSoc(p : integer) ;
Var
   sseule,sVirer : String ;
BEGIN
SourisSablier ;
SauverLaSoc(False) ;
sSeule:='' ; sVirer:='' ;
Case p of
   0 : sseule:='SCO_COORDONNEES' ;
   1 : sseule:='SCO_COMPTABLES' ;
   2 : sseule:='SCO_FOURCHETTES' ;
   3 : sseule:='SCO_COMPTESSPECIAUX' ;
   4 : sseule:='SCO_DATESDIVERS;SCO_DIVERS;SCO_CPFOLIO' ;
   5 : sseule:='SCO_REGLEMENTSTVA' ;
   6 : sseule:='SCO_EURO' ;
  10 : sseule:='SCO_REGLEMENTSTVA;SCO_EURO;SCO_AUTRES' ;
   END ;
if sSeule='' then sSeule:='SCO_PARAMETRESGENERAUX' ;
//ParamSociete(False,sVirer,sSeule,'',ChargeSocieteHalley,ChargePageSoc,SauvePageSocSansVerif,InterfaceSoc,1105000) ;
ParamSociete(False,sVirer,sSeule,'',ChargeSocieteHalley,ChargePageSoc,SauvePageSocSansVerif,InterfaceSoc,1105000) ;
if p=0 then SetCoord ;
SourisNormale ;
END ;

// Création du plan comptable
procedure TFAssistSoc.CreerPlan ;
var Q : TQuery ;
    s : String ;
BEGIN
if _Blocage(['nrBatch','nrCloture'],True,'nrCloture') then Exit ;
Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+W_W+'"',False) ;
while not QPlanRef.EOF do
   BEGIN
   s := BourreLaDonc(QPlanRef.FindField('PR_COMPTE').AsString,fbGene) ;
   if Not Presence('GENERAUX', 'G_GENERAL', s) then
      BEGIN
      Q.Insert ;
      InitNew(Q) ;
      Q.FindField('G_GENERAL').AsString := s ;
      Q.FindField('G_LIBELLE').AsString := QPlanRef.FindField('PR_LIBELLE').AsString ;
      Q.FindField('G_ABREGE').AsString := QPlanRef.FindField('PR_ABREGE').AsString ;
      Q.FindField('G_CENTRALISABLE').AsString := QPlanRef.FindField('PR_CENTRALISABLE').AsString ;
      Q.FindField('G_SOLDEPROGRESSIF').AsString := QPlanRef.FindField('PR_SOLDEPROGRESSIF').AsString ;
      Q.FindField('G_SAUTPAGE').AsString := QPlanRef.FindField('PR_SAUTPAGE').AsString ;
      Q.FindField('G_TOTAUXMENSUELS').AsString := QPlanRef.FindField('PR_TOTAUXMENSUELS').AsString ;
      Q.FindField('G_COLLECTIF').AsString := QPlanRef.FindField('PR_COLLECTIF').AsString ;
      Q.FindField('G_BLOCNOTE').AsString := QPlanRef.FindField('PR_BLOCNOTE').AsString ;
      Q.FindField('G_SENS').AsString := QPlanRef.FindField('PR_SENS').AsString ;
      Q.FindField('G_LETTRABLE').AsString := QPlanRef.FindField('PR_LETTRABLE').AsString ;
      Q.FindField('G_POINTABLE').AsString := QPlanRef.FindField('PR_POINTABLE').AsString ;
      Q.FindField('G_VENTILABLE1').AsString := QPlanRef.FindField('PR_VENTILABLE1').AsString ;
      Q.FindField('G_VENTILABLE2').AsString := QPlanRef.FindField('PR_VENTILABLE2').AsString ;
      Q.FindField('G_VENTILABLE3').AsString := QPlanRef.FindField('PR_VENTILABLE3').AsString ;
      Q.FindField('G_VENTILABLE4').AsString := QPlanRef.FindField('PR_VENTILABLE4').AsString ;
      Q.FindField('G_VENTILABLE5').AsString := QPlanRef.FindField('PR_VENTILABLE5').AsString ;
      Q.FindField('G_NATUREGENE').AsString := QPlanRef.FindField('PR_NATUREGENE').AsString ;
      Q.Post ;
      END ;
   QPlanRef.Next ;
   END ;
Ferme(Q) ;
_Bloqueur('nrCloture',False) ;
END ;

procedure TFAssistSoc.bSoc1Click(Sender: TObject);
begin
ParamLaSoc(0) ;
RestoreButton(bSoc1) ;
end;

procedure TFAssistSoc.bSoc2Click(Sender: TObject);
begin
ParamLaSoc(1) ;
RestoreButton(bSoc2) ;
end;

procedure TFAssistSoc.bSoc5Click(Sender: TObject);
begin
// Autres paramètres
{$IFDEF SPEC302}
ParamLaSoc(1) ;
{$ELSE}
ParamLaSoc(4) ;
{$ENDIF}
RestoreButton(bSoc5) ;
end;

procedure TFAssistSoc.bSoc6Click(Sender: TObject);
begin
{$IFDEF SPEC302}
ParamLaSoc(1) ;
{$ELSE}
ParamLaSoc(10) ; 
{$ENDIF}
RestoreButton(bSoc6) ;
end;

procedure TFAssistSoc.SO_NUMPLANREFChange(Sender: TObject);
begin
if SO_NUMPLANREF.Value = 0 then
   BEGIN
   QPlanRef.Close ;
   QPlanRef.SQL.Clear ;
   GPlanRef.Enabled := False ;
   bCreerPlanComptable.Enabled := False ;
   END
else
   BEGIN
   ExecuteSQL('UPDATE PLANREF SET PR_REPORTDETAIL="X" WHERE PR_NUMPLAN = ' + IntToStr(SO_NUMPLANREF.Value) + ' AND PR_REPORTDETAIL = "-"') ;
   QPlanRef.SQL.Clear ;
   QPlanRef.SQL.Add('SELECT * FROM PLANREF WHERE PR_NUMPLAN = ' + IntToStr(SO_NUMPLANREF.Value) + ' ORDER BY PR_COMPTE') ;
   ChangeSQL(QPlanRef) ;
   QPlanRef.Open ;
   GPlanRef.Enabled := True ;
   bCreerPlanComptable.Enabled := True ;
   END ;
end;

procedure TFAssistSoc.GPlanRefDblClick(Sender: TObject);
begin
if QPlanRef.Eof then Exit ;
QPlanRef.Edit ;
if QPlanRef.FindField('PR_REPORTDETAIL').AsString = '-' then
   QPlanRef.FindField('PR_REPORTDETAIL').AsString := 'X' else
   QPlanRef.FindField('PR_REPORTDETAIL').AsString := '-' ;
QPlanRef.Post ;
end;

procedure TFAssistSoc.bCreerPlanComptableClick(Sender: TObject);
begin
if ((GetParamSoc('SO_LGCPTEGEN')=0) or (Trim(GetParamSoc('SO_BOURREGEN'))='')) then
   begin
   Msg.Execute(2,Caption,'') ;
   Exit ;
   end ;
CreerPlan ;
end;

procedure TFAssistSoc.bSoc3Click(Sender: TObject);
begin
ParamLaSoc(2) ; // fourchettes
RestoreButton(bSoc3) ;
end;

procedure TFAssistSoc.bSoc4Click(Sender: TObject);
begin
ParamLaSoc(3) ; // comptes spéciaux
RestoreButton(bSoc4) ;
end;


procedure TFAssistSoc.VidePopAxes ;
Var M,N : TMenuItem ;
    PP  : TPopupMenu ;
    Limite : integer ;
BEGIN
if EstSerie(S3) then Limite:=1 else Limite:=5 ;		// Modification
PP:=BAna3.DropDownMenu ;
if PP=Nil then Exit ;
if PP.Items.Count<=0 then Exit ;
While PP.Items.Count>Limite do
   BEGIN
   M:=PP.Items[PP.Items.Count-1] ;
   While M.Count>0 do BEGIN N:=M.Items[0] ; M.Remove(N) ; N.Free ; END ;
   PP.Items.Remove(M) ; M.Free ;
   END ;
PP:=nAna7.DropDownMenu ;
if PP=Nil then Exit ;
if PP.Items.Count<=0 then Exit ;
While PP.Items.Count>Limite do
   BEGIN
   M:=PP.Items[PP.Items.Count-1] ;
   While M.Count>0 do BEGIN N:=M.Items[0] ; M.Remove(N) ; N.Free ; END ;
   PP.Items.Remove(M) ; M.Free ;
   END ;
END ;

procedure TFAssistSoc.FormShow(Sender: TObject);
var i,iNum : Integer ;
    mi : TMenuItem ;
    s,sCode,SLibelle,sEtapes  : String ;
    Q  : TQuery ;
    VisEuro  : Boolean ;
begin
GeneCharge:=True ;
for i:=low(Steps) to high(Steps) do Steps[i] := False ;
sCode:='' ; sLibelle:='' ; sEtapes:='' ; ChoixEuro:=False ;
{$IFDEF SPEC302}
Q:=OpenSQL('Select SO_SOCIETE, SO_LIBELLE, SO_ETAPES, SO_NUMPLANREF from SOCIETE',True) ;
if Not Q.EOF then
   BEGIN
   sCode:=Q.Fields[0].AsString ; sLibelle:=Q.Fields[1].AsString ;
   sEtapes:=Q.Fields[2].AsString ; iNum:=Q.Fields[3].AsInteger ;
   END ;
Ferme(Q) ;
VisEuro:=False ;
{$ELSE}
Q:=OpenSQL('Select SO_SOCIETE from SOCIETE',True) ;
if Not Q.EOF then sCode:=Q.Fields[0].AsString ;
Ferme(Q) ;
sLibelle:=GetParamSoc('SO_LIBELLE') ; sEtapes:=GetParamSoc('SO_ETAPES') ;
iNum:=GetParamSoc('SO_NUMPLANREF') ;
BSoc7.Visible:=False ;
VisEuro:=True ;
if VH^.TenueEuro then VisEuro:=False ;
if V_PGI.DevisePivot='' then VisEuro:=False ;
if VisEuro then if ExisteSQL('SELECT E_JOURNAL FROM ECRITURE') then VisEuro:=False ;
if VisEuro then if ExisteSQL('SELECT Y_JOURNAL FROM ANALYTIQ') then VisEuro:=False ;
if VisEuro then if ExisteSQL('SELECT BE_BUDJAL FROM BUDECR') then VisEuro:=False ;
{$ENDIF}
if Not VisEuro then BEGIN DossierEuro.Visible:=False ; HDossierEuro.Visible:=False ; END ;
SO_CODE.Text:=sCode ; SO_LIBELLE.Text:=sLibelle ; SO_NUMPLANREF.Value:=iNum ;
s:=sEtapes ;
for i:=1 to length(s) do
    BEGIN
    if Copy(s,i,1)='1' then Steps[i]:=True else Steps[i]:=False ;
    END ;
for i:=0 to P.ControlCount-1 do
   BEGIN
   mi := TMenuItem.Create(Self);
   mi.Caption := TTabSheet(P.Controls[i]).Hint ;
   mi.OnClick := StepMenuClick ;
   StepMenu.Items.Add(mi);
   END ;
inherited ;
SourisNormale ;
Caption:=NomHalley+' - '+Caption ; UpdateCaption(Self) ;
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
   VidePopAxes ;
   BAna6.Visible:=False ;
   {$IFDEF CCS3}
   BAna3.Visible:=False ;
   BCpt1.Visible:=False ;
   BTiers6.Visible:=False ;
   BUnites.Visible:=False ;
   BTaxes3.Visible:=False ;
   BTiers3.Enabled:=FALSE ;
   BTiers4.Enabled:=FALSE ;
   BDivers5.enabled:=FALSE ;
   {$ENDIF}
   END ;
GeneCharge:=False ;
end;

procedure TFAssistSoc.bUsersClick(Sender: TObject);
begin
FicheUser('') ;
RestoreButton(bUsers) ;
end;

procedure TFAssistSoc.bUserGroupsClick(Sender: TObject);
begin
FicheUserGrp ;
RestoreButton(bUserGroups) ;
end;

procedure TFAssistSoc.bEtabClick(Sender: TObject);
begin
FicheEtablissement_AGL(taCreat) ;
RestoreButton(bEtab) ;
end;

procedure TFAssistSoc.bExosClick(Sender: TObject);
begin
//FicheExercice('',taCreat,1) ;
YYLanceFiche_Exercice('1');
RestoreButton(bExos) ;
end;

procedure TFAssistSoc.bDevisesClick(Sender: TObject);
begin
FicheDevise('',taModif,True) ;
RestoreButton(bDevises) ;
end;

procedure TFAssistSoc.bUnitesClick(Sender: TObject);
begin
ParamTable('ttQualUnitMesure',tacreat,1190030,Nil) ;
RestoreButton(bUnites) ;
end;

procedure TFAssistSoc.StepMenuClick(Sender: TObject);
var i : Integer ;
begin
RestorePage ;
i := TMenuItem(Sender).MenuIndex-2 ;
P.ActivePage := P.Pages[i] ;
PChange(nil) ;
end;

procedure TFAssistSoc.PChange(Sender: TObject);
begin
inherited;
FinEtape.Checked := Steps[P.ActivePage.PageIndex+1] ;
ChangeStepBitmap ;
if Not GeneCharge then
   BEGIN
   if Not ChoixEuro then EnregistreEuro ;
   if DossierEuro.Checked then ChoixEuro:=True ;
   END ;
end ;

procedure TFAssistSoc.ChangeStepBitmap ;
begin
if Steps[P.ActivePage.PageIndex + 1] then bStep.Glyph := ip2.Glyph
                                     else bStep.Glyph := ip1.Glyph ;
end ;

procedure TFAssistSoc.FinEtapeClick(Sender: TObject);
begin
FinEtape.Checked := not FinEtape.Checked ;
Steps[P.ActivePage.PageIndex + 1] := FinEtape.Checked ;
ChangeStepBitmap ;
end;

procedure TFAssistSoc.bModePaiementClick(Sender: TObject);
begin
FicheModePaie_AGL('');
RestoreButton(bModePaiement) ;
end;

procedure TFAssistSoc.bModeReglClick(Sender: TObject);
begin
FicheRegle_AGL('',FALSE,taModif) ;
RestoreButton(bModeRegl) ;
end;

procedure TFAssistSoc.bAna2Click(Sender: TObject);
begin
ParamPlanAnal('') ;
RestoreButton(bAna2) ;
end;

procedure TFAssistSoc.bAna4Click(Sender: TObject);
begin
FicheSection(Nil,'A1','',taCreatEnSerie,0) ;
RestoreButton(bAna4) ;
end;

procedure TFAssistSoc.bAna5Click(Sender: TObject);
begin
ParamVentilType ;
RestoreButton(bAna5) ;
end;

procedure TFAssistSoc.bAna6Click(Sender: TObject);
begin
{$IFNDEF CCS3}
ParamRepartCle ;
RestoreButton(bAna6) ;
{$ENDIF}
end;

procedure TFAssistSoc.AxesClick(Sender: TObject);
var i : Integer ;
begin
{$IFNDEF CCS3}
i := TMenuItem(Sender).MenuIndex + 1 ;
CCLanceFiche_Correspondance('A' + IntToStr(i)) ;	// Modification
{$ENDIF}
end;

procedure TFAssistSoc.bCpt1Click(Sender: TObject);
begin
{$IFNDEF CCS3}
CCLanceFiche_Correspondance('GE') ;			// Modification
RestoreButton(bCpt1) ;
{$ENDIF}
end;

procedure TFAssistSoc.bCpt2Click(Sender: TObject);
begin
FicheGene(Nil,'','',taCreatEnSerie,0) ;
RestoreButton(bCpt2) ;
end;

procedure TFAssistSoc.bCpt3Click(Sender: TObject);
begin
PlanRupture('RUG') ;
RestoreButton(bCpt3) ;
end;

procedure TFAssistSoc.bTaxes1Click(Sender: TObject);
begin
ParamTable('TTREGIMETVA',taCreat,1165000,Nil) ;
RestoreButton(bTaxes1) ;
end;

procedure TFAssistSoc.bTaxes2Click(Sender: TObject);
begin
ParamTvaTpf(True) ;
RestoreButton(bTaxes2) ;
end;

procedure TFAssistSoc.bTaxes3Click(Sender: TObject);
begin
ParamTvaTpf(False) ;
RestoreButton(bTaxes3) ;
end;

procedure TFAssistSoc.bTiers1Click(Sender: TObject);
begin
ParamTable('ttFormeJuridique',tacreat,1120000,Nil) ;
RestoreButton(bTiers1) ;
end;

procedure TFAssistSoc.bTiers2Click(Sender: TObject);
begin
ParamTable('ttFonction',tacreat,1125000,Nil) ;
RestoreButton(bTiers2) ;
end;

procedure TFAssistSoc.bTiers3Click(Sender: TObject);
begin
{$IFDEF CCS3}
{$ELSE}
//ParamRelance('RTR','',taModif) ;
CCLanceFiche_ParamRelance('RTR');
RestoreButton(bTiers3) ;
{$ENDIF}
end;

procedure TFAssistSoc.bTiers4Click(Sender: TObject);
begin
{$IFDEF CCS3}
{$ELSE}
//ParamRelance('RRG','',taModif) ;
CCLanceFiche_ParamRelance('RRG');
RestoreButton(bTiers4) ;
{$ENDIF}
end;

procedure TFAssistSoc.bTiers5Click(Sender: TObject);
begin
PlanRupture('RUT') ;
RestoreButton(bTiers5) ;
end;

procedure TFAssistSoc.bTiers6Click(Sender: TObject);
begin
{$IFNDEF CCS3}
CCLanceFiche_Correspondance('AU') ;			// Modification
RestoreButton(bTiers6) ;
{$ENDIF}
end;

procedure TFAssistSoc.bTiers9Click(Sender: TObject);
begin
FicheTiers(Nil,'','',taCreatEnSerie,1) ;
RestoreButton(bTiers9) ;
end;

procedure TFAssistSoc.bDivers1Click(Sender: TObject);
begin
  FicheJournal(Nil,'','',taCreatEnSerie,0) ;
  RestoreButton(bDivers1) ;
end;

procedure TFAssistSoc.bDivers2Click(Sender: TObject);
begin
  YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT');
  RestoreButton(bDivers2) ;
end;

procedure TFAssistSoc.bDivers3Click(Sender: TObject);
begin
ParamLibelle ;
RestoreButton(bDivers3) ;
end;

procedure TFAssistSoc.bDivers4Click(Sender: TObject);
begin
ParamScenario('','') ;
RestoreButton(bDivers4) ;
end;

procedure TFAssistSoc.bDivers5Click(Sender: TObject);
begin
{$IFDEF CCS3}
{$ELSE}
ParamAbonnement(TRUE,'',taModif) ;
RestoreButton(bDivers5) ;
{$ENDIF}
end;

(*
Function TFAssistSoc.CodeDevisePris ( NewCode : String ) : boolean ;
BEGIN
Result:=ExisteSQL('SELECT D_DEVISE FROM DEVISE WHERE D_DEVISE="'+NewCode+'"') ;
END ;
*)

{=============================================================================}
procedure TFAssistSoc.EnregistreEuro ;
(*
Var OldFF : String ;
    TOBPiv,TOBContr : TOB ;
    QD   : TQuery ;
*)
BEGIN
DossierEuro.Enabled:=False ;
if Not DossierEuro.Visible then Exit ;
if Not DossierEuro.Checked then Exit ;
if V_PGI.DevisePivot='' then Exit ;
CreerDeviseTenue('E') ;
(*
OldFF:=V_PGI.DevisePivot ;
TOBPiv:=TOB.Create('DEVISE',Nil,-1) ;
QD:=OpenSQL('SELECT * FROM DEVISE WHERE D_DEVISE="'+V_PGI.DevisePivot+'"',True) ;
if Not QD.EOF then TOBPiv.SelectDB('',QD) ;
Ferme(QD) ;
TOBContr:=TOB.Create('DEVISE',Nil,-1) ; TOBContr.Dupliquer(TOBPiv,False,True,False) ;
// TOBPiv reste pivot et donc devient l'Euro
TOBPiv.PutValue('D_LIBELLE','Euro') ; TOBPiv.PutValue('D_SYMBOLE','') ;
TOBPiv.PutValue('D_FONGIBLE','-')   ; TOBPiv.PutValue('D_MONNAIEIN','-') ;
TOBPiv.PutValue('D_DECIMALE',2)     ; TOBPiv.PutValue('D_QUOTITE',1) ;
TOBPiv.PutValue('D_PARITEEURO',1)   ;
TOBPiv.PutValue('D_CODEISO','')     ;
// TOBContr devient le franc, rien à changer sauf le code devise
OldFF:='' ;
if Not CodeDevisePris('FF') then OldFF:='FF' ;
if OldFF='' then if Not CodeDevisePris('FFF') then OldFF:='FFF' ;
if OldFF='' then if Not CodeDevisePris('FRF') then OldFF:='FRF' ;
if OldFF='' then OldFF:='SRF' ;
TOBContr.PutValue('D_DEVISE',OldFF) ;
// Insertions
TOBContr.SetAllModifie(True) ; TOBPiv.SetAllModifie(True) ;
TOBContr.UpdateDB(False) ;
TOBPiv.InsertDB(Nil) ;
SetParamSoc('SO_TENUEEURO',True) ; VH^.TenueEuro:=True ;
*)
SetParamSoc('SO_TAUXEURO',1) ;
END ;

procedure TFAssistSoc.SauverLaSoc ( AvecEtapes : boolean ) ;
var s : String ;
    i : Integer ;
BEGIN
s:='' ;
if AvecEtapes then
   BEGIN
   for i:=1 to high(Steps) do BEGIN if Steps[i] then s := s + '1' else s := s + '0' ; END ;
   END ;
{$IFDEF SPEC302}
if AvecEtapes then ExecuteSQL('UPDATE SOCIETE SET SO_SOCIETE="'+SO_CODE.Text+'", SO_LIBELLE="'+SO_LIBELLE.Text+'", SO_ETAPES="'+s+'", SO_NUMPLANREF='+IntToStr(SO_NUMPLANREF.Value))
              else ExecuteSQL('UPDATE SOCIETE SET SO_SOCIETE="'+SO_CODE.Text+'", SO_LIBELLE="'+SO_LIBELLE.Text+'", SO_NUMPLANREF='+IntToStr(SO_NUMPLANREF.Value)) ;
{$ELSE}
ExecuteSQL('UPDATE SOCIETE SET SO_SOCIETE="'+SO_CODE.Text+'"') ;
SetParamSoc('SO_SOCIETE',SO_CODE.Text) ; SetParamSoc('SO_LIBELLE',SO_LIBELLE.Text) ;
SetParamSoc('SO_NUMPLANREF',SO_NUMPLANREF.Value) ;
if AvecEtapes then SetParamSoc('SO_ETAPES',s) ;
if Not ChoixEuro then EnregistreEuro ;
ChoixEuro:=True ;
{$ENDIF}
END ;

procedure TFAssistSoc.bFinClick(Sender: TObject);
begin
inherited;
// Sauvegarde de la chaîne des étapes
SauverLaSoc(True) ;
Close ;
end;

procedure TFAssistSoc.bAna1Click(Sender: TObject);
begin
FicheAxeAna ('') ;
RestoreButton(bAna1) ;
end;

procedure TFAssistSoc.bSoc7Click(Sender: TObject);
begin
// Comptes collectifs
{$IFDEF SPEC302}
ParamLaSoc(4) ;
{$ELSE}
ParamLaSoc(1) ;
{$ENDIF}
RestoreButton(bSoc7) ;
end;

procedure TFAssistSoc.bTiers10Click(Sender: TObject);
begin
ParamTable('ttCivilite',tacreat,1122000,Nil) ;
RestoreButton(bTiers10) ;
end;

procedure TFAssistSoc.AxesRuptClick(Sender: TObject);
var i : Integer ;
begin
i:=TMenuItem(Sender).MenuIndex + 1 ;
PlanRupture('RU' + IntToStr(i)) ;
end;

end.
