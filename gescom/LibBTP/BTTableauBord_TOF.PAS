{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTABLEAUBORD ()
Mots clefs ... : TOF;BTTABLEAUBORD
*****************************************************************}
unit BTTableauBord_TOF;

interface

uses StdCtrls,
  Controls,
  DialogEx,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  Menus,
  HCtrls,
  HEnt1,
  HQry,
  HMsgBox,
  UTOF,
  Stat,
  UTOB,
  utobview,
  AGLInit,
  AGLInitBTP,
  HTB97,
  facture,
  Btputil,
  UtilSaisieConso,
  UtilReglementAffaire,
  uTofAfBaseCodeAffaire,
  Affaireutil,
  Lookup,
  Ent1,
  CalcOLEGenericBTP,
  StrUtils,
  //
  QRS1,
  EdtEtat,
  EdtDoc,
  EdtREtat,
  EdtRDoc,
  //
  UtilsEtat,
  UtilTableauBord,
  UtilsGrille,
  UtilPGI;

type

  TOF_BTTABLEAUBORD = class(TOF_AFBASECODEAFFAIRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;

  private
    TV1           : TTobViewer;
    //
    TTForm        : TToolWindow97;
    TTGrille      : THGrid;
    GridEvent     : TGestionGS;
    //
    TOBBTB        : TOB;
    //
    DateDeb       : TdateTime;
    DateFin       : TDateTime;
    //
    TResponsable      : THEdit;
    TDomaine          : THEdit;
    TEtablissement    : THEdit;
    LibResponsable    : THLabel;
    LibDomaine        : THLabel;
    LibEtablissement  : THLabel;        
    //
    CodeAffaire   : String;
    CodeAffaire0  : String;
    CodeAffaire1  : String;
    CodeAffaire2  : String;
    CodeAffaire3  : string;
    CodeAvenant   : string;
    CodeClient    : String;
    Etablissement : string;
    CodeChantier  : String;
    Apporteur     : string;
    Responsable   : String;
    //
    fModeleAnaXls : String;
    //
    // Ajout LS
    DateMvtDeb    : TdateTime;
    DateMvtFin    : TdateTime;
    //
    OptionChoixPrevuAvanc : TchoixPrevuAvanc;
    // ---

    // --- Ajout FV ---
    CChantier     : TCheckBox;
    CContrat      : TCheckBox;
    CAppel        : TCheckBox;
    CCumulAffaire : TCheckBox;
    CSelSansConso : TCheckBox;
    CSelReglement : TCheckBox;
    CEclatNatPrest: TCheckBox;
    CEclatFamNiv1 : TCheckBox;
    CSELACHATSTOCK: TCheckBox;
    CHKPRORATA    : TCheckBox;
    CHKREVISION   : TCheckBox;
    CFactHorsDevis: TCheckBox;
    CTotauxAvoir  : TCheckBox;
    //
    THAffaire     : THEdit;
    THAffaire0    : THEdit;
    THAffaire1    : THEdit;
    THAffaire2    : THEdit;
    THAffaire3    : THEdit;
    THAvenant     : THEdit;
    //
    T_Affaire     : THLabel;
    //
    BtbApporteur  : THEdit;
    TBtBApporteur : THLabel;
    //
    MontantPA     : Double;
    MontantPR     : Double;
    MontantPV     : Double;
    Qte           : Double;
    //
    //---
    // Modified by f.vautrain 29/11/2017 16:29:31 -
    Prorata       : String;
    Revision      : String;
    //
    // ---
    //Modif FV : Dev. prioritaire DSL le 05/06/2012
    CoefFG_Param  : Double;
    TauxHoraire   : Double;
    // -- Partage de referentiel
    NomTabNatPrest: string;

    //Modif FV : Ajout pour gestion du modèle d'état...
    OptionEdition : TOptionEdition;
    //
    TheType       : String;
    BParamEtat    : TToolBarButton97;
    //
    ChkApercu     : TCheckBox;
    ChkReduire    : TCheckBox;
    //
    FETAT         : THValComboBox;
    TEtat         : ThLabel;
    //
    Pages         : TPageControl;
    //
    procedure OnDblClickTV(Sender: TObject);

    procedure ChargelesFacturesansCommandes(TOBTMP : TOB; CodeAffaire: String);
    procedure ChargeLesChantiers;
    procedure ChargelesCommandesLivrees(TOBTMP  : TOB; CodeAffaire: String);
    procedure ChargelesCommandesNonLivrees(TOBTMP  : TOB; CodeAffaire: String);
    procedure ChargeLesConsos;
    procedure ChargeProrata(TOBTMP: TOB; CodeAffaire: String);
    procedure ChargeRevision(TOBTMP: TOB; CodeAffaire: String);

    procedure AjouteChampSup(var TOBTMP: TOB);
    procedure CompletelesChantiers;
    procedure InsereChampsDesire(TOBL, TOBAFF: TOB);

    procedure ChargeLeReste(TOBTMP: TOB);
    procedure ChargeLeResteADepenser(TOBTMP, TOBEclate: TOB; CodeAffaire : String);
    //
    //function AvanceVide(TOBL: TOB): boolean;
    //function EngageVide(TOBL: TOB): boolean;
    //function FactureVide(TOBL: TOB): boolean;
    //function LigneVide(TOBL: TOB): boolean;
    //function PrevueVide(TOBL: TOB): boolean;
    //function RealiseVide(TOBL: TOB): boolean;
    //
    function ChargeAffaireRef(CodeAffaire, AffaireSaisie: String): String;
    //
    //procedure Repartition_Eclatement(Eclatement, TypeMt: String; Montant: Double; TOBTMP: Tob);
    procedure Repartition_Fournisseur(FamNiveau1: String; TOBTMP: TOB);
    procedure Repartition_Frais(TypeRessource, FamNiveau1 : String; TobTMP : TOB);
    procedure Repartition_IntSsTraitLocAut(TypeRessource, NatPrestation : String; TobTMP : TOB);
    procedure Repartition_MoInterne(TypeRessource, NatPrestation: String; TOBTMP: Tob);
    procedure Repartition_MatOutil(TypeRessource, NatPrestation: String; TobTMP: TOB);
    procedure Traitement_Repartition(TOBEclate, TOBTMP : TOB);

		//Procedure MAJ_PremiereLigne(NomChampSup : String);

    // Modif FV puis LS et enfin le pape
    procedure CChoixClick(Sender: TObject);
    procedure AppelAnalyseDetail;
    procedure GenerationExcel(TOBBTB : TOB);
    procedure ConstitutionFeuilleXLS(WorkBook: Olevariant; NomFeuille: String; TobExport: TOB);
    // --
    procedure CCumulAffaireClick(Sender: TObject);
    procedure DateTransfoOK (TOBEclate : TOB; var Montant : Double);
    procedure DateFactureOK (Nummouv : Double; var Montant : Double);

    //FV : intégration reste à dépenser dans tableau de bord ventilé par type de ressource
    procedure Traitement_Repartition_RAD(TOBL, TOBTMP: TOB);
    procedure Repartition_MoInterne_RAD(TypeRessource,NatPrestation: String; TOBTMP: Tob; MTReste, QTReste: Double);
    procedure Repartition_Frais_RAD(TypeRessource, FamNiveau1: String; TobTMP: TOB; MTReste, QTReste: Double);
    procedure Repartition_MatOutil_RAD(TypeRessource, NatPrestation: String; TobTMP: TOB; MtReste, QtReste: Double);
    procedure Repartition_IntSsTraitLocAut_RAD(TypeRessource, NatPrestation: String; TobTMP: TOB; MtReste, QTReste: Double);
    procedure Repartition_Facturation_RAD(TypeRessource, NatPrestation: String; TobTMP: TOB; MtReste, QtReste: Double);
    Procedure Repartition_Fourniture_RAD(TypeRessource, NatPrestation : String; TobTMP : TOB; MtReste, QtReste : Double);
    procedure ChargeAchatStocks(TOBTMP : TOB; CodeAffaire: String);
    procedure GenerationTotal;
    procedure ChargeTOBTotalLigneTB(TOBTOTGEN, TOBGEN: TOB; Nomchamp: String);
    procedure ChargeTOBTotalEnteteTB(TOBTOTGEN, TOBGEN: TOB);
    procedure ChargeInfosTiers(TOBTEMP: TOB);
    procedure ChargeInfoResponsable(TOBTEMP: TOB);
    procedure ChargeInfoLibre(TOBTEMP: TOB; Index: String);
    procedure choixModeleExcel(TOBBTB: TOB; Totalisation : Boolean = False);
    procedure AppelJournalConso;
    procedure ChargeAchatFrais(TOBTMP: TOB; CodeAffaire: String);
    procedure TraitementColonnePrevu(Nature: string);
    procedure OnChangeFEtat(Sender: TObject);
    procedure BParamEtatClick(Sender: TObject);
    procedure ChargeEtatTableauBord;
    procedure ExitResponsable(Sender: TObject);
    procedure ExitDomaine(Sender: TObject);
    procedure ExitEtablissement(Sender: TObject);
    procedure AppelListeDesCdes;
    procedure CloseTTForm(Sender: TOBJECT);
    procedure ChargeHorsDevis(TOBTMP: TOB; CodeAffaire: String);
    procedure ChargeTotalAvoir(TOBTMP: TOB; CodeAffaire: String);

  end;

implementation
uses Paramsoc,
     Dialogs,
     UtilXlsBTP,
     UtilRessource,
     Variants,
     FactComm,
     EntGC,
     FactTob;


procedure TOF_BTTABLEAUBORD.OnNew;
begin
  inherited;
end;

procedure TOF_BTTABLEAUBORD.OnDelete;
begin
  inherited;
end;

procedure TOF_BTTABLEAUBORD.OnUpdate;
begin
  inherited;

	CodeChantier := CodeAffaireRegroupe(CodeAffaire0, CodeAffaire1,CodeAffaire2,CodeAffaire3,CodeAvenant,TaModif,false,false,false);

  TOBBTB.cleardetail;

  ChargeLesChantiers;
  
  if CSELSANSCONSO.Checked then CompletelesChantiers;

  if TOBBTB <> nil then
  begin
    ChargeLesConsos;
    TFStat(Ecran).LaTOB := TOBBTB;
  end;

end;

procedure TOF_BTTABLEAUBORD.OnLoad;
begin
  inherited;
  Apporteur    := GetControltext('BTBAPPORTEUR');
  Responsable  := GetControltext('AFF_RESPONSABLE');
//    
  DateDeb := StrtoDate(GetControltext('BTBDATE'));
  DateFin := StrtoDate(GetControltext('BTBDATE_'));

  CodeAffaire0  := GetControltext('AFF_AFFAIRE0');
  CodeAffaire  := GetControltext('BTBAFFAIRE');
  CodeAffaire1 := GetControltext('BTBAFFAIRE1');
  CodeAffaire2 := GetControltext('BTBAFFAIRE2');
  CodeAffaire3 := GetControltext('BTBAFFAIRE3');
  CodeAvenant  := GetControltext('BTBAVENANT');

  CodeClient   := GetControltext('BTBTIERS');
  
  Etablissement:= GetControltext('GP_ETABLISSEMENT');
  
  // Ajout LS
  DateMVTDeb   := StrtoDate(GetControltext('DATEMVTDEB'));
  DateMVTFin   := StrtoDate(GetControltext('DATEMVTDEB_'));

  OptionChoixPrevuAvanc := [OptGlobal];

  if TcheckBox(GetControl('CDBT')).Checked then
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc + [OptDetailDev]
  else
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc - [OptDetailDev];

  if TcheckBox(GetControl('CPBT')).Checked then
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc + [OptDetailPBT]
  else
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc - [OptDetailPBT];

  if TcheckBox(GetControl('CETU')).Checked then
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc + [OptDetailETU]
  else
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc - [OptDetailETU];

  if TcheckBox(GetControl('CBCE')).Checked then
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc + [OptDetailBCE]
  else
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc - [OptDetailBCE];

  if TcheckBox(GetControl('CFAC')).Checked then
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc + [OptDetailFAC]
  else
     OptionChoixPrevuAvanc := OptionChoixPrevuAvanc - [OptDetailFAC];
  // --
end;

procedure TOF_BTTABLEAUBORD.OnArgument(S: string);
Var Req : String;
    QQ  : TQuery;
begin
  inherited;
  //
  TOBBTB := TOB.create('Ma Tob', nil, -1);

  // Gestion des évènements du TB
  TV1 := TTobViewer(GetControl('TV'));
  TV1.OnDblClick := OnDblClickTV;
  //
  if Assigned(GetControl('AFF_RESPONSABLE'))   Then
  begin
    TResponsable          := THEdit(GetControl('AFF_RESPONSABLE'));
    if Assigned(GetControl('LIBRESSOURCE'))      Then LibResponsable   := THLabel(Getcontrol('LIBRESSOURCE'));
    //TResponsable.OnElipsisClick := BRechResponsable;
    TResponsable.OnEXit   := ExitResponsable;
  end;

  If Assigned(GetControl('AFF_DOMAINE'))       then
  Begin
    TDomaine := THEdit(GetControl('AFF_DOMAINE'));
    if Assigned(GetControl('LIBDOMAINE'))        Then LibDomaine       := THLabel(Getcontrol('LIBDOMAINE'));
    TDomaine.OnEXit := ExitDomaine;
  end;

  If Assigned(GetControl('AFF_ETABLISSEMENT')) then
  begin
    TEtablissement        := THEdit(GetControl('AFF_ETABLISSEMENT'));
    if Assigned(GetControl('LIBETABLISSEMENT'))  Then LibEtablissement := THLabel(Getcontrol('LIBETABLISSEMENT'));
    TEtablissement.OnEXit := ExitEtablissement;
  end;

  //
  THAffaire  := THEdit(GetControl('BTBAFFAIRE'));
  //
  if ecran.name  = 'BTRESULTATS' then
    THAffaire0 := THEdit(GetControl('BTBAFFAIRE0'))
  else
    THAffaire0 := THEdit(GetControl('AFF_AFFAIRE0'));

  THAffaire1 := THEdit(GetControl('BTBAFFAIRE1'));
  THAffaire2 := THEdit(GetControl('BTBAFFAIRE2'));
  THAffaire3 := THEdit(GetControl('BTBAFFAIRE3'));
  THAvenant  := THEdit(GetControl('BTBAVENANT'));
  //
  if Assigned(GetControl('BTBAPPORTEUR'))  Then BtbApporteur  := THEdit(GetControl('BTBAPPORTEUR'));
  if Assigned(GetControl('TBTB_APPORTEUR')) Then TBtBApporteur := THLabel(GetControl('TBTB_APPORTEUR'));
  //
  CChantier := ThCheckBox(ecran.FindComponent('CCHANTIER'));
  CChantier.OnClick := CChoixClick;
  //
  CContrat  := ThCheckBox(ecran.FindComponent('CCONTRAT'));
  CContrat.OnClick := CChoixClick;
  //
  CAppel    := ThCheckBox(ecran.FindComponent('CAPPEL'));
  CAppel.OnClick := CChoixClick;
  //
  CSelSansConso  := ThCheckBox(ecran.FindComponent('CSELSANSCONSO'));
  CCumulAffaire  := ThCheckBox(ecran.FindComponent('CCUMAFFAIRE'));
  CCumulAffaire.OnClick := CCumulAffaireClick;
  CSelReglement  := ThCheckBox(ecran.FindComponent('CSELREGLEMENT'));
  //
  CEclatNatPrest := ThCheckBox(ecran.FindComponent('CECLATNATPREST'));
  CEclatFamNiv1  := ThCheckBox(ecran.FindComponent('CECLATFAMNIV1'));
  CSELACHATSTOCK := ThCheckBox(ecran.FindComponent('CSELACHATSTOCK'));
  //
  CHKPRORATA     := ThCheckBox(ecran.FindComponent('CHKPRORATA'));
  CHKREVISION    := ThCheckBox(ecran.FindComponent('CHKREVISION'));
  CFactHorsDevis := ThCheckBox(ecran.FindComponent('CHKFACTHORSDEVIS'));
  CTotauxAvoir   := ThCheckBox(ecran.FindComponent('CHKTOTAUXAVOIR'));
  //
  T_Affaire :=  THLabel (GetControl('TBTB_AFFAIRE'));
  //
  //FV1 : 09/03/2017 - Intégration gestion des modèles d'état dans le Tableau de bord
  ChargeEtatTableauBord;
  //
  //Vérification si des appels existent et si des contrats existent
  Req := 'Select AFF_AFFAIRE0 FROM AFFAIRE WHERE AFF_AFFAIRE0="I"';
  QQ := OpenSQL(req, true,-1,'',true);
  if QQ.eof then CContrat.visible := False;
  Ferme(QQ);

  Req := 'Select AFF_AFFAIRE0 FROM AFFAIRE WHERE AFF_AFFAIRE0="W"';
  QQ := OpenSQL(req, true,-1,'',true);
  if QQ.eof then CAppel.visible := False;
  ferme(QQ);

  if (CAppel.visible = false) and (CContrat.Visible = false) then
  begin
     SetControlproperty('GCHOIXAFFAIRE', 'Visible', False);
     SetControlproperty('CChantier', 'Visible', False);
  end;
  // --
  THAffaire0.text := 'A';

  if not (VH_GC.AFGestionCom) then
  begin
    BtbApporteur.visible  := False;
    TBtbApporteur.visible := False
  end
  else
  begin
    BtbApporteur.visible  := True;
    TBtbApporteur.visible := True;
  end;

//uniquement en line
{*  SetControlProperty('TOPTIONS', 'TabVisible', False);
  SetControlProperty('PCOMPLEMENT', 'TabVisible', False);
  SetControlProperty('PAVANCE', 'TabVisible', False);
  SetControlProperty('GCHOIXAFFAIRE', 'Visible', False);
*}

  //Modif FV : Dev. prioritaire DSL le 05/06/2012
  CoefFG_param  := GetParamSocSecur('SO_COEFFG', 0);
  Tauxhoraire   := GetParamSocSecur('SO_TAUXHORAIRE', 0);
  // -- Recup Info partage de referentiel
  NomTabNatPrest := GetShare('BTNATPRESTATION')+'NATUREPREST';
  // --
  Prorata       := GetParamSocSecur('SO_PRORATA',  '');
  Revision      := GetParamSocSecur('SO_REVISION', '');
  //
  Prorata   := FabriqueConditionIn(Prorata);
  Revision  := FabriqueConditionIn(Revision);
  //
  if Prorata  = '' then CHKPRORATA.visible := False;
  if Revision = '' then CHKREVISION.Visible := False;
  //
end;


Procedure TOF_BTTABLEAUBORD.ExitResponsable(Sender: TObject);
Var Lib1, Lib2  : string;
    Zmail       : String;
    Zmail1      : string;
    Zmail2      : string;
begin

  if TResponsable = nil then Exit;

  if LibResponsable = nil then Exit;

  LibResponsable.Caption  := '';
  THLabel(GetControl('ZMRESSOURCE')).caption := '';

  IF TResponsable.text <> '' then
  begin
    LibelleRessource(TResponsable.text, lib1, lib2);
    LibResponsable.Visible := True;
    LibResponsable.caption := Lib1 + ' ' + Lib2;
    //
    //FV1 : 09/01/2018 - FS#2840 - FLIPO - afficher l'adresse mail du responsable à l'édition du tableau de bord
    EmailRessource(TResponsable.text, Zmail);
    if pos('@', Zmail) <> 0 then
    begin
      Zmail1 := Copy(Zmail, 0, Pos('@', Zmail)-1);
      Zmail2 := Copy(Zmail, Pos('@', Zmail)+1, Length(Zmail) - Pos('@', Zmail)+1);
    end;
    THLabel(GetControl('ZMAIL1')).caption := Zmail1;
    THLabel(GetControl('ZMAIL2')).caption := Zmail2;
  end;


end;

Procedure TOF_BTTABLEAUBORD.ExitEtablissement(Sender: TObject);
begin

  if TEtablissement = nil then Exit;

  if LibEtablissement = nil then Exit;

  LibEtablissement.Caption := TEtablissement.text;

end;

Procedure TOF_BTTABLEAUBORD.ExitDomaine(Sender: TObject);
begin

  if TDomaine = nil then Exit;

  if LibDomaine = nil then Exit;

  LibDomaine.Caption := TDomaine.Text;

end;

procedure TOF_BTTABLEAUBORD.ChargeEtatTableauBord;
Var Idef : Integer; 
begin

  TheType   := 'E';

  BParamEtat     := TToolBarButton97(GetControl('BParamEtat'));
  ChkApercu      := TCheckBox(GetControl('FApercu'));
  ChkReduire     := TCheckBox(GetControl('FReduire'));
  //
  FETAT          := THValComboBox(GetControl('FEtat'));
  TEtat          := ThLabel(GetControl('TEtat'));
  //
  Pages          := TPageControl(GetControl('Pages'));
  //
  OptionEdition := TOptionEdition.Create(TheType, TFStat(Ecran).NatureEtat, FEtat.Value, Ecran.Caption, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, Pages, fEtat);

  OptionEdition.first := True;
  OptionEdition.ChargeListeEtat(fEtat, Idef);
  //
  TFStat(Ecran).NatureEtat := 'BTB';

  FEtat.OnChange := OnChangeFEtat;
  FETAT.Plus     := 'MO_TYPE="E" AND MO_NATURE="' + TFStat(Ecran).NatureEtat + '"';
  FEtat.Value    := GetParamSocSecur('SO_BTTABLEAUBORD', '');

  BParamEtat.OnClick  := BParamEtatClick;
  BParamEtat.Visible  := BParamEtat.Visible and JaiLeDroitConcept(ccParamEtat,False);

end;


Procedure TOF_BTTABLEAUBORD.OnChangeFEtat(Sender : TObject);
begin

  TFStat(Ecran).CodeEtat  := FEtat.Value;
  OptionEdition.Modele    := FEtat.Value;

end;

procedure TOF_BTTABLEAUBORD.BParamEtatClick(Sender: TObject);
begin

  OptionEdition.Appel_Generateur
  
end;

procedure TOF_BTTABLEAUBORD.OnClose;
begin
  inherited;
  FreeAndNil(TOBBTB);
  FreeAndNil(OptionEdition);
end;

//******************************************************************************
//********************* Zooms / double clic  ***********************************
//******************************************************************************

//FV1 - gestion de la TOB dans EXCEL
procedure TOF_BTTABLEAUBORD.OnDblClickTV(Sender: TObject);
Var CBtnTxt         : PCustBtnText;
Begin

  fModeleAnaXls := '';

  system.New(CBtnTxt);

  if TV1.ColName[TV1.CurrentCol] = 'AFF_LIBELLE' then
  begin
    CBtnTxt[mbCust1] := 'Analyse détaillée';
    CBtnTxt[mbCust3] := 'Génération Excel avec Choix modèle';
    CBtnTxt[mbCust4] := 'Génération Excel';

    Case ExMessageDlg('Que voulez-vous faire ?' + #13#10, mtxInformation, [mbCust1,mbCust4,mbCust3], 0, mbCust2, Application.Icon, CBtnTxt) Of
      mrCust1: AppelAnalyseDetail;
      mrCust3: ChoixModeleExcel(TOBBTB);
      mrCust4: GenerationExcel(TOBBTB);
    End;
    dispose(CBtnTxt);
  end
  else if (Pos('CONSOMME', TV1.ColName[TV1.CurrentCol])<>0) Or (Pos('REALISE', TV1.ColName[TV1.CurrentCol])<>0) then
  begin
    AppelJournalConso;
  end
  else
  begin
    if (tv1.CurrentCol=0) And (tv1.Currentrow < 0) then
    begin
      CBtnTxt[mbCust1] := 'Export Totalisations Tableau';
      CBtnTxt[mbCust3] := 'Génération Excel avec Choix modèle';
      CBtnTxt[mbCust4] := 'Annuler';
      //
      Case ExMessageDlg('Que voulez-vous faire ?' + #13#10, mtxInformation, [mbCust1,mbCust4, mbCust3], 0, mbCust2, Application.Icon, CBtnTxt) Of
        mrCust1: GenerationTotal;
        mrCust3: ChoixModeleExcel(TOBBTB,True);
        mrCust4: Exit;
      End;
      dispose(CBtnTxt);
    end
    else
      AppelAnalyseDetail;
  end;

End;

Procedure TOF_BTTABLEAUBORD.AppelListeDesCdes;
Var TobCde    : TOB;
    ToblCde   : TOB;
    TobTempo  : TOB;
    Tobl      : TOB;
    ToblTempo : TOB;
    II        : Integer;
    StSQL     : string;
    QQ        : TQuery;
    Arg       : string;
Begin

  CodeAffaire := TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];

  if CodeAffaire = '' then
  begin
    PgiError ('Le code Affaire n''est pas renseigné', application.name);
    SourisNormale;
    exit;
  end;

  TobCde := Tob.create('CDELIVREE', nil, -1);
  TobCde.ClearDetail;

  //Arg := 'GP_NATUREPIECEG=CF;GP_AFFAIRE=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
  //Arg := Arg + ';GP_AFFAIRE1=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE1'), TV1.CurrentRow];
  //Arg := Arg + ';GP_AFFAIRE2=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE2'), TV1.CurrentRow];
  //Arg := Arg + ';GP_AFFAIRE3=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE3'), TV1.CurrentRow];
  //Arg := Arg + ';GP_DATEPIECE=' +

  //AGLLanceFiche('BTP', 'BTPIECECDE', Arg + ';CONSULTATION=X', 'AND GL_QTERESTE = 0', ''); // affichage factures

  //Création de la fiche

  TTForm         := TToolwindow97.Create(nil);
  TTForm.Parent  := Ecran;
  TTForm.Width   := Round(Ecran.Width);
  TTForm.Height  := Round(Ecran.Height/2);
  TTForm.Top     := Round((Ecran.Height/2) - (TTForm.Height/2));
  TTForm.Left    := Round((ecran.Width/2) - (TTForm.Width/2));
  TTForm.OnClose := CloseTTForm;
  //
  //Création de la Grille
  TTGrille       := THGrid.Create(TTForm);
  TTGrille.parent:= TTForm;
  TTGrille.Width := Round(TTForm.Width -20);
  TTGrille.Height:= Round(TTForm.Height-20);
  TTGrille.Top   := 0;
  TTGrille.Left  := 0;
  TTGrille.Align := alClient;
  TTGrille.BorderStyle := bsNone;
  TTGrille.Ctl3D := False;

  //
  //Initiualisation et Dessin de la grille
  GridEvent      := TGestionGS.Create;
  GridEvent.Ecran:= Ecran;
  GridEvent.GS   := TTGrille;
  GridEvent.TOBG := TobCde;

  //chargement de la grille de présentation
  GridEvent.NomListe  := 'BTPIECECDE';
  GridEvent.ChargeInfoListe;

  //Dessin de la grille
  if GridEvent.NomListe <> '' then GridEvent.DessineGrille;

  //
  StSQL := 'SELECT * FROM BTDEVIS LEFT JOIN LIGNE ';
  StSQL := StSQL + '   ON GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG ';
  StSQL := StSQL + 'WHERE GL_NATUREPIECEG = "CF" ';

  //
  if TV1.ColName[TV1.CurrentCol] = 'CDELIVREE' then
  begin
    TTForm.Caption := 'Commandes livrées';

    StSQL := StSQL + '   AND GL_QTERESTE = 0';
  end
  else
  begin
    TTForm.Caption := 'Commandes Non livrées';

    StSQL := StSQL + '   AND GL_QTERESTE<>0';
  end;

  StSQL := StSQL + '   AND GL_AFFAIRE = "' + CodeAffaire + '" ';

  QQ := OpenSQL(StSQL, true,-1,'',true);

  //chargement de la Tob en fonction de la requête initale
  If Not QQ.eof then
  begin
    TobTempo := Tob.create('COMMANDES', nil, -1);
    TobTempo.LoadDetailDB('PIECE','','',QQ, False);
    For II := 0 to TobTempo.detail.count -1 do
    begin
      TOBL := TobCde.FindFirst(['GP_NUMERO'],[TobTempo.detail[ii].GetValue('GP_NUMERO')], false);
      If TOBL = nil then
      begin
        TOBLCde := Tob.Create('PIECE', TobCde, -1);
        TobLCde.Dupliquer(TobTempo.detail[ii],True,True);
      end;
    end;
    FreeAndNil(TobTempo);
  end;

  Ferme(QQ);

  TTForm.Visible := True;
  TTGrille.visible := True;

  if TobCde.Detail.count >0 then
  begin
    GridEvent.ChargementGrille;
  end
  else
  begin
    CloseTTForm(Self);
  end;
  //
  FreeAndNil(TobCde);
  //
end;

procedure TOF_BTTABLEAUBORD.CloseTTForm(Sender : TOBJECT);
begin

  TTform.Visible := False;

  FreeAndNil(TTForm);

  FreeAndNil(GridEvent);


end;

Procedure TOF_BTTABLEAUBORD.AppelJournalConso;
Var NatPrestation : String;
    TypeRessource : String;
    TOBL          : TOB;
    LastAutoSearch: boolean;
    LibEntCol     : String;
begin

  //if CodeAffaire = '' then
  CodeAffaire := TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];

  if CodeAffaire = '' then
  begin
    PgiError ('Le code Affaire n''est pas renseigné', application.name);
    SourisNormale;
    exit;
  end;

  TOBL := TOBBTB.Detail[TV1.CurrentRow];


  StArgument := 'AFFAIRE=' + CodeAffaire;
  StArgument := StArgument + ';DATEDEB=' + DateToStr(DateMvtDeb) + ';DATEFIN=' + DatetoStr(DateMvtfin);

  LibEntCol := TV1.ColName[TV1.CurrentCol];

  //Découpage de l'entete de colonne pour récupération Nature Prestation et Type de ressource
  if LibEntCol = 'REALISEMOPA' then
    StArgument := StArgument + ';NATUREMOUV=MO'
  else
  begin
  if CEclatNatPrest.checked then
  begin
    //
    NatPrestation := TOBL.GetString('NATUREPRESTATION');
    TypeRessource := TOBL.GetString('ARS_TYPERESSOURCE');
    //
    if TypeRessource = '' then TypeRessource := TOBL.GetString('BCO_TYPERESSOURCE');
    if (TypeRessource = 'INT') then StArgument := StArgument + ';AVANCE';
    if      (TypeRessource = 'ACH') OR (TypeRessource = 'STK')                            then StArgument := StArgument + ';NATUREMOUV=FOU'
    Else if (TypeRessource = 'AUT') OR (TypeRessource = 'LOC') OR (TypeRessource = 'ST')  then StArgument := StArgument + ';NATUREMOUV=EXT'
    Else if (TypeRessource = 'MAT') OR (TypeRessource = 'OUT')                            then StArgument := StArgument + ';NATUREMOUV=RES'
    Else if  TypeRessource = 'FRA'                                                        then StArgument := StArgument + ';NATUREMOUV=FRS'
    Else if  TypeRessource = 'SAL'                                                        then StArgument := StArgument + ';NATUREMOUV=MO';
  end;
  end;

  StArgument := StArgument + ';FULLSCREEN';

  LastAutoSearch := V_PGI.AutoSearch;

  V_PGI.AutoSearch := true;

  AGLLanceFiche('BTP','BTJOUCON','','',StArgument) ;

  V_PGI.AutoSearch := lastautoSearch;

end;

//Cumul pour gestion des totaux dans excel
procedure TOF_BTTABLEAUBORD.GenerationTotal;
var i         : Integer;
    Indice    : integer;
    //
		NomChamp  : string;
    //
    TOBGEN    : TOB;
    TOBTOTGEN : TOB;
begin

  TOBTOTGEN := TOB.Create('TOTGENTB', nil, -1);

  if assigned(TOBBTB) then
  begin
    //chargement de la Tob de totalisation général
    For Indice := 0 to TobBTB.detail.count -1 do
    begin
      i := 1000;
      TOBGEN := TobBTB.detail[Indice];
      NomChamp := TOBGEN.GetNomChamp(i);
      TOBTOTGEN.AddChampSupValeur('OKTIERS', 'X');
      if (VH_GC.AFGestionCom) then
        TOBTOTGEN.AddChampSupValeur('OKAPPORTEUR', 'X')
      else
        TOBTOTGEN.AddChampSupValeur('OKAPPORTEUR', '-');
      //
      TOBTOTGEN.AddChampSupValeur('OKRESPNSABLE', 'X');
      //
      while NomChamp <> '' do
      begin
        //
        if not TOBTOTGEN.FieldExists (NomChamp) then TOBTOTGEN.AddchampSupValeur(NomChamp, 0);
        //
        if (Pos('AFF_TIERS', Nomchamp)<>0) then
        begin
          if TOBGEN.GetString('AFF_TIERS') <> CodeClient then
          begin
            TOBTOTGEN.PutValue('OKTIERS', '-');
          end;
        end;
        //
        if BTBApporteur <> nil then
        begin
        if (Pos('AFF_APPORTEUR', Nomchamp)<>0) then
        begin
          if TOBGEN.GetString('AFF_APPORTEUR') <> Apporteur then
          begin
            TOBTOTGEN.PutValue('OKAPPORTEUR', '-');
          end;
        end;
        end;
        //
        if TResponsable <> nil then
        Begin
        if (Pos('AFF_RESPONSABLE', Nomchamp)<>0) then
        begin
          if TOBGEN.GetString('AFF_RESPONSABLE') <> Responsable then
          begin
            TOBTOTGEN.PutValue('OKRESPONSABLE', '-');
          end;
        end;
        end;
        //
        if (Pos('REALISE'  , NomChamp)<>0) OR
           (pos('RESTEADEP', NomChamp)<>0) OR
           (Pos('ENGAGE'   , NomChamp)<>0) OR
           (Pos('AVANCE'   , NomChamp)<>0) OR
           (Pos('PREVU'    , NomChamp)<>0) OR
           (Pos('FACT'     , NomChamp)<>0) OR
           (Pos('ANNEXE'   , NomChamp)<>0) OR
           (Pos('REGLE'    , NomChamp)<>0) OR
           (Pos('TPS'      , NomChamp)<>0) OR
           (Pos('CONSOMME' , NomChamp)<>0) Then
        begin
          //Chargement des lignes du Tobviewer
          ChargeTOBTotalLigneTB(TOBTOTGEN, TOBGEN, NomChamp);
        end
        Else
        begin
          //Chargement des zones d'entêtes
          ChargeTOBTotalEnteteTB(TOBTOTGEN, TOBGEN);
        end;
        inc(i);
        NomChamp := TOBGEN.GetNomChamp(i);
      end;
    end;
  end;

  GenerationExcel(TOBTOTGEN);;

  FreeAndNil(TobTotGen);

end;

//chargement de l'entête
Procedure TOF_BTTABLEAUBORD.ChargeTOBTotalEnteteTB(TOBTOTGEN, TOBGEN : TOB);
Begin

  if TOBTOTGEN.GetString('OKTIERS') = 'X' then
  begin
    TOBTOTGEN.PutValue('AFF_TIERS', CodeClient);
    ChargeInfosTiers(TOBTOTGEN);
  end
  Else
    TOBTOTGEN.PutValue('AFF_TIERS', '');
  //
  if BtbApporteur <> nil then
  begin
  If TOBTOTGEN.GetString('OKAPPORTEUR') = 'X' then
    TOBTOTGEN.PutValue('AFF_APPORTEUR', Apporteur)
  else
    TOBTOTGEN.PutValue('AFF_APPORTEUR', '');
  end;
  //
  if TResponsable <> nil then
  begin
  If TOBTOTGEN.GetString('OKRESPONSABLE') = 'X' then
  begin
    TOBTOTGEN.PutValue('AFF_RESPONSABLE', Responsable);
    ChargeInfoResponsable(TOBTOTGEN);
  end
  else
    TOBTOTGEN.AddChampSupValeur('NOMRESPONSABLE', '');
  end;
  //
  ChargeInfoLibre(TOBTOTGEN, '1');
  ChargeInfoLibre(TOBTOTGEN, '2');
  ChargeInfoLibre(TOBTOTGEN, '3');
  ChargeInfoLibre(TOBTOTGEN, '4');
  ChargeInfoLibre(TOBTOTGEN, '5');
  ChargeInfoLibre(TOBTOTGEN, '6');
  ChargeInfoLibre(TOBTOTGEN, '7');
  ChargeInfoLibre(TOBTOTGEN, '8');
  ChargeInfoLibre(TOBTOTGEN, '9');
  ChargeInfoLibre(TOBTOTGEN, 'A');

  TOBTOTGEN.PutValue('AFF_LIBELLE', 'TOTALISATIONS TABLEAU DE BORD');

  TOBTOTGEN.PutValue('AFF_DATEDEBUT', TOBGEN.GetDateTime('AFF_DATEDEBUT'));
  TOBTOTGEN.PutValue('AFF_DATESIGNE', TOBGEN.GetDateTime('AFF_DATESIGNE'));
  TOBTOTGEN.PutValue('AFF_DATEFIN',   TOBGEN.GetDateTime('AFF_DATEFIN'));

  TOBTOTGEN.PutValue('CECLATNATPREST',TOBGEN.GetDateTime('CECLATNATPREST'));

  TOBTOTGEN.PutValue('DATEMVTDEB',    TOBGEN.GetDateTime('DATEMVTDEB'));
  TOBTOTGEN.PutValue('DATEMVTFIN',    TOBGEN.GetDateTime('DATEMVTFIN'));

end;

procedure TOF_BTTABLEAUBORD.ChargeTOBTotalLigneTB(TOBTOTGEN, TOBGEN : TOB; Nomchamp : String);
var PrefixeZone : String;
    SuffixeZone : string;
    TotGeneral  : Double;
    MtSection   : Double;
begin
  //
  PrefixeZone := LeftSTR(NomChamp,Pos('_', Nomchamp));
  SuffixeZone := RightSTR(NomChamp,(Length(Nomchamp) - Length(PrefixeZone)));
  //
  MtSection   := TOBGEN.GetDOUBLE(PrefixeZone + SuffixeZone);
  //
  If MtSection = 0 then exit;
  //
  TotGeneral  := TOBTOTGEN.GetDOUBLE(PrefixeZone + SuffixeZone);
  TotGeneral  := TotGeneral + MtSection;
  TOBTOTGEN.PutValue(PrefixeZone + SuffixeZone, TotGeneral);

end;

Procedure TOF_BTTABLEAUBORD.ChoixModeleExcel(TOBBTB : TOB; Totalisation : Boolean = False);
var TT : TOpenDialog;
    Rep: String;
begin

  //ouverture du sélecteur de fichier windows dans le répertoire des modèles...
	TT := TOpenDialog.Create(Ecran);
  TRY
    TT.DefaultExt := '.xlsx';
    TT.Filter := 'Modèle Excel (*.xls)|*.xls;*.xlsx';
    Rep := GetParamSocSecur('SO_BTMODELEANAXLS','');
    Rep := ExtractFilePath(rep);
    TT.InitialDir := rep;
    if TT.Execute then
    begin
      fModeleanaxls := TT.FileName;
    end;
  FINALLY
  	TT.Free;
  end;

  if totalisation then
    GenerationTotal
  else
    GenerationExcel(TOBBTB);

end;

procedure TOF_BTTABLEAUBORD.GenerationExcel(TOBBTB : TOB);
var FileNameXLS   : String;
    Repert        : String;
    NomFeuille    : String;
    NomLot        : String;
    CodeAffaire   : String;
    //
    MotDePasse    : String;
    //
    fWinExcel     : OleVariant;
    fWorkBook     : Variant;
    //
    fnewInst      : Boolean;
    //
    TOBT          : TOB;
begin

  // transfert de prévision vers excel pour récupération dans visual projet
  SourisSablier;

  MotDepasse := GetParamSocSecur('SO_PASSWDXLS', 'PASSWD');

  //chargement du document excel
  Repert := getparamSocSecur('SO_BTSTOCKAGEEXPORT','');
  if Repert = '' then
  begin
    PgiInfo ('veuillez renseigner l''emplacement de stockage des exports', application.name);
    SourisNormale;
    exit;
  end;

  if fModeleanaxls = '' then
  begin
    fModeleAnaXls := GetparamSocSecur('SO_BTMODELEANAXLS','C:\PGI00\STD\MODELEANAXLS.xlsx');
    if fModeleAnaXls = '' then
    begin
      PGIError('Modèle d''export non défini', Application.Name);
      SourisNormale;
      Exit;
    end;
  end;

  if not FileExists(fModeleAnaXls) then
  begin
    PGIError('Le Modèle d''export n''existe pas dans le répertoire paramètré', Application.Name);
    SourisNormale;
    Exit;
  end;

  if not OpenExcel(true,fWinExcel,fNewInst) then
  begin
    PgiError ('Excel n''est pas installé sur ce poste', application.name);
    SourisNormale;
    exit;
  end;

  if TV1.ColName[TV1.CurrentCol] = 'AFF_LIBELLE' then
  Begin
    CodeAffaire := TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];

    if codeAffaire = '' then
    begin
      PgiError ('Le code Affaire n''est pas renseigné', application.name);
      SourisNormale;
      exit;
    end;

    FileNameXLS := 'Analyse chantier ' + BTPCodeAffaireAffiche(codeAffaire) + '.xlsx';
  end
  else
  begin
    FileNameXLS := 'Analyse chantier ' + TOBBTB.GEtString('DATEDEB') + '_' +  TOBBTB.GEtString('DATEFIN') + '.xlsx';
  end;

  //Positionnement sur la feuille 1
  fWorkBook := OpenWorkBook(fModeleAnaXls, fWinExcel);

  NomFeuille := CopyFeuille(fWorkBook,'Feuil1',mcbefore);
  NomLot := 'Analyse '+ BTPCodeAffaireAffiche(CodeAffaire);
  RenommeFeuille (fWorkBook,NomFeuille,NomLot);

  if assigned(TOBBTB) then
  begin
    if TV1.ColName[TV1.CurrentCol] = 'AFF_LIBELLE' then
      TOBT := TOBBTB.FindFirst(['BCO_AFFAIRE'], [CodeAffaire], false)
    else
      TOBT := TOBBTB;
    //
    ConstitutionFeuilleXLS(fWorkBook, NomLot, TobT);
    //
    if not VarIsEmpty(fWinExcel) then
    begin
      //Protection du Document par mot de passe
      if MotDePasse <> '' then fWorkBook.activeSheet.Protect(MotDePasse, True, True);
      //Sauvegarde du document d'export
      ExcelSave (fWorkBook ,IncludeTrailingBackslash (Repert)+FileNameXls);
      if PGIask('Voulez-vous visualiser le fichier : ' + FileNameXls)= MrYes then
      Begin
        //Ouverture d'Excel            //
        fWinExcel.visible := true;
      end
      else
      begin
        //Fermeture d'excel et libération de la mémoire
        fWinExcel.quit;
        fWinExcel := unassigned;
      end;
    end;
  end;

  SourisNormale ;

end;

Procedure TOF_BTTABLEAUBORD.ConstitutionFeuilleXLS(WorkBook : Olevariant; NomFeuille : String; TobExport : TOB);
Var fWorkSheet  : OleVariant;
    //
    i           : Integer;
    //
    Etiquette   : String;
    ValeurChar  : String;
    ValeurNum   : Double;
begin

  SourisSablier;

  if not assigned(TOBExport) then
  begin
    PgiError ('Aucune ligne à exporter', Application.name);
    exit;
  end;

  if Not varIsEmpty(WorkBook) then
  begin
    fWorkSheet := SelectSheet (WorkBook,NomFeuille);
    fWorkSheet := WorkBook.activeSheet;
    //
    For i := 1 to WorkBook.Names.Count do
    begin
      Etiquette := Workbook.Names.Item(i).Name;
      ValeurChar := TobExport.GetValue(Etiquette);
      if Etiquette = 'BCO_AFFAIRE' then
        ExcelRangeValue(WorkBook,Nomfeuille,Etiquette,BTPCodeAffaireAffiche(ValeurChar))
      else if IsNumeric(ValeurChar) then
      begin
       ValeurNum   := StrToFloat(StringReplace(TobExport.GetString(Etiquette),'.',',',[rfReplaceAll]));
       ExcelRangeValue(WorkBook,Nomfeuille,Etiquette,ValeurNum);
      end
      else
        ExcelRangeValue(WorkBook,Nomfeuille,Etiquette,ValeurChar);
    end;

  end;

end;

procedure TOF_BTTABLEAUBORD.AppelAnalyseDetail;
var Req, Arg, Nature: string;
    QQ              : Tquery;
    OneTOb          : TOB;
    lastautoSearch  : boolean;
begin

  if Copy(TV1.ColName[TV1.CurrentCol], 1, 5) = 'PREVU' then
     begin
     if Copy(TV1.ColName[TV1.CurrentCol], 1, 6) = 'PREVU_' then
      Nature := Copy(TV1.ColName[TV1.CurrentCol], 7, 3)
     else
      Nature := 'PBT';
     //
     TraitementcolonnePrevu(Nature);
     end
  else if Copy(TV1.ColName[TV1.CurrentCol], 1, 6) = 'AVANCE' then
     begin
     Req := 'SELECT GP_NATUREPIECEG, GP_DATEPIECE, GP_SOUCHE, GP_NUMERO, GP_INDICEG, GP_TIERS, GP_AFFAIRE ' +
            'FROM PIECE ' +
            'WHERE GP_AFFAIRE="' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow] + '" ' +
            'AND GP_NATUREPIECEG="PBT"';
     QQ := OpenSQL(Req, true,-1,'',true);
     if not QQ.EOF then
        begin
        SaisieAvancementChantier(QQ.Fields[0].AsString,QQ.Fields[1].AsString,QQ.Fields[2].AsString,
                                 QQ.Fields[5].AsString, QQ.Fields[6].AsString,
                                 QQ.Fields[3].AsInteger, QQ.Fields[4].AsInteger, TaModif);
        end;
     ferme(QQ);
     end
  else if Copy(TV1.ColName[TV1.CurrentCol], 1, 3) = 'CDE' then
  Begin
    AppelListeDesCdes;
  end
  else if Copy(TV1.ColName[TV1.CurrentCol], 1, 7) = 'FACTURE' then
     begin
     Arg := 'GP_NATUREPIECEG=ABT:FBT:AVC:FAC:FBC:B00;GP_AFFAIRE=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
     Arg := Arg + ';GP_AFFAIRE1=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE1'), TV1.CurrentRow];
     Arg := Arg + ';GP_AFFAIRE2=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE2'), TV1.CurrentRow];
     Arg := Arg + ';GP_AFFAIRE3=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE3'), TV1.CurrentRow];
     AGLLanceFiche('BTP', 'BTPIECESAFF_MUL', Arg + ';CONSULTATION=X', '', ''); // affichage factures
     end
  else if Copy(TV1.ColName[TV1.CurrentCol], 1, 5) = 'REGLE' then
     begin
     Arg := 'GP_NATUREPIECEG=DAC:FBT:FBC:B00;GP_AFFAIRE=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
     Arg := Arg + ';GP_AFFAIRE1=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE1'), TV1.CurrentRow];
     Arg := Arg + ';GP_AFFAIRE2=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE2'), TV1.CurrentRow];
     Arg := Arg + ';GP_AFFAIRE3=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE3'), TV1.CurrentRow];
     AGLLanceFiche('BTP', 'BTREGLAFFAIRE', Arg + ';CONSULTATION=X', '', ''); // affichage reglements
     end
  else if Copy(TV1.ColName[TV1.CurrentCol], 1, 7) = 'REALISE' then
     begin
     Arg := 'BCOAFFAIRE=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
     Arg := Arg + ';BCOAFFAIRE1=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE1'), TV1.CurrentRow];
     Arg := Arg + ';BCOAFFAIRE2=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE2'), TV1.CurrentRow];
     Arg := Arg + ';BCOAFFAIRE3=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE3'), TV1.CurrentRow];
     AGLLanceFiche('BTP','BTCONSOS',Arg,'','AUTO;'+TV1.ColName[TV1.CurrentCol]) ;
     end
  else if (TV1.ColName[TV1.CurrentCol] = 'ENGAGEFOU') or (TV1.ColName[TV1.CurrentCol] = 'ENGAGECHA') or (TV1.ColName[TV1.CurrentCol] = 'MTRESTE') then
     begin
     Arg := 'CF:BFA';
     if (TV1.ColName[TV1.CurrentCol] = 'ENGAGECHA') then Arg := Arg + ':BLF:FF:AF:AFS';
     Arg := 'GP_NATUREPIECEG=' + Arg + ';GP_AFFAIRE=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
     AGLLanceFiche('BTP', 'BTPIECESAFF_MUL', Arg + ';CONSULTATION=X', '', ''); // affichage commandes et réceptions
     end
  else if TV1.ColName[TV1.CurrentCol] = 'FACTNONPAR' then
     begin
     Arg := 'CF:BFA:BLF:FF:AF:AFS';
     Arg := 'GP_NATUREPIECEG=' + Arg + ';GP_AFFAIRE=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
     AGLLanceFiche('BTP', 'BTPIECESAFF_MUL', Arg + ';CONSULTATION=X', '', ''); // affichage commandes et réceptions
     end
  else if TV1.ColName[TV1.CurrentCol] = 'BCO_AFFAIRE' then
     begin
     Arg := 'BCPAFFAIRE=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
     Arg := Arg + ';BCPAFFAIRE1=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE1'), TV1.CurrentRow];
     Arg := Arg + ';BCPAFFAIRE2=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE2'), TV1.CurrentRow];
     Arg := Arg + ';BCPAFFAIRE3=' + TV1.AsString[TV1.ColIndex('BCO_AFFAIRE3'), TV1.CurrentRow];
     AGLLanceFiche('BTP','BTCOMPARATIF',Arg,'','AUTO') ;
     end
  else if TV1.ColName[TV1.CurrentCol] = 'AFF_LIBELLE' then
     begin
     // ESSAIS ANALYSE GRAPHIQUE
     //TheTOB:=TOBBTB;
     //AGLLanceFiche('BTP','BTBRL','','','') ;
     //---------
//     V_PGI.DispatchTT(5, taConsult, TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow], '', '');

       lastautoSearch := V_PGI.AutoSearch;
			 V_PGI.AutoSearch := true;
			 OneTOB := TOB.Create ('LE LIEN',nil,-1);
       ONETOB.AddChampSupValeur ('DATEMVTDEB',DateMvtDeb);
       ONETOB.AddChampSupValeur ('DATEMVTFIN',DateMvtFin);
       ONETOB.AddChampSupValeur ('CHANTIER',TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow]);
       ONETOB.AddChampSupValeur ('CHANTIER1',TV1.AsString[TV1.ColIndex('BCO_AFFAIRE1'), TV1.CurrentRow]);
       ONETOB.AddChampSupValeur ('CHANTIER2',TV1.AsString[TV1.ColIndex('BCO_AFFAIRE2'), TV1.CurrentRow]);
       ONETOB.AddChampSupValeur ('CHANTIER3',TV1.AsString[TV1.ColIndex('BCO_AFFAIRE3'), TV1.CurrentRow]);
       ONETOB.AddChampSupValeur ('LIBELLECHANTIER',TV1.AsString[TV1.ColIndex('AFF_LIBELLE'), TV1.CurrentRow]);
       TheTOB := OneTOB;
       AglLanceFiche('BTP','BTRESULTCHADET','','','');
       TheTOB := nil;
       OneTOB.free;
			 V_PGI.AutoSearch := lastautoSearch;
     end
  else if (TV1.ColName[TV1.CurrentCol] = 'AFF_APPORTEUR') and (TV1.AsString[TV1.ColIndex('AFF_APPORTEUR'), TV1.CurrentRow] <> '') then
     begin
     V_PGI.DispatchTT(14, taConsult, TV1.AsString[TV1.ColIndex('AFF_APPORTEUR'), TV1.CurrentRow], '', '');
     end
  else if TV1.ColName[TV1.CurrentCol] = 'AFF_RESPONSABLE' then
     begin
     V_PGI.DispatchTT(6, taConsult, TV1.AsString[TV1.ColIndex('AFF_RESPONSABLE'), TV1.CurrentRow], '', '');
     end
  else if TV1.ColName[TV1.CurrentCol] = 'AFF_TIERS' then
     begin
     Req := 'SELECT T_AUXILIAIRE FROM  TIERS '+
            'WHERE T_TIERS="' + TV1.AsString[TV1.ColIndex('AFF_TIERS'), TV1.CurrentRow] + '" ' +
            'AND T_NATUREAUXI="CLI"';
     QQ := OpenSQL(Req, true,-1,'',true);
     if not QQ.EOF then
        V_PGI.DispatchTT(8, taConsult, QQ.findfield('T_AUXILIAIRE').AsString, '', '');
     ferme(QQ);
     end;

end;

//FV1 : 29/11/2016 - FS#2234 - CLOSSUR : si double-clic sur CA prévu devis, afficher liste des devis acceptés et non un devis seul.
procedure TOF_BTTABLEAUBORD.TraitementColonnePrevu(Nature : string);
Var StSQL   : string;
    CodeAff : String;
    Arg     : string;
    QQ      : TQuery;
    //
    Aff0    : String;
    Aff1    : String;
    Aff2    : String;
    Aff3    : String;
    Avenant : String;
    //
    StRange : string;
begin

  CodeAff   := TV1.AsString[TV1.ColIndex('BCO_AFFAIRE'), TV1.CurrentRow];
  BTPCodeAffaireDecoupe(CodeAff, Aff0, Aff1, Aff2, Aff3, Avenant, taConsult, True);

  Try
    StSQL := 'SELECT GP_NATUREPIECEG, GP_DATEPIECE, GP_SOUCHE, GP_NUMERO, GP_INDICEG FROM PIECE ' +
          'WHERE GP_AFFAIRE="' + CodeAff + '" AND GP_NATUREPIECEG="'+ Nature +'"';
    //
    QQ := OpenSQL(StSQL, True,-1,'',false);
    //
    if not QQ.EOF then
    begin
      if QQ.RecordCount > 1 then
      begin
        StRange := 'AFFAIRE0=' + Aff0 + ';GP_AFFAIRE1=' + Aff1 + ';GP_AFFAIRE2=' + Aff2 + ';GP_AFFAIRE3=' + Aff3 + ';AFF_ETATAFFAIRE=ACP;GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN';
        Arg     := 'ACTION=CONSULTATION;';
        if Aff0 = 'A' then
          AGLLanceFiche('BTP','BTDEVIS_MUL', StRange,'',Arg + 'STATUT=AFF;DATEDEB=' + DatetoStr(DateDeb) + ';DATEFIN=' + DatetoStr(DateFin))
        else if Aff0 = 'I' then
          AGLLanceFiche('BTP','BTDEVIS_MUL',StRange,'', Arg + 'STATUT=INT;ETAT=ACP')
        else if Aff0 = 'W' then
          AGLLanceFiche('BTP','BTDEVIS_MUL',StRange,'', Arg + 'STATUT=APP;ETAT=ACC')
        else if Aff0 = 'P' then
          AGLLanceFiche('BTP','BTDEVIS_MUL',StRange,'', Arg + 'STATUT=PRO;ETAT=ACA')
        Else
          AGLLanceFiche('BTP','BTDEVIS_MUL',StRange,'', Arg + 'STATUT=ALL;ETAT=ACP')
      end
      else
      begin
        Arg := QQ.Fields[0].AsString + ';' + QQ.Fields[1].AsString + ';' + QQ.Fields[2].AsString + ';' + IntToStr(QQ.Fields[3].AsInteger) + ';' + IntToStr(QQ.Fields[4].AsInteger);
        AppelPiece([Arg, 'ACTION=CONSULTATION'], 2);
      end;
    end;
  finally
    ferme(QQ);
  end;

end;

procedure TOF_BTTABLEAUBORD.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin

  Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));
  Aff := THEdit(GetControl('BTBAFFAIRE'));
  Aff1 := THEdit(GetControl('BTBAFFAIRE1'));
  Aff2 := THEdit(GetControl('BTBAFFAIRE2'));
  Aff3 := THEdit(GetControl('BTBAFFAIRE3'));
  Aff4 := THEdit(GetControl('BTBAVENANT'));

  Tiers := THEdit(GetControl('BTBTIERS'));

end;

// Ajout des champs dans la TOB

procedure TOF_BTTABLEAUBORD.AjouteChampSup(var TOBTMP: TOB);
begin

  //ajout par FV pour gestion envoie dans fichier XLS
  TOBTMP.addchampsupValeur ('DATEMVTDEB',     DateToStr(DateMvtDeb));
  TOBTMP.addchampsupValeur ('DATEMVTFIN',     DateToStr(DateMvtFin));

  //Champ Ajouté par FV
  TOBTMP.addchampsup('COUTANNEXE', false); TOBTMP.PutValue('COUTANNEXE', 0.0);
  TOBTMP.addchampsup('RECETTEANNEXE', false); TOBTMP.PutValue('RECETTEANNEXE', 0.0);

  if Not TOBTMP.FieldExists('CHANTIER_ASSOCIE') then  // a pu être créé avant par inserechampdesire appelé par completeleschantiers
  begin
     TOBTMP.addchampsup('CHANTIER_ASSOCIE', false); TOBTMP.PutValue('CHANTIER_ASSOCIE', '');
  end;

  if assigned(CSELACHATSTOCK) then
  begin
    if CSELACHATSTOCK.Checked then
    begin
      TOBTMP.addchampsup('REALISESTOCKPA', false); TOBTMP.PutValue('REALISESTOCKPA', 0.0);
      TOBTMP.addchampsup('REALISESTOCKPR', false); TOBTMP.PutValue('REALISESTOCKPR', 0.0);
      TOBTMP.addchampsup('REALISESTOCKPV', false); TOBTMP.PutValue('REALISESTOCKPV', 0.0);
      //
      TOBTMP.addchampsup('REALISEACHATPA', false); TOBTMP.PutValue('REALISEACHATPA', 0.0);
      TOBTMP.addchampsup('REALISEACHATPR', false); TOBTMP.PutValue('REALISEACHATPR', 0.0);
      TOBTMP.addchampsup('REALISEACHATPV', false); TOBTMP.PutValue('REALISEACHATPV', 0.0);
      //
      TOBTMP.addchampsup('PREVUACHATPA', false);   TOBTMP.PutValue('PREVUACHATPA', 0.0);
      TOBTMP.addchampsup('PREVUACHATPR', false);   TOBTMP.PutValue('PREVUACHATPR', 0.0);
      TOBTMP.addchampsup('PREVUACHATPV', false);   TOBTMP.PutValue('PREVUACHATPV', 0.0);
    end;
  end;
  //
  // Modified by f.vautrain 29/11/2017 16:47:11
  if assigned(CHKPRORATA) then
  begin
    if CHKPRORATA.Checked then
    begin
      if Prorata <> '' then
      Begin
        TOBTMP.addchampsup('PRORATA', false);
        TOBTMP.PutValue('PRORATA', 0.0);
      end;
    end;
  end; 
  //
  if assigned(CHKREVISION) then
  begin
    if CHKREVISION.Checked then
    begin
      if Revision <> '' then
      Begin
        TOBTMP.addchampsup('REVISION', false);
        TOBTMP.PutValue('REVISION', 0.0);
      end;
    end;
  end;
  //
  if CFactHorsDevis.Checked then
  begin
    TOBTMP.addchampsup('FACTHORSDEVIS', false);
    TOBTMP.PutValue('FACTHORSDEVIS', 0.0);
  end;
  //
  If CTotauxAvoir.Checked then
  begin
    TOBTMP.addchampsup('TOTALAVOIR', false);
    TOBTMP.PutValue('TOTALAVOIR', 0.0);
  end;
  //
  TOBTMP.addchampsup('PREVUPA', false); TOBTMP.PutValue('PREVUPA', 0.0);
  TOBTMP.addchampsup('PREVUPR', false); TOBTMP.PutValue('PREVUPR', 0.0);
  TOBTMP.addchampsup('PREVUPV', false); TOBTMP.PutValue('PREVUPV', 0.0);
  TOBTMP.addchampsup('TPS_PREVU', false); TOBTMP.PutValue('TPS_PREVU', 0.0);
  //
  TOBTMP.addchampsup('REALISETOTALPA', false); TOBTMP.PutValue('REALISETOTALPA', 0.0);
  TOBTMP.addchampsup('REALISETOTALPR', false); TOBTMP.PutValue('REALISETOTALPR', 0.0);
  TOBTMP.addchampsup('REALISETOTALPV', false); TOBTMP.PutValue('REALISETOTALPV', 0.0);
  //
  TOBTMP.addchampsup('REALISEMOPA', false); TOBTMP.PutValue('REALISEMOPA', 0.0);
  TOBTMP.addchampsup('REALISEMOPR', false); TOBTMP.PutValue('REALISEMOPR', 0.0);
  TOBTMP.addchampsup('REALISEMOPV', false); TOBTMP.PutValue('REALISEMOPV', 0.0);
  TOBTMP.addchampsup('TPS_REALISE', false); TOBTMP.PutValue('TPS_REALISE', 0.0);
  TOBTMP.addchampsup('TPS_REALISEINT', false); TOBTMP.PutValue('TPS_REALISEINT', 0.0);

  TOBTMP.addchampsup('REALISEFOURNITURESPA', false); TOBTMP.PutValue('REALISEFOURNITURESPA', 0.0);
  TOBTMP.addchampsup('REALISEFOURNITURESPR', false); TOBTMP.PutValue('REALISEFOURNITURESPR', 0.0);
  TOBTMP.addchampsup('REALISEFOURNITURESPV', false); TOBTMP.PutValue('REALISEFOURNITURESPV', 0.0);
  //
  TOBTMP.addchampsup('REALISEPRESTATIONSPA', false); TOBTMP.PutValue('REALISEPRESTATIONSPA', 0.0);
  TOBTMP.addchampsup('REALISEPRESTATIONSPR', false); TOBTMP.PutValue('REALISEPRESTATIONSPR', 0.0);
  TOBTMP.addchampsup('REALISEPRESTATIONSPV', false); TOBTMP.PutValue('REALISEPRESTATIONSPV', 0.0);
  //
  TOBTMP.addchampsup('REALISEMATERIELSPA', false); TOBTMP.PutValue('REALISEMATERIELSPA', 0.0);
  TOBTMP.addchampsup('REALISEMATERIELSPR', false); TOBTMP.PutValue('REALISEMATERIELSPR', 0.0);
  TOBTMP.addchampsup('REALISEMATERIELSPV', false); TOBTMP.PutValue('REALISEMATERIELSPV', 0.0);
  // AJOUT LS
	AjouteChampSupRealiseDetail (TOBTmp);
  // AJOUT FV
	AjouteChampSupRestADep(TOBTmp);
  //
  //FV1 - 30/08/2017 : FS#2667 - ARKAL : Nouvelle présentation de tableau d'analyse chantier
  TOBTMP.addchampsup('CDENONLIVREE',false); TOBTMP.PutValue('CDENONLIVREE', 0.0);
  TOBTMP.addchampsup('CDELIVREE',false);    TOBTMP.PutValue('CDELIVREE', 0.0);
  TOBTMP.addchampsup('FACTSANSCDE',false);  TOBTMP.PutValue('FACTSANSCDE', 0.0);
  //
  TOBTMP.addchampsup('FRAISACHATPA',false); TOBTMP.PutValue('AVANCEPA', 0.0);
  //
  TOBTMP.addchampsup('AVANCEPA',  false); TOBTMP.PutValue('AVANCEPA', 0.0);
  TOBTMP.addchampsup('AVANCEPR',  false); TOBTMP.PutValue('AVANCEPR', 0.0);
  TOBTMP.addchampsup('AVANCEPV',  false); TOBTMP.PutValue('AVANCEPV', 0.0);
  TOBTMP.addchampsup('TPS_AVANCE',false); TOBTMP.PutValue('TPS_AVANCE', 0.0);

  TOBTMP.addchampsup('FACTUREHT', false); TOBTMP.PutValue('FACTUREHT', 0.0);
  TOBTMP.addchampsup('FACTURE',   false); TOBTMP.PutValue('FACTURE', 0.0);
  TOBTMP.addchampsup('REGLE',     false); TOBTMP.PutValue('REGLE', 0.0);
  TOBTMP.addchampsup('REGLEHT',   false); TOBTMP.PutValue('REGLEHT', 0.0);

  TOBTMP.addchampsup('PREVU_MONTANTFG', false); TOBTMP.PutValue('PREVU_MONTANTFG', 0.0);
  TOBTMP.addchampsup('PREVU_MONTANTFC', false); TOBTMP.PutValue('PREVU_MONTANTFC', 0.0);
  TOBTMP.addchampsup('PREVU_MONTANTFR', false); TOBTMP.PutValue('PREVU_MONTANTFR', 0.0);
  //
  TOBTMP.addchampsup('ENGAGEFOU', false); TOBTMP.PutValue('ENGAGEFOU', 0.0);
  TOBTMP.addchampsup('ENGAGECHA', false); TOBTMP.PutValue('ENGAGECHA', 0.0);
  TOBTMP.addchampsup('MTRESTE',   false); TOBTMP.PutValue('MTRESTE',   0.0);
  TOBTMP.addchampsup('FACTNONPAR',false); TOBTMP.PutValue('FACTNONPAR',0.0);

  TOBTMP.addchampsup('COURECANN', false); TOBTMP.PutValue('COURECANN', 0.0);

  // AJOUT LS
  if (OptDetailETU in OptionChoixPrevuAvanc) then AjouteChampSupPrevu(TOBTmp,'ETU');
  if (OptDetailDEV in OptionChoixPrevuAvanc) then AjouteChampSupPrevu(TOBTmp,'DBT');
  if (OptDetailBCE in OptionChoixPrevuAvanc) then AjouteChampSupPrevu(TOBTmp,'BCE');
  if (OptDetailPBT in OptionChoixPrevuAvanc) or (OptGlobal in OptionChoixPrevuAvanc) then AjouteChampSupPrevu(TOBTmp,'PBT');

  // AJOUT BRL
  if (OptDetailFAC in OptionChoixPrevuAvanc) then AjouteChampSupFactureDetail (TOBTmp);

end;

procedure TOF_BTTABLEAUBORD.InsereChampsDesire(TOBL,TOBAFF : TOB);
begin

	TOBL.AddChampSupValeur ('AFF_TIERS',        TOBAFF.geTValue('AFF_TIERS'));
	TOBL.AddChampSupValeur ('AFF_APPORTEUR',    TOBAFF.geTValue('AFF_APPORTEUR'));
	TOBL.AddChampSupValeur ('AFF_ETABLISSEMENT',TOBAFF.geTValue('AFF_ETABLISSEMENT'));
	TOBL.AddChampSupValeur ('AFF_RESPONSABLE',  TOBAFF.geTValue('AFF_RESPONSABLE'));
  //
	TOBL.AddChampSupValeur ('AFF_LIBELLE',    TOBAFF.geTValue('AFF_LIBELLE'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF1',  TOBAFF.geTValue('AFF_LIBREAFF1'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF2',  TOBAFF.geTValue('AFF_LIBREAFF2'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF3',  TOBAFF.geTValue('AFF_LIBREAFF3'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF4',  TOBAFF.geTValue('AFF_LIBREAFF4'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF5',  TOBAFF.geTValue('AFF_LIBREAFF5'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF6',  TOBAFF.geTValue('AFF_LIBREAFF6'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF7',  TOBAFF.geTValue('AFF_LIBREAFF7'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF8',  TOBAFF.geTValue('AFF_LIBREAFF8'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFF9',  TOBAFF.geTValue('AFF_LIBREAFF9'));
	TOBL.AddChampSupValeur ('AFF_LIBREAFFA',  TOBAFF.geTValue('AFF_LIBREAFFA'));
	TOBL.AddChampSupValeur ('AFF_BOOLLIBRE1', TOBAFF.geTValue('AFF_BOOLLIBRE1'));
	TOBL.AddChampSupValeur ('AFF_BOOLLIBRE2', TOBAFF.geTValue('AFF_BOOLLIBRE2'));
	TOBL.AddChampSupValeur ('AFF_BOOLLIBRE3', TOBAFF.geTValue('AFF_BOOLLIBRE3'));
	TOBL.AddChampSupValeur ('AFF_CHARLIBRE1', TOBAFF.geTValue('AFF_CHARLIBRE1'));
	TOBL.AddChampSupValeur ('AFF_CHARLIBRE2', TOBAFF.geTValue('AFF_CHARLIBRE2'));
	TOBL.AddChampSupValeur ('AFF_CHARLIBRE3', TOBAFF.geTValue('AFF_CHARLIBRE3'));
	TOBL.AddChampSupValeur ('AFF_VALLIBRE1', TOBAFF.geTValue('AFF_VALLIBRE1'));
	TOBL.AddChampSupValeur ('AFF_VALLIBRE2', TOBAFF.geTValue('AFF_VALLIBRE2'));
	TOBL.AddChampSupValeur ('AFF_VALLIBRE3', TOBAFF.geTValue('AFF_VALLIBRE3'));
	TOBL.AddChampSupValeur ('AFF_DATELIBRE1', TOBAFF.geTValue('AFF_DATELIBRE1'));
	TOBL.AddChampSupValeur ('AFF_DATELIBRE2', TOBAFF.geTValue('AFF_DATELIBRE2'));
	TOBL.AddChampSupValeur ('AFF_DATELIBRE3', TOBAFF.geTValue('AFF_DATELIBRE3'));
  TOBL.AddChampSupValeur ('AFF_REFEXTERNE', TOBAFF.geTValue('AFF_REFEXTERNE'));
	TOBL.AddChampSupValeur ('AFF_DATEDEBUT', TOBAFF.geTValue('AFF_DATEDEBUT'));
	TOBL.AddChampSupValeur ('AFF_DATEFIN', TOBAFF.geTValue('AFF_DATEFIN'));
	TOBL.AddChampSupValeur ('AFF_ETATAFFAIRE', TOBAFF.geTValue('AFF_ETATAFFAIRE'));
  //
	TOBL.AddChampSupValeur ('BCO_AFFAIRE',    TOBAFF.geTValue('AFF_AFFAIRE'));
	TOBL.AddChampSupValeur ('BCO_AFFAIRE0',   TOBAFF.geTValue('AFF_AFFAIRE0'));
	TOBL.AddChampSupValeur ('BCO_AFFAIRE1',   TOBAFF.geTValue('AFF_AFFAIRE1'));
	TOBL.AddChampSupValeur ('BCO_AFFAIRE2',   TOBAFF.geTValue('AFF_AFFAIRE2'));
	TOBL.AddChampSupValeur ('BCO_AFFAIRE3',   TOBAFF.geTValue('AFF_AFFAIRE3'));

  //Test de gestion du cumul sur affaire
  //Gestion des affaires de référence
  if CCumulAffaire.checked then
    TOBL.AddChampSupValeur('CHANTIER_ASSOCIE', ChargeAffaireRef(TOBAFF.getValue('AFF_AFFAIRE'), ''));

end;

procedure TOF_BTTABLEAUBORD.CompletelesChantiers;
var Req       : string;
    QQ        : Tquery;
    TOBAFF    : TOB;
    TOBL      : TOB;
    Indice    : integer;
    WhereSql  : string;
    SelAPartir, CodCha: String;
begin

  WhereSql := RecupWhereCritere (TPageControl(GetControl('PAGES')));

  Req := 'SELECT AFF_TIERS, AFF_APPORTEUR, AFF_RESPONSABLE,'+
         'AFF_LIBELLE, AFF_LIBREAFF1, AFF_LIBREAFF2, AFF_LIBREAFF3,'+
         'AFF_LIBREAFF4, AFF_LIBREAFF5,AFF_LIBREAFF6,AFF_LIBREAFF7,'+
         'AFF_LIBREAFF8, AFF_LIBREAFF9,AFF_LIBREAFFA,'+
         'AFF_BOOLLIBRE1,AFF_BOOLLIBRE2,AFF_BOOLLIBRE3,'+
         'AFF_CHARLIBRE1,AFF_CHARLIBRE2,AFF_CHARLIBRE3,'+
         'AFF_VALLIBRE1,AFF_VALLIBRE2,AFF_VALLIBRE3,'+
         'AFF_DATELIBRE1,AFF_DATELIBRE2,AFF_DATELIBRE3,AFF_REFEXTERNE, '+
         'AFF_DATEDEBUT,AFF_DATEFIN, AFF_ETABLISSEMENT,AFF_ETATAFFAIRE,AFF_REFEXTERNE,'+
         'AFF_AFFAIRE, AFF_AFFAIRE0,AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3  FROM AFFAIRE ';

  if WhereSql = '' then
     Req := Req + ' WHERE '
  else
     begin
     WhereSql := StringReplace (WhereSql, 'BCO_', 'AFF_', [rfReplaceAll]);
     Req := Req + WhereSql+' AND ';
     end;

  if Not CCumulAffaire.checked then
    Req := Req + 'AFF_AFFAIRE IN (SELECT DISTINCT GP_AFFAIRE FROM PIECE WHERE '+
                 'GP_NATUREPIECEG IN ("DBT","ETU","BCE","PBT","AFF","FBT","FAC","AVC","ABT","FBC","B00") AND GP_AFFAIRE <> "")';

  //Ne pas prendre l'affaire si la date de fin est antérieure ou égale à la date d'entrée
  // mis en commentaire par BRL le 11/01/2011 : remplacé par l'ajout du code état affaire dans les champs dispos
  //Req := Req + ' AND (AFF_DATEFIN > "' + UsDateTime(V_PGI.DateEntree) + '") ';

  //Gestion des Sélection à Partir de ...
  SelAPartir := '';
  if CChantier.checked then SelAPartir := 'AFF_AFFAIRE0 = "A" ';

  if CContrat.checked  then
     if Selapartir <> '' then
        Selapartir := SelAPartir + ' OR AFF_AFFAIRE0 = "I" '
     else
        Selapartir := 'AFF_AFFAIRE0 = "I" ';

  if CAppel.Checked Then
     if Selapartir <> '' then
        Selapartir := SelAPartir + ' OR AFF_AFFAIRE0 = "W" '
     else
        Selapartir := 'AFF_AFFAIRE0 = "W" ';

  if Selapartir <> '' then Req := Req + ' AND (' + SelAPartir + ')';

  if CCumulAffaire.checked then
  Begin
    CodCha := Trim(Copy(CodeChantier,2,14));
  	if (CodCha <> '') then
      Req := Req + ' AND (AFF_CHANTIER LIKE "%' + CodCha + '%" OR (AFF_AFFAIRE LIKE "%' + Codcha + '%" AND AFF_AFFAIRE0 = "A"))'
  end else
  Begin
  	if CodeAffaire <> '' then
     	Req := Req + ' AND AFF_AFFAIRE = "' + CodeAffaire + '"'
  	else
     	begin
     	if CodeAffaire1 <> '' then
      	Req := Req + ' AND AFF_AFFAIRE1 = "' + CodeAffaire1 + '"';
     	if CodeAffaire2 <> '' then
        Req := Req + ' AND AFF_AFFAIRE2 = "' + CodeAffaire2 + '"';
     	if CodeAffaire3 <> '' then
        Req := Req + ' AND AFF_AFFAIRE3 = "' + CodeAffaire3 + '"';
     	end;
  end;

  if CodeClient <> '' then
     Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';

(*
  if Responsable <> '' then
     Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
*)
  if Apporteur <> '' then
     Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';
(*
  if Etablissement <> '' then
     Req := Req + ' AND AFF_ETABLISSEMENT = "' + Etablissement + '"';
*)
  // On ne prend que les chantiers dont la date de fin n'est pas atteinte c'est à dire supérieure à la date de début sélectionnée
  // pour les appels, il faut qu'ils ne soient pas annulés ou clôturés (demande delaboudiniere Mars 2015)
  if DateDeb <> iDate1900 then
  begin
    Req := Req + ' AND ((AFF_AFFAIRE0 <> "W" AND AFF_DATEFIN >= "' + UsDateTime(DateDeb)+ '")';
    //FV1 : 21/09/2017 - FS#2591 - DELABOUDINIERE :en analyse chantiers,factures d'interventions sur appels clôturés n'apparaissent pas
    Req := Req + ' OR (AFF_AFFAIRE0 = "W" AND AFF_ETATAFFAIRE<>"ANN"))'; //,"CL1")))';
    //Req := Req + ' OR (AFF_AFFAIRE0 = "W" AND AFF_ETATAFFAIRE NOT IN ("ANN","CL1")))';
  end;

  QQ := OpenSql (Req,True,-1,'',true);

  TOBAFF := TOB.Create ('LES AFFAIRES',nil,-1);
  TOBAFF.LoadDetailDB ('AFFAIRE','','',QQ,false,true);

  ferme (QQ);

  for Indice := 0 TO TOBAFF.detail.count -1 do
      begin
      if TOBBTB.FindFirst (['BCO_AFFAIRE'],[TOBAFF.detail[Indice].getValue('AFF_AFFAIRE')],true) = nil then
         begin
    	   TOBL := TOB.Create ('',TOBBTB,-1);
         insereChampsDesire(TOBL, TOBAFF.detail[Indice]);
         ChargeInfoResponsable(TOBL);
         end;
      end;

  TOBAFF.free;

end;

procedure TOF_BTTABLEAUBORD.ChargeLesChantiers;
var Req         : string;
    QQ          : Tquery;
    SelAPartir, CodCha  : String;
    Ind         : Integer;
begin

  Req := TFStat(Ecran).stSQL;

  // On met à jour la TOB uniquement pour les chantiers
  // ayant des mouvements dans les dates sélectionnées
  Req := Req + '  AND BCO_DATEMOUV BETWEEN "' + UsDateTime(DateDeb) +
               '" AND "' + UsDateTime(DateFin) + '"';

  //Gestion des Sélection à Partir de ...
  SelAPartir := '';
  if CChantier.checked then SelAPartir := 'BCO_AFFAIRE0 = "A" ';

  if CContrat.checked  then
     if Selapartir <> '' then
        Selapartir := SelAPartir + ' OR BCO_AFFAIRE0 = "I" '
     else
        Selapartir := 'BCO_AFFAIRE0 = "I" ';

  if CAppel.Checked Then
     if Selapartir <> '' then
        Selapartir := SelAPartir + ' OR BCO_AFFAIRE0 = "W" '
     else
        Selapartir := 'BCO_AFFAIRE0 = "W" ';

  if Selapartir <> '' then Req := Req + ' AND (' + SelAPartir + ')';

  if CCumulAffaire.checked then
  Begin
    CodCha := Trim(Copy(CodeChantier,2,14));
  	if (CodCha <> '') then
      Req := Req + ' AND (AFF_CHANTIER LIKE "%' + CodCha + '%" OR BCO_AFFAIRE LIKE "%' + Codcha + '%")'
  end else
  Begin
  	if CodeAffaire <> '' then
    	Req := Req + ' AND BCO_AFFAIRE = "' + CodeAffaire + '"'
  	else
    	begin
     	if CodeAffaire1 <> '' then
        Req := Req + ' AND BCO_AFFAIRE1 = "' + CodeAffaire1 + '"';
     	if CodeAffaire2 <> '' then
        Req := Req + ' AND BCO_AFFAIRE2 = "' + CodeAffaire2 + '"';
     	if CodeAffaire3 <> '' then
        Req := Req + ' AND BCO_AFFAIRE3 = "' + CodeAffaire3 + '"';
     	end;
	end;

  if CodeClient <> '' then
  begin
    Req := Req + ' AND AFF_TIERS = "' + CodeClient + '"';
  end;

  if Responsable <> '' then
  begin
    Req := Req + ' AND AFF_RESPONSABLE = "' + Responsable + '"';
  end;

  if Apporteur <> '' then
     Req := Req + ' AND AFF_APPORTEUR = "' + Apporteur + '"';

  if Etablissement <> '' then
     Req := Req + ' AND AFF_ETABLISSEMENT = "' + Etablissement + '"';

  QQ := OpenSQL(Req, true,-1,'',true);
  TOBBTB.LoadDetailDB('', '', '', QQ, true);
  ferme(QQ);

  For ind:=0 to TOBBTB.detail.count-1 do
  begin
    //
    if CEclatNatPrest.checked then
      TOBBTB.Detail[ind].AddchampSupValeur('CEclatNatPrest', 'X')
    else
      TOBBTB.Detail[ind].AddchampSupValeur('CEclatNatPrest', '-');
    //
    ChargeInfosTiers(TOBBTB.Detail[ind]);
    //
    ChargeInfoResponsable(TOBBTB.Detail[ind]);
    //
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '1');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '2');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '3');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '4');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '5');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '6');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '7');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '8');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], '9');
    ChargeInfoLIBRE(TOBBTB.Detail[ind], 'A');

    //Ajout date d'acceptation et date de fin
    //la date de fin c'est la date de fin sauf qu'il faut la ficher si <> 31/12/99
    //la date d'acceptation  c'est la date d'acceptation (AFF_DATESIGNE) Du premier devis accepté
    Req := 'SELECT ##TOP 1## AFF_DATESIGNE FROM AFFAIRE WHERE ' +
           'AFF_AFFAIREREF="' + TOBBTB.detail[Ind].GetString('BCO_AFFAIRE') + '" AND AFF_AFFAIRE0="Z" AND AFF_ETATAFFAIRE="ACP" ORDER BY AFF_DATESIGNE';
    QQ := OpenSQL(Req, True, -1, '', True);
    If Not QQ.eof then
    begin
      TOBBTB.detail[Ind].AddChampSupValeur('AFF_DATESIGNE', QQ.Findfield('AFF_DATESIGNE').AsDateTime);
    end;
    Ferme(QQ);

  end;

end;

Procedure TOF_BTTABLEAUBORD.ChargeInfosTiers(TOBTEMP : TOB);
Var Req : String;
    QQ  : TQuery;
begin

    //Ajout FV1 : charge information Table Tiers
    Req :=  'SELECT T_TIERS, T_NATUREAUXI, T_LIBELLE AS NOMTIERS, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, T_CODEPOSTAL, T_VILLE ' +
            '  FROM TIERS '+
            ' WHERE T_TIERS="' + TOBTEMP.GetString('AFF_TIERS') + '" AND T_NATUREAUXI = "CLI"';
    QQ := OpenSQL(Req, True, -1, '', True);
    If Not QQ.eof then
    begin
      TOBTEMP.AddChampSupValeur('NOMTIERS', QQ.Findfield('NOMTIERS').AsString);
      TOBTEMP.AddChampSupValeur('ADRESSE1', QQ.Findfield('T_ADRESSE1').AsString);
      TOBTEMP.AddChampSupValeur('ADRESSE2', QQ.Findfield('T_ADRESSE2').AsString);
      TOBTEMP.AddChampSupValeur('ADRESSE3', QQ.Findfield('T_ADRESSE3').AsString);
      TOBTEMP.AddChampSupValeur('CDPOSTAL', QQ.Findfield('T_CODEPOSTAL').AsString);
      TOBTEMP.AddChampSupValeur('VILLE',    QQ.Findfield('T_VILLE').AsString);
    end;
    ferme(QQ);

end;

Procedure TOF_BTTABLEAUBORD.ChargeInfoResponsable(TOBTEMP : TOB);
var REQ : String;
    QQ  : TQuery;
begin
      //Ajout FV1 : chargement des informations du Responsable
    Req := 'SELECT ARS_LIBELLE+" "+ARS_LIBELLE2 as NOMRESPONSABLE ' +
           '  FROM RESSOURCE WHERE ARS_RESSOURCE="' + TOBTEMP.GetString('AFF_RESPONSABLE') + '" AND  ARS_TYPERESSOURCE="SAL"';
    QQ := OpenSQL(Req, True, -1, '', True);
    If Not QQ.eof then
    begin
      TOBTEMP.AddChampSupValeur('NOMRESPONSABLE', QQ.Findfield('NOMRESPONSABLE').AsString);
    end;
    Ferme(QQ);

end;

procedure TOF_BTTABLEAUBORD.ChargeInfoLibre(TOBTEMP : TOB; Index : String);
var Req : String;
    QQ  : Tquery;
    LibAff : String;
begin

    LibAFF := 'AFF_LIBREAFF' + Index;

    //Ajout FV1 : chargement des informations des codes libre
    Req := 'SELECT YX_LIBELLE as LIBLIBREAFF' + Index +
           '  FROM CHOIXEXT WHERE YX_TYPE="LF' + Index + '" AND YX_CODE="'+ TOBTEMP.GetString(LibAFF) +'"';
    QQ := OpenSQL(Req, True, -1, '', True);
    If Not QQ.eof then
    begin
      TOBTEMP.AddChampSupValeur('LIBLIBREAFF'+Index , QQ.Findfield('LIBLIBREAFF'+Index).AsString);
    end;
    Ferme(QQ);

end;

procedure TOF_BTTABLEAUBORD.ChargeLesConsos;
var Req           : String;
    Indice        : integer;
    QQ            : Tquery;
    TOBTMP        : TOB;
    TOBECLATE     : TOB;
    I             : integer;
    AffaireSaisie : String;
    CodeAff       : String;
begin

  // Mise à jour des montants réalisés et engagés pour les chantiers concernés
  for i := 0 to TOBBTB.Detail.Count - 1 do
  begin
      TOBTMP := TOBBTB.Detail[i];
      if (TOBTMP <> nil) then
      begin
         AjouteChampSup(TOBTMP);
         CodeAff := TOBTMP.GetValue('BCO_AFFAIRE');
         Req := 'SELECT BNP_TYPERESSOURCE,BCO_NATUREPIECEG, BCO_SOUCHE, BCO_NUMERO, BCO_NATUREMOUV, BCO_NUMMOUV, BCO_TRANSFORME, BCO_TRAITEVENTE, ' +
                'BCO_RESSOURCE, BCO_ARTICLE, BCO_QUANTITE, ' +
                'BCO_MONTANTACH AS ACHAT, ' +
                'BCO_MONTANTPR AS REVIENT,' +
                'BCO_MONTANTHT AS VENTE,  ' +
                'ARS_TYPERESSOURCE, ' +
                'BCO_AFFAIRESAISIE AS CONTRATASSOCIE, ' +
                'GA_NATUREPRES AS NATUREPRESTATION, GA_FAMILLENIV1 AS FAMILLEART ' +
                'FROM CONSOMMATIONS ' +
                'LEFT JOIN ARTICLE ON GA_ARTICLE=BCO_ARTICLE ' +
                'LEFT JOIN '+NomTabNatPrest+' ON BNP_NATUREPRES=GA_NATUREPRES '+
                'LEFT JOIN RESSOURCE ON ARS_RESSOURCE=BCO_RESSOURCE '+
                'WHERE BCO_AFFAIRE="' + CodeAff + '" ';
         Req := Req + '  AND BCO_DATEMOUV BETWEEN "' + UsDateTime(DateMvtDeb) +
                      '" AND "' + UsDateTime(DateMvtFin) + '" ';
//         Req := Req + 'ORDER BY BCO_AFFAIRE, BCO_NATUREPIECEG, BCO_NATUREMOUV';
         Req := Req + 'ORDER BY BCO_AFFAIRE, BCO_NATUREPIECEG, ARS_TYPERESSOURCE, BCO_NATUREMOUV, BNP_NATUREPRES';
         QQ := OpenSQL(Req, true,-1,'',true);
         TOBECLATE := TOB.create('Repartition', nil, -1);
         TOBECLATE.LoadDetailDB('', '', '', QQ, true);
         ferme(QQ);
         //Boucle sur tob pour gestion des différents éclatements...
         for Indice := 0 TO TOBECLATE.detail.count -1 do
         Begin
           //Test de gestion du cumul sur affaire
           //Gestion des affaires de référence : chargement niveau ligne
           AffaireSaisie := TOBECLATE.Detail[Indice].GetString('CONTRATASSOCIE');
           if CCumulAffaire.checked then
           TOBTMP.PutValue('CHANTIER_ASSOCIE', ChargeAffaireRef(CodeAff, AffaireSaisie));
           //
           Traitement_Repartition(TOBEclate.detail[Indice], TOBTMP);
         end;
         FreeAndNil(TOBECLATE);

         // Calcul de l'engagé fournisseurs
         // Cumul des commandes fournisseurs non livrées
         Req := 'SELECT SUM(GL_QTERESTE*GL_PUHTNETDEV) AS ACHAT, ' +
                'SUM(GL_MTRESTE) AS MTRESTE ' + // --- GUINIER ---
                'FROM LIGNE ' +
                'WHERE GL_NATUREPIECEG IN ("CF","CFR","BFA") AND GL_TYPELIGNE="ART" ' +
                'AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
         Req := Req + ' AND GL_DATEPIECE BETWEEN "' + UsDateTime(DateMvtDeb) + '" ';
         Req := Req + ' AND "' + UsDateTime(DateMvtFin) + '" ';
         QQ := OpenSQL(Req, true,-1,'',true);
         if not QQ.EOF then
         begin
            MontantPA := QQ.findfield('ACHAT').AsFloat;
            TOBTMP.PutValue('ENGAGEFOU', TOBTMP.GetValue('ENGAGEFOU') + MontantPA);
            TOBTMP.PutValue('ENGAGECHA', TOBTMP.GetValue('ENGAGECHA') + MontantPA);
            // --- GUINIER ---
            MontantPA := QQ.findfield('MTRESTE').AsFloat;
            TOBTMP.PutValue('MTRESTE',   TOBTMP.GetValue('MTRESTE')   + MontantPA);
         end;
         ferme(QQ);

         // Cumul des réceptions fournisseurs et factures issues de commandes et hors période (postérieures à date de fin) : donc à ajouter
         Req := 'SELECT SUM(GL_QTEFACT*GL_PUHTNETDEV) AS ACHAT FROM LIGNE ' +
                'WHERE GL_NATUREPIECEG IN ("BLF","LFR","FF") AND GL_TYPELIGNE="ART" ' +
                'AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
         Req := Req + ' AND GL_PIECEPRECEDENTE LIKE "%;CF;%"';
         Req := Req + ' AND GL_DATEPIECE > "' + UsDateTime(DateMvtFin) + '" ';
         QQ := OpenSQL(Req, true,-1,'',true);
         if not QQ.EOF then
         begin
            MontantPA := QQ.findfield('ACHAT').AsFloat;
            TOBTMP.PutValue('ENGAGEFOU', TOBTMP.GetValue('ENGAGEFOU') + MontantPA);
            TOBTMP.PutValue('ENGAGECHA', TOBTMP.GetValue('ENGAGECHA') + MontantPA);
         end;
         ferme(QQ);

         // Complément pour l'engagé chantier
         // Cumul des livraisons chantiers issues de réceptions fournisseurs et factures et hors période (postérieures à date de fin) : donc à ajouter
         Req := 'SELECT SUM(GL_QTEFACT*GL_PUHTNETDEV) AS ACHAT FROM LIGNE ' +
                'WHERE GL_NATUREPIECEG = "LBT" AND GL_TYPELIGNE="ART" ' +
                'AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
         Req := Req + ' AND ((GL_PIECEPRECEDENTE LIKE "%;BLF;%") OR (GL_PIECEPRECEDENTE LIKE "%;FF;%"))';
         Req := Req + ' AND GL_DATEPIECE > "' + UsDateTime(DateMvtFin) + '" ';
         QQ := OpenSQL(Req, true,-1,'',true);
         if not QQ.EOF then
         begin
            MontantPA := QQ.findfield('ACHAT').AsFloat;
            TOBTMP.PutValue('ENGAGECHA', TOBTMP.GetValue('ENGAGECHA') + MontantPA);
         end;
         ferme(QQ);

         // Calcul des factures fournisseurs non parvenues
         Req := 'SELECT SUM(GL_QTERESTE*GL_PUHTNETDEV) AS ACHAT, ' +
                'SUM(GL_MTRESTE) AS MTRESTE ' + // --- GUINIER ---
                ' FROM LIGNE ' +
                'WHERE GL_NATUREPIECEG IN ("CF","CFR","BLF","LFR") AND GL_TYPELIGNE="ART" ' +
                'AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
         Req := Req + ' AND GL_DATEPIECE BETWEEN "' + UsDateTime(DateMvtDeb) + '" ';
         Req := Req + ' AND "' + UsDateTime(DateMvtFin) + '" ';
         QQ := OpenSQL(Req, true,-1,'',true);
         if not QQ.EOF then
         begin
            MontantPA := QQ.findfield('ACHAT').AsFloat;
            TOBTMP.PutValue('FACTNONPAR', TOBTMP.GetValue('FACTNONPAR') + MontantPA);
            // --- GUINIER ---*
            MontantPA := QQ.findfield('MTRESTE').AsFloat;
            TOBTMP.PutValue('FACTNONPAR', TOBTMP.GetValue('FACTNONPAR') + MontantPA);
         end;
         ferme(QQ);

         ChargeLeReste(TOBTMP);

         // Modified by f.vautrain 04/12/2017 16:55:55 - GUINIER -
         if CHKPRORATA.Checked      then ChargeProrata(TOBTMP, CodeAff);
         if CHKREVISION.Checked     then ChargeRevision(TOBTMP, CodeAff);
         if CFactHorsDevis.Checked  then ChargeHorsDevis(TOBTMP, CodeAff);
         if CTotauxAvoir.Checked    then ChargeTotalAvoir(TOBTMP, CodeAff);

         if (TOBTMP.GetDateTime('AFF_DATEFIN') > DateMvtFin) then
            ChargeLeResteADepenser(TOBTMP,TOBEclate,CodeAff);

         if CSELACHATSTOCK.Checked then ChargeAchatStocks(TOBTMP, CodeAff);

         ChargeAchatFrais(TOBTMP, CodeAff);

         //FV1 : 30/08/2017 - FS#2667 - ARKAL : Nouvelle présentation de tableau d'analyse chantier
         if CSELACHATSTOCK.Checked then
         begin
          ChargelesCommandesNonLivrees(TOBTMP, CodeAff);
          ChargelesCommandesLivrees(TOBTMP, CodeAff);
          ChargelesFacturesansCommandes(TOBTMP, CodeAff);
         end;

    end;
  end;

end;

Procedure TOF_BTTABLEAUBORD.Traitement_Repartition(TOBEclate, TOBTMP : TOB);
Var NatureMouv    : String;
    TypeRessource : String;
    NatPrestation : String;
    FamNiveau1    : String;
    Montant       : Double;
Begin

  // Traitements du réalisé
  MontantPA     := TOBECLATE.GetDouble('ACHAT');
  MontantPR     := TOBECLATE.GetDouble('REVIENT');
  MontantPV     := TOBECLATE.GetDouble('VENTE');
  Qte           := TOBECLATE.GetDouble('BCO_QUANTITE');
  //
  NatPrestation := TOBECLATE.GetString('NATUREPRESTATION');
  FamNiveau1    := TOBECLATE.GetString('FAMILLEART');
  NatureMouv    := TOBECLATE.GetString('BCO_NATUREMOUV');
  TypeRessource := TOBECLATE.GetString('ARS_TYPERESSOURCE');
  //
  if TypeRessource = '' then
     TypeRessource := TOBECLATE.GetString('BNP_TYPERESSOURCE');
  //
  if (TOBECLATE.GetString('BCO_NATUREPIECEG') = '')    or   // consos saisies
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'AFS') or   // correction BRL le 06/04/05 prise ne compte des avoirs en consos
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'AF')  or   // ...
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'BFC') or   // Retour chantier
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'FPR') or   // consos issues de retour d'intervention avec facturation provisoire
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'FAC') or   // consos issues de retour d'intervention avec facturation définitive
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'LBT') then // consos issues de bons de livraison clients sur chantier
     begin
     // Répartition Fournitures
     if NatureMouv = 'FOU' then
        Repartition_Fournisseur(FamNiveau1, TOBTMP)
     // Répartition MO Interne
     Else if NatureMouv = 'MO' then
     begin
        //Modif FV : Dev. prioritaire DSL le 05/06/2012
        if (TypeRessource = 'SAL') and (CoefFG_Param <> 0) then
        begin
          MontantPA := Arrondi(TauxHoraire * Qte,V_PGI.OkDecV);
          MontantPR := Arrondi(MontantPA * CoefFG_Param, V_PGI.OkDecV);
        end;
        Repartition_MOInterne(TypeRessource, NatPrestation, TobTMP);
     end
     // Répartition Frais
     Else if NatureMouv = 'FRS' then
        Repartition_Frais(TypeRessource, FamNiveau1, TobTMP)
     // Répartition Matériels + Outils
     Else if NatureMouv = 'RES' then
        Repartition_MatOutil(TypeRessource, NatPrestation, TobTMP)
     // Répartition Intérim + Sous-traitance + Location + Autres
     Else if NatureMouv = 'EXT' then
        Repartition_IntSsTraitLocAut(TypeRessource, NatPrestation, TobTMP)
     // Répartition coûts et recettes annexes
     Else if NatureMouv = 'RAN' Then
        Begin
        TOBTMP.PutValue('COURECANN', TOBTMP.GetValue('COURECANN') - MontantPV);
        MontantPV := MontantPV * -1;
        TOBTMP.PutValue('RECETTEANNEXE', TOBTMP.GetValue('RECETTEANNEXE') + MontantPV);
        end
     Else If NatureMouv = 'FAN' then
        Begin
        TOBTMP.PutValue('COURECANN', TOBTMP.GetValue('COURECANN') - MontantPV);
        TOBTMP.PutValue('COUTANNEXE', TOBTMP.GetValue('COUTANNEXE') + MontantPV);
        end;
     end;
  // consos issues de bons de réception fournisseurs (engagé chantier)
  if (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'BLF') or
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'LFR')  or
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'FF')  then
  begin
     if (TOBECLATE.getValue('BCO_TRANSFORME') = '-') and (TOBECLATE.getValue('BCO_TRAITEVENTE') = '-') then
//        TOBTMP.PutValue('ENGAGECHA', TOBTMP.GetValue('ENGAGECHA') + MontantPA);
  // consos issues de bons de commande fournisseurs (engagé fournisseur)
  end else if (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'CF') or
          (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'CFR')  or
          (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'BFA')then
  begin
     if (TOBECLATE.getValue('BCO_TRANSFORME') = '-') then
     begin
//       TOBTMP.PutValue('ENGAGEFOU', TOBTMP.GetValue('ENGAGEFOU') + MontantPA);
//       TOBTMP.PutValue('ENGAGECHA', TOBTMP.GetValue('ENGAGECHA') + MontantPA);
     end;
     DateTransfoOK(TOBECLATE, Montant);
     if Montant <> 0 then
     begin
//       TOBTMP.PutValue('ENGAGEFOU', TOBTMP.GetValue('ENGAGEFOU') + Montant);
//       TOBTMP.PutValue('ENGAGECHA', TOBTMP.GetValue('ENGAGECHA') + Montant);
     end;
  end;

end;

Procedure TOF_BTTABLEAUBORD.Traitement_Repartition_RAD(TOBL, TOBTMP : TOB);
Var NatureMouv    : String;
    TypeRessource : String;
    NatPrestation : String;
    FamNiveau1    : String;
    MtReste       : Double;
    QtReste       : Double;
begin

  // Traitements du réalisé
  MtReste    := TOBL.GetValue('MTRESTE');
  QtReste    := TOBL.GetValue('QTRESTE');
  //
  TypeRessource := TOBL.GetString('TYPERESSOURCE');
  NatPrestation := TOBL.GetString('NATUREPRESTATION');
  NatureMouv    := '';
  FamNiveau1    := TOBL.GetString('FAMILLEART');
  //
  if TOBL.GetSTring('NATUREPRESTATION') = 'FRA' then
    NatureMouv := 'FRS'
  else
  begin
    if TOBL.GetSTring('TYPERESSOURCE') = 'SAL' then //Main d'Oeuvre
      NatureMouv := 'MO'
    Else If TOBL.GetSTring('TYPERESSOURCE') = 'ST' then  //Sous-Traitantce
      NatureMouv := 'EXT'
    Else If TOBL.GetSTring('TYPERESSOURCE') = 'MAT' then //Matériels Internes
      NatureMouv := 'RES'
    Else IF TOBL.GetString('TYPERESSOURCE') = 'AUT' then //Diverses
      NatureMouv := 'EXT'
    Else IF TOBL.GetString('TYPERESSOURCE') = 'LOC' then //Prestation Externes
      NatureMouv := 'EXT'
    Else IF TOBL.GetString('TYPERESSOURCE') = 'INT' then //Main d'Oeuvre Interim
      NatureMouv := 'MO'
    Else IF TOBL.GetString('TYPERESSOURCE') = 'OUT' then
      NatureMouv := 'RES'
    Else IF TOBL.GetString('TYPERESSOURCE') = 'FAC' then
      NatureMouv := 'FAC'
    Else
      NatureMouv := '';
  end;
  //
  // Répartition MO interne
  if NatureMouv = 'MO' then
    Repartition_MOInterne_RAD(TypeRessource, NatPrestation, TobTMP, Mtreste, QtReste)
  // Répartition Frais
  Else if NatureMouv = 'FRS' then
    Repartition_Frais_RAD(TypeRessource, FamNiveau1, TobTMP, Mtreste, QtReste )
  // Répartition Matériels + Outils
  Else if NatureMouv = 'RES' then
    Repartition_MatOutil_RAD(TypeRessource, NatPrestation, TobTMP,  Mtreste, QtReste)
  // Répartition Intérim + Sous-traitance + Location + Autres
  Else if NatureMouv = 'EXT' then
    Repartition_IntSsTraitLocAut_RAD(TypeRessource, NatPrestation, TobTMP, Mtreste, QtReste)
  // Répartition Facturation
  Else if NatureMouv = 'FAC' then
    Repartition_Facturation_RAD(TypeRessource, NatPrestation, TobTMP, Mtreste, QtReste)
  Else if NatureMouv = '' then
    Repartition_Fourniture_RAD(TypeRessource, FamNiveau1, TobTMP, Mtreste, QtReste);

  //FV1 : 05/02/2014 - FS#804 - BAGE : en analyses chantiers, ajouter les champs totalisant des RAD et Fin d'affaire par catégorie
  if NatureMouv <> 'FAC' then
  begin
    TOBTMP.PutValue('RESTEADEP_MTRESTE',TOBTMP.GetValue('RESTEADEP_MTRESTE') + MtReste);
    TOBTMP.PutValue('RESTEADEP_QTRESTE',TOBTMP.GetValue('RESTEADEP_QTRESTE') + QtReste);
  end;

end;

Procedure TOF_BTTABLEAUBORD.Repartition_Fournisseur(FamNiveau1 : String; TOBTMP : TOB);
begin

  TOBTMP.PutValue('REALISEFOURNITURESPA', TOBTMP.GetValue('REALISEFOURNITURESPA') + MontantPA);
  TOBTMP.PutValue('REALISEFOURNITURESPR', TOBTMP.GetValue('REALISEFOURNITURESPR') + MontantPR);
  TOBTMP.PutValue('REALISEFOURNITURESPV', TOBTMP.GetValue('REALISEFOURNITURESPV') + MontantPV);

  if CEclatFamNiv1.checked then
     begin
     if FamNiveau1 = '' Then FamNiveau1 := 'NONRENSEIGNE';
     Repartition_Eclatement('CONSOMME','_FOU_'+FamNiveau1, '_PA', MontantPA, TOBTMP);
     Repartition_Eclatement('CONSOMME','_FOU_'+FamNiveau1, '_PV', MontantPV, TOBTMP);
     Repartition_Eclatement('CONSOMME','_FOU_'+FamNiveau1, '_PR', MontantPR, TOBTMP);
     end;
end;

Procedure TOF_BTTABLEAUBORD.Repartition_MoInterne(TypeRessource, NatPrestation : String; TOBTMP : Tob);
Begin
  // Traitement des intérimaires saisis en détail dans les heures
     if TypeRessource = 'INT' then
        begin
        TOBTMP.PutValue('REALISEPRESTATIONSPA', TOBTMP.GetValue('REALISEPRESTATIONSPA') + MontantPA);
        TOBTMP.PutValue('REALISEPRESTATIONSPR', TOBTMP.GetValue('REALISEPRESTATIONSPR') + MontantPR);
        TOBTMP.PutValue('REALISEPRESTATIONSPV', TOBTMP.GetValue('REALISEPRESTATIONSPV') + MontantPV);
        //
        TOBTMP.PutValue('REALISE_INTERIM_PA', TOBTMP.GetValue('REALISE_INTERIM_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_INTERIM_PR', TOBTMP.GetValue('REALISE_INTERIM_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_INTERIM_PV', TOBTMP.GetValue('REALISE_INTERIM_PV') + MontantPV);
        TOBTMP.PutValue('TPS_REALISEINT',     TOBTMP.GetValue('TPS_REALISEINT') + Qte);
  			//
				if CEclatNatPrest.checked then
     			Begin
     			Repartition_Eclatement('CONSOMME', '_INT_'+ NatPrestation, '_PA', MontantPA, TOBTMP);
     			Repartition_Eclatement('CONSOMME', '_INT_'+ NatPrestation, '_PV', MontantPV, TOBTMP);
     			Repartition_Eclatement('CONSOMME', '_INT_'+ NatPrestation, '_PR', MontantPR, TOBTMP);
     			Repartition_Eclatement('CONSOMME', '_INT_'+ NatPrestation, '_QTE', Qte, TOBTMP);
     			end;
        end
     else
        // Traitement des salariés
        begin
        TOBTMP.PutValue('REALISEMOPA', TOBTMP.GetValue('REALISEMOPA') + MontantPA);
        TOBTMP.PutValue('REALISEMOPR', TOBTMP.GetValue('REALISEMOPR') + MontantPR);
        TOBTMP.PutValue('REALISEMOPV', TOBTMP.GetValue('REALISEMOPV') + MontantPV);
        TOBTMP.PutValue('TPS_REALISE', TOBTMP.GetValue('TPS_REALISE') + Qte);
        //
        TOBTMP.PutValue('REALISE_SAL_PA', TOBTMP.GetValue('REALISE_SAL_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_SAL_PR', TOBTMP.GetValue('REALISE_SAL_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_SAL_PV', TOBTMP.GetValue('REALISE_SAL_PV') + MontantPV);
  			//
  			if CEclatNatPrest.checked then
     			Begin
     			Repartition_Eclatement('CONSOMME','_MO_'+ NatPrestation, '_PA', MontantPA, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_MO_'+ NatPrestation, '_PV', MontantPV, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_MO_'+ NatPrestation, '_PR', MontantPR, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_MO_'+ NatPrestation, '_QTE', Qte, TOBTMP);
     			end;
        end;
end;

Procedure TOF_BTTABLEAUBORD.Repartition_Frais(TypeRessource, FamNiveau1 : String; TobTMP : TOB);
begin
     if TypeRessource = 'INT' then
        begin
  	    // Traitement des intérimaires saisis en détail dans les frais
        TOBTMP.PutValue('REALISEPRESTATIONSPA', TOBTMP.GetValue('REALISEPRESTATIONSPA') + MontantPA);
        TOBTMP.PutValue('REALISEPRESTATIONSPR', TOBTMP.GetValue('REALISEPRESTATIONSPR') + MontantPR);
        TOBTMP.PutValue('REALISEPRESTATIONSPV', TOBTMP.GetValue('REALISEPRESTATIONSPV') + MontantPV);
        //
        TOBTMP.PutValue('REALISE_INTERIM_PA', TOBTMP.GetValue('REALISE_INTERIM_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_INTERIM_PR', TOBTMP.GetValue('REALISE_INTERIM_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_INTERIM_PV', TOBTMP.GetValue('REALISE_INTERIM_PV') + MontantPV);
			  //
        if CEclatFamNiv1.checked then
        	begin
     			if FamNiveau1 = '' Then FamNiveau1 := 'NONRENSEIGNE';
     			Repartition_Eclatement('CONSOMME','_FRSINT_'+FamNiveau1, '_PA', MontantPA, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_FRSINT_'+FamNiveau1, '_PV', MontantPV, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_FRSINT_'+FamNiveau1, '_PR', MontantPR, TOBTMP);
     			end;
        end
     else
        begin
        // Realise main d'oeuvre salarie
        TOBTMP.PutValue('REALISE_SAL_PA', TOBTMP.GetValue('REALISE_SAL_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_SAL_PR', TOBTMP.GetValue('REALISE_SAL_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_SAL_PV', TOBTMP.GetValue('REALISE_SAL_PV') + MontantPV);
        //
        TOBTMP.PutValue('REALISEMOPA', TOBTMP.GetValue('REALISEMOPA') + MontantPA);
        TOBTMP.PutValue('REALISEMOPR', TOBTMP.GetValue('REALISEMOPR') + MontantPR);
        TOBTMP.PutValue('REALISEMOPV', TOBTMP.GetValue('REALISEMOPV') + MontantPV);
			  //
        if CEclatFamNiv1.checked then
        	begin
     			if FamNiveau1 = '' Then FamNiveau1 := 'NONRENSEIGNE';
     			Repartition_Eclatement('CONSOMME','_FRSMO_'+FamNiveau1, '_PA', MontantPA, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_FRSMO_'+FamNiveau1, '_PV', MontantPV, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_FRSMO_'+FamNiveau1, '_PR', MontantPR, TOBTMP);
     			end;
        end;
end;

Procedure TOF_BTTABLEAUBORD.Repartition_MatOutil(TypeRessource, NatPrestation : String; TobTMP : TOB);
begin
     if TypeRessource = 'MAT' then
        begin
        TOBTMP.PutValue('REALISE_MAT_PA', TOBTMP.GetValue('REALISE_MAT_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_MAT_PR', TOBTMP.GetValue('REALISE_MAT_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_MAT_PV', TOBTMP.GetValue('REALISE_MAT_PV') + MontantPV);
  			if CEclatNatPrest.checked then
     			Begin
     			Repartition_Eclatement('CONSOMME','_MAT_'+NatPrestation, '_PA', MontantPA, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_MAT_'+NatPrestation, '_PV', MontantPV, TOBTMP);
          Repartition_Eclatement('CONSOMME','_MAT_'+NatPrestation, '_PR', MontantPR, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_MAT_'+NatPrestation, '_QTE', Qte, TOBTMP);
          end;
     		end
     else
        begin
        TOBTMP.PutValue('REALISE_OUTIL_PA', TOBTMP.GetValue('REALISE_OUTIL_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_OUTIL_PR', TOBTMP.GetValue('REALISE_OUTIL_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_OUTIL_PV', TOBTMP.GetValue('REALISE_OUTIL_PV') + MontantPV);
  			if CEclatNatPrest.checked then
     			Begin
     			Repartition_Eclatement('CONSOMME','_OUT_'+NatPrestation, '_PA', MontantPA, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_OUT_'+NatPrestation, '_PV', MontantPV, TOBTMP);
          Repartition_Eclatement('CONSOMME','_OUT_'+NatPrestation, '_PR', MontantPR, TOBTMP);
     			Repartition_Eclatement('CONSOMME','_OUT_'+NatPrestation, '_QTE', Qte, TOBTMP);
          end;
        end;
     TOBTMP.PutValue('REALISEMATERIELSPA', TOBTMP.GetValue('REALISEMATERIELSPA') + MontantPA);
     TOBTMP.PutValue('REALISEMATERIELSPR', TOBTMP.GetValue('REALISEMATERIELSPR') + MontantPR);
     TOBTMP.PutValue('REALISEMATERIELSPV', TOBTMP.GetValue('REALISEMATERIELSPV') + MontantPV);
end;

//Répartition Intérim + Sous-traitance + Location + Autres
Procedure TOF_BTTABLEAUBORD.Repartition_IntSsTraitLocAut(TypeRessource, NatPrestation : String; TobTMP : TOB);
begin
     if TypeRessource = 'INT' then
        begin
        TOBTMP.PutValue('REALISE_INTERIM_PA', TOBTMP.GetValue('REALISE_INTERIM_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_INTERIM_PR', TOBTMP.GetValue('REALISE_INTERIM_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_INTERIM_PV', TOBTMP.GetValue('REALISE_INTERIM_PV') + MontantPV);
        //FV1 : 15/05/2014 - FS#986 - ERCC - différence de montant indiqué pour le champ TPS_REALISE
        //FV1 : 16/02/2015 - FS#1465 - GROUPE LB : Dans les analyses, le temps réalisé intérim est à 0
        TOBTMP.PutValue('TPS_REALISEINT',     TOBTMP.GetValue('TPS_REALISEINT') + Qte);
        end
     else if TypeRessource = 'ST' then
        begin
        TOBTMP.PutValue('REALISE_ST_PA', TOBTMP.GetValue('REALISE_ST_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_ST_PR', TOBTMP.GetValue('REALISE_ST_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_ST_PV', TOBTMP.GetValue('REALISE_ST_PV') + MontantPV);
        end
     else if TypeRessource = 'LOC' then
        begin
        TOBTMP.PutValue('REALISE_LOC_PA', TOBTMP.GetValue('REALISE_LOC_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_LOC_PR', TOBTMP.GetValue('REALISE_LOC_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_LOC_PV', TOBTMP.GetValue('REALISE_LOC_PV') + MontantPV);
        end
     else if TypeRessource = 'AUT' then
        begin
        TOBTMP.PutValue('REALISE_AUTRE_PA', TOBTMP.GetValue('REALISE_AUTRE_PA') + MontantPA);
        TOBTMP.PutValue('REALISE_AUTRE_PR', TOBTMP.GetValue('REALISE_AUTRE_PR') + MontantPR);
        TOBTMP.PutValue('REALISE_AUTRE_PV', TOBTMP.GetValue('REALISE_AUTRE_PV') + MontantPV);
        end;
     TOBTMP.PutValue('REALISEPRESTATIONSPA', TOBTMP.GetValue('REALISEPRESTATIONSPA') + MontantPA);
     TOBTMP.PutValue('REALISEPRESTATIONSPR', TOBTMP.GetValue('REALISEPRESTATIONSPR') + MontantPR);
     TOBTMP.PutValue('REALISEPRESTATIONSPV', TOBTMP.GetValue('REALISEPRESTATIONSPV') + MontantPV);

  //
  if CEclatNatPrest.checked then
     Begin
     Repartition_Eclatement('CONSOMME','_EXT_'+NatPrestation, '_PA', MontantPA, TOBTMP);
     Repartition_Eclatement('CONSOMME','_EXT_'+NatPrestation, '_PV', MontantPV, TOBTMP);
     Repartition_Eclatement('CONSOMME','_EXT_'+NatPrestation, '_PR', MontantPR, TOBTMP);
     Repartition_Eclatement('CONSOMME','_EXT_'+NatPrestation, '_QTE', Qte, TOBTMP);
     end;

end;

Procedure TOF_BTTABLEAUBORD.Repartition_MoInterne_RAD(TypeRessource, NatPrestation : String; TOBTMP : Tob; MTReste, QTReste : Double);
Begin
  // Traitement des intérimaires saisis en détail dans les heures
  if TypeRessource = 'INT' then
    begin
    TOBTMP.PutValue('RESTADEP_INT_MTRESTE', TOBTMP.GetValue('RESTADEP_INT_MTRESTE') + MtReste);
    TOBTMP.PutValue('RESTADEP_INT_QTRESTE', TOBTMP.GetValue('RESTADEP_INT_QTRESTE') + QtReste);
    //
    if CEclatNatPrest.checked then
      Begin
      Repartition_Eclatement('RESTADEP','_INT_'+NatPrestation, '_MT', MtReste, TOBTMP);
      Repartition_Eclatement('RESTADEP','_INT_'+NatPrestation, '_QTE', QtReste, TOBTMP);
      end;
    end
  else
    // Traitement des salariés
    begin
    //
    TOBTMP.PutValue('RESTADEP_SAL_MTRESTE', TOBTMP.GetValue('RESTADEP_SAL_MTRESTE') + MtReste);
    TOBTMP.PutValue('RESTADEP_SAL_QTRESTE', TOBTMP.GetValue('RESTADEP_SAL_QTRESTE') + QtReste);
    //
    if CEclatNatPrest.checked then
      Begin
      Repartition_Eclatement('RESTADEP','_MO_'+NatPrestation, '_MT', MtReste, TOBTMP);
      Repartition_Eclatement('RESTADEP','_MO_'+NatPrestation, '_QTE', QtReste, TOBTMP);
      end;
    end;

end;

Procedure TOF_BTTABLEAUBORD.Repartition_Frais_RAD(TypeRessource, FamNiveau1 : String; TobTMP : TOB; MTReste, QTReste : Double);
begin

  if TypeRessource = 'INT' then
    begin
    // Traitement des intérimaires saisis en détail dans les frais
    TOBTMP.PutValue('RESTADEP_INT_MTRESTE', TOBTMP.GetValue('RESTADEP_INT_MTRESTE') + MtReste);
    //
// BRL 14/08 : La gestion par famille ne peut fonctionner car les codes articles ne sont pas mémorisés dans la table BTRESTEADEP
{
    if CEclatFamNiv1.checked then
      begin
      if FamNiveau1 = '' Then FamNiveau1 := 'NONRENSEIGNE';
      Repartition_Eclatement('RESTADEP','_FRSINT'+FamNiveau1, '_MT', MtReste, TOBTMP);
      end;
}
    end
  else
    begin
    // Realise main d'oeuvre salarie
    TOBTMP.PutValue('RESTADEP_SAL_MTRESTE', TOBTMP.GetValue('RESTADEP_SAL_MTRESTE') + MtReste);
    //
// BRL 14/08 : La gestion par famille ne peut fonctionner car les codes articles ne sont pas mémorisés dans la table BTRESTEADEP
{
    if CEclatFamNiv1.checked then
      begin
//      if FamNiveau1 = '' Then FamNiveau1 := 'NONRENSEIGNE';
      Repartition_Eclatement('RESTADEP','_FRSMO'+FamNiveau1, '_MT', MtReste, TOBTMP);
      end;
}
    end;
end;

Procedure TOF_BTTABLEAUBORD.Repartition_MatOutil_RAD(TypeRessource, NatPrestation : String; TobTMP : TOB; MtReste, QtReste : Double);
begin

  if TypeRessource = 'MAT' then
  begin
    TOBTMP.PutValue('RESTADEP_MAT_MTRESTE', TOBTMP.GetValue('RESTADEP_MAT_MTRESTE') + MtReste);
    if CEclatNatPrest.checked then
      Begin
      Repartition_Eclatement('RESTADEP','_MAT_'+NatPrestation, '_MT', MtReste, TOBTMP);
      end;
  end
  else
  begin
    TOBTMP.PutValue('RESTADEP_OUT_MTRESTE', TOBTMP.GetValue('RESTADEP_OUT_MTRESTE') + MtReste);
    if CEclatNatPrest.checked then
      Begin
      Repartition_Eclatement('RESTADEP','_OUT_'+NatPrestation, '_MT', MtReste, TOBTMP);
      end;
  end;

end;

//Répartition Intérim + Sous-traitance + Location + Autres
Procedure TOF_BTTABLEAUBORD.Repartition_IntSsTraitLocAut_RAD(TypeRessource, NatPrestation : String; TobTMP : TOB; MtReste, QTReste : Double);
begin

  if TypeRessource = 'INT' then
  begin
    TOBTMP.PutValue('RESTADEP_INT_MTRESTE', TOBTMP.GetValue('RESTADEP_INT_MTRESTE') + MtReste);
  end
  else if TypeRessource = 'ST' then
  begin
    TOBTMP.PutValue('RESTADEP_ST_MTRESTE', TOBTMP.GetValue('RESTADEP_ST_MTRESTE') + MtReste);
  end
  else if TypeRessource = 'LOC' then
  begin
    TOBTMP.PutValue('RESTADEP_LOC_MTRESTE', TOBTMP.GetValue('RESTADEP_LOC_MTRESTE') + MtReste);
  end
  else if TypeRessource = 'AUT' then
  begin
    TOBTMP.PutValue('RESTADEP_AUT_MTRESTE', TOBTMP.GetValue('RESTADEP_AUT_MTRESTE') + MtReste);
  end;
  //
  if CEclatNatPrest.checked then
  Begin
    Repartition_Eclatement('RESTADEP','_EXT_'+NatPrestation, '_MT', MtReste, TOBTMP);
  end;

end;

Procedure TOF_BTTABLEAUBORD.Repartition_Facturation_RAD(TypeRessource, NatPrestation : String; TobTMP : TOB; MtReste, QtReste : Double);
begin

  TOBTMP.PutValue('RESTADEP_FAC_MTRESTE', TOBTMP.GetValue('RESTADEP_FAC_MTRESTE') + MtReste);

end;

Procedure TOF_BTTABLEAUBORD.Repartition_Fourniture_RAD(TypeRessource, NatPrestation : String; TobTMP : TOB; MtReste, QtReste : Double);
begin

  TOBTMP.PutValue('RESTADEP_FOU_MTRESTE', TOBTMP.GetValue('RESTADEP_FOU_MTRESTE') + MtReste);
  if TypeRessource = 'ACH' then
    TOBTMP.PutValue('RESTADEP_ACH_MTRESTE', TOBTMP.GetValue('RESTADEP_ACH_MTRESTE') + MtReste)
  else
    TOBTMP.PutValue('RESTADEP_STK_MTRESTE', TOBTMP.GetValue('RESTADEP_STK_MTRESTE') + MtReste);

end;

{*
Procedure TOF_BTTABLEAUBORD.MAJ_PremiereLigne(NomChampSup : String);
Begin

  if TOBBTB.Detail.count = 0 Then Exit;
  if Not TOBBTB.Detail[0].FieldExists(NomChampSup) then
  begin
     TOBBTB.Detail[0].AddChampSupValeur(NomChampSup, 0.0);
  end;

end;
*}
{*Procedure TOF_BTTABLEAUBORD.Repartition_Eclatement(Eclatement, TypeMt : String; Montant : Double; TOBTMP : Tob);
var MtAvant, MtApres : Double;
Begin

  if Not TobTMP.FieldExists('CONSOMME' + Eclatement + TypeMt) then
  begin
    TOBTMP.AddChampSupValeur('CONSOMME' + Eclatement + TypeMt, Montant);
    MAJ_PremiereLigne('CONSOMME' + Eclatement + TypeMt);
  end else
  begin
    MtAvant := TOBTMP.GetValue('CONSOMME' + Eclatement + TypeMt);
    MtAvant := Arrondi(MtAvant, V_PGI.OkDecV);
    MtApres := MtAvant + Montant;
    MtApres := Arrondi(MtApres, V_PGI.OkDecV);
    TOBTMP.PutValue('CONSOMME' + Eclatement + TypeMt,  MtApres);
  end;
end;
*}

procedure TOF_BTTABLEAUBORD.ChargeLeReste(TOBTMP : TOB);
var Req         : String;
    WherePiece  : String;
    DateFinAff  : string;
    QQ          : Tquery;
    TobFact     : TOB;
    TOBF        : TOB;
    indice      : Integer;
  //TOBTMP: TOB;
    Montant     : DOuble;
    Coef        : Double;
    TTC         : Double;
    HTX         : Double;
    Cloture     : Boolean;
    TypeFac     : String;
begin

  // Mise à jour des autres montants pour les chantiers concernés

  // recalcul du montant reglé par situations
  if CSelReglement.Checked then
     CalcReglementSituations(TOBTMP.GetValue('BCO_AFFAIRE'));

  // Calcul des Totaux Réalisés
  Montant := TOBTMP.GetValue('REALISEMOPA') +
             TOBTMP.GetValue('REALISEFOURNITURESPA') +
             TOBTMP.GetValue('REALISEMATERIELSPA') +
             TOBTMP.GetValue('REALISEPRESTATIONSPA') ;
  TOBTMP.PutValue('REALISETOTALPA', Montant);


  Montant := TOBTMP.GetValue('REALISEMOPR') +
             TOBTMP.GetValue('REALISEFOURNITURESPR') +
             TOBTMP.GetValue('REALISEMATERIELSPR') +
             TOBTMP.GetValue('REALISEPRESTATIONSPR') ;
  TOBTMP.PutValue('REALISETOTALPR', Montant);

  Montant := TOBTMP.GetValue('REALISEMOPV') +
             TOBTMP.GetValue('REALISEFOURNITURESPV') +
             TOBTMP.GetValue('REALISEMATERIELSPV') +
             TOBTMP.GetValue('REALISEPRESTATIONSPV') ;
  TOBTMP.PutValue('REALISETOTALPV', Montant);

  if (OptDetailDev in OptionChoixPrevuAvanc) then SetPrevuAvance (TOBTMP,'DBT',OptionChoixPrevuAvanc);

  if (OptDetailPBT in OptionChoixPrevuAvanc) or(OptGlobal in OptionChoixPrevuAvanc) then
     begin
     if TOBTMP.GetValue('BCO_AFFAIRE0')='I' then
        SetPrevuAvance (TOBTMP,'AFF',OptionChoixPrevuAvanc)
     else
       	SetPrevuAvance (TOBTMP,'PBT',OptionChoixPrevuAvanc);
     end;

  if (OptDetailETU in OptionChoixPrevuAvanc) then  SetPrevuAvance (TOBTMP,'ETU',OptionChoixPrevuAvanc);
  if (OptDetailBCE in OptionChoixPrevuAvanc) then  SetPrevuAvance (TOBTMP,'BCE',OptionChoixPrevuAvanc);

  if TOBTMP.GetValue('BCO_AFFAIRE0')='A' then
// MODIF LS
     WherePiece := ' (GP_NATUREPIECEG IN ("FBT","ABT","FAC","AVC","DAC","FBC","B00") OR (GP_NATUREPIECEG In ("FBP","DAP") AND (GP_VIVANTE="X")))  '
//     WherePiece := ' GP_NATUREPIECEG IN ("FBT") '
  else
     WherePiece := ' GP_NATUREPIECEG IN ("FBT","ABT","FAC","AVC","FBC") ';

  // Récupération du facturé
  Req := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_AFFAIREDEVIS,GP_DATEPIECE,GP_TOTALHTDEV AS FACTUREHT, GP_TOTALTTCDEV AS FACTURE ';

  //vérification si le piecetrait existe pour cette facture !!!
  if TobTMP.GetString('AFF_MANDATAIRE') <> '' then
  begin
    Req := Req + ',BPE_TOTALHTDEV AS FACTHTTRAIT, BPE_TOTALTTCDEV AS FACTTTCTRAIT ';

    req := req + 'FROM PIECE LEFT JOIN PIECETRAIT ON ' +
                 'BPE_NATUREPIECEG=GP_NATUREPIECEG AND BPE_SOUCHE=GP_SOUCHE AND ' +
                 'BPE_NUMERO=GP_NUMERO AND BPE_INDICEG=GP_INDICEG ';
    req := req + 'WHERE '+ WherePiece;
    Req := req + '  AND BPE_FOURNISSEUR="" ';
  end
  else
  begin
    req := req + 'FROM PIECE WHERE '+ WherePiece;
  end;

  Req := Req + '  AND GP_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
  Req := Req + '  AND GP_DATEPIECE BETWEEN "' + UsDateTime(DateMvtDeb) + '"';
  Req := Req + '  AND "' + UsDateTime(DateMvtFin) + '" ';
  //Req := Req + 'GROUP BY GP_AFFAIRE';

  QQ := OpenSql (req, true);
  TOBFact := TOB.Create('FACTURE', nil, -1);
  TOBFact.loadDetailDb ('FACTURE','','',QQ , false);
  Ferme (QQ);

  if TOBFACT.Detail.count = 0 then
  begin
    freeandnil (TobFact);
    exit;
  end;

  HTX := 0;
  TTC := 0;

  //Boucle facture ht et facture ttc en fonction valeurs renseignées
  For indice := 0 to TobFact.detail.count - 1  do
  begin
    TOBF := TOBFACT.Detail[Indice];
    //test sur DGD...
    Cloture := ISFactureCloture(TOB2Cledoc(TOBF), DateFinAff);
    TypeFac := RenvoieTypeFact(TOBF.GetValue('GP_AFFAIREDEVIS'));
    if (TOBF.getString('GP_NATUREPIECEG')='FBT') and (Typefac = 'DAC') And (Cloture) then continue;
    //...suite
    if TobTMP.GetString('AFF_MANDATAIRE')<>'' then
    begin
      if TOBF.GetValue('FACTHTTRAIT') <> 0 then
        HTX := HTX + TOBF.GetDouble('FACTHTTRAIT')
      else
        HTX := HTX + TOBF.GetDouble('FACTUREHT');
    end
    else
      HTX := HTX + TOBF.GetDouble('FACTUREHT');
    if TobTMP.GetString('AFF_MANDATAIRE')<>'' then
    begin
      if TOBF.GetValue('FACTTTCTRAIT') <> 0 then
        TTC := TTC + TOBF.GetDouble('FACTTTCTRAIT')
      else
        TTC := TTC + TOBF.GetValue('FACTURE');
    end
    else
      TTC := TTC + TOBF.GetDouble('FACTURE');
  end;

  FreeAndNil (TobFact);

  TOBTMP.PutValue('FACTUREHT', Arrondi(HTX,V_PGI.OkDecV));
  TOBTMP.PutValue('FACTURE'  , Arrondi(TTC,V_PGI.OkDecV));

  // Calcul approximatif du REGLE HT
  if HTX <> 0 then
     Coef := TTC / HTX
  else
     Coef := 1;

//Récupération facturé détaillé
	if (OptDetailFAC in OptionChoixPrevuAvanc) then
    begin
    // Clause Where à transmettre à la fonction pour prise en compte
    // des natures de pièces à traiter et des dates
  	WherePiece := WherePiece + ' AND GP_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
  	WherePiece := WherePiece + ' AND GP_DATEPIECE BETWEEN "' + UsDateTime(DateMvtDeb) + '"';
  	WherePiece := WherePiece + ' AND "' + UsDateTime(DateMvtFin) + '" ';

  	SetPrevuAvance (TOBTMP,'FBT',OptionChoixPrevuAvanc,WherePiece);
  end;

//Recuperation du reglé
  if CSelReglement.Checked then
  Begin
    TTC := GetReglementAffaire(TOBTMP.GetValue('BCO_AFFAIRE'));
    TOBTMP.PutValue('REGLE', TTC);
    if Coef = 0.0 then HTX := 0.0
    else HTX := TTC / Coef;
    TOBTMP.PutValue('REGLEHT', HTX);
  end;

(* ----
      // Récupération du prévu et de l'avancé
      Req := 'SELECT SUM(GL_QTEFACT*GL_DPA) AS ACHAT, SUM(GL_QTEFACT*GL_DPR) AS REVIENT, SUM(GL_QTEFACT*GL_PUHTNETDEV) AS VENTE, ' +
        'SUM(GL_QTEPREVAVANC*GL_DPA) AS AVANCEPA, SUM(GL_QTEPREVAVANC*GL_DPR) AS AVANCEPR, SUM(GL_QTEPREVAVANC*GL_PUHTNETDEV) AS AVANCEPV ' +
        'FROM LIGNE ' +
        'WHERE GL_NATUREPIECEG = "PBT" AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"' +
        'GROUP BY GL_AFFAIRE';
      QQ := OpenSQL(Req, true);
      if not QQ.EOF then
      begin
        Montant := QQ.findfield('ACHAT').AsFloat;
        TOBTMP.PutValue('PREVUPA', Montant);
        Montant := QQ.findfield('REVIENT').AsFloat;
        TOBTMP.PutValue('PREVUPR', Montant);
        Montant := QQ.findfield('VENTE').AsFloat;
        TOBTMP.PutValue('PREVUPV', Montant);
        Montant := QQ.findfield('AVANCEPA').AsFloat;
        TOBTMP.PutValue('AVANCEPA', Montant);
        Montant := QQ.findfield('AVANCEPR').AsFloat;
        TOBTMP.PutValue('AVANCEPR', Montant);
        Montant := QQ.findfield('AVANCEPV').AsFloat;
        TOBTMP.PutValue('AVANCEPV', Montant);
      end;
      ferme(QQ);

      // Récupération du total des heures prévues et avancées
      Req := 'SELECT SUM(GL_QTEFACT) AS TPS_PREVU, SUM(GL_QTEPREVAVANC) AS TPS_AVANCE ' +
        'FROM LIGNE ' +
        'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
        'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=GA_NATUREPRES '+
        'WHERE GL_NATUREPIECEG = "PBT" AND BNP_TYPERESSOURCE="SAL" AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"' +
        ' GROUP BY GL_AFFAIRE';
      QQ := OpenSQL(Req, true);
      if not QQ.EOF then
      begin
        Montant := QQ.findfield('TPS_PREVU').AsFloat;
        TOBTMP.PutValue('TPS_PREVU', Montant);
        Montant := QQ.findfield('TPS_AVANCE').AsFloat;
        TOBTMP.PutValue('TPS_AVANCE', Montant);
      end;
      ferme(QQ);
*)
      //
    //Req := 'SELECT SUM(GP_TOTALHTDEV) AS FACTUREHT, SUM(GP_TOTALTTCDEV) AS FACTURE, SUM(GP_ACOMPTE) AS REGLE ' +
(*
     Req := 'SELECT SUM(BST_MONTANTREGL) AS REGLE ' +
            'FROM BSITUATIONS ' +
            'WHERE (BST_NATUREPIECE = "FBT" OR BST_NATUREPIECE = "ABT") AND BST_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
     QQ := OpenSQL(Req, true);
     if not QQ.EOF then
        begin
        TTC := QQ.findfield('REGLE').AsFloat;
        TOBTMP.PutValue('REGLE', TTC);
        HTX := TTC / Coef;
        TOBTMP.PutValue('REGLEHT', HTX);
        end;
     ferme(QQ);
*)
//
    //end;

end;


Procedure TOF_BTTABLEAUBORD.ChargeAchatFrais(TOBTMP : TOB; CodeAffaire : String);
Var Req : string;
    QQ  : TQuery;
begin

  //Recherche des Frais sur Facture d'Achat
  Req := 'SELECT SUM(GPT_TOTALHT) AS FRAISPA FROM PIEDPORT LEFT JOIN PIECE ' +
         'ON GPT_NATUREPIECEG=GP_NATUREPIECEG AND GPT_SOUCHE=GP_SOUCHE AND GPT_NUMERO=GP_NUMERO ' +
         'WHERE GPT_NATUREPIECEG IN ("FF","AF","AFS")' +
         '  AND GP_AFFAIRE="' + CodeAffaire + '" ';
  QQ  := OpenSQL(Req, true,-1,'',true);

  if not QQ.EOF then
  begin
    TOBTMP.PutValue('FRAISACHATPA', QQ.FindField('FRAISPA').AsFloat);
  end;

  ferme(QQ);

end;

Procedure TOF_BTTABLEAUBORD.ChargeAchatStocks(TOBTMP : TOB; CodeAffaire : String);
var Req : String;
    QQ  : TQuery;
    Ind : Integer;
    TobEclate     : TOB;
    Tobl          : TOB;
    //
    StockPA   : Double;
    StockPR   : Double;
    StockPV   : Double;
    //
    AchatPA   : Double;
    AchatPR   : Double;
    AchatPV   : Double;
begin

 // BRL 14/08 : lecture des lignes de fournitures et cumul direct dans la requête
  Req := 'SELECT '+
           'SUM(GL_MONTANTPR)  AS BUDGETPR,'+
           'SUM(GL_MONTANTPA)  AS BUDGETPA,'+
           'SUM(GL_MONTANTHTDEV) AS BUDGETPV '+
           'FROM LIGNE '+
           'WHERE GL_NATUREPIECEG = "PBT" AND GL_TYPEARTICLE IN ("MAR","ARP") AND GL_AFFAIRE="' + CodeAffaire + '" ';

  QQ := OpenSQL(Req, False);

  If Not QQ.eof then
  begin
    TOBTMP.PutValue('PREVUACHATPA', QQ.FindField('BUDGETPA').AsFloat);
    TOBTMP.PutValue('PREVUACHATPR', QQ.FindField('BUDGETPR').AsFloat);
    TOBTMP.PutValue('PREVUACHATPV', QQ.FindField('BUDGETPV').AsFloat);
  end;

  Ferme(QQ);

 {
  //chargement du Prévu Achat
  Req := 'SELECT GA_FAMILLENIV1 AS FAMILLE,'+
           'GL_CODEARTICLE     AS CODEARTICLE,'+
           'GL_LIBELLE         AS LIBELLEART,'+
           'GL_QTEFACT         AS BUDGETQTE,'+
           'GL_QTEFACT*GL_DPR  AS BUDGETPR,'+
           'GL_QTEFACT*GL_DPA  AS BUDGETPA,'+
           'GL_QTEFACT*GL_PUHT AS BUDGETPV,'+
           'GL_AFFAIRE         AS CODEAFFAIRE, '+
           'GL_DPR             AS REVIENT, '+
           'GL_DPA             AS ACHAT, '+
           'GL_PUHT            AS VENTE '+
           'FROM LIGNE LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
           'WHERE GL_NATUREPIECEG = "PBT" AND GL_TYPEARTICLE IN ("MAR","ARP") AND GL_AFFAIRE="' + CodeAffaire + '" '+
           'ORDER BY FAMILLE';

  QQ := OpenSQL(Req, False);

  If Not QQ.eof then
  begin
    TOBEclate := Tob.Create('PREVU', nil, -1);
    TobEclate.LoadDetailDB('', '','',QQ, False);
    For ind := 0 to TOBEclate.detail.count - 1 do
    begin
      Tobl    := TobEclate.Detail[Ind];
      TypeRessource := TOBL.GetSTring('TYPERESSOURCE');
      NaturePresta  := TOBL.GetSTring('NATUREPRESTATION');
      if (TypeRessource = 'FAC') OR (TypeRessource = 'ST')  OR (TypeRessource = 'MAT') OR
         (TypeRessource = 'AUT') Or (TypeRessource = 'LOC') OR (TypeRessource = 'INT') OR
         (TypeRessource = 'OUT') Or (TypeRessource = 'SAL') Then
         Continue
      Else
      Begin
        if TOBL.GetSTring('NATUREPRESTATION') <> 'FRA' then //Fournitures
        begin
          MontantPA := MontantPA + TOBL.GetDouble('BUDGETPA');
          MontantPR := MontantPR + TOBL.GetDouble('BUDGETPR');
          MontantHT := MontantHT + TOBL.GetDouble('BUDGETPV');
        end;
      end;
    end;
    TOBTMP.PutValue('PREVUACHATPA', MontantPA);
    TOBTMP.PutValue('PREVUACHATPR', MontantPR);
    TOBTMP.PutValue('PREVUACHATPV', MontantPV);
  	if CEclatNatPrest.checked then
		Begin
      Repartition_Eclatement('PREVUACHAT', '_FOU_'+ NaturePresta, '_PA', MontantPA, TOBTMP);
      Repartition_Eclatement('PREVUACHAT', '_FOU_'+ NaturePresta, '_PR', MontantPR, TOBTMP);
      Repartition_Eclatement('PREVUACHAT', '_FOU_'+ NaturePresta, '_PV', MontantPV, TOBTMP);
		end;
  end;

  Ferme (QQ);
 }
 //BRL : 14/08

  //chargement du Réalisé Stock & Achat
  Req := 'SELECT '+
         '"STOCKPR" = CASE WHEN (GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3) = "PBT") THEN BCO_MONTANTPR ELSE 0 END, '+
         '"STOCKPA" = CASE WHEN (GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3) = "PBT") THEN BCO_MONTANTACH ELSE 0 END, '+
         '"STOCKPV" = CASE WHEN (GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3) = "PBT") THEN BCO_MONTANTHT ELSE 0 END, '+
         '"ACHATPR" = CASE WHEN NOT (GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3) = "PBT") THEN BCO_MONTANTPR  ELSE 0 END, ' +
         '"ACHATPA" = CASE WHEN NOT (GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3) = "PBT") THEN BCO_MONTANTACH ELSE 0 END, ' +
         '"ACHATPV" = CASE WHEN NOT (GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3) = "PBT") THEN BCO_MONTANTHT  ELSE 0 END, ' +
         'GA_FAMILLENIV1 AS FAMILLE, BCO_ARTICLE AS CODEARTICLE, BCO_LIBELLE AS LIBELLEART, BCO_AFFAIRE AS CODEAFFAIRE ' +
         'FROM CONSOMMATIONS LEFT JOIN LIGNE ' +
         '  ON GL_NATUREPIECEG= BCO_NATUREPIECEG ' +
         ' AND GL_SOUCHE= BCO_SOUCHE ' +
         ' AND GL_NUMERO= BCO_NUMERO ' +
         ' AND GL_NUMORDRE= BCO_NUMORDRE '+
         'LEFT JOIN ARTICLE AA ON GA_ARTICLE=BCO_ARTICLE WHERE BCO_DATEMOUV BETWEEN "' + UsdateTime(DateMvtDeb) + '" AND "' + UsdateTime(DateMvtFin) + '" ' +
         //FV1 - 16/05/2017 : FS#2541 - TREUIL - informations quant au champ REALISESTOCKPA
         //' AND BCO_NATUREPIECEG IN ("LBT", "", "BFC","AFS","AF","FAC") AND BCO_NATUREMOUV="FOU" AND BCO_AFFAIRE="' + CodeAffaire + '" ';
         ' AND BCO_NATUREPIECEG IN ("LBT","","BFC","AFS","FAC") AND BCO_NATUREMOUV="FOU" AND BCO_AFFAIRE="' + CodeAffaire + '" ';

  QQ := OpenSQL(Req, False);

  If Not QQ.eof then
  begin
     TOBEclate := Tob.Create('REALISE', nil, -1);
     TobEclate.LoadDetailDB('', '','',QQ, False);
     For ind := 0 to TOBEclate.detail.count - 1 do
     begin
       Tobl    := TobEclate.Detail[Ind];
       //
       StockPA := TOBTMP.GetDouble('REALISESTOCKPA');
       StockPR := TOBTMP.GetDouble('REALISESTOCKPR');
       StockPV := TOBTMP.GetDouble('REALISESTOCKPV');
       //
       AchatPA := TOBTMP.GetDouble('REALISEACHATPA');
       AchatPR := TOBTMP.GetDouble('REALISEACHATPR');
       AchatPV := TOBTMP.GetDouble('REALISEACHATPV');
       //
       TOBTMP.PutValue('REALISESTOCKPA', StockPA + TOBL.GetDouble('STOCKPA'));
       TOBTMP.PutValue('REALISESTOCKPR', StockPR + TOBL.GetDouble('STOCKPR'));
       TOBTMP.PutValue('REALISESTOCKPV', StockPV + TOBL.GetDouble('STOCKPV'));
       //
       TOBTMP.PutValue('REALISEACHATPA', AchatPA + TOBL.GetDouble('ACHATPA'));
       TOBTMP.PutValue('REALISEACHATPR', AchatPR + TOBL.GetDouble('ACHATPR'));
       TOBTMP.PutValue('REALISEACHATPV', AchatPV + TOBL.GetDouble('ACHATPV'));
       //
       if CEclatNatPrest.checked then
	     Begin
        Repartition_Eclatement('REALISESTOCK', '_FOU_'+ TOBL.GetString('FAMILLE'), '_PA', TOBL.GetDouble('STOCKPA'), TOBTMP);
        Repartition_Eclatement('REALISESTOCK', '_FOU_'+ TOBL.GetString('FAMILLE'), '_PR', TOBL.GetDouble('STOCKPR'), TOBTMP);
        Repartition_Eclatement('REALISESTOCK', '_FOU_'+ TOBL.GetString('FAMILLE'), '_PV', TOBL.GetDouble('STOCKPV'), TOBTMP);
        //
        Repartition_Eclatement('REALISEACHAT', '_FOU_'+ TOBL.GetString('FAMILLE'), '_PA', TOBL.GetDouble('ACHATPA'), TOBTMP);
        Repartition_Eclatement('REALISEACHAT', '_FOU_'+ TOBL.GetString('FAMILLE'), '_PR', TOBL.GetDouble('ACHATPR'), TOBTMP);
        Repartition_Eclatement('REALISEACHAT', '_FOU_'+ TOBL.GetString('FAMILLE'), '_PV', TOBL.GetDouble('ACHATPV'), TOBTMP);
		   end;
     end;
     FreeAndNil(TobEclate);
  end;

  Ferme(QQ);

end;

// Modified by f.vautrain 29/11/2017 16:49:48 - GUINIER - Ajout des Totaux de Proratas frais et ports
Procedure TOF_BTTABLEAUBORD.ChargeProrata(TOBTMP : TOB; CodeAffaire : String);
var Req : String;
    QQ  : TQuery;
    Ind : Integer;
begin

  If Prorata = '' Then Exit;

  Req := 'SELECT SUM(GPT_TOTALHTDEV) AS TOTALPRORATA FROM PIECE ';
  Req := Req + 'LEFT JOIN PIEDPORT ON GPT_NATUREPIECEG=GP_NATUREPIECEG AND GPT_SOUCHE=GP_SOUCHE AND GPT_NUMERO=GP_NUMERO AND GPT_INDICEG=GP_INDICEG ';
  Req := Req + 'WHERE GP_NATUREPIECEG="FBT" AND GP_AFFAIRE="' + CodeAffaire + '" AND GPT_FRAISREPARTIS="-" AND GPT_CODEPORT IN ' + Prorata;
  Req := Req + '  AND GP_DATEPIECE BETWEEN "' + UsDateTime(DateDeb) + '"';
  Req := Req + '  AND "' + UsDateTime(DateFin) + '" ';

  QQ := OpenSQL(Req, true,-1,'',true);
  if not QQ.eof then
    TOBTMP.PutValue('PRORATA', QQ.FindField('TOTALPRORATA').AsFloat)
  else
    TOBTMP.PutValue('PRORATA', 0.00);

  ferme(QQ);

end;

// Modified by f.vautrain 04/12/2017 16:53:55 - GUINIER - Ajout des totaux des Révisions Frais & Ports
Procedure TOF_BTTABLEAUBORD.ChargeRevision(TOBTMP : TOB; CodeAffaire : String);
var Req : String;
    QQ  : TQuery;
    Ind : Integer;
begin

  If Revision = '' Then Exit;

  Req := 'SELECT SUM(GPT_TOTALHTDEV) AS TOTALREVISION FROM PIECE ';
  Req := Req + 'LEFT JOIN PIEDPORT ON GPT_NATUREPIECEG=GP_NATUREPIECEG AND GPT_SOUCHE=GP_SOUCHE AND GPT_NUMERO=GP_NUMERO AND GPT_INDICEG=GP_INDICEG ';
  Req := Req + 'WHERE GPT_NATUREPIECEG="FBT" AND GP_AFFAIRE="' + CodeAffaire + '" AND GPT_FRAISREPARTIS="-" AND GPT_CODEPORT IN ' + Revision;
  Req := Req + '  AND GP_DATEPIECE BETWEEN "' + UsDateTime(DateDeb) + '"';
  Req := Req + '  AND "' + UsDateTime(DateFin) + '" ';

  QQ := OpenSQL(Req, true,-1,'',true);
  if not QQ.eof then
    TOBTMP.PutValue('REVISION', QQ.FindField('TOTALREVISION').AsFloat)
  else
    TOBTMP.PutValue('REVISION', 0.00);

  ferme(QQ);

end;

// Modified by f.vautrain 04/12/2017 16:53:32 - GUINIER - Ajouts des Factures Hors Devis
Procedure TOF_BTTABLEAUBORD.ChargeHorsDevis(TOBTMP : TOB; CodeAffaire : String);
var Req : String;
    QQ  : TQuery;
    Ind : Integer;
begin

  Req := 'SELECT SUM(GL_QTEFACT * GL_PUHTDEV) AS TOTALFACTHD FROM LIGNE';
  Req := Req + ' WHERE GL_NATUREPIECEG = "FBT" ';
  Req := Req + '   AND GL_PIECEORIGINE = ""';
  if CodeAffaire <> '' then  Req := Req + '   AND GL_AFFAIRE = "' + CodeAffaire + '" ';
  Req := Req + '   AND GL_DATEPIECE BETWEEN "' + UsDateTime(DateDeb) + '"';
  Req := Req + '   AND "' + UsDateTime(DateFin) + '" ';

  QQ := OpenSQL(Req, true,-1,'',true);
  if not QQ.eof then
  begin
    TOBTMP.PutValue('FACTHORSDEVIS', QQ.findfield('TOTALFACTHD').AsFloat);
  end;

  Ferme(QQ);

end;

// Modified by f.vautrain 04/12/2017 16:53:04 - GUINIER - Ajouts des totaux des Avoirs
Procedure TOF_BTTABLEAUBORD.ChargeTotalAvoir(TOBTMP : TOB; CodeAffaire : String);
var Req : String;
    QQ  : TQuery;
    Ind : Integer;
begin

  Req := 'SELECT SUM(GL_QTEFACT * GL_PUHTDEV) AS TOTALAVOIR FROM LIGNE';
  Req := Req + ' WHERE GL_NATUREPIECEG = "ABT" ';
  Req := Req + '   AND GL_AFFAIRE = "' + CodeAffaire + '" ';
  Req := Req + '   AND GL_DATEPIECE BETWEEN "' + UsDateTime(DateDeb) + '"';
  Req := Req + '   AND "' + UsDateTime(DateFin) + '" ';

  QQ := OpenSQL(Req, true,-1,'',true);
  if not QQ.eof then
  begin
    TOBTMP.PutValue('TOTALAVOIR', QQ.findfield('TOTALAVOIR').AsFloat);
  end;

  Ferme(QQ);

end;

procedure TOF_BTTABLEAUBORD.ChargeLeResteADepenser(TOBTMP, TOBEclate : TOB; CodeAffaire : String);
var Req : String;
    QQ  : TQuery;
    Ind : Integer;
    GestionEnPa       : Boolean;
    //
    TheYear       : Word;
    TheMonth      : Word;
    TheDay        : Word;
    //
    MtFinAff      : Double;
    //
    Tobl          : TOB;
begin

  //Chargement des ligne de reste à dépenser (???)
  Req := 'SELECT DISTINCT(rad_datearretee) FROM BTRESTEADEP WHERE RAD_AFFAIRE="' + CodeAffaire +
        '" AND rad_datearretee <= "' + UsDateTime(DateMvtFin) + '" order by rad_datearretee desc';
  QQ := OpenSQL(Req, true,-1,'',true);
  if not QQ.eof then
  begin
   DecodeDate( QQ.FindField('RAD_DATEARRETEE').AsDateTime,TheYear, TheMonth, TheDay);
   TOBTMP.PutValue('RESTEADEP_DATEARRETEE', QQ.FindField('RAD_DATEARRETEE').AsDateTime);


   Req := 'SELECT RAD_DATEARRETEE, RAD_QTERESTE AS QTRESTE, RAD_MTRESTE AS MTRESTE, ' +
          ' RAD_TYPERESSOURCE AS TYPERESSOURCE, ' +
          ' RAD_NATUREPRES AS NATUREPRESTATION, ' +
          ' GA_FAMILLENIV1 AS FAMILLEART ' +
          ' FROM BTRESTEADEP ' +
          ' LEFT JOIN ARTICLE ON GA_ARTICLE=RAD_ARTICLE ' +
          'WHERE RAD_AFFAIRE = "' + CodeAffaire + '" ' +
          '  AND RAD_ANNEE = "' + IntToStr(TheYear)  + '" ' +
          '  AND RAD_MOIS = "'  + IntToStr(TheMonth) + '" ' +
          'ORDER BY TYPERESSOURCE, NATUREPRESTATION';

   Ferme(QQ);

   QQ := OpenSQL(Req, true,-1,'',true);
   //
   if not QQ.eof then
   begin
     TOBEclate := Tob.Create('ADEPENSER', nil, -1);
     TobEclate.LoadDetailDB('', '','',QQ, False);
     For ind := 0 to TOBEclate.detail.count - 1 do
     begin
       Tobl    := TobEclate.Detail[Ind];
       //répartition par prestation et nature de prestation
       Traitement_Repartition_RAD(TOBL, TOBTMP);
       //
     end;

     GestionEnPa := GetParamSocSecur('SO_BTRADENPA',false);

     if GestionEnPa then
     begin
       MtFinAff   := TOBTMP.GetDouble('REALISETOTALPA') + TOBTMP.GetDouble('RESTEADEP_MTRESTE');
       TOBTMP.PutValue('RESTEADEP_FINAFF', MtFinAff);
       //
       TOBTMP.PutValue('RESTADEP_AUT_FINAFF', TOBTMP.GetDouble('REALISE_AUTRE_PA')    + TOBTMP.GetDouble('RESTADEP_AUT_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_SAL_FINAFF', TOBTMP.GetDouble('REALISE_SAL_PA')      + TOBTMP.GetDouble('RESTADEP_SAL_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_LOC_FINAFF', TOBTMP.GetDouble('REALISE_LOC_PA')      + TOBTMP.GetDouble('RESTADEP_LOC_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_OUT_FINAFF', TOBTMP.GetDouble('REALISE_OUTIL_PA')    + TOBTMP.GetDouble('RESTADEP_OUT_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_ST_FINAFF' , TOBTMP.GetDouble('REALISE_ST_PA')       + TOBTMP.GetDouble('RESTADEP_ST_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_MAT_FINAFF', TOBTMP.GetDouble('REALISE_MAT_PA')      + TOBTMP.GetDouble('RESTADEP_MAT_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_INT_FINAFF', TOBTMP.GetDouble('REALISE_INTERIM_PA')  + TOBTMP.GetDouble('RESTADEP_INT_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_FOU_FINAFF', TOBTMP.GetDouble('REALISEFOURNITURESPA')+ TOBTMP.GetDouble('RESTADEP_FOU_MTRESTE'));
       TOBTMP.PutValue('RESTADEP_FAC_FINAFF', TOBTMP.GetDouble('FACTUREHT')           + TOBTMP.GetDouble('RESTADEP_FAC_MTRESTE'));
     end else
     begin
     MtFinAff   := TOBTMP.GetDouble('REALISETOTALPR') + TOBTMP.GetDouble('RESTEADEP_MTRESTE');
     TOBTMP.PutValue('RESTEADEP_FINAFF', MtFinAff);
     //
     TOBTMP.PutValue('RESTADEP_AUT_FINAFF', TOBTMP.GetDouble('REALISE_AUTRE_PR')    + TOBTMP.GetDouble('RESTADEP_AUT_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_SAL_FINAFF', TOBTMP.GetDouble('REALISE_SAL_PR')      + TOBTMP.GetDouble('RESTADEP_SAL_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_LOC_FINAFF', TOBTMP.GetDouble('REALISE_LOC_PR')      + TOBTMP.GetDouble('RESTADEP_LOC_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_OUT_FINAFF', TOBTMP.GetDouble('REALISE_OUTIL_PR')    + TOBTMP.GetDouble('RESTADEP_OUT_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_ST_FINAFF' , TOBTMP.GetDouble('REALISE_ST_PR')       + TOBTMP.GetDouble('RESTADEP_ST_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_MAT_FINAFF', TOBTMP.GetDouble('REALISE_MAT_PR')      + TOBTMP.GetDouble('RESTADEP_MAT_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_INT_FINAFF', TOBTMP.GetDouble('REALISE_INTERIM_PR')  + TOBTMP.GetDouble('RESTADEP_INT_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_FOU_FINAFF', TOBTMP.GetDouble('REALISEFOURNITURESPR')+ TOBTMP.GetDouble('RESTADEP_FOU_MTRESTE'));
     TOBTMP.PutValue('RESTADEP_FAC_FINAFF', TOBTMP.GetDouble('FACTUREHT')           + TOBTMP.GetDouble('RESTADEP_FAC_MTRESTE'));
     end;
     //
     FreeAndNil(TobEclate);
   end;
  end;

  Ferme(QQ);

end;

//FV1 : 24/08/2017 - FS#2667 - ARKAL : Nouvelle présentation de tableau d'analyse chantier
procedure TOF_BTTABLEAUBORD.ChargelesCommandesNonLivrees(TOBTMP : TOB; CodeAffaire : String);
Var STSQL : String;
    QQ    : TQuery;
begin

  StSQL := 'SELECT SUM(GL_QTERESTE * GL_PUHTDEV) AS CDENONLIVREE FROM LIGNE';
  StSQL := StSQL + ' WHERE GL_NATUREPIECEG = "CF" ';
  StSQL := StSQL + '   AND GL_QTERESTE <> 0';
  StSQL := StSQL + '   AND GL_AFFAIRE = "' + CodeAffaire + '" ';

  QQ := OpenSQL(StSQL, true,-1,'',true);
  if not QQ.eof then
  begin
    TOBTMP.PutValue('CDENONLIVREE', QQ.findfield('CDENONLIVREE').AsFloat);
  end;
  Ferme(QQ);

end;

procedure TOF_BTTABLEAUBORD.ChargelesCommandesLivrees(TOBTMP : TOB; CodeAffaire : String);
Var STSQL : String;
    QQ    : TQuery;
begin

  StSQL := 'SELECT SUM(GL_QTEFACT * GL_PUHTDEV) AS CDELIVREE FROM LIGNE';
  StSQL := StSQL + ' WHERE GL_NATUREPIECEG = "CF" ';
  StSQL := StSQL + '   AND GL_QTERESTE = 0';
  StSQL := StSQL + '   AND GL_AFFAIRE = "' + CodeAffaire + '" ';

  QQ := OpenSQL(StSQL, true,-1,'',true);
  if not QQ.eof then
  begin
    TOBTMP.PutValue('CDELIVREE', QQ.findfield('CDELIVREE').AsFloat);
  end;
  Ferme(QQ);

end;

procedure TOF_BTTABLEAUBORD.ChargelesFacturesansCommandes(TOBTMP: TOB; CodeAffaire : String);
Var STSQL : String;
    QQ    : TQuery;
begin

  StSQL := 'SELECT SUM(GL_QTEFACT * GL_PUHTDEV) AS FACTSANSCDE FROM LIGNE';
  StSQL := StSQL + ' WHERE GL_NATUREPIECEG = "FF" ';
  StSQL := StSQL + '   AND GL_PIECEORIGINE NOT LIKE "%CF%"';
  StSQL := StSQL + '   AND GL_AFFAIRE = "' + CodeAffaire + '" ';

  QQ := OpenSQL(StSQL, true,-1,'',true);
  if not QQ.eof then
  begin
    TOBTMP.PutValue('FACTSANSCDE', QQ.findfield('FACTSANSCDE').AsFloat);
  end;
  Ferme(QQ);

end;

{*
Function TOF_BTTABLEAUBORD.PrevueVide(TOBL : TOB) : boolean;
begin
    result := (TOBL.GetValue('PREVUPA')= 0.0) and
    					(TOBL.GetValue('PREVUPR')=  0.0) and
    					(TOBL.GetValue('PREVUPV')= 0.0) and
    					(TOBL.GetValue('TPS_PREVU')= 0.0);
end;

Function TOF_BTTABLEAUBORD.RealiseVide(TOBL : TOB) : boolean;
begin
    result := (TOBL.GetValue('REALISETOTALPA')= 0.0) and
    					(TOBL.GetValue('REALISETOTALPR')=  0.0) and
    					(TOBL.GetValue('REALISETOTALPV')= 0.0) and
    					(TOBL.GetValue('TPS_REALISE')= 0.0);
end;

function TOF_BTTABLEAUBORD.AvanceVide(TOBL : TOB) : boolean;
begin
   result := (TOBL.GetValue('AVANCEPA')= 0.0) and
    					(TOBL.GetValue('AVANCEPR')=  0.0) and
    					(TOBL.GetValue('AVANCEPV')= 0.0) and
    					(TOBL.GetValue('TPS_AVANCE')= 0.0);
end;

function TOF_BTTABLEAUBORD.FactureVide(TOBL : TOB) : boolean;
begin
    result := (TOBL.GetValue('FACTUREHT')= 0.0) and
    					(TOBL.GetValue('FACTURE')=  0.0) and
    					(TOBL.GetValue('REGLE')= 0.0) and
    					(TOBL.GetValue('REGLEHT')= 0.0);
end;

function TOF_BTTABLEAUBORD.EngageVide(TOBL : TOB) : boolean;
begin
    result := (TOBL.GetValue('ENGAGEFOU')= 0.0) and
    					(TOBL.GetValue('ENGAGECHA')=  0.0) and
    					(TOBL.GetValue('FACTNONPAR')= 0.0);
end;

Function  TOF_BTTABLEAUBORD.LigneVide (TOBL : TOB) : boolean;
begin
  	result := false;
  	if PrevueVide(TOBL) and RealiseVide(TOBL) and
       AvanceVide(TOBL) and FactureVide(TOBL) and
       EngageVide(TOBL) and (TOBL.GetValue('COURECANN')=0.0) then result:= true;
end;
*}

//var Indice : integer;
//		TOBTMP : TOB;
//begin
//  Indice := 0;
//  repeat
//  	TOBTMP := TOBBTB.Detail[indice];
//    if LigneVide(TOBTMP) then TOBTMP.free else Inc(Indice);
//  until Indice >= TOBBTB.detail.count;
//end;

procedure TOF_BTTABLEAUBORD.CChoixClick(Sender: TObject);
begin

  THAffaire0.text := '';

  	if CCumulAffaire.checked then
  	begin
      T_Affaire.Caption := 'Chantier associé';
      Exit;
  	end;

   if (CChantier.checked) and (not CContrat.Checked ) And (not CAppel.Checked ) then
      Begin
      T_Affaire.Caption := 'Chantier';
      THAffaire0.text := 'A';
      end
   Else if (CChantier.checked) and (CContrat.Checked) and (CAppel.Checked) then
      begin
      T_Affaire.Caption := 'Affaire';
      end
   Else if (CChantier.checked) and (not CContrat.Checked) and (CAppel.Checked) then
      begin
      T_Affaire.Caption := 'Affaire';
      end
   Else if (CChantier.checked) and (CContrat.Checked) and (not CAppel.Checked) then
      begin
      T_Affaire.Caption := 'Affaire';
      end
   Else if (not CChantier.checked) and (not CContrat.Checked) And (not CAppel.Checked) then
      begin
      T_Affaire.Caption := 'Affaire';
      end
   Else if (not CChantier.checked) and (CContrat.Checked) And (CAppel.Checked) then
      begin
      T_Affaire.Caption := 'Affaire';
      end
   Else if (not CChantier.checked) and (CContrat.Checked) And (not CAppel.Checked) then
      Begin
      T_Affaire.Caption := 'Contrat';
      THAffaire0.text := 'I';
      end
   Else if (not CChantier.checked) and (not CContrat.Checked) And (CAppel.Checked) then
      Begin
      T_Affaire.Caption := 'Appel';
      THAffaire0.text := 'W';
      end
   Else
      Begin
      T_Affaire.Caption := 'Chantier';
      THAffaire0.text := 'A';
      end;
  //
  //ChargeCleAffaire(THAffaire0, THAffaire1, THAffaire2, THAffaire3, THAvenant, nil, TaModif, THAffaire.Text, false);
  //

end;

Function TOF_BTTABLEAUBORD.ChargeAffaireRef(CodeAffaire, AffaireSaisie : String) : String;
Var QQ  : Tquery;
    Req : String;
Begin

  Result := '';

  If Copy(CodeAffaire, 1, 1) =  'A' then
     Begin
     Result := CodeAffaire;
     exit;
     end;

  //Récupération du code affaire dans les lignes d'appel
  Req := 'SELECT AFF_CHANTIER FROM AFFAIRE WHERE AFF_AFFAIRE = "' + CodeAffaire + '"';

  QQ := OpenSQL(Req, true,-1,'',true);
  if QQ.EOF then
     begin
     ferme(QQ);
     exit;
     end;

  Result := QQ.findfield('AFF_CHANTIER').AsString;

  ferme(QQ);

  if AffaireSaisie = '' then exit;

  if Result = '' then
     Begin
     Req := 'SELECT AFF_CHANTIER FROM AFFAIRE WHERE AFF_AFFAIRE = "' + AffaireSaisie + '"';
     QQ := OpenSQL(Req, true,-1,'',true);
     if QQ.eof then
        Begin
        ferme(QQ);
        exit;
        end;
     Result := QQ.findfield('AFF_CHANTIER').AsString;
     ferme(QQ);
     end;

end;

procedure TOF_BTTABLEAUBORD.CCumulAffaireClick(Sender: TObject);
begin
	CChoixClick(Self);
end;

procedure TOF_BTTABLEAUBORD.DateTransfoOK (TOBEclate : TOB; var Montant : Double);
var Q : Tquery;
    DateTransfo : TDateTime ;
    req : String;
begin
  Montant := 0;
  req := 'SELECT BCO_DATEMOUV, BCO_MONTANTACH FROM CONSOMMATIONS WHERE BCO_QUANTITE<>0 and BCO_LIENTRANSFORME='+FloatToStr(TOBEclate.GetValue('BCO_NUMMOUV'));
  Q := OpenSQL(req, True);
  while Not Q.EOF do
  begin
	  DateTransfo := StrToDate(Q.FindField('BCO_DATEMOUV').AsString);
    if not ((DateTransfo >= DateMvtDeb) and (DateTransfo <= DateMvtFin)) then
       Montant := Montant + Q.FindField('BCO_MONTANTACH').AsFloat;
    Q.Next;
  end;
	Ferme (Q);
end;

procedure TOF_BTTABLEAUBORD.DateFactureOK (Nummouv : Double; var Montant : Double);
var Q : Tquery;
    DateTransfo : TDateTime ;
    req : String;
    Mt : Double;
begin
  Montant := 0;
  req := 'SELECT BCO_DATEMOUV, BCO_MONTANTACH, BCO_NATUREPIECEG, BCO_NUMMOUV, BCO_TRANSFORME FROM CONSOMMATIONS WHERE BCO_LIENTRANSFORME='+FloatToStr(Nummouv);
  Q := OpenSQL(req, True);
  while Not Q.EOF do
  begin
    if (Q.FindField('BCO_NATUREPIECEG').AsString <> 'FF') then
    begin
      if (Q.FindField('BCO_TRANSFORME').AsString = '-') then
        Montant := Montant + Q.FindField('BCO_MONTANTACH').AsFloat
      else
      begin
        Mt := 0.0;
        DateFactureOK (Q.FindField('BCO_NUMMOUV').AsFloat, Mt);
        Montant := Montant + Mt;
      end;
      Q.Next;
    end else
    begin
      if (Q.FindField('BCO_TRANSFORME').AsString = 'X') then
      begin
  	    DateTransfo := StrToDate(Q.FindField('BCO_DATEMOUV').AsString);
        if not ((DateTransfo >= DateMvtDeb) and (DateTransfo <= DateMvtFin)) then
          Montant := Montant + Q.FindField('BCO_MONTANTACH').AsFloat;
      end;
      Q.Next;
    end;
  end;
	Ferme (Q);
end;

initialization
  registerclasses([TOF_BTTABLEAUBORD]);
end.

