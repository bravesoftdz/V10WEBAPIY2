{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/11/2002
Modifié le ... : 09/02/2004
Description .. : Source TOF de la FICHE : CPGLAUXI ()
Suite ........ : GCO - 09/02/2004
Suite ........ : -> Correction du bug E_QUALIFPIECE pour la justification de
Suite ........ : solde ( on force le type d'écritures à N )
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Passage des paramètres à l'état avec un CritEdt
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Suite ........ : GCO - 05/09/2005 - FQ 16320
Suite ........ :
Mots clefs ... : TOF;CPGLAUXI
*****************************************************************}
unit uTofCPGLAuxi;

interface

uses
  StdCtrls,
  Controls,
  Classes,
  Graphics,
{$IFDEF EAGLCLIENT}
  Maineagl,
  eQRS1,
  uTileAGL,     // LanceEtatTob
{$ELSE}
  Fe_main,
  DB,
{$IFDEF DBXPRESS}uDbxDataSet,{$ELSE}dbtables,{$ENDIF}
  QRS1,
  EdtREtat,     // LanceEtatTob
{$ENDIF}
  Hqry,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  Htb97,
  Ent1,
  HMsgBox,
  TofMeth,
  uLibWindows,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UTOF,
  uTob,         // Tob
  uLibExercice, // CExerciceVersRelatif
  Spin,         // TSpinEdit
  SaisUtil,     // QuelExoDT
  AGLInit,      // TheData
  CritEdt,      // ClassCritEdt
  LicUtil,      // CryptageSt
  uPrintF1Book, // PDFBatch
  HPdfPrev,     // PreviewPDFFile
  HPdfViewer,   // PrintPDF
  Printers,     // Printer
  ED_TOOLS,     // InitMoveProgressForm
  ParamSoc,     // GetParamSocSecur           {FP 25/04/2006 FQ18001}
  UtilEdt;      // MajEditionLegal


type
  TOF_CPGLAUXI = class(TOF_METH)

  private
    EdLegale         : Boolean;
    E_AUXILIAIRE     : THEdit;
    E_AUXILIAIRE_    : THEdit;
    E_GENERAL        : THEdit;
    E_GENERAL_       : THEdit;
    AuxiParGene      : THEdit;
    E_DATECOMPTABLE  : THEdit;
    E_DATECOMPTABLE_ : THEdit;
    E_NUMEROPIECE    : THEdit;
    E_NUMEROPIECE_   : THEdit;
    CPTEXCEPT        : THEdit;
    LIBEQUALIFPIECE  : THEdit;
    CUMULAU          : THEdit;
    E_REFINTERNE     : THEdit;

    // GCO - 24/03/2006
    OKENTETE         : THEdit;
    OKRESUME         : THEdit;
    CBPreview        : TCheckBox;
    // FIN GCO

    E_EXERCICE       : THValComboBox;
    E_DEVISE         : THValComboBox;

    T_NATUREAUXI    : THValComboBox;
    MODESELECT      : THValComboBox;
    E_QUALIFPIECE   : THMultiValComboBox;

    // Correspondances
    CORRESP    : THValComboBox;
    CORRESPDE  : THValComboBox;
    CORRESPA   : THValComboBox;

    // Tables Libres
    TABLELIBRE : THValComboBox;
    LIBREDE    : THEdit;
    LIBREA     : THEdit;

    CBEnSituation  : TCheckBox;
    ECRVALIDE      : TCheckBox;
    ECRLETTRE      : TCheckBox;
    AFFANOUVEAU    : TCheckBox;
    AFFCENTRAECHE  : TCheckBox;
    AFFCUMUL       : TCheckBox;
    LIBANO         : THEdit;
    
    EDateSituation : THEdit;
    DateDebutExo   : THEdit;
    Exercice       : THEdit;

    RGDevise: THRadioGroup;

    Pages: TPageControl;
    RUPTURE: THRadioGroup;
    RUPTURES: TTabSheet;
    RUPTURETYPE: THRadioGroup;
    RUPCORRESP: TGroupBox;
    RUPLIBRES: TGroupBox;
    RUPGROUPES: TGroupBox;

    AFFDEVISE: THEdit;
    DEVPIVOT: THEdit;

    NIVORUPTURE     : TSpinEdit;
    XX_RUPTURE      : THEdit;
    AVECRUPTYPE     : THEdit;
    AVECRUPTURE     : THEdit;
    AVECNIVORUPTURE : THEdit;

    CPTCorrespExist : TCheckBox;
    CPTLibreExist   : TCheckBox;

    CHOIXTYPETALI   : THRadioGroup;

    procedure OnExitE_General       (Sender: TObject);
    procedure OnExitE_Auxiliaire    (Sender: TObject);
    procedure OnChangeT_NatureAuxi  (Sender: TObject);
    procedure OnExitE_DateComptable (Sender: TObject);
    procedure OnChangeE_Exercice    (Sender: TObject);
    procedure OnChangeE_Devise      (Sender: TObject);
    procedure OnClickCBEnSituation  (Sender: TObject);
    procedure OnClickRupture        (Sender: TObject);
    procedure OnClickRuptureType    (Sender: TObject);
    procedure FTimerTimer           (Sender: TObject);
    procedure OnChangeCorresp       (Sender: TObject);
    procedure OnChangeTableLibre    (Sender: TObject);
    procedure OnClickAffCentraEche  (Sender: TObject);
    procedure OnEnterQualifPiece(Sender : TObject);     // SBO 12/10/2004
    procedure OnClickChoixTypeTaLi  (Sender : TObject);
    procedure AuxiElipsisClick      (Sender: TObject) ;

  public
    procedure OnArgument(S: string); override;
    procedure OnNew; override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
    {$IFNDEF GCGC}
    procedure OnChangeFiltre (Sender: TObject); override;
    {$ENDIF}
    procedure ChargementCritEdt; override;
    procedure MySelectFiltre; override;

  private
    FStArgument           : string;
    FStWhereEtat          : string;

   // FStSelectTL           : string;
   // FStLeftJoinTIERSCOMPL : string;

    FExoDate       : TExoDate; // Contexte de travail sur les dates

    FBoAuxiParGene : Boolean;
    FBoOkDecoupe   : Boolean; // Découpage du grand livre

//    function GestionTableLibre        : string;
    function GenereConditionWhereEtat : string;
    function RecupAutresCriteres      : string;
    function RecupCriteresCommuns     : string;
    function RecupCriteresModeSelection( vStWhereEtat : string ) : string;
    function RenvoiGroupByOrderBy     : string;

    // =*= Centralisation des échéances =*=
    procedure GestionCentralisation;       // Point d'éntrée de la centralisation
    procedure SupprimeCentralisation;      // Suppression des écritures "Z"
    procedure EffectueCentralisation;      // Insertion des écritures "Z"
    procedure CBlocageGLGene;              // Blocage dans la Table COURRIER
    procedure CDeBlocageGLGene;            // Déblocage
    function  CEstBloqueGLGene : Boolean ; // Controle de la présence du Verrou
    // =*= Fin =*=

    procedure InitChoixTableLibre;
    procedure ControleAvantEdition;
    procedure GestionCritereDEV;

{$IFDEF EAGLCLIENT}
{$ELSE}
    procedure GestionTailleGrandLivre( vStwhereCritere : string );
{$ENDIF}

    procedure GenereCumulGL;
  end;

function CPLanceFiche_CPGLAUXI        ( vStParam: string = '') : string ;
function CPLanceFiche_CPGLAUXIPARGENE ( vStParam: string = '') : string ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  uLibCalcEdtCompta, // FTobCumulAu
  UtilPgi,           // EstMultiSoc
  uMultiDossierUtil, // RequeteMultiDossier
  UTofMulParamGen;

const cNbEcrMax = 10000;

////////////////////////////////////////////////////////////////////////////////
function CPLanceFiche_CPGLAUXI(vStParam: string = ''): string;
begin
  // Ne jamais toucher au vStParam
  Result := AGLLanceFiche('CP', 'CPGLAUXI', '', '', vStParam);
end;

////////////////////////////////////////////////////////////////////////////////
function CPLanceFiche_CPGLAUXIPARGENE ( vStParam: string = '') : string ;
begin
  // Ne jamais toucher au vStParam
  Result := AGLLanceFiche('CP', 'CPGLAUXIPARGENE', '', '', vStParam);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnArgument(S: string);
var     
    DeviseAff                       : RDevise;
begin
  inherited;
  { FQ 20047 BVE 25.04.07 }
  RGDEVISE         := THRadioGroup(GetControl('RGDEVISE', True));
  DeviseAff.Code := GetParamSoc('SO_DEVISEPRINC');
  GetInfosDevise(DeviseAff);
  if DeviseAff.Libelle = '' then DeviseAff.Libelle := 'Euro';
  RGDEVISE.Items.Strings[0] := DeviseAff.Libelle;
  { END FQ 20047 }

  FStArgument := S; // Récupératrion des paramètres

  // GCO - 11/07/2006 - FQ 18363
  PAvances := TTabSheet(GetControl('AVANCES', true));
  Pavances.TabVisible := True;
  TFQRS1(Ecran).CritAvancesVisibled := True;

  if FStArgument = 'LEGALE' then EdLegale := true;

  if ReadTokenSt(FStArgument) = 'MULTIDOSSIER' then
  begin
    SetControlVisible( 'MULTIDOSSIER',   True ) ;
    SetControlVisible( 'TMULTIDOSSIER',  True ) ;
    SetControlVisible( 'BVMULTIDOSSIER', True ) ;
    //SetControlVisible( 'RUPTUREMAX',     True ) ;
  end;

  // Détermination du Type de Grand Livre Auxiliaire
  FBoAuxiParGene := IIF( Ecran.Name = 'CPGLAUXIPARGENE', True, False);

  // Branchement du HelpContext en fonction de l'état
  Ecran.HelpContext := IIF( FBoAuxiParGene, 7430000, 7560000);

  E_Auxiliaire     := THEdit(GetControl('EAUXILIAIRE', True));
  E_Auxiliaire_    := THEdit(GetControl('EAUXILIAIRE_', True));
  AuxiParGene      := THEdit(GetControl('AUXIPARGENE', True));

  // Grand Livre Auxiliaire Par Général
  if FBoAuxiParGene then
  begin
    E_GENERAL            := THEdit(GetControl('EGENERAL', True));
    E_GENERAL_           := THEdit(GetControl('EGENERAL_', True));
    E_General.MaxLength  := VH^.CPta[fbGene].Lg;
    E_General_.MaxLength := VH^.CPta[fbGene].Lg;
    E_General.OnExit     := OnExitE_General;
    E_General_.OnExit    := OnExitE_General;
  end;

  E_DATECOMPTABLE  := THEdit(GetControl('EDATECOMPTABLE', True));
  E_DATECOMPTABLE_ := THEdit(GetControl('EDATECOMPTABLE_', True));
  E_NUMEROPIECE    := THEdit(GetControl('E_NUMEROPIECE', True));
  E_NUMEROPIECE_   := THEdit(GetControl('E_NUMEROPIECE_', True));
  E_REFINTERNE     := THEdit(GetControl('E_REFINTERNE', True));

  E_EXERCICE       := THValComboBox(GetControl('EEXERCICE', True));
  CInitComboExercice(E_EXERCICE);

  E_DEVISE         := THValComboBox(GetControl('E_DEVISE', True));
  E_QUALIFPIECE    := THMultiValComboBox(GetControl('EQUALIFPIECE', True));
  T_NATUREAUXI     := THValComboBox(GetControl('T_NATUREAUXI', True));

  MODESELECT       := THValComboBox(GetControl('MODESELECT', True));
  AFFDEVISE        := THEdit(GetControl('AFFDEVISE', True));
  DEVPIVOT         := THEdit(GetControl('DEVPIVOT', True));
  CPTEXCEPT        := THEdit(GetControl('CPTEXCEPT', True));

  ECRValide        := TCheckBox(GetControl('ECRVALIDE', True));
  ECRLETTRE        := TCheckBox(GetControl('ECRLETTRE', True));
  AFFANOUVEAU      := TCheckBox(GetControl('AFFANOUVEAU', True));
  AFFCUMUL         := TCheckBox(GetControl('AFFCUMUL', True));

  CBEnSituation    := TCheckBox(GetControl('CBENSITUATION', True));
  AFFCentraEche    := TCheckBox(GetControl('AFFCENTRAECHE', True));
  EDateSituation   := THEdit(GetControl('EDATESITUATION', True));
  DateDebutExo     := THEdit(GetControl('DATEDEBUTEXO', True));
  Exercice         := THEdit(GetControl('EXERCICE', True));

  Pages           := TPageControl(GetControl('Pages', True));
  RUPTURE         := THRadioGroup(GetControl('RUPTURE', True));
  RUPTURES        := TTabSheet(GetControl('RUPTURES', True));
  RUPTURETYPE     := THRadioGroup(GetControl('RUPTURETYPE', True));
  RUPCORRESP      := TGroupBox(GetControl('RUPCORRESP', True));
  RUPLIBRES       := TGroupBox(GetControl('RUPLIBRES', True));
  RUPGROUPES      := TGroupBox(GetControl('RUPGROUPES', True));
  CORRESP         := THValComboBox(GetControl('CORRESP', True));
  CORRESPDE       := THValComboBox(GetControl('CORRESPDE', True));
  CORRESPA        := THValComboBox(GetControl('CORRESPA', True));

  CBPreview       := TCheckBox(GetControl('FAPERCU', True));
  // Case à cocher : Uniquement les comptes associés
  CPTCORRESPEXIST := TCheckBox(GetControl('CPTCORRESPEXIST', True));
  CPTLIBREEXIST   := TCheckBox(GetControl('CPTLIBRESEXIST', True));

  // Zones cachées
  XX_RUPTURE      := THEdit(GetControl('XX_RUPTURE', True));
  AVECRUPTYPE     := THEdit(GetControl('AVECRUPTYPE', True));
  AVECRUPTURE     := THEdit(GetControl('AVECRUPTURE', True));
  AVECNIVORUPTURE := THEdit(GetControl('AVECNIVORUPTURE', True));
  NIVORUPTURE     := TSpinEdit(GetControl('NIVORUPTURE', True));
  OKENTETE        := THEdit(GetControl('OKENTETE', True));
  OKRESUME        := THEdit(GetControl('OKRESUME', True));
  LIBANO          := THEdit(GetControl('LIBANO', True));

  TABLELIBRE      := THValComboBox(Getcontrol('TABLELIBRE', True));
  LIBREDE         := THEdit(GetControl('LIBREDE', True));
  LIBREA	  := THEdit(GetControl('LIBREA', True));
  CHOIXTYPETALI   := THRadioGroup(GetControl('CHOIXTYPETALI',False));

  LIBEQUALIFPIECE := THEdit(GetControl('LIBEQUALIFPIECE', True));
  CUMULAU         := THEdit(GetControl('CUMULAU', True));

  // ----------------- Init des PROPERTY ---------------------------------------
  E_Auxiliaire.MaxLength  := VH^.CPta[fbAux].Lg;
  E_Auxiliaire_.MaxLength := VH^.CPta[fbAux].Lg;

  RUPTURES.TabVisible := False;
  RUPCORRESP.Visible  := False;
  RUPLIBRES.Visible   := False;
  RUPGROUPES.Visible  := True;

  // Type de plan comptable :
  //if V_PGI.LaSerie = S7 then // FP 25/04/2006 FQ18001
  if GetParamSocSecur('SO_CORSAU2', False, True) then 
    Corresp.plus := 'AND (CO_CODE = "AU1" OR CO_CODE = "AU2")'
  else
    Corresp.plus := 'AND CO_CODE = "AU1"';

  // Centralisation des échéances visible en mode PGE uniqeuement
  //AffCentraEche.Visible := not ( Ctxpcl in V_PGI.PGIContexte ) ;

  // ----------------- Branchement des événements ------------------------------
  T_NATUREAUXI.OnChange := OnChangeT_NATUREAUXI;
  E_EXERCICE.OnChange   := OnChangeE_Exercice;

  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' ) then
    E_EXERCICE.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    E_EXERCICE.Value := CExerciceVersRelatif(VH^.Entree.Code) ;

  E_Auxiliaire.OnExit    := OnExitE_Auxiliaire;
  E_Auxiliaire_.OnExit   := OnExitE_Auxiliaire;
  CORRESP.OnChange       := OnChangeCorresp;
  TABLELIBRE.OnChange    := OnChangeTableLibre;
  E_DATECOMPTABLE.OnExit := OnExitE_DateComptable;
  E_DATECOMPTABLE.OnExit := OnExitE_DateComptable;
  E_DEVISE.OnChange      := OnChangeE_Devise;
  CBEnSituation.OnClick  := OnClickCBEnSituation;
  Rupture.OnClick        := OnClickRupture;
  RuptureType.OnClick    := OnClickRuptureType;
  FTimer.OnTimer         := FTimerTimer;
  E_QUALIFPIECE.OnEnter  := OnEnterQualifPiece ; // SBO 12/10/2004
  AffCentraEche.OnClick  := OnClickAffCentraEche;

  if CHOIXTYPETALI <> nil then
    CHOIXTYPETALI.OnClick  := OnClickChoixTypeTaLi;

  {$IFDEF GIL}
    SetControlProperty('TABDEV', 'TABVISIBLE', True);
  {$ELSE}
    SetControlProperty('TABDEV', 'TABVISIBLE', (V_PGI.Sav) and (V_PGI.PassWord = CryptageSt(DayPass(Date))));
  {$ENDIF}

  // CA - 03/06/2004 - Initialisation du choix du type de table libre
  if CHOIXTYPETALI <> nil then
  begin
    CHOIXTYPETALI.Visible := TL_TIERSCOMPL_Actif;
    CHOIXTYPETALI.ItemIndex := 0;
  end;
  InitChoixTableLibre;

  // GCO - 03/01/2007
  FBoOkDecoupe := True;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    E_Auxiliaire.OnElipsisClick:=AuxiElipsisClick;
    E_Auxiliaire_.OnElipsisClick:=AuxiElipsisClick;
  end;
end;



////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function OnMyValideCritereEvent : Boolean;
begin
  Result := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/12/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.GenereCumulGL;
var lStSelect      : string;
    lStWhere       : string;
    lStWhereCommun : string;
    lStWhereAuxi   : string;
    lStRgrp        : string;
    lQuery         : TQuery;
begin
  if FTobCumulAnoGL = nil then
    FTobCumulAnoGL := Tob.Create('', nil, -1)
  else
    FTobCumulAnoGL.ClearDetail;

  if FTobCumulAuGL = nil then
    FTobCumulAuGL := Tob.Create('', nil, -1)
  else
    FTobCumulAuGL.ClearDetail;

  if FTobBilanGestion = nil then
  begin
    FTobBilanGestion := Tob.Create('', nil, -1);
    FTobBilanGestion.AddChampSupValeur('DEBIT', 0, False);
    FTobBilanGestion.AddChampSupValeur('CREDIT', 0, False);
  end
  else
  begin
    FTobBilanGestion.SetDouble('DEBIT', 0);
    FTobBilanGestion.SetDouble('CREDIT', 0);
  end;

  lStSelect := 'SELECT E_AUXILIAIRE, ' + IIF(FBoAuxiParGene, 'E_GENERAL,', '');

  if AffDevise.Text = 'X' then
    lStSelect := lStSelect + ' SUM(E_DEBITDEV) DEBIT, SUM(E_CREDITDEV) CREDIT'
  else
    lStSelect := lStSelect + ' SUM(E_DEBIT) DEBIT, SUM(E_CREDIT) CREDIT';

  lStSelect := lStSelect + ' FROM ECRITURE WHERE E_QUALIFPIECE = "N"';

    // Where Commun aux 2 requêtes
  lStWhereCommun := '';
  if E_Exercice.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND E_EXERCICE = "' + CRelatifVersExercice(E_EXERCICE.Value) + '"';

  if ComboEtab.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND E_ETABLISSEMENT = "' + ComboEtab.Value + '"';

  if E_Devise.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND E_DEVISE = "' + E_DEVISE.Value + '"';

  // GCO - 30/08/2006 - FQ 18698 - Ajout test Auxiliaire et Généraux pour optimisation
  lStWhereAuxi   := ConvertitCaractereJokers(E_Auxiliaire, E_Auxiliaire_, 'E_AUXILIAIRE');
  if Trim( lStWhereAuxi ) <> '' then
    lStWhereCommun := lStWhereCommun + ' AND ' + lStWhereAuxi ;

  if FBoAuxiParGene then
  begin
    lStWhereAuxi := ConvertitCaractereJokers(E_GENERAL, E_GENERAL_, 'E_GENERAL');
    if Trim( lStWhereAuxi ) <> '' then
      lStWhereCommun := lStWhereCommun + ' AND ' + lStWhereAuxi ;
  end ;
  // Fin FQ 18698

  lStWhereCommun := lStWhereCommun +
                    ' GROUP BY E_AUXILIAIRE ' + IIF(FBoAuxiParGene, ',E_GENERAL', '') +
                    ' ORDER BY E_AUXILIAIRE ' + IIF(FBoAuxiParGene, ',E_GENERAL', '');

  lStWhere := ' AND E_DATECOMPTABLE = "' + UsDateTime(StrToDate(DateDebutExo.Text)) + '"' +
              ' AND (E_ECRANOUVEAU = "OAN" OR E_ECRANOUVEAU = "H")';

  // GCO - 08/11/2007 - FQ 20377 Mode Multi-société
  if EstMultiSoc then
  begin
    try
      RequeteMultiDossier( Pages, lStRgrp );
      lQuery := OpenSelect( lStSelect + lStWhere + lStWhereCommun, lStRgrp);
      FTobCumulAnoGL.LoadDetailDB( '', '', '', lQuery, False);
    finally
      Ferme(lQuery);
    end;
  end
  else
    FTobCumulAnoGL.LoadDetailFromSql( lStSelect + lStWhere + lStWhereCommun );

  // Calcul du Cumul AU
  lStWhere := ' AND E_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DateDebutExo.Text)) + '" AND ' +
               'E_DATECOMPTABLE < "' + UsDateTime(StrToDate(E_DateComptable.Text)) + '" AND ' +
               'E_ECRANOUVEAU = "N"';

  // GCO - 08/11/2007 - FQ 20377
  if EstMultiSoc then
  begin
    try
      lQuery := OpenSelect( lStSelect + lStWhere + lStWhereCommun, lStRgrp);
      FTobCumulAuGL.LoadDetailDB( '', '', '', lQuery, False);
    finally
      Ferme(lQuery);
    end;
  end
  else
    FTobCumulAuGL.LoadDetailFromSql( lStSelect + lStWhere + lStWhereCommun );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.ChargementCritEdt;
begin
  inherited;

  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    E_Auxiliaire.Text  := ClassCritEdt(TheData).CritEdt.Cpt1;
    E_Auxiliaire_.Text := ClassCritEdt(TheData).CritEdt.Cpt2;

    if FBoAuxiParGene then
    begin
      E_General.Text  := ClassCritEdt(TheData).CritEdt.SCpt1;
      E_General_.Text := ClassCritEdt(TheData).CritEdt.SCpt2;
    end;

    E_Exercice.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);

    // Date de Début de l'édition
    E_DateComptable.Text  := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);

    // Date de Fin de l'édition
    E_DateComptable_.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);

    // Numéro de Pièce
    // JP 26/08/04 : FQ 14062 : avec l'autocomplétion, si NumPiece1 = 0 et NumPiece2 = 9999999,
    //                          on se retrouve E_NUMEROPIECE à 99999999
    // GCO - 05/09/2005 - FQ 16320
    if ClassCritEdt(TheData).CritEdt.GL.NumPiece1 > 0 then
      E_NUMEROPIECE.Text  := IntToStr(ClassCritEdt(TheData).CritEdt.GL.NumPiece1);

    if ClassCritEdt(TheData).CritEdt.GL.NumPiece2 > 0 then
      E_NUMEROPIECE_.Text := IntToStr(ClassCritEdt(TheData).CritEdt.GL.NumPiece2);

    // En situation
    CBEnSituation.Checked := ClassCritEdt(TheData).CritEdt.GL.EnDateSituation;
    if CBEnsituation.Checked then
    begin
      Ecran.Caption := TraduireMemoire('Justificatif de solde');
      UpdateCaption(Ecran);
      EDateSituation.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);
    end
    else
    begin // Détail A-Nouveaux
      AFFANOUVEAU.Checked := ClassCritEdt(TheData).CritEdt.GL.DetailAno;
    end;

    // Référence Interne
    E_REFINTERNE.Text := ClassCritEdt(TheData).CritEdt.ReferenceInterne;

    // E_Qualifpièce
    E_QUALIFPIECE.Text := ClassCritEdt(TheData).CritEdt.QualifPiece;

    // Etablissement
   ComboEtab.Value := ClassCritEdt(TheData).CritEdt.Etab;

    // Devise
    E_DEVISE.Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;

    // Ecritures validées ( '' = Grayed, OUI = Checked sinon not Checked
    if (ClassCritEdt(TheData).CritEdt.Valide = '') then
      ECRVALIDE.State := cbGrayed
    else
      ECRVALIDE.Checked := (ClassCritEdt(TheData).CritEdt.Valide = 'OUI');

    // Ecritures Lettrées ( 0 = Grayed, 1 = not Checked sinon 2 = Checked
    if (ClassCritEdt(TheData).CritEdt.GL.Lettrable = 0) then
      ECRLETTRE.State   := cbGrayed
    else
      ECRLETTRE.Checked := (ClassCritEdt(TheData).CritEdt.GL.Lettrable = 2);

    TheData := nil;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2004
Modifié le ... : 26/10/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnExitE_DateComptable(Sender: TObject);
var lExoDate : TExoDate;
begin
  // GCO - 26/10/2005 - FQ 14967
  if E_Exercice.ItemIndex <> 0 then
  begin
    lExoDate := CtxExercice.QuelExoDate( CRelatifVersExercice(E_Exercice.Value));
    AFFANOUVEAU.Enabled := StrToDate(E_DateComptable.Text) = lExoDate.Deb;
    if not AFFANouveau.Enabled then
      AffANouveau.Checked := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.OnChangeE_Exercice(Sender: TObject);
begin
  if E_Exercice.ItemIndex = 0 then
  begin
    //AFFANouveau.Enabled   := False;
    AFFANouveau.Checked   := False;
    if VH^.ExoV8.Code <> '' then
      E_DateComptable.Text := DateToStr(VH^.ExoV8.Deb)
    else
      E_DateComptable.Text := DateToStr(VH^.Exercices[1].Deb);
    E_DateComptable_.Text  := '30/12/2099';
  end
  else
  begin
    CExoRelatifToDates ( E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
    if StrToDate(E_DATECOMPTABLE.Text) < VH^.ExoV8.Deb then
    begin
      AFFANouveau.Checked   := False;
      AFFANouveau.Enabled   := False;
      CBENSITUATION.Checked := False;
      CBENSITUATION.Enabled := False;
    end
    else
    begin
      AFFANouveau.Enabled   := True;
      CBENSITUATION.Enabled := True;
    end;
  end;  
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : Changement de la Devise et
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnChangeE_Devise(Sender: TObject);
begin
  RGDevise.Enabled := (E_Devise.ItemIndex <> 0) and (E_Devise.Value <>  V_PGI.DevisePivot);
  if E_Devise.ItemIndex = 0 then
    RGDevise.ItemIndex := 0;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : Selection des écritures en situation
Suite ........ : et Activation la zone de saisie de la date
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnClickCBEnSituation(Sender: TObject);
begin
  EDateSituation.Enabled   := CBEnSituation.Checked;
  EDateSituation.Color     := IIF(CBEnSituation.Checked, ClWindow, ClBtnFace);
  E_DateComptable.Enabled  := not CBEnSituation.Checked;
  E_DateComptable_.Enabled := not CBEnSituation.Checked;
  E_DateComptable.Color    := IIF(not CBEnSituation.Checked, ClWindow, ClBtnFace);
  E_DateComptable_.Color   := IIF(not CBEnSituation.Checked, ClWindow, ClBtnFace);
  E_QUALIFPIECE.Enabled    := not CBEnSituation.Checked;

  // GCO - 10/11/2004 - FQ 13498 et 14865
  AFFCumul.Checked := not CBEnSituation.Checked;
  AFFCumul.Enabled := not CBEnSituation.Checked;

  if CBEnSituation.Checked then
  begin
    E_Exercice.ItemIndex := 0;
    E_Exercice.Enabled   := False;
    E_Exercice.OnChange(Self);
    AFFANouveau.Enabled  := False;
    AFFANouveau.Checked  := True;
    E_NumeroPiece.Text   := '';
    E_NumeroPiece_.Text  := '';
    E_QualifPiece.Text   := 'N;'
  end
  else
  begin
    E_Exercice.Enabled   := True;
    AFFANouveau.Enabled  := True;
    AFFANouveau.Checked  := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.OnClickRupture(Sender: TObject);
begin
  Ruptures.TabVisible := (Rupture.ItemIndex <> 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.OnClickRuptureType(Sender: TObject);
begin
  RupGroupes.Visible := (RuptureType.ItemIndex = 0);
  RupLibres.Visible  := (RuptureType.ItemIndex = 1);
  RupCorresp.Visible := (RuptureType.ItemIndex = 2);

  if RuptureType.ItemIndex = 1 then
  begin
    TableLibre.ItemIndex := 0;
    OnChangeTableLibre(nil);
  end;

  if RuptureType.ItemIndex = 2 then
  begin
    if Corresp.Items.Count >= 1 then
      Corresp.ItemIndex := 1
    else
      Corresp.ItemIndex := 0;

    OnChangeCorresp(nil);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2004
Modifié le ... : 06/09/2005    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.FTimerTimer(Sender: TObject);
begin
  if FCritEdtChaine <> nil then
  begin
    // GCO - 03/01/2007 - FQ 19375
    FBoOkDecoupe := False;

    with FCritEdtChaine do
    begin
      if CritEdtChaine.UtiliseCritStd then
      begin
        E_Exercice.Value      := CritEdtChaine.Exercice.Code;
        E_DateComptable.Text  := DateToStr(CritEdtChaine.Exercice.Deb);
        E_DateComptable_.Text := DateToStr(CritEdtChaine.Exercice.Fin);

        // GCO - 03/08/2007 - FQ 19377
        T_NatureAuxi.Value    := CritEdtChaine.NatureCompte;
        ModeSelect.Value      := CritEdtChaine.ModeSelection;
        E_QualifPiece.Value   := CritEdtchaine.TypeEcriture;
      end;
    end;
  end;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. : Restreint la tablette des comptes en fonction de la nature
Suite ........ : de compte sélectionnée
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnChangeT_NATUREAUXI(Sender: TObject);
begin
  E_Auxiliaire.Text := '';
  E_Auxiliaire_.Text := '';

  E_Auxiliaire.Plus := IIF(T_NatureAuxi.ItemIndex > 0, ' AND T_NATUREAUXI="' + T_Natureauxi.Value + '"', '');
  E_Auxiliaire_.Plus := E_Auxiliaire.Plus;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLAUXI.GenereConditionWhereEtat: string;
var lStWhereCritere : string;
begin
  Result := '(';

  // La requête va se présenter sous cette forme WHERE ((1) or (2) or (3)) AND Criteres Communs
  // Début Récupération du détail des A-Nouveaux pour les comptes Lettrables uniquement (1)
  if AFFANOUVEAU.Checked then
  begin
    Result := Result + '((E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H") AND ' +
    'E_QUALIFPIECE = "N" AND ' + // Ajout GCO - 09/02/20004
    'T_LETTRABLE = "X" AND ' + 'E_DATECOMPTABLE <= "' + UsDateTime(FExoDate.Deb) + '" AND ' +
    '(E_ETATLETTRAGE <> "TL" OR (E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX >= "' + UsDateTime(FExoDate.Deb) + '"';

    if CBEnSituation.Checked then
      Result := Result + ' AND E_DATEPAQUETMAX > "' + UsDateTime( StrToDate( EDateSituation.Text)) + '"))'
    else
      Result := Result + '))';

    Result := Result + ') OR ' ;
  end;
  // Fin récupération du détail des A-Nouveaux pour les comptes Lettrables uniquement (1) : Terminé

  // Début de récupération des écritures A-Nouveaux PGI(2)
  // NB : Ne jamais prendre les A-Nouveaux uniquement à la Date de début car si on édite le GL
  //      et que la requête ne renvoie pas d'enregistrement pour le compte, il faut au moins
  //      provoquer la présence du compte dans le Grand livre , afin que les cumuls s'éditent
  //      pour avoir une info sur le solde du compte
  Result := Result + '(';
  Result := Result + 'E_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DateDebutExo.Text)) + '"'
                               + ' AND E_DATECOMPTABLE <= "' + UsDateTime(StrToDate(E_DateComptable.Text)) + '"'
                               + ' AND E_QUALIFPIECE = "N"';
  if E_Exercice.ItemIndex <> 0 then
  begin
    Result := Result + ' AND E_EXERCICE ="' + CRelatifVersExercice(E_EXERCICE.Value) + '"';
  end;
  Result := Result + ' AND (E_ECRANOUVEAU = "OAN" OR E_ECRANOUVEAU = "H")';

  if AFFANouveau.Checked then
    Result := Result + ' AND T_LETTRABLE = "-"';
  Result := Result + ') OR ';
  // Fin de Récupération des écritures OAN (PGI)
  // On a récupéré les écritures OAN ou le détail des A-Nouveaux

  // Début de Récupération des écritures concernées par les critères (3)
  Result := Result + '(E_ECRANOUVEAU = "N"';

  //
  if AFFCentraEche.Checked then
    Result  := Result  + ' AND E_NUMECHE < 1';

  lStWhereCritere := RecupAutresCriteres ;
  Result  := Result  + ' ' + lStWhereCritere;

  //lStWhereCritere := RecupWhereCritere(Pages);
  //Result  := Result  + ' AND ' + Copy(lStWhereCritere, 6, length(lStWhereCritere));

  // Ecritures en situation à une date donnée
  if CBEnSituation.Checked then
  begin
    if IsValidDate(EDateSituation.Text) then
    begin
      Result := Result + ' AND E_DATECOMPTABLE <= "' + UsDateTime(StrToDate(EDateSituation.Text)) + '" AND ' +
        '(E_ETATLETTRAGE<>"TL" OR ' +
        '(T_LETTRABLE="-" OR (E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX > "' +
        UsDateTime(StrToDate(EDateSituation.Text)) + '")))';
    end;
  end;
  Result := Result + ')';
  // Fin de Récupération des écritures concernées par les critères

  // Ajout des écritures de Centralisation dans la condition Where de la requête
  // On ne prend que les Ecritures "Z" du User, en effet plusieurs utilisateurs
  // peuvent faire un grand livre avec centralisation en même temps.
  if AFFCentraEche.Checked then
    Result := Result + ' OR (E_QUALIFPIECE = "Z" AND E_UTILISATEUR = "' + string(V_PGI.User) + '")';

  // Fermeture dela parenthèse principale qui englobe tous les ( (1) OR (2) OR (3) )
  Result := Result + ')';

  // PARAMETRES COMMUNS AUX ECRITURES ET AUX A-NOUVEAUX
  Result := Result + RecupCriteresCommuns;

  Result := Result + RecupCriteresModeSelection(Result);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLAUXI.RecupAutresCriteres : string;
var lStValEQualifPiece : string;
    lStLibEQualifPiece : string;
begin
  // Borne Max de E_DATECOMPTABLE
  Result := 'AND E_DATECOMPTABLE <="' + USDateTime(StrToDate(E_DateComptable_.Text)) + '"' ;

  // E_Exercice, E_DATECOMPTABLE ne peuvent être pris par le RecupWhereCriteres
  // car ne doit pas intervenir en date de Situation,
  if not CbEnsituation.Checked then
    Result := Result + ' AND E_DATECOMPTABLE >="' + USDateTime(StrToDate(E_DateComptable.Text)) + '"'
  else
    Result := Result + ' AND E_DATECOMPTABLE >="' + USDateTime(StrToDate(DateDebutExo.Text)) + '"';

  if (not CBEnSituation.Checked) then
  begin
    if E_Exercice.ItemIndex <> 0 then // Tous
      Result := Result + ' AND E_EXERCICE ="' + cRelatifVersExercice(E_Exercice.Value) + '"';
  end;
  
  // Ajout GCO - 09/02/2004
  if CBEnSituation.Checked then
  begin
    // E_QUALIFPIECE est forcé à "N" en justificatif de solde
    lStLibEQualifPiece := TraduireMemoire('Normal');
    lStValEQualifPiece := '(E_QUALIFPIECE = "N")';
  end
  else
  begin
    // Sélection du type d'écritures et  Traduction du E_QualifPiece
    TraductionTHMultiValComboBox( E_QualifPiece, lStValEQualifPiece, lStLibEQualifPiece, 'E_QUALIFPIECE', False );
  end;

  LibEQualifPiece.Text := lStLibEQualifPiece;

  if Trim(lStValEqualifPiece) <> '' then
    Result := Result + ' AND ' + lStValEQualifPiece;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLAUXI.RecupCriteresCommuns: string;
var lStListeCptExcept  : string;
    lStCptExcept       : string;
    lStValEQualifPiece : string;
    lStLibEQualifPiece : string;
    lStTemp            : string;
begin

  // Sélection du type d'écritures et  Traduction du E_QualifPiece
  TraductionTHMultiValComboBox( E_QualifPiece, lStValEQualifPiece, lStLibEQualifPiece, 'E_QUALIFPIECE', False, '"Z"');

  // --> Compte Auxiliaire
  // JP 01/07/05 : Gestion des caractères jokers
  Result := ' AND ' + ConvertitCaractereJokers(E_Auxiliaire, E_Auxiliaire_, 'E_AUXILIAIRE') +
            ' AND E_AUXILIAIRE <> ""';

  // --> Compte Général si Grand Livre Auxiliaire par Général
  if FBoAuxiParGene then
  begin
    // JP 01/07/05 : Gestiondes  caractères jokers
    Result := Result + ' AND ' + ConvertitCaractereJokers(E_GENERAL, E_GENERAL_, 'E_GENERAL');
  end;

  // Prise en compte du VH^.ExoV8 afin de cacher les ecritures antérieures
  // uniquement en date de situation
  if (CBEnSituation.Checked or AFFANOUVEAU.Checked) and (VH^.ExoV8.Code <> '') then
  begin
    Result := Result + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '" ';
  end;

  // GCO - 11/07/2006 - FQ 18363
  lStTemp := RecupWhereCritere(Pages);
  if lStTemp <> '' then
    lStTemp := Copy(lStTemp, 6, length(lStTemp)); // Dégage le mot WHERE
  Result := Result + ' AND ' + lStTemp;

  // Sélection du type d'écritures et  Traduction du E_QualifPiece
  TraductionTHMultiValComboBox( E_QualifPiece, lStValEQualifPiece, lStLibEQualifPiece, 'E_QUALIFPIECE', False, '"Z"');
  LibEQualifPiece.Text := lStLibEQualifPiece;
  if Trim(lStValEqualifPiece) <> '' then
    Result := Result + ' AND ' + lStValEQualifPiece;


  // Justif de Solde pour les comptes généraux lettrables uniquement
  if CBEnSituation.Checked then
    Result := Result + ' AND T_LETTRABLE = "X"';

  // Ecritures validées
  if ECRValide.State <> CbGrayed then
  begin
    if ECRValide.Checked then
      Result := Result + ' AND E_VALIDE = "X" '
    else
      Result := Result + ' AND E_VALIDE = "-" ';
  end;

  // Ecritures Lettrés
  if ECRLettre.State <> CbGrayed then
  begin
    if ECRLettre.Checked then
      Result := Result + ' AND (E_ETATLETTRAGE = "TL"  OR E_ETATLETTRAGE = "PL") '
    else
      Result := Result + ' AND (E_ETATLETTRAGE = "RI" OR E_ETATLETTRAGE = "AL") ';
  end;

  Result := Result + ' AND (E_DEBIT <> 0 OR E_CREDIT <> 0) AND E_CREERPAR <> "DET"' ;

  // Comptes d'exceptions (les comptes peuvent être séparés par des ',' ou ';')
  if CptExcept.Text <> '' then
  begin
    lStListeCptExcept := FindEtReplace(CptExcept.Text, ',', ';', True);
    while (lStListeCptExcept <> '') do
    begin
      lStCptExcept := Trim(ReadTokenSt(lStListeCptExcept));
      if lStCptExcept <> '' then
        Result := Result + ' AND E_AUXILIAIRE NOT LIKE "' + lStCptExcept + '%"';
    end;
  end;

  // Gestion du V_Pgi.Confidentiel
  Result := Result  + ' AND ' + CGenereSQLConfidentiel('T');
  if FBoAuxiParGene then
    Result := Result + ' AND ' + CGenereSQLConfidentiel('G');

  // GCO - 12/09/2006 - FQ 18784
  if Rupture.Value <> 'SANS' then
  begin
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
    begin
      // Mise à jour de la requête pour rupture
      if CPTLibreExist.Checked then
      begin
        // GCO - 04/09/2006 - FQ 18432
        Result := Result + ' AND ' + XX_Rupture.Text + ' <> ''';

        if Trim(LibreDe.Text) <> '' then
          Result := Result + ' AND '+ XX_Rupture.Text + ' >= "' + LibreDe.Text + '"';

        if Trim(LibreA.Text) <> '' then
          Result := Result + ' AND '+ XX_RUPTURE.Text + ' <= "' + LibreA.Text + '"';
      end;
    end;

    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
    begin
      if CPTCorrespExist.Checked then
      begin
        if Corresp.Value = 'AU1' then
        begin
          // GCO - 04/09/2006 - FQ 18432
          Result := Result + ' AND T_CORRESP1 <> ''';

          if Trim(CorrespDe.Text) <> '' then
            Result := Result + ' AND T_CORRESP1 >= "' + CorrespDe.Text + '"';

          if Trim(CorrespA.Text) <> '' then
            Result := Result + ' AND T_CORRESP1 <= "' + CorrespA.Text + '"';
        end
        else
        begin
          // GCO - 04/09/2006 - FQ 18432
          Result := Result + ' AND T_CORRESP2 <> ''';

          if Trim(CorrespDe.Text) <> '' then
            Result := Result + ' AND T_CORRESP2 >= "' + CorrespDe.Text + '"';

          if Trim(CorrespA.Text) <> '' then
            Result := Result + ' AND T_CORRESP2 <= "' + CorrespA.Text + '"';
        end;
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLAUXI.RecupCriteresModeSelection( vStWhereEtat: string): string;
begin
  Result := '';

  if ModeSelect.Value = 'FOU' then Exit; // Tous les Comptes

  // Mode de sélection
  if ModeSelect.Value = 'EXO' then
  begin // Comptes mouvementés sur l'exercice
    Result := Result + ' AND (E_AUXILIAIRE in (SELECT DISTINCT E_AUXILIAIRE FROM ECRITURE WHERE ' +
              vStWhereEtat + '))';
  end
  else
  if ModeSelect.Value = 'NSL' then
  begin // Comptes non soldés
    Result := Result + ' AND (E_AUXILIAIRE in (SELECT DISTINCT E_AUXILIAIRE FROM ECRITURE WHERE ' +
              vStWhereEtat + ' GROUP BY E_AUXILIAIRE HAVING SUM(E_DEBIT) <> SUM( E_CREDIT))) ';
  end
  else
  if ModeSelect.Value = 'PER' then
  begin // Comptes mouvementés sur la période
    Result := Result + ' AND (E_AUXILIAIRE in (SELECT DISTINCT E_AUXILIAIRE FROM ECRITURE WHERE ' +
              vStWhereEtat + '))';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/08/2003
Modifié le ... :   /  /
Description .. : Changement de la rupture sur les plans de correspondance
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnChangeCorresp(Sender: TObject);
begin
  if RuptureType.value <> 'RUPCORRESP' then Exit;
  CorrespToCodes(Corresp,TComboBox(CorrespDe),TComboBox(CorrespA));
  CorrespDe.ItemIndex := 0 ;
  CorrespA.ItemIndex  := CorrespA.Items.Count - 1 ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/08/2003
Modifié le ... :   /  /    
Description .. : Changement de la rupture sur les Tables Libres
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.OnChangeTableLibre(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
  begin
    LibreDe.DataType := '';
    LibreA.DataType  := '';
    LibreDe.Text     := '';
    LibreA.Text      := '';
  end
  else
  begin
    if CHOIXTYPETALI <> nil then
    begin
      if CHOIXTYPETALI.Value = 'AUX' then
      begin
        LibreDe.DataType := 'TZNATTIERS' + GetNumTableLibre;
        LibreA.DataType  := 'TZNATTIERS' + GetNumTableLibre;
      end else
      if CHOIXTYPETALI.Value = 'CLI' then
      begin
        LibreDe.DataType := 'GCLIBRETIERS' + GetNumTableLibre;
        LibreA.DataType  := 'GCLIBRETIERS' + GetNumTableLibre;
      end else
      if CHOIXTYPETALI.Value = 'SCL' then
      begin
        LibreDe.DataType := 'AFLRESSOURCE';
        LibreA.DataType  := 'AFLRESSOURCE';
      end;
    end // CHOIXTYPETALI <> nil then
    else
    begin // GCO - 23/06/2006 - FQ 18432
      LibreDe.DataType := 'TZNATTIERS' + GetNumTableLibre;
      LibreA.DataType  := 'TZNATTIERS' + GetNumTableLibre;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/08/2003
Modifié le ... : 07/04/2004
Description .. : Prepare les compsoants cachés afin de programmer les ruptures
Suite ........ : sur l'état du Grand Livre
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.GestionCritereDEV;
var lStTemp : string;
begin
  AvecRupture.Text     := 'SANS' ;
  AvecRupType.Text     := '' ;
  AvecNivoRupture.Text := '0' ;
  XX_Rupture.Text      := '' ;
  AuxiParGene.Text     := IIF(FBoAuxiParGene, 'X', '-');
  OkEntete.Text        := 'X';
  OkResume.Text        := 'X';

  LibAno.Text := TraduireMemoire('Total des A-Nouveaux');
  if E_Exercice.ItemIndex = 0 then
    Exercice.Text := '---'
  else
  begin
    Exercice.Text  := FExoDate.Code;
    // Test de la présence d' A-Nouveaux
    if ExisteSQL('SELECT E_GENERAL FROM ECRITURE WHERE ' +
       'E_EXERCICE = "' + FExoDate.Code + '" AND ' +
       'E_DATECOMPTABLE = "' + UsDateTime(FExoDate.Deb) + '" AND ' +
       '(E_ECRANOUVEAU = "OAN" OR E_ECRANOUVEAU = "H")') then
    begin
      lStTemp := GetColonneSQL('EXERCICE', 'EX_ETATCPTA', 'EX_DATEFIN = "' + UsDateTime(FExoDate.Deb-1) + '"');
      if (lStTemp <> '') and (lStTemp <> 'CDE') then
        LibAno.Text := TraduireMemoire('Total des A-Nouveaux provisoires');
    end;
  end;

  DateDebutExo.Text := DateToStr(FExoDate.Deb);

  // Libellé Cumul au
  if CBEnSituation.Checked then
    CumulAu.Text := DateToStr(StrToDate(DateDebutExo.Text) - 1)
  else
    CumulAu.Text := DateToStr(StrToDate(E_DateComptable.Text) - 1);

  // Affichage en DEVISE ou en EURO
  RGDevise.Items[0] := '&' + VH^.LibDevisePivot;
  if (RGDevise.Enabled) and (RGDevise.ItemIndex = 1) then
  begin // Affichage en devise
    AFFDEVISE.Text := 'X';
    DEVPIVOT.Text  := TraduireMemoire('Devise');
  end
  else
  begin
    AFFDEVISE.Text := '-';
    DEVPIVOT.Text  := VH^.LibDevisePivot;
  end;

  // Gestion des ruptures
  if Rupture.Value <> 'SANS' then
  begin
    AvecRupType.Text := RuptureType.Value;
    AvecRupture.Text := Rupture.Value;
    // Rupture sur Numéro de Compte
    if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.Value > 0) then
    begin
      AvecNivoRupture.Text := IntToStr(NivoRupture.Value);
    end;

    // Rupture sur champ libre généraux
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
    begin
      // GCO - 23/06/2006 - FQ 18432
      XX_Rupture.Text := 'T_TABLE' + GetNumTableLibre;

      // Mise à jour du champs de rupture
      if CHOIXTYPETALI <> nil then
      begin
        if CHOIXTYPETALI.Value = 'AUX' then
          XX_Rupture.Text := 'T_TABLE' + GetNumTableLibre
        else
        if CHOIXTYPETALI.Value = 'CLI' then
          XX_Rupture.Text := 'YTC_TABLELIBRETIERS' + GetNumTableLibre
        else
        if CHOIXTYPETALI.Value = 'SCL' then
          XX_Rupture.Text := 'YTC_RESSOURCE' + GetNumTableLibre
      end;
    end;

    // Rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
    begin
      if Corresp.value = 'AU1' then
        XX_Rupture.Text := 'T_CORRESP1'
      else
        XX_Rupture.Text := 'T_CORRESP2';
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnNew;
begin
  inherited;
  SupprimeCentralisation;

  // ----------------- Init des composants si pas de Filtre DEFAUT -------------
  if FFiltres.Text = '' then
  begin
    T_NatureAuxi.ItemIndex := 0;
    ModeSelect.Value       := 'FOU';
    E_QualifPiece.Text     := 'N;';
    {JP 04/10/05 : FQ 16149 : si l'établissement est positionné (PositionneEtabUser),
                   il est préférable d'éviter de le "dépositionner" }
    if (ComboEtab.Enabled) and (ComboEtab.ItemIndex < 0) then
      ComboEtab.ItemIndex := 0;
    E_Devise.ItemIndex := 0;
  end;

  OnChangeE_Devise( nil );
  OnClickCBEnSituation( nil );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2003
Modifié le ... : 12/09/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnExitE_General(Sender: TObject);
begin
  if (csDestroying in Ecran.ComponentState) then Exit ;

  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  if ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + BourreEtLess(THEdit(Sender).Text, fbGene) + '"') then
  begin
    if Length(THEdit(Sender).Text) < VH^.Cpta[fbGene].Lg then
      THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbGene);
  end
  else
    THEdit(Sender).ElipsisClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2003
Modifié le ... : 12/09/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnExitE_Auxiliaire(Sender: TObject);
begin
  if (csDestroying in Ecran.ComponentState) then Exit ;

  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  if ExisteSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = "' + BourreEtLess(THEdit(Sender).Text, fbAux) + '"') then
  begin
    if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
      THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
  end
  else
    THEdit(Sender).ElipsisClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnClickAffCentraEche(Sender: TObject);
begin
  CBEnSituation.Enabled  := not AffCentraEche.Checked;
  TCheckBox(GetControl('AFFInfoComp',False)).Enabled := not AffCentraEche.Checked;
  TCheckBox(GetControl('AFFMemo', False)).Enabled := not AffCentraEche.Checked;

  if AffCentraEche.Checked then
  begin
    CBEnsituation.Checked  := False;
    TCheckBox(GetControl('AFFInfoComp',False)).Checked := False;
    TCheckBox(GetControl('AFFMemo', False)).Checked    := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{$IFNDEF GCGC}
procedure TOF_CPGLAUXI.OnChangeFiltre(Sender: TObject);
begin
  inherited;
  CExoRelatifToDates ( E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_, True);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 12/10/2004
Modifié le ... :   /  /
Description .. : Afin de rendre apparent le focus sur la
Suite ........ : THMultiValComboBax, car en readOnly, on perd le curseur
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnEnterQualifPiece(Sender: TObject);
begin
  CSelectionTextControl( Sender ) ;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPGLAUXI.InitChoixTableLibre;
begin
  if CHOIXTYPETALI <> nil then
  begin
    if CHOIXTYPETALI.Value = 'AUX' then
    begin
      SetControlProperty('TABLELIBRE','DataType','TTTABLESLIBRESAUX');
      SetControlProperty('TABLELIBRE','Plus','');
    end
    else if CHOIXTYPETALI.Value = 'CLI' then
    begin
      SetControlProperty('TABLELIBRE','DataType','GCZONELIBRETIE');
      SetControlProperty('TABLELIBRE','Plus','AND CC_CODE LIKE "CT%"');
    end
    else if CHOIXTYPETALI.Value ='SCL' then
    begin
      SetControlProperty('TABLELIBRE','DataType','GCZONELIBRETIE');
      SetControlProperty('TABLELIBRE','Plus','AND CC_CODE LIKE "CR%"');
    end;
  end else
  begin
    SetControlProperty('TABLELIBRE','DataType','TTTABLESLIBRESAUX');
    SetControlProperty('TABLELIBRE','Plus','');
  end;
  TableLibre.ReLoad;
  SetControlText ('LIBREDE','');
  SetControlText ('LIBREA','');
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.AuxiElipsisClick(Sender: TObject);
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPGLAUXI.OnClickChoixTypeTaLi(Sender: TObject);
begin
  InitChoixTableLibre;
  OnChangeTableLibre ( nil );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.MySelectFiltre;
var lExoDate : TExoDate;
begin
  inherited;
  lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(E_Exercice.Value));

  if (CtxExercice.QuelExo(E_DateComptable.Text, False) <> lExoDate.Code) and
     (CtxExercice.QuelExo(E_DateComptable_.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  end;

  // GCO - 06/09/2007 - FQ 21180
  if E_Qualifpiece.Value = '' then
  begin
    E_Qualifpiece.SelectAll;
    E_Qualifpiece.Text := TraduireMemoire('<<Tous>>');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnLoad;
begin
  inherited;

  // Remplissage des Zones GENERAL ET AUXILIAIRE et Vérification des dates
  ControleAvantEdition;

  // Gestion des critères DEV
  GestionCritereDEV;

  // GCO - 19/09/2006 - Gestion BOI (Norme NF)
  FDateFinEdition := StrToDate(E_DateComptable_.Text);

  GestionCentralisation;

  // GestionTableLibre; // FStSelectTL OK, FSTLeftJoinTiersCompl OK

  FStWhereEtat := GenereConditionWhereEtat;

  // Ajout GCO - 09/092003
  //FStWhereEtat := CMajRequeteExercice(E_EXERCICE.Value, FStWhereEtat);

  GenereCumulGL;

{$IFDEF EAGLCLIENT}
{$ELSE}
  // GCO - 20/12/2005 - Correction PB BDE si trop d'ecritures à imprimer 900000
  // GCO - 03/01/2007 - FQ 19375
  if FBoOkDecoupe then
    GestionTailleGrandLivre( FStWhereEtat );
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnUpdate;
begin
  inherited;
  TFQRS1(Ecran).WhereSQL := FStWhereEtat + RenvoiGroupByOrderBy;
  If EdLegale then MajEditionLegal('GLT',CRelatifVersExercice(E_EXERCICE.Value),E_DateComptable.Text,E_DateComptable_.Text);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/04/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.GestionCentralisation;
begin
  SupprimeCentralisation;
  if AFFCentraEche.Checked then
  begin
    if Transactions( EffectueCentralisation, 1 ) <> OeOk then
    begin
      PgiInfo('Erreur lors de la centralisation', Ecran.Caption);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.SupprimeCentralisation;
begin
  try
    if CEstBloqueGLGene then
    begin
      BeginTrans;
      ExecuteSQL('DELETE FROM ECRITURE WHERE E_QUALIFPIECE = "Z" AND E_UTILISATEUR = "' + V_PGI.User + '"');
      CDeBlocageGLGene; // Suppression du verrou dans la Table COURRIER
      CommitTrans;
    end;
  except
    on E: Exception do
    begin
      PgiError('Impossible de supprimer les enregistrements temporaires du Grand Livre', Ecran.Caption);
      RollBack;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.EffectueCentralisation;
var
  lQuery         : TQuery;
  lStSql         : string;
  lSt            : string;
  lNumLigne      : Integer;
  lNumeroPiece   : Integer;
  lStRgrp        : string;
begin
  try
    try
      lQuery := OpenSql('SELECT MAX(E_NUMEROPIECE) LENUM FROM ECRITURE WHERE E_QUALIFPIECE = "Z"', True);
      if lQuery.FindField('LENUM').AsInteger <> 0 then
        lNumeroPiece := lQuery.FindField('LENUM').AsInteger
      else
        lNumeroPiece := 1;
      Ferme(lQuery);

      // Requete de la centralisation des échéances
      lStSql := 'SELECT E_AUXILIAIRE, ' + IIF(FBoAuxiParGene, 'E_GENERAL, ', '') +
                'E_EXERCICE, E_PERIODE, E_JOURNAL, E_DATECOMPTABLE, E_LIBELLE, '  +
                'E_DEVISE, E_ETABLISSEMENT, E_NUMEROPIECE, E_NUMLIGNE, E_QUALIFPIECE, ' +
                'SUM(E_DEBIT) DEBIT, SUM(E_CREDIT) CREDIT, ' +
                'SUM(E_DEBITDEV) DEBITDEV, SUM(E_CREDITDEV) CREDITDEV FROM ECRITURE ' +
                'LEFT JOIN GENERAUX G ON G_GENERAL = E_GENERAL ' +
                'LEFT JOIN TIERS T ON T_AUXILIAIRE = E_AUXILIAIRE ' +
                'LEFT JOIN JOURNAL J ON J_JOURNAL = E_JOURNAL ' +
                'LEFT JOIN TIERSCOMPL ON YTC_AUXILIAIRE = T_AUXILIAIRE WHERE ' +
                'E_NUMECHE >= 1 AND E_ECRANOUVEAU = "N" ' +
                RecupAutresCriteres +
                RecupCriteresCommuns + ' ' +
                'GROUP BY E_AUXILIAIRE, ' + IIF(FBoAuxiParGene, 'E_GENERAL, ', '') +
                'E_EXERCICE, E_PERIODE, E_JOURNAL, E_DATECOMPTABLE, E_LIBELLE, E_DEVISE, E_ETABLISSEMENT,' +
                'E_NUMEROPIECE, E_NUMLIGNE, E_QUALIFPIECE ' + 
                'ORDER BY E_AUXILIAIRE, ' + IIF(FBoAuxiParGene, 'E_GENERAL, ', '') +
                'E_EXERCICE, E_PERIODE, E_JOURNAL';

      // GCO - 20/11/2007 - FQ 18875
      if EstMultiSoc then
      begin
        RequeteMultiDossier( Pages, lStRgrp );
        lQuery := OpenSelect( lStSql, lStRgrp);
      end
      else
        lQuery := OpenSql(lStSql, True);

      if not lQuery.Eof then
      begin
        // Création d'un verrou en Table COURRIER pour la présence de
        // Centralisation des échéances dans la Table ECRITURE
        CBlocageGLGene;
        InitMoveProgressForm(Ecran, Ecran.Caption, 'Traitement de la centralisation des échéances...', RecordsCount(lQuery), True, True);

        lNumLigne := 1;
        while not lQuery.Eof do
        begin
          if not MoveCurProgressForm('Compte ' + lQuery.FindField('E_AUXILIAIRE').AsString + ' - ' + 'Centralisation des échéances') then
            Break;

          if FBoAuxiParGene then
            lSt := 'INSERT INTO ECRITURE (E_UTILISATEUR, E_AUXILIAIRE,  E_GENERAL, '
          else
            lSt := 'INSERT INTO ECRITURE (E_UTILISATEUR, E_AUXILIAIRE, ';

          lSt := lSt + 'E_EXERCICE, E_JOURNAL, ' +
                 'E_DATECOMPTABLE, E_PERIODE, E_QUALIFPIECE, E_LIBELLE, E_DATEMODIF, ' +
                 'E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_ECRANOUVEAU,' +
                 'E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV, E_CONFIDENTIEL, E_CREERPAR, ' +
                 'E_QTE1, E_QTE2, E_DATEECHEANCE, E_LETTRAGE, E_DEVISE, E_ETABLISSEMENT) ' +
                 'VALUES (' +
                 '"' + string(V_PGI.User) + '",' +
                 '"' + lQuery.FindField('E_AUXILIAIRE').AsString + '",';

          if FBoAuxiParGene then
            lSt := lSt + '"' + lQuery.FindField('E_GENERAL').AsString + '",';

          lSt := lSt + '"' + lQuery.FindField('E_EXERCICE').AsString + '",' +
                 '"' + lQuery.FindField('E_JOURNAL').AsString + '",' +
                 '"' + UsDateTime(lQuery.FindField('E_DATECOMPTABLE').AsDateTime) + '",' +
                 '"' + lQuery.FindField('E_PERIODE').AsString + '",' +
                 '"Z", "' + lQuery.FindField('E_LIBELLE').AsString + '", ' +
                 '"' + UsDateTime(lQuery.FindField('E_DATECOMPTABLE').AsDateTime) + '",' +
                 IntToStr( lNumeroPiece ) + ',' + IntToStr(lNumLigne) + ', 0, "N",' +
                 VariantToSql(lQuery.FindField('DEBIT').AsFloat) + ',' +
                 VariantToSql(lQuery.FindField('CREDIT').AsFloat) + ',' +
                 VariantToSql(lQuery.FindField('DEBITDEV').AsFloat) + ',' +
                 VariantToSql(lQuery.FindField('CREDITDEV').AsFloat) + ',' +
                 '"0", "SAI", ' + // E_CONFIDENTIEL A ZERO, E_CREERPAR = SAI
                 '"0", "0", "' + UsDateTime( iDate1900 ) + '",' +
                 '"####",' +'"' + lQuery.FindField('E_DEVISE').AsString + '",' +
                 '"' + lQuery.FindField('E_ETABLISSEMENT').AsString + '")' ;

          ExecuteSql(lSt);
          Inc(lNumLigne);
          lQuery.Next;
        end;
      end;
    except
      on E: Exception do PgiError(E.Message, Ecran.Caption);
    end;
  finally
    Ferme(lQuery);
    FiniMoveProgressForm;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.CBlocageGLGene;
begin
  try
    if not CEstBloqueGLGene then
    begin
      ExecuteSQL('INSERT INTO COURRIER(MG_UTILISATEUR, MG_TYPE, MG_LIBELLE, '+
                 'MG_DATE, MG_EXPEDITEUR) VALUES("' + W_W + '",4000, ' +
                 '"Centralisation des échéances","'+USTime(Now)+'","' + V_PGI.User + '")');
    end;
  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CBlocageGLGene');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.CDeBlocageGLGene;
begin
  try
    ExecuteSQL('DELETE FROM COURRIER WHERE MG_UTILISATEUR = "' + W_W + '" ' +
               'AND MG_TYPE=4000 ' +
               'AND MG_EXPEDITEUR = "' + String(V_PGI.User) + '"');
  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CDeBlocageGLGene');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/04/2006
Modifié le ... : 03/04/2006
Description .. : Permet de savoir si l'on a généré des écritures typées "Z"
Suite ........ : pour ne pas exécuter la requête DELETE inutilement.
Mots clefs ... :
*****************************************************************}
function TOF_CPGLAUXI.CEstBloqueGLGene: Boolean;
var lStSql  : string ;
    lQuery  : TQuery ;
begin
  Result := False;

  lStSql := 'SELECT MG_EXPEDITEUR FROM COURRIER WHERE ' +
            'MG_UTILISATEUR = "' + W_W + '" AND MG_TYPE=4000 AND ' +
            'MG_EXPEDITEUR = "' + String(V_PGI.User) + '"';

  lQuery := nil;
  try
    try
      lQuery := OpenSql(lStSql,True) ;
      Result := not lQuery.Eof ;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CEstBloqueGLGene');
    end;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLCLIENT}
{$ELSE}
procedure TOF_CPGLAUXI.GestionTailleGrandLivre(vStwhereCritere: string);
var lStSelect   : string;          //
    lStGroupBy  : string;          //
    lStMini     : string;          //
    lStCompte   : string;          // Fourchette de comte que l'on ajoute au Where
    lStFileName : string;          // Nom du fichier temp du PDFBatch
    lBoOldNoPrintDialog : Boolean; //
    lBoOldQRPdf         : Boolean;
    lBoAbort            : Boolean;
    lNumPage, lNbEcr, i : integer; //
    lTotalPage          : integer;
    lNbEcrMax           : integer; // Nombre d'écritures MAX avant de couper le GL
    lTobTrancheCompte   : Tob;     // Liste des fourchettes de Compte
    lTobDecoupage       : Tob;     // Nb d'écritures à imprimer par Compte
    lTobTemp            : Tob;     // Tob de Traitement
begin
  TFQRS1(Ecran).OnValideCritereEvent := nil; // Conditionne le CritOk du TQRS1

  lBoAbort := False;

  lTobTrancheCompte := Tob.Create('', nil, -1);

  lStSelect  := 'SELECT E_AUXILIAIRE';
  lStGroupBy := ' GROUP BY E_AUXILIAIRE';

  // GCO - 27/10/2006 - 18923
  //if FBoAuxiParGene then
  //begin
  //  lStSelect := lStSelect + ', E_GENERAL';
  //  lStGroupBy := lStGroupBy + ', E_GENERAL';
  //end;

  lStSelect := lStSelect + ', COUNT(*) TOTAL FROM ECRITURE ' +
               'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
               'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
               'LEFT JOIN JOURNAL ON J_JOURNAL  = E_JOURNAL ' +
               'LEFT JOIN TIERSCOMPL ON YTC_AUXILIAIRE = T_AUXILIAIRE ';

  lStSelect := lStSelect + ' WHERE ';

  lTobDecoupage := Tob.Create('', nil, -1);
  lTobDecoupage.LoadDetailFromSQL( lStSelect + vStWhereCritere + lStGroupBy + ' ORDER BY E_AUXILIAIRE');

  // GCO - 26/12/2005 - Découpage possible si au moins deux comptes édités
  if lTobDecoupage.Detail.Count > 1 then
  begin
    lNbEcr := 0;
    lNbEcrMax := GetSynRegKey('CPGLENREGMAX', cNbEcrMax, True);

    lStMini := E_Auxiliaire.Text;
    for i := 0 to lTobDecoupage.Detail.Count-1 do
    begin
      lNbEcr := lNbEcr + lTobDecoupage.Detail[i].GetInteger('TOTAL');
      if lNbEcr >= lNbEcrMax then
      begin
        lNbEcr := 0;
        lTobTemp := Tob.Create('', lTobTrancheCompte, -1);
        lTobTemp.AddChampSupValeur('MINI', lStMini);
        lTobTemp.AddChampSupValeur('MAXI', lTobDecoupage.Detail[i].GetString('E_AUXILIAIRE'));
        if i <> lTobDecoupage.Detail.Count-1 then
          lStMini := lTobDecoupage.Detail[i+1].GetString('E_AUXILIAIRE')
        else
          lStMini := E_Auxiliaire_.Text; // Dernière ligne
      end;
    end;

    // Faut Faire la dernière tranche
    if lTobTrancheCompte.Detail.Count <> 0 then
    begin
      lTobTemp := Tob.Create('', lTobTrancheCompte, -1);
      lTobTemp.AddChampSupValeur('MINI', lStMini);
      lTobTemp.AddChampSupValeur('MAXI', E_Auxiliaire_.Text);
    end;

    // Edition découpée de l'état
    if lTobTrancheCompte.Detail.Count <> 0 then
    begin
      lBoOldNoPrintDialog := V_Pgi.NoPrintDialog;
      lBoOldQRPdf         := V_Pgi.QRPdf;

      V_Pgi.QRPdf := True;

      // CritOk devient FALSE, Plus de OnUpdate pour le QRS1
      TFQRS1(Ecran).OnValideCritereEvent := OnMyValideCritereEvent;

      lStFileName := TempFileName();
    {$IFDEF EAGLCLIENT}
      StartPdfBatch(lStFileName);
    {$ELSE}
      THPrintBatch.StartPdfBatch(lStFileName);
    {$ENDIF}

      // Changement de la CLAUSE WHERE de l'état
      lTotalPage := 0;
      for i := 0 to lTobTrancheCompte.Detail.count -1 do
      begin
        if not lBoAbort then
        begin
          OkEntete.Text := IIF(i = 0, 'X', '-');
          OkResume.Text := IIF(i = lTobTrancheCompte.Detail.count -1, 'X', '-');
          lStCompte := ' AND E_AUXILIAIRE >= "' + lTobTrancheCompte.Detail[i].GetString('MINI') + '" AND ' +
                       'E_AUXILIAIRE <= "' + lTobTrancheCompte.Detail[i].GetString('MAXI') + '"';

        {$IFDEF EAGLCLIENT}
          lNumPage := LanceEtatTob('E', TFQRS1(Ecran).NatureEtat, TFQRS1(Ecran).CodeEtat, nil, CBPreview.Checked, False, False, Pages,  vStWhereCritere
           + lStCompte + RenvoiGroupByOrderBy, Ecran.caption, False,0,'', lTotalPage);
        {$ELSE}
          lNumPage := LanceEtatTob('E', TFQRS1(Ecran).NatureEtat, TFQRS1(Ecran).CodeEtat, nil, CBPreview.Checked, False, False, Pages,  vStWhereCritere
           + lStCompte + RenvoiGroupByOrderBy, Ecran.caption, False,0,'',nil, lTotalPage);
        {$ENDIF}
          // GCO - 13/09/2006 - FQ 18667
          lTotalPage := lTotalPage + lNumPage;

          if lNumPage = -1 then
            lBoAbort := True;

          if not CBPreview.Checked then
            V_PGI.NoPrintDialog := True;

          THPrintBatch.AjoutPdf(V_PGI.QrPdfQueue, True);
        end;
      end;

    {$IFDEF EAGLCLIENT}
      CancelPDFBatch;
    {$ELSE}
      THPrintBatch.StopPdfBatch();
    {$ENDIF}

      if CbPreview.Checked then
        PreviewPDFFile(Ecran.Caption, lStFileName, True)
      else
        PrintPDF(lStFileName, Printer.Printers[Printer.PrinterIndex],'');

      // Remise en place du Contexte
      V_Pgi.NoPrintDialog := lBoOldNoPrintDialog;
      V_Pgi.QRPdf := lBoOldQRPdf;
    end;
  end;

  FreeAndNil(lTobDecoupage);
  FreeAndNil(lTobTrancheCompte);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLAUXI.ControleAvantEdition;
var lStSql : string;
    lSql   : TQuery;
begin
  // Remplissage des bornes des comptes auxiliaires
  if (Trim(E_Auxiliaire.Text) = '') or (Trim(E_Auxiliaire_.Text) = '') then
  begin
    {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
    if not TestJoker(E_Auxiliaire.Text) then
    begin
      try
        lStSql := 'SELECT MIN(T_AUXILIAIRE)MINI, MAX(T_AUXILIAIRE)MAXI FROM TIERS';

        // Nature du Compte Auxiliaire
        if T_NATUREAUXI.ItemIndex <> 0 then
          lStSql := lStSql + ' WHERE T_NATUREAUXI = "' + T_NATUREAUXI.Value + '"';

        lSql := OpenSql( lStSql, True);
        if not lSql.Eof then
        begin
          if Trim(E_Auxiliaire.Text) = '' then
            E_Auxiliaire.Text := lSql.FindField('MINI').AsString;

          if Trim(E_Auxiliaire_.Text) = '' then
            E_Auxiliaire_.Text := lSql.FindField('MAXI').AsString;
        end;

      finally
        Ferme(lSql);
      end;
    end;
  end;

  if FBoAuxiParGene then
  begin
    if (Trim(E_General.Text) = '') or (Trim(E_General_.Text) = '') then
    begin
      {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
      if not TestJoker(E_General.Text) then begin
        try
          lSql := OpenSql('SELECT MIN(G_GENERAL)MINI, MAX(G_GENERAL)MAXI FROM GENERAUX WHERE G_COLLECTIF = "X"', True);
          if Trim(E_General.Text) = '' then
            E_General.Text := lSql.FindField('MINI').AsString;

          if Trim(E_General_.Text) = '' then
            E_General_.Text := lSql.FindField('MAXI').AsString;
        finally
          Ferme(lSql);
        end;
      end;
    end;
  end;

  if E_Exercice.ItemIndex = 0 then
  begin
    if VH^.ExoV8.Code <> '' then
    begin
      if StrToDate(E_DateComptable.Text) < VH^.ExoV8.Deb then
      begin
        PgiInfo('Vous ne pouvez pas saisir une date inférieure au ' + DateToStr(VH^.ExoV8.Deb) + '.' + #13 + #10 +
                'La date a été modifiée.', Ecran.Caption);
        E_DateComptable.Text := DateToStr(VH^.ExoV8.Deb);
      end;
    end
    else
    begin
      // GCO - 02/12/2003 FB 12796
      if StrToDate(E_DateComptable.Text) < VH^.Exercices[1].Deb then
      begin
        PgiInfo('Vous ne pouvez pas saisir une date inférieure au ' + DateToStr(VH^.Exercices[1].Deb) + '.' + #13 + #10 +
                'La date a été modifiée.', Ecran.Caption);
        E_DateComptable.Text := DateToStr(VH^.Exercices[1].Deb);
      end;
    end;

    if CbEnSituation.Checked then
      FExoDate := ctxExercice.QuelExoDate(StrToDate(EDateSituation.Text))
    else
      FExoDate := CtxExercice.QuelExoDate(StrToDate(E_DateComptable.Text));
  end
  else
  begin
    FExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(E_Exercice.Value));
    if (StrToDate(E_DateComptable.Text) < FExoDate.Deb) then
    begin
      PgiInfo('Vous ne pouvez pas saisir une date inférieure au ' + DateToStr(FExoDate.Deb) + '.' + #13 + #10 +
              'La date a été modifiée.', Ecran.Caption);
      E_DateComptable.Text := DateToStr(FExoDate.Deb);
    end;
  end;

  // Correction du E_QUALIFPIECE
  if E_QualifPiece.Text = '' then
    E_QualifPiece.Text := 'N;';

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLAUXI.RenvoiGroupByOrderBy: string;
begin
  Result := '';
 
  if (Rupture.ItemIndex = 0) or ((Rupture.ItemIndex <> 0) and (RuptureType.ItemIndex = 0)) then
  begin
    Result := Result + ' ORDER BY E_AUXILIAIRE, ' + IIF( FBoAuxiParGene, 'E_GENERAL, ', '');
  end
  else
  begin
    if XX_Rupture.Text <> '' then
      Result := Result + ' ORDER BY ' + XX_Rupture.Text + ', E_AUXILIAIRE, ' +
                    IIF( FBoAuxiParGene, 'E_GENERAL, ', '')
    else
      Result := Result + ' ORDER BY E_AUXILIAIRE, ' +
                    IIF( FBoAuxiParGene, 'E_GENERAL, ', '')
  end;

  // GCO - 25/05/2005 - FQ 15707 - Pas de tri sur exercice si un exo selectionné
  // et A-Nouveaux détaillés.
  if (E_Exercice.ItemIndex = 0) or
     ((E_Exercice.ItemIndex <> 0) and (AFFAnouveau.Checked)) then
    Result := Result + 'E_DATECOMPTABLE, E_EXERCICE, E_QUALIFPIECE, E_PERIODE, ' +
              'E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE'
  else
    Result := Result + 'E_EXERCICE, E_DATECOMPTABLE, E_QUALIFPIECE, E_PERIODE, ' +
              'E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE';
end;

(*
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLAUXI.GestionTableLibre: string;
begin
  FStSelectTL := '';
  FStLeftJoinTIERSCOMPL := '';

  // Initialisation des champs du SELECT en fonction des tables libres choisies
  if CHOIXTYPETALI <> nil then
  begin
    if CHOIXTYPETALI.Value = 'AUX' then
      FStSelectTL := 'T_TABLE0, T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5,T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9,'
    else
    if CHOIXTYPETALI.Value = 'CLI' then
      FStSelectTL := 'YTC_TABLELIBRETIERS1, YTC_TABLELIBRETIERS2, YTC_TABLELIBRETIERS3, YTC_TABLELIBRETIERS4'+
                     ', YTC_TABLELIBRETIERS5, YTC_TABLELIBRETIERS6,YTC_TABLELIBRETIERS7, YTC_TABLELIBRETIERS8'+
                     ', YTC_TABLELIBRETIERS9, YTC_TABLELIBRETIERSA,'
    else
    if CHOIXTYPETALI.Value = 'SCL' then
      FStSelectTL := 'YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3,';

    if FStSelectTL[1] = 'Y' then
      FStLeftJoinTIERSCOMPL := ' LEFT JOIN TIERSCOMPL ON YTC_AUXILIAIRE = T_AUXILIAIRE'
    else
      FStLeftJoinTIERSCOMPL := '';
  end
  else
  begin
    FStSelectTL := 'T_TABLE0, T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5,T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9,';
    FStLeftJoinTIERSCOMPL := '';
  end;
end;
*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLAUXI.OnClose;
begin
  SupprimeCentralisation;
  if Assigned(FTobCumulAnoGL)   then FreeAndNil(FTobCumulAnoGL);
  if Assigned(FTobCumulAuGL)    then FreeAndNil(FTobCumulAuGL);
  if Assigned(FTobBilanGestion) then FreeAndNil(FTobBilanGestion);
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  registerclasses([TOF_CPGLAUXI]);
end.

