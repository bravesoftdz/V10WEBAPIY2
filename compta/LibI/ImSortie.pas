{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/06/2003
Modifié le ... : 12/10/2004
Description .. : - CA - 02/06/2003 - Calcul des + ou - values : prise en
Suite ........ : compte de l'exceptionnel
Suite ........ : - FQ 12440 - CA - 02/07/2003 - Contrôle des racines des
Suite ........ : comptes généraux
Suite ........ : - FQ 14664 - CA - 12/10/2004 - Affichage de la VNC réelle
Suite ........ : et non de la valeur résiduelle
Suite ........ : BTY 07/05 FQ 16295 Formater montants avec nb décimales de la monnaie du dossier
Suite ........ : YCP 12/06/2006  date ope dans la serie date entree par
Suite......... : - FQ 15280 - MBO - 30/09/2005 - saisie exceptionnel sur immo avec dérog sauf plan variable
Suite......... : - FQ 16776 - MBO - 30/09/2005 - touche F6 pour récup exceptionnel maximum
Suite......... : - MBO - 24/10/2005 - crc 2002/10 appel getexcepexercice avec depreciation
Suite......... : - FQ 16804 - MBO - 28/11/2005 - autoriser exceptionnel sortie sur plan variable
Suite......... : - FQ 17093 - MBO - 28/11/2005 - affichage valeurs correctes suite erreur dans la saisie
Suite......... : - MBO - 05/12/2005 - dans le cas d'une reprise inférieure à dotation calculée :
                                      affichage de l'exceptionnel non signé
Suite......... : - FQ 17155 - MBO - 07/12/2005 - pb comparaison entre 2 montants identiques
Suite......... : - BTY - 01/06 FQ 17259 Nouveau top dépréciation dans IMMO
                 - TGA - 23/01/2006 Ajout champs sur table IMMO
Suite......... : - BTY - 04/06 FQ 17516 Nouveau top changement de regroupement dans IMMO
Suite......... : - BTY - 04/06 FQ 17629 Choix pour déterminer la règle de calcul de la PM Value
Suite......... : - 0 CT Tout en court terme
Suite......... : - 1 LT Tout en long terme
Suite......... : - 2 NOR Règle normale (comme avant)
Suite......... : - 3 RSD Règle normale en ignorant la durée (<=> NOR sans tenir compte de l'âge)
Suite......... : - MBO - 22/05/2006 - fq 18213 - pb masque sur zone dotation si saisie except négatif
Suite......... : - MBO - 15/06/2006 - FQ 18395 - masquer la saisie exceptionnel si date deb eco > exo en cours
Suite......... : - BTY - 10/06 - FQ 18937 Divers pbs après cession dans certains cas :
Suite......... :   PValue / VNC erronés / absence enreg CES dans IMMOLOG
Suite......... :   Causes PutValue trop complexes et stockage du BlocNotes non standard
Suite......... : - MBO - 10/06 - Gestion de la prime et de la subvention
Suite......... : - BTY - 10/06 - FQ 18916 En répartition libre sur un regroupement, contrôler
                   que le total réparti est égal au prix de sortie
Suite......... : - MBO - 09/07/2007 - ds le cas de sortie de regroupement : masquer l'entête de la fenêtre
Suite......... : - MBO - 23/07/2007 - masquer le motif de sortie Remplacement composant
Suite..........: - BTY - 09/07 - Combo des motifs de sortie : filtre Plus pour exclure 999 Remplac. composant 2e catégorie
Suite..........: -  attention va avec &#@  nouveau paramétrage de la tablette TIMOTIFCESSION
Suite..........: - MBO - 29/10/07 - FQ 21746 -  coche "sortie calculée" en assistance plus visible et plus utilisée
Suite..........: - MBO - 31/10/07 - FQ 21754 - appel planInfo.Calcul avec nveau paramètre pour prise en cpte du paramètre
                                               'amortissement du jour de cession'

*****************************************************************}
unit ImSortie;

// CA - 09/07/1999 - Calcul VNC sans amortissement à la sortie
interface

uses
  windows, ParamSoc,
  Classes,
  StdCtrls,
  Forms,
  HSysMenu,
  ComCtrls,
  HRichEdt,
  HRichOLE,
  Mask,
  Controls,
  ExtCtrls,
  SysUtils,
  hmsgbox,
  Hctrls,
  HTB97,
  HPanel,
  HEnt1,
  db,
  utob,
  {$IFDEF EAGLCLIENT}
  etablette,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  tablette,
  {$ENDIF}
  ImAncOpe,
  ImPlan,
  ImEnt,
  Grids,
  ImPlanInfo,
  LicUtil,
  AmSortie
  ;


type
  TEnregCession = class
    Code : string;
    Vo : double;
    QteAchat : double;
    QteCedee : double;
    DateOpe : TDateTime;
    MontantHt: double ;
    MontantAchat : double;
    MontantCession : double;
    Tva : double;
    TvaR : double;
    TvaE : double;
    AmortEco : double;
    AmortFisc : double;
    RepriseEco : double;
    RepriseFisc : double;
    VNCCedee : double;
    PValue : double;
    MotifCession : string;
    CalcCession : string;
    TypeCession : string;
    BlocNote : HTStrings;
    PlanPrec : integer;
    PlanSuiv : integer;
    ChangePlan : boolean;
    TypeExc : string;
    MontantExc : double;
    Dotation : double;
    BaseEco : double;
    BaseFisc : double;
    BaseTP: double ;
    Reglecession : string;
    ReprisePri : double;
    BasePri : double;
    RepriseSBV : double;
end;
  TFSortie = class(TFAncOpe)
    HM2: THMsgBox;
    HLabel4: THLabel;
    IL_MOTIFCES: THValComboBox;
    HLabel5: THLabel;
    IL_CALCCESSION: THValComboBox;
    HLabel6: THLabel;
    IL_VOCEDEE: THNumEdit;
    lIL_QTECEDEE: THLabel;
    IL_QTECEDEE: THNumEdit;
    HLabel8: THLabel;
    IL_MONTANTCES: THNumEdit;
    lIL_TVAAREVERSER: THLabel;
    IL_TVAAREVERSER: THNumEdit;
    PAMORTISSEMENT: TTabSheet;
    GBAmortissement: TGroupBox;
    HLabel10: THLabel;
    HLabel11: THLabel;
    HLabel12: THLabel;
    HLabel13: THLabel;
    VNC: THNumEdit;
    DotPartieCedee: THNumEdit;
    rbAjouter: TRadioButton;
    rbRemplacer: TRadioButton;
    DotationCalc: THNumEdit;
    Excep: THNumEdit;
    Dotation: THNumEdit;
    HPanel4: THPanel;
    PREPARTITION: TTabSheet;
    HLabel2: THLabel;
    MODEREPARTITION: THValComboBox;
    LISTEREPARTITION: THGrid;
    PANELREGROUPEMENT: THPanel;
    BSORTIECALCULEE: TCheckBox;
    HLabel9: THLabel;
    ReglePValue: THValComboBox;
    procedure DATEOPEExit(Sender: TObject);
    procedure DATEOPEEnter(Sender: TObject);
    procedure AccelCreationMotifClick(Sender: TObject);
    procedure OnModifDotation(Sender: TObject);
    procedure IL_CALCCESSIONExit(Sender: TObject);
    procedure DotationKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure IL_TVAAREVERSEREnter(Sender: TObject);
    procedure IL_TVAAREVERSERExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure IL_MONTANTCESExit(Sender: TObject);
    procedure IL_CALCCESSIONChange(Sender: TObject);
    procedure MODEREPARTITIONChange(Sender: TObject);
    procedure LISTEREPARTITIONKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure LISTEREPARTITIONCellExit(Sender: TObject; var ACol,
      ARow: Integer; var Cancel: Boolean);
  private
    { Déclarations privées }
    fVo : double;
    fPlanAvantCession : TPlanAmort;
    fQteCedee, fTva, fMontantCession, fVNC, fExcep: double;
    fTypeExc : string;
    fCurDateOp : string;
    fCurVCedee : double;
    fRegroupement : string;
    fTRepartition : TOB;

    // ajout mbo  fq 15280
    DotTest : double;
    DotEcoExe : double;
    DotFiscExe : double;
    MaxAjout : double;
    MaxRemplace : double;
    Derog : boolean;
    initVNCeco : double;
    initVNCfisc: double;

    AmortJsortie : string;  //fq 21754

    procedure RecalculDotation;
    procedure RecalculDetailDotation;
    procedure EnregCession ;
    procedure GereDotationMax;
    procedure InitGrilleRepartition;
    procedure CalculRepartition ( bRecalculVNC : boolean );
    // mbo - 30/10/2007 procedure EnregistreLaSortieRegroupement(CodeImmo: string; PrixSortie, VNC: double); // n'est plus utilisée
    procedure SoldePrixSortie ( bDerniereLigne : boolean );
    procedure ExecuteSortie;
    // BTY 07/05 FQ 16295 Formater montants avec nb décimales de la monnaie du dossier
    procedure FormaterDecimales ;
  public
    { Déclarations publiques }
    procedure InitZones; override;
    function ControleZones : boolean; override;
  end;

function ExecuteCession(Code : string) : TModalResult;
function ExecuteCessionRegroupement ( Regroupement : string ) : TModalResult;
procedure TraiteImmoCedee(var PlanImmoCedee : TPlanAmort;var EnregCession : TEnregCession);
procedure TraiteImmoOrig(PlanImmoCedee : TPlanAmort; CessionPartielle : boolean;var EnregCession : TEnregCession);

// n'est plus utlisée (appelée dans opeserie mais opeserie pour cession n'existe pas)
procedure EnregLogCession(PlanImmoCedee : TPlanAmort;CessionPartielle : boolean; EnregCession : TEnregCession; OrdreS : integer);

procedure TraiteChangementPlan (Plan : TPlanAmort);
procedure ExecuteChangementPlan (var PlanCede : TPlanAmort; var EnregCession : TEnregCession);

implementation

uses ImOuPlan,Outils,ChanPlan;

{$R *.DFM}

function ExecuteCession(Code : string) : TModalResult;
var
  FSortie: TFSortie;
begin
  FSortie:=TFSortie.Create(Application) ;
  FSortie.fCode:=Code ;
  FSortie.fRegroupement := '';
  FSortie.fProcEnreg:=FSortie.EnregCession;
  try
    FSortie.ShowModal ;
  finally
    result := FSortie.ModalResult;
    FSortie.Free ;
  end ;
end;

function ExecuteCessionRegroupement ( Regroupement : string ) : TModalResult;
var
  FSortie: TFSortie;
begin
  FSortie:=TFSortie.Create(Application) ;
  FSortie.fCode := '';
  FSortie.fRegroupement := Regroupement ;
  FSortie.fProcEnreg:=FSortie.EnregCession;
  try
    FSortie.ShowModal ;
  finally
    result := FSortie.ModalResult;
    FSortie.Free ;
  end ;
end;

procedure TFSortie.FormCreate(Sender: TObject);
begin
  inherited;
  fPlanAvantCession:=TPlanAmort.Create(true) ;
{$IFDEF SERIE1}
HelpContext:=511030 ;
{$ELSE}
HelpContext:=2111300 ;
{$ENDIF}
end;

procedure TFSortie.FormShow(Sender: TObject);

begin
  inherited;
  FVo:=IL_VOCEDEE.Value;
  FQteCedee:=IL_QTECEDEE.Value ;
  // BTY 07/05 FQ 16295 Formater montants avec nb décimales de la monnaie du dossier
  FormaterDecimales;
  if fRegroupement <> '' then
  begin
    IL_CALCCESSION.Items.Delete(IL_CALCCESSION.Values.IndexOf('NOR'));
    IL_CALCCESSION.Values.Delete(IL_CALCCESSION.Values.IndexOf('NOR'));
  end;
  InitGrilleRepartition;
  // mbo 26.09.05 RecalculDotation;
  //RecalculDetailDotation;

end;

procedure TFSortie.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  fTRepartition.Free;
  fPlanAvantCession.free ;
  inherited;
end;

// BTY 07/05 FQ 16295 Formater montants avec nb décimales de la monnaie du dossier
procedure TFSortie.FormaterDecimales;
var
  Masque : string;
begin
  Masque := StrfMask(V_PGI.OkDecV, '', True);

  IL_VOCEDEE.Masks.PositiveMask := Masque;
  IL_VOCEDEE.Masks.NegativeMask := Masque;
  IL_VOCEDEE.Masks.ZeroMask := Masque;

  IL_Montantces.Masks.PositiveMask := IL_VOCEDEE.Masks.PositiveMask ;
  IL_Montantces.Masks.NegativeMask := IL_VOCEDEE.Masks.NegativeMask ;
  IL_Montantces.Masks.ZeroMask := IL_VOCEDEE.Masks.ZeroMask ;

  IL_TVAAReverser.Masks.PositiveMask := IL_VOCEDEE.Masks.PositiveMask ;
  IL_TVAAReverser.Masks.NegativeMask := IL_VOCEDEE.Masks.NegativeMask ;
  IL_TVAAReverser.Masks.ZeroMask := IL_VOCEDEE.Masks.ZeroMask ;

  VNC.Masks.PositiveMask := IL_VOCEDEE.Masks.PositiveMask ;
  VNC.Masks.NegativeMask := IL_VOCEDEE.Masks.NegativeMask ;
  VNC.Masks.ZeroMask := IL_VOCEDEE.Masks.ZeroMask ;

  Dotation.Masks.PositiveMask := IL_VOCEDEE.Masks.PositiveMask ;
  // fq 18213 - mbo - 22/05/2006 Dotation.Masks.NegativeMask := IL_VOCEDEE.Masks.NegativeMask ;
  Dotation.masks.NegativeMask := '-' + Masque;
  Dotation.Masks.ZeroMask := IL_VOCEDEE.Masks.ZeroMask ;

  DotPartieCedee.Masks.PositiveMask := IL_VOCEDEE.Masks.PositiveMask ;
  DotPartieCedee.Masks.NegativeMask := IL_VOCEDEE.Masks.NegativeMask ;
  DotPartieCedee.Masks.ZeroMask := IL_VOCEDEE.Masks.ZeroMask ;

  DotationCalc.Masks.PositiveMask := IL_VOCEDEE.Masks.PositiveMask ;
  DotationCalc.Masks.NegativeMask := IL_VOCEDEE.Masks.NegativeMask ;
  DotationCalc.Masks.ZeroMask := IL_VOCEDEE.Masks.ZeroMask ;

  Excep.Masks.PositiveMask := IL_VOCEDEE.Masks.PositiveMask ;
  // mbo 5/12/2005  Excep.Masks.NegativeMask := '-' + IL_VOCEDEE.Masks.NegativeMask ; // modif mbo 15280
  Excep.masks.NegativeMask := '-' + Masque;

  Excep.Masks.ZeroMask := IL_VOCEDEE.Masks.ZeroMask ;

  if fRegroupement <> '' then
  begin
     { Colonne n°2 = Valeur d'achat }
     LISTEREPARTITION.ColTypes[2]:='R';
     LISTEREPARTITION.ColFormats[2]:= Masque;
     { Colonne n°3 = VNC }
     LISTEREPARTITION.ColTypes[3]:='R';
     LISTEREPARTITION.ColFormats[3]:= Masque;
     { Colonne n°4 = Prix de sortie }
     LISTEREPARTITION.ColTypes[4]:='R';
     LISTEREPARTITION.ColFormats[4]:= Masque;
  end;
end;
//

procedure TFSortie.InitZones;
var QImmo : TQuery;
    bGestionTVA : boolean;
    DateOpe : TDateTime;

begin
  inherited;
  // fq 21746 - mbo - 26.10.07 - BSORTIECALCULEE.Visible :=  (V_PGI.PassWord = CryptageSt(DayPass(Date)));
  BSORTIECALCULEE.Visible := false;
  bGestionTVA := False;
  PAMORTISSEMENT.TabVisible := (fRegroupement = '');
  PREPARTITION.TabVisible := (fRegroupement <> '');
  PANELREGROUPEMENT.Visible := (fRegroupement <> '');

  //ajout mbo car le panelregroupement ne masque plus les zones en dessous - 09.07.07
  tI_IMMO.Visible :=(fRegroupement = '');
  tI_LIBELLE.Visible :=(fRegroupement = '');
  tI_MONTANTHT.Visible :=(fRegroupement = '');
  tI_DATEPIECEA.Visible :=(fRegroupement = '');
  HLabel1.Visible :=(fRegroupement = '');
  HLabel15.Visible :=(fRegroupement = '');
  laI_IMMO.Visible :=(fRegroupement = '');
  laI_DATEPIECEA.visible :=(fRegroupement = '');
  laI_COMPTEIMMO.visible :=(fRegroupement = '');
  laI_MONTANTHT.visible :=(fRegroupement = '');
  laI_QUANTITE.visible :=(fRegroupement = '');
  laI_LIBELLE.visible :=(fRegroupement = '');

  Derog := false; //ajout mbo - fq 15280

  //ajout mbo pour fq 21754 - 30/10/2007
  AmortJsortie := 'OUI';
  if (GetParamSocSecur('SO_AMORTJSORTIE', True) = cbUnChecked) then
      AmortJsortie := 'NON';

  // ajout mbo 18.04.2007 masquer le motif Remplacement composant de 2ème catégorie
  // mbo 23.07.07 IL_MOTIFCES.Items.Delete(IL_MOTIFCES.Items.IndexOf('Rempl. composant 2ème catégorie'));
  // BTY 09/07 Exclure 999 via la propriété Plus de la combo
  //IL_MOTIFCES.Items.Delete(IL_MOTIFCES.Values.IndexOf('999'));

  // BTY 07/05 FQ 16295
  FormaterDecimales;
  //
  if fRegroupement <> '' then
  begin
    PAGES.ActivePage := PREPARTITION;
    lIL_QTECEDEE.Visible := False;
    IL_QTECEDEE.Visible := False;
    lIL_TVAAREVERSER.Visible := False;
    IL_TVAAREVERSER.Visible := False;
    fTRepartition := TOB.Create ('',nil,-1);
    PANELREGROUPEMENT.Caption := fRegroupement+' - '+RechDom('AMREGROUPEMENT',fRegroupement,False);
    MODEREPARTITION.Value := 'VNC';
    fTRepartition.LoadDetailFromSQL('SELECT I_IMMO,I_LIBELLE,I_MONTANTHT FROM IMMO WHERE I_GROUPEIMMO="'+fRegroupement+'"');
    CalculRepartition ( False );
  end else PAGES.ActivePage := PAMORTISSEMENT;
(*  if Presence ( 'GENERAUX','G_GENERAL',ImBourreLaDoncSurLesComptes('462') ) then
    if Presence ( 'GENERAUX','G_GENERAL',ImBourreLaDoncSurLesComptes('775') ) then
      if Presence ( 'GENERAUX','G_GENERAL',ImBourreLaDoncSurLesComptes('44571') ) then*)
  { FQ 12440 - CA - 02/07/2003 - Contrôle des racines des comptes généraux }
  if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL LIKE "462%"') then
    if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL LIKE "775%"') then
      if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL LIKE "44571%"') then
        bGestionTVA := True;
  IL_TVAAREVERSER.Visible :=bGestionTVA and (fRegroupement='');
  lIL_TVAAREVERSER.Visible:=bGestionTVA and (fRegroupement='');
  if fRegroupement<>'' then
  begin
    QImmo := OpenSQL ('SELECT SUM(I_MONTANTHT) TOTALMONTANT FROM IMMO WHERE I_GROUPEIMMO="'+fRegroupement+'"',True);
    try
      FVo := QImmo.FindField('TOTALMONTANT').AsFloat;
    finally
      Ferme ( QImmo );
    end;
  end else FVo:=fMontantHT;
  QImmo := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fCode+'"',True);
  if not QImmo.Eof then
  begin
    fPlanAvantCession.Charge (QImmo);
    fPlanAvantCession.Recupere(fCode,QImmo.FindField('I_PLANACTIF').AsString);
  end;
  Ferme (QImmo);

  //mbo 24.10.05 fExcep:=fPlanAvantCession.GetExcepExercice(VHImmo^.Encours.Fin, DateOpe); //dot° exceptionnelle
  fExcep:=fPlanAvantCession.GetExcepExercice(VHImmo^.Encours.Fin, DateOpe, false, false); //dot° exceptionnelle

  fTypeExc := '';
  // Quand la quantité est égale à 1 pas de cession partielle
  if fQuantite = 1 then IL_QTECEDEE.Enabled:=FALSE;
  FQteCedee:=fQuantite;
  FTva:=0 ;
  FMontantCession:=FVo;
  IL_VOCEDEE.Min:=0.00; IL_VOCEDEE.Max:=FVo;
  IL_QTECEDEE.Min:=0.00; IL_QTECEDEE.Max:=FQteCedee;
  IL_QTECEDEE.Value:=FQteCedee ;
  IL_VOCEDEE.Value:=FVo ;
  IL_MONTANTCES.Value:=0;
  IL_TVAAREVERSER.Value:=FTva ;
  IL_MOTIFCES.ItemIndex:=0;
  IL_CALCCESSION.ItemIndex:=0;
  //mbo 26.09.05RecalculDotation;
  //RecalculDetailDotation;
  // BTY 04/06 FQ 17629 Règle de calcul de PMValue à choisir par l'utilisateur
  // Supprimer l'enreg 3 inutilisé 'RAD' de la tablette (remplacé par l'enreg 2 'NOR')
  ReglePValue.Items.Delete(3);
  ReglePValue.Values.Delete(3);
  if fRegroupement<>'' then
     ReglePValue.ItemIndex := ReglePValue.Items.IndexOf(RechDom('AMREGLECESSION', 'RSD', FALSE))
  else
     ReglePValue.ItemIndex := ReglePValue.Items.IndexOf(RechDom('AMREGLECESSION', 'NOR', FALSE)) ;

  // MBO - FQ 18395 - 15.06.2006 GBAmortissement.enabled:=not (fNature='FI') ;
  GBAmortissement.enabled:=(not (fNature='FI')) and (not (fDateDebEco > VHImmo^.Encours.Fin));

  { MBO - FQ 18395 - 15.06.2006
   IL_CALCCESSION.enabled :=(not (fNature='FI')) and (not (fPlanAvantCession.AmortEco.Methode='NAM'));
   PAmortissement.TabVisible := (not (fNature='FI')) and (not (fPlanAvantCession.AmortEco.Methode='NAM') and (fRegroupement=''));
   if ((fNature='FI') or (fPlanAvantCession.AmortEco.Methode='NAM')) then PAGES.ActivePage := PBlocNote;
  }
  IL_CALCCESSION.enabled :=(not (fNature='FI')) and (not (fPlanAvantCession.AmortEco.Methode='NAM'))
                           and (not (fDateDebEco > VHImmo^.Encours.Fin));

  PAmortissement.TabVisible := (not (fNature='FI')) and (not (fPlanAvantCession.AmortEco.Methode='NAM')
                               and (fRegroupement='')) and (not (fDateDebEco > VHImmo^.Encours.Fin));
  if ((fNature='FI') or (fPlanAvantCession.AmortEco.Methode='NAM')
                     or (fDateDebEco > VHImmo^.Encours.Fin)) then PAGES.ActivePage := PBlocNote;
end;

procedure TFSortie.RecalculDotation;
var CoeffQteCedee,fDotationExe : double;
    dtOpe(*, dtFinAmort, dtDebutCalc *): TDateTime;
    Year,Month,Day,DayFin : Word;
    PlanInfo : TPlanInfo;
begin
  if not IsValidDate(DATEOPE.Text) then exit;

  dtOpe:=StrToDate(DATEOPE.Text) ;
  CoeffQteCedee:=1 ;
  if fMontantHT<>0 then
    begin
    if fQuantite=1 then CoeffQteCedee:=(IL_VOCEDEE.Value/fMontantHT)
                   else CoeffQteCedee:=FQteCedee/fQuantite;
    end ;

  //YCP 09/07/01
(*  dtFinAmort:=fPlanAvantCession.AmortEco.DateFinAmort;
  if (dtFinAmort<VHImmo^.Encours.Deb) or (dtFinAmort>VHImmo^.Encours.Fin) then
     dtFinAmort:=VHImmo^.Encours.fin
  else
  begin
    if dtOpe>dtFinAmort then dtOpe:=dtFinAmort;
  end;*)
  //YCP 09/07/01 fin

  DecodeDate(dtOpe,Year,Month,Day);
  DecodeDate(FinDeMois(dtOpe),Year,Month,DayFin);
  if Day = DayFin then Day:=0;

{//EPZ 15/12/00
  dtDebutCalc:=VHImmo^.Encours.Deb;

  if fPlanAvantCession.AmortEco.Methode = 'LIN' then
    Prorata := GetProrataLinAvant(dtDebutCalc,dtFinAmort,fDateAmort,dtFinAmort,dtOpe,false)
  else Prorata := GetProrataDegAvant(dtDebutCalc,dtFinAmort,fDateAmort,dtOpe-Day);
}//EPZ 15/12/00

//EPZ 23/11/00
{
  Prorata := Prorata * CoeffQteCedee; //EPZ 21/01/99
  fVNCEco := fPlanAvantCession.GetVNCAvecMethode(fPlanAvantCession.AmortEco, VHImmo^.Encours.Deb);
  DotPartieCedeeEco.Value := Arrondi ((fDotationExeEco-fExcep)*Prorata,V_PGI.OkDecV);
  // ### Modif CA le 07/07/2000 - Arrondi car si VNC = 0 VNC en fait égale à -1e-13 !!!
  fVNCEco := Arrondi((fVNCEco* CoeffQteCedee) - Arrondi ((fDotationExeEco-fExcep)*Prorata + (fExcep*CoeffQteCedee),V_PGI.OkDecV),V_PGI.OkDecV);
  VNCEco.Value := fVNCEco;
}
  PlanInfo := TPlanInfo.Create(fCode);
  try
    PlanInfo.Plan.copie(fPlanAvantCession);
    PlanInfo.Calcul(dtOpe,True, true, AmortJsortie);   //fq 21754 ajout des 2 derniers paramètres

    if fPlanAvantCession.AmortEco.Methode='NAM' then
    begin
      fDotationExe:=0.0 ;
      fVNC:=PlanInfo.VNCEco ;
    end
    else
    begin
      // YCP 09/07/01 prend en compte l'amort fiscal s'il existe
(*      if fPlanAvantCession.AmortFisc.Methode<>'' then
      begin
        fDotationExe:=PlanInfo.DotationFisc*CoeffQteCedee;
        DotPartieCedee.Value := fDotationExe;
        fVNC:=fPlanAvantCession.GetVNCAvecMethode(fPlanAvantCession.AmortFisc,VHImmo^.Encours.Deb) ;
      end
      else
      begin*)
        fDotationExe:=PlanInfo.DotationEco*CoeffQteCedee; // dot éco exercice proratisée + dot except°

        // ajout mbo 15280
        DotEcoExe :=fDotationExe;
        DotFiscExe:= PlanInfo.DotationFisc*CoeffQteCedee; // dot exercice fiscale proratisée

        DotPartieCedee.Value := fDotationExe;
        fVNC:=fPlanAvantCession.GetVNCAvecMethode(fPlanAvantCession.AmortEco, VHImmo^.Encours.Deb)+fTVA;
        initVNCeco := fVNC*CoeffQteCedee ;   // fVNC = vnc debut exercice
        InitVNCFisc:=fPlanAvantCession.GetVNCAvecMethode(fPlanAvantCession.AmortFisc,VHImmo^.Encours.Deb) ;
        InitVNCFisc := InitVNCFisc * CoeffQteCedee ;
///      end ;
      //YCP 09/07/01
    end ;
    fVNC:=Arrondi((fVNC*CoeffQteCedee)-fDotationExe,V_PGI.OkDecV); // vnc fin d'exercice en cours
    // VNC.Value:=fVNC;
    if not ISvalidDate(DateOpe.Text) then VNC.Value := 0
//    else VNC.Value := fVNC + (fMontantHT + fTVARecuperable  - fTVARecuperee - fBaseEco );  // FQ 12083
    else VNC.Value := fVNC; // FQ 14664 : On affiche la VNC
  finally
    PlanInfo.Free;
  end ;

  if IL_CALCCESSION.Value='NOR' then
  begin
    // Dotation.Value:=VNC.Value ;
    // fq 17093 mbo 29/11/05 - Dotation.Value:=fVNC ; // FQ 12083
    if rbAjouter.checked then
       Dotation.Value := MaxAjout

    else
       Dotation.value := MaxRemplace;

    DotTest := Dotation.value;
  end
  else if IL_CALCCESSION.Value='SAN' then
  begin
    DotPartieCedee.Value := 0.00;
    // YCP 09/07/01 prend en compte l'amort fiscal s'il existe
    if fPlanAvantCession.AmortFisc.Methode<>'' then
      fVNC:=fPlanAvantCession.GetVNCAvecMethode(fPlanAvantCession.AmortFisc,VHImmo^.Encours.Deb)*CoeffQteCedee
    else
      fVNC:=fPlanAvantCession.GetVNCAvecMethode(fPlanAvantCession.AmortEco, VHImmo^.Encours.Deb)*CoeffQteCedee ;
    //YCP 09/07/01
  end;
  GBAmortissement.Enabled:=(IL_CALCCESSION.Value<>'NOR') ;
end;

procedure TFSortie.RecalculDetailDotation;
var //DotationExe: double; BTY 04/06 Pour diminuer les messages compilateur
    RepriseEco : double;
    RepriseFisc : double;
    TotalEco: double;
    TotalFisc : double;
    DotInitCalc:double;
    //Typeexc:string; BTY 04/06 Pour diminuer les messages compilateur

begin
  //ajout mbo 30.09.05 - fq 15280
  RepriseEco := fPlanAvantCession.AmortEco.Reprise;   // antérieurs saisis (n'inclut pas cumul antérieurs calculés)
  RepriseFisc := fPlanAvantCession.AmortFisc.Reprise;

  If (fPlanAvantCession.AmortFisc.Methode<>'') AND (fPlanAvantCession.AmortEco.Methode<>'VAR') then derog:=true;
  //dotationExe = dotation eco exercice non proratisée + exceptionnel
  //DotationExe := fPlanAvantCession.GetDotationExercice(VHImmo^.EnCours.Fin,fPlanAvantCession.AmortEco,false);
  //DotFiscExe := fPlanAvantCession.GetDotationExercice(VHImmo^.EnCours.Fin,fPlanAvantCession.AmortFisc,false);

  fPlanAvantCession.GetCumulsDotExercice(VHImmo^.Encours.Deb,TotalEco, TotalFisc,false,false,true);


  TotalEco := TotalEco + RepriseEco + DotEcoExe;
  DotInitCalc := Arrondi ((DotEcoExe - fExcep), V_PGI.OkDecV);
  TotalFisc := TotalFisc + RepriseFisc + DotFiscExe;
  // TotalEco = cumul antérieurs eco + reprise  + dot eco exo en cours proratisée
  // Totalfisc = cumul antérieurs fisc +  reprise + dot fiscale exo en cours proratisée

  // ds DotEcoExe sont comprises : dotation exercice proratisée
  //                             + exceptionnel saisi (non proratisé)

  //FPlanAvantCession.GetElementExceptionnel(DotEcoExe, MontantExc, TypeExc);
  //if typeexc = 'RDO' then MontantExc := MontantExc * (-1);

  if derog = false then
    begin
      MaxAjout := Arrondi((InitVNCeco-DotEcoExe), V_PGI.OkDecV);
      MaxRemplace := Arrondi ( (initVNCeco-fExcep), V_PGI.OkDecV ); // prendre en cpte exceptionnel saisi
    end
  else
    begin
      MaxAjout := Arrondi ( (InitVNCFisc-DotFiscExe), V_PGI.OkDecV );
      if maxAjout > Arrondi (( InitVNCeco-DotEcoExe), V_PGI.OkDecV ) then MaxAjout :=Arrondi ( (InitVNCeco-DotEcoExe), V_PGI.OkDecV );
      MaxRemplace := Arrondi (( InitVNCFisc-fExcep), V_PGI.OkDecV );
      If MaxRemplace > Arrondi ((InitVNCeco-fExcep), V_PGI.OkDecV)then MaxRemplace :=Arrondi ((InitVNCeco-fExcep), V_PGI.OkDecV);
      If MaxRemplace > Arrondi ((InitVNCFisc-fExcep), V_PGI.OkDecV)then MaxRemplace :=Arrondi ((InitVNCfisc-fExcep), V_PGI.OkDecV);
  end;

  //ajout mbo 20.09.2005
  if derog = false then
  begin
    if rbAjouter.Checked then
      begin
        DotationCalc.Value := DotPartieCedee.Value+Dotation.Value;
        Excep.Value := Arrondi((fExcep + Dotation.Value), V_PGI.OkDecV);
        VNC.Value := Arrondi ((InitVNCeco - DotEcoExe - Dotation.Value), V_PGI.OkDecV);
      end
    else
      begin
        DotationCalc.Value := Dotation.Value;
        Excep.Value := Arrondi ((Dotation.Value-DotInitCalc),V_PGI.OkDecV);
        VNC.Value := Arrondi((InitVNCeco - Dotation.Value),V_PGI.OkDecV);
      end;

    if not ISvalidDate(DateOpe.Text) then VNC.Value := 0;

    if (VNC.Value) < 0 then
    begin
      HM2.Execute(1,Caption,'');
      // ajout mbo 28/11/05 fq 17093
      Dotation.value := 0;
      // ajout mbo 7.12.05 fq 17155
      DotTest := 0;

      if rbAjouter.Checked then
        begin
        VNC.Value := Arrondi ((InitVNCeco - DotEcoExe), V_PGI.OkDecV);
        DotationCalc.Value := DotPartieCedee.Value;
        Excep.Value := fExcep;
        end
      else
        begin
        VNC.value := InitVNCeco;
        DotationCalc.Value := 0;
        Excep.Value := Arrondi ((Dotation.Value-DotInitCalc),V_PGI.OkDecV);
        end;
      // fin ajout mbo 17093
    FocusControl(Dotation);
    end
  end else
    if rbAjouter.Checked then
      begin
       DotationCalc.Value := DotPartieCedee.Value+Dotation.Value;
       Excep.Value := Arrondi((fExcep + Dotation.Value), V_PGI.OkDecV);
       VNC.Value := Arrondi ((InitVNCeco - DotEcoExe - Dotation.Value), V_PGI.OkDecV);
      end
    else
      begin
        DotationCalc.Value := Dotation.Value;
        Excep.Value := Arrondi ((Dotation.Value-DotInitCalc),V_PGI.OkDecV);
        VNC.Value := Arrondi((InitVNCeco - Dotation.Value),V_PGI.OkDecV);
      end;

    // fq 17155 7.12.05 mbo  if (rbAjouter.Checked) AND (Dotation.Value > MaxAjout) then
    if (rbAjouter.Checked) AND (DotTest > MaxAjout) then
    begin
        HM2.execute(4,Caption,FloatToStr(MaxAjout));
        // ajout mbo 29.11.05
        Dotation.Value := 0;
        DotTest := 0;  // fq 17155 ajout mbo 7.12.05
        DotationCalc.Value := DotPartieCedee.Value+Dotation.Value;
        Excep.Value := Arrondi((fExcep + Dotation.Value), V_PGI.OkDecV);
        VNC.Value := Arrondi ((InitVNCeco - DotEcoExe - Dotation.Value), V_PGI.OkDecV);
        FocusControl(Dotation);
    end
    else
      // fq 17155 - mbo 7.12.05 if (rbRemplacer.Checked) AND (Dotation.Value>MaxRemplace) then
      if (rbRemplacer.Checked) AND (DotTest>MaxRemplace) then
        begin
        HM2.execute(4,Caption,FloatToStr(MaxRemplace));
        // ajout mbo 29.11.2005
        Dotation.Value := 0;
        DotTest := 0; // fq 17155 - ajout mbo 07.12.05
        DotationCalc.Value := 0;
        Excep.Value := Arrondi((fExcep + Dotation.Value), V_PGI.OkDecV);
        VNC.Value := Arrondi ((InitVNCeco - DotEcoExe - Dotation.Value), V_PGI.OkDecV);
        FocusControl(Dotation);
        end;
end;

function TFSortie.ControleZones : boolean;
var cumul : double;
    i : integer;
begin
  result:=false ;
  if (inherited ControleZones) then
    begin
    if (VNC.Value) < 0 then
      begin
      HM2.Execute(1,Caption,'');
      FocusControl (Dotation);
      end
    else if IL_MONTANTCES.Value<0 then
      begin
      HM2.execute(0,Caption,'') ;
      IL_MONTANTCES.SetFocus ;
      end
    else if ((fRegroupement='') and (IL_VOCEDEE.Value>fMontantHT)) then
      begin
      HM2.execute(3,Caption,'') ;
      IL_VOCEDEE.SetFocus ;
      IL_VOCEDEE.Value:=fMontantHT ;
      end
    // mbo - fq 15280 - autoriser exceptionnel si immo avec derog sauf plan variabl
    //else  if (Dotation.Value>0) and (existeSQL('SELECT I_METHODEFISC FROM IMMO WHERE I_IMMO="'+fCode+'" AND I_METHODEFISC<>"" ')) then
      { mbo 16804  autoriser exceptionnel sur plan variable
      else  if (Dotation.Value>0) and (existeSQL('SELECT I_METHODEFISC FROM IMMO WHERE I_IMMO="'+fCode+'" AND I_METHODEFISC<>"" AND I_METHODEECO ="VAR"')) then
      begin
      // mbo 15280 PGIBox('Vous ne pouvez pas effectuer de dotations ou reprises exceptionnelles sur une immobilisation gérant l''amortissement dérogatoire','Changement de plan') ;

      PGIBox('Dotations ou reprises exceptionnelles non autorisées sur une immobilisation ayant un plan économique variable','Changement de plan') ;
      if GBAmortissement.Enabled then FocusControl(Dotation)
                                 else FocusControl(IL_CALCCESSION) ;
      end }

    // FQ 18916 En répartition libre sur un regroupement, comparer le total réparti avec le prix de sortie
    else if (fRegroupement <>'') then
    begin
      if MODEREPARTITION.Value ='LIB' then
      begin
        Cumul := 0;
        for i:=0 to fTRepartition.Detail.Count - 1 do
           Cumul := Arrondi (Cumul + Valeur(LISTEREPARTITION.Cells[4,i+1]), V_PGI.OkDecV);
        if (IL_MONTANTCES.Value - Cumul) <> 0 then
        begin
           HM2.execute(5,Caption,'') ;
           IL_MONTANTCES.SetFocus ;
        end else
           result := true;
      end else
        result := true;
    end
    // -------
    else result := true;
    end ;
end;

procedure ExecuteChangementPlan (var PlanCede : TPlanAmort; var EnregCession : TEnregCession);
var ChangePlan : TChangePlanAmort;
begin
  ChangePlan := TChangePlanAmort.Create (Application);
  ChangePlan.Plan.copie(PlanCede);
  ChangePlan.AncienPlan.copie(PlanCede);
  ChangePlan.CodeImmo:=PlanCede.CodeImmo;
  ChangePlan.DateOpe :=EnregCession.DateOpe;
  ChangePlan.TypeChangement:=2;
  if EnregCession.MontantExc>0  then
  begin
    EnregCession.TypeExc := 'DOT';
    ChangePlan.Plan.MontantExc:=ChangePlan.Plan.MontantExc+EnregCession.MontantExc;
  end  else if EnregCession.MontantExc < 0 then
  begin
    ChangePlan.Plan.MontantExc:=ChangePlan.Plan.MontantExc-EnregCession.MontantExc;
    EnregCession.TypeExc := 'RDO';
  end;
  ChangePlan.Plan.TypeOpe := 'ELC';
  ChangePlan.Plan.SetTypeExc(EnregCession.TypeExc);
  ChangePlan.TypeExc := EnregCession.TypeExc;
  ChangePlan.TypeDot := EnregCession.TypeExc;
  ChangePlan.MontantDot := EnregCession.Dotation;
  ChangePlan.Plan.CalculExceptionnelSurSortie(ChangePlan.Plan.AmortEco,EnregCession.DateOpe);
  PlanCede.copie(ChangePlan.Plan);
  ChangePlan.free;
end;

procedure TFSortie.EnregCession ;

{var CessionPartielle : boolean ;  // fq 21746
    PlanImmoCedee : TPlanAmort ;
    EnregCession : TEnregCession;
    i : integer;
}
begin
    ExecuteSortie;


 {fq 21746 - mbo - 26.10.07 plus de coche spécifique assistance
  if not BSORTIECALCULEE.Checked then
  begin
    ExecuteSortie;
    exit;
  end;
  if fRegroupement <> '' then
  // Sortie complète du regroupement
  begin
    // Calcul du numéro d'ordre pour opération en série
    fOrdreS := TrouveNumeroOrdreSerieLogSuivant;
    // Récupération des informations de la grille de saisie de la répartition
    fTRepartition.GetGridDetail(ListeRepartition,ListeRepartition.RowCount-1,'','I_IMMO;I_LIBELLE;I_MONTANTHT;I_VNC;PRIXSORTIE',True);
    // Enregistrement des sortie pour chaque immobilisation concernée
    for i:=0 to fTRepartition.Detail.Count - 1 do
      EnregistreLaSortieRegroupement(fTRepartition.Detail[i].GetValue('I_IMMO'),fTRepartition.Detail[i].GetValue('PRIXSORTIE'),fTRepartition.Detail[i].GetValue('I_VNC'));
  end else
  // Sortie simple d'une immobilisation
  begin
    fOrdreS := TrouveNumeroOrdreSerieLogSuivant;
    CessionPartielle:=(FQteCedee<fQuantite) or (FVo<>fMontantHT) ; // Cession partielle ?
    if (fQuantite=1) and (CessionPartielle) then FQteCedee := 0;
    EnregCession                :=TEnregCession.Create;
    EnregCession.Code           :=FCode;
    EnregCession.Vo             :=FVo;
    EnregCession.QteAchat       :=fQuantite;
    EnregCession.QteCedee       :=FQteCedee;
    EnregCession.DateOpe        :=StrToDate(DATEOPE.Text);
    EnregCession.Tva            :=FTva;
    EnregCession.MontantHT      :=fMttHT  ;
    EnregCession.MontantAchat   :=fMontantHT;
    EnregCession.MontantCession :=IL_MONTANTCES.Value;
    EnregCession.MotifCession   :=IL_MOTIFCES.Value ;
    EnregCession.CalcCession    :=IL_CALCCESSION.Text ;
    EnregCession.TypeCession    :=IL_CALCCESSION.Value;
    EnregCession.BlocNote       :=IL_BLOCNOTE.LinesRTF;
    EnregCession.Tva            :=FTva;

    // BTY 04/06 FQ 17629
    EnregCession.RegleCession   :=ReglePValue.Value ;

    // modif mbo fq 15280 car excep = element exceptionnel + excep cession
    // EnregCession.ChangePlan     :=Excep.Value <> 0;
    EnregCession.ChangePlan     :=Arrondi((Excep.Value - fExcep),V_PGI.OkDecV) <> 0;

    // mbo fq 15280 - EnregCession.MontantExc     :=Excep.Value;
    // mbo fq 15280 - EnregCession.TypeExc        :=fTypeExc;
    if Arrondi((Excep.Value - fExcep),V_PGI.OkDecV) < 0 then
       begin
       EnregCession.MontantExc  :=Arrondi(((Excep.Value - fExcep)*(-1)),V_PGI.OkDecV);
       EnregCession.TypeExc     := 'RDO';
       end
    else
       begin
       EnregCession.MontantExc  :=Arrondi((Excep.Value - fExcep),V_PGI.OkDecV);
       EnregCession.TypeExc     := 'DOT';
    end;

    EnregCession.Dotation       :=Valeur(Dotation.Text);
    EnregCession.BaseEco        :=fBaseEco;
    EnregCession.BaseFisc       :=fBaseFisc;
    EnregCession.BaseTP         :=fBaseTP;
    EnregCession.VNCCedee       :=fVnc;
    PlanImmoCedee:=TPlanAmort.Create(true) ;
    TraiteImmoCedee(PlanImmoCedee,EnregCession);
    fTypeExc := EnregCession.TypeExc;
    TraiteImmoOrig(PlanImmoCedee,CessionPartielle,EnregCession);
  //  if fPlanAvantCession.AmortFisc.Methode<>'' then
  //        EnregCession.PValue:=CalcVCession + EnregCession.AmortFisc + EnregCession.MontantExc
  //  else  EnregCession.PValue:=CalcVCession + EnregCession.AmortEco  + EnregCession.MontantExc ;
    { CA - 02/06/2003 - Calcul des + ou - values : prise en compte de l'exceptionnel }
  {  if fPlanAvantCession.AmortFisc.Methode<>'' then
          EnregCession.PValue:=CalcVCession + EnregCession.AmortFisc + EnregCession.MontantExc
    else  EnregCession.PValue:=IL_MONTANTCES.Value - (EnregCession.VNCCedee - EnregCession.MontantExc) ;}
    { OL& CJ le 19/06/2003 - +/- value économique UNIQUEMENT
//    EnregCession.PValue:=IL_MONTANTCES.Value - (EnregCession.VNCCedee - EnregCession.MontantExc);
    EnregCession.PValue:=IL_MONTANTCES.Value - (EnregCession.VNCCedee - EnregCession.MontantExc);  // FQ 12083
    EnregLogCession(PlanImmoCedee,CessionPartielle,EnregCession,fOrdreS);
    PlanImmoCedee.free ;
    EnregCession.free ;
  end;

  fin fq 21746 }
end;

procedure TFSortie.DATEOPEExit(Sender: TObject);
begin
  inherited;
  if (StrToDate(DATEOPE.Text) <VHImmo^.Encours.Deb) or
     (StrToDate(DATEOPE.Text)>VHImmo^.Encours.Fin)
     or (StrToDate(DATEOPE.Text)< fDateAchat) then exit;
  RecalculDotation;
  RecalculDetailDotation;
  CalculRepartition ( True );
end;

procedure TFSortie.DATEOPEEnter(Sender: TObject);
begin
  inherited;
  fCurDateOp := DATEOPE.Text;
end;

procedure TFSortie.OnModifDotation(Sender: TObject);
begin
  inherited;
  // fq 17155 - mbo 7.12.05 if (Dotation.Value<0) then
  if (DotTest<0) then
    begin
    HM.Execute(8,caption,'') ;
    FocusControl(Dotation) ;
    end
  else
    { mbo - fq 15280 - autoriser saisie exceptionnel même sur immo avec dérog.
    if (Dotation.Value>0) and (existeSQL('SELECT I_METHODEFISC FROM IMMO WHERE I_IMMO="'+fCode+'" AND I_METHODEFISC<>""')) then
    begin
    PGIBox('Vous ne pouvez pas effectuer de dotations ou reprises exceptionnelles sur une immobilisation gérant l''amortissement dérogatoire','Changement de plan') ;
    FocusControl(Dotation) ;
    end
  else}
    RecalculDotation;
    RecalculDetailDotation;
end;

procedure TFSortie.IL_CALCCESSIONExit(Sender: TObject);
begin
  inherited;
  RecalculDotation;
  RecalculDetailDotation;
end;

procedure TFSortie.AccelCreationMotifClick(Sender: TObject);
var OldVal : string;
begin
  inherited;
  OldVal := IL_MOTIFCES.Value;
  ParamTable('TIMOTIFCESSION',taCreat,0,nil) ;
  IL_MOTIFCES.Reload;
  IL_MOTIFCES.Value := OldVal;
end;

procedure TraiteImmoCedee(var PlanImmoCedee : TPlanAmort; var EnregCession : TEnregCession);
var QOrig : TQuery;
    {QteOrig,}i : integer;
    TTvaR,TTvaE,wAmortEco,wAmortFisc, wCoef : double;
    TCedee, TPlanUO : TOB;
    wAmortPri, wAmortSbv : double;
begin
  TTvaR := 0.0; TTvaE := 0.0 ; wCoef:=0.0 ;
  QOrig:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+EnregCession.Code+'"', TRUE);
  if not QOrig.EOF then
  begin
    if EnregCession.MontantAchat<>0 then wCoef:=EnregCession.Vo/EnregCession.MontantAchat ;
    TCedee := TOB.Create ('IMMO', nil,-1);
    for i:=0 to QOrig.FieldCount-1 do TCedee.PutValue(QOrig.Fields[i].FieldName,QOrig.Fields[i].AsVariant) ;
    // BTY 01/06 FQ 17259 Nouveau top dépréciation
    // BTY 04/06 FQ 17516 Nouveau top changement de regroupement
    //InitOpeEnCoursImmo(TCedee,'-','-','-','-','-','-','-','-','-');
    //InitOpeEnCoursImmo(TCedee,'-','-','-','-','-','-','-','-','-','-');
    InitOpeEnCoursImmo(TCedee,'-','-','-','-','-','-','-','-','-','-','-');
    TCedee.PutValue('I_IMMO','TMP~20011999');
    TCedee.PutValue('I_MONTANTHT',EnregCession.Vo) ;
    TCedee.PutValue('I_QUANTITE',EnregCession.QteCedee);
    TCedee.PutValue('I_QTCEDE',EnregCession.QteCedee );
    TCedee.PutValue('I_DATECESSION',EnregCession.DateOpe);
    TCedee.PutValue('I_IMMOORIGINEECL',EnregCession.Code) ;
    TTvaR:=TCedee.GetValue('I_TVARECUPERABLE')*wCoef ;
    TTvaE:=TCedee.GetValue('I_TVARECUPEREE')*wCoef ;
    TCedee.PutValue('I_BASETAXEPRO',EnregCession.BaseTP*wCoef );
    TCedee.PutValue('I_BASEECO',EnregCession.BaseEco*wCoef) ;
    TCedee.PutValue('I_BASEFISC',EnregCession.BaseFisc*wCoef) ;
    TCedee.PutValue('I_TVARECUPERABLE',TTvaR) ;
    TCedee.PutValue('I_TVARECUPEREE',TTvaE);
    TCedee.PutValue('I_REPCEDECO',TCedee.GetValue('I_REPCEDECO')+TCedee.GetValue('I_REPRISEECO')*(1-wCoef));
    TCedee.PutValue('I_REPRISEECO',TCedee.GetValue('I_REPRISEECO')*wCoef);
    TCedee.PutValue('I_REPCEDFISC',TCedee.GetValue('I_REPCEDFISC')+TCedee.GetValue('I_REPRISEFISCAL')*(1-wCoef));
    TCedee.PutValue('I_REPRISEFISCAL',TCedee.GetValue('I_REPRISEFISCAL')*wCoef);
    TCedee.PutValue('I_ETAT','OUV');

    // TGA Ajout zones sur immo 23/01/2006
    TCedee.PutValue('I_REPRISEDEP',TCedee.GetValue('I_REPRISEDEP')*wCoef);
    TCedee.PutValue('I_ECCLEECR','');
    TCedee.PutValue('I_DATEFINCB',iDate1900);

    // Si changement de plan avec sortie , alors exceptionnel sur sortie
    if EnregCession.ChangePlan then
      TCedee.PutValue('I_MONTANTEXC',TCedee.GetValue('I_MONTANTEXC'))
    else
      TCedee.PutValue('I_MONTANTEXC',TCedee.GetValue('I_MONTANTEXC')*wCoef);

    // ajout pour la prime
    TCedee.PutValue('I_REPRISEUO',TCedee.GetValue('I_REPRISEUO')*wCoef);
    TCedee.PutValue('I_SBVPRI',TCedee.GetValue('I_SBVPRI')*wCoef);
    TCedee.PutValue('I_REPRISEUOCEDEE',TCedee.GetValue('I_REPRISEUOCEDEE')+TCedee.GetValue('I_REPRISEUO')*(1-wCoef));

    // pour la subvention on ne fait rien pour les antérieurs
    // car on n'a pas de zone pour stocker les antérieurs cédés
    TCedee.PutValue('I_SBVMT',TCedee.GetValue('I_SBVMT')*wCoef);

    TCedee.InsertDB(nil);
    { Enregistrement du plan d'unités d'oeuvre le cas échéant }
    if (TCedee.GetValue('I_METHODEECO')='VAR') then
    begin
      TPlanUO := TOB.Create ('Maman',nil,-1);
      try
        TPlanUO.LoadDetailDBFromSQL ('IMMOUO','SELECT * FROM IMMOUO WHERE IUO_IMMO="'+EnregCession.Code+'"');
        for i:=0 to TPlanUO.Detail.Count -1 do
          TPlanUO.Detail[i].PutValue('IUO_IMMO','TMP~20011999');
        TPlanUO.InsertDB(nil);
      finally
        TPlanUO.Free;
      end;
    end;
    PlanImmoCedee.ChargeTOB(TCedee);
    PlanImmoCedee.AmortEco.Reprise:=(TCedee.GetValue('I_REPRISEECO'));
    PlanImmoCedee.AmortFisc.Reprise:=(TCedee.GetValue('I_REPRISEFISCAL'));
    PlanImmoCedee.AmortPRI.Reprise:=(TCedee.GetValue('I_REPRISEUO'));
    PlanImmoCedee.AmortSBV.Reprise:=(TCedee.GetValue('I_CORRECTIONVR'));

    PlanImmoCedee.SetInfosCession(EnregCession.TypeCession,EnregCession.DateOpe);
    PlanImmoCedee.CalculPlanCede(TCedee, TCedee.GetValue('I_DATEAMORT')) ;
    PlanImmoCedee.CalculProrataDotationCedee(PlanImmoCedee.AmortEco,EnregCession.DateOpe);
    PlanImmoCedee.CalculProrataDotationCedee(PlanImmoCedee.AmortFisc,EnregCession.DateOpe);
    PlanImmoCedee.CalculProrataDotationCedee(PlanImmoCedee.AmortPri,EnregCession.DateOpe);
    PlanImmoCedee.CalculProrataDotationCedee(PlanImmoCedee.AmortSbv,EnregCession.DateOpe);

    //EPZ 20/12/2000    EnregCession.VNCCedee := PlanImmoCedee.GetVNCAvecMethode(PlanImmoCedee.AmortEco,EnregCession.DateOpe);
    PlanImmoCedee.ImmoCloture := QOrig.FindField('I_ETAT').AsString<>'OUV' ; //restaure
    PlanImmoCedee.GetCumulsDotExercice(VHImmo^.Encours.Fin, wAmortEco, wAmortFisc,false,true,true);

    PlanImmoCedee.GetCumulPri(VHImmo^.Encours.Fin, wAmortPri, false,true,true);
    PlanImmoCedee.GetCumulSbv(VHImmo^.Encours.Fin, wAmortSbv, false,true,true);

    PlanImmoCedee.SetBaseAmortFin(EnregCession.Vo);
    PlanImmoCedee.TypeOpe:='CES';
    TCedee.Free;
  end ;
  Ferme(QOrig);
  if EnregCession.ChangePlan then ExecuteChangementPlan(PlanImmoCedee,EnregCession);
  ExecuteSQL('DELETE FROM IMMO WHERE I_IMMO="TMP~20011999"');
  ExecuteSQL('DELETE FROM IMMOUO WHERE IUO_IMMO="TMP~20011999"');
  EnregCession.TvaR := TTvaR;
  EnregCession.TvaE := TTvaE;
  EnregCession.AmortEco := wAmortEco;
  EnregCession.AmortFisc := wAmortFisc;
  EnregCession.RepriseEco := PlanImmoCedee.AmortEco.Reprise;
  EnregCession.RepriseFisc := PlanImmoCedee.AmortFisc.Reprise;
  // ajout pour prime et subvention
  EnregCession.ReprisePri := PlanImmoCedee.AmortPri.Reprise;
  EnregCession.RepriseSBV := PlanImmoCedee.AmortSbv.Reprise;

end;

// cette procedure n'est plus utilisée
procedure EnregLogCession(PlanImmoCedee : TPlanAmort;CessionPartielle : boolean; EnregCession : TEnregCession; OrdreS : integer);
var TLog : TOB;
    CodeCession : string;
    Ordre : integer;
    CumulAntEco, CumulAntFisc, dDotEco,dExcEco,dCesEco : double;
    CommentaireOpe : string;
begin
  Ordre := TrouveNumeroOrdreLogSuivant(EnregCession.Code);
  TLog := TOB.Create ('IMMOLOG',nil,-1);
  try
    TLog.PutValue('IL_IMMO',EnregCession.Code);
    if CessionPartielle then CodeCession:='CEP' else CodeCession:='CES' ;
    // BTY 04/06 FQ 17629 Commentaire de l'opération en fonction du calcul de la PMValue
    //TLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', CodeCession, FALSE)+' '+DateToStr(EnregCession.DateOpe));
    if (EnregCession.RegleCession='CT') then
        CommentaireOpe := ' (répartition +/- value à CT)'
    else if (EnregCession.RegleCession='LT') then
        CommentaireOpe := ' (répartition +/- value à LT)'
    else if (EnregCession.RegleCession='RSD') then
        CommentaireOpe := ' (répart. +/- ignorant la durée)'
    else
        CommentaireOpe := '';
    TLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', CodeCession, FALSE)+' '+DateToStr(EnregCession.DateOpe)+CommentaireOpe);
    TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('CES'));
    TLog.PutValue('IL_DATEOP',EnregCession.DateOpe);
    TLog.PutValue('IL_TYPEOP',CodeCession);
    TLog.PutValue('IL_MOTIFCES',EnregCession.MotifCession);
    TLog.PutValue('IL_CALCCESSION',EnregCession.CalcCession);
    TLog.PutValue('IL_VOCEDEE',EnregCession.Vo);
    TLog.PutValue('IL_VNC',EnregCession.VNCCedee);
    if EnregCession.QteAchat <> 1 then TLog.PutValue('IL_QTECEDEE',EnregCession.QteCedee);
    TLog.PutValue('IL_MONTANTCES',EnregCession.MontantCession);
    TLog.PutValue('IL_TVAAREVERSER',EnregCession.Tva);
    TLog.PutValue('IL_TVARECUPEREE',EnregCession.TvaE);
    TLog.PutValue('IL_TVARECUPERABLE',EnregCession.TvaR);
    TLog.PutValue('IL_REPRISEECO',EnregCession.RepriseEco);
    TLog.PutValue('IL_REPRISEFISC',EnregCession.RepriseFisc);
    TLog.PutValue('IL_PVALUE',EnregCession.PValue);
    TLog.PutValue('IL_BLOCNOTE',EnregCession.BlocNote.Text);
    TLog.PutValue('IL_PLANACTIFAV',EnregCession.PLanPrec);
    TLog.PutValue('IL_PLANACTIFAP',EnregCession.PlanSuiv);
    TLog.PutValue('IL_ORDRE',Ordre);
    TLog.PutValue('IL_ORDRESERIE',OrdreS);
    PlanImmoCedee.GetCumulsDotExercice(VHImmo^.Encours.Deb,CumulAntEco,CumulAntFisc, True, true, true);
    TLog.PutValue('IL_DOTCESSECO',Arrondi(PlanImmoCedee.GetDotationGlobale(1,VHImmo^.Encours.Fin,dDotEco,dCesEco,dExcEco),V_PGI.OKDecV));

    // ajout mbo 27.09.2005 si pas de plan fiscal on ne met pas à jour
    if PlanImmoCedee.AmortFisc.Methode<>'' then
       TLog.PutValue('IL_DOTCESSFIS',Arrondi(PlanImmoCedee.GetDotationGlobale(2,VHImmo^.Encours.Fin,dDotEco,dCesEco,dExcEco),V_PGI.OKDecV));

    TLog.PutValue('IL_CUMANTCESECO',Arrondi(CumulAntEco,V_PGI.OKDecV));
    TLog.PutValue('IL_CUMANTCESFIS',Arrondi(CumulAntFisc,V_PGI.OKDecV));
    TLog.PutValue('IL_MONTANTEXC',PlanImmoCedee.MontantExc);
    TLog.PutValue('IL_TYPEEXC',EnregCession.TypeExc);
    TLog.PutValue('IL_BASEECOAVMB',EnregCession.BaseEco);
    TLog.PutValue('IL_BASEFISCAVMB',EnregCession.BaseFisc);
    TLog.PutValue('IL_BASETAXEPRO',EnregCession.BaseTP);
    TLog.InsertDB(nil);
  finally
    TLog.Free;
  end;
end;

procedure TraiteImmoOrig(PlanImmoCedee : TPlanAmort; CessionPartielle : boolean; var EnregCession : TEnregCession);
var
   QOrigine : TQuery;
   wVNC,TTvaRO,TTvaEO,wCoef : double ;
   PlanImmoOrig : TPlanAmort ;
begin
  wCoef:=0.0 ;
  if EnregCession.MontantAchat<>0 then wCoef:=EnregCession.Vo/EnregCession.MontantAchat ;
  QOrigine :=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+EnregCession.Code+'"', FALSE);
  QOrigine.Edit ;
  QOrigine.FindField('I_MONTANTHT').AsFloat     :=EnregCession.MontantHT*(1-wCoef);
  QOrigine.FindField('I_BASEAMORFINEXO').AsFloat:=EnregCession.MontantAchat-EnregCession.Vo;
  TTvaRO:=QOrigine.FindField('I_TVARECUPERABLE').AsFloat-EnregCession.TvaR ;
  TTvaEO:=QOrigine.FindField('I_TVARECUPEREE').AsFloat-EnregCession.TvaE ;
  QOrigine.FindField('I_TVARECUPERABLE').AsFloat:=TTvaRO ;
  QOrigine.FindField('I_TVARECUPEREE').AsFloat  :=TTvaEO ;
  if (EnregCession.QteAchat=1) and (CessionPartielle) then QOrigine.FindField('I_QUANTITE').AsFloat:=EnregCession.QteAchat
  else QOrigine.FindField('I_QUANTITE').AsFloat :=EnregCession.QteAchat-EnregCession.QteCedee;
  QOrigine.FindField('I_BASETAXEPRO').AsFloat   :=EnregCession.BaseTP*(1-wCoef);
  QOrigine.FindField('I_BASEECO').AsFloat       :=EnregCession.BaseEco*(1-wCoef);
  QOrigine.FindField('I_BASEFISC').AsFloat      :=EnregCession.BaseFisc*(1-wCoef);
  if EnregCession.QteAchat = 1 then QOrigine.FindField('I_QTCEDE').AsFloat:=1
  else QOrigine.FindField('I_QTCEDE').AsFloat   :=QOrigine.FindField('I_QTCEDE').AsFloat+EnregCession.QteCedee;
  QOrigine.FindField('I_REPCEDECO').AsFloat     :=QOrigine.FindField('I_REPCEDECO').AsFloat+QOrigine.FindField('I_REPRISEECO').AsFloat*wCoef;
  QOrigine.FindField('I_REPRISEECO').AsFloat    :=QOrigine.FindField('I_REPRISEECO').AsFloat*(1-wCoef);
  QOrigine.FindField('I_REPCEDFISC').AsFloat    :=QOrigine.FindField('I_REPCEDFISC').AsFloat+QOrigine.FindField('I_REPRISEFISCAL').AsFloat*wCoef;
  QOrigine.FindField('I_REPRISEFISCAL').AsFloat :=QOrigine.FindField('I_REPRISEFISCAL').AsFloat*(1-wCoef);
  QOrigine.FindField('I_MONTANTEXCCED').AsFloat :=QOrigine.FindField('I_MONTANTEXC').AsFloat*wCoef;

  // BTY 04/06 FQ 17629
  QOrigine.FindField('I_REGLECESSION').AsString       :=EnregCession.RegleCession;

  if not EnregCession.ChangePlan then // Uniquement si exceptionnel pas joint avec sortie
    QOrigine.FindField('I_MONTANTEXC').AsFloat:=QOrigine.FindField('I_MONTANTEXC').AsFloat*(1-wCoef);

  PlanImmoOrig:=TPlanAmort.Create(true) ;
  PlanImmoOrig.Charge(QOrigine);
  EnregCession.PlanPrec:=PlanImmoOrig.NumSeq ;
  PlanImmoOrig.Recupere(PlanImmoOrig.CodeImmo,IntToStr(PlanImmoOrig.NumSeq));
  PlanImmoOrig.TypeOpe:='CES';
  PlanImmoOrig.SetBaseAmortFin(EnregCession.Vo);
  PlanImmoOrig.AmortEco.Reprise:=(QOrigine.FindField('I_REPRISEECO').AsFloat);
  PlanImmoOrig.AmortFisc.Reprise:=(QOrigine.FindField('I_REPRISEFISCAL').AsFloat);
  PlanImmoOrig.CalculProrataDotationOrigine(PlanImmoOrig.AmortEco,1-wCoef);
  PlanImmoOrig.CalculProrataDotationOrigine(PlanImmoOrig.AmortFisc,1-wCoef);
  PlanImmoOrig.Calcul(QOrigine, EnregCession.DateOpe);

  if CessionPartielle then
  begin
    QOrigine.FindField('I_DATEDERMVTECO').AsDateTime := PlanImmoOrig.GetDateFinAmortEx(PlanImmoOrig.AmortEco);
    QOrigine.FindField('I_DATEDERNMVTFISC').AsDateTime := PlanImmoOrig.GetDateFinAmortEx(PlanImmoOrig.AmortFisc);
  end
  else
  begin
    QOrigine.FindField('I_DATEDERMVTECO').AsDateTime := EnregCession.DateOpe;
    if PlanImmoOrig.Fiscal then QOrigine.FindField('I_DATEDERNMVTFISC').AsDateTime := EnregCession.DateOpe;
  end;
  if CessionPartielle then PlanImmoOrig.BasculeDotationCedee(PlanImmoCedee,EnregCession.DateOpe)
  else PlanImmoOrig.AffecteDotationCedee(PlanImmoCedee,EnregCession.DateOpe);
  PlanImmoOrig.Sauve;
  EnregCession.PlanSuiv:=PlanImmoOrig.NumSeq ;
  QOrigine.FindField('I_PLANACTIF').AsInteger:=EnregCession.PlanSuiv;
  //YVP 09/07/01
  if PlanImmoOrig.AmortFisc.Methode<>'' then
    wVNC := PlanImmoOrig.GetVNCAvecMethode(PlanImmoOrig.AmortFisc,EnregCession.DateOpe)
  else
    wVNC := PlanImmoOrig.GetVNCAvecMethode(PlanImmoOrig.AmortEco,EnregCession.DateOpe);
  //YCP 09/07/01 fin
  QOrigine.FindField('I_VNC').AsFloat := wVNC;
  CocheChampOperation (QOrigine,'','I_OPECESSION');
  if EnregCession.ChangePlan then CocheChampOperation (QOrigine,'','I_OPECHANGEPLAN');
  QOrigine.Post ;
  PlanImmoOrig.free ;
  Ferme(QOrigine);
end;

procedure TraiteChangementPlan (Plan : TPlanAmort);
var QOrigine : TQuery;
begin
  QOrigine :=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+Plan.CodeImmo+'"', FALSE);
  QOrigine.Edit ;
  QOrigine.FindField('I_TYPEEXC').AsString:=Plan.TypeExc;
  QOrigine.FindField('I_MONTANTEXC').AsFloat:= Plan.MontantExc;
  Plan.SortiePlusExcep := True;
  Plan.Calcul(QOrigine, iDate1900);
  CocheChampOperation (QOrigine,'','I_OPECHANGEPLAN');
  QOrigine.Post ;
  Ferme(QOrigine);
end;

procedure TFSortie.DotationKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  //if Key=VK_F3 then GereDotationMax; mbo - FQ 16776
  if Key=VK_F6 then GereDotationMax;
end;

procedure TFSortie.GereDotationMax;
begin
  // Dotation.Value:=VNC.Value; si pas de derog ou si derog + amort variable
  //  mbo 27.09.2005 Dotation.Value := fVNC; // FQ 12083
  if rbAjouter.Checked then
     Dotation.Value := MaxAJout
  else
     Dotation.Value := MaxRemplace;
  DotTest := Dotation.Value; // fq 17155 - ajout mbo 7.12.05 

  // mbo - fq 15280 - suite autorise saisie exceptionnel sur dérogatoire
//  if (existeSQL('SELECT I_METHODEFISC FROM IMMO WHERE I_IMMO="'+fCode+'" AND I_METHODEFISC<>""')) then
//    if (existeSQL('SELECT I_METHODEECO FROM IMMO WHERE I_IMMO="'+fCode+'" AND I_METHODEECO <> "VAR"')) then
//         Dotation.Value := (RepriseFisc + fVNC;
  RecalculDotation;
  RecalculDetailDotation;
end;

procedure TFSortie.IL_TVAAREVERSEREnter(Sender: TObject);
begin
  inherited;
  fCurVCedee := IL_TVAAREVERSER.Value;
end;

procedure TFSortie.IL_TVAAREVERSERExit(Sender: TObject);
begin
  inherited;
  fTVA := Valeur(IL_TVAAREVERSER.Text);
  RecalculDotation;
end;

procedure TFSortie.IL_MONTANTCESExit(Sender: TObject);
begin
  inherited;
  CalculRepartition ( False );
end;

procedure TFSortie.IL_CALCCESSIONChange(Sender: TObject);
begin
  inherited;
  Dotation.Value := 0;
  DotTest := 0; // fq 17155 - ajout mbo 7.12.05
  RecalculDotation;
  RecalculDetailDotation;
  CalculRepartition ( True );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/06/2004
Modifié le ... : 09/06/2004
Description .. : Initialise la grille de répartition
Mots clefs ... : 
*****************************************************************}
procedure TFSortie.InitGrilleRepartition;
begin
  { Affichage de la grille }
    { Colonne n°0 = Code de l'immobilisation }
  LISTEREPARTITION.ColEditables[0]:=False;
  LISTEREPARTITION.ColWidths[0]:=70;
  LISTEREPARTITION.ColAligns[0]:=taLeftJustify;
  LISTEREPARTITION.Cells[0,0] := TraduireMemoire('Code');

    { Colonne n°1 = Libellé de l'immobilisation }
  LISTEREPARTITION.ColEditables[1]:=False;
  LISTEREPARTITION.ColWidths[1]:=100;
  LISTEREPARTITION.ColAligns[1]:=taLeftJustify;
  LISTEREPARTITION.Cells[1,0] := TraduireMemoire('Libellé');

    { Colonne n°2 = Valeur d'achat }
  LISTEREPARTITION.ColEditables[2]:=False;
  LISTEREPARTITION.ColWidths[2]:=80;
  LISTEREPARTITION.ColAligns[2]:=taRightJustify;
  LISTEREPARTITION.Cells[2,0] := TraduireMemoire('Val. d''achat');

    { Colonne n°3 = VNC }
  LISTEREPARTITION.ColEditables[3]:=False;
  LISTEREPARTITION.ColWidths[3]:=80;
  LISTEREPARTITION.ColAligns[3]:=taRightJustify;
  LISTEREPARTITION.Cells[3,0] := TraduireMemoire('VNC');

    { Colonne n°4 = Prix de sortie }
  LISTEREPARTITION.ColEditables[4]:=False;
  LISTEREPARTITION.ColWidths[4]:=80;
  LISTEREPARTITION.ColAligns[4]:=taRightJustify;
  LISTEREPARTITION.Cells[4,0] := TraduireMemoire('Prix de sortie');
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/06/2004
Modifié le ... :   /  /
Description .. : Calcul de la répartition du prix de sortie en fonction du mode
Suite ........ : de répartition choisi
Mots clefs ... :
*****************************************************************}
procedure TFSortie.CalculRepartition ( bRecalculVNC : boolean );
var i : integer;
    TotalMontant , PrixSortie : double;
    PlanInfo : TPlanInfo;
begin
  if fRegroupement='' then exit;
  if ((fTRepartition=nil) or (fTRepartition.Detail.Count = 0)) then exit;
  { Initialisation }
  TotalMontant := 0;
  LISTEREPARTITION.ColEditables[4]:=(MODEREPARTITION.Value='LIB');

  { Recalcul des VNC }
  if (bRecalculVNC and IsValidDate(DATEOPE.Text)) then
  begin
    PlanInfo := TPlanInfo.Create ('');
    for i:=0 to fTRepartition.Detail.Count - 1 do
    begin
      PlanInfo.ChargeImmo (fTRepartition.Detail[i].GetValue('I_IMMO'));
      PlanInfo.Calcul (StrToDate(DATEOPE.Text),False, true, AmortJsortie); //fq 21754 mbo 
      if IL_CALCCESSION.Value = 'SAN' then fTRepartition.Detail[i].AddChampSupValeur('I_VNC',PlanInfo.VNCEco+PlanInfo.DotationEco)
      else fTRepartition.Detail[i].AddChampSupValeur('I_VNC',PlanInfo.VNCEco);
    end;
    PlanInfo.Free;
  end;

  { Calcul des prix de sortie }
  if (MODEREPARTITION.Value = 'VNC') then
    TotalMontant := Arrondi(fTRepartition.Somme('I_VNC',[''],[''],False), V_PGI.OkDecV)
  else if MODEREPARTITION.Value = 'ACH' then
    TotalMontant := Arrondi(fTRepartition.Somme('I_MONTANTHT',[''],[''],False), V_PGI.OkDecV);
  for i:=0 to fTRepartition.Detail.Count - 1 do
  begin
    if MODEREPARTITION.Value = 'LIB' then PrixSortie := 0
    else if MODEREPARTITION.Value = 'VNC' then
    begin
      if ( TotalMontant = 0 ) then PrixSortie := Arrondi ( ( IL_MONTANTCES.Value / fTRepartition.Detail.Count ) , V_PGI.OkDecV )
      else PrixSortie := Arrondi ( (fTRepartition.Detail[i].GetValue('I_VNC')*IL_MONTANTCES.Value)/TotalMontant, V_PGI.OkDecV);
    end
    else if MODEREPARTITION.Value = 'ACH' then PrixSortie := Arrondi ( (fTRepartition.Detail[i].GetValue('I_MONTANTHT')*IL_MONTANTCES.Value)/TotalMontant, V_PGI.OkDecV)
    else PrixSortie := 0;
    fTRepartition.Detail[i].AddChampSupValeur ('PRIXSORTIE',PrixSortie);
  end;
  if MODEREPARTITION.Value<>'LIB' then
  begin
    SoldePrixSortie ( True );
    fTRepartition.PutGridDetail(LISTEREPARTITION,False,False,'I_IMMO;I_LIBELLE;I_MONTANTHT;I_VNC;PRIXSORTIE');
  end;
end;

procedure TFSortie.MODEREPARTITIONChange(Sender: TObject);
begin
  inherited;
  CalculRepartition ( False );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/06/2004
Modifié le ... : 10/06/2004
Description .. : Enregistre la sortie pour une immobilisation dans le cadre de
Suite ........ : la sortie complète d'un regroupement
Mots clefs ... :  ATTENTION : cette procedure n'est plus utilisée
*****************************************************************}
{ n'est plus utilisée

procedure TFSortie.EnregistreLaSortieRegroupement(CodeImmo: string; PrixSortie, VNC : double);
var TImmo : TOB;
    EnregCession : TEnregCession;
    PlanImmoCedee : TPlanAmort ;
begin
  TImmo := TOB.Create ('IMMO', nil, -1);
  try
    if TImmo.SelectDB('"'+CodeImmo+'"',nil,True) then
    begin
      EnregCession                  :=TEnregCession.Create;
      try
        EnregCession.Code           :=TImmo.GetValue('I_IMMO');
        EnregCession.Vo             :=TImmo.GetValue('I_MONTANTHT');
        EnregCession.QteAchat       :=TImmo.GetValue('I_QUANTITE');
        EnregCession.QteCedee       :=TImmo.GetValue('I_QUANTITE');
        EnregCession.DateOpe        :=StrToDate(DATEOPE.Text);
        EnregCession.Tva            :=0;
        EnregCession.MontantHT      :=TImmo.GetValue('I_MONTANTHT');
        EnregCession.MontantAchat   :=TImmo.GetValue('I_MONTANTHT');
        EnregCession.MontantCession :=PrixSortie;
        EnregCession.MotifCession   :=IL_MOTIFCES.Value ;
        EnregCession.CalcCession    :=IL_CALCCESSION.Text ;
        EnregCession.TypeCession    :=IL_CALCCESSION.Value;
        EnregCession.BlocNote       :=IL_BLOCNOTE.LinesRTF;
        EnregCession.ChangePlan     :=False;
        EnregCession.MontantExc     :=0;
        EnregCession.TypeExc        :='';
        EnregCession.Dotation       :=0;  // Inutile si pas de changement de plan
        EnregCession.BaseEco        :=TImmo.GetValue('I_BASEECO');
        EnregCession.BaseFisc       :=TImmo.GetValue('I_BASEFISC');
        EnregCession.BaseTP         :=TImmo.GetValue('I_BASETAXEPRO');
        EnregCession.VNCCedee       :=VNC;

        // BTY 04/06 FQ 17629
        EnregCession.RegleCession   :=ReglePValue.Value ;

        PlanImmoCedee:=TPlanAmort.Create(true) ;
        try
          TraiteImmoCedee(PlanImmoCedee,EnregCession);
          fTypeExc := '';
          TraiteImmoOrig(PlanImmoCedee,False,EnregCession);
          //EnregCession.PValue:=PrixSortie - (EnregCession.VNCCedee + TImmo.GetValue('I_MONTANTHT') + TImmo.GetValue('I_BASEECO') );
          EnregCession.PValue:=PrixSortie - (EnregCession.VNCCedee); // FQ 12083
          EnregLogCession(PlanImmoCedee,False,EnregCession,fOrdreS);
        finally
          PlanImmoCedee.free ;
        end;
      finally
        EnregCession.free ;
      end;
    end;
  finally
    TImmo.Free;
  end;
end;
}


procedure TFSortie.LISTEREPARTITIONKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  if Key=VK_F6 then SoldePrixSortie ( False );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/06/2004
Modifié le ... :   /  /
Description .. : Solde le prix de sortie sur la ligne en cours
Mots clefs ... :
*****************************************************************}
procedure TFSortie.SoldePrixSortie ( bDerniereLigne : boolean );
var lEncours : integer;
    i : integer;
    TotalPrix, EcartPrix : double;
begin
  if (fRegroupement = '') then exit;
  TotalPrix := 0;
  if bDerniereLigne then lEnCours := LISTEREPARTITION.RowCount - 1
  else lEncours := LISTEREPARTITION.Row;
  for i:=0 to fTRepartition.Detail.Count - 1 do
    TotalPrix := TotalPrix + fTRepartition.Detail[i].GetValue('PRIXSORTIE');
  EcartPrix := Arrondi(IL_MONTANTCES.Value - TotalPrix, V_PGI.OkDecV);
  if EcartPrix > 0 then
    fTRepartition.Detail[lEnCours-1].PutValue('PRIXSORTIE', Arrondi(EcartPrix+fTRepartition.Detail[lEnCours-1].GetValue('PRIXSORTIE'), V_PGI.OkDecV))
  else fTRepartition.Detail[lEnCours-1].PutValue('PRIXSORTIE', Arrondi(EcartPrix+fTRepartition.Detail[lEnCours-1].GetValue('PRIXSORTIE'), V_PGI.OkDecV));
  LISTEREPARTITION.Cells[4,lEncours] := StrFMontant ( fTRepartition.Detail[lEnCours-1].GetValue('PRIXSORTIE'), 15 , V_PGI.OkDecV, '' , true);
  LISTEREPARTITION.Refresh;
end;

procedure TFSortie.LISTEREPARTITIONCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var RdMontant : double;
begin
  inherited;
  case ACol of
    4 :
      if IsNumeric(LISTEREPARTITION.Cells[ACol,ARow]) then
      begin
        RdMontant := Valeur(LISTEREPARTITION.Cells[ACol,ARow]);
        Cancel    := RdMontant < 0; // on ne peut pas saisir de valeur negative
        if Cancel then
          PGIInfo('Montant négatif !','Attention');
        if Arrondi(RdMontant,V_PGI.OkDecV) = 0 then
        begin
          LISTEREPARTITION.Cells[ACol,ARow] := '';
          fTRepartition.Detail[ARow-1].PutValue('PRIXSORTIE',0);
        end
        else
        begin
          LISTEREPARTITION.Cells[ACol,ARow] := STRFMONTANT ( RdMontant , 15 , V_PGI.OkDecV, '' , true);
          fTRepartition.Detail[ARow-1].PutValue('PRIXSORTIE',RdMontant);
        end;
      end;
  end;
end;

procedure TFSortie.ExecuteSortie;
var
  DP_BlocNote : THRichEditOle;  // FQ 18937
  //BlocNote : HTStringList ;   FQ 18937
  LaSortie : TAmSortie;
  MontantExc : double;
  stTypeExc : string;
  i : integer;

begin


  if fRegroupement <> '' then
  { Sortie complète du regroupement }
  begin
    { Calcul du numéro d'ordre pour opération en série }
    fOrdreS := TrouveNumeroOrdreSerieLogSuivant;
    { Récupération des informations de la grille de saisie de la répartition }
    fTRepartition.GetGridDetail(ListeRepartition,ListeRepartition.RowCount-1,'','I_IMMO;I_LIBELLE;I_MONTANTHT;I_VNC;PRIXSORTIE',True);
    { Enregistrement des sortie pour chaque immobilisation concernée }
    LaSortie := TAmSortie.Create ('');
    for i:=0 to fTRepartition.Detail.Count - 1 do
    begin
      LaSortie.Charge(fTRepartition.Detail[i].GetValue('I_IMMO'));

      // BTY 04/06 FQ 17269 Nouveau paramètre de règle de cession
      //LaSortie.Init ( StrToDate(DATEOPE.Text), IL_MOTIFCES.Value, IL_CALCCESSION.Value, fTRepartition.Detail[i].GetValue('PRIXSORTIE'), IL_TVAAREVERSER.Value);

      //fq 21754 - mbo - 30/10/2007 ajout du paramètre amortJsortie
      LaSortie.Init ( StrToDate(DATEOPE.Text), IL_MOTIFCES.Value, IL_CALCCESSION.Value,
                      fTRepartition.Detail[i].GetValue('PRIXSORTIE'), IL_TVAAREVERSER.Value,
                      ReglePValue.Value, AmortJsortie);

      // FQ 18937  Ecriture non standard du blocnotes
      //BlocNote := HTStringList.Create ;
      //RichToStrings(THRichEdit(IL_BLOCNOTE),BlocNote) ;
      //LaSortie.BlocNote := BlocNote;
      //FreeAndNil(BlocNote) ;
      DP_BlocNote := THRichEditOle (IL_BLOCNOTE);
      LaSortie.BlocNote :=   RichToString (DP_BLOCNOTE);
      //
      LaSortie.Execute;
    end;
    FreeAndNil(LaSortie);
  end else
  begin
    MontantExc := 0;
    LaSortie := TAmSortie.Create (fCode);
    // BTY 04/06 FQ 17269 Nouveau paramètre de règle de cession
    //LaSortie.Init ( StrToDate(DATEOPE.Text), IL_MOTIFCES.Value, IL_CALCCESSION.Value, IL_MONTANTCES.Value, IL_TVAAREVERSER.Value);
    LaSortie.Init ( StrToDate(DATEOPE.Text), IL_MOTIFCES.Value, IL_CALCCESSION.Value,
                    IL_MONTANTCES.Value, IL_TVAAREVERSER.Value, ReglePValue.Value, AmortJsortie);
    // mbo 15280 dans excep.value il y a élément excep + excep sortie - if Excep.Value > 0  then
    if Arrondi((Excep.Value-fExcep),V_PGI.OkDecV) > 0 then
    begin
      stTypeExc := 'DOT';
      // mbo 15280 MontantExc := Excep.Value;
      MontantExc := Arrondi((Excep.Value-fExcep),V_PGI.OkDecV)
    end
    else if Arrondi((Excep.Value-fExcep),V_PGI.OkDecV) < 0 then
    begin
      stTypeExc := 'RDO';
      // mbo 15280 MontantExc := (-1)*Excep.Value;
      MontantExc := (-1)* Arrondi((Excep.Value-fExcep),V_PGI.OkDecV);
    end;
    LaSortie.InitExceptionnel ( MontantExc, stTypeExc  );
    // FQ 18937  Ecriture non standard du blocnotes
    //BlocNote := HTStringList.Create ;
    //RichToStrings(THRichEdit(IL_BLOCNOTE),BlocNote) ;
    //LaSortie.BlocNote := BlocNote;
    //FreeAndNil(BlocNote) ;
    DP_BlocNote := THRichEditOle (IL_BLOCNOTE);
    LaSortie.BlocNote :=   RichToString (DP_BLOCNOTE);
    //
    LaSortie.Execute;
    FreeAndNil(LaSortie);
  end;
end;


end.
