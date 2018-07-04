unit MulImmo;

// CA - 18/05/1999 - Renseignement de la nature pour mutation en série
// CA - 02/07/1999 - Ajout de l'appel à la génération des écritures.
// CA - 05/07/1999 - Operation en cours = TypeOp <> "ACQ"
// CA - 07/07/1999 - Repositionnement sur la dernière ligne sélectionnée dans le Mul
// CA - 09/07/1999 - Intégration : paramètres supplémentaires : intégration dans la compta et détail
// CA - 09/07/1999 - Ajout appel écritures dans popup
// CA - 09/07/1999 - ElementExceptionnel initialisé à false dans popup
// CA - 09/07/1999 - Etalonnage barre d'avancement
// CA - 15/08/2000 - Relance de la requête pour les sorties en série (plus de pb de maj de Liste).
// CA - 18/07/2003 - Suppression champs euro
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, Db, DBTables, Hqry, ComCtrls, HRichEdt, HTB97,
  Grids, DBGrids, HDB, StdCtrls, ColMemo, Hctrls, ExtCtrls, Buttons, Mask,
  HEnt1, hmsgbox, ParamDat, Outils, ImEnt, Hstatus, UiUtil,
  HPanel,OpeSerie,LookUp,ParamSoc, HRichOLE,FE_Main
  {$IFDEF SERIE1}
  , S1Util
  {$ENDIF}
  ;

procedure ConsultationImmo(Comment : TActionFiche; Operation : TypeOperation; Bureau : boolean; Compte : string);

type
  TFMulImmo = class(TFMul)
    I_IMMO: TEdit;
    I_LIBELLE: TEdit;
    HM: THMsgBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    Nature: THLabel;
    HLabel3: THLabel;
    tI_ETABLISSEMENT: THLabel;
    I_ETABLISSEMENT: THValComboBox;
    HLabel7: THLabel;
    I_DATEAMORT: THCritMaskEdit;
    Label10: TLabel;
    I_DATEAMORT_: THCritMaskEdit;
    I_DATEPIECEA_: THCritMaskEdit;
    Label12: TLabel;
    HLabel8: THLabel;
    I_DATEPIECEA: THCritMaskEdit;
    I_NATUREIMMO: THValComboBox;
    I_QUALIFIMMO: THValComboBox;
    XX_WHERE: TEdit;
    PopZoomAction: TPopupMenu;
    Cession: TMenuItem;
    Mutation: TMenuItem;
    Eclatement: TMenuItem;
    ChangePlan: TMenuItem;
    LeveeOption: TMenuItem;
    AnnulationOperation: TMenuItem;
    SeparAnnulAction: TMenuItem;
    DureeMethode: TMenuItem;
    ElementExceptionnel: TMenuItem;
    PopZoomVisu: TPopupMenu;
    ZoomPlan: TMenuItem;
    ZoomEcheance: TMenuItem;
    SeparPlanVisu: TMenuItem;
    ZoomOperations: TMenuItem;
    BZoomAction: TToolbarButton97;
    BZoomVisu: TToolbarButton97;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TI_TABLE0: TLabel;
    I_TABLE0: THCritMaskEdit;
    TI_TABLE5: TLabel;
    I_TABLE5: THCritMaskEdit;
    I_TABLE6: THCritMaskEdit;
    TI_TABLE6: TLabel;
    I_TABLE1: THCritMaskEdit;
    TI_TABLE1: TLabel;
    TI_TABLE2: TLabel;
    TI_TABLE3: TLabel;
    TI_TABLE4: TLabel;
    I_TABLE4: THCritMaskEdit;
    I_TABLE3: THCritMaskEdit;
    I_TABLE2: THCritMaskEdit;
    TI_TABLE7: TLabel;
    I_TABLE7: THCritMaskEdit;
    I_TABLE8: THCritMaskEdit;
    TI_TABLE8: TLabel;
    TI_TABLE9: TLabel;
    I_TABLE9: THCritMaskEdit;
    POperations: TTabSheet;
    Bevel6: TBevel;
    HLabel5: THLabel;
    PopZoomCreation: TPopupMenu;
    PleinePropriete: TMenuItem;
    CreditBail: TMenuItem;
    LocationFinanciere: TMenuItem;
    ImmoFinanciere: TMenuItem;
    N1: TMenuItem;
    CreationSerie: TMenuItem;
    PopupListe: TPopupMenu;
    HLabel4: THLabel;
    ZoomPlan1: TMenuItem;
    ZoomEcheance1: TMenuItem;
    SeparOpe: TMenuItem;
    ZoomOperations1: TMenuItem;
    N3: TMenuItem;
    Mutation1: TMenuItem;
    Eclatement1: TMenuItem;
    Cession1: TMenuItem;
    ChangePlan1: TMenuItem;
    DureeMethode1: TMenuItem;
    ElementExceptionnel1: TMenuItem;
    LeveeOption1: TMenuItem;
    SeparAnnul: TMenuItem;
    AnnulationOperation1: TMenuItem;
    HLabel11: THLabel;
    ChangeLieu: TMenuItem;
    ChangeEtablissement: TMenuItem;
    ChangeLieu1: TMenuItem;
    ChangeEtablissement1: TMenuItem;
    BDelete: TToolbarButton97;
    SeparPlan: TMenuItem;
    PleinePropriete1: TMenuItem;
    ImmoFinanciere1: TMenuItem;
    CreditBail1: TMenuItem;
    LocationFinanciere1: TMenuItem;
    N6: TMenuItem;
    CreationSerie1: TMenuItem;
    I_LIEUGEO: THValComboBox;
    OPEECLATEMENT: TCheckBox;
    OPECHANGEPLAN: TCheckBox;
    OPELIEUGEO: TCheckBox;
    OPEETABLISSEMENT: TCheckBox;
    OPELEVEEOPTION: TCheckBox;
    OPECESSION: TCheckBox;
    OPEMUTATION: TCheckBox;
    HLabel9: THLabel;
    FExercice2: THValComboBox;
    PopZoomFiltreOpe: TPopupMenu;
    SelectOpe: TMenuItem;
    ToutesOpe: TMenuItem;
    ToolbarButton971: TToolbarButton97;
    DupliqueImmo: TMenuItem;
    DupliqueImmoListe: TMenuItem;
    AucuneOpe: TMenuItem;
    ModifBases: TMenuItem;
    ModifBases1: TMenuItem;
    HPanel1: THPanel;
    sOperationEnCours: THLabel;
    HM2: THMsgBox;
    REINTEGRATION: TCheckBox;
    QUOTEPART: TCheckBox;
    FiltreOpe: TCheckBox;
    HLabel12: THLabel;
    I_METHODEECO: THValComboBox;
    I_METHODEFISC: THValComboBox;
    HLabel13: THLabel;
    AMORTDEROG: TCheckBox;
    HLabel10: THLabel;
    I_COMPTELIE: THCritMaskEdit;
    I_COMPTEIMMO: THCritMaskEdit;
    SeparOpeVisu: TMenuItem;
    Ecritures: TMenuItem;
    N8: TMenuItem;
    Ecritures1: TMenuItem;
    I_ORGANISMECB: THCritMaskEdit;
    PzlibreS1: TTabSheet;
    Bevel7: TBevel;
    TABLELIBRE1: THValComboBox;
    TT_TABLELIBREIMMO1: THLabel;
    TABLELIBRE2: THValComboBox;
    TT_TABLELIBREIMMO2: THLabel;
    TABLELIBRE3: THValComboBox;
    TT_TABLELIBREIMMO3: THLabel;
    BREINIT: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override ;
    procedure I_DATEAMORTKeyPress(Sender: TObject; var Key: Char);
    procedure CessionClick(Sender: TObject);
    procedure MutationClick(Sender: TObject);
    procedure EclatementClick(Sender: TObject);
    procedure LeveeOptionClick(Sender: TObject);
    procedure AnnulationOperationClick(Sender: TObject);
    procedure DureeMethodeClick(Sender: TObject);
    procedure ElementExceptionnelClick(Sender: TObject);
    procedure PopZoomActionPopup(Sender: TObject);
    procedure ZoomPlanClick(Sender: TObject);
    procedure ZoomEcheanceClick(Sender: TObject);
    procedure ZoomOperationsClick(Sender: TObject);
    procedure PopZoomVisuPopup(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure CreationClick(Sender: TObject);
    procedure CreationSerieClick(Sender: TObject);
    procedure PleineProprieteClick(Sender: TObject);
    procedure ImmoFinanciereClick(Sender: TObject);
    procedure CreditBailClick(Sender: TObject);
    procedure LocationFinanciereClick(Sender: TObject);
    procedure bSupprimeClick(Sender: TObject);
    procedure SuppressionFicheImmo;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BOuvrirClick(Sender: TObject); override ;
    procedure DessineListeImmo(ACol,ARow : LongInt; Canvas : TCanvas);
    procedure PopupListePopup(Sender: TObject);
    procedure ChangeEtablissementClick(Sender: TObject);
    procedure ChangeLieuClick(Sender: TObject);
    procedure ExecuteOperationSerie(TypeOpe : TypeOperation);
    procedure SQDataChange(Sender: TObject; Field: TField);
    procedure FExercice2Change(Sender: TObject);
    procedure ToutesOpeClick(Sender: TObject);
    procedure SelectOpeClick(Sender: TObject);
    procedure AucuneOpeClick(Sender: TObject);
    procedure PopZoomCreationPopup(Sender: TObject);
    procedure DupliqueImmoClick(Sender: TObject);
    procedure OpeIndicateurClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject); override ;
    procedure BNouvRechClick(Sender: TObject);
    procedure ModifBasesClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject); override ;
    procedure BSaveFiltreClick(Sender: TObject);
    procedure AMORTDEROGClick(Sender: TObject);
    procedure OnCompteElipsisClick(Sender: TObject);
    procedure EcrituresClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TABLELIBRE1Change(Sender: TObject);
    procedure I_ORGANISMECBElipsisClick(Sender: TObject);
    procedure FormCreate(Sender: TObject); override ;
    procedure BREINITClick(Sender: TObject);
  private
    FCodeImmo : string;
    TypeOp : TypeOperation ;
    fBureau : boolean;
    XX_Where_Ope : string;
    XX_Where_Orig : string;
    XX_Where_OpeAucune : string;
    WhereDerog : string;
    procedure EnableZoomAction;
    procedure EnableZoomVisu;
    procedure EnableZoomActionSerie;
    procedure EnableZoomVisuSerie;
    //function RecupereListeImmo : TStrings;
    procedure RechargeLieuGeographique; //EPZ 15/01/99
    procedure GereFiltreOpeEnCours(Etat : TCheckBoxState);
    procedure EnableCheckBoxOperation;
    procedure EnableBoutonsAction(Action: TActionFiche);
  public
  end;

function VerifNature (NatureIni, Nature : string) : boolean;

implementation

uses Imogen, ImSortie, ImMutati, ImEclate, ChanPlan,  ImLevOpt,OpEnCour ,ImPlan,
     ListeEch,SupImmo,ImGenEcr,ImDupImo,ImModBas,PlanAmor,ImContra,ImOutGen,ImCreSer
{$IFDEF AMORTISSEMENT}
     ,AMIMMO, AmExport
{$ENDIF}
    {$IFDEF SERIE1}
    ,Integecr ;
    {$ELSE}
    ,Integecr ;
    {$ENDIF}

{$R *.DFM}

procedure ConsultationImmo(Comment : TActionFiche; Operation : TypeOperation; Bureau : boolean; Compte : string);
var FMulImmo: TFMulImmo;
    PP : THPanel;
begin
     if GetParamSoc('SO_EXOCLOIMMO')=VHImmo^.Encours.Code then Comment := taConsult;
     FMulImmo:=TFMulImmo.Create(Application) ;
     if Compte <> '' then
     begin
       FMulImmo.I_COMPTEIMMO.Text := Compte;
       FMulImmo.I_COMPTEIMMO.Enabled := false;
     end;
     FMulImmo.CheckBoxStyle := csCheckBox;
     FMulImmo.fBureau := Bureau;
     FMulImmo.TypeAction:=Comment;
     FMulImmo.TypeOp:=Operation ;
     Case Comment Of
     taConsult : begin
                   FMulImmo.Caption:=FMulImmo.HM.Mess[0] ;
                   FMulImmo.FNomFiltre:='MULVIMMOS' ;
                   FMulImmo.Q.Liste:='MULVIMMOS' ;
                 end ;
     taModif :   begin
                   FMulImmo.Caption:=FMulImmo.HM.Mess[1] ;
                   FMulImmo.FNomFiltre:='MULMIMMOS' ;
                   FMulImmo.Q.Liste:='MULMIMMOS' ;
                 end ;
     end;
     PP:=FindInsidePanel ;
     if (PP=Nil) or (V_PGI.ZoomOLE) then
     BEGIN
          try
          FMulImmo.ShowModal ;
          finally
            FMulImmo.Free ;
          end ;
     END else
     BEGIN
          InitInside(FMulImmo,PP) ;
          FMulImmo.Show ;
     END ;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulImmo.FormShow(Sender : TObject) ;
begin
  MakeZoomOLE(Handle) ;
  {$IFDEF SERIE1}
  I_ETABLISSEMENT.Visible := False;
  tI_ETABLISSEMENT.Visible := False;
  OPEETABLISSEMENT.Visible:=false ;
  {$ENDIF}
  I_DATEAMORT.text:=StDate1900 ; I_DATEAMORT_.text:=StDate2099 ;
  I_DATEPIECEA.text:=StDate1900 ; I_DATEPIECEA_.text:=StDate2099 ;
  I_METHODEFISC.Items.Delete(I_METHODEFISC.Values.IndexOf('NAM'));
  AMORTDEROG.State := cbGrayed;
  if XX_WHERE.Text = '' then  XX_WHERE.Text:='I_ETAT<>"FER"'
  else XX_WHERE.Text:=XX_WHERE.Text+' AND I_ETAT<>"FER"';
  // On ne passe que toNone en paramètre. Cette ligne n'a donc pas lieu d'être.
  //  if TypeOp<>toNone then begin Q.TransformSQL ; Q.Manuel:=FALSE ; Q.UpdateCriteres ; end ;
  BInsert.Visible:=fBureau;
  BZoomAction.Visible:=fBureau;
  bDelete.Visible:= (fBureau) ;
  bSelectAll.Visible:=fBureau;
  XX_Where_OpeAucune := ' AND (I_OPEMUTATION="-" AND I_OPECHANGEPLAN="-" AND ' +
                        'I_OPELIEUGEO="-" AND I_OPEETABLISSEMENT="-" AND ' +
                        'I_OPELEVEEOPTION="-" AND I_OPECESSION="-" AND '+
                        'I_OPEECLATEMENT="-" AND I_REINTEGRATION=0.0 AND '+
                        'I_QUOTEPART=0)';
  //EPZ 15/01/99
  XX_Where_Orig := XX_WHERE.Text;
  inherited ;
  ToutesOpe.Checked := FiltreOpe.State = cbChecked;
  AucuneOpe.Checked := FiltreOpe.State = cbGrayed;
  SelectOpe.Checked := FiltreOpe.State = cbUnchecked;
  EnableCheckBoxOperation;
  AMORTDEROGClick(nil);
  EnableBoutonsAction (TypeAction);
{$IFNDEF AMORTISSEMENT}
  BReinit.Visible := False;
{$ENDIF}
end;

procedure TFMulImmo.FListeDblClick(Sender: TObject);
var
  stCode : string;
begin
  if Q.Eof and Q.Bof then Exit ;
  inherited;
  stCode:=Q.FindField ('I_IMMO').AsString;
  if TypeOp=toNone then FicheImmobilisation(Q,Q.FindField('I_IMMO').AsString,TypeAction,'') ;
  BChercheClick(Sender);
  RechargeLieuGeographique;
  EnableBoutonsAction(TypeAction) ;
  Q.Locate('I_IMMO',stCode,[]);
end;

procedure TFMulImmo.I_DATEAMORTKeyPress(Sender: TObject; var Key: Char);
begin
  inherited;
ParamDate(Self,Sender,Key) ;
end;

procedure TFMulImmo.CessionClick(Sender: TObject);
var mr,i : integer; stCode : string;
begin
  inherited;
  stCode:=Q.FindField('I_IMMO').AsString;
  mr:=mrYes;
  if FListe.AllSelected then
    begin
    Q.First;
    while not Q.EOF do
      begin
      mr:=ExecuteCession(Q.FindField('I_IMMO').AsString) ;
      if mr<>mrYes then break ;
      Q.Next;
      end;
    FListe.AllSelected:=False;
    end
  else if (FListe.NbSelected>=1) then
    begin
    for i:=0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookmark(i);
      mr:=ExecuteCession(Q.FindField('I_IMMO').AsString) ;
      if mr<>mrYes then break ;
      end;
    FListe.ClearSelected;
    end
  else
    mr := ExecuteCession(Q.FindField('I_IMMO').AsString) ;
  if mr=mrYes then
    begin
    BChercheClick(Sender);
    Q.Locate('I_IMMO',stCode,[]);
    end;
end;

procedure TFMulImmo.MutationClick(Sender: TObject);
var mr : integer;
    stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  mr := mrYes;
  if (FListe.NbSelected<=1) and (not FListe.AllSelected) then
    mr := ExecuteMutation(Q.FindField('I_IMMO').AsString)
  else ExecuteOperationSerie(toMutation);
  if mr = mrYes then
  begin
    BChercheClick(Sender);
    Q.Locate('I_IMMO',stCode,[]);
  end;
end;

procedure TFMulImmo.EclatementClick(Sender: TObject);
var mr : integer;
    stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  mr := ExecuteEclatement(Q.FindField('I_IMMO').AsString) ;
  if mr = mrYes then
  begin
    BChercheClick(Sender);
    Q.Locate('I_IMMO',stCode,[]);
  end;
end;

procedure TFMulImmo.LeveeOptionClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if ExecuteLeveeOption (Q.FindField('I_IMMO').AsString) = mrYes then
  begin
    BChercheClick(Sender);
    Q.Locate('I_IMMO',stCode,[]);
  end;
end;

procedure TFMulImmo.AnnulationOperationClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  OperationsEnCours (Q.FindField('I_IMMO').AsString,Q.FindField('I_LIBELLE').AsString,true,TypeAction   );
  BChercheClick(Sender);
  Q.Locate('I_IMMO',stCode,[]);
end;

procedure TFMulImmo.DureeMethodeClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if ExecuteChangePlan(Q.FindField('I_IMMO').AsString,1) = mrYes then
  begin
    BChercheClick(Sender);
    Q.Locate('I_IMMO',stCode,[]);
  end;
end;

procedure TFMulImmo.ElementExceptionnelClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if ExecuteChangePlan(Q.FindField('I_IMMO').AsString,2) = mrYes then
  begin
    BChercheClick(Sender);
    Q.Locate('I_IMMO',stCode,[]);
  end;
end;

procedure TFMulImmo.EnableZoomAction;
var Nat : string;
    bDepotGar,bAcquisitionAnnee : boolean;
    bCedee : boolean;
begin
  bCedee := (Q.FindField('I_QUANTITE').AsFloat=0);
  DupliqueImmo.Visible:= not Q.IsEmpty;
  bAcquisitionAnnee := (Q.FindField('I_DATEPIECEA').AsDateTime >= VHImmo^.Encours.Deb)
                    and (Q.FindField('I_DATEPIECEA').AsDateTime <= VHImmo^.Encours.Fin);
  Cession.Visible := false;
  Eclatement.Visible := false;
  Mutation.Visible :=  false;
  DureeMethode.Visible := false;
  ElementExceptionnel.Visible := false;
  LeveeOption.Visible := false;
  AnnulationOperation.Visible := false;
  ChangeEtablissement.Visible := false;
  ChangeLieu.Visible := false;
  ModifBases.Visible := false;
  Ecritures1.Visible := not Q.IsEmpty;
  if Q.IsEmpty then exit;
  Nat := Q.FindField('I_NATUREIMMO').AsString;
  // Autorisation Cession
  Cession.Visible := ((Nat = 'PRO') or (Nat = 'FI')) and (not bCedee);
  // Autorisation Eclatement
  //ATTENTION : le champ immolie est utilisé pour identifier l'immo d'origine pour un depot de gar.
  bDepotGar := (Nat = 'FI') and (Q.FindField('I_IMMOLIE').AsString <> '');
  Eclatement.Visible := (Nat <> 'CB') and (Nat <> 'LOC') and (not bDepotGar) and (not bCedee) and (not (Q.FindField('I_OPELEVEEOPTION').AsString = 'X'));
  // Autorisation Mutation
  Mutation.Visible :=  (not bCedee) and (not (Q.FindField('I_OPELEVEEOPTION').AsString='X')) and (not bAcquisitionAnnee);
  // Autorisation Changement de plan
  DureeMethode.Visible := (Nat = 'PRO') and (not bAcquisitionAnnee) and (not bCedee) and (not (Q.FindField('I_OPECHANGEPLAN').AsString='X'));
  ElementExceptionnel.Visible := (Nat = 'PRO') and (not bCedee) and (not (Q.FindField('I_OPECHANGEPLAN').AsString='X'));
  // Autorisation Levée option
  LeveeOption.Visible := (Nat = 'CB') and (Q.FindField('I_IMMOLIE').AsString = '') and (not bCedee) and (not (Q.FindField('I_OPELEVEEOPTION').AsString='X'));
  // Opérations ?
  AnnulationOperation.Visible := Q.FindField('I_OPERATION').AsString = 'X';
  SeparAnnulAction.Visible := AnnulationOperation.Visible;
  ChangeEtablissement.Visible := not bCedee and (not (Q.FindField('I_OPELEVEEOPTION').AsString='X')) and (VHImmo^.EtablisCpta=TRUE);
  ChangeLieu.Visible := not bCedee and (not (Q.FindField('I_OPELEVEEOPTION').AsString='X'));
  ModifBases.Visible := (not bCedee) and (not bAcquisitionAnnee) and (not (Q.FindField('I_OPELEVEEOPTION').AsString='X'));
end;

procedure TFMulImmo.EnableZoomVisu;
var QLog:TQuery;
    Nature : string;
begin
  inherited;
  ZoomPlan.Visible := false;
  ZoomEcheance.Visible := false;
  ZoomOperations.Visible := false;
  if Q.IsEmpty then exit;
  ZoomPlan.Visible:=ImmoAmortie(Q.FindField('I_METHODEECO').AsString,
                                Q.FindField('I_METHODEFISC').AsString,
                                Q.FindField('I_NATUREIMMO').AsString);
  Nature := Q.FindField('I_NATUREIMMO').AsString;
  ZoomEcheance.Visible := ((Nature='CB') or (Nature = 'LOC'));
//21/04/99
//  QLog := OpenSQL ('SELECT * from IMMOLOG WHERE IL_IMMO="'+Q.FindField('I_IMMO').AsString+'"',True);
  QLog := OpenSQL ('SELECT * from IMMOLOG WHERE IL_IMMO="'+
        Q.FindField('I_IMMO').AsString+'" AND IL_TYPEOP<>"ACQ" AND IL_TYPEOP<>"CLO" AND (IL_DATEOP>="'+
        USDateTime(VHImmo^.Encours.Deb)+'" AND IL_DATEOP<="'+
        USDateTime(VHImmo^.Encours.Fin)+'")' ,True);
//21/04/99
  ZoomOperations.Visible := not (QLog.Eof);
  SeparOpeVisu.Visible := ZoomOperations.Visible;
  Ferme (QLog);
  SeparPlanVisu.Visible := ZoomEcheance.Visible or ZoomPlan.Visible;
end;

procedure TFMulImmo.PopZoomActionPopup(Sender: TObject);
begin
  inherited;
  if (FListe.NbSelected<=1) and (not FListe.AllSelected) then
    EnableZoomAction
  else EnableZoomActionSerie;
end;

procedure TFMulImmo.ZoomPlanClick(Sender: TObject);
var QTmp:TQuery;
    PlanAmor:TPlanAmort;
begin
  inherited;
  QTmp:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+Q.FindField('I_IMMO').AsString+'"',true);
  if not QTmp.Eof then
  begin
    PlanAmor:=TPlanAmort.Create(true) ;// := CreePlan(true);
    try
      PlanAmor.Charge(QTmp);
      PlanAmor.Recupere(QTmp.FindField('I_IMMO').AsString,QTmp.FindField('I_PLANACTIF').AsString);
      AffichePlanAmortissement(PlanAmor);
    finally
      PlanAmor.free ; //Detruit;
    end ;
  end;
  Ferme (QTmp);
end;

procedure TFMulImmo.ZoomEcheanceClick(Sender: TObject);
var QTmp:TQuery;
    Contrat : TImContrat;
begin
  inherited;
  QTmp:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+Q.FindField('I_IMMO').AsString+'"',true);
  if not QTmp.Eof then
  begin
    Contrat := TImContrat.Create;
    Contrat.Charge (QTmp);
    Contrat.ChargeTableEcheance;
    ListeDesEcheances (Contrat.ListeEcheances,QTmp.FindField('I_RESIDUEL').AsFloat);
    Contrat.Free;
  end;
  Ferme (QTmp);
end;

procedure TFMulImmo.ZoomOperationsClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  OperationsEnCours (Q.FindField('I_IMMO').AsString,Q.FindField('I_LIBELLE').AsString,false,TypeAction);
  BChercheClick(Sender);  //EPZ 24/11/98
  Q.Locate('I_IMMO',stCode,[]);
end;

procedure TFMulImmo.PopZoomVisuPopup(Sender: TObject);
begin
  inherited;
  if (FListe.NbSelected<=1) and (not FListe.AllSelected) then
    EnableZoomVisu
  else EnableZoomVisuSerie;
end;

procedure TFMulImmo.HMTradBeforeTraduc(Sender: TObject);
var Okok: boolean ;
begin
  inherited ;
  {$IFDEF SERIE1}
  ImLibellesTableLibre(PzLibreS1,'TT_TABLELIBREIMMO','','I') ;
  Okok:=false ;
  {$ELSE}
  ImLibellesTableLibre(PzLibre,'TI_TABLE','I_TABLE','I') ;
  Okok:=true ;
  {$ENDIF}
  PzLibre.TabVisible:=Okok ;
  PzLibreS1.TabVisible:=not Okok ;
end;

procedure TFMulImmo.CreationClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode:='';
  if not q.eof then stCode:=Q.FindField ('I_IMMO').AsString ;
  FicheImmobilisation(Nil,'',taCreat, '') ;
  BChercheClick(Sender);  //EPZ 03/11/98
  EnableBoutonsAction(TypeAction) ;
  if stCode<>'' then Q.Locate('I_IMMO',stCode,[]);
end;

procedure TFMulImmo.CreationSerieClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode:='';
  if not q.eof then stCode:=Q.FindField ('I_IMMO').AsString ;
  FicheImmobilisationSerie (nil,'',taCreat,'');
  BChercheClick(Sender);  //EPZ 03/11/98
  RechargeLieuGeographique; //EPZ 15/01/99
  if stCode<>'' then Q.Locate('I_IMMO',stCode,[]);
  EnableBoutonsAction(TypeAction) ;
end;

procedure TFMulImmo.PleineProprieteClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode:='';
  if not q.eof then stCode:=Q.FindField ('I_IMMO').AsString ;
  FicheImmobilisation(nil,'',taCreat,'PRO');
  BChercheClick(Sender);  //EPZ 23/11/98
  RechargeLieuGeographique; //EPZ 15/01/99
  if stCode<>'' then Q.Locate('I_IMMO',stCode,[]);
  EnableBoutonsAction(TypeAction) ;
end;

procedure TFMulImmo.ImmoFinanciereClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode:='';
  if not q.eof then stCode:=Q.FindField ('I_IMMO').AsString ;
  FicheImmobilisation(nil,'',taCreat,'FI');
  BChercheClick(Sender);  //EPZ 23/11/98
  RechargeLieuGeographique; //EPZ 15/01/99
  if stCode<>'' then Q.Locate('I_IMMO',stCode,[]);
  EnableBoutonsAction(TypeAction) ;
end;

procedure TFMulImmo.CreditBailClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode:='';
  if not q.eof then stCode:=Q.FindField ('I_IMMO').AsString ;
  FicheImmobilisation(nil,'',taCreat,'CB');
  BChercheClick(Sender);  //EPZ 23/11/98
  RechargeLieuGeographique; //EPZ 15/01/99
  if stCode<>'' then Q.Locate('I_IMMO',stCode,[]);
  EnableBoutonsAction(TypeAction) ;
end;

procedure TFMulImmo.LocationFinanciereClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode:='';
  if not q.eof then stCode:=Q.FindField ('I_IMMO').AsString ;
  FicheImmobilisation(nil,'',taCreat,'LOC');
  BChercheClick(Sender);  //EPZ 23/11/98
  RechargeLieuGeographique; //EPZ 15/01/99
  if stCode<>'' then Q.Locate('I_IMMO',stCode,[]);
  EnableBoutonsAction(TypeAction) ;
end;

procedure TFMulImmo.bSupprimeClick(Sender: TObject);
var i : integer;
    stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if (FListe.NbSelected=0) and (not FListe.AllSelected) then
  begin
    MessageAlerte(HM.Mess[9]);
    exit;
  end;
  if Q.FindField('I_ETAT').AsString<>'OUV' then
  begin
    if HM.execute(15,Caption,'')<>mrYes then exit ;
  end;
  if HM.execute(8,Caption,'')<>mrYes then exit ;
  inherited;
  if FListe.AllSelected then
  BEGIN
    InitMove(Q.RecordCount,'');
    Q.First;
    while Not Q.EOF do
    BEGIN
      MoveCur(False);
      FCodeImmo:=Q.FindField('I_IMMO').AsString ;
      if Transactions(SuppressionFicheImmo, 3)<>oeOk then
      BEGIN
        MessageAlerte(HM.Mess[0]) ;
        Break ;
      END else ImMarquerPublifi (True); // CA le 01/02/2001
      Q.Next;
    END;
    FListe.AllSelected:=False;
  END
  ELSE
  BEGIN
    InitMove(FListe.NbSelected,'');
    for i:=0 to FListe.NbSelected-1 do
    BEGIN
      MoveCur(False);
      FListe.GotoLeBookmark(i);
      FCodeImmo:=Q.FindField('I_IMMO').AsString ;
      if Transactions(SuppressionFicheImmo, 3)<>oeOk then
      BEGIN
        MessageAlerte(HM.Mess[0]) ;
        Break ;
      END ;
    END;
    FListe.ClearSelected;
  END;
  FiniMove;
  BChercheClick(Sender);  //EPZ 03/11/98
  Q.Locate('I_IMMO',stCode,[]);  
end;

procedure TFMulImmo.SuppressionFicheImmo;
begin
  if not IsOpeEnCours (nil,FCodeImmo,False) then // CA le 01/02/1999
  begin
    if  Q.FindField('I_CHANGECODE').AsString = '' then
      SupprimeFicheImmo(FCodeImmo)
    else  HM.Execute (12,Caption,'');
  end
  else HM.Execute(11,Caption,'');
end;

procedure TFMulImmo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
if isInside(Self) then Action:=caFree ;
end;

procedure TFMulImmo.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(FListe);
end;

procedure TFMulImmo.DessineListeImmo(ACol,ARow : LongInt; Canvas : TCanvas);
begin
//  if (Q.FindField('I_QUANTITE').AsInteger=0) then
  begin
     FListe.Canvas.Brush.Style:=bsBDiagonal;
     FListe.Canvas.Brush.Color:=clRed;
  end;
end;


function VerifNature (NatureIni, Nature : string) : boolean;
begin
  if ((NatureIni = 'LOC') or (NatureIni = 'CB')) and ((Nature='LOC') or (Nature='CB'))
  then Result := True else
  if ((NatureIni = 'FI') or (NatureIni = 'PRO')) and ((Nature='PRO') or (Nature='FI'))
  then Result := True else
  Result := False;
end;

(*function TFMulImmo.RecupereListeImmo : TStrings;
var i : integer;
    StringList : TStrings;
    NatureIni : string;
    bErreur : boolean;
begin
  bErreur := false;
  if FListe.AllSelected then
  begin
    InitMove(100,'');
    Q.First;
    NatureIni := Q.FindField ('I_NATUREIMMO').AsString;
    StringList := TStringList.Create;
    while not Q.EOF do
    begin
      MoveCur(False);
      if not VerifNature (NatureIni,Q.FindField ('I_NATUREIMMO').AsString) then
      begin
        HM.Execute(10,Caption,'');
        bErreur := True;
        break;
      end;
      if (Q.FindField ('I_TIERSA').AsString ='') then
      begin
        HM.Execute(14,Caption,'');
        bErreur := True;
        break;
      end;
      StringList.Add (Q.FindField('I_IMMO').AsString );
      Q.Next;
    end;
    FListe.AllSelected:=False;
  end else
  begin
    InitMove(FListe.NbSelected,'');
    StringList := TStringList.Create;
    NatureIni := Q.FindField ('I_NATUREIMMO').AsString;
    for i:=0 to FListe.NbSelected-1 do
    begin
      MoveCur(False);
      if not VerifNature (NatureIni,Q.FindField ('I_NATUREIMMO').AsString) then
      begin
        HM.Execute(10,Caption,'');
        bErreur := True;
        break;
      end;
      if (Q.FindField ('I_TIERSA').AsString ='') then
      begin
        HM.Execute(14,Caption,'');
        bErreur := True;
        break;
      end;
      FListe.GotoLeBookmark(i);
      StringList.Add (Q.FindField('I_IMMO').AsString );
    end;
    FListe.ClearSelected;
  end;
  FiniMove;
  if bErreur then  StringList.Clear;
  result := StringList;
end;         *)

procedure TFMulImmo.PopupListePopup(Sender: TObject);
begin
  inherited;
  PopZoomVisuPopup(Sender);
  PopZoomActionPopup(Sender);
  DupliqueImmoListe.Visible := (not Q.IsEmpty) and (FListe.NbSelected <=1);
  Cession1.Visible := Cession.Visible;
  Eclatement1.Visible := Eclatement.Visible;
  Mutation1.Visible := Mutation.Visible;
//02/02/99
  //ChangePlan1.Visible := ChangePlan.Visible;
  DureeMethode1.Visible := DureeMethode.Visible;
  ElementExceptionnel1.Visible := ElementExceptionnel.Visible;
//02/02/99
  LeveeOption1.Visible := LeveeOption.Visible;
  AnnulationOperation1.Visible := AnnulationOperation.Visible;
  SeparAnnul.Visible := AnnulationOperation1.Visible;
  ZoomPlan1.Visible := ZoomPlan.Visible;
  ZoomEcheance1.Visible := ZoomEcheance.Visible;
  SeparPlan.Visible := ZoomPlan1.Visible or ZoomEcheance1.Visible ;
  ZoomOperations1.Visible := ZoomOperations.Visible;
  SeparOpe.visible := ZoomOperations1.Visible;
  ChangeEtablissement1.Visible := ChangeEtablissement.Visible;
  ChangeLieu1.Visible := ChangeLieu.Visible;
  ModifBases1.Visible := ModifBases.Visible;
end;

procedure TFMulImmo.ChangeEtablissementClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if (FListe.NbSelected<=1) and (not FListe.AllSelected) then
    ExecuteEtablissement(Q.FindField('I_IMMO').AsString)
  else ExecuteOperationSerie(toChanEtabl);
  BChercheClick(Sender);
  Q.Locate('I_IMMO',stCode,[]);
end;

procedure TFMulImmo.ChangeLieuClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if (FListe.NbSelected<=1) and (not FListe.AllSelected) then
    ExecuteLieu(Q.FindField('I_IMMO').AsString)
  else ExecuteOperationSerie(toChanLieu);
  BChercheClick(Sender);
  Q.Locate('I_IMMO',stCode,[]);
end;

procedure TFMulImmo.EnableZoomVisuSerie;
begin
  ZoomPlan.Visible := false;
  ZoomEcheance.Visible := false;
  SeparPlanVisu.Visible := false;
  ZoomOperations.Visible := false;
end;

procedure TFMulImmo.EnableZoomActionSerie;
var i : integer;
    EnabledCession : boolean;
    PremNature : string;
    bCedee, bAcquisitionAnnee : boolean;
begin
  bAcquisitionAnnee := (Q.FindField('I_DATEPIECEA').AsDateTime >= VHImmo^.Encours.Deb)
                    and (Q.FindField('I_DATEPIECEA').AsDateTime <= VHImmo^.Encours.Fin);
  bCedee := (Q.FindField('I_QUANTITE').AsFloat = 0);
  DupliqueImmo.Visible:= false;
  Cession.Visible := true;
  Mutation.Visible :=  true;
  Eclatement.Visible := false;
//02/02/99
  //ChangePlan.Visible := false;
  DureeMethode.Visible := false;
  ElementExceptionnel.Visible := false;
//02/02/99
  LeveeOption.Visible := false;
  AnnulationOperation.Visible := false;
  SeparAnnulAction.Visible := false;
  ChangeEtablissement.Visible := true;
  ChangeLieu.Visible := true;
  ModifBases.Visible := false;
  Ecritures.Visible := false;
  //InitMove(FListe.NbSelected,'');
  if FListe.NbSelected > 0 then
  begin
      FListe.GotoLeBookmark(0);
      PremNature :=  Q.FindField('I_NATUREIMMO').AsString;
  end;
  for i:=0 to FListe.NbSelected-1 do
  begin
      //MoveCur(False);
      FListe.GotoLeBookmark(i);
      // Autorisation Cession
      EnabledCession := ((Q.FindField('I_NATUREIMMO').AsString = 'PRO') or (Q.FindField('I_NATUREIMMO').AsString = 'FI')) and (Q.FindField('I_QUANTITE').AsFloat <> 0);
      Cession.Visible := (Cession.Visible) and (EnabledCession) and (not bCedee);
      if (bCedee) or bAcquisitionAnnee or ((Mutation.Visible) and (Q.FindField('I_NATUREIMMO').AsString <> PremNature)) then
        Mutation.Visible := false;
      ChangeEtablissement.Visible := ChangeEtablissement.Visible and (not bCedee) and (VHImmo^.EtablisCpta=TRUE);
      ChangeLieu.Visible := ChangeLieu.Visible and (not bCedee);
  end;
  //FiniMove;
end;

procedure TFMulImmo.ExecuteOperationSerie(TypeOpe : TypeOperation);
var FOpeSerie: TOperationSerie;
    i,OrdreSerie,ret : integer;
    QTmp, QSortie : TQuery;
    FCodeImmo : string;
begin
  FOpeSerie := TOperationSerie.Create(Application);
  try
    if TypeOpe=toMutation then FOpeSerie.fNature := Q.FindField('I_NATUREIMMO').AsString;
    FOpeSerie.InitOperationSerie(TypeOpe,FListe.NbSelected);
    ret := FOpeSerie.AfficheOperationSerie;
    if (ret = mrOK) or (ret = mrYes) then
    begin
//21/01/99
      if TypeOpe=toCession then FOpeSerie.CalculCumulImmoCedee(FListe,Q);
//21/01/99
      QTmp:=OpenSQL('SELECT MAX(IL_ORDRESERIE) FROM IMMOLOG',TRUE) ;
      if not QTmp.EOF then OrdreSerie:=(QTmp.Fields[0].AsInteger+1)
      else OrdreSerie := 1;
      Ferme(QTmp);
      if FListe.AllSelected then
      BEGIN
          InitMove(100,'');
          Q.First;
          while Not Q.EOF do
          BEGIN
              MoveCur(False);
              FCodeImmo:=Q.FindField('I_IMMO').AsString ;
              case TypeOpe of
                toCession:begin
                   QSortie := OpenSQL ('SELECT I_QUANTITE,I_MONTANTHT,I_BASEECO,I_BASEFISC FROM IMMO WHERE I_IMMO="'+FCodeImmo+'"',True);
                   FOpeSerie.OperationSerieCession(FCodeImmo,OrdreSerie,QSortie);
                   Ferme (QSortie);
                   end;
                toMutation: begin
                   FOpeSerie.OperationSerieMutation(FCodeImmo,OrdreSerie);
                   end;
                toChanEtabl: begin
                   FOpeSerie.OperationSerieEtabl(FCodeImmo,OrdreSerie);
                   end;
                toChanLieu: begin
                   FOpeSerie.OperationSerieLieu(FCodeImmo,OrdreSerie);
                   end;
              end ;
              Q.Next;
          END;
          FListe.AllSelected:=False;
      END
      ELSE
      BEGIN
        InitMove(FListe.NbSelected,'');
        for i:=0 to FListe.NbSelected-1 do
        begin
              MoveCur(False);
              FListe.GotoLeBookmark(i);
              FCodeImmo:=Q.FindField('I_IMMO').AsString ;
              case TypeOpe of
                toCession:begin
                   QSortie := OpenSQL ('SELECT I_QUANTITE,I_MONTANTHT,I_BASEECO,I_BASEFISC FROM IMMO WHERE I_IMMO="'+FCodeImmo+'"',True);
                   FOpeSerie.OperationSerieCession(FCodeImmo,OrdreSerie,QSortie);
                   Ferme (QSortie);
                   end;
                toMutation: begin
                   FOpeSerie.OperationSerieMutation(FCodeImmo,OrdreSerie);
                   end;
                toChanEtabl: begin
                   FOpeSerie.OperationSerieEtabl(FCodeImmo,OrdreSerie);
                   end;
                toChanLieu: begin
                   FOpeSerie.OperationSerieLieu(FCodeImmo,OrdreSerie);
                   end;
              end ;
        end;
        FListe.ClearSelected;
      END;
      FiniMove;
    end;
  finally
    bSelectAll.Down := FListe.AllSelected ;
    FOpeSerie.DetruitOperationSerie;
  end;
end;

procedure TFMulImmo.SQDataChange(Sender: TObject; Field: TField);
var
  sListeOpe : string;
begin
  inherited;
  sListeOpe := '';
  if Q.FindField('I_OPEMUTATION').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[1] + ', ';
  if Q.FindField('I_OPEECLATEMENT').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[2] + ', ';
  if Q.FindField('I_OPECESSION').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[3] + ', ';
  if Q.FindField('I_OPEMODIFBASES').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[4] + ', ';
  if Q.FindField('I_OPELIEUGEO').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[5] + ', ';
  if Q.FindField('I_OPEETABLISSEMENT').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[6] + ', ';
  if Q.FindField('I_OPECHANGEPLAN').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[7] + ', ';
  if Q.FindField('I_OPELEVEEOPTION').AsString = 'X' then
    sListeOpe := sListeOpe + HM2.Mess[8] + ', ';
  if Length (sListeOpe) > 0 then
  begin
    Delete (sListeOpe,Length(sListeOpe)-1,2);
    sOperationEnCours.Caption := HM2.Mess[0] + sListeOpe;
  end else sOperationEnCours.Caption := '';
end;

procedure TFMulImmo.RechargeLieuGeographique; //EPZ 15/01/99
var OldVal : string;
begin
  inherited;
  OldVal := I_LIEUGEO.Value;
  I_LIEUGEO.Reload;
  I_LIEUGEO.Value := OldVal;
end;


procedure TFMulImmo.FExercice2Change(Sender: TObject);
begin
  inherited;
  if FExercice2.Value<>'' then
    begin
    ImExoToDates(FExercice2.Value,I_DATEPIECEA,I_DATEPIECEA_) ;
    ImExoToDates(FExercice2.Value,I_DATEAMORT,I_DATEAMORT_) ;
    end
  else
    begin
    I_DATEAMORT.text:=StDate1900 ; I_DATEAMORT_.text:=StDate2099 ;
    I_DATEPIECEA.text:=StDate1900 ; I_DATEPIECEA_.text:=StDate2099 ;
    end;
end;

procedure TFMulImmo.ToutesOpeClick(Sender: TObject);
begin
  inherited;
  ToutesOpe.Checked := true;
  AucuneOpe.Checked := false;
  SelectOpe.Checked := false;
  FiltreOpe.State := cbChecked;
  GereFiltreOpeEnCours(cbChecked);
end;

procedure TFMulImmo.SelectOpeClick(Sender: TObject);
begin
  inherited;
  ToutesOpe.Checked := false;
  AucuneOpe.Checked := false;
  SelectOpe.Checked := true;
  FiltreOpe.State := cbUnChecked;
  XX_Where_Ope := XX_Where_OpeAucune;
  GereFiltreOpeEnCours(cbUnchecked);
end;

procedure TFMulImmo.AucuneOpeClick(Sender: TObject);
begin
  inherited;
  ToutesOpe.Checked := false;
  AucuneOpe.Checked := true;
  SelectOpe.Checked := false;
  FiltreOpe.State := cbGrayed;
  GereFiltreOpeEnCours(cbUnchecked);
end;

procedure TFMulImmo.GereFiltreOpeEnCours(Etat : TCheckBoxState);
begin
  XX_Where_Ope := '';
  EnableCheckBoxOperation;
  OPEECLATEMENT.State := Etat;
  OPECHANGEPLAN.State := Etat;
  OPELIEUGEO.State := Etat;
  OPEETABLISSEMENT.State := Etat;
  OPELEVEEOPTION.State := Etat;
  OPECESSION.State := Etat;
  OPEMUTATION.State := Etat;
  REINTEGRATION.State := Etat;
  QUOTEPART.State := Etat;
end;

procedure TFMulImmo.PopZoomCreationPopup(Sender: TObject);
begin
  inherited;
  DupliqueImmo.Visible:= (not Q.EOF) and (FListe.NbSelected <= 1);
end;

procedure TFMulImmo.DupliqueImmoClick(Sender: TObject);
var
  NewCodeImmo : string;
  ret : integer;
  stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if not IsOpeEnCours (nil,Q.FindField('I_IMMO').AsString,false) then    // CA le 22/01/1999
  begin
    ret := ExecuteDuplication(Q.FindField('I_IMMO').AsString,NewCodeImmo);
    if ret = mrYes then FicheImmobilisation(nil,NewCodeImmo,taModif,'');
    BChercheClick(Sender);
    Q.Locate('I_IMMO',stCode,[]);
  end
  else HM.Execute(13,Caption,'');
end;

procedure TFMulImmo.OpeIndicateurClick(Sender: TObject);
var sOR, sSEP : string;
begin
  inherited;
  sSEP := ' AND ('; sOR := ' OR';
  XX_Where_Ope := '';
  if OPEECLATEMENT.Checked then
  begin
    XX_Where_Ope:=XX_Where_Ope+sSep+' I_OPEECLATEMENT="X"';
    sSep:=sOR;
  end;
  if OPECHANGEPLAN.Checked then
  begin
    XX_Where_Ope:=XX_Where_Ope+sSep+' I_OPECHANGEPLAN="X"';
    sSep := sOR;
  end;
  if OPELIEUGEO.Checked then
  begin
    XX_Where_Ope:=XX_Where_Ope+sSep+' I_OPELIEUGEO="X"';
    sSep := sOR;
  end;
  if OPEETABLISSEMENT.Checked then
  begin
    XX_Where_Ope:=XX_Where_Ope+sSep+' I_OPEETABLISSEMENT="X"';
    sSep := sOR;
  end;
  if OPELEVEEOPTION.Checked then
  begin
    XX_Where_Ope:=XX_Where_Ope+sSep+' I_OPELEVEEOPTION="X"';
    sSep := sOR;
  end;
  if OPECESSION.Checked then
  begin
    XX_Where_Ope := XX_Where_Ope+sSep+' I_OPECESSION="X"';
    sSep := sOR;
  end;
  if OPEMUTATION.Checked then
  begin
    XX_Where_Ope := XX_Where_Ope+sSep+' I_OPEMUTATION="X"';
    sSep := sOR;
  end;
  if REINTEGRATION.Checked then
  begin
    XX_Where_Ope:=XX_Where_Ope +sSep+' I_REINTEGRATION<>0.0';
    sSep := sOR;
  end;
  if QUOTEPART.Checked then
  begin
    XX_Where_Ope := XX_Where_Ope+sSep+' I_QUOTEPART<>0';
    sSep := sOR;
  end;
  if XX_Where_Ope<>'' then
    XX_Where_Ope := XX_Where_Ope + ')'
  else
    XX_Where_Ope := XX_Where_OpeAucune;
end;

procedure TFMulImmo.BChercheClick(Sender: TObject);
begin
  XX_WHERE.Text:=XX_WHERE.Text+WhereDerog;
  if AucuneOpe.Checked then XX_WHERE.Text:=XX_WHERE.Text+XX_Where_OpeAucune
  else if ToutesOpe.Checked then XX_WHERE.Text:=XX_WHERE.Text
  else XX_WHERE.Text:=XX_WHERE.Text+XX_Where_Ope;
  inherited;
  XX_WHERE.Text := XX_Where_Orig;
end;

procedure TFMulImmo.BNouvRechClick(Sender: TObject);
begin
  inherited;
  XX_WHERE.Text := XX_Where_Orig ;
  ToutesOpe.Checked := true;
  AucuneOpe.Checked := false;
  SelectOpe.Checked := false;
  GereFiltreOpeEnCours(cbChecked);
end;

procedure TFMulImmo.ModifBasesClick(Sender: TObject);
var stCode : string;
begin
  inherited;
  stCode := Q.FindField ('I_IMMO').AsString;
  if ExecuteModificationBases(Q.FindField('I_IMMO').AsString) = mrYes then
  begin
    bChercheClick (Sender);
    Q.Locate('I_IMMO',stCode,[]);
  end;
end;

procedure TFMulImmo.FFiltresChange(Sender: TObject);
begin
  inherited;
  ToutesOpe.Checked := FiltreOpe.State = cbChecked;
  AucuneOpe.Checked := FiltreOpe.State = cbGrayed;
  SelectOpe.Checked := FiltreOpe.State = cbUnchecked;
  EnableCheckBoxOperation;
end;

procedure TFMulImmo.BSaveFiltreClick(Sender: TObject);
begin
  // CA le 12/02/1999 - la sauvegarde n'est effectué
  // que pour les critères visibles !!!
  FiltreOpe.Color := clBtnFace;
  FiltreOpe.Visible := True;
  inherited;
  FiltreOpe.Visible := False;
end;

procedure TFMulImmo.EnableCheckBoxOperation;
begin
  OPEECLATEMENT.Enabled := SelectOpe.Checked;
  OPECHANGEPLAN.Enabled := SelectOpe.Checked;
  OPELIEUGEO.Enabled := SelectOpe.Checked;
  OPEETABLISSEMENT.Enabled := SelectOpe.Checked;
  OPELEVEEOPTION.Enabled := SelectOpe.Checked;
  OPECESSION.Enabled := SelectOpe.Checked;
  OPEMUTATION.Enabled := SelectOpe.Checked;
  REINTEGRATION.Enabled := SelectOpe.Checked;
  QUOTEPART.Enabled := SelectOpe.Checked;
end;

procedure TFMulImmo.AMORTDEROGClick(Sender: TObject);
begin
  inherited;
  if AMORTDEROG.State = cbGrayed then WhereDerog := ''
  else if AMORTDEROG.State = cbChecked then WhereDerog := 'AND I_METHODEFISC <>""'
  else WhereDerog := 'AND I_METHODEFISC =""';
end;

procedure TFMulImmo.OnCompteElipsisClick(Sender: TObject);
var stWhere : string;
begin
  inherited;
  if THDBEdit(Sender).Name = 'I_COMPTEIMMO'
    then stWhere := 'G_GENERAL LIKE "2%"'
  else  if THDBEdit(Sender).Name = 'I_COMPTELIE'
    then stWhere := 'G_GENERAL LIKE "612%" OR G_GENERAL LIKE "613%"';
  LookupList(TControl(Sender),'','GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,0)  ;
end;

procedure TFMulImmo.EcrituresClick(Sender: TObject);
var Nature : string;
    ListeImmo : TStrings;
begin
  inherited;
  ListeImmo := TStringList.Create;
  Nature := Q.FindField ('I_NATUREIMMO').AsString;
  ListeImmo.Add (Q.FindField ('I_IMMO').AsString);
  if (Nature = 'PRO') or (Nature = 'FI') then
      IntegrationEcritures (toDotation,ListeImmo,FALSE,TRUE)
  else if (Nature = 'CB') or (Nature = 'LOC') then
      IntegrationEcritures (toEcheance,ListeImmo,FALSE,TRUE);
  ListeImmo.Free;
end;

procedure TFMulImmo.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  Case Key of
    VK_F3 : BEGIN Key:=0 ; if Shift = [ssAlt] then LanceCreationEnSerie ; BChercheClick(nil);END ;
  END ;
end;

procedure TFMulImmo.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
Case Key of
  VK_F11: BEGIN Key:=0 ; if TypeAction<>taConsult then PopupListe.Popup(Mouse.CursorPos.x,Mouse.CursorPos.y) ; END ;
  VK_F5 : BEGIN Key:=0 ; FListeDblClick(nil); END ;
  END ;
end;

procedure TFMulImmo.EnableBoutonsAction ( Action : TActionFiche);
begin
  BZoomAction.Visible := (Action <> taConsult) and (not Q.IsEmpty);
  BInsert.Visible := (Action <> taConsult);
  {$IFDEF SERIE1}
  BDelete.Visible:=(Action<>taConsult) and (not Q.IsEmpty);
  {$ELSE}
  BDelete.Visible:=(Action<>taConsult) and (V_PGI.LaSerie=S5) and (not Q.IsEmpty);
  {$ENDIF}
end;

procedure TFMulImmo.TABLELIBRE1Change(Sender: TObject);
begin
  inherited;
  if Sender=TABLELIBRE1 then I_TABLE0.Text:=TABLELIBRE1.Value
  else if Sender=TABLELIBRE2 then I_TABLE1.Text:=TABLELIBRE2.Value
  else if Sender=TABLELIBRE3 then I_TABLE2.Text:=TABLELIBRE3.Value
end;

procedure TFMulImmo.I_ORGANISMECBElipsisClick(Sender: TObject);
begin
  inherited;
  {$IFDEF SERIE1}
  LookUpList (TControl(Sender),TraduireMemoire('Auxiliaire'),'TIERS','T_AUXILIAIRE','T_LIBELLE','T_NATUREAUXI="FOU"','T_AUXILIAIRE',True,1) ;
  {$ELSE}
  LookUpList (TControl(Sender),TraduireMemoire('Auxiliaire'),'TIERS','T_AUXILIAIRE','T_LIBELLE','T_NATUREAUXI="FOU"','T_AUXILIAIRE',True,2) ;
  {$ENDIF}
end;

procedure TFMulImmo.FormCreate(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
HelpContext:=511000 ;
if (VS1^.TypeProduit=S1BNC) and (not VS1^.OKModImmo) then
  begin
  ImmoFinanciere.Visible:=false ;
  CreditBail.Visible:=false ;
  LocationFinanciere.Visible:=false ;
  CreationSerie.Visible:=false ;
  bZoomVisu.Visible:=false ;
  end ;
{$ELSE}
if TypeAction=taConsult then HelpContext := 2110000 else HelpContext:= 2111000;
{$ENDIF}
end;

procedure TFMulImmo.BREINITClick(Sender: TObject);
var stCode : string;
    i : integer;
begin
  stCode := Q.FindField('I_IMMO').AsString;
  if (FListe.AllSelected) or (FListe.NbSelected>=1) then
  begin
    if (PGIAsk('Attention : cette opération va réinitialiser les fiches sélectionnées.#10#13Voulez-vous continuer ?.','') = mrYes) then
    begin
      if FListe.AllSelected then
      begin
        Q.First;
        while not Q.Eof do
        begin
{$IFDEF AMORTISSEMENT}
          ReinitImmo ( Q.FindField('I_IMMO').AsString );
          Q.Next;
{$ENDIF}
        end;
      end else
      begin
        for i:=0 to FListe.NbSelected-1 do
        begin
          FListe.GotoLeBookmark(i);
{$IFDEF AMORTISSEMENT}
          ReinitImmo ( Q.FindField('I_IMMO').AsString );
{$ENDIF}
        end;
      end;
      BChercheClick(Sender);
      Q.Locate('I_IMMO',stCode,[]);
      FListe.ClearSelected;
      PGIInfo('Traitement terminé.');
    end;
  end;
end;

end.

