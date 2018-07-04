unit OpeSerie;

// --- Liste des modifications et ajouts ---
// 06/05/1999 - CA - Prise en compte de I_COMPTEREF
// 18/05/1999 - CA - Mutation : cas différents suivant nature (ajout public fNature)
// 24/06/1999 - CA - Cession : pb à la création du plan de l'immo cédée
// 26/05/2004 - CA - FQ 13235 : On ne propose pas de date par défaut.
// 29/07/2005 - MB - FQ 15019 : repositionnement fenetre regroupement apres bdetail et binfo
// 01/06     - BTY - FQ 17378 Message bloquant en Informations du regroupement s'il n'a pas de modèle
// 04/06     - BTY - FQ 17516 Le changement de regroupement devient une opération ('REG')
// 12/06/2006 - YCP - Pour serie 1 mettre la date d'entree par defaut
// 09/06     - BTY - FQ 18775 Changement de regroupement : l'icône d'accès à la saisie des regroupements
//                   doit appeler le nouveau pgm de saisie avec notion de composant
// BTY - 11/06 - FQ 19099 Mutation en série, compte non paddé et divers modifs
// BTY - 11/06 - FQ 13234 Mutation d'un compte corporel à un compte incorporel, mettre basetaxepro à 0
// (fait pour la mutation unitaire en 2005)

// REMARQUE MBO 31.10.07 - la cession en série n'est pas utilisée actuellement ds l'appli

// --- Fin Liste ---
interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  Mask,
  Hctrls,
  ComCtrls,
  HTB97,
  ExtCtrls,
  HPanel,
  ImEnt,
  Hent1,
  Outils,
  HRichEdt,
  HRichOLE,
  hmsgbox,
  Db,
  uTob,
  {$IFDEF EAGLCLIENT}
  eTablette,
  {$ELSE}
  Tablette,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  AMSREGRO_TOF,    // FQ 18775
  HDB,
  Hqry,
  HStatus,
  ParamDat,
  LookUp,
  HSysMenu(*, UtilSoc*)
  {$IFNDEF SERIE1}
  ,Ent1
  {$ENDIF}
  ;

type

  RCession = record
    DateOpe : TDateTime;
    Motif : string;
    CalculDot : string;
    PrixCession : double;
    TvaReverser : double;
    BlocNote : HTStrings;
    CumulImmoSel : double;
    PrixParImmo : boolean;
    RepartitionImmo : boolean;
  end;
  RMutation = record
    DateOpe : TDateTime;
    Compte : string;
    BlocNote : HTStrings;
  end;
  RLieu = record
    DateOpe : TDateTime;
    Lieu : string;
    BlocNote : HTStrings;
  end;
  // BTY 04/06 FQ 17516
  RGroupe = record
    DateOpe : TDateTime;
    Regroupement : string;
    BlocNote : string;
  end;

  REtablissement = record
    DateOpe : TDateTime;
    Etablissement : string;
    BlocNote : HTStrings;
  end;
  TOperationSerie = class(TForm)
    PDATEOPERATION: THPanel;
    HPanel2: THPanel;
    HPanel3: THPanel;
    HLabel3: THLabel;
    DateOperation: THCritMaskEdit;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    Panel: TPageControl;
    tsCession: TTabSheet;
    HLabel1: THLabel;
    MotifCession: THValComboBox;
    CalculDot: THValComboBox;
    PrixCession: THNumEdit;
    HLabel6: THLabel;
    HLabel4: THLabel;
    tsMutation: TTabSheet;
    tsLieu: TTabSheet;
    tsEtablissement: TTabSheet;
    tI_COMPTELIE: THLabel;
    HLabel2: THLabel;
    HLabel5: THLabel;
    HLabel8: THLabel;
    tsBlocNote: TTabSheet;
    BlocNote: THRichEditOLE;
    HM: THMsgBox;
    RepartitionPrix: TRadioButton;
    PrixParImmo: TRadioButton;
    Etablissement: THValComboBox;
    LieuGeo: THValComboBox;
    CompteMutation: THCritMaskEdit;
    HMTrad: THSystemMenu;
    PREGROUPEMENT: TTabSheet;
    HLabel7: THLabel;
    FCODEREGROUPEMENT: THValComboBox;
    BINFOREG: TToolbarButton97;
    BDETAILREG: TToolbarButton97;
    TGroupement: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
//    procedure AccelCreationMotifClick(Sender: TObject);
//    procedure AccelCreationLieuClick(Sender: TObject);
    procedure CompteMutationExit(Sender: TObject);
    procedure DateOperationKeyPress(Sender: TObject; var Key: Char);
    procedure CompteMutationEnter(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure CompteMutationElipsisClick(Sender: TObject);
    procedure BINFOREGClick(Sender: TObject);
    procedure BDETAILREGClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure OnClickSaisieRegroupement(Sender: TObject); // BTY 09/06 FQ 18775
  private
    TypeOpe : TypeOperation;
    CodeI : string;
    OrdreSerie : integer;
    NbImmoSelect : integer;
    EnregCession : RCession;
    EnregMutation : RMutation;
    EnregEtablissement : REtablissement;
    EnregLieu : RLieu;
    // BTY 04/06 FQ 17516
    //fRegroupement : string;
    EnregGroupe : RGroupe;
    QteAchat : double;
    MontantHT : double;
    BaseEco : double;
    BaseFisc : double;
    EtablissementPrec : string;
    RegroupementPrec : string;
    LieuPrec : string;
    ValeurPrec : string;
    fCurCompte : string;
    fPlanActif : integer;
    { Déclarations privées }
    procedure AffecteOperationSerie;
    procedure EnregistreMutationSerie;
    procedure EnregistreCessionSerie;
    procedure EnregistreEtablissement;
    procedure EnregistreLieu;
    procedure EnregistreRegroupement;
    procedure TraiteEnregistrementOperation;
    function ControleCompteMutation : boolean;
    function DateOperationOK : boolean;
  public
    fNature : string;
    fDateDerniereOpe : TDateTime;     // FQ 19099
    fCode : string;     // FQ 19099
    function AfficheOperationSerie : integer;
    procedure InitOperationSerie(Operation : TypeOperation; NbImmo : integer);
    procedure DetruitOperationSerie;
    procedure OperationSerieMutation(vCodeImmo : string;NumOrdreLog : integer);
    procedure OperationSerieCession(vCodeImmo : string;NumOrdreLog : integer; QImmo : TQuery);
    procedure OperationSerieEtabl(vCodeImmo : string; NumOrdreLog : integer);
    procedure OperationSerieLieu(vCodeImmo : string; NumOrdreLog : integer);
    // BTY FQ 17516
    //procedure OperationSerieRegroupement(vCodeImmo : string );
    procedure OperationSerieRegroupement(vCodeImmo : string; NumOrdreLog : integer);
    procedure CalculCumulImmoCedee(Liste : THGrid ; Q : THQuery);
    { Déclarations publiques }
  end;

procedure ExecuteEtablissement(CodeImmo : string);
procedure ExecuteLieu(CodeImmo : string);
procedure ExecuteRegroupement ( CodeImmo : string );

implementation

uses  ImSortie,
      ImMutati,
      ImPlan,
      AMLISTE_TOF,
      AMREGROUPEMENT_TOF;

const
MSG_NOMODELE = 'Il n''existe pas de modèle pour ce regroupement.';
MSG_REG = 'Regroupement ';

{$R *.DFM}

procedure TOperationSerie.InitOperationSerie(Operation : TypeOperation; NbImmo : integer);
begin
  TypeOpe := Operation;
  NbImmoSelect := NbImmo;
  fDateDerniereOpe := iDate1900;
end;

function TOperationSerie.AfficheOperationSerie : integer;
begin
  result := ShowModal ;
end;

procedure TOperationSerie.DetruitOperationSerie;
begin
  Free;
end;

procedure TOperationSerie.FormShow(Sender: TObject);
begin
{$IFDEF SERIE1}
DateOperation.EditText:=DateToStr(V_PGI.DateEntree); //YCP 12/06/2006
//AccelCreationMotif.visible:=false ;
//AccelCreationLieu.visible:=false ;
{$ELSE}
{$ENDIF}
//  DateOperation.EditText:=DateToStr(Date);
// CA - 26/05/2004 - FQ 13235 : On ne propose pas de date par défaut.
//  DateOperation.EditText:=DateToStr(VHImmo^.Encours.Fin);
  tsCession.TabVisible := false;
  tsMutation.TabVisible := false;
  tsLieu.TabVisible := false;
  tsEtablissement.TabVisible := false;
  PREGROUPEMENT.TabVisible := False;
  BINFOREG.Visible := (TypeOpe = toRegroupement);
  BDETAILREG.Visible := (TypeOpe = toRegroupement);  
  case TypeOpe of
       toCession:  begin
                   Caption := HM.Mess[5];
                   Panel.ActivePage:=tsCession;
                   tsCession.TabVisible := true;
                   //FocusControl(MotifCession);//EPZ 12/03/99
                   MotifCession.ItemIndex:=0;
                   CalculDot.ItemIndex:=0;
                   end;
       toMutation: begin
                   Caption := HM.Mess[6];
                   Panel.ActivePage:=tsMutation;
                   tsMutation.TabVisible := true;
                   //FocusControl(CompteMutation);//EPZ 12/03/99
                   end;
       toChanEtabl: begin
                   Caption := HM.Mess[7];
                   Panel.ActivePage:=tsEtablissement;
                   tsEtablissement.TabVisible := true;
                   Etablissement.ItemIndex:=Etablissement.Items.IndexOf(EtablissementPrec) ;
                   Etablissement.Text:=EtablissementPrec;
                   HelpContext := 2111700;
                  {$IFNDEF SERIE1}
                  PositionneEtabUser(Etablissement);
                  {$ENDIF}

                   //FocusControl(Etablissement);//EPZ 12/03/99
                   end;
       toChanLieu: begin
                   Caption := HM.Mess[8];
                   Panel.ActivePage:=tsLieu;
                   tsLieu.TabVisible := true;
                   LieuGeo.ItemIndex:=LieuGeo.Items.IndexOf(LieuPrec) ;
                   LieuGeo.Text:=LieuPrec;
                   HelpContext := 2111600;
                   //FocusControl(LieuGeo);//EPZ 12/03/99
                   end;
       toRegroupement:
                    begin
                      Caption := HM.Mess[19];
                      Panel.ActivePage:=PREGROUPEMENT;
                      PREGROUPEMENT.TabVisible := true;
                      // BTY 04/06 FQ 17516 Date+Bloc-note accessibles, Regroupement alimenté
                      tsBlocNote.TabVisible := True; //False;
                      PDATEOPERATION.Visible := True; //False;
                      FCodeRegroupement.ItemIndex:=FCodeRegroupement.Items.IndexOf(RegroupementPrec) ;
                      FCodeRegroupement.Text:=RegroupementPrec;
                    end;
  end ;
  if PDateOperation.Visible then
     begin
     FocusControl(DateOperation);
     if fDateDerniereOpe > iDate1900 then
        DateOperation.EditText := DateToStr (fDAteDerniereOpe);
     end;
end;

procedure TOperationSerie.BValiderClick(Sender: TObject);
var Err: integer ;
begin
  Err:=-1 ;

  //BTY 04/06 FQ 17516 if (TypeOpe<>toRegroupement) then
  //begin
    if not IsValidDate(DateOperation.EditText) then Err := 14
    else if (StrToDate(DateOperation.EditText) < VHImmo^.EnCours.Deb)
         or (StrToDate(DateOperation.EditText) > VHImmo^.EnCours.Fin) then
             err:=15
    // FQ 19099 La date est-elle postérieure à la date de dernière opération ?
    else if ((fDateDerniereOpe>iDate1900) and (fDateDerniereOpe>StrToDate(DateOperation.EditText)))
         or ((fCode <> '') and (not DateOperationOK)) then
         err :=18;
  //end;
  //BTY 04/06 FQ 17516 Changement de regroupement
  if (TypeOpe = toRegroupement) then
  begin
    if (fCodeRegroupement.Text=RegroupementPrec) and (RegroupementPrec <> '') then Err:=20;
  end
  else if (TypeOpe=toMutation) then
  //
  //if (TypeOpe=toMutation) then
  begin
    if (CompteMutation.Text='') then Err:=2// HM.execute(2,'',''); Panel.ActivePage:=tsMutation ; FocusControl (CompteMutation);
    else if (not ControleCompteMutation) then Err:=11 ; //HM.execute(11,'',''); Panel.ActivePage:=tsMutation ; FocusControl (CompteMutation);
  end
  else if (TypeOpe=toCession) then
  begin
    if (MotifCession.Text='') then Err:=3 //HM.execute(3,'',''); Panel.ActivePage:=tsMutation ; FocusControl (MotifCession);
    else if (CalculDot.Text='') then Err:=4 ; //HM.execute(4,'',''); Panel.ActivePage:=tsMutation ; FocusControl (CalculDot);
  end
  else if (TypeOpe=toChanEtabl) then
  begin
    if (Etablissement.Text='') then Err:=9  //HM.execute(9,'','') ; Panel.ActivePage:=tsMutation ; FocusControl (Etablissement);
    else if (Etablissement.Text=EtablissementPrec) then Err:=12 ; //HM.execute(12,'',''); Panel.ActivePage:=tsMutation ; FocusControl(Etablissement);
  end
  else if (TypeOpe = toChanLieu) then
  begin
    if (LieuGeo.Text='') then Err:=10  //HM.execute(10,'',''); FocusControl (LieuGeo);
    else if (LieuGeo.Text=LieuPrec) then Err:=13 ; //HM.execute(13,'',''); FocusControl (LieuGeo); exit;
  end;

  if Err=-1 then
  begin
    ModalResult:=mrYes;
    AffecteOperationSerie;
  end
  else
  begin
    ModalResult := mrNone;
    case TypeOpe of
      toCession:  Panel.ActivePage:=tsCession;
      toMutation: Panel.ActivePage:=tsMutation;
      toChanEtabl:Panel.ActivePage:=tsEtablissement;
      toChanLieu: Panel.ActivePage:=tsLieu;
      // BTY  04/06 FQ 17516
      toRegroupement : Panel.ActivePage:=PREGROUPEMENT;
    end ;
    HM.execute(Err,'','');
  end ;
end;

procedure TOperationSerie.BFermeClick(Sender: TObject);
var mr : integer;
begin
  mr := HM.execute(0,Caption,'');
  if mr = mrYes then
  begin
    BValiderClick(BValider);
  end
  else if mr = mrNo then ModalResult := mrNo
  else ModalResult := mrNone;
end;

procedure TOperationSerie.AffecteOperationSerie;
begin
  case TypeOpe of
       toCession:  begin
                   EnregCession.DateOpe:=StrToDate(DateOperation.EditText);
                   EnregCession.Motif:=MotifCession.Value;
                   EnregCession.CalculDot:=CalculDot.Value;
                   EnregCession.PrixParImmo:=PrixParImmo.checked;
                   EnregCession.RepartitionImmo:= not PrixParImmo.checked;
                   EnregCession.PrixCession:=StrToFloat(PrixCession.Text);
//                   EnregCession.TvaReverser:=TvaReverser.Value;
                   EnregCession.BlocNote := BlocNote.LinesRTF;
                   end;
       toMutation: begin
                   EnregMutation.DateOpe:=StrToDate(DateOperation.EditText);
                   EnregMutation.Compte:=CompteMutation.Text;
                   EnregMutation.BlocNote := BlocNote.LinesRTF;
                   end;
       toChanEtabl: begin
                   EnregEtablissement.DateOpe:=StrToDate(DateOperation.EditText);
                   EnregEtablissement.Etablissement:=Etablissement.Value;
                   EnregEtablissement.BlocNote := BlocNote.LinesRTF;
                    end;
       toChanLieu: begin
                   EnregLieu.DateOpe:=StrToDate(DateOperation.EditText);
                   EnregLieu.Lieu:=LieuGeo.Value;
                   EnregLieu.BlocNote := BlocNote.LinesRTF;
                   end;
       toRegroupement : begin
                   // BTY FQ 17516
                   //       fRegroupement := FCODEREGROUPEMENT.Value;
                   EnregGroupe.DateOpe:=StrToDate(DateOperation.EditText);
                   EnregGroupe.Regroupement:= FCODEREGROUPEMENT.Value;
                   EnregGroupe.BlocNote := RichToString (Blocnote);
                        end;
  end ;
end;

(*procedure TOperationSerie.AccelCreationMotifClick(Sender: TObject);
var OldVal : string;
begin
  OldVal := MotifCession.Value;
  ParamTable('TIMOTIFCESSION',taCreat,0,nil) ;
  MotifCession.Reload;
  MotifCession.Value := OldVal;
end;
*)
procedure TOperationSerie.OperationSerieMutation(vCodeImmo : string;NumOrdreLog : integer);
begin
  CodeI := vCodeImmo;
  OrdreSerie := NumOrdreLog;
  if Transactions(EnregistreMutationSerie, 1)<>oeOK then HM.execute(1,Caption,'')
  else
  begin
       VHImmo^.ChargeOBImmo := True;
       ImMarquerPublifi (True);   // CA le 06/10/2000
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. : Enregistrement d'une mutation en série
Suite ........ : 24/02/2003 : affectation correcte du plan actif ( était
Suite ........ : auparavant sytématiquement forcé à 0 )
Mots clefs ... : MUTATION;PLANACTIF
*****************************************************************}
procedure TOperationSerie.EnregistreMutationSerie;
var Query, QueryC : TQuery ;
    AncienCompte, AncienCompteLie : string;
    Ordre, PlanActif : integer;
    Nature : string;
    NaturePRO : boolean;
begin
  Ordre := TrouveNumeroOrdreLogSuivant(CodeI);
  QueryC:=OpenSQL('SELECT I_COMPTEIMMO,I_NATUREIMMO,I_PLANACTIF, I_COMPTELIE FROM IMMO WHERE I_IMMO="'+CodeI+'"',TRUE) ;
  if not QueryC.EOF then
  begin
    AncienCompte:=QueryC.FindField('I_COMPTEIMMO').AsString;
    AncienCompteLie:=QueryC.FindField('I_COMPTELIE').AsString;
    NaturePRO := ((QueryC.FindField('I_NATUREIMMO').AsString='PRO') or
                  (QueryC.FindField('I_NATUREIMMO').AsString='FI'));

    // FQ 19099 Ne pas muter l'immo si compte origine = compte destination
    if ((NaturePRO) and (AncienCompte <> EnregMutation.Compte))
    or ((not NaturePRO) and (AncienCompteLie <> EnregMutation.Compte)) then
    begin

      Nature := QueryC.FindField('I_NATUREIMMO').AsString;
      PlanActif := QueryC.FindField('I_PLANACTIF').AsInteger;
      Ferme(QueryC) ;
      Query:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+CodeI+'"',FALSE);
      Query.Insert ;
      Query.FindField('IL_IMMO').AsString:=CodeI;
      {$IFDEF EAGLCLIENT}
      Query.PutValue('IL_BLOCNOTE',EnregMutation.BlocNote.Text);
      {$ELSE}
      TBlobField(Query.FindField('IL_BLOCNOTE')).Assign(EnregMutation.BlocNote);
      {$ENDIF}
      Query.FindField('IL_TYPEOP').AsString:='MUT' ;
      Query.FindField('IL_ORDRE').AsInteger:=Ordre;
      Query.FindField('IL_ORDRESERIE').AsInteger:=OrdreSerie;
      Query.FindField('IL_CODEMUTATION').AsString:='';
      // FQ 19208
      if NaturePRO then
         Query.FindField('IL_CPTEMUTATION').AsString := AncienCompte
      else
         Query.FindField('IL_CPTEMUTATION').AsString := AncienCompteLie;
      Query.FindField('IL_DATEOP').AsDateTime:= EnregMutation.DateOpe;
      Query.FindField('IL_PLANACTIFAV').AsInteger:=PlanActif;
      Query.FindField('IL_PLANACTIFAP').AsInteger:=PlanActif;
      Query.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', Query.FindField('IL_TYPEOP').AsString, FALSE)+' '+DateToStr(EnregMutation.DateOpe) ;
      //  Query.FindField('IL_TYPEMODIF').AsString:=RechDom('TIOPEAMOR', Query.FindField('IL_TYPEOP').AsString, FALSE)+' '+DateToStr(EnregMutation.DateOpe) ;
      Query.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('MUT'); //15/01/99 EPZ
      Query.Post ;
      Ferme(Query) ;
      // FQ 19208
      if NaturePRO then
         begin
         ExecuteSQL('UPDATE IMMOAMOR SET IA_COMPTEIMMO="'+EnregMutation.Compte+'" WHERE IA_IMMO="'+CodeI+'"');
         ExecuteSQL('UPDATE IMMO SET I_COMPTEIMMO="'+EnregMutation.Compte+'",I_COMPTEREF="'+EnregMutation.Compte+'" WHERE I_IMMO="'+CodeI+'"') ;
         end
      else
         ExecuteSQL('UPDATE IMMO SET I_COMPTELIE="'+EnregMutation.Compte+'",I_COMPTEREF="'+EnregMutation.Compte+'" WHERE I_IMMO="'+CodeI+'"') ;
      // 15/01/99
      //  ExecuteSQL('UPDATE IMMO SET I_OPEMUTATION="X" WHERE I_IMMO="'+CodeI+'"') ;
      Query := nil;
      CocheChampOperation (Query,CodeI,'I_OPEMUTATION');
      // 15/01/99
      if Nature = 'PRO' then
         MajComptesAssocies (CodeI,EnregMutation.Compte);
      // FQ 13234
      if (Copy(EnregMutation.Compte, 1, 2) = '20') and (Copy(AncienCompte, 1, 2) <> '20') then
         ExecuteSQL ('UPDATE IMMO SET I_BASETAXEPRO=0 WHERE I_IMMO="'+CodeI+'"');
    end;
  end else Ferme(QueryC) ;
end;

procedure TOperationSerie.OperationSerieCession(vCodeImmo : string;NumOrdreLog : integer; QImmo : TQuery);
begin
  QteAchat := QImmo.FindField('I_QUANTITE').AsFloat;
  MontantHT := QImmo.FindField('I_MONTANTHT').AsFloat;
  BaseEco := QImmo.FindField('I_BASEECO').AsFloat;
  BaseFisc := QImmo.FindField('I_BASEFISC').AsFloat;
  CodeI := vCodeImmo;
  OrdreSerie := NumOrdreLog;
  if Transactions(EnregistreCessionSerie, 1)<>oeOK then HM.execute(1,Caption,'')
  else
  begin
       VHImmo^.ChargeOBImmo := True;
       ImMarquerPublifi (True);   // CA le 06/10/2000
  end;
end;

procedure TOperationSerie.EnregistreCessionSerie;
var CessionPartielle : boolean ;
    PlanImmoCedee : TPlanAmort ;
    VoCedee : double;
    ElemCession : TEnregCession;
begin
  ElemCession := TEnregCession.Create;
  ElemCession.Code := CodeI;
  ElemCession.Vo := MontantHT ;//YCP 19-06-01 EnregCession.PrixCession;
  ElemCession.QteAchat := QteAchat;
//  ElemCession.QteCedee := (EnregCession.PrixCession/MontantHT)*QteAchat;
  ElemCession.QteCedee := QteAchat; // on cede tout
  ElemCession.DateOpe := EnregCession.DateOpe;
  ElemCession.MontantAchat := MontantHT;
  ElemCession.ChangePlan := false;
  if EnregCession.RepartitionImmo then
    ElemCession.MontantCession := EnregCession.PrixCession*(MontantHT/EnregCession.CumulImmoSel)
  else ElemCession.MontantCession := EnregCession.PrixCession;
  ElemCession.Tva := EnregCession.TvaReverser;
  ElemCession.MotifCession:=EnregCession.Motif;
  ElemCession.CalcCession:=EnregCession.CalculDot;
  ElemCession.BlocNote := EnregCession.BlocNote;
  ElemCession.BaseEco := BaseEco;
  ElemCession.BaseFisc := BaseFisc;
  CessionPartielle:=(ElemCession.QteCedee<>ElemCession.QteAchat); // Cession partielle ?
  if (ElemCession.QteAchat=1) and CessionPartielle then ElemCession.QteCedee := 0;
  // Mise à jour de la fiche immo d'origine
  // Modif CA le 24/06/99 - Création plan immo cédée
  PlanImmoCedee := TPlanAmort.Create(true);
  TraiteImmoCedee(PlanImmoCedee,ElemCession);
  VoCedee:=ElemCession.QteCedee*(ElemCession.MontantAchat/ElemCession.QteAchat);
  ElemCession.PValue := ElemCession.MontantCession-VoCedee+ElemCession.AmortEco;
  TraiteImmoOrig(PlanImmoCedee,CessionPartielle,ElemCession);
  EnregLogCession(PlanImmoCedee,CessionPartielle,ElemCession,OrdreSerie);
  PlanImmoCedee.free ; //Detruit;
  ElemCession.free ;//Destroy;
end;

procedure TOperationSerie.OperationSerieEtabl(vCodeImmo : string; NumOrdreLog : integer);
var Qtmp : TQuery;
begin
CodeI := vCodeImmo;
OrdreSerie := NumOrdreLog;
Qtmp := OpenSQL('SELECT I_ETABLISSEMENT, I_PLANACTIF FROM IMMO WHERE I_IMMO="'+vCodeImmo+'"',TRUE);
ValeurPrec := Qtmp.FindField('I_ETABLISSEMENT').AsString;
fPlanActif := Qtmp.FindField('I_PLANACTIF').AsInteger;
EtablissementPrec:=RechDom('TTETABLISSEMENT', ValeurPrec, FALSE);
Ferme(Qtmp);
if EnregEtablissement.Etablissement <>  ValeurPrec then // FQ 19096
begin
  if Transactions(TraiteEnregistrementOperation, 1)<>oeOK then HM.execute(1,Caption,'')
  else
  begin
     VHImmo^.ChargeOBImmo := True;
     ImMarquerPublifi (True);   // CA le 06/10/2000
  end;
end;
end;

procedure TOperationSerie.OperationSerieLieu(vCodeImmo : string; NumOrdreLog : integer);
var Qtmp : TQuery;
begin
CodeI := vCodeImmo;
OrdreSerie := NumOrdreLog;
Qtmp := OpenSQL('SELECT I_LIEUGEO, I_PLANACTIF FROM IMMO WHERE I_IMMO="'+vCodeImmo+'"',TRUE);
ValeurPrec := Qtmp.FindField('I_LIEUGEO').AsString;
fPlanActif := Qtmp.FindField('I_PLANACTIF').AsInteger;
LieuPrec:=RechDom('TILIEUGEO', ValeurPrec, FALSE);
Ferme(Qtmp);
if EnregLieu.Lieu <> ValeurPrec then // FQ 19096
begin
  if Transactions(TraiteEnregistrementOperation, 1)<>oeOK then HM.execute(1,Caption,'')
  else
  begin
     VHImmo^.ChargeOBImmo := True;
     ImMarquerPublifi (True);   // CA le 06/10/2000
  end;
end;
end;

procedure ExecuteEtablissement(CodeImmo : string);
var FOpeSerie: TOperationSerie;
    ret : integer;
    PlanActif : integer;
    QueryC : TQuery ;
    AncienEtabl : string;
begin
  QueryC:=OpenSQL('SELECT I_ETABLISSEMENT, I_PLANACTIF FROM IMMO WHERE I_IMMO="'+CodeImmo+'"',TRUE) ;
  if QueryC.EOF then AncienEtabl:=''
  else AncienEtabl:=QueryC.FindField('I_ETABLISSEMENT').AsString;
  PlanActif := QueryC.FindField('I_PLANACTIF').AsInteger;
  Ferme(QueryC) ;
  FOpeSerie := TOperationSerie.Create(Application);
  try
    FOpeSerie.InitOPerationSerie(toChanEtabl,1);
    FOpeSerie.CodeI := CodeImmo;
    FOpeSerie.OrdreSerie := -1;
    FOpeSerie.ValeurPrec := AncienEtabl;
    FOpeSerie.fPlanActif := PlanActif;
    FOpeSerie.EtablissementPrec:=RechDom('TTETABLISSEMENT',AncienEtabl,FALSE);
    ret := FOpeSerie.AfficheOperationSerie;
    if (ret = mrOK) or (ret = mrYes) then
        FOpeSerie.TraiteEnregistrementOperation;
  finally
    FOpeSerie.Free;
  end;
end;

procedure TOperationSerie.TraiteEnregistrementOperation;
begin
  case TypeOpe of
    toChanEtabl: begin
                if Transactions(EnregistreEtablissement,1) <> oeOk  then
                        HM.execute(1,Caption,'') else
                        begin
                             VHImmo^.ChargeOBImmo := True;
                             ImMarquerPublifi (True);   // CA le 06/10/2000
                        end;
                end;
    toChanLieu: begin
                if Transactions(EnregistreLieu,1) <> oeOk  then
                        HM.execute(1,Caption,'') else
                        begin
                             VHImmo^.ChargeOBImmo := True;
                             IMMarquerPublifi (True);   // CA le 06/10/2000
                        end;

                end;
    toRegroupement: begin
                      if Transactions(EnregistreRegroupement,1) <> oeOk  then
                        HM.execute(1,Caption,'') else
                      begin
                        VHImmo^.ChargeOBImmo := True;
                        // FQ 18119 Positionner l'indicateur de modif de compta
                        IMMarquerPublifi (True);
                      end;
                    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : Enregistrement de l'opération changement d'établissement
Suite ........ : 24/02/2003 : Mise à jour correcte du plan actif
Mots clefs ... : CHANGEMENT ETABLISSEMENT
*****************************************************************}
procedure TOperationSerie.EnregistreEtablissement;
var Query : TQuery ;
    OrdreS,Ordre : integer;
begin
  Ordre := TrouveNumeroOrdreLogSuivant(CodeI);
  if OrdreSerie=-1 then OrdreS := TrouveNumeroOrdreSerieLogSuivant // modif serie ou non
  else OrdreS := OrdreSerie;
  Query:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+CodeI+'"',FALSE);
  Query.Insert ;
  Query.FindField('IL_IMMO').AsString:=CodeI;
  {$IFDEF EAGLCLIENT}
  Query.PutValue('IL_BLOCNOTE',EnregEtablissement.BlocNote.Text);
  {$ELSE}
  TBlobField(Query.FindField('IL_BLOCNOTE')).Assign(EnregEtablissement.BlocNote);
  {$ENDIF}
  Query.FindField('IL_TYPEOP').AsString:='ETA' ;
  Query.FindField('IL_ORDRE').AsInteger:=Ordre;
  Query.FindField('IL_ORDRESERIE').AsInteger:=OrdreS ;
  Query.FindField('IL_CODEMUTATION').AsString:='';
  Query.FindField('IL_CPTEMUTATION').AsString := '';
  Query.FindField('IL_DATEOP').AsDateTime:= EnregEtablissement.DateOpe;
  Query.FindField('IL_PLANACTIFAV').AsFloat:=fPlanActif;
  Query.FindField('IL_PLANACTIFAP').AsFloat:=fPlanActif;
  Query.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', Query.FindField('IL_TYPEOP').AsString, FALSE)+' '+DateToStr(EnregEtablissement.DateOpe) ;
//  Query.FindField('IL_TYPEMODIF').AsString:=RechDom('TIOPEAMOR', Query.FindField('IL_TYPEOP').AsString, FALSE)+' '+DateToStr(EnregMutation.DateOpe) ;
  Query.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('ETA'); //15/01/99 EPZ
  Query.FindField('IL_ETABLISSEMENT').AsString:=ValeurPrec;
  Query.Post ;
  Ferme(Query) ;
  ExecuteSQL('UPDATE IMMO SET I_ETABLISSEMENT="'+EnregEtablissement.Etablissement+'" WHERE I_IMMO="'+CodeI+'"') ;
// 15/01/99
//  ExecuteSQL('UPDATE IMMO SET I_OPEETABLISSEMENT="X" WHERE I_IMMO="'+CodeI+'"') ;
  Query := nil;
  CocheChampOperation (Query,CodeI,'I_OPEETABLISSEMENT');
// 15/01/99
end;

(*procedure TOperationSerie.AccelCreationLieuClick(Sender: TObject);
var OldVal : string;
begin
  OldVal := LieuGeo.Value;
  ParamTable('TILIEUGEO',taCreat,0,nil) ;
  LieuGeo.Reload;
  LieuGeo.Value := OldVal;
end;*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : Enregistrement de l'opération changement de lieu
Suite ........ : 24/02/2003 : Mise à jour correcte du plan actif
Mots clefs ... : CHANGEMENT LIEU
*****************************************************************}
procedure TOperationSerie.EnregistreLieu;
var Query : TQuery ;
    OrdreS,Ordre : integer;
begin
    Ordre := TrouveNumeroOrdreLogSuivant(CodeI);
    if OrdreSerie=-1 then OrdreS := TrouveNumeroOrdreSerieLogSuivant // modif serie ou non
    else OrdreS := OrdreSerie;
    Query:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+CodeI+'"',FALSE);
    Query.Insert ;
    Query.FindField('IL_IMMO').AsString:=CodeI;
    {$IFDEF EAGLCLIENT}
    Query.PutValue('IL_BLOCNOTE',EnregLieu.BlocNote.Text);
    {$ELSE}
    TBlobField(Query.FindField('IL_BLOCNOTE')).Assign(EnregLieu.BlocNote);
    {$ENDIF}
    Query.FindField('IL_TYPEOP').AsString:='LIE' ;
    Query.FindField('IL_ORDRE').AsInteger:=Ordre;
    Query.FindField('IL_ORDRESERIE').AsInteger:=OrdreS ;
    Query.FindField('IL_CODEMUTATION').AsString:='';
    Query.FindField('IL_CPTEMUTATION').AsString := '';
    Query.FindField('IL_DATEOP').AsDateTime:= EnregLieu.DateOpe;
    //  Query.FindField('IL_PLANACTIFAV').AsFloat:=0;
    Query.FindField('IL_PLANACTIFAV').AsFloat:=fPlanActif;
    //  Query.FindField('IL_PLANACTIFAP').AsFloat:=0;
    Query.FindField('IL_PLANACTIFAP').AsFloat:=fPlanActif;
    Query.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', Query.FindField('IL_TYPEOP').AsString, FALSE)+' '+DateToStr(EnregLieu.DateOpe) ;
    //  Query.FindField('IL_TYPEMODIF').AsString:=RechDom('TIOPEAMOR', Query.FindField('IL_TYPEOP').AsString, FALSE)+' '+DateToStr(EnregMutation.DateOpe) ;
    Query.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('LIE'); //15/01/99 EPZ
    Query.FindField('IL_LIEUGEO').AsString:=ValeurPrec;
    Query.Post ;
    Ferme(Query) ;
    ExecuteSQL('UPDATE IMMO SET I_LIEUGEO="'+EnregLieu.Lieu+'" WHERE I_IMMO="'+CodeI+'"') ;
    // 15/01/99
    //  ExecuteSQL('UPDATE IMMO SET I_OPELIEUGEO="X" WHERE I_IMMO="'+CodeI+'"') ;
    Query := nil;
    CocheChampOperation (Query,CodeI,'I_OPELIEUGEO');
    // 15/01/99
end;

procedure ExecuteLieu(CodeImmo : string);
var FOpeSerie: TOperationSerie;
    ret, PlanActif : integer;
    QueryC : TQuery ;
    AncienLieu : string;
begin
  QueryC:=OpenSQL('SELECT I_LIEUGEO, I_PLANACTIF FROM IMMO WHERE I_IMMO="'+CodeImmo+'"',TRUE) ;
  if QueryC.EOF then AncienLieu:=''
  else AncienLieu:=QueryC.FindField('I_LIEUGEO').AsString;
  PlanActif := QueryC.FindField('I_PLANACTIF').AsInteger;
  Ferme(QueryC) ;
  FOpeSerie := TOperationSerie.Create(Application);
  try
    FOpeSerie.InitOPerationSerie(toChanLieu,1);
    FOpeSerie.CodeI := CodeImmo;
    FOpeSerie.OrdreSerie := -1;
    FOpeSerie.ValeurPrec := AncienLieu;
    FOpeSerie.fPlanActif := PlanActif;
    FOpeSerie.LieuPrec:=RechDom('TILIEUGEO', AncienLieu, FALSE);
    ret := FOpeSerie.AfficheOperationSerie;
    if (ret = mrOK) or (ret = mrYes) then
        FOpeSerie.TraiteEnregistrementOperation;
  finally
    FOpeSerie.Free;
  end;
end;

procedure TOperationSerie.CalculCumulImmoCedee(Liste : THGrid ; Q : THQuery);
var i : integer;
begin
  EnregCession.CumulImmoSel:=0.00;
  if Liste.AllSelected then
  BEGIN
    InitMove(100,'');
    Q.First;
    while Not Q.EOF do
    begin
      MoveCur(False);
      EnregCession.CumulImmoSel := EnregCession.CumulImmoSel+
                                Q.FindField('I_MONTANTHT').AsFloat ;
      Q.Next;
    end;
  END
  ELSE
  BEGIN
    InitMove(Liste.NbSelected,'');
    for i:=0 to Liste.NbSelected-1 do
    begin
      MoveCur(False);
      Liste.GotoLeBookmark(i);
      EnregCession.CumulImmoSel := EnregCession.CumulImmoSel+
                                Q.FindField('I_MONTANTHT').AsFloat ;
    end;
  END;
  FiniMove;
end;

procedure TOperationSerie.DateOperationKeyPress(Sender: TObject;
  var Key: Char);
begin
  ParamDate(Self,Sender,Key);
end;

// FQ 19099 Mutation en série mais immo unique, contrôle date opération
function TOperationSerie.DateOperationOK : boolean;
var Q: TQuery;
begin
  result := True;
  // La date est-elle postérieure à la date de dernière opération dans l'exo ?
  Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="' + fCode +
                '" ORDER BY IL_DATEOP DESC', True);
  if not Q.Eof then
    if Q.FindField ('IL_DATEOP').AsDateTime > StrToDate(DateOperation.Text) then
       result := False;
end;



function TOperationSerie.ControleCompteMutation : boolean;
var CpteSup,CpteInf,Compte : string;
begin
  // FQ 19099 Compléter d'abord à la longueur standard
  Compte := CompteMutation.Text;
  Compte := ImBourreEtLess (Compte, ImGeneTofb);
  CompteMutation.Text := Compte;

  if fNature = 'CB' then
  begin CpteSup := VHImmo^.CpteCBSup; CpteInf := VHImmo^.CpteCBInf;end
  else if fNature = 'LOC' then
  begin CpteSup := VHImmo^.CpteLocSup; CpteInf := VHImmo^.CpteLocInf;end
  else  if fNature = 'PRO' then
  begin CpteSup := VHImmo^.CpteImmoSup; CpteInf := VHImmo^.CpteImmoInf;end
  else if fNature = 'FI' then
  begin CpteSup := VHImmo^.CpteFinSup; CpteInf := VHImmo^.CpteFinInf;end;
  if (CompteMutation.Text < CpteInf) or (CompteMutation.Text > CpteSup) then result := false
  else if not Presence ('GENERAUX','G_GENERAL',CompteMutation.Text) then result := false
  else result := true;
end;

procedure TOperationSerie.CompteMutationEnter(Sender: TObject);
begin
  fCurCompte := CompteMutation.Text;
end;

procedure TOperationSerie.CompteMutationExit(Sender: TObject);
var Compte : string;
begin
  inherited;
  Compte := CompteMutation.Text;
  if Compte = fCurCompte then exit;
  if Compte = '' then exit;
  Compte := ImBourreEtLess ( Compte,ImGeneTofb);
  if Presence('GENERAUX','G_GENERAL',Compte) then
    CompteMutation.Text := Compte;
end;

procedure TOperationSerie.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TOperationSerie.CompteMutationElipsisClick(Sender: TObject);
var CpteSup,CpteInf : string;
    sWhere : string;
begin
  if fNature = 'CB' then
  begin CpteSup := VHImmo^.CpteCBSup; CpteInf := VHImmo^.CpteCBInf;end
  else if fNature = 'LOC' then
  begin CpteSup := VHImmo^.CpteLocSup; CpteInf := VHImmo^.CpteLocInf;end
  else  if fNature = 'PRO' then
  begin CpteSup := VHImmo^.CpteImmoSup; CpteInf := VHImmo^.CpteImmoInf;end
  else if fNature = 'FI' then
  begin CpteSup := VHImmo^.CpteFinSup; CpteInf := VHImmo^.CpteFinInf;end;
  sWhere := 'G_GENERAL<="'+CpteSup+'" AND G_GENERAL>="'+CpteInf+'"';
  LookUpList (TControl (Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',sWhere,'G_GENERAL',True,21) ;
end;

procedure ExecuteRegroupement(CodeImmo : string);
var FOpeSerie: TOperationSerie;
    ret : integer;
    PlanActif : integer;
    QueryC : TQuery ;
    AncienRegroupement : string;
begin
  QueryC:=OpenSQL('SELECT I_GROUPEIMMO, I_PLANACTIF FROM IMMO WHERE I_IMMO="'+CodeImmo+'"',TRUE) ;
  if QueryC.EOF then AncienRegroupement:=''
  else AncienRegroupement:=QueryC.FindField('I_GROUPEIMMO').AsString;
  PlanActif := QueryC.FindField('I_PLANACTIF').AsInteger;
  Ferme(QueryC) ;
  FOpeSerie := TOperationSerie.Create(Application);
  try
    FOpeSerie.InitOPerationSerie(toRegroupement,1);
    FOpeSerie.CodeI := CodeImmo;
    FOpeSerie.OrdreSerie := -1;
    FOpeSerie.ValeurPrec := AncienRegroupement;
    FOpeSerie.fPlanActif := PlanActif;
    FOpeSerie.RegroupementPrec:=RechDom('AMREGROUPEMENT',AncienRegroupement,FALSE);
    ret := FOpeSerie.AfficheOperationSerie;
    if (ret = mrOK) or (ret = mrYes) then
        FOpeSerie.TraiteEnregistrementOperation;
  finally
    FOpeSerie.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/02/2003
Modifié le ... : 18/05/2004
Description .. : Enregistrement de l'opération changement de regroupement
Suite ........ :
Mots clefs ... : CHANGEMENT REGROUPEMENT
*****************************************************************}
procedure TOperationSerie.EnregistreRegroupement;
var Query : TQuery;
    TLog : TOB;
    OrdreS,Ordre : integer;
begin
  { Dans ce cas précis, on ne renseigne pas IMMOLOG : on ne souhaite pas garder
    d'historique des changements de regroupement }
  // BTY 17516 Enregister en tant qu'opération REG
  //ExecuteSQL('UPDATE IMMO SET I_GROUPEIMMO="'+fRegroupement+'" WHERE I_IMMO="'+CodeI+'"') ;

  Ordre := TrouveNumeroOrdreLogSuivant(CodeI);
  if OrdreSerie=-1 then OrdreS := TrouveNumeroOrdreSerieLogSuivant // modif serie ou non
  else OrdreS := OrdreSerie;

  // Ajout IMMOLOG
  TLog := TOB.Create ('IMMOLOG',nil,-1);
  try
    TLog.PutValue('IL_IMMO', CodeI);
    TLog.PutValue('IL_TYPEOP', 'REG');
    TLog.PutValue('IL_ORDRE', Ordre);
    TLog.PutValue('IL_ORDRESERIE', OrdreS);
    TLog.PutValue('IL_DATEOP', EnregGroupe.DateOpe);
    TLog.PutValue('IL_PLANACTIFAV', fPlanActif);
    TLog.PutValue('IL_PLANACTIFAP', fPlanActif);
    TLog.PutValue('IL_LIBELLE', RechDom('TIOPEAMOR', 'REG', FALSE)+
                                ' ' + DateToStr(EnregGroupe.DateOpe) );
    TLog.PutValue('IL_TYPEMODIF', AffecteCommentaireOperation('REG') );
    TLog.PutValue('IL_LIEUGEO', ValeurPrec);
    TLog.PutValue('IL_BLOCNOTE', EnregGroupe.BlocNote);
    TLog.InsertDB(nil);
  finally
    TLog.Free;
  end;

    {  Query.FindField('IL_IMMO').AsString:= CodeI ;
  Query.FindField('IL_TYPEOP').AsString:='REG' ;
  Query.FindField('IL_ORDRE').AsInteger:=Ordre;
  Query.FindField('IL_ORDRESERIE').AsInteger:=OrdreS ;
  Query.FindField('IL_CODEMUTATION').AsString:='';
  Query.FindField('IL_CPTEMUTATION').AsString := '';
  Query.FindField('IL_DATEOP').AsDateTime:= EnregGroupe.DateOpe;
  Query.FindField('IL_PLANACTIFAV').AsFloat:=fPlanActif;
  Query.FindField('IL_PLANACTIFAP').AsFloat:=fPlanActif;
  Query.FindField('IL_LIBELLE').AsString:= RechDom('TIOPEAMOR',
                                           Query.FindField('IL_TYPEOP').AsString, FALSE) +
                                           ' ' + DateToStr(EnregGroupe.DateOpe) ;
  Query.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('REG');
  Query.FindField('IL_LIEUGEO').AsString:=ValeurPrec;
  Query.Post ;
  Ferme(Query) ;}

  // Maj IMMO
  ExecuteSQL('UPDATE IMMO SET I_GROUPEIMMO="'+EnregGroupe.Regroupement+'" WHERE I_IMMO="'+CodeI+'"') ;
  Query := nil;
  CocheChampOperation (Query,CodeI,'I_OPEREG');
end;

//procedure TOperationSerie.OperationSerieRegroupement(vCodeImmo: string );
procedure TOperationSerie.OperationSerieRegroupement(vCodeImmo: string; NumOrdreLog : integer );
var Qtmp : TQuery;
begin
  CodeI := vCodeImmo;
  OrdreSerie := NumOrdreLog; // BTY FQ 17516
  Qtmp := OpenSQL('SELECT I_GROUPEIMMO, I_PLANACTIF FROM IMMO WHERE I_IMMO="'+vCodeImmo+'"',TRUE);
  ValeurPrec := Qtmp.FindField('I_GROUPEIMMO').AsString;
  fPlanActif := Qtmp.FindField('I_PLANACTIF').AsInteger;
  RegroupementPrec:=RechDom('AMREGROUPEMENT', ValeurPrec, FALSE);
  Ferme(Qtmp);
  if EnregGroupe.Regroupement <> ValeurPrec then // FQ 19096
  begin
    if Transactions(TraiteEnregistrementOperation, 1)<>oeOK then HM.execute(1,Caption,'')
    // BTY FQ 18119 Positionner l'indicateur de modif de la compta
    //else VHImmo^.ChargeOBImmo := True;
    else
    begin
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi (True);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/06/2004
Modifié le ... :   /  /
Description .. : Affichage des informations du regroupement
Mots clefs ... :
*****************************************************************}
procedure TOperationSerie.BINFOREGClick(Sender: TObject);
var T : TOB;
    ModeleExiste:boolean;
begin
  if (FCODEREGROUPEMENT.Value <> '') then
    begin
    // BTY 01/06 FQ 17378 Afficher la fiche du regroupement uniquement si les infos existent
    // sinon message
    T := TOB.Create ('IMMO',nil,-1);
    try
     ModeleExiste := (T.SelectDB('"&#@MODELE_'+FCODEREGROUPEMENT.Value+'"', nil, True));
    finally
     T.Free;
    end;
    if not ModeleExiste then
       PGIInfo (MSG_NOMODELE, RechDom('AMREGROUPEMENT',FCODEREGROUPEMENT.Value,False)) // (MSG_NOMODELE)
    else
       // PGR 01/07/2005 FQ 15021
       AmLanceFiche_FicheRegroupement ('&#@MODELE_' +  FCODEREGROUPEMENT.Value);
    end;

  // fq 15019 - mbo 28.07.2005
  BinfoReg.ModalResult := MrNone;

{  if (FCODEREGROUPEMENT.Value <> '') then
    // PGR 01/07/2005 FQ 15021
    // AmLanceFiche_FicheRegroupement ('&#@MODELE_' +  FCODEREGROUPEMENT.Text);
    AmLanceFiche_FicheRegroupement ('&#@MODELE_' +  FCODEREGROUPEMENT.Value);
    // fq 15019 - mbo 28.07.2005
    BinfoReg.ModalResult := MrNone;
}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/06/2004
Modifié le ... :   /  /
Description .. : Affichage des éléments qui composent le regroupement
Mots clefs ... : 
*****************************************************************}
procedure TOperationSerie.BDETAILREGClick(Sender: TObject);
begin
  if (FCODEREGROUPEMENT.Value <> '') then
    // PGR 01/07/2005 FQ 15021
    //AMLanceFiche_ListeDesImmobilisations ( '', False, taConsult, FCODEREGROUPEMENT.Text);
    AMLanceFiche_ListeDesImmobilisations ( '', False, taConsult, FCODEREGROUPEMENT.Value);
    // fq 15019 - mbo 28.07.2005
    BDetailReg.ModalResult := MrNone;

end;

procedure TOperationSerie.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_F10 then bValiderClick(nil);
end;

// BTY 09/06 FQ 18775
procedure TOperationSerie.OnClickSaisieRegroupement(Sender: TObject);
begin
  AMLanceFiche_SaisieRegroupement('IGI');

  THValComboBox(FCodeRegroupement).Reload;
end;



end.
