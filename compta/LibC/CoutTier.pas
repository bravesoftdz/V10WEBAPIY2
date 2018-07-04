(*------------------------------------------------------------------------------------
Description       :   Calcul des coûts financiers des tiers
Procédure d'appel :   LanceCoutTiers
Paramètres        :   Compte auxiliaire & Date de référence
Dernière version  :   04/04/97
--------------------------------------------------------------------------------------*)
               
unit CoutTier ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, Buttons, ExtCtrls, Hctrls, Mask, ComCtrls, hmsgbox, DB,
  Hqry, Menus, HDB, HEnt1, Ent1,
  CPTIERS_TOM,
  HStatus,
  Saisie{, Filtre SG6 17/11/04 Gestion de filtres FQ 14976}, ParamDat, {EncUtil,} SaisComm, CritEdt, UtilEdt,
  SaisUtil, Graph, Math, Teengine, Chart, Series, HSysMenu, Hcompte, LettUtil,
  HTB97, TeeProcs, HPanel, UiUtil,
  SaisBor,UObjFiltres {SG6 17/11/04 Gestion des filtres FQ 14976},
  Utob,
  {$IFDEF EAGLCLIENT}
  UtileAGL
  {$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  EdtREtat,
  PrintDBG
  {$ENDIF}
;

procedure LanceCoutTiers ( LeAuxi : String ; LaDateRef : TDateTime ) ;
function  EcartTypeGrid(G : THGrid ; col : Integer) : Extended ;
function  SommeGrid(G : THGrid ; col : Integer) : Extended ;
function  MoyenneGrid(G : THGrid ; col : Integer) : Extended ;

type
  TFCoutTiers = class(TForm)
    PFiltres: TToolWindow97;
    TT_FILTRESCRITERES: TLabel;
    FFiltres: THValComboBox;
    BChercher: TToolbarButton97;
    HPB: TToolWindow97;
    BRecherche: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    FindMvt: THFindDialog;
    POPS: TPopupMenu;
    BZoomTiers: THBitBtn;
    MsgBox: THMsgBox;
    POPZ: TPopupMenu;
    Pages: TPageControl;
    Princ: TTabSheet;
    Bevel1: TBevel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel2: THLabel;
    HLabel1: THLabel;
    rSelection: TRadioGroup;
    E_GENERAL_SUP: THCpteEdit;
    E_GENERAL_INF: THCpteEdit;
    E_AUXILIAIRE_INF: THCpteEdit;
    E_AUXILIAIRE_SUP: THCpteEdit;
    Complements: TTabSheet;
    Bevel2: TBevel;
    TE_JOURNAL: THLabel;
    HLabel11: THLabel;
    Label2: THLabel;
    HLabel8: THLabel;
    HLabel10: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    E_JOURNAL: THValComboBox;
    E_NUMEROPIECE: THCritMaskEdit;
    E_NUMEROPIECE_: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_ECHE: TEdit;
    E_NUMECHE: THCritMaskEdit;
    TabSheet1: TTabSheet;
    Bevel4: TBevel;
    TT_LIBELLE: THLabel;
    TT_EAN: THLabel;
    TT_ADRESSE1: THLabel;
    TT_CODEPOSTAL: THLabel;
    TT_VILLE: THLabel;
    TT_Secteur: THLabel;
    HLabel13: THLabel;
    T_LIBELLE: TEdit;
    T_EAN: TEdit;
    T_ADRESSE1: TEdit;
    T_CODEPOSTAL: TEdit;
    T_VILLE: TEdit;
    T_SECTEUR: THValComboBox;
    T_TELEPHONE: TEdit;
    TabSheet2: TTabSheet;
    Bevel3: TBevel;
    HLabel6: THLabel;
    cAvoirs: TCheckBox;
    DateReference: THCritMaskEdit;
    E_QUALIFPIECE: TEdit;
    HLabel3: THLabel;
    HLabel7: THLabel;
    TauxNormal: THNumEdit;
    TauxRetard: THNumEdit;
    HLabel9: THLabel;
    HLabel12: THLabel;
    GT: THGrid;
    TRacine: THLabel;
    RacineEscompte: THCritMaskEdit;
    PanelPieces: TPanel;
    GP: THGrid;
    TitrePanelPieces: TPanel;
    Panel1: TPanel;
    BStop: TToolbarButton97;
    bZoomMvts: THBitBtn;
    BAgrandir: TToolbarButton97;
    BReduire: TToolbarButton97;
    Panel2: TPanel;
    bImprimerP: THBitBtn;
    bValideP: THBitBtn;
    bFermerPieces: THBitBtn;
    BGraph: TToolbarButton97;
    bGraph1: THBitBtn;
    bGraph2: THBitBtn;
    TSeuil1: THLabel;
    MontantSeuil: THNumEdit;
    TSeuil2: THLabel;
    RatioSeuil: THNumEdit;
    HLabel20: THLabel;
    RatioMoyenP: THLabel;
    TotalGeneralP: TPanel;
    HLabel18: THLabel;
    PanelStats: TPanel;
    Panel4: TPanel;
    GS: THGrid;
    bStats: TToolbarButton97;
    bFermerS: THBitBtn;
    bImprimerS: THBitBtn;
    Pie: TPieSeries;
    Chart1: TChart;
    Chart2: TChart;
    Pie2: TPieSeries;
    HLabel15: THLabel;
    NbTiers: TPanel;
    HMTrad: THSystemMenu;
    PLibres: TTabSheet;
    Bevel5: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    HLabel14: THLabel;
    TotalGeneral: TPanel;
    BImprimer: TToolbarButton97;
    BDetail: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock: TDock97;
    Dock971: TDock97;
    procedure BAnnulerClick(Sender: TObject);
    procedure BZoomTiersClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure E_DATEECHEANCEKeyPress(Sender: TObject; var Key: Char);

    procedure BDetailClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure bFermerPClick(Sender: TObject);
    procedure GTDblClick(Sender: TObject);
    procedure bZoomMvtsClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure bValidePClick(Sender: TObject);
    procedure GPDblClick(Sender: TObject);
    procedure bImprimerPClick(Sender: TObject);
    procedure DateReferenceExit(Sender: TObject);
    procedure bGraph3Click(Sender: TObject);
    procedure bGraph2Click(Sender: TObject);
    procedure bGraph1Click(Sender: TObject);
    procedure MontantSeuilEnter(Sender: TObject);
    procedure TauxNormalEnter(Sender: TObject);
    procedure TauxRetardEnter(Sender: TObject);
    procedure RatioSeuilEnter(Sender: TObject);
    procedure bStatsClick(Sender: TObject);
    procedure bFermerSClick(Sender: TObject);
    procedure GSDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure bImprimerSClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);

    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BGraphMouseEnter(Sender: TObject);
  private
    { Déclarations privées }
    ObjFiltre:TObjFiltre;  //SG6 17/11/04 Gestion des Filtres FQ 14976
    StopProcess,AppelAutre : Boolean ;
    DateRef       : TDateTime ;
    OldBrush      : TBrush ;
    OldPen        : TPen ;
    Autres        : Boolean ;
    WMinX,WMinY : Integer ;
    QEcr:TQuery ;
    TobP, TobS, TobT : Tob ;
    {$IFDEF EAGLCLIENT}
    TLig          : Tob ;
    {$ENDIF}
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    function  SelectQL : String ;
    procedure ReInitGridTiers ;
    procedure ReInitGridPieces ;
    procedure EnableBack(Flag : Boolean) ;
    procedure AfficheGridTiers ;
    procedure AfficheGridPieces(Aux : String) ;
    procedure TraiteAuxiliaire(Auxi : String  ; Var fc_aux, ce_aux, cd_aux, cr_aux : Double) ;
    procedure TraitePiece(Auxi,Journal,Exercice : String ; NumeroPiece : Integer ; Var fc, ce, cd, cr : Double) ;
    procedure ChercherEcritures(Aux : String) ;
    function  TrouveLigneAuxiliaire(Auxiliaire : String ; Cout : Double) : Integer ;
    procedure DrawGraph(NumGraph : Integer) ;
    procedure InsereLigneAuxiliaire(lig : Integer) ;
    function  AjouteLigneAuxiliaire(Position : boolean) : Integer ;
    function  StrStat(d : Double ; percent : boolean) : String ;
    procedure ToucheEscape ;
    procedure AffichePanelPieces ;
    procedure AffichePanelStats ;
    procedure EndProcess ;
    function  CopyGrid(G1 : THGrid) : THGrid ;
    {$IFDEF EAGLCLIENT}       
    function  RenseigneCriteres : String ;
    Procedure AlimLaFille(LeNom, LaValeur : String);
    {$ENDIF}

  public
    { Déclarations publiques }
    FindFirst : Boolean ;
    GeneAuxi  : String ;
    GeneDateRef : TDateTime ;
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPTypeCons,
  {$ENDIF MODENT1}
  UtilPgi ;


// Appel de la fenêtre
Procedure LanceCoutTiers ( LeAuxi : String ; LaDateRef : TDateTime ) ;
Var X  : TFCoutTiers ;
    PP : THPanel ;
    Composants : TControlFiltre; //SG6   Gestion des Filtes 17/11/04   FQ 14976
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
X:=TFCoutTiers.Create(Application) ;

//SG6 17/11/04 Gestion des Filtres      FQ 14976
Composants.PopupF   := X.POPF;
Composants.Filtres  := X.FFILTRES;
Composants.Filtre   := X.BFILTRE;
Composants.PageCtrl := X.Pages;
X.ObjFiltre := TObjFiltre.Create(Composants, '');

X.GeneAuxi:=LeAuxi ; X.GeneDateRef:=LaDateRef ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
    End ;
   Screen.Cursor:=SynCrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

function EcartTypeGrid(G : THGrid ; col : Integer) : Extended ;
Var lig,n    : Integer ;
    x,sx,sx2 : Double ;
BEGIN
sx := 0 ; sx2 := 0 ;
n := G.RowCount-1 ;
if n<=1 then BEGIN result:=0 ; exit ; END ;
for lig:=1 to G.RowCount-1 do
   BEGIN
   x := Valeur(G.Cells[col,lig]) ;
   sx := sx + x ;
   sx2 := sx2 + Sqr(x) ;
   END ;
result := Sqrt(((n * sx2) - sqr(sx)) / (n * (n-1))) ;
END ;

function SommeGrid(G : THGrid ; col : Integer) : Extended ;
Var lig : Integer ;
BEGIN
result := 0 ;
for lig:=1 to G.RowCount-1 do result := result + Valeur(G.Cells[col,lig]) ;
END ;

function MoyenneGrid(G : THGrid ; col : Integer) : Extended ;
BEGIN
if G.RowCount>0 then Result := SommeGrid(G,col) / (G.RowCount-1) else Result:=0 ;
END ;

function TFCoutTiers.SelectQL : String ;
BEGIN
SelectQL:='Select Ecriture.*, '+GetChampsTiers(ctCoutTiers)+' from Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ' ;
END ;

procedure TFCoutTiers.AfficheGridPieces(Aux : String) ;
Var AJournal, AExercice          : String ;
    ANumeroPiece                 : Integer ;
    lig                          : Integer ;
    OBM                          : TOBM ;
    fc, ce, cd, cr, tc, rt       : Double ;
    fc_total, tc_total, rt_total : Double ;
BEGIN
ReInitGridPieces ;
AJournal := '' ; AExercice := '' ; ANumeroPiece := 0 ; lig := 0 ;
fc_total := 0 ; rt_total := 0 ; tc_total := 0 ;

if Not QEcr.Eof then
{$IFDEF EAGLCLIENT}
  InitMove(QEcr.RecordCount, '') ;
  TobP.Detail.Clear ;
{$ELSE}
  InitMove(RecordsCount(QEcr),'') ;
{$ENDIF}

While Not QEcr.EOF do
   BEGIN
   if (QEcr.FindField('E_JOURNAL').AsString <> AJournal) Or
      (QEcr.FindField('E_EXERCICE').AsString <> AExercice) Or
      (QEcr.FindField('E_NUMEROPIECE').AsInteger <> ANumeroPiece) then
      BEGIN
      AJournal     := QEcr.FindField('E_JOURNAL').AsString ;
      AExercice    := QEcr.FindField('E_EXERCICE').AsString ;
      ANumeroPiece := QEcr.FindField('E_NUMEROPIECE').AsInteger ;
      TraitePiece(Aux,AJournal, AExercice, ANumeroPiece, fc, ce, cd, cr) ;
      tc := ce + cd + cr ;
      if fc <> 0 then rt := (tc / fc) * 100
                 else rt := 0 ;
      fc_total := fc_total + fc ;
      tc_total := tc_total + tc ;
      if fc_total <> 0 then rt_total := (tc_total / fc_total) * 100
                       else rt_total := 0 ;
      if (fc<>0) then
         BEGIN
         Inc(lig) ;
         if GP.RowCount - 1 < lig then GP.RowCount := lig + 1 ;
         GP.Cells[0,lig] := IntToStr(ANumeroPiece) ;
         GP.Cells[1,lig] := StrS0(fc) ;
         GP.Cells[2,lig] := StrS0(ce) ;
         GP.Cells[3,lig] := StrS0(cd) ;
         GP.Cells[4,lig] := StrS0(cr) ;
         GP.Cells[5,lig] := StrS0(tc) ;
         GP.Cells[6,lig] := StrS0(rt) ;

         {$IFDEF EAGLCLIENT}
         TLig := Tob.Create('$Fille', TobP, -1);

         AlimLaFille('num', IntToStr(ANumeroPiece));
         AlimLaFille('fc', StrS0(fc));
         AlimLaFille('ce', StrS0(ce));
         AlimLaFille('cd', StrS0(cd));
         AlimLaFille('cr', StrS0(cr));
         AlimLaFille('tc', StrS0(tc));
         AlimLaFille('rt', StrS0(rt));
         {$ENDIF}

         // Création de l'objet OBM pour ZoomPiece
         OBM:=TOBM.Create(EcrGen,'',True) ;
         OBM.PutMvt('E_EXERCICE',QEcr.FindField('E_EXERCICE').AsString) ;
         OBM.PutMvt('E_JOURNAL',QEcr.FindField('E_JOURNAL').AsString) ;
         OBM.PutMvt('E_NUMEROPIECE',QEcr.FindField('E_NUMEROPIECE').AsInteger) ;
         OBM.PutMvt('E_DATECOMPTABLE',QEcr.FindField('E_DATECOMPTABLE').AsDateTime) ;
         GP.Objects[0,lig]:=OBM ;
         END ;
      END ;
   MoveCur(False) ;
   QEcr.Next ;
   END ;
FiniMove ;
TotalGeneralP.Caption := StrS0(tc_total) + ' ' ;
RatioMoyenP.Caption := '(' + StrS0(rt_total) + ' %)' ;
TitrePanelPieces.Caption := MsgBox.Mess[11] + ' ' + GT.Cells[0,GT.Row] + ' - ' + GT.Cells[1,GT.Row] ;
AffichePanelPieces ;
END ;

procedure TFCoutTiers.AffichePanelPieces ;
BEGIN
PanelPieces.Left := (ClientWidth - PanelPieces.Width) div 2 ;
PanelPieces.Top  := (ClientHeight - PanelPieces.Height) div 2 ;
PanelPieces.Visible := True ;
PanelPieces.BringToFront;           {FP 09/11/2005 FQ16952}
if GP.CanFocus then GP.SetFocus ;
EnableBack(False) ;
END ;

procedure TFCoutTiers.EnableBack(Flag : Boolean) ;
BEGIN
Pages.Enabled := Flag ;
PFiltres.Enabled := Flag ;
GT.Enabled := Flag ;
HPB.Enabled := Flag ;
END ;

procedure TFCoutTiers.ReInitGridPieces ;
Var c : Integer ;
BEGIN
GP.RowCount:=2 ; GP.ListeParam:='COUTPIECES' ;
For c:=0 to GP.ColCount - 1 do GP.Cells[c,1] := '' ;
END ;

procedure TFCoutTiers.ReInitGridTiers ;
Var c : Integer ;
BEGIN
GT.RowCount:=2 ; GT.ListeParam:='COUTTIERS' ;
For c:=0 to GT.ColCount - 1 do GT.Cells[c,1] := '' ;
GT.Enabled := False ;
Autres := False ;
END ;

procedure TFCoutTiers.AfficheGridTiers ;
Var Auxi, lAuxi : String ; lig : Integer ;
    fc_aux, ce_aux, cd_aux, cr_aux, tc_aux, rt_aux                   : Double ;  // auxiliaire
    fc_autres, ce_autres, cd_autres, cr_autres, tc_autres, rt_autres : Double ;  // autres
    fc_total, ce_total, cd_total, cr_total, tc_total                 : Double ;  // totaux
BEGIN
fc_total := 0 ; ce_total := 0 ; cd_total := 0 ; cr_total := 0 ;
if Not QEcr.EOF then HPB.Enabled:=True ;

if Not QEcr.Eof then
{$IFDEF EAGLCLIENT}
  InitMove(QEcr.RecordCount, '') ;
  TobT.Detail.Clear ;
{$ELSE}
  InitMove(RecordsCount(QEcr),'') ;
{$ENDIF}

Auxi := '' ;
fc_autres := 0 ; ce_autres := 0 ; cd_autres := 0 ; cr_autres := 0 ; tc_autres := 0 ; rt_autres := 0 ;

While Not QEcr.EOF do
   BEGIN
   if QEcr.FindField('E_AUXILIAIRE').AsString <> Auxi then
      BEGIN
      Auxi := QEcr.FindField('E_AUXILIAIRE').AsString ;
      lAuxi := QEcr.FindField('T_LIBELLE').AsString ;
      TraiteAuxiliaire(Auxi,fc_aux,ce_aux,cd_aux,cr_aux) ;
      if StopProcess then exit ;
      // calcul du total cout & ratio
      tc_aux := ce_aux + cd_aux + cr_aux ;
      if fc_aux <> 0 then rt_aux := (tc_aux / fc_aux) * 100 else rt_aux := 0 ;
      // gestion des seuils
      if (tc_aux < Valeur(MontantSeuil.Text)) Or
         (rt_aux < Valeur(RatioSeuil.Text)) then
         BEGIN // inférieur à l'un des 2 seuils dont cumuler dans "autres"
         fc_autres := fc_autres + fc_aux ;
         ce_autres := ce_autres + ce_aux ;
         cd_autres := cd_autres + cd_aux ;
         cr_autres := cr_autres + cr_aux ;
         tc_autres := ce_autres + cd_autres + cr_autres ;
         if fc_autres <> 0 then rt_autres := (tc_autres / fc_autres) * 100 else rt_autres := 0 ;
         END
      else
         BEGIN // Affichage de la ligne auxiliaire
         lig := TrouveLigneAuxiliaire(Auxi,tc_aux) ;
         GT.Cells[0,lig] := Auxi ;
         GT.Cells[1,lig] := lAuxi ;
         GT.Cells[2,lig] := StrS0(fc_aux) ;
         GT.Cells[3,lig] := StrS0(ce_aux) ;
         GT.Cells[4,lig] := StrS0(cd_aux) ;
         GT.Cells[5,lig] := StrS0(cr_aux) ;
         GT.Cells[6,lig] := StrS0(tc_aux) ;
         GT.Cells[7,lig] := StrS0(rt_aux) ;

         {$IFDEF EAGLCLIENT}
         TLig := Tob.Create('$Fille', TobT, -1);

         AlimLaFille('Auxi', Auxi);
         AlimLaFille('lAuxi', lAuxi);
         AlimLaFille('fc_aux', StrS0(fc_aux));
         AlimLaFille('ce_aux', StrS0(ce_aux));
         AlimLaFille('cd_aux', StrS0(cd_aux));
         AlimLaFille('cr_aux', StrS0(cr_aux));
         AlimLaFille('tc_aux', StrS0(tc_aux));
         AlimLaFille('rt_aux', StrS0(rt_aux));
         {$ENDIF}

         END ;
      // Cumul dans total général
      fc_total := fc_total + fc_aux ;
      ce_total := ce_total + ce_aux ;
      cd_total := cd_total + cd_aux ;
      cr_total := cr_total + cr_aux ;
      tc_total := ce_total + cd_total + cr_total ;
//      if fc_total <> 0 then rt_total := (tc_total / fc_total) * 100 else rt_total := 0 ;
      TotalGeneral.Caption := StrS0(tc_total) + ' ' ;
//      RatioMoyen.Caption := '('+StrS0(rt_total) + ' %)' ;
      END ;
   END ;
if (GT.RowCount = 2) and (GT.Cells[0,1]='') then
   BEGIN
   MsgBox.Execute(1,'','') ;
   HPB.Enabled:=False ;
   FiniMove ;
   exit ;
   END ;
if (fc_autres <> 0) Or (tc_autres <> 0) then // il existe une ligne "autres"
   BEGIN
   lig := TrouveLigneAuxiliaire('',tc_autres) ;
   GT.Cells[0,lig] := MsgBox.Mess[12] ;
   GT.Cells[1,lig] := MsgBox.Mess[13] + MontantSeuil.Text + MsgBox.Mess[14] + RatioSeuil.Text + ' %' ;
   GT.Cells[2,lig] := StrS0(fc_autres) ;
   GT.Cells[3,lig] := StrS0(ce_autres) ;
   GT.Cells[4,lig] := StrS0(cd_autres) ;
   GT.Cells[5,lig] := StrS0(cr_autres) ;
   GT.Cells[6,lig] := StrS0(tc_autres) ;
   GT.Cells[7,lig] := StrS0(rt_autres) ;
   Autres := True ;
   END ;
FiniMove ;
END ;

// Traite un auxiliaire
procedure TFCoutTiers.TraiteAuxiliaire(Auxi : String  ; Var fc_aux, ce_aux, cd_aux, cr_aux : Double) ;
Var Ajournal, AExercice : String ;
    ANumeroPiece : Integer ;
    fc, ce, cd, cr : Double ;
BEGIN
fc_aux := 0 ; ce_aux := 0 ; cd_aux := 0 ; cr_aux := 0 ;
AJournal := '' ; AExercice := '' ; ANumeroPiece := 0 ;
While (Not QEcr.EOF) and (QEcr.FindField('E_AUXILIAIRE').AsString = Auxi) do
   BEGIN
   Application.ProcessMessages ;
   if StopProcess then exit ;
   if (QEcr.FindField('E_JOURNAL').AsString <> AJournal) Or
      (QEcr.FindField('E_EXERCICE').AsString <> AExercice) Or
      (QEcr.FindField('E_NUMEROPIECE').AsInteger <> ANumeroPiece) then
      BEGIN
      AJournal     := QEcr.FindField('E_JOURNAL').AsString ;
      AExercice    := QEcr.FindField('E_EXERCICE').AsString ;
      ANumeroPiece := QEcr.FindField('E_NUMEROPIECE').AsInteger ;
      TraitePiece(Auxi,AJournal, AExercice, ANumeroPiece, fc, ce, cd, cr) ;
      if StopProcess then exit ;
      fc_aux := fc_aux + fc ;
      ce_aux := ce_aux + ce ;
      cd_aux := cd_aux + cd ;
      cr_aux := cr_aux + cr ;
      END ;
   MoveCur(False) ;
   QEcr.Next ;
   END ;
END ;

// Traite une pièce
procedure TFCoutTiers.TraitePiece(Auxi,Journal,Exercice : String ; NumeroPiece : Integer ; Var fc, ce, cd, cr : Double) ;
Var
    Q      : TQuery ;
    m,mc,n : Double ;
    d      : TDateTime ;
BEGIN
    Q := OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL = "' + Journal + '" AND E_EXERCICE = "' + Exercice + '" AND E_QUALIFPIECE="N"  AND E_NUMEROPIECE = ' + IntToStr(NumeroPiece),True) ;

    fc := 0;
    ce := 0;
    cd := 0;
    cr := 0;

    While Not Q.EOF do
    BEGIN
        // BPY le 09/09/2004 : fiche n° 12260 => piece multi tiers ....
        if (Q.FindField('E_ECHE').AsString = 'X') and ((Q.FindField('E_AUXILIAIRE').AsString = Auxi) or (Auxi = '')) then // Echéance
        BEGIN
            m := 0 ;
            // Facturé
            if Q.FindField('E_NATUREPIECE').AsString = 'FC' then m := Q.FindField('E_DEBIT').AsFloat;
            fc := Arrondi(fc+m,4) ;
            // Délai normal
            n := m * (Q.FindField('E_DATEECHEANCE').AsDateTime - Q.FindField('E_DATECOMPTABLE').AsDateTime);
            if n < 0 then n:= 0;
            cd := cd + (n * (TauxNormal.Value / 36000));
            // Retard
            if (Q.FindField('E_ETATLETTRAGE').AsString = 'PL') Or (Q.FindField('E_ETATLETTRAGE').AsString = 'TL') And (Q.FindField('E_DATEPAQUETMAX').AsDateTime <> 0) then d := Q.FindField('E_DATEPAQUETMAX').AsDateTime
            else d := DateRef ;
            if Q.FindField('E_ETATLETTRAGE').AsString <> 'TL' then mc := Q.FindField('E_COUVERTURE').AsFloat else mc := 0;
            n  := (m - mc) * (d - Q.FindField('E_DATEECHEANCE').AsDateTime);
            if n < 0 then n:= 0;
            cr := cr + (n * (TauxRetard.Value / 36000));
        END ;
        // Escompte
        if Copy(Q.FindField('E_GENERAL').AsString,1,Length(RacineEscompte.Text)) = RacineEscompte.Text then
        BEGIN
            // Facture
            if Q.FindField('E_NATUREPIECE').AsString = 'FC' then ce := ce + Q.FindField('E_DEBIT').AsFloat - Q.FindField('E_CREDIT').AsFloat ;
            // Avoirs
            if (Q.FindField('E_NATUREPIECE').AsString = 'AC') And (cAvoirs.Checked) then ce := ce - Q.FindField('E_CREDIT').AsFloat + Q.FindField('E_DEBIT').AsFloat ;
        END ;
        Q.Next
    END ;
    Ferme(Q) ;
END ;

// Trouve la ligne correspondant à un compte auxiliaire
function TFCoutTiers.TrouveLigneAuxiliaire(Auxiliaire : String ; Cout : Double) : Integer ;
Var l,li : Integer ;
BEGIN
result := 0 ;
if (GT.RowCount = 2) And (GT.Cells[0,1]='') then BEGIN result := 1 ; exit ; END ;
if Auxiliaire = '' then BEGIN result := AjouteLigneAuxiliaire(True) ; exit ; END ;
For l := 1 to GT.RowCount - 1 do
   BEGIN
   if GT.Cells[0,l] = Auxiliaire then BEGIN result := l ; END ;
   END ;

if result = 0 then // nouvelle ligne
   BEGIN

   result:=AjouteLigneAuxiliaire(True) ;
   exit ;

   // Tri par insertion par coût total décroissant
   li := 0 ;
   For l := 1 to GT.RowCount - 1 do if Valeur(GT.Cells[6,l]) < Cout then BEGIN li := l ; break ; END ;
   if li = 0 then BEGIN result := AjouteLigneAuxiliaire(True) ; exit ; END ;
   InsereLigneAuxiliaire(li) ; result := li ;
   END ;
END ;

procedure TFCoutTiers.InsereLigneAuxiliaire(lig : Integer) ;
Var c,l : Integer ;
BEGIN
AjouteLigneAuxiliaire(False) ;
For l:=GT.RowCount-2 Downto lig do For c:=0 to GT.ColCount-1 do GT.Cells[c,l+1]:=GT.Cells[c,l] ;
GT.Row := lig ;
END ;

function TFCoutTiers.AjouteLigneAuxiliaire(Position : boolean) : Integer ;
BEGIN
GT.RowCount := GT.RowCount + 1 ;
if Position then GT.Row := GT.RowCount - 1 ;
result := GT.RowCount - 1 ;
END ;

procedure TFCoutTiers.ChercherEcritures(Aux : String) ;
Var StV8, StSQL : String ;
BEGIN

{$IFDEF EAGLCLIENT}
QEcr:=TQuery.Create('ECRITURE',nil,0) ;
{$ELSE}
QEcr:=TQuery.Create(Application) ;
QEcr.DataBaseName:=DBSOC.DataBaseName ;
{$ENDIF}

StSQL:=SelectQL ;

StSQL:=StSQL+RecupWhereCritere(Pages) ;
StV8:=LWhereV8 ; if StV8<>'' then StSQL:=StSQL+' AND '+StV8 ;
StSQL:=StSQL+' AND (T_NATUREAUXI = "CLI" OR T_NATUREAUXI = "AUD")' ;
StSQL:=StSQL+' AND (E_NATUREPIECE = "FC"' ;
if cAvoirs.Checked then StSQL:=StSQL+' OR E_NATUREPIECE = "AC")' else StSQL:=StSQL+') AND E_DEBIT>0' ;
if Aux = '' then StSQL:=StSQL+' AND E_AUXILIAIRE <> ""'
            else StSQL:=StSQL+' AND E_AUXILIAIRE = "' + Aux + '"' ;
Case rSelection.ItemIndex Of
   1 : StSQL:=StSQL+' AND (E_ETATLETTRAGE = "AL" OR E_ETATLETTRAGE = "PL")' ;
   2 : StSQL:=StSQL+' AND (E_ETATLETTRAGE = "AL" OR E_ETATLETTRAGE = "PL") AND E_DATEECHEANCE < "' + USDateTime(DateRef) + '"' ;
   END ;
StSQL:=StSQL+' ORDER BY E_AUXILIAIRE, E_EXERCICE, E_JOURNAL, E_NUMEROPIECE' ;

QEcr:=OpenSQL(StSQL,true) ;

END ;

procedure TFCoutTiers.DrawGraph(NumGraph : Integer) ;
Var GTitres : TStrings ;
    lig     : Integer ;
BEGIN
GTitres := TStringList.Create ;
GTitres.Add('*') ; // pour être en phase avec l'indice grid
for lig:=1 to GT.RowCount - 1 do GTitres.Add(GT.Cells[0,lig]) ;
Case NumGraph Of
   1 : VisuGraph('COUTTIERS1',BGraph1.Hint,GT,GTitres,0,1,GT.RowCount-1,False,[7],Nil,Nil) ;
   2 : VisuGraph('COUTTIERS2',BGraph2.Hint,GT,GTitres,0,1,GT.RowCount-1,False,[6],Nil,Nil) ;
   END ;
GTitres.Free ;
END ;

function  TFCoutTiers.StrStat(d : Double ; percent : boolean) : String ;
BEGIN
result := StrfMontant(d,15,V_PGI.OkDecV,'',TRUE) ;
if percent then result := result + ' %' ;
END ;

procedure TFCoutTiers.ToucheEscape ;
BEGIN
if PanelPieces.Visible then bFermerPClick(Nil) ;
if PanelStats.visible then bFermerSClick(Nil) ;
END ;

{-------------------------------- Evenements de la form -------------------------------}

procedure TFCoutTiers.BAnnulerClick(Sender: TObject);
begin
  Close ;
  //SG6 04.03.05 FQ 15441
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
end;

procedure TFCoutTiers.BZoomTiersClick(Sender: TObject);
Var Lig  : integer ;
    Auxi : String ;
begin
Lig:=GT.Row ; if ((Lig<=0) or (Lig>GT.RowCount-1)) then Exit ;
Auxi:=GT.Cells[0,Lig] ;
if Auxi<>'' then
   FicheTiers(Nil,'',Auxi,taConsult,1) ;
end;

procedure TFCoutTiers.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_ESCAPE : ToucheEscape ;
          80 : if Shift=[ssCtrl] then if bImprimer.Enabled then BImprimerClick(Nil) ;
   VK_F9 : BChercherClick(nil) //SG6 17/11/04 Gestion des filtres FQ 14976
   END ;
end;

procedure TFCoutTiers.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FreeAndNil(ObjFiltre); //SG 17/11/04 Gestion des filtres FQ 14976
GT.VidePile(True) ; GP.VidePile(True) ;
if Assigned( TobT ) then FreeAndNil(TobT) ;
if Assigned( TobS ) then FreeAndNil(TobS) ;
if Assigned( TobP ) then FreeAndNil(TobP) ;

end;

procedure TFCoutTiers.FormShow(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
TobT := Tob.Create('$Maman', nil, -1);
TobS := Tob.Create('$Maman', nil, -1);
TobP := Tob.Create('$Maman', nil, -1);
{$ENDIF}
E_DateEcheance.Text := StDate1900 ; E_DateEcheance_.Text := StDate2099 ;
E_DateComptable.Text := StDate1900 ; E_DateComptable_.Text := StDate2099 ;
DateReference.Text := DateToStr(V_PGI.DateEntree) ; AppelAutre:=False ;

ObjFiltre.FFI_TABLE:='COUTTIERS'; //SG6 17/11/04 Gestion des filtres FQ 14976
ObjFiltre.Charger;

InitTablesLibresTiers(PLibres) ;
ReInitGridTiers ;
TauxNormal.Value:=VH^.TauxCoutTiers ; TauxRetard.Value:=VH^.TauxCoutTiers ;
if (GeneAuxi <> '') Or (GeneDateRef <> 0) then // Appel depuis une autre fonction
   BEGIN
   AppelAutre:=True ;
   if GeneAuxi <> '' then BEGIN E_AUXILIAIRE_SUP.Text := GeneAuxi ; E_AUXILIAIRE_INF.Text := GeneAuxi ; END ;
   if GeneDateRef <> 0 then DateReference.Text := DateToStr(GeneDateRef) ;
   bChercherClick(Nil) ;
   END ;

end;

procedure TFCoutTiers.BChercherClick(Sender: TObject);
begin
if ((Valeur(TauxNormal.Text)=0) And
   (Valeur(TauxRetard.Text)=0)) Or
   (Trim(RacineEscompte.Text)='') Or
   (Length(Trim(RacineEscompte.Text))<3) Or
   (StrToDate(DateReference.Text)=0) then if Not AppelAutre then
   BEGIN
   MsgBox.Execute(0,'','') ;
   exit ;
   END ;
HEnt1.EnableControls(Self,False) ;
//TotalGeneral.Visible := False ;
bStop.Visible := True ;
//bChercher.Enabled := False ;
StopProcess := False ;
DateRef:=StrToDate(DateReference.Text) ;
ReInitGridTiers ;
ChercherEcritures('') ;
AfficheGridTiers ;

ferme(QEcr) ;

bStop.Visible := False ;
bChercher.Enabled := True ;
GT.TopRow := 1 ; GT.Row :=1 ;
GT.Enabled := True ;
if Autres then GT.SortRowExclude := 1 else GT.SortRowExclude := 0 ;
GT.SortGrid(6,True) ;
if StopProcess then
   BEGIN
   FiniMove ;
   END
else
   BEGIN
   EndProcess ;
   GT.SetFocus ;
   END ;
END ;

procedure TFCoutTiers.EndProcess ;
BEGIN
HEnt1.EnableControls(Self,True) ;
BAgrandir.Enabled := True ;
BRecherche.Enabled := True ;
BMenuZoom.Enabled := True ;
BGraph.Enabled := True ;
BStats.Enabled := True ;
BDetail.Enabled := True ;
BImprimer.Enabled := True ;
END ;

procedure TFCoutTiers.POPSPopup(Sender: TObject);
begin
if ((PanelStats.Visible) or (PanelPieces.Visible)) then Exit ; 
InitPopup(Self) ;
end;

procedure TFCoutTiers.FindMvtFind(Sender: TObject);
begin
Rechercher(GT,FindMvt,FindFirst) ;
end;

procedure TFCoutTiers.BRechercheClick(Sender: TObject);
begin
FindFirst:=True ; FindMvt.Execute ;
end;

procedure TFCoutTiers.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
if Assigned(TobT) and (TobT.Detail.count > 0) then
    LanceEtatTob('E','EST','LST',TobT,True,false,False,nil,'','Estimation des coûts financiers',False,0,RenseigneCriteres,0);
{$ELSE}
PrintDbGrid(GT,Pages,Caption,'');
{$ENDIF}

end;

procedure TFCoutTiers.E_DATEECHEANCEKeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFCoutTiers.BDetailClick(Sender: TObject);
begin
// Valide (affiche les pièces d'un tiers)
if ((GT.RowCount = 2) And (GT.Cells[0,1]='')) Or (GT.Cells[0,GT.Row]=MsgBox.Mess[12])then Exit ;
ChercherEcritures(GT.Cells[0,GT.Row]) ;
AfficheGridPieces(GT.Cells[0,GT.Row]) ;

ferme(QEcr) ;

end;

procedure TFCoutTiers.BStopClick(Sender: TObject);
begin
StopProcess := True ;
EndProcess ;
end;

procedure TFCoutTiers.bFermerPClick(Sender: TObject);
begin
EnableBack(True) ;
PanelPieces.Visible := False ;
end;

procedure TFCoutTiers.GTDblClick(Sender: TObject);
begin
if bDetail.Enabled then bDetailClick(Nil) ;
end;

procedure TFCoutTiers.bZoomMvtsClick(Sender: TObject);
Var
StSQL, StV8, StCri : String ;
begin
StSQL:=StSQL+RecupWhereCritere(Pages) ;
StV8:=LWhereV8 ; if StV8<>'' then StSQL:=StSQL+' AND '+StV8 ;
StSQL:=StSQL+' AND (T_NATUREAUXI = "CLI" OR T_NATUREAUXI = "AUD")' ;
StSQL:=StSQL+' AND E_AUXILIAIRE = "'+GT.Cells[0,GT.Row]+'"' ;
StSQL:=StSQL+' AND (E_NATUREPIECE = "FC"' ;
if cAvoirs.Checked then StSQL:=StSQL+' OR E_NATUREPIECE = "AC")' else StSQL:=StSQL+') AND E_DEBIT>0' ;

Case rSelection.ItemIndex Of
   1 : StSQL:=StSQL+' AND (E_ETATLETTRAGE = "AL" OR E_ETATLETTRAGE = "PL")' ;
   2 : StSQL:=StSQL+' AND (E_ETATLETTRAGE = "AL" OR E_ETATLETTRAGE = "PL") AND E_DATEECHEANCE < "' + USDateTime(DateRef) + '"' ;
   END ;
StSQL:=StSQL+' ORDER BY E_AUXILIAIRE, E_EXERCICE, E_JOURNAL, E_NUMEROPIECE' ;

StSQL:=StringReplace(StSQL, 'WHERE', '', []);

// Critères
StCri:= 'EAUXILIAIRE=' + GT.Cells[0,GT.Row] + '`EAUXILIAIRE_=' + GT.Cells[0,GT.Row]
      + '`E_NUMEROPIECE=' + E_NUMEROPIECE.Text + '`E_NUMEROPIECE_=' + E_NUMEROPIECE_.Text
      + '`EDATECOMPTABLE=' + E_DATECOMPTABLE.Text + '`EDATECOMPTABLE_=' + E_DATECOMPTABLE_.Text ;

LanceEtat('E','GLA','GLA',True,False,False,Nil,StSQL,'',False,0,StCri);
end;

procedure TFCoutTiers.BAgrandirClick(Sender: TObject);
begin
ChangeListeCrit(Self,True) ;
end;

procedure TFCoutTiers.BReduireClick(Sender: TObject);
begin
ChangeListeCrit(Self,False) ;
end;

procedure TFCoutTiers.bValidePClick(Sender: TObject);
Var M   : RMvt ;
    OBM : TOBM ;
    Q   : TQuery ;
    d   : TDateTime ;
BEGIN
OBM:=TOBM(GP.Objects[0,GP.Row]) ; if OBM=Nil then Exit ;
d:=OBM.GetMvt('E_DATECOMPTABLE') ;
Q:=OpenSQL('SELECT ' + SQLForIdent( fbGene ) + ' FROM ECRITURE WHERE E_JOURNAL="'+OBM.GetMvt('E_JOURNAL')+'"'
          +' AND E_EXERCICE="'+OBM.GetMvt('E_EXERCICE')+'"'
          +' AND E_QUALIFPIECE="N" '
          +' AND E_DATECOMPTABLE="'+USDateTime(d)+'"'
          +' AND E_NUMEROPIECE='+IntToStr(OBM.GetMvt('E_NUMEROPIECE')),True) ;
if Q.Eof then BEGIN Ferme(Q) ; exit ; END ;
M:=MvtToIdent(Q,fbGene,False) ; Ferme(Q) ;
if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>'')) then
  LanceSaisieFolioOBM(OBM,taConsult) else
LanceSaisie(Nil,taConsult,M) ;
END ;

procedure TFCoutTiers.GPDblClick(Sender: TObject);
begin
if bValideP.Enabled then bValidePClick(Nil) ;
end;

procedure TFCoutTiers.bImprimerPClick(Sender: TObject);
var
  G : THGrid ;
begin
G:=CopyGrid(GP) ;
if G<>nil then

{$IFDEF EAGLCLIENT}
if Assigned(TobP) and (TobP.Detail.count > 0) then
    LanceEtatTob('E','EST','EST',TobP,True,false,False,nil,'','Estimation des coûts financiers',False,0,RenseigneCriteres,0);
{$ELSE}
PrintDBGrid(G,Pages,Caption,'');
{$ENDIF}

G.VidePile(False) ; G.Free ;
end;

procedure TFCoutTiers.DateReferenceExit(Sender: TObject);
begin
if Not IsValidDate(DateReference.Text) then if DateReference.CanFocus then DateReference.SetFocus ;
end;

{------------------------------------ Graphiques -------------------------------------}

procedure TFCoutTiers.bGraph3Click(Sender: TObject);
begin DrawGraph(3) ; end ;

procedure TFCoutTiers.bGraph2Click(Sender: TObject);
begin DrawGraph(2) ; end ;

procedure TFCoutTiers.bGraph1Click(Sender: TObject);
begin DrawGraph(1) ; end ;

procedure TFCoutTiers.MontantSeuilEnter(Sender: TObject);
begin
MontantSeuil.SelectAll ;
end;

procedure TFCoutTiers.TauxNormalEnter(Sender: TObject);
begin
TauxNormal.SelectAll ;
end;

procedure TFCoutTiers.TauxRetardEnter(Sender: TObject);
begin
TauxRetard.SelectAll ;
end;

procedure TFCoutTiers.RatioSeuilEnter(Sender: TObject);
begin
RatioSeuil.SelectAll ;
end;

procedure TFCoutTiers.bStatsClick(Sender: TObject);
Var i : Integer ;
begin
{$IFDEF EAGLCLIENT}
TobS.Detail.Clear ;
{$ENDIF}

if (GT.RowCount = 2) And (GT.Cells[0,1]='') then Exit ;
For i:=1 to 4 do GS.Cells[i,0] := MsgBox.Mess[i+1] ;
For i:=1 to 5 do GS.Cells[0,i] := MsgBox.Mess[i+5] ;

NbTiers.Caption := IntToStr(GT.RowCount - 1) ;  // Nb de tiers (!!! tenir compte de Autres)

GS.Cells[1,1] := StrStat(SommeGrid(GT,2),False) ;    // total facturé
GS.Cells[2,1] := '-' ;                               // poids facturé
GS.Cells[3,1] := StrStat(MoyenneGrid(GT,2),False) ;  // Moyenne facturé
GS.Cells[4,1] := '-' ;

GS.Cells[1,2] := StrStat(SommeGrid(GT,6),False) ;     // total couts
GS.Cells[2,2] := StrStat(Valeur(GS.Cells[1,2]) / Valeur(GS.Cells[1,1]) * 100, True) ; // poids coût
GS.Cells[3,2] := StrStat(MoyenneGrid(GT,6),False) ;   // moyenne couts
GS.Cells[4,2] := StrStat(EcartTypeGrid(GT,6),False) ; // Ecart Type

GS.Cells[1,3] := StrStat(SommeGrid(GT,3),False) ;     // total escompte
GS.Cells[2,3] := StrStat(Valeur(GS.Cells[1,3]) / Valeur(GS.Cells[1,1]) * 100,True) ; // poids escompte
GS.Cells[3,3] := StrStat(MoyenneGrid(GT,3),False) ;   // moyenne escompte
GS.Cells[4,3] := StrStat(EcartTypeGrid(GT,3),False) ; // Ecart Type

GS.Cells[1,4] := StrStat(SommeGrid(GT,4),False) ;     // total délai
GS.Cells[2,4] := StrStat(Valeur(GS.Cells[1,4]) / Valeur(GS.Cells[1,1]) * 100,True) ; // poids délai
GS.Cells[3,4] := StrStat(MoyenneGrid(GT,4),False) ;   // moyenne délai
GS.Cells[4,4] := StrStat(EcartTypeGrid(GT,4),False) ; // Ecart Type

GS.Cells[1,5] := StrStat(SommeGrid(GT,5),False) ;     // total retard
GS.Cells[2,5] := StrStat(Valeur(GS.Cells[1,5]) / Valeur(GS.Cells[1,1]) * 100,True) ; // poids retard
GS.Cells[3,5] := StrStat(MoyenneGrid(GT,5),False) ;   // moyenne retard
GS.Cells[4,5] := StrStat(EcartTypeGrid(GT,5),False) ; // Ecart Type

{$IFDEF EAGLCLIENT}
 TLig := Tob.Create('$Fille', TobS, -1);
 AlimLaFille('totfac', 'GS.Cells[1,1]');
 AlimLaFille('moyfac', GS.Cells[3,1]);
 AlimLaFille('totcou', GS.Cells[1,2]);
 AlimLaFille('pctcou', GS.Cells[2,2]);
 AlimLaFille('moycou', GS.Cells[3,2]);
 AlimLaFille('ectcou', GS.Cells[4,2]);
 AlimLaFille('totesc', GS.Cells[1,3]);
 AlimLaFille('pctesc', GS.Cells[2,3]);
 AlimLaFille('moyesc', GS.Cells[3,3]);
 AlimLaFille('ectesc', GS.Cells[4,3]);
 AlimLaFille('totdel', GS.Cells[1,4]);
 AlimLaFille('pctdel', GS.Cells[2,4]);
 AlimLaFille('moydel', GS.Cells[3,4]);
 AlimLaFille('ectdel', GS.Cells[4,4]);
 AlimLaFille('totret', GS.Cells[1,5]);
 AlimLaFille('pctret', GS.Cells[2,5]);
 AlimLaFille('moyret', GS.Cells[3,5]);
 AlimLaFille('ectret', GS.Cells[4,5]);
{$ENDIF}

Pie.Clear ;
Pie.AddPie(Valeur(GS.Cells[1,3]), '', clGreen) ;    // Escompte
Pie.AddPie(Valeur(GS.Cells[1,4]), '', clYellow) ;   // Délai négocié
Pie.AddPie(Valeur(GS.Cells[1,5]), '', clRed) ;      // Retard

Pie2.Clear ;
Pie2.AddPie(Valeur(GS.Cells[1,3])/Valeur(GS.Cells[1,2])*100, '', clGreen) ;    // Escompte
Pie2.AddPie(Valeur(GS.Cells[1,4])/Valeur(GS.Cells[1,2])*100, '', clYellow) ;   // Délai négocié
Pie2.AddPie(Valeur(GS.Cells[1,5])/Valeur(GS.Cells[1,2])*100, '', clRed) ;      // Retard
AffichePanelStats ;
end;

procedure TFCoutTiers.AffichePanelStats ;
BEGIN
PanelStats.Left := (ClientWidth - PanelStats.Width) div 2 ;
PanelStats.Top := (ClientHeight - PanelStats.Height) div 2 ;
PanelStats.Visible := True ;
EnableBack(False) ;
END ;

procedure TFCoutTiers.bFermerSClick(Sender: TObject);
begin
EnableBack(True) ;
PanelStats.Visible := False ;
end;

procedure TFCoutTiers.GSDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
VAR Text      : array[0..255] of Char;
    F         : TAlignment ;
    R         : TRect ;
BEGIN
OldBrush.assign(GS.Canvas.Brush) ;
OldPen.assign(GS.Canvas.Pen) ;
StrPCopy(Text,GS.Cells[Col,Row]);
if (gdfixed in State) then
   BEGIN
   GS.Canvas.Brush.Color := clBtnFace ;
   GS.Canvas.Font.Color := clWindowText ;
   if Col = 0 then F:=taLeftJustify else F := taCenter ;
   END else
   BEGIN
   F:=taRightJustify ;
   GS.Canvas.Brush.Color :=  clWindow ;
   GS.Canvas.Font.Color := clWindowText ;
   END ;
if GS.Cells[Col,Row] <> '-' then
   BEGIN
      Case F of
        taRightJustify : ExtTextOut(GS.Canvas.Handle,Rect.Right-GS.Canvas.TextWidth(GS.Cells[Col,Row])-3,
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
        taCenter       : ExtTextOut(GS.Canvas.Handle,Rect.Left+((Rect.Right-Rect.Left-GS.canvas.TextWidth(GS.Cells[Col,Row])) div 2),
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
        else             ExtTextOut(GS.Canvas.Handle,Rect.Left+2,
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
        END ;
   END ;

if ((gdfixed in State) and GS.Ctl3D) then
   BEGIN
   DrawEdge(GS.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
   DrawEdge(GS.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
   END;
if GS.Cells[Col,Row] = '-' then
    BEGIN
    GS.Canvas.Brush.Color := clSilver ;
    GS.Canvas.Brush.Style := bsBDiagonal ;
    GS.Canvas.Pen.Color := clGrayText ;
    GS.Canvas.Pen.Mode := pmCopy ;
    GS.Canvas.Pen.Style := psClear ;
    GS.Canvas.Pen.Width:= 1 ;
    R.Left:=Rect.Left ; R.Top:=Rect.Top ; R.Right:=Rect.right+1 ; R.Bottom:=Rect.bottom+1 ;
    GS.Canvas.rectangle(R.Left,R.Top,R.Right,R.Bottom) ;
    END ;
if Col = 0 then
   BEGIN
   Case Row of
      3 : GS.Canvas.Brush.Color:=clGreen ;
      4 : GS.Canvas.Brush.Color:=clYellow ;
      5 : GS.Canvas.Brush.Color:=clRed ;
      END ;
   if (Row >=3) And (Row<=5) then
      BEGIN
      R.Left:=Rect.Left+2 ; R.Top:=Rect.Bottom-10 ;R.Right:=R.Left+4 ;R.Bottom:=R.Top+4 ;
      GS.Canvas.FillRect(R) ;
      END ;
   END ;
GS.Canvas.Brush.assign(OldBrush) ;
GS.Canvas.Pen.assign(OldPen) ;

end;

procedure TFCoutTiers.FormCreate(Sender: TObject);
begin
OldBrush:=TBrush.Create ;
OldPen:=TPen.Create ;
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFCoutTiers.FormDestroy(Sender: TObject);
begin
OldBrush.Free ;
OldPen.Free ;
end;

procedure TFCoutTiers.BitBtn2Click(Sender: TObject);
begin
GS.Height := GS.Height -1 ;
end;

procedure TFCoutTiers.bImprimerSClick(Sender: TObject);
var G : THGrid ;
begin
G:=CopyGrid(GS) ;
if G<>nil then
{$IFDEF EAGLCLIENT}
if Assigned(TobS) and (TobS.Detail.count > 0) then
    LanceEtatTob('E','EST','SST',TobS,True,false,False,nil,'','Estimation des coûts financiers',False,0,RenseigneCriteres,0);
{$ELSE}
PrintDBGrid(G,Pages,Caption,'');
{$ENDIF}
G.VidePile(False) ; G.Free ;
end;

function TFCoutTiers.CopyGrid(G1 : THGrid) : THGrid ;
var G : THGrid ;
    i,j : integer ;
BEGIN
// Permet d'éditer un grid de taille inférieure au multi-critères
G:=THGrid.Create(Application) ; G.Visible:=False ;
G.Parent:=G1.Parent ;
G.ListeParam:=G1.ListeParam ;
G.Width:=GT.Width ;
G.RowCount:=G1.RowCount ; G.ColCount:=G1.ColCount ;
G.FixedCols:=0 ;
if G1=GS then G.FixedCols:=1 ;
G.FixedRows:=1 ;
for i:= 0 to G1.RowCount-1 do
  BEGIN
  for j:=0 to G1.ColCount-1 do
    BEGIN
    G.ColWidths[j]:=G1.ColWidths[j] ;
    G.Cells[j,i]:=G1.Cells[j,i] ;
    END ;
  END ;
Result:=G ;
END ;

procedure TFCoutTiers.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCoutTiers.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;


procedure TFCoutTiers.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFCoutTiers.BGraphMouseEnter(Sender: TObject);
begin
PopZoom97(BGraph,POPZ) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 29/06/2006
Modifié le ... :   /  /
Description .. : Critéres
Mots clefs ... : CRITERES
*****************************************************************}
{$IFDEF EAGLCLIENT}
function TFCoutTiers.RenseigneCriteres : String;
var
  i : integer ;
  CC : THCpteEdit ;
  CL : THLabel ;
Begin

Result := 'E_GENERAL_INF=' + E_General_Inf.text + '`E_GENERAL_SUP=' + E_General_Sup.text
        + '`E_AUXILIAIRE_INF=' + E_Auxiliaire_Inf.text + '`E_AUXILIAIRE_SUP=' + E_Auxiliaire_Sup.text
        + '`E_NUMEROPIECE=' + E_Numeropiece.text + '`E_NUMEROPIECE_=' + E_Numeropiece_.text
        + '`E_JOURNAL=' + E_Journal.text + '`T_VILLE=' + T_Ville.text
        + '`T_LIBELLE=' + T_Libelle.text + '`T_SECTEUR=' + T_Secteur.text
        + '`T_EAN=' + T_Ean.text + '`T_CODEPOSTAL=' + T_CodePostal.text
        + '`T_ADRESSE1=' + T_Adresse1.text + '`T_TELEPHONE=' + T_Telephone.text
        + '`E_DATEECHEANCE=' + E_DateEcheance.text + '`E_DATEECHEANCE_=' + E_DateEcheance_.text
        + '`E_DATECOMPTABLE=' + E_DateComptable.text + '`E_DATECOMPTABLE_=' + E_DateComptable_.text
        + '`TAUXNORMAL=' + TauxNormal.text + '`TAUXRETARD=' + TauxRetard.text
        + '`MONTANTSEUIL=' + MontantSeuil.text + '`RATIOSEUIL=' + RatioSeuil.text
        + '`RACINEESCOMPTE=' + RacineEscompte.text + '`DATEREFERENCE=' + DateReference.text ;

for i:=0 to 9 do
begin
  CL:=THLabel(FindComponent('TT_TABLE'+IntToStr(i))) ;
  if (CL=Nil) Or (not CL.Visible) then Continue ;
  if CL.Caption<>'' then Result:=Result+'`TT_TABLE'+IntToStr(i)+'=' + CL.Caption ;

  CC:=THCpteEdit(FindComponent('T_TABLE'+IntToStr(i))) ;
  if CC=Nil then Continue ;
  if CC.Text<>'' then Result:=Result+'`T_TABLE'+IntToStr(i)+'=' + CC.text ;
end;


End;
{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 30/06/2006
Modifié le ... :   /  /
Description .. : Alimentation de la tob pour l'état
Mots clefs ... :
*****************************************************************}
Procedure TFCoutTiers.AlimLaFille(LeNom, LaValeur : String);
begin
  TLig.AddChampSup(LeNom, False);
  TLig.PutValue(LeNom, LaValeur);
end;
{$ENDIF}

end.


