{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 18/02/2003
Modifié le ... :   /  /
Description .. : Passage en EAGL                              
Mots clefs ... :
*****************************************************************}
unit VENTIL;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Forms, Dialogs, StdCtrls,  Buttons, Grids, ExtCtrls,HPanel,
{$IFDEF GCGC}
     EntGC,
{$ENDIF}
{$IFDEF EAGLCLIENT}
     LookUp,
{$ELSE}
     DB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     Hcompte,  // TGSection
     UTob,     // GetGridDetail
     Ent1, ComCtrls, HEnt1, hmsgbox,HSysMenu, Hctrls,UtilPGI ,
     ULibAnalytique,uEntCommun, TntStdCtrls, TntGrids, TntButtons
     ;

Procedure ParamVentil ( Nature,Compte,Axes : String ; Comment : TActionFiche ; Reseau : boolean );

type
  TFVentil = class(TForm)
    HPanel1: TPanel;
    Pages: TPageControl;
    PAxe1: TTabSheet;
    FListe1: THGrid;
    PAxe2: TTabSheet;
    PAxe3: TTabSheet;
    PAxe4: TTabSheet;
    PAxe5: TTabSheet;
    Panel1: TPanel;
    FTotQte21: THNumEdit;
    FTotQte11: THNumEdit;
    FTotVal1: THNumEdit;
    Panel2: TPanel;
    FTotQte22: THNumEdit;
    FTotQte12: THNumEdit;
    FTotVal2: THNumEdit;
    Panel3: TPanel;
    FTotQte23: THNumEdit;
    FTotQte13: THNumEdit;
    FTotVal3: THNumEdit;
    Panel4: TPanel;
    FTotQte24: THNumEdit;
    FTotQte14: THNumEdit;
    FTotVal4: THNumEdit;
    Panel5: TPanel;
    FTotQte25: THNumEdit;
    FTotQte15: THNumEdit;
    FTotVal5: THNumEdit;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    TCache: TLabel;
    FListe2: THGrid;
    FListe3: THGrid;
    FListe4: THGrid;
    FListe5: THGrid;
    TFType: THLabel;
    FType: THValComboBox;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    Panel6: TPanel;
    BDelLigne: THBitBtn;
    BInsLigne: THBitBtn;
    BImprimer: THBitBtn;
    OKBtn: THBitBtn;
    BFermer: THBitBtn;
    HelpBtn: THBitBtn;
    FTotalUOe5: THNumEdit;
    FTotalUOe4: THNumEdit;
    FTotalUOe3: THNumEdit;
    FTotalUOe2: THNumEdit;
    FTotalUOe1: THNumEdit;
    BCalculUOE: THBitBtn;
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListe1DblClick(Sender: TObject);
    procedure FListe1CellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure BDelLigneClick(Sender: TObject);
    procedure BInsLigneClick(Sender: TObject);
    procedure FListe1CellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure FTypeClick(Sender: TObject);
    procedure BFermerClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FListe1RowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FListe1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure BCalculUOEClick(Sender: TObject);
  private { Déclarations private }
    FNature,FCompte,FAxes : String ;
    Comment : TActionFiche ;
    CodeSectOk,FClosing : Boolean ;
    Modifier : Boolean ;
    FRestriction: TRestrictionAnalytique;        // Modèlde de restriction ana FP 29/12/2005}
    procedure Recalc (F : THGrid);
    procedure ChargeVentil (Nat,Cpte : String) ;
    Function  ChercheSect( GS : THGrid ; C,L : integer ) : byte ;
    Procedure NumeroteLigne(Sender : TObject) ;
    Function  SectionExiste(Sender :  TObject ; ARow : Integer) : Boolean ;
    Procedure InitTaux(F : THGrid ; ARow : Integer) ;
    Function  OnSauve : Boolean ;
    procedure DoSolde(Grid : THGrid); {JP 17/01/06 : FQ 17268}
    procedure DoZoom (Grid : THGrid); {JP 18/01/06}
    procedure DoTab; {JP 18/01/06}
{$IFDEF EAGLCLIENT}
    function  GenererTOBEtat ( vGrille : THGrid ) : TOB ;
{$ENDIF}
    procedure PrepareGrille ;
    function  EstModeUOE( vAxe : integer ) : Boolean ;
    function  AxeCourant( vGrille : THGrid ) : integer;
    procedure CalculDepuisUOE( vGrille : THGrid ; vBoQte : Boolean = True );
//    function  GetTotalPourc( vGrille : THGrid ) : Double ;
    function  EstUOActif : Boolean ;
    procedure GereModeAction ;
  public  { Déclarations public }
  end;

implementation   

uses
{$IFDEF MODENT1}
CPTypeCons,
{$ENDIF MODENT1}
{$IFDEF EAGLCLIENT}
  UtileAGL,      // LanceEtatTob
{$ELSE}
  PrintDBG,  // PrintDBGrid
{$ENDIF}
  UFonctionsCBP,
  CPSECTION_TOM; // FicheSection

{$R *.DFM}


Const GV_NUM  = 0 ;
      GV_SECT = 1 ;
      GV_LIB  = 2 ;
      GV_UOE  = 3 ;
      GV_POUR = 4 ;
      GV_QTE1 = 5 ;
      GV_QTE2 = 6 ;

Procedure ParamVentil ( Nature,Compte,Axes : String ; Comment : TActionFiche ; Reseau : boolean );
var FVentil: TFVentil;
BEGIN
if Reseau then if _Blocage(['nrCloture','nrBatch'],True,'nrBatch') then Exit ;
FVentil:=TFVentil.Create(Application) ;
try
  FVentil.FCompte:=Compte ;
  FVentil.FNature:=Nature ;
  FVentil.FAxes:=Axes ;
  FVentil.Comment:=Comment ;
  FVentil.ShowModal ;
  finally
  FVentil.Free ;
  if Reseau then _Bloqueur('nrBatch',False) ;
  end;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFVentil.ChargeVentil (Nat,Cpte : String) ;
Var F : THGrid ;
    i : integer ;
    sax,SQL : String ;
    QVentil : TQuery ;
    CompteAna:  array[1..MaxAxe] of String;      {FP 29/12/2005}
    {b FP 19/04/2006 FQ17725}
    QrySansFiltre:  String;
    OrderBy,ChpLibelle:String;
    {e FP 19/04/2006}
BEGIN
FillChar(CompteAna, sizeof(CompteAna), #0);       {FP 29/12/2005}
for i:=1 to MaxAxe do
  if Pos(InttoStr(i),FAxes)>0 then
    BEGIN
    F:=THGrid(FindComponent('FListe'+IntToStr(i))) ; F.VidePile(false) ;
    sax:='A'+IntToStr(i) ;
    {b FP 19/04/2006 FQ17725: Requête sans filtre sur la restriction analytique}
    if VH^.LiaisonY2ViaShare then
    begin
      QrySansFiltre := 'SELECT VENTIL.*,CSP_LIBELLE '+
                       'FROM VENTIL '+
                       'LEFT JOIN SECTION ON V_SECTION=CSP_SECTION '+
                       'WHERE V_NATURE="'+Nat+IntToStr(i)+'" AND '+
                       'V_COMPTE="'+Cpte+'" AND CSP_AXE="'+sax+'"';
      ChpLibelle := 'CSP_LIBELLE';
    end else
    begin
      QrySansFiltre := 'SELECT VENTIL.*,S_LIBELLE FROM VENTIL, SECTION WHERE VENTIL.V_NATURE="'+Nat+IntToStr(i)+'" '
          + 'AND VENTIL.V_COMPTE="'+Cpte+'" AND S_SECTION=V_SECTION AND S_AXE="'+sax+'"';
      ChpLibelle := 'S_LIBELLE';
    end;
    SQL := QrySansFiltre
       {b FP 29/12/2005}
       +' AND ' + FRestriction.GetClauseCompteAutorise(
         FCompte, sax, 'VENTIL', CompteAna);
       {e FP 29/12/2005}
    OrderBy := ' ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
    QrySansFiltre := QrySansFiltre + OrderBy;
    SQL           := SQL           + OrderBy;
    if (Nat = 'TY') and (Cpte = FCompte) then  {Paramétrage des Ventilation type}
      begin
      QVentil:=OpenSQL(QrySansFiltre,True) ;
      end
    else
      begin
      QVentil:=OpenSQL(SQL,True) ;
      FRestriction.VerifModelVentil(FCompte, QVentil, 'VENTIL', sax, QrySansFiltre, Cpte<>FCompte); {FP 29/12/2005}
      end;
    {e FP 19/04/2006}
    While Not QVentil.EOF do
       BEGIN
       F.Cells[GV_NUM,F.RowCount-1]   := QVentil.FindField('V_NUMEROVENTIL').AsString ;
       F.Cells[GV_SECT,F.RowCount-1]  := QVentil.FindField('V_SECTION').AsString ;
       F.Cells[GV_LIB,F.RowCount-1]   := QVentil.FindField(ChpLibelle).AsString ;
       // 19/04/07 Paie rajout des natures SA et PG
       if (Nat = 'TY') or (Nat = 'GE') or (Nat = 'SA') or (Nat = 'PG') then // DEV 4158 : Alimentation de l'Unité d'Oeuvre depuis V_MONTANT
         F.Cells[GV_UOE,F.RowCount-1] := StrFMontant( QVentil.FindField('V_MONTANT').AsFloat,     15, V_PGI.OkDecQ, '', True ) ;
       F.Cells[GV_POUR,F.RowCount-1]  := StrFMontant( QVentil.FindField('V_TAUXMONTANT').AsFloat, 15, 4, '', TRUE ) ;
       F.Cells[GV_QTE1,F.RowCount-1]  := StrFMontant( QVentil.FindField('V_TAUXQTE1').AsFloat,    15, 4, '', TRUE ) ;
       F.Cells[GV_QTE2,F.RowCount-1]  := StrFMontant( QVentil.FindField('V_TAUXQTE2').AsFloat,    15, 4, '', TRUE ) ;

       F.RowCount:=F.RowCount+1 ;
       QVentil.Next ;
       END ;
    Ferme(QVentil) ;
    Recalc(F) ;
    END ;
END ;

procedure TFVentil.Recalc (F : THGrid);
Var Leq,i : Integer ;
    V,Q1,Q2 : Double ;
    lTotUOe : Double ;
BEGIN
  V   := 0 ;
  Q1  := 0 ;
  Q2  := 0 ;
  Leq := StrToInt(Copy(F.Name,Length(F.Name),1)) ;
  lTotUOe := 0 ;

  For i:=1 to F.RowCount-1 do
    BEGIN
    V  := V+Valeur(F.Cells[GV_POUR,i]) ;
    Q1 := Q1+Valeur(F.Cells[GV_QTE1,i]) ;
    Q2 := Q2+Valeur(F.Cells[GV_QTE2,i]) ;
    lTotUOe := lTotUOe + Valeur(F.Cells[GV_UOE,i]) ;
    END ;

  if F.Cells[GV_SECT,F.RowCount-1]<>'' then
    F.RowCount:=F.RowCount+1 ;

  NumeroteLigne(F) ;

  THNumEdit(FindComponent('FTotVal'+IntToStr(Leq))).Value:=V ;
  THNumEdit(FindComponent('FTotQte1'+IntToStr(Leq))).Value:=Q1 ;
  THNumEdit(FindComponent('FTotQte2'+IntToStr(Leq))).Value:=Q2 ;
  THNumEdit(FindComponent('FTotalUOe'+IntToStr(Leq))).Value:=lTotUOe ;

  // ligne d'UOE saisie ?
  F.ColEditables[ GV_POUR ] := not EstModeUOE( Leq ) ;

END ;

Procedure TFVentil.InitTaux(F : THGrid ; ARow : Integer) ;
Var i : Integer ;
BEGIN
for i:=GV_UOE to GV_QTE2 do
    if F.Cells[i,ARow]='' then
       F.Cells[i,ARow]:=StrFMontant(Valeur(F.Cells[i,ARow]),15,4,'',TRUE) ;
END ;

procedure TFVentil.FormShow(Sender: TObject);
Var i : integer ;
    First : integer  ;
    QLoc : TQuery ;
begin

  // === MAJ CAPTION ===
  // clé de répartition
  if FNature='CL' then
    begin
    QLoc    := OpenSql('Select RE_LIBELLE From CLEREPAR Where RE_AXE="A'+FAxes+'" And RE_CLE="'+FCompte+'"',True) ;
    Caption := MsgBox.Mess[2]+' '+FCompte+' '+QLoc.Fields[0].AsString ; Ferme(QLoc) ;
    Ferme(QLoc) ;
    HelpContext:=1455120 ;  // sections réceptrices
    end
  // Ventil type
  else if FNature='TY' then
     begin
     QLoc:=OpenSql('Select CC_LIBELLE From CHOIXCOD Where CC_TYPE="VTY" And CC_CODE="'+FCompte+'"',True) ;
     Caption:=MsgBox.Mess[4]+' '+FCompte +' '+QLoc.Fields[0].AsString ;
     Ferme(QLoc) ;
     end
  // Autres ?
  else if FNature='HA' then
     Caption:=MsgBox.Mess[5]
  else if FNature='HV' then
     Caption:=MsgBox.Mess[6]
  else Caption:=MsgBox.Mess[3]+' '+Fcompte ;
  UpdateCaption(Self) ;
  // === FIN MAJ CAPTION ===

  // Désactivation de l'appel des ventilations types si mode saisie des ventils types...
  FType.Visible  := (FNature<>'TY') ;
  TFType.Visible := (FNature<>'TY') ;
  BCalculUOE.Visible := EstUOActif ;

  // === Détermination 1er onglet visible ===
  First:=0 ;
  For i:=1 to MaxAxe do
    begin
    if ((First=0) and (Pos(IntToStr(i),FAxes)>0)) then
      first:=i ;
    TTabSheet(FindComponent('PAxe'+IntToStr(i))).TabVisible := (Pos(IntToStr(i),FAxes)>0) ;
    end ;
  Pages.ActivePage := TTabSheet(FindComponent('PAxe'+IntToStr(First))) ;
  { FQ 19039 BVE 11.04.07 }
  if (Comment <> taConsult) and (First <> 0) then
     THGrid(FindComponent('FListe' + IntToStr(First))).SetFocus;
  { END FQ 19039}

  // === Chargement des données ===
  GereModeAction ;
  ChargeVentil(FNature,FCompte) ;
  CodeSectOk:=False ;

  // === Adaptation de la grille au différent produits / série / fonctionnel ===
  PrepareGrille ;

  // === Mode Consultation ===
  if Comment=taConsult then
    BEGIN
    FListe1.Enabled   := False ;
    FListe2.Enabled   := False ;
    FListe3.Enabled   := False ;
    FListe4.Enabled   := False ;
    FListe5.Enabled   := False ;
    FType.Enabled     := False ;
    BDelLigne.Enabled := False ;
    BInsLigne.Enabled := False ;
    OkBtn.Enabled     := False ;
    END ;

  // === Pour le mode croiseaxe, pas de gestion des ventilations type === //SG6 23.02.05
  if VH^.AnaCroisaxe and (FNature = 'CL') then
    begin
    FType.Visible := False;
    TFType.Visible := False;
    end;

end;

procedure TFVentil.OKBtnClick(Sender: TObject);
Var i,j : integer ;
    F : THGrid ;
    TOBG,TOBV : TOB ;
begin

BeginTrans ;
TOBG:=TOB.Create('LESVENTILS',Nil,-1) ;
for i:=1 to MaxAxe do
    BEGIN
    ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE="'+FNature+IntToStr(i)+'" AND V_COMPTE="'+FCompte+'" AND (V_SOUSPLAN1="" OR V_SOUSPLAN1 IS NULL) AND (V_SOUSPLAN2="" OR V_SOUSPLAN2 IS NULL) AND '+
     '(V_SOUSPLAN3="" OR V_SOUSPLAN3 IS NULL) AND (V_SOUSPLAN4="" OR V_SOUSPLAN5 IS NULL) AND (V_SOUSPLAN5="" OR V_SOUSPLAN5 IS NULL)') ;
    F:=THGrid(FindComponent('FListe'+IntToStr(i))) ;
    // Si saisie en UOE, il faut s'assurer que le total des pourcentage fait 100%
    if EstUOActif and EstModeUOE( i ) then // DEV 4158 : gestion unité d'oeuvre
//      if Arrondi( GetTotalPourc( F ) - 100.0, 4 ) <> 0 then // Si pas 100% recalcul
        CalculDepuisUOE( F, Comment=taCreat ) ;
    // Parcours de la grille
    For j:=1 to F.RowCount-1 do
      if F.Cells[GV_SECT,j]<>'' then
        BEGIN
        TOBV:=TOB.Create('VENTIL',TOBG,-1) ;
        TOBV.PutValue('V_NATURE',FNature+IntToStr(i)) ;
        TOBV.PutValue('V_COMPTE',FCompte) ;
        TOBV.PutValue('V_NUMEROVENTIL',j) ;
        TOBV.PutValue('V_SECTION',F.Cells[GV_SECT,j]) ;
        if FNature='CL' then
           BEGIN
           TOBV.PutValue('V_TAUXMONTANT',0) ;
           TOBV.PutValue('V_TAUXQTE1',0) ;
           TOBV.PutValue('V_TAUXQTE2',0) ;
           END else
           BEGIN
           if EstUOActif then // DEV 4158 : gestion unité d'oeuvre
             TOBV.PutValue('V_MONTANT',Valeur(F.Cells[GV_UOE,j])) ;
           TOBV.PutValue('V_TAUXMONTANT',Valeur(F.Cells[GV_POUR,j])) ;
           TOBV.PutValue('V_TAUXQTE1',Valeur(F.Cells[GV_QTE1,j])) ;
           TOBV.PutValue('V_TAUXQTE2',Valeur(F.Cells[GV_QTE2,j])) ;
           END ;
        END ;
    END ;
TOBG.InsertDB(Nil) ; TOBG.Free ;
CommitTrans ;
ChargeVentil(FNature,FCompte) ;
Modifier:=False ;
end;

procedure TFVentil.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FNature:='' ; FCompte:='' ; FAxes:='' ; Modifier:=False ;
FClosing:=False ;
FRestriction := TRestrictionAnalytique.Create;          {FP 29/12/2005}
end;

Function TFVentil.ChercheSect( GS : THGrid ; C,L : integer ) : byte ;
Var St : String ;
    CSect : TGSection ;
    Cache : THCpteEdit ;
    CompteAna:   array[1..MaxAxe] of String;      {FP 29/12/2005}
BEGIN
ChercheSect:=0 ;
Cache:=THCpteEdit.Create(Self) ; Cache.Visible:=False ; Cache.Libelle:=TCache ;
Cache.Parent := Self;
{b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
FillChar(CompteAna, sizeof(CompteAna), #0);
Cache.SynPlus := FRestriction.GetClauseCompteAutorise(
    FCompte, 'A'+IntToStr(Pages.ActivePage.PageIndex+1), 'SECTION', CompteAna);
Case Pages.ActivePage.PageIndex of
  0 : Cache.ZoomTable:=tzSection ;
  1 : Cache.ZoomTable:=tzSection2 ;
  2 : Cache.ZoomTable:=tzSection3 ;
  3 : Cache.ZoomTable:=tzSection4 ;
  4 : Cache.ZoomTable:=tzSection5 ;
  END ;
St:=uppercase(GS.Cells[C,L]) ; Cache.Text:=St ;
if GChercheCompte(Cache,Nil) then
   BEGIN
   if St<>Cache.Text then
      BEGIN
      GS.Cells[C,L]:=Cache.Text ;
      GS.Cells[GV_LIB,L]:=TCache.Caption ;
      ChercheSect:=1 ;
      CSect:=TGSection.Create(GS.Cells[C,L],Cache.ZoomTable) ;
      if ((CSect<>Nil) and (CSect.Ferme)) then MsgBox.Execute(1,'','') ;
      CSect.Free ;
      END else
      BEGIN
      ChercheSect:=2 ;
      END ;
   END ;
Cache.Free ;
END ;

procedure TFVentil.FListe1DblClick(Sender: TObject);
var
  F : THGrid ;
begin
  F:=THGrid(Sender) ;
  DoZoom(F); {JP 18/01/06}
end;

procedure TFVentil.FListe1CellExit(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
Var GS : THGrid ;
    iRow: Integer;   {FP 19/04/2006 FQ17726}
begin
if (ACol=GV_SECT) then
  Fliste1.Cells[ACol,ARow]:=UpperCase(Fliste1.Cells[ACol,ARow]) ;
GS:=THGrid(Sender) ;
if Not CodeSectOk then
   BEGIN
   if Not SectionExiste(Sender,ARow) then
      BEGIN
      {b FP 19/04/2006 FQ17726: Si le compte est renseigné, on interdit la possibilité de cliquer sur une autre ligne}
      if GS.Cells[GV_SECT,ARow]<>''then
        iRow := ARow
      else
        iRow := GS.Row;
      if ChercheSect(GS,1, iRow)<=0 then
      {e FP 19/04/2006 FQ17726}
         BEGIN
         Cancel:=True ; Exit ;
         END else
         BEGIN
         CodeSectOk:=True ;
         if GS.Cells[GV_SECT,ARow]<>''then InitTaux(GS,Arow) ;
         END ;
      END ;
   END ;
if ACol>GV_LIB then
  if ACol=GV_UOE
    then GS.Cells[ACol,ARow]:=StrFMontant(Valeur(GS.Cells[ACol,ARow]),15, V_PGI.OkDecQ,'',TRUE)
    else GS.Cells[ACol,ARow]:=StrFMontant(Valeur(GS.Cells[ACol,ARow]),15,4,'',TRUE) ;
// Si une UOE est saisie, il faut désactiver la saisie de la colonne suivante
if ACol=GV_UOE then
  if EstModeUOE( AxeCourant( GS ) ) then
    GS.ColEditables[ GV_POUR ] := False ;
Recalc(GS) ;
Modifier:=True ;
end;

Procedure TFVentil.NumeroteLigne(Sender : TObject) ;
Var i : Integer ;
BEGIN
for i:=1 to THGrid(Sender).RowCount-1 do
  THGrid(Sender).Cells[GV_NUM,i]:=IntToStr(i) ;
END ;

procedure TFVentil.BDelLigneClick(Sender: TObject);
Var Leq : integer ;
    F : THGrid ;
    Dessine : Boolean ;
begin
Leq:=Pages.ActivePage.PageIndex+1 ;
F:=THGrid(FindComponent('FListe'+IntToStr(Leq))) ;
Dessine:=(F.Row=12) ;
if F.Row>=1 then F.DeleteRow(F.Row) ;
if F.RowCount=1 then BEGIN F.RowCount:=F.RowCount+1 ; F.Row:=1 ; F.FixedRows:=1 ; END ;
NumeroteLigne(F) ; Recalc(F) ;
if Dessine then BEGIN F.Visible:=False ; F.Visible:=True ; F.SetFocus ;  END ;
F.Col:=1 ;
Modifier:=True ; CodeSectOk:=False ;
end;

procedure TFVentil.BInsLigneClick(Sender: TObject);
Var Leq : integer ;
    F : THGrid ;
    i : Integer ;
begin
Leq:=Pages.ActivePage.PageIndex+1 ;
F:=THGrid(FindComponent('FListe'+IntToStr(Leq))) ;
F.InsertRow(F.Row) ;
// Ligne UOE
F.Cells[GV_UOE,F.Row-1] := StrFMontant( Valeur(F.Cells[GV_UOE,F.Row-1]), 15, V_PGI.OkDecQ, '', TRUE ) ;
for i:=GV_POUR to GV_QTE2 do
  F.Cells[i,F.Row-1] := StrFMontant( Valeur(F.Cells[i,F.Row-1]), 15, 4, '', TRUE ) ;
NumeroteLigne(F) ;
F.Row:=F.RowCount-1 ;
F.Col:=GV_SECT ;
Modifier:=True ;
end;

procedure TFVentil.FListe1CellEnter(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
Var G : THGrid ;
begin
G:=THGrid(Sender) ;
if FNature='CL' then
   BEGIN
   if G.Cells[GV_SECT,ARow]<>'' then
      BEGIN
      G.Col:=GV_SECT ;
      if G.Row=ARow then G.Row:=G.Row+1 ;
      Exit ;
      END else
      BEGIN
      G.Col:=GV_SECT ; Exit ;
      END ;
   END ;
// Gestion du libellé...
if(G.Col=GV_LIB)then
   BEGIN
   if EstUOActif then
     begin
     if Acol=GV_UOE then G.Col:=GV_SECT ;
     if Acol=GV_SECT then G.Col:=GV_UOE ;
     end
   else
     begin
     if Acol=GV_POUR then G.Col:=GV_SECT ;
     if Acol=GV_SECT then G.Col:=GV_POUR ;
     end ;
   END ;
// Si Unité d'Oeuvre renseigné, la saisie du pourcentage est interdite
if EstUOActif and (G.Col=GV_POUR) then
  begin
  if (G.Cells[GV_UOE,G.Row] <> '') and ( Valeur( G.Cells[GV_UOE,G.Row] ) <> 0 ) then
    begin
    if Acol>=GV_POUR
      then G.Col:=GV_UOE
      else if Comment=taCreat then
             begin
             G.Col:=GV_SECT ;
             if G.Row=ARow then
               G.Row:=G.Row+1 ;
             end
           else G.Col:=GV_QTE1 ;
    end ;
  end ;
end;

procedure TFVentil.FTypeClick(Sender: TObject);
BEGIN
ChargeVentil('TY',FType.Value) ;
end;

procedure TFVentil.BFermerClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

procedure TFVentil.BImprimerClick(Sender: TObject);
var
  i : integer ;
{$IFDEF EAGLCLIENT}
  TobAna : Tob;
{$ENDIF}
begin
i:=Pages.ActivePageIndex;
{$IFDEF EAGLCLIENT}
  // Crée une tob avec les éléments de la grille
  Case i of
    0 : TobAna := GenererTOBEtat(FListe1);
    1 : TobAna := GenererTOBEtat(FListe2);
    2 : TobAna := GenererTOBEtat(FListe3);
    3 : TobAna := GenererTOBEtat(FListe4);
    4 : TobAna := GenererTOBEtat(FListe5);
    else TobAna := GenererTOBEtat(FListe1);
  end ;

  // Lance l'état
  LanceEtatTOB('E','CST','VAN',TobAna,True,False,False,nil,'',Caption,False);
  TobAna.Free;
{$ELSE}
Case i of
  0 :PrintDBGrid(FListe1,Nil,Caption,'') ;
  1 :PrintDBGrid(FListe2,Nil,Caption,'') ;
  2 :PrintDBGrid(FListe3,Nil,Caption,'') ;
  3 :PrintDBGrid(FListe4,Nil,Caption,'') ;
  4 :PrintDBGrid(FListe5,Nil,Caption,'') ;
  end ;
{$ENDIF}
end ;

procedure TFVentil.FListe1RowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin CodeSectOk:=False ; end;

Function TFVentil.SectionExiste(Sender :  TObject ; ARow : Integer) : Boolean ;
Var St,StC : String ;
    Q : TQuery ;
    CompteAna:   array[1..MaxAxe] of String;  {FP 29/12/2005} 
BEGIN
FillChar(CompteAna, sizeof(CompteAna), #0);  {FP 29/12/2005}
CodeSectOk:=False ; Result:=False ;
St:='A'+THGrid(Sender).Name[Length(THGrid(Sender).Name)] ;
StC:=THGrid(Sender).Cells[GV_SECT,ARow] ;
if Stc='' then Exit ;
Q:=OpenSql('Select S_SECTION From SECTION Where S_SECTION="'+StC+'" And S_AXE="'+St+'"'
{b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
  +' AND ' + FRestriction.GetClauseCompteAutorise(FCompte, St, 'SECTION', CompteAna)
{e FP 29/12/2005}
{GC_NFO_FQ;010;16036_Debut}
  +' AND S_FERME <> "X"'
{GC_NFO_FQ;010;16036_Fin}
,True) ;
Result:=Not Q.Eof ;
Ferme(Q) ; CodeSectOk:=Result ;
END ;

Function TFVentil.OnSauve : Boolean ;
Var Rep : Integer ;
BEGIN
Result:=True ;
if Modifier then
   BEGIN
   Rep:=MsgBox.Execute(0,'','') ;
   Case Rep of
      mrYes    : OKBtnClick(Nil) ;
      mrNo     : Exit ;
      mrCancel : Result:=False ;
      End ;
   END ;
END ;

procedure TFVentil.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin Canclose:=OnSauve ; end;

{---------------------------------------------------------------------------------------}
procedure TFVentil.FListe1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if not(Key in [VK_PRIOR,VK_NEXT,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT,VK_TAB]) then Modifier:=True ;
  if Sender is THGrid then begin
    case Key of
      {JP 18/01/06 : J'en profite pour brancher le F5}
      VK_F5 : if (Shift=[]) then begin
                Key := 0;
                DoZoom(THGrid(Sender));
              end;
      {JP 17/01/06 : FQ 17268 : Gestion du calcul automatique du solde}
      VK_F6 : if (Shift=[]) then begin
                Key := 0;
                DoSolde(THGrid(Sender));
              end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TFVentil.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F10 : OKBtnClick(OKBtn);
    VK_TAB : if (Shift = [ssCtrl]) and (ActiveControl is THGrid) then DoTab;
  end;
end;

{JP 17/01/06 : FQ 17268 : Gestion du calcul automatique du solde
{---------------------------------------------------------------------------------------}
procedure TFVentil.DoSolde(Grid : THGrid);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Double;
begin
  {On s'assure que l'on n'est pas en consultation, que l'on a le droit au concept et
   que la section est bien renseignée}
  if Comment = taConsult then Exit;
  if Grid.Cells[GV_SECT, Grid.Row] = '' then Exit;
  if not ExJaiLeDroitConcept(TConcept(ccSaisSolde), True) then Exit;

  if Grid.RowCount > 1 then begin
    p := 0;
    {Calcul du total de la saisie}
    for n := 1 to Grid.RowCount - 1 do
      p := p + Valeur(Grid.Cells[GV_POUR, n]);

    {Si la ligne courante n'a pas de montant, on met le reste (100 - p), dans le cas contraire
     il faut rajouter la valeur de la ligne car elle a été comptabilisée dans p}
    if Valeur(Grid.Cells[GV_POUR, Grid.Row]) = 0
      then p := 100 - p
      else p := 100 - p + Valeur(Grid.Cells[3, Grid.Row]) ;
    {Mise à jour du montant, au bon format}
    Grid.Cells[GV_POUR, Grid.Row] := StrFMontant(p, 15, 4, '', True);
  end;
end;

{18/01/06 : Pour afficher le zoom sur les sections
{---------------------------------------------------------------------------------------}
procedure TFVentil.DoZoom(Grid : THGrid);
{---------------------------------------------------------------------------------------}
var
  b : Byte;
  A : TActionFiche;
begin
  if not ExJaiLeDroitConcept(TConcept(ccSecModif),False) then A := taConsult
                                                       else A := taModif;
  if Grid.Col = GV_SECT then begin
    b := ChercheSect(Grid, Grid.Col, Grid.Row);
    if b = 2 then
      FicheSection(nil, 'A' + Copy(Grid.Name, Length(Grid.Name), 1), Grid.Cells[Grid.Col, Grid.Row], A, 0)
    else begin
      if Grid.Cells[GV_SECT, Grid.Row] <> '' then
        InitTaux(Grid, Grid.Row);
    end;
  end;
end;

{18/01/06 : Pour passer d'un Onglet à l'autre
{---------------------------------------------------------------------------------------}
procedure TFVentil.DoTab;
{---------------------------------------------------------------------------------------}
begin
  Pages.ActivePage := Pages.FindNextPage(Pages.ActivePage, True, True);
end;

procedure TFVentil.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe1.VidePile(True) ;
FListe2.VidePile(True) ;
FListe3.VidePile(True) ;
FListe4.VidePile(True) ;
FListe5.VidePile(True) ;
FClosing:=True ;
FreeAndNil(FRestriction);                   {FP 29/12/2005}
end;

procedure TFVentil.HelpBtnClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

{$IFDEF EAGLCLIENT}
function TFVentil.GenererTOBEtat(vGrille: THGrid): TOB;
var
  lInCpt    : Integer ;
  lTobLigne : TOB ;
begin
  Result := TOB.Create('GRILLE', nil, -1) ;
  for lInCpt := 1 to (vGrille.RowCount - 2) do begin
    lTobLigne := TOB.Create('_VENTILANA', Result , -1) ;

    lTobLigne.AddChampSup('V_COMPTE',      False);
    lTobLigne.AddChampSup('V_SECTION',     False);
    lTobLigne.AddChampSup('V_TAUXMONTANT', False);
    lTobLigne.AddChampSup('S_LIBELLE',     False);

    lTobLigne.PutValue( 'V_COMPTE',      FCompte);
    lTobLigne.PutValue( 'V_SECTION',     vGrille.Cells[GV_SECT, lInCpt]);
    lTobLigne.PutValue( 'V_TAUXMONTANT', vGrille.Cells[GV_POUR, lInCpt]);
    lTobLigne.PutValue( 'S_LIBELLE',     vGrille.Cells[GV_LIB, lInCpt]);
  end ;
end;
{$ENDIF}

procedure TFVentil.PrepareGrille;
var i : integer ;
    GG : THGrid ;
begin

  if FNature='CL' then
    begin
    for i := 1 to 5 do
      begin
      GG  := THGrid( FindComponent('FListe'+IntToStr(i)+'') ) ;
      GG.ColCount := 3 ;
      GG.ColWidths[ GV_SECT ] := 100 ;
      GG.ColWidths[ GV_LIb  ] := 318 ;
      end ;
    Panel1.Visible:=False ;
    end
  else
    begin
    for i := 1 to 5 do
      begin
      GG  := THGrid(FindComponent('FListe'+IntToStr(i)+'') );
      GG.ColWidths[ GV_SECT ] := 80 ;
      GG.ColWidths[ GV_LIB  ] := 200 ;
      if EstUOActif then
        begin
        GG.ColWidths[ GV_UOE  ] := 55 ;
        GG.ColWidths[ GV_POUR ] := 55 ;
        GG.ColWidths[ GV_QTE1 ] := 55 ;
        GG.ColWidths[ GV_QTE2 ] := 55 ;
        end
      else
        begin
        GG.ColLengths[GV_UOE]   := -1 ;
        GG.ColWidths[ GV_UOE  ] := 0 ;
        GG.ColWidths[ GV_POUR ] := 70 ;
        GG.ColWidths[ GV_QTE1 ] := 70 ;
        GG.ColWidths[ GV_QTE2 ] := 70 ;
        end ;
      end ;
    end ;

  // Gestion des séries ??? encore valable ?
  { 19/04/07 Paie les séries ne sont plus gérées car fonctionnel identique
   if ((EstSerie(S3)) or (EstSerie(S5))) then
     for i:=1 to 5 do
         begin
         GG := THGrid(FindComponent('Fliste'+IntToStr(i))) ;
         GG.ColLengths[GV_QTE1] := -1 ;
         GG.ColWidths [GV_QTE1] := 0 ;
         GG.ColLengths[GV_QTE2] := -1 ;
         GG.ColWidths [GV_QTE2] := 0 ;
         THnumEdit(FindComponent('FTotQte1'+IntToStr(i))).Visible:=False ;
         THnumEdit(FindComponent('FTotQte2'+IntToStr(i))).Visible:=False ;
         end ;
   }
  // Spécif par produit (GC, Paie)
{$IFDEF GCGC}
  PAxe4.TabVisible:=False ;
  PAxe5.TabVisible:=False ;
  //DP 26092003 Modif avec accord Compta
  {$IFDEF PAIEGRH}
  if EstSerie(S5) then
    PAxe3.TabVisible:=TRUE ;
  {$ELSE}
  if EstSerie(S5) then
    PAxe3.TabVisible:=VH_GC.GCventAxe3 ;
  {$ENDIF}
{$ENDIF}

  // 1 seul axe pour S3...?? encore valable ?
{  if EstSerie(S3) then
     BEGIN
     PAxe2.TabVisible:=False ;
     PAxe3.TabVisible:=False ;
     PAxe4.TabVisible:=False ;
     PAxe5.TabVisible:=False ;
     END ;
}
end;

function TFVentil.EstModeUOE(vAxe: integer): Boolean;
var i : integer ;
    F : THGrid ;
begin
  result := False ;
  if not EstUOActif then Exit ;

  // Recup Grille de l'axe
  F:=THGrid(FindComponent('FListe'+IntToStr(vAxe))) ;

  // Parcours des lignes
  For i:=1 to F.RowCount-1 do
    // Ligne valide uniquement si section renseigné
      if F.Cells[GV_SECT,i]<>'' then
        // Est-ce qu'une UOE a été saisie ?
        if ( F.Cells[GV_UOE,i]<>'' ) and ( Valeur(F.Cells[GV_UOE,i])<>0 ) then
          begin
          result := True ;
          Exit ;
          end ;

end;

function TFVentil.AxeCourant( vGrille : THGrid ) : integer;
begin
  result := StrToInt( Copy( vGrille.Name, Length(vGrille.Name), 1 ) ) ;
  // Pages.ActivePageIndex + 1 ;
end;

procedure TFVentil.CalculDepuisUOE( vGrille : THGrid ; vBoQte : Boolean = True ) ;
Var i       : integer ;
    lTotUOE : double ;
    lPourc  : double ;
    lTotP   : double ;
    lLasti  : integer ;
begin
  lTotUOe := 0 ;
  lTotP   := 0 ;
  lLasti  := 0 ;

  // Calcul du total d'Unité d'Oeuvre
  For i:=1 to vGrille.RowCount-1 do
    if vGrille.Cells[ GV_SECT, i]<>'' then
      if (vGrille.Cells[ GV_UOE, i]<>'') and ( Valeur(vGrille.Cells[ GV_UOE, i])<>0 ) then
        lTotUOE := lTotUOE + Arrondi( Valeur(vGrille.Cells[ GV_UOE, i]), V_PGI.OkDecQ ) ;

//  if lTotUOE = 0 then Exit ;

  // Parcours des lignes pour calcul
  For i:=1 to vGrille.RowCount-1 do
    if vGrille.Cells[ GV_SECT, i]<>'' then
      if (vGrille.Cells[ GV_UOE, i]<>'') then
        if ( Valeur(vGrille.Cells[ GV_UOE, i])<>0 ) then
          begin
          lPourc  := Arrondi( ( Valeur(vGrille.Cells[ GV_UOE, i]) / lTotUOE) * 100, V_PGI.OkDecP ) ;
          lTotP   := lTotP + lPourc ;
          lLasti  := i ;
          // MAJ grille
          vGrille.Cells[GV_POUR, i] := StrFMontant( lPourc, 15, 4, '', TRUE) ;
          if vBoQte then
            begin
            vGrille.Cells[GV_QTE1, i] := StrFMontant( lPourc, 15, 4, '', TRUE) ;
            vGrille.Cells[GV_QTE2, i] := StrFMontant( lPourc, 15, 4, '', TRUE) ;
            end ;
          end
        else
          begin
          // MAJ grille
          vGrille.Cells[GV_POUR, i] := StrFMontant( 0, 15, 4, '', TRUE) ;
          if vBoQte then
            begin
            vGrille.Cells[GV_QTE1, i] := StrFMontant( 0, 15, 4, '', TRUE) ;
            vGrille.Cells[GV_QTE2, i] := StrFMontant( 0, 15, 4, '', TRUE) ;
            end ;
          end ;

  // Gestion de l'arrondi
  if lTotP <> 0 then
    if Arrondi(100.0 - lTotP, 4)<>0 then
      begin
      lPourc := Valeur( vGrille.Cells[ GV_POUR, lLasti ] ) + Arrondi(100.0 - lTotP, 4) ;
      // MAJ grille
      vGrille.Cells[GV_POUR, lLasti] := StrFMontant( lPourc, 15, 4, '', TRUE) ;
      if vBoQte then
        begin
        vGrille.Cells[GV_QTE1, lLasti] := StrFMontant( lPourc, 15, 4, '', TRUE) ;
        vGrille.Cells[GV_QTE2, lLasti] := StrFMontant( lPourc, 15, 4, '', TRUE) ;
        end ;
      end ;


end;

procedure TFVentil.BCalculUOEClick(Sender: TObject);
var i : integer ;
    F : THGrid ;
begin

  if not EstUOActif then Exit ;

  for i:=1 to MaxAxe do
    if Pos(InttoStr(i),FAxes)>0 then
      begin
      F := THGrid(FindComponent('FListe'+IntToStr(i))) ;
      CalculDepuisUOE( F ) ;
      end ;

end;
(*
function TFVentil.GetTotalPourc(vGrille: THGrid): Double;
Var i     : integer ;
begin
  result   := 0 ;
  // Calcul du total des pourcentages
  For i:=1 to vGrille.RowCount-1 do
    if vGrille.Cells[ GV_SECT, i]<>'' then
      if (vGrille.Cells[ GV_POUR, i]<>'') and ( Valeur(vGrille.Cells[ GV_POUR, i])<>0 ) then
        result := result + Arrondi( Valeur(vGrille.Cells[ GV_POUR, i]), 4 ) ;
end;
*)
function TFVentil.EstUOActif: Boolean;
begin
  // 19/04/07 Paie rajout des natures SA et PG
  result := (FNature='TY') or (FNature='GE') or (FNature='SA') or (FNature='PG');
end;

procedure TFVentil.GereModeAction;
begin
  if Comment <> taModif then Exit ;
  if not EstUOActif then Exit ;
  if not ExisteSQL('SELECT V_NATURE FROM VENTIL WHERE V_NATURE LIKE "' + FNature + '%" AND V_COMPTE="'+FCompte+'"')
    then Comment := taCreat ;
end;

end.


