unit CtrRubBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, Grids, Hctrls, StdCtrls, Menus, Buttons, ExtCtrls, ComCtrls,
  Hcompte, Hqry, HEnt1, DB, DBTables, Ent1, HStatus, HSysMenu, HTB97, HPanel, UiUtil ,UObjFiltres {SG6 16/11/04 FQ 14976 Gestion des Filtres} ;

Type TInfoRub = Class
     Libelle : String ;
     Famille : String ;
     Compte1 : String ;
     Exclu1  : String ;
     Compte2 : String ;
     Exclu2  : String ;
     TabLib  : String ;
     Axe     : String ;
     End ;

Procedure ControleRubriqueBud ;

type
  TFCtrRubbu = class(TForm)
    Pages: TPageControl;
    ChoixEtat: TTabSheet;
    Bevel1: TBevel;
    LRubFam: TLabel;
    TCpte: TLabel;
    Tcpt1: TLabel;
    FamRub: THValComboBox;
    PFiltres: TToolWindow97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    FindDialog: TFindDialog;
    FListe: THGrid;
    MsgBox: THMsgBox;
    BRub: TToolbarButton97;
    BGen: TToolbarButton97;
    PopZ: TPopupMenu;
    TRub1: THLabel;
    Trub2: THLabel;
    C1: THCpteEdit;
    C2: THCpteEdit;
    Typrub: THValComboBox;
    QRub: TQuery;
    QGen: TQuery;
    QEcr: TQuery;
    CbTous: TCheckBox;
    HMTrad: THSystemMenu;
    RgChoix: TRadioGroup;
    TTypRub: TLabel;
    CbAxe: THValComboBox;
    TCbAxe: TLabel;
    TCbBud: TLabel;
    CbBud: THValComboBox;
    Rub1: THCpteEdit;
    Rub2: THCpteEdit;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    BReduire: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BRechercher: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    Dock971: TDock97;
//    procedure BCreerFiltreClick(Sender: TObject);
  //  procedure BSaveFiltreClick(Sender: TObject);
//    procedure BDelFiltreClick(Sender: TObject);
//    procedure BRenFiltreClick(Sender: TObject);
//    procedure BNouvRechClick(Sender: TObject);
//    procedure FFiltresChange(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BRubClick(Sender: TObject);
    procedure BGenClick(Sender: TObject);
    procedure C1Exit(Sender: TObject);
    procedure C2Exit(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure RgChoixClick(Sender: TObject);
    procedure TyprubChange(Sender: TObject);
    procedure CbAxeChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FamRubChange(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure CbBudChange(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    ObjFiltres:TObjFiltre; //SG6 16/11/04 Gestion des filtres FQ 14976
    FNomFiltre : String ;
    FirstFind,Stop : boolean;
    LFam,LRub1,LRub2,Lcpt1,Lcpt2 : String ;
    MemoFam : String ;
    Lefb : TFichierBase ;
    WMinX,WMinY : Integer ;
    MemoTyp : String ;
    FiltreEnCours : Boolean ;
    ListeRub : TStringList ;
    ListeGen : TStringList ;
    OnNeSortPas : Boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure InitFliste ;
    Procedure RunControle ;
    Function  ChecheLeCompte : String ;
    Procedure PositionneRequete ;
    Procedure ControleLesRubriques ;
    Procedure ControleLesComptes ;
    Function  FaitRequeteGene(Compte1,Cexlu1:String ; Unfb : TFichierBase ; SurTabLib : Boolean ) : String ;
    Function  ChercheCompteDansRub(CptRub,ComptGene : String) : Boolean ;
    Function  TestStop : Boolean ;
    Procedure ChargeTypRub ;
    Procedure FaitRequeteEcr(Unfb : TFichierBase) ;
    Procedure RempliComboFamRub ;
    Procedure SwapFamille(St : String ; TypeRubri : Boolean) ;
    Procedure QuelEstLeFb ;
    Function  QuelfbComposite : TFichierBase ;
    Procedure RempliListeRubGen(AvecGen : Boolean) ;
    Procedure VideLaListe ;
    Procedure TraiteRequeteGen(Ind : Integer ; Composite : Boolean ; Unfb : TFichierBase) ;
    Procedure BloqueControle(InRun : Boolean) ;
    Procedure ChargeRub ;
    Procedure PremierToDernier ;
  public
    { Déclarations publiques }
  end;


implementation

Uses PrintDBG,
     Filtre,
{$IFNDEF CCS3}
     Budgene,
{$ENDIF}
     Rubrique_TOM,
     CPGeneraux_TOM,
     CPTiers_TOM,
     HDB,
     Calcole,
     CPSection_TOM  ;

{$R *.DFM}

Procedure ControleRubriqueBud ;
var FCtrRub : TFCtrRubBu ;
    PP : THPanel ;

BEGIN
FCtrRub:=TFCtrRubBu.Create(Application) ;
FCtrRub.FNomFiltre:='CTRLRUBBUD' ;
FCtrRub.HelpContext:=15430000 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FCtrRub.ShowModal ;
    Finally
     FCtrRub.Free ;
    End ;
   END else
   BEGIN
   InitInside(FCtrRub,PP) ;
   FCtrRub.Show ;
   END ;
END ;

procedure TFCtrRubbu.BAnnulerClick(Sender: TObject);
begin
  //SG6 04/01/05 Vide le panel
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFCtrRubbu.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil, Caption,'') ; end;

procedure TFCtrRubbu.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFCtrRubbu.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

Procedure TFCtrRubbu.InitFliste ;
Var Typr : String ;
BEGIN
Typr:=Typrub.Value ;
FListe.ColAligns[4]:=taCenter ; FListe.ColAligns[5]:=taCenter ;
if Pos('/',Typr)<=0 then
   BEGIN
   FListe.Cells[1,0]:=MsgBox.Mess[1] ; FListe.Cells[3,0]:=MsgBox.Mess[1] ;
   if Pos('A',Typr)=1 then FListe.Cells[5,0]:=MsgBox.Mess[16]
                      else FListe.Cells[5,0]:=MsgBox.Mess[7] ;
   if RgChoix.ItemIndex=0 then
      BEGIN
      FListe.Cells[0,0]:=MsgBox.Mess[2] ;
      Case Lefb of
           fbGene : FListe.Cells[2,0]:=MsgBox.Mess[0] ;
           fbAux :  FListe.Cells[2,0]:=MsgBox.Mess[9] ;
           fbBudgen : FListe.Cells[2,0]:=MsgBox.Mess[11] ;
           fbAxe1..fbAxe5 : FListe.Cells[2,0]:=MsgBox.Mess[10] ;
         End ;
      FListe.Cells[4,0]:=MsgBox.Mess[4] ;
      END else
      BEGIN
      Case Lefb of
           fbGene : FListe.Cells[0,0]:=MsgBox.Mess[0] ;
           fbAux :  FListe.Cells[0,0]:=MsgBox.Mess[9] ;
           fbBudgen : FListe.Cells[0,0]:=MsgBox.Mess[11] ;
           fbAxe1..fbAxe5 : FListe.Cells[0,0]:=MsgBox.Mess[10] ;
         End ;
      FListe.Cells[2,0]:=MsgBox.Mess[2] ;
      END ;
   END else
   BEGIN
   FListe.Cells[0,0]:=MsgBox.Mess[2] ;
   if Typr='G/A' then
      BEGIN
      FListe.Cells[2,0]:=MsgBox.Mess[27] ;
      FListe.Cells[5,0]:=MsgBox.Mess[31] ;
      END else
      if Typr='G/T' then
         BEGIN
         FListe.Cells[2,0]:=MsgBox.Mess[28] ;
         FListe.Cells[5,0]:=MsgBox.Mess[7] ;
         END else
         if Typr='A/G' then
            BEGIN
            FListe.Cells[2,0]:=MsgBox.Mess[29] ;
            FListe.Cells[5,0]:=MsgBox.Mess[32] ;
            END else
            if Typr='T/G' then
               BEGIN
               FListe.Cells[2,0]:=MsgBox.Mess[30] ;
               FListe.Cells[5,0]:=MsgBox.Mess[7] ;
               END ;
   END ;
END ;

Function TFCtrRubbu.QuelfbComposite : TFichierBase ;
BEGIN
Result:=fbGene ;
if Pos('/',Typrub.Value)<=0 then Exit ;
Case Typrub.Value[3] of
     'A' : BEGIN
           Case CbAxe.Value[2] of
                '1' : Result:=fbAxe1 ;
                '2' : Result:=fbAxe2 ;
                '3' : Result:=fbAxe3 ;
                '4' : Result:=fbAxe4 ;
                '5' : Result:=fbAxe5 ;
             End ;
           END ;
     'G' : Result:=fbGene ;
     'T' : Result:=fbAux ;
   End ;
END ;

Procedure TFCtrRubbu.QuelEstLeFb ;
BEGIN
if Pos('A',TypRub.Value)=1 then
   BEGIN
   Case CbAxe.Value[2] of
       '1' : LeFb:=fbAxe1 ;
       '2' : LeFb:=fbAxe2 ;
       '3' : LeFb:=fbAxe3 ;
       '4' : LeFb:=fbAxe4 ;
       '5' : LeFb:=fbAxe5 ;
     End ;
   END else
   BEGIN
   if Length(TypRub.Value) = 0 then Exit;
   Case TypRub.Value[1] of
        'B' : Lefb:=fbBudgen ;
        'G' : Lefb:=fbGene ;
        'T' : Lefb:=fbAux ;
      End ;
   END ;
END ;

Procedure TFCtrRubbu.SwapFamille(St : String ; TypeRubri : Boolean) ;
Var St1 : String ;
BEGIN
if TypeRubri then
   BEGIN
   if St='CBG' then St1:='GEN' else
     if St='CBS' then St1:='ANA' else
        if St='G/S' then St1:='G/A' else
           if St='S/G' then St1:='A/G' ;
   TypRub.ItemIndex:=TypRub.Values.IndexOf(St1) ;
//   TCbAxe.Enabled:=(Pos('S',St)>0) ; CbAxe.Enabled:=(Pos('S',St)>0) ;
   END else
   BEGIN
   if St='GEN' then St1:='CBG' else
     if St='ANA' then St1:='CBS' else
        if St='G/A' then St1:='G/S' else
           if St='A/G' then St1:='S/G' ;
   Famrub.ItemIndex:=Famrub.Values.IndexOf(St1) ;
   END ;
END ;

procedure TFCtrRubbu.TyprubChange(Sender: TObject);
begin
TCbAxe.Enabled:=(Pos('A',TypRub.Value)>0) ; CbAxe.Enabled:=(Pos('A',TypRub.Value)>0) ;
if TypRub.Value='' then Exit ;
if (MemoTyp=TypRub.Value) then Exit ;
if Pos('/',TypRub.Value)>0 then if RgChoix.ItemIndex=1 then RgChoix.ItemIndex:=0 ;
MemoTyp:=TypRub.Value ;
SwapFamille(TypRub.Value,False) ;
Case TypRub.Value[1] of
     'A' : BEGIN If CbAxe.Values.Count>0 Then CbAxe.Value:=CbAxe.Values[0] ; BGen.Hint:=MsgBox.Mess[15] ; END ;
     'B' : BEGIN Lefb:=fbBudgen ; BGen.Hint:=MsgBox.Mess[14] ; C1.ZoomTable:=tzBudGen ; C2.ZoomTable:=tzBudGen ; END ;
     'G' : BEGIN Lefb:=fbGene ; BGen.Hint:=MsgBox.Mess[12] ; C1.ZoomTable:=tzGeneral ; C2.ZoomTable:=tzGeneral ; END ;
     'T' : BEGIN Lefb:=fbAux ; BGen.Hint:=MsgBox.Mess[13] ; C1.ZoomTable:=tzTiers ; C2.ZoomTable:=tzTiers ; END ;
   End ;
if TypRub.Value[1]<>'A' then BEGIN PremierToDernier ; FamRubChange(Nil) ; END ;
end;

procedure TFCtrRubbu.FamRubChange(Sender: TObject);
Var Fam : String ;
begin
  Fam:=FamRub.Value ;
  if Pos('/',Fam)>0 then if RgChoix.ItemIndex=1 then RgChoix.ItemIndex:=0 ;
  if (MemoFam=Fam)  then Exit ;
  MemoFam:=Fam ; SwapFamille(Fam,True) ; ChargeRub ;
end;

procedure TFCtrRubbu.CbBudChange(Sender: TObject);
Var Q : TQuery ;
begin
  ChargeRub ;
  Q:=OpenSQL('SELECT BJ_AXE FROM BUDJAL WHERE BJ_BUDJAL="'+cbBud.Value+'" ',TRUE) ;
  If Not Q.Eof Then cbAxe.Value:=Q.Fields[0].AsString ;
  Ferme(Q) ;
end;

Procedure TFCtrRubbu.ChargeRub ;
Var C1,C2 : String ;
BEGIN
if FamRub.Value='CBG' then BEGIN Rub1.ZoomTable:=tzRubBUDG ; Rub2.ZoomTable:=tzRubBUDG ; END else
if FamRub.Value='CBS' then BEGIN Rub1.ZoomTable:=tzRubBUDS ; Rub2.ZoomTable:=tzRubBUDS ; END else
if FamRub.Value='G/S' then BEGIN Rub1.ZoomTable:=tzRubBUDGS ; Rub2.ZoomTable:=tzRubBUDGS ; END else
if FamRub.Value='S/G' then BEGIN Rub1.ZoomTable:=tzRubBUDSG ; Rub2.ZoomTable:=tzRubBUDSG ; END ;
Rub1.SynPlus:=CbBud.Value ; Rub2.SynPlus:=CbBud.Value ;
if Not ObjFiltres.InChargement then
   BEGIN
   PremierDernierRub(Rub1.ZoomTable,Rub1.SynPlus,C1,C2) ;
   Rub1.Text:=C1 ; Rub2.Text:=C2 ;
   END ;
PremierToDernier ;
END ;

procedure TFCtrRubbu.FormShow(Sender: TObject);
var
  Composants : TControlFiltre; //SG6   Gestion des Filtes 10/11/04   FQ 14976
begin
//SG6 16/11/04 Gestion des Filtres      FQ 14976
Composants.PopupF   := POPF;
Composants.Filtres  := FFILTRES;
Composants.Filtre   := BFILTRE;
Composants.PageCtrl := Pages;
ObjFiltres := TObjFiltre.Create(Composants, 'CTRLRUBBUD');

//ChargeFiltre(FNomFiltre,FFiltres,Pages) ;
LRubFam.Caption:=MsgBox.Mess[33] ; TTypRub.Caption:=MsgBox.Mess[34] ;
Typrub.Visible:=False ; CbBud.Left:=Rub2.Left ; CbBud.Width:=Rub2.Width ;
CbBud.Visible:=True ;
LFam:='' ; LRub1:='' ; LRub2:='' ; Lcpt1:='' ; Lcpt2:='' ; Stop:=False ; MemoFam:='' ;
MemoTyp:='' ; ChargeTypRub ; CbBud.ItemIndex:=0 ; FamRub.ItemIndex:=0 ;
CbAxe.ItemIndex:=0 ; Typrub.Value:='GEN' ;
FamRubChange(Nil) ; RgChoixClick(Nil) ; InitFliste ;
OnNeSortPas:=FALSE ;
ObjFiltres.Charger;

end;

procedure TFCtrRubbu.C1Exit(Sender: TObject);
begin C1.Text:=BourrelaDonc(C1.Text,LeFb) ; end;

procedure TFCtrRubbu.C2Exit(Sender: TObject);
begin C2.Text:=BourrelaDonc(C2.Text,LeFb) ; end;

procedure TFCtrRubbu.BChercheClick(Sender: TObject);
begin
if FListe.Cells[0,1]<>'' then FListe.VidePile(False) ;
VideLaListe ; OnNeSortPas:=TRUE ; RunControle ; OnNeSortPas:=FALSE ;
end;

Procedure TFCtrRubbu.PositionneRequete ;
Var St : String ;
BEGIN
if (LRub1='') or (LRub2='') then Exit ;
QRub.Close ; QRub.Sql.Clear ; QGen.Close ; QGen.Sql.Clear ;
St:=' And (RB_NATRUB="BUD" And RB_BUDJAL="'+CbBud.Value+'")' ;
Case RgChoix.ItemIndex of
     0: BEGIN
        QRub.Sql.Add('Select * From RUBRIQUE Where RB_RUBRIQUE>="'+LRub1+'" And RB_RUBRIQUE<="'+LRub2+'" And '+
                     'RB_TYPERUB="'+Typrub.Value+'"') ;
        QRub.Sql.Add(St) ;
        If (TypRub.Value='ANA') Or (TypRub.Value='G/A') Or (TypRub.Value='A/G')Then
          BEGIN
          QRub.Sql.Add(' And RB_AXE="'+CbAxe.Value+'" ') ;
          END ;
        ChangeSql(QRub) ; QRub.Open ;
        END ;
     1: BEGIN
        QRub.Sql.Add('Select * From RUBRIQUE Where RB_TYPERUB="'+Typrub.Value+'"') ;
        QRub.Sql.Add(St) ;
        If (TypRub.Value='ANA') Or (TypRub.Value='G/A') Or (TypRub.Value='A/G')Then
          BEGIN
          QRub.Sql.Add(' And RB_AXE="'+CbAxe.Value+'" ') ;
          END ;
        Case Lefb of
             fbGene : QGen.Sql.Add('Select G_GENERAL,G_LIBELLE From GENERAUX Where G_GENERAL>="'+Lcpt1+'" And '+
                     'G_GENERAL<="'+Lcpt2+'" Order by G_GENERAL') ;
             fbAux :  QGen.Sql.Add('Select T_AUXILIAIRE,T_LIBELLE From TIERS Where T_AUXILIAIRE>="'+Lcpt1+'" And '+
                     'T_AUXILIAIRE<="'+Lcpt2+'" Order by T_AUXILIAIRE') ;
             fbBudgen : QGen.Sql.Add('Select BG_BUDGENE,BG_LIBELLE From BUDGENE Where BG_BUDGENE>="'+Lcpt1+'" And '+
                     'B_BUDGENE<="'+Lcpt2+'" Order by BG_BUDGENE') ;
             fbAxe1..fbAxe5 : QGen.Sql.Add('Select S_SECTION,S_LIBELLE From SECTION Where S_SECTION>="'+Lcpt1+'" And '+
                     'S_SECTION<="'+Lcpt2+'" And S_AXE="'+CbAxe.Value+'" Order by S_SECTION') ;
           End ;
        ChangeSql(QRub) ; QRub.Open ; ChangeSql(QGen) ; QGen.Open ;
        END ;
  End ;
END ;

Procedure TFCtrRubbu.BloqueControle(InRun : Boolean) ;
BEGIN
Pages.Enabled:=Not InRun ; PFiltres.Enabled:=Not InRun ; BAgrandir.Enabled:=Not InRun ;
BReduire.Enabled:=Not InRun ; BRechercher.Enabled:=Not InRun ;
BMenuZoom.Enabled:=Not InRun ; 
END ;

Procedure TFCtrRubbu.RunControle ;
BEGIN
LFam:=FamRub.Value ; LRub1:=Rub1.Text ; LRub2:=Rub2.Text ; InitFliste ;
Lcpt1:=C1.Text ; Lcpt2:=C2.Text ; PositionneRequete ;
BloqueControle(True) ; Stop:=False ;
Case RgChoix.ItemIndex of
     0: BEGIN
        RempliListeRubGen(False) ;
        InitMove(ListeRub.Count,MsgBox.Mess[3]) ;
        ControleLesRubriques ; FiniMove ;
        END ;
     1: BEGIN
        RempliListeRubGen(True) ;
        InitMove(ListeGen.Count,MsgBox.Mess[3]) ;
        ControleLesComptes ; FiniMove ;
        END ;
  End ;
BloqueControle(False) ;
END ;

Procedure TFCtrRubbu.RempliListeRubGen(AvecGen : Boolean ) ;
Var X : TInfoRub ;
    Cod,Lib : String ;
BEGIN
While Not QRub.Eof do
   BEGIN
   X:=TInfoRub.Create ;
   X.Libelle:=QRub.FindField('RB_LIBELLE').AsString ;
   X.Famille:=QRub.FindField('RB_FAMILLES').AsString ;
   X.Compte1:=QRub.FindField('RB_COMPTE1').AsString ;
   X.Exclu1:=QRub.FindField('RB_EXCLUSION1').AsString ;
   X.Compte2:=QRub.FindField('RB_COMPTE2').AsString ;
   X.Exclu2:=QRub.FindField('RB_EXCLUSION2').AsString ;
   X.TabLib:=QRub.FindField('RB_TABLELIBRE').AsString ;
   X.Axe:=QRub.FindField('RB_AXE').AsString ;
   ListeRub.AddObject(QRub.FindField('RB_RUBRIQUE').AsString,X) ;
   QRub.Next ;
   END ;
if AvecGen then
   BEGIN
   Case Lefb of
        fbGene : BEGIN Cod:='G_GENERAL' ; Lib:='G_LIBELLE' ; END ;
        fbAux :  BEGIN Cod:='T_AUXILIAIRE' ; Lib:='T_LIBELLE' ; END ;
        fbAxe1..fbAxe5 : BEGIN Cod:='S_SECTION' ; Lib:='S_LIBELLE' ; END ;
      End ;
   While Not QGen.Eof do
      BEGIN
      X:=TInfoRub.Create ;
      X.Libelle:=QGen.FindField(Lib).AsString ;
      X.Famille:='' ; X.Compte1:='' ; X.Exclu1:='' ; X.Compte2:='' ; X.Exclu2:='' ; X.Axe:='' ;
      ListeGen.AddObject(QGen.FindField(Cod).AsString,X) ;
      QGen.Next ;
      END ;
   END ;
QRub.Close ; QGen.Close ;
END ;

Function TFCtrRubbu.FaitRequeteGene(Compte1,Cexlu1:String ; Unfb : TFichierBase ; SurTabLib : Boolean ) : String ;
Var Sql,St,Where : String ;
    Cod,Lib,Table : String ;
BEGIN
Result:='' ; St:='' ;
Case Unfb of
    fbAxe1..fbAxe5 : BEGIN Cod:='S_SECTION' ; Lib:='S_LIBELLE' ; Table:='SECTION' ; St:='And S_AXE="'+CbAxe.Value+'" ' ; END ;
    fbGene : BEGIN Cod:='G_GENERAL' ; Lib:='G_LIBELLE' ; Table:='GENERAUX' ; END ;
    fbAux : BEGIN Cod:='T_AUXILIAIRE' ; Lib:='T_LIBELLE' ; Table:='TIERS' ; END ;
    fbBudgen : BEGIN Cod:='BG_BUDGENE' ; Lib:='BG_LIBELLE' ; Table:='BUDGENE' ; END ;
  End ;
Where:=AnalyseCompte(Compte1,Unfb,False,SurTabLib) ;
if Where<>'' then
   BEGIN
   Sql:='Select '+Cod+','+Lib+' From '+Table+'' ;
   Sql:=Sql+' Where '+Where ;
   END else Exit ;
if Cexlu1<>'' then
   BEGIN
   Where:=AnalyseCompte(Cexlu1,Unfb,True,FALSE) ;
   if Where<>'' then Sql:=Sql+' And '+Where ;
   END ;
Sql:=Sql+St+' Order by '+Cod ;
Result:=Sql ;
END ;

Procedure TFCtrRubbu.TraiteRequeteGen(Ind : Integer ; Composite : Boolean ; Unfb : TFichierBase) ;
BEGIN
While Not QGen.Eof do
   BEGIN
   Fliste.Cells[2,FListe.RowCount-1]:=QGen.Fields[0].AsString ;
   Fliste.Cells[3,FListe.RowCount-1]:=QGen.Fields[1].AsString ;
   Fliste.Cells[4,FListe.RowCount-1]:=MsgBox.Mess[5+Ind] ;
   if Not Composite then
      BEGIN
      if Pos('A',Typrub.Value)=1 then
         BEGIN
         if VH^.Cpta[AxeToFb(CbAxe.Value)].Chantier then Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[6]
                                                   else Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[5] ;
         END else
         BEGIN
         QEcr.Close ; QEcr.Params[0].AsString:=QGen.Fields[0].AsString ; QEcr.Open ;
         if QEcr.Eof then Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[5]
                     else Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[6] ;
         END ;
      END else
      BEGIN
      if UnFb in [fbAxe1..fbAxe5] then
         BEGIN
         if VH^.Cpta[AxeToFb(CbAxe.Value)].Chantier then Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[6]
                                                   else Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[5] ;
         END else
         BEGIN
         QEcr.Close ; QEcr.Params[0].AsString:=QGen.Fields[0].AsString ; QEcr.Open ;
         if QEcr.Eof then Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[5]
                     else Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[6] ;
         END ;
      END ;
   QGen.Next ; FListe.RowCount:=FListe.RowCount+1 ;
   END ;
END ;

Procedure TFCtrRubbu.ControleLesRubriques ;
Var Compte1,Cexlu1,St,Compte2,Cexlu2 : String ;
    i : Integer ;
    fb1,fb2 : TFichierBase ;
BEGIN
fb1:=Lefb ; fb2:=QuelfbComposite ;
for i:=0 to ListeRub.Count-1 do
  BEGIN
  if Pos(LFam,TInfoRub(ListeRub.Objects[i]).Famille)>0 then
     BEGIN
     FListe.Cells[0,FListe.RowCount-1]:=ListeRub.Strings[i] ;
     FListe.Cells[1,FListe.RowCount-1]:=TInfoRub(ListeRub.Objects[i]).Libelle ;
     Compte1:=TInfoRub(ListeRub.Objects[i]).Compte1 ;
     Cexlu1:=TInfoRub(ListeRub.Objects[i]).Exclu1 ;
     St:=FaitRequeteGene(Compte1,Cexlu1,fb1,TInfoRub(ListeRub.Objects[i]).TabLib='X') ;
     if St='' then Continue ;
     QGen.Close ; QGen.Sql.Clear ; QGen.Sql.Add(St) ; ChangeSQL(QGen) ; QGen.Open ;
     if Pos('A',Typrub.Value)=1 then else
       BEGIN
       FaitRequeteEcr(fb1) ; QEcr.Close ; ChangeSQL(QEcr) ;
       //QEcr.Prepare ;
       PrepareSQLODBC(QEcr) ;
       END ;
     TraiteRequeteGen(0,False,fb1) ;
     St:=FaitRequeteGene(Cexlu1,'',fb1,False) ;
     if St<>'' then
        BEGIN
        QGen.Close ; QGen.Sql.Clear ; QGen.Sql.Add(St) ; ChangeSQL(QGen) ; QGen.Open ;
        TraiteRequeteGen(1,False,fb1) ;
        END ;
     if Pos('/',Typrub.Value)>0 then
        BEGIN
        Compte2:=TInfoRub(ListeRub.Objects[i]).Compte2 ;
        Cexlu2:=TInfoRub(ListeRub.Objects[i]).Exclu2 ;
        St:=FaitRequeteGene(Compte2,Cexlu2,fb2,False) ;
        if St<>'' then
           BEGIN
           FListe.Cells[0,FListe.RowCount-1]:=ListeRub.Strings[i] ;
           FListe.Cells[1,FListe.RowCount-1]:=TInfoRub(ListeRub.Objects[i]).Libelle ;
           QGen.Close ; QGen.Sql.Clear ; QGen.Sql.Add(St) ; ChangeSQL(QGen) ; QGen.Open ;
           if Pos('A',Typrub.Value)=3 then else
             BEGIN
             FaitRequeteEcr(fb2) ; QEcr.Close ; ChangeSQL(QEcr) ;
             //QEcr.Prepare ;
             PrepareSQLODBC(QEcr) ;
             END ;
           TraiteRequeteGen(0,True,fb2) ;
           END ;
        St:=FaitRequeteGene(Cexlu2,'',fb2,False) ;
        if St<>'' then
           BEGIN
           QGen.Close ; QGen.Sql.Clear ; QGen.Sql.Add(St) ; ChangeSQL(QGen) ; QGen.Open ;
           TraiteRequeteGen(1,True,fb2) ;
           END ;
        END ;
     END ;
  MoveCur(False) ;
  if TestStop then Break ;
  END ;
if Fliste.RowCount>2 then Fliste.RowCount:=Fliste.RowCount-1 ;
END ;

Function TFCtrRubbu.ChercheCompteDansRub(CptRub,ComptGene : String) : Boolean ;
Var St,StTemp : String ;
BEGIN
Result:=False ;
While CptRub<>'' do
  BEGIN
  St:=ReadTokenSt(CptRub) ; StTemp:='' ;
  if St='' then Continue ;
  if Pos('(',St)>0 then St:=Copy(St,1,Pos('(',St)-1) ;
  if Pos(':',St)>0 then BEGIN StTemp:=Copy(St,Pos(':',St)+1,200) ; Delete(St,Pos(':',St),200) ; END ;
  if StTemp<>'' then
     BEGIN
     St:=BourrelaDonc(St,Lefb) ; StTemp:=BourreLaDonc(StTemp,Lefb) ;
     if(ComptGene>=St) And (ComptGene<=StTemp) then
        BEGIN Result:=True ; Exit ; END ;
     END else
     BEGIN
     if Length(St)<=VH^.Cpta[Lefb].Lg then
        if(Pos(St,ComptGene)=1)Or(St=ComptGene) then
           BEGIN Result:=True ; Exit ; END ;
     END ;
  END ;
END ;

Procedure TFCtrRubbu.ControleLesComptes ;
Var Compte,CptRub,CptExRub : String ;
    PremierTour : Boolean ;
    i,j : Integer ;
BEGIN
for j:=0 to ListeGen.Count-1 do
  BEGIN
  MoveCur(False) ; PremierTour:=True ;
  for i:=0 to ListeRub.Count-1 do
    BEGIN
    if PremierTour then
       BEGIN
       Compte:=ListeGen.Strings[j] ;
       FListe.Cells[0,FListe.RowCount-1]:=Compte ;
       FListe.Cells[1,FListe.RowCount-1]:=TInfoRub(ListeGen.Objects[j]).Libelle ;
       if Pos('A',Typrub.Value)=1 then
          BEGIN
          if VH^.Cpta[AxeToFb(CbAxe.Value)].Chantier then Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[6]
                                                    else Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[5] ;
          END else
          BEGIN
          QEcr.Close ; QEcr.Params[0].AsString:=Compte ; QEcr.Open ;
          if QEcr.Eof then Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[5]
                      else Fliste.Cells[5,FListe.RowCount-1]:=MsgBox.Mess[6] ;
          END ;
       PremierTour:=False ;
       END ;
    if Pos(LFam,TInfoRub(ListeRub.Objects[i]).Famille)>0 then
       BEGIN
       CptRub:=TInfoRub(ListeRub.Objects[i]).Compte1 ;
       CptExRub:=TInfoRub(ListeRub.Objects[i]).Exclu1 ;
       if ChercheCompteDansRub(CptRub,Compte)then
          BEGIN
          FListe.Cells[2,FListe.RowCount-1]:=ListeRub.Strings[i] ;
          FListe.Cells[3,FListe.RowCount-1]:=TInfoRub(ListeRub.Objects[i]).Libelle ;
          FListe.Cells[4,FListe.RowCount-1]:=MsgBox.Mess[5] ;
          FListe.RowCount:=FListe.RowCount+1 ;
          END ;
       if CptExRub<>'' then
          BEGIN
          if ChercheCompteDansRub(CptExRub,Compte)then
             BEGIN
             FListe.Cells[2,FListe.RowCount-1]:=ListeRub.Strings[i] ;
             FListe.Cells[3,FListe.RowCount-1]:=TInfoRub(ListeRub.Objects[i]).Libelle ;
             FListe.Cells[4,FListe.RowCount-1]:=MsgBox.Mess[6] ;
             FListe.RowCount:=FListe.RowCount+1 ;
             END ;
          END ;
       END ;
    END ;
  if CbTous.Checked then
     if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount:=FListe.RowCount+1 ;
  if TestStop then Break ;
  END ;
if Fliste.RowCount>2 then Fliste.RowCount:=Fliste.RowCount-1 ;
if(Fliste.RowCount=2) And (Fliste.Cells[2,1]='') then
   for i:=0 to Fliste.ColCount-1 do Fliste.Cells[i,1]:='' ;
END ;

procedure TFCtrRubbu.BRubClick(Sender: TObject);
Var C : String ;
begin
if FListe.Cells[0,1]='' then Exit ;
Case RgChoix.ItemIndex of
     0:BEGIN
       if FListe.Cells[0,FListe.Row]='' then C:=ChecheLeCompte
                                        else C:=FListe.Cells[0,FListe.Row] ;
       ParametrageRubrique(C,taConsult, CtxBudget)
       END ;
     1: BEGIN
        ParametrageRubrique(FListe.Cells[2,Fliste.Row],taConsult,CtxBudget) ;
        END ;
    End ;
end;

procedure TFCtrRubbu.BGenClick(Sender: TObject);
Var C : String ;
begin
if FListe.Cells[0,1]='' then Exit ;
Case RgChoix.ItemIndex of
  0: BEGIN
     Case Lefb of
        fbAxe1..fbAxe5 :FicheSection(Nil,CbAxe.Value,FListe.Cells[2,FListe.Row],taConsult,0) ;
        fbGene : FicheGene(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0) ;
        fbAux : FicheTiers(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0) ;
{$IFNDEF CCS3}
        fbBudgen : FicheBudgene(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0) ;
{$ENDIF}
        End ;
     END ;

  1: BEGIN
     if FListe.Cells[0,FListe.Row]='' then C:=ChecheLeCompte
                                      else C:=FListe.Cells[0,FListe.Row] ;
     Case Lefb of
        fbAxe1..fbAxe5 :FicheSection(Nil,CbAxe.Value,FListe.Cells[2,FListe.Row],taConsult,0) ;
        fbGene : FicheGene(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0) ;
        fbAux : FicheTiers(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0) ;
{$IFNDEF CCS3}
        fbBudgen : FicheBudgene(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0) ;
{$ENDIF}
        End ;
     END ;
  End ;
end;

Function TFCtrRubbu.ChecheLeCompte : String ;
Var i : Integer ;
BEGIN
Result:='' ;
for i:=Fliste.Row Downto 1 do
    if Fliste.Cells[0,i]<>'' then
       BEGIN Result:=Fliste.Cells[0,i] ; Exit ; END ;
END ;

procedure TFCtrRubbu.FListeDblClick(Sender: TObject);
begin
Case RgChoix.ItemIndex of
     0: BRubClick(Nil) ;
     1: BGenClick(Nil) ;
   End ;
end;

procedure TFCtrRubbu.BStopClick(Sender: TObject);
begin Stop:=True ; end;

Function TFCtrRubbu.TestStop : Boolean ;
BEGIN
Application.ProcessMessages ; Result:=False ;
if Not Stop then Exit ;
Stop:=False ;
if MsgBox.Execute(8,'','')=mrYes then Result:=True ;
END ;

Procedure TFCtrRubbu.VideLaListe ;
Var i : Integer ;
BEGIN
for i:=0 to ListeRub.Count-1 do TObject(ListeRub.Objects[i]).Free ;
ListeRub.Clear ;
for i:=0 to ListeGen.Count-1 do TObject(ListeGen.Objects[i]).Free ;
ListeGen.Clear ;
END ;

procedure TFCtrRubbu.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(False) ; VideLaListe ; ListeRub.Free ; ListeGen.Free ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFCtrRubbu.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCtrRubbu.FormCreate(Sender: TObject);
begin
FNomFiltre:='' ; PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=Height ; FiltreEnCours:=False ;
ListeRub:=TStringList.Create ; ListeGen:=TStringList.Create ;
end;

procedure TFCtrRubbu.RgChoixClick(Sender: TObject);
begin
if Pos('/',Typrub.Value)>0 then RgChoix.ItemIndex:=0 ;
Case RgChoix.ItemIndex of
     0: BEGIN
        Rub1.Enabled:=True ; Rub2.Enabled:=True ;
        C1.Enabled:=False ; C2.Enabled:=False ; CbTous.Enabled:=False ;
        END ;
     1: BEGIN
        C1.Enabled:=True ; C2.Enabled:=True ; CbTous.Enabled:=True ;
        PremierToDernier ;
        Rub1.Enabled:=False ; Rub2.Enabled:=False ;
        END ;
   End ;
end;

Procedure TFCtrRubbu.ChargeTypRub ;
Var QLoc : TQuery ;
BEGIN
TypRub.Values.Clear ; TypRub.Items.Clear ;
TypRub.DataType:='ttRubTypeBud' ; RempliComboFamRub ;
END ;

Procedure TFCtrRubbu.RempliComboFamRub ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CO_CODE,CO_LIBELLE from COMMUN Where CO_TYPE="RBB"',True) ;
FamRub.Values.Clear ; FamRub.Items.Clear ;
While not Q.Eof do
  BEGIN
  if Q.Fields[0].AsString='GEN' then FamRub.Values.Add('CBG') else
     if Q.Fields[0].AsString='ANA' then FamRub.Values.Add('CBS') else
        if Q.Fields[0].AsString='G/A' then FamRub.Values.Add('G/S') else
           if Q.Fields[0].AsString='A/G' then FamRub.Values.Add('S/G') ;
  FamRub.Items.Add(Q.FindField('CO_LIBELLE').AsString) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

procedure TFCtrRubbu.CbAxeChange(Sender: TObject);
begin
if Not CbAxe.Enabled then Exit ;
if Pos('A',TypRub.Value)=1 then
   BEGIN
   Case CbAxe.Value[2] of
       '1' : BEGIN LeFb:=fbAxe1 ; C1.ZoomTable:=tzSection  ; C2.ZoomTable:=tzSection  ; END ;
       '2' : BEGIN LeFb:=fbAxe2 ; C1.ZoomTable:=tzSection2 ; C2.ZoomTable:=tzSection2 ; END ;
       '3' : BEGIN LeFb:=fbAxe3 ; C1.ZoomTable:=tzSection3 ; C2.ZoomTable:=tzSection3 ; END ;
       '4' : BEGIN LeFb:=fbAxe4 ; C1.ZoomTable:=tzSection4 ; C2.ZoomTable:=tzSection4 ; END ;
       '5' : BEGIN LeFb:=fbAxe5 ; C1.ZoomTable:=tzSection5 ; C2.ZoomTable:=tzSection5 ; END ;
     End ;
   END ;
PremierToDernier ; FamRubChange(Nil) ;
end;

Procedure TFCtrRubbu.FaitRequeteEcr(Unfb : TFichierBase) ;
BEGIN
QEcr.Close ; QEcr.Sql.Clear ;
Case Unfb of
    fbGene : QEcr.Sql.Add('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL=:Cpte '+
                          'AND ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL=:Cpte))'+
                          'Or (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL=:Cpte )))') ;

    fbAux :  QEcr.Sql.Add('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE=:Cpte '+
                          'AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE=:Cpte)');

    fbBudgen :  QEcr.Sql.Add('SELECT BG_BUDGENE FROM BUDGENE WHERE BG_BUDGENE=:Cpte '+
                          'AND (EXISTS(SELECT BE_BUDGENE FROM BUDECR WHERE BE_BUDGENE=:Cpte))') ;

  End ;
END ;

procedure TFCtrRubbu.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFCtrRubbu.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FirstFind); end ;

procedure TFCtrRubbu.BRechercherClick(Sender: TObject);
begin FirstFind:=true; FindDialog.Execute ; end;

Procedure TFCtrRubbu.PremierToDernier ;
Var Cpt1,Cpt2 : String ;
BEGIN
if ObjFiltres.InChargement then Exit;
//if FiltreEnCours then Exit ;
QuelEstLeFb ; Cpt1:='' ; Cpt2:='' ;
PremierDernier(Lefb,Cpt1,Cpt2) ; C1.Text:=Cpt1 ; C2.Text:=Cpt2 ;
END ;

procedure TFCtrRubbu.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TFCtrRubbu.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ; 
end;

procedure TFCtrRubbu.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
If OnNeSortPas Then CanClose:=FALSE ;
end;

procedure TFCtrRubbu.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=vk_F9 then BChercheClick(nil); //SG6 17/11/04 Gestion des filtres FQ 14976
end;

end.
