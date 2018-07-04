unit AffBqe;

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
  ExtCtrls,
  ComCtrls,
  StdCtrls,
  Buttons,
  Grids,
  Hctrls,
  HDB,
  Ent1,
  HEnt1,
  ParamDat,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry,
  Menus,
  hmsgbox,
  HSysMenu,
  Hspliter,
  HTB97,
  HPanel,
  UiUtil,
  UTOB,
  Approxim, DB, Mask, TntButtons, TntStdCtrls, TntGrids ;

Const CodError=-123456789.321  ;

Procedure AffecteBanquePrevi(Client : Boolean ; Action : tActionFiche) ;

Type TCol=Class
       LCol : Integer ;
       Libelle : String ;
     End ;
type
  TFAffBQEED = class(TForm)
    BReduire: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BRechercher: TToolbarButton97;
    FindDialog: TFindDialog;
    PFiltres: TToolWindow97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    HM: THMsgBox;
    FListe: THGrid;
    HMTrad: THSystemMenu;
    QBud: TQuery;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    SM: THNumEdit;
    HPB: TToolWindow97;
    Dock: TDock97;
    BImprimer: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock971: TDock97;
    bExport: TToolbarButton97;
    Pages: TPageControl;
    PCritere: TTabSheet;
    Bevel1: TBevel;
    PComplement: TTabSheet;
    Bevel2: TBevel;
    HLabel3: THLabel;
    E_NOMLOT: THCritMaskEdit;
    HLabel5: THLabel;
    E_NOMLOT_: THCritMaskEdit;
    HLabel6: THLabel;
    E_BANQUEPREVI: THCritMaskEdit;
    HLabel8: THLabel;
    E_BANQUEPREVI_: THCritMaskEdit;
    SD: TSaveDialog;
    BSelectMvt: TToolbarButton97;
    PVentil: TPanel;
    H_CODEVENTIL: THLabel;
    HLabel2: THLabel;
    H_TITREVENTIl: TLabel;
    Y_CODEVENTIL: TEdit;
    Y_LIBELLEVENTIL: TEdit;
    BNewVentil: THBitBtn;
    PFenCouverture: TPanel;
    BOuvrir: TToolbarButton97;
    BCalcul: TToolbarButton97;
    PCumul: TPanel;
    FTotPSaisi: THNumEdit;
    FTotMSaisi: THNumEdit;
    TotRepart: THNumEdit;
    TFType: THLabel;
    CVType: THValComboBox;
    BSauveVentil: TToolbarButton97;
    FTotPCalcul: THNumEdit;
    FTotMCalcul: THNumEdit;
    WFListe2: TToolWindow97;
    PFListe2: TPanel;
    FListe2: THGrid;
    ToolbarButton972: TToolbarButton97;
    ToolbarButton973: TToolbarButton97;
    bSelectAll: TToolbarButton97;
    BSelectBqe: TToolbarButton97;
    bPoubelle: TToolbarButton97;
    BDeselectAll: TToolbarButton97;
    BDeselectBqe: TToolbarButton97;
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeTopLeftChanged(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure FListe2TopLeftChanged(Sender: TObject);
    procedure FListe2DblClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure FListe2Enter(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure bExportClick(Sender: TObject);
    procedure BSelectMvtClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure CVTypeChange(Sender: TObject);
    procedure BNewVentilClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BCalculClick(Sender: TObject);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure BOuvrirClick(Sender: TObject);
    procedure FListe2MouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);
    procedure FListe2DragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure FListe2EndDrag(Sender, Target: TObject; X, Y: Integer);
    procedure FListeDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure FListeDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure bSelectAllClick(Sender: TObject);
    procedure BSelectBqeClick(Sender: TObject);
    procedure bPoubelleClick(Sender: TObject);
    procedure BDeselectBqeClick(Sender: TObject);
    procedure BDeselectAllClick(Sender: TObject);
    procedure BSauveVentilClick(Sender: TObject);
  private
    CurGrid : Integer ;
    FNomFiltre : String ;
    FirstFind : boolean;
    MemoLig : Integer ;
    Symb : String ;
    WMinX,WMinY : Integer ;
    OldBrush    : TBrush ;
    OldPen      : TPen ;
    FiltreEnCours : Boolean ;
    OnNeSortPas : Boolean ;
    Client : Boolean ;
    TotDAffecte,TotCAffecte : Double ;
    Action : tActionFiche ;
    TobMvt : TOB ;
    ModeVentil : Boolean ;
    OkDrop : Boolean ;
    NbFListe2Selected : Integer ;
    CalculFait : Boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure InitChamp ;
    Procedure BloqueControle(InRun : Boolean) ;
    Procedure AfficheFliste2(Avec : Boolean ) ;
    procedure RempliGridCptBqe ;
    Function  TrouveLig(Cpt : String) : Integer ;
    Procedure CalculMontantDejaAffecte ;
    Function  TrouveLigne(Cpt : String) : Integer ;
    procedure Preventil (NomGrille : String) ;
    procedure ClickSauveVentil ;
    procedure FinVentil ;
    procedure AfficheTotauxSaisi ;
    procedure AfficheTotauxCalcul(What : Integer) ;
    Procedure ResizeLesTotaux(Sender: TObject) ;
    procedure AfficheTotARepartir ;
    function  ConstruitListe(var LM : T_DX; var LP,LB : T_IX) : integer ;
    procedure AffichePendantCalcul(R : Integer ; D,C : Double) ;
    Procedure GoCalcul ;
    Function  FiniCalcul(What : Integer) : Boolean ;
    Function  TenteAffectation(TOBL : TOB ; What : Integer) : Boolean ;
    procedure AnnuleBanquePrevi ;
    Procedure ValideLeLot ;
    Function  MontantBqe(Cpt : String) : Double ;
    Procedure UpdateBanquePrevi(Cpt : String) ;
    Procedure TraiteFListe2(What : Integer) ;
    procedure SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
    procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    Procedure ReinitMontant ;
    procedure InitTitres ;
  public
  end;

implementation

Uses
  PrintDBG,
  Filtre,
  HStatus,
  HXLSPas,
  MulSuiviBP,
  SaisLot;

{$R *.DFM}

Const ab_Banque = 1 ;
      ab_Compte = 0 ;
      ab_SoldeAvant = 2 ;
      ab_SoldeApres = 3 ;
      ab_PourcentSaisi = 4 ;
      ab_MontantSaisi = 5 ;
      ab_PourcentCalcul = 6 ;
      ab_MontantCalcul = 7 ;
(*
1 Banque
0 Compte
2 Solde avant affectation
3 Solde après affectation
4 % à saisir
5 Montant à saisir
6 % déjà affecté / % Calculé
7 Montant déjà affecté / Montant Calculé
*)

Procedure AffecteBanquePrevi(Client : Boolean ; Action : tActionFiche) ;
var XX : TFAffBQEED ;
    PP : THPanel ;
BEGIN
XX:=TFAffBQEED.Create(Application) ;
XX.FNomFiltre:='AFFBQE' ;
XX.Client:=Client ;
XX.Action:=Action ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     XX.ShowModal ;
    Finally
     XX.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;
END ;

procedure TFAffBQEED.BCreerFiltreClick(Sender: TObject);
begin NewFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFAffBQEED.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFAffBQEED.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre(FNomFiltre,FFiltres) ; end;

procedure TFAffBQEED.BRenFiltreClick(Sender: TObject);
begin RenameFiltre(FNomFiltre,FFiltres) ; end;

procedure TFAffBQEED.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres,Pages) ; InitChamp ; end;

procedure TFAffBQEED.FFiltresChange(Sender: TObject);
begin
FiltreEnCours:=True ;
LoadFiltre(FNomFiltre,FFiltres,Pages) ;
FiltreEnCours:=False ;
end;

procedure TFAffBQEED.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFAffBQEED.BReduireClick(Sender: TObject);
begin
if FListe2.Visible then FListe.Height:=80 ;
ChangeListeCrit(Self,False) ;
end;

procedure TFAffBQEED.BRechercherClick(Sender: TObject);
begin FirstFind:=true; FindDialog.Execute ; end;

procedure TFAffBQEED.BAnnulerClick(Sender: TObject);
begin

if FListe2.Visible then
   BEGIN FListe2.VidePile(FALSE) ; FListe2.Visible:=False ; END ;
//FListe.VidePile(True) ; Close ;
end;

procedure TFAffBQEED.BImprimerClick(Sender: TObject);
(*
Var LCpte : TStringList ;
    I, J : Integer ; St : String ;
    InfGene : TCol ;
    LeCrit : SaveMCrit ;
*)
begin
//PrintDBGrid (FListe,Nil, Caption,'') ;

end;

procedure TFAffBQEED.FormCreate(Sender: TObject);
begin
FNomFiltre:='' ; PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; Symb:='' ;
WMinX:=Width ; WMinY:=Height ; OldBrush:=TBrush.Create ; OldPen:=TPen.Create ;
TOBMvt:=TOB.Create('',Nil,-1) ;
If CLient Then RegLoadToolbarPos(Self,'AFFBQECLI') Else RegLoadToolbarPos(Self,'AFFBQEFOU') ;
FListe.OnColumnWidthsChanged:=ResizeLestotaux ;
If Action=taCreat Then BEGIN Pages.Visible:=FALSE ; PFiltres.DockedTo:=NIL ; PFiltres.Visible:=FALSE ; END ;
Fliste.PostDrawCell:=PostDrawCell ;
Fliste.GetCellCanvas:=GetCellCanvas ;
FListe.Refresh ;
end;

Procedure TFAffBQEED.InitChamp ;
Var i : Integer ;
(*
Const ab_Banque = 0 ;
      ab_Compte = 1 ;
      ab_SoldeAvant = 2 ;
      ab_SoldeApres = 3 ;
      ab_PourcentSaisi = 4 ;
      ab_MontantSaisi = 5 ;
      ab_PourcentCalcul = 6 ;
      ab_MontantCalcul = 7 ;
*)
BEGIN
Pages.ActivePage:=Pages.Pages[0] ;
For i:=ab_Compte to ab_montantCalcul  Do
  If (i<>ab_PourcentSaisi) And (i<>ab_MontantSaisi) Then FListe.FColLengths[i]:=-1 ;
(*
FListe.FColFormats[ab_PourcentSaisi]:='#,##0.00' ;
FListe.FColFormats[ab_MontantSaisi]:='#,##0.00' ;
*)
END ;

Procedure TFAffBQEED.ReinitMontant ;
Var i : Integer ;
BEGIN
For i:=1 To FListe.RowCount-1 Do
  BEGIN
  FListe.Cells[ab_PourcentCalcul,i]:='' ;FListe.Cells[ab_MontantCalcul,i]:='' ; FListe.Cells[ab_SoldeApres,i]:='' ;
  END ;
END ;

procedure TFAffBQEED.InitTitres ;
BEGIN
If Action=taCreat Then Caption:=HM.Mess[11] Else Caption:=HM.Mess[12] ; 
END ;

procedure TFAffBQEED.FormShow(Sender: TObject);
begin
ChargeFiltre(FNomFiltre,FFiltres,Pages) ;
SM.Visible:=False ; MemoLig:=0 ; ModeVentil:=False ;
FListe2.Visible:=False ;
InitChamp ; FListe.ALign:=alClient ;
if FFiltres.Text='DEFAUT' then FFiltresChange(FFiltres) ;
OnNeSortPas:=FALSE ;
RempliGridCptBqe ;
AfficheTotARepartir ;
WFliste2.Resizable:=TRUE ; WFListe2.Visible:=FALSE ; Fliste2.Visible:=FALSE ;
CalculFait:=FALSE ;
InitTitres ;
UpdateCaption(Self) ;
If Action=taCreat Then BSelectMvtClick(Nil);
end;

Procedure TFAffBQEED.BloqueControle(InRun : Boolean) ;
BEGIN
Pages.Enabled:=Not InRun ; PFiltres.Enabled:=Not InRun ; BAgrandir.Enabled:=Not InRun ;
BReduire.Enabled:=Not InRun ; BRechercher.Enabled:=Not InRun ;
OnNeSortPas:=InRun ;
END ;

procedure TFAffBQEED.RempliGridCptBqe ;
Var Q : TQuery ;
    St : String ;
    S : Double ;
(*
Const ab_Banque = 0 ;
      ab_Compte = 1 ;
      ab_SoldeAvant = 2 ;
      ab_SoldeApres = 3 ;
      ab_PourcentSaisi = 4 ;
      ab_MontantSaisi = 5 ;
      ab_PourcentCalcul = 6 ;
      ab_MontantCalcul = 7 ;
*)
BEGIN
St:='SELECT G_GENERAL, G_LIBELLE, G_TOTDEBE, G_TOTCREE, G_TOTDEBS, G_TOTCRES FROM GENERAUX WHERE G_NATUREGENE="BQE"' ;
Q:=OpenSQL(St,TRUE) ;
While Not Q.Eof Do
  BEGIN
  FListe.Cells[ab_banque,FListe.RowCount-1]:=Q.FindField('G_LIBELLE').AsString ;
  FListe.Cells[ab_compte,FListe.RowCount-1]:=Q.FindField('G_GENERAL').AsString ;
  S:=Q.FindField('G_TOTDEBE').AsFloat+Q.FindField('G_TOTDEBS').AsFloat-Q.FindField('G_TOTCREE').AsFloat-Q.FindField('G_TOTCRES').AsFloat ;
  FListe.Cells[ab_SoldeAvant,FListe.RowCount-1]:=StrfMontant(S,0,V_PGI.OkDecV,'',True) ;
  Q.Next ;
  FListe.RowCount:=FListe.RowCount+1 ;
  END ;
Ferme(Q) ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
ReinitMontant ;
END ;

procedure TFAffBQEED.BChercheClick(Sender: TObject);
begin
if FListe2.Visible then AfficheFliste2(False) ;
If Action<>taCreat Then
  BEGIN
  BloqueControle(True) ;
  CalculMontantDejaAffecte ;
  BloqueControle(False) ;
  END ;
CurGrid:=1 ;
end;

Function TFAffBQEED.TrouveLig(Cpt : String) : Integer ;
Var i : Integer ;
BEGIN
Result:=0 ;
For i:=1 To FListe.RowCount-1 Do If FListe.Cells[ab_Compte,i]=Cpt Then BEGIN Result:=i ; Exit ; END ;
END ;

Procedure TFAffBQEED.CalculMontantDejaAffecte ;
Var Q : TQuery ;
    St : String ;
    i : Integer ;
    StNum,StDen : String ;
    StF1,StF2 : String ;
    XX : Double ;
    StV8 : String ;
(*
Const ab_Banque = 0 ;
      ab_Compte = 1 ;
      ab_SoldeAvant = 2 ;
      ab_SoldeApres = 3 ;
      ab_PourcentSaisi = 4 ;
      ab_MontantSaisi = 5 ;
      ab_PourcentCalcul = 6 ;
      ab_MontantCalcul = 7 ;
*)
BEGIN
TotDAffecte:=0 ; TotCAffecte:=0 ;
StNum:='(E_DEBIT+E_DEBITDEV)' ;
StDen:='(E_DEBIT+E_DEBITDEVE_CREDIT+E_CREDITDEV)' ;
StF1:='('+StNum+'/'+StDen+')' ;
StNum:='(E_CREDIT+E_CREDITDEV)' ;
StF2:='('+StNum+'/'+StDen+')' ;
St:='SELECT E_BANQUEPREVI,' ;
St:=St+' SUM(E_COUVERTURE*'+StF1+') AS COD, SUM(E_COUVERTURE*'+StF2+') AS COC, SUM(E_DEBIT) AS D, SUM(E_CREDIT) AS C FROM ECRITURE '  ;
St:=St+' WHERE E_BANQUEPREVI<>"" AND E_NOMLOT<>"" ' ;
If E_BANQUEPREVI.Text<>'' Then St:=St+' AND E_BANQUEPREVI>="'+E_BANQUEPREVI.Text+'" ' ;
If E_BANQUEPREVI_.Text<>'' Then St:=St+' AND E_BANQUEPREVI<="'+E_BANQUEPREVI_.Text+'" ' ;
If E_NOMLOT.Text<>'' Then St:=St+' AND E_NOMLOT>="'+E_NOMLOT.Text+'" ' ;
If E_NOMLOT_.Text<>'' Then St:=St+' AND E_NOMLOT<="'+E_NOMLOT_.Text+'" ' ;
StV8:=LWhereV8 ; if StV8<>'' then St:=St+' AND '+StV8+' ' ;
St:=St+' And (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") and E_QUALIFPIECE="N" ' ;
St:=St+' GROUP BY E_BANQUEPREVI ' ;
Q:=OpenSQL(St,TRUE) ;
While Not Q.Eof Do
  BEGIN
  i:=TrouveLig(Q.FindField('E_BANQUEPREVI').AsString) ;
  If i>0 Then
    BEGIN
    If Client Then XX:=Arrondi(Q.FindField('D').AsFloat-Q.FindField('C').AsFloat-Q.FindField('COD').AsFloat+Q.FindField('COC').AsFloat,V_PGI.OkDecV)
              Else XX:=Arrondi(Q.FindField('C').AsFloat-Q.FindField('D').AsFloat-Q.FindField('COC').AsFloat-Q.FindField('COD').AsFloat,V_PGI.OkDecV) ;
    FListe.Cells[ab_MontantCalcul,i]:=StrfMontant(XX,0,V_PGI.OkDecV,'',True) ;
    FListe.Cells[ab_SoldeApres,i]:=StrfMontant(Valeur(FListe.Cells[ab_SoldeAvant,i])-XX,0,V_PGI.OkDecV,'',True) ;
    END ;
  Q.Next ;
  END ;
Ferme(Q) ;
AfficheTotauxCalcul(1) ;
END ;

Procedure TFAffBQEED.AfficheFliste2(Avec : Boolean ) ;
BEGIN
WFliste2.Visible:=Avec ; FListe2.Visible:=Avec ; //Split.Visible:=Avec ; Split.Top:=FListe.Top+FListe.Height+1 ;
END ;

procedure TFAffBQEED.FListeDblClick(Sender: TObject);
Var i : Integer ;
begin
//if Fliste.Cells[0,1]='' then Exit ;
AfficheFliste2(True) ; NbFListe2Selected:=0 ;
MeMoLig:=FListe.Row ; 
FListe2.SortEnabled:=FALSE ; FListe2.SortedCol:=-1 ;
If Client Then TobMvt.PutGridDetail(Fliste2,TRUE,TRUE,'E_BANQUEPREVI;E_JOURNAL;E_GENERAL;E_AUXILIAIRE;E_DATECOMPTABLE;E_DATEECHEANCE;E_NUMEROPIECE;E_DEBIT;E_CREDIT;IND;',TRUE)
          Else TobMvt.PutGridDetail(Fliste2,TRUE,TRUE,'E_BANQUEPREVI;E_JOURNAL;E_GENERAL;E_AUXILIAIRE;E_DATECOMPTABLE;E_DATEECHEANCE;E_NUMEROPIECE;E_CREDIT;E_DEBIT;IND;',TRUE) ;
Fliste2.ColWidths[FListe2.ColCount-1]:=0 ;
For i:=1 To Fliste2.RowCount-1 Do
  BEGIN
  FListe2.Cells[Fliste2.ColCount-1,i]:=IntToStr(i-1) ;
  END ;
FListe2.SortEnabled:=TRUE ; FListe2.SortedCol:=0 ;
Fliste2.Refresh ;
end;

procedure TFAffBQEED.FListeTopLeftChanged(Sender: TObject);
begin
if Not FListe2.Visible then Exit ;
FListe.Row:=FListe.TopRow ;
if MemoLig=FListe.Row then Exit ;
MemoLig:=FListe.Row ;
end;

procedure TFAffBQEED.FListeRowEnter(Sender: TObject; Ou: Longint;
  var Cancel: Boolean; Chg: Boolean);
begin
if Not FListe2.Visible then Exit ;
MemoLig:=Ou ; 
end;

procedure TFAffBQEED.FListe2TopLeftChanged(Sender: TObject);
begin FListe2.Row:=Fliste2.TopRow ; end;

procedure TFAffBQEED.FListe2DblClick(Sender: TObject);
begin
If TobMvt=NIL Then Exit ;
end;

procedure TFAffBQEED.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ; FListe2.VidePile(True) ;
OldBrush.Free ; OldPen.Free ;
TOBMvt.Free ;
If Client Then RegSaveToolbarPos(Self,'AFFBQECLI') Else RegSaveToolbarPos(Self,'AFFBQEFOU') ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFAffBQEED.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFAffBQEED.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFAffBQEED.FindDialogFind(Sender: TObject);
begin
If CurGrid=2 Then
  BEGIN
  Rechercher(FListe2,FindDialog, FirstFind) ;
  END Else
  BEGIN
  Rechercher(FListe,FindDialog, FirstFind) ;
  END ;
end;

procedure TFAffBQEED.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TFAffBQEED.FListeEnter(Sender: TObject);
begin
CurGrid:=1 ; 
end;

procedure TFAffBQEED.FListe2Enter(Sender: TObject);
begin
CurGrid:=2 ;
end;

procedure TFAffBQEED.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
If OnNeSortPas Then CanClose:=FALSE ;
end;

procedure TFAffBQEED.bExportClick(Sender: TObject);
begin
if not ExJaiLeDroitConcept(ccExportListe,True) then exit ;
if SD.Execute then
  BEGIN
  If CurGrid=1 Then ExportGrid(FListe,Nil,SD.FileName,SD.FilterIndex,TRUE)
               Else ExportGrid(FListe2,Nil,SD.FileName,SD.FilterIndex,TRUE) ;
  END ;
end;

procedure TFAffBQEED.AfficheTotARepartir ;
Var St : String ;
BEGIN
St:=StrfMontant(TotRepart.Value,0,V_PGI.OkDecV,'',True) ;
PCumul.Caption:=' Totaux  ('+St+' à répartir)' ;
END ;

procedure TFAffBQEED.BSelectMvtClick(Sender: TObject);
Var XX,OldXX : Double ;
    i : Integer ;
    Cancel : Boolean ;
    C,L : Integer ;
begin
OldXX:=TotRepart.Value ;
XX:=SuiviBP(Client,TobMvt) ; If XX<>-1 Then TotRepart.Value:=XX ;
BCalcul.Enabled:=(TobMvt.Detail.Count>0) And (TotRepart.Value>0) ;
AfficheTotARepartir ;
If (Arrondi(OldXX-XX,V_PGI.OkDecV)<>0) And (Arrondi(XX,V_PGI.OkDecV)<>0) Then
  BEGIN
  CalculFait:=FALSE ;
  ReinitMontant ;
  END ;
For i:=1 To Fliste.RowCount-1 Do
  BEGIN
  C:=ab_pourcentSaisi ; l:=i ; FListeCellExit(Nil,C,L,Cancel);
  END ;
end;

procedure TFAffBQEED.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : Boolean ;
begin
OkG:=(Screen.ActiveControl=FListe) ; Vide:=(Shift=[]) ;
Case Key Of
  VK_RETURN : if ((OkG) and (Vide)) then KEY:=VK_TAB ;
  END ;
end;

procedure TFAffBQEED.FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var XX,YY : Double ;
begin
If (ACol<>ab_PourcentSaisi) And (ACol<>ab_MontantSaisi) Then Exit ;
if (ACol=ab_PourcentSaisi) Or (ACol=ab_MontantSaisi) Then
  FListe.Cells[ACol,ARow]:=StrfMontant(Valeur(FListe.Cells[ACol,ARow]),0,V_PGI.OkDecV,'',True) ;
If ACol=ab_PourcentSaisi Then
  BEGIN
  XX:=Valeur(FListe.Cells[ACol,ARow]) ; YY:=TotRepart.value*XX/100 ;
  FListe.Cells[ab_montantSaisi,ARow]:=StrfMontant(YY,0,V_PGI.OkDecV,'',True) ;
  END ;
If ACol=ab_MontantSaisi Then
  BEGIN
  XX:=Valeur(FListe.Cells[ACol,ARow]) ; YY:=0 ;
  If Arrondi(TotRepart.Value,V_PGI.OkDecV)<>0 Then YY:=(XX*100)/TotRepart.Value ;
  FListe.Cells[ab_PourcentSaisi,ARow]:=StrfMontant(YY,0,V_PGI.OkDecV,'',True) ;
  END ;
AfficheTotauxSaisi ;
end;

Function TFAffBQEED.TrouveLigne(Cpt : String) : Integer ;
Var i : Integer ;
BEGIN
Result:=0 ;
For i:=1 To FListe.RowCount-1 Do If FListe.Cells[ab_Compte,i]=Cpt Then BEGIN Result:=i ; Break ; END ;
END ;

procedure TFAffBQEED.Preventil (NomGrille : String) ;
Var QVentil : TQuery ;
    Cpt     : String ;
    Taux,TotR,X    : Double ;
    i       : Integer ;
    STV : String ;
BEGIN
StV:='SELECT * FROM VENTIL WHERE V_NATURE="ABP" AND V_COMPTE="'+NomGrille+'" '
    +'ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
QVentil:=OpenSQL(STV,True) ;
While Not QVentil.EOF do
    BEGIN
    Cpt:=QVentil.FindField('V_SECTION').AsString ; Taux:=QVentil.FindField('V_TAUXMONTANT').AsFloat ;
    i:=TrouveLigne(Cpt) ;
    TotR:=Arrondi(TotRepart.Value,V_PGI.OkDecV) ;
    If i>0 Then
      BEGIN
      FListe.Cells[ab_PourcentSaisi,i]:=StrfMontant(Taux,0,V_PGI.OkDecV,'',True) ;
      (*
      TotPSaisi:=Arrondi(TotPSaisi+Taux,V_PGI.OkDecV) ;
      *)
      X:=Arrondi(Taux*TotR/100.0,V_PGI.OkDecV) ;
      (*
      TotMSaisi:=Arrondi(TotMSaisi+X,V_PGI.OkDecV) ;
      *)
      FListe.Cells[ab_MontantSaisi,i]:=StrfMontant(X,0,V_PGI.OkDecV,'',True) ;
      END ;
    QVentil.Next ;
    END ;
Ferme(QVentil) ;
AfficheTotauxSaisi ;
END ;

procedure TFAffBQEED.CVTypeChange(Sender: TObject);
begin
if CVType.Value='' then exit ;
ReinitMontant ;
PreVentil(CVType.Value) ;
Fliste.SetFocus ;
end;

procedure TFAffBQEED.ClickSauveVentil ;
BEGIN
if Fliste.RowCount<=2 then Exit ;
PVentil.Left:=FListe.Left+(Fliste.Width-PVentil.Width) div 2 ;
PVentil.Visible:=True ; Y_CODEVENTIL.SetFocus ; ModeVentil:=True ;
Pages.Enabled:=False ; PFiltres.Enabled:=False ; FListe.Enabled:=FALSE ;
HPB.Enabled:=False ;
END ;

procedure TFAffBQEED.FinVentil ;
BEGIN
Pages.Enabled:=TRUE ; PFiltres.Enabled:=TRUE ; FListe.Enabled:=TRUE ;
HPB.Enabled:=TRUE ;
CVType.SetFocus ; PVentil.Visible:=False ; ModeVentil:=False ;
END ;

procedure TFAffBQEED.BNewVentilClick(Sender: TObject);
Var Lib       : String ;
    Existe    : boolean ;
    i,j         : integer ;
    Q,T       : TQuery ;
    Taux      : Double ;
begin
if ((Y_CODEVENTIL.Text='') or (Y_LIBELLEVENTIL.Text='')) then BEGIN HM.Execute(0,Caption,'') ; Exit ; END ;
Lib:=uppercase(Y_LIBELLEVENTIL.Text) ;
Q:=OpenSQL('Select * from CHOIXCOD Where CC_TYPE="VAB" AND (CC_CODE="'+Y_CODEVENTIL.Text+'" OR UPPER(CC_LIBELLE)="'+Lib+'")',True) ;
Existe:=Not Q.EOF ; Ferme(Q) ;
if Existe then BEGIN HM.Execute(1,Caption,'') ; Exit ; END ;
Try
 BeginTrans ;
 T:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="Ed#"',False) ;
 T.Insert ; InitNew(T) ;
 T.FindField('CC_TYPE').AsString:='VAB' ;
 T.FindField('CC_CODE').AsString:=Y_CODEVENTIL.Text ;
 T.FindField('CC_LIBELLE').AsString:=Y_LIBELLEVENTIL.Text ;
 T.Post ; Ferme(T) ;
 T:=OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE="Ed#"',False) ; j:=0 ;
 for i:=1 to Fliste.RowCount-1 do
   BEGIN
   Taux:=Arrondi(Valeur(FListe.Cells[ab_pourcentSaisi,i]),V_PGI.OkDecV) ;
   if taux<>0 then
     BEGIN
     T.Insert ; InitNew(T) ;
     T.FindField('V_NATURE').AsString:='ABP' ;
     T.FindField('V_COMPTE').AsString:=Y_CODEVENTIL.text ;
     T.FindField('V_SECTION').AsString:=FListe.Cells[ab_Compte,i] ;
     T.FindField('V_TAUXMONTANT').AsFloat:=Taux ;
     Inc(j) ; T.FindField('V_NUMEROVENTIL').AsInteger:=j ;
     T.Post ;
     END ;
   END ;
 Ferme(T) ;
 CommitTrans ;
Except
 Rollback ;
End ;
FinVentil ;
AvertirTable('CPVENTILAFFBQE') ;
CVType.DataType:='' ; CVType.DataType:='CPVENTILAFFBQE' ;
end;

procedure TFAffBQEED.FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
if ((ModeVentil) and (Key<>VK_F10) and (Key<>VK_ESCAPE)) then Exit ;
Case Key of
   VK_F10    : if ModeVentil then BNewVentilClick(Nil) ;
   VK_ESCAPE : if ModeVentil then FinVentil  ;
   END ;

end;

procedure TFAffBQEED.AfficheTotauxSaisi ;
Var XX,YY : Double ;
    i : Integer ;
BEGIN
XX:=0 ; YY:=0 ;
For i:=1 To Fliste.RowCount-1 Do
  BEGIN
  XX:=Arrondi(XX+Valeur(FListe.Cells[ab_PourcentSaisi,i]),V_PGI.OkDecV) ;
  YY:=Arrondi(YY+Valeur(FListe.Cells[ab_MontantSaisi,i]),V_PGI.OkDecV) ;
  END ;
FTotPSaisi.Value:=XX/100 ; FTotMSaisi.Value:=YY ;
END ;

Function TFAffBQEED.MontantBqe(Cpt : String) : Double ;
Var i : Integer ;
    XX,D,C,ZZ : Double ;
    Cpt1 : String ;
    TOBL : TOB ;
BEGIN
Result:=0 ; XX:=0 ;
If TobMvt.Detail=Nil THEN Exit ;
For i:=0 To TobMvt.Detail.Count-1 Do
  BEGIN
  TOBL:=TobMvt.Detail[i] ; Cpt1:=TOBL.GetValue('E_BANQUEPREVI') ;
  If Cpt1=Cpt Then
    BEGIN
    D:=TOBL.GetValue('E_DEBIT') ; C:=TOBL.GetValue('E_CREDIT') ;
    If Client Then ZZ:=XX+D-C Else ZZ:=XX+C-D ;
    XX:=Arrondi(ZZ,V_PGI.OkDecV) ;
    END ;
  END ;
Result:=XX ;
END ;

procedure TFAffBQEED.AfficheTotauxCalcul(What : Integer) ;
Var TotP,TotM,P,M,MB : Double ;
    i : Integer ;
BEGIN
TotM:=0 ; TotP:=0 ;
If What=1 Then
  BEGIN
  TotRepart.Value:=0 ;
  For i:=1 To Fliste.RowCount-1 Do
    BEGIN
    M:=Valeur(FListe.Cells[ab_montantCalcul,i]) ;
    TotRepart.Value:=Arrondi(TotRepart.Value+M,V_PGI.OkDecV) ;
    END ;
  END ;
For i:=1 To Fliste.RowCount-1 Do
  BEGIN
  If What=0 Then
    BEGIN
    M:=MontantBqe(Fliste.Cells[ab_compte,i]) ;
    FListe.Cells[ab_MontantCalcul,i]:=StrfMontant(M,0,V_PGI.OkDecV,'',True) ;
    END Else M:=Valeur(FListe.Cells[ab_montantCalcul,i]) ;
  MB:=Valeur(FListe.Cells[ab_SoldeAvant,i]) ;
  If Client Then FListe.Cells[ab_SoldeApres,i]:=StrfMontant(MB+M,0,V_PGI.OkDecV,'',True)
            Else FListe.Cells[ab_SoldeApres,i]:=StrfMontant(MB-M,0,V_PGI.OkDecV,'',True) ;
  TotM:=Arrondi(TotM+M,V_PGI.OkDecV) ; P:=0 ;
  If Arrondi(TotRepart.Value,V_PGI.OkDecV)<>0 Then P:=(M*100)/TotRepart.Value ;
  FListe.Cells[ab_PourcentCalcul,i]:=StrfMontant(P,0,V_PGI.OkDecV,'',True) ;
  TotP:=Arrondi(TotP+P,V_PGI.OkDecV) ;
  END ;
FTotPCalcul.Value:=TotP/100 ; FTotMCalcul.Value:=TotM ;
END ;

Procedure TFAffBQEED.ResizeLesTotaux(Sender: TObject) ;
Var i,t : Integer ;
BEGIN
t:=0 ;
FTotPSaisi.Width:=Fliste.ColWidths[ab_pourcentSaisi]+1+1 ;
For i:=0 To ab_pourcentSaisi-1 Do t:=t+Fliste.ColWidths[i]+1 ; FTotPSaisi.Left:=t ;
FTotMSaisi.Width:=Fliste.ColWidths[ab_MontantSaisi]+1+1 ;
t:=t+Fliste.ColWidths[ab_PourcentSaisi]+1 ; FTotMSaisi.Left:=t ;
FTotPCalcul.Width:=Fliste.ColWidths[ab_PourcentCalcul]+1+1 ;
t:=t+Fliste.ColWidths[ab_MontantSaisi]+1 ; FTotPCalcul.Left:=t ;
FTotMCalcul.Width:=Fliste.ColWidths[ab_MontantCalcul]+1+1 ;
t:=t+Fliste.ColWidths[ab_PourcentCalcul]+1 ; FTotMCalcul.Left:=t ;
END ;

procedure TFAffBQEED.FListeKeyPress(Sender: TObject; var Key: Char);
begin
If (Fliste.Col<>ab_PourcentSaisi) And (Fliste.Col<>ab_MontantSaisi) Then Exit ;
If key in ['0'..'9',',','.',chr(VK_BACK)]=FALSE Then BEGIN Key:=#0 ; Exit ; END ;
end;

function  TFAffBQEED.ConstruitListe(var LM : T_DX; var LP,LB : T_IX) : integer ;
var j : integer ;
    BanquePrevi : String ;
    Indice : Integer ;
    TOBL : TOB ;
begin
Indice:=0 ;
FillChar(LM,Sizeof(LM),#0) ; FillChar(LP,Sizeof(LP),#0) ; FillChar(LB,Sizeof(LB),#0) ;
For j:=0 To TobMvt.Detail.Count-1 Do
  begin
  TOBL:=TOBMvt.Detail[j] ; If TOBL=NIL Then Continue ;
  BanquePrevi:=TobL.GetValue('E_BANQUEPREVI') ;
  if BanquePrevi='' then
    begin
    If Client Then LM[j]:=Arrondi(TOBL.GetValue('E_DEBIT')-TOBL.GetValue('E_CREDIT'),V_PGI.OkDecV)
              Else LM[j]:=Arrondi(TOBL.GetValue('E_CREDIT')-TOBL.GetValue('E_DEBIT'),V_PGI.OkDecV) ;
    LP[Indice]:=0 ; LB[Indice]:=j ; Inc(Indice) ;
    end ;
  end ;
Result:=Indice ;
end;

procedure TFAffBQEED.AffichePendantCalcul(R : Integer ; D,C : Double) ;
Var XX,YY,ZZ : Double ;
BEGIN
XX:=Valeur(FListe.Cells[ab_MontantCalcul,R]) ; YY:=0 ;
If Client Then ZZ:=D-C Else ZZ:=C-D ;
XX:=Arrondi(XX+ZZ,V_PGI.OkDecV) ;
FListe.Cells[ab_MontantCalcul,R]:=StrfMontant(XX,0,V_PGI.OkDecV,'',True) ;
If Arrondi(TotRepart.Value,V_PGI.OkDecV)<>0 Then YY:=(XX*100)/TotRepart.Value ;
FListe.Cells[ab_PourcentCalcul,R]:=StrfMontant(YY,0,V_PGI.OkDecV,'',True) ;
Application.ProcessMessages ;
END ;

Function TFAffBQEED.TenteAffectation(TOBL : TOB ; What : Integer) : Boolean ;
(*
 What 0 : Par épuisement
      10 : à 1O%
      20 : à 30% etc...
      >100 : Au mieux pour compléter à What-100 près
*)
Var D,C,S : Double ;
    i,j : Integer ;
    MS,MC,MA : Double ;
BEGIN
Result:=FALSE ; j:=-1 ;
D:=TOBL.GetValue('E_DEBIT') ; C:=TOBL.GetValue('E_CREDIT') ;
If Client Then S:=D-C Else S:=C-D ;
For i:=1 To Fliste.RowCount-1 Do
  BEGIN
  MS:=Arrondi(Valeur(Fliste.Cells[ab_montantSaisi,i]),V_PGI.OkDecV) ;
  MC:=Arrondi(Valeur(Fliste.Cells[ab_montantCalcul,i]),V_PGI.OkDecV) ;
  If What=0 Then
    BEGIN
    If MC+S<MS Then BEGIN j:=i ; Result:=TRUE ; Break ; END ;
    END Else If What>100 then
    BEGIN
    MA:=MS*(What-100)/100 ;
    If S<=MA Then BEGIN j:=i ; Result:=TRUE ; Break ; END ;
    END Else
    BEGIN
    MA:=MS*What/100 ;
    If (MC+S>MS-MA) And (MC+S<MS+MA) Then BEGIN j:=i ; Result:=TRUE ; Break ; END ;
    END ;
  END ;
If Result And (j>=0) Then
  BEGIN
  TOBL.PutValue('E_BANQUEPREVI',FListe.Cells[ab_Compte,j]) ;
  AffichePendantCalcul(j,TobL.GetValue('E_DEBIT'),TobL.GetValue('E_CREDIT')) ;
  END ;
END ;

Function TFAffBQEED.FiniCalcul(What : Integer) : Boolean ;
Var TOBL : TOB ;
    i : Integer ;
    OkAffecte : Boolean ;
BEGIN
OkAffecte:=TRUE ;
For i:=0 To TobMvt.Detail.Count-1 Do
  BEGIN
  TOBL:=TOBMvt.Detail[i] ; If TOBL=NIL Then Continue ;
  If TOBL.GetValue('E_BANQUEPREVI')='' Then If Not TenteAffectation(TOBL,0) Then OkAffecte:=FALSE ;
  END ;
Result:=OkAffecte ;
END ;

Procedure TFAffBQEED.GoCalcul ;
Var LM    : T_DX ;
    LP,LB    : T_IX ;
    Solde,Delta : double ;
    InfoX : REC_AUTOX ;
    i,j : integer ;
    TOBL : TOB ;
begin
//
If TobMvt=Nil Then Exit ;
If TobMvt.Detail.Count<=0 Then Exit ;
Fillchar(InfoX,SizeOf(InfoX),#0) ;
InfoX.Nival:=TobMvt.Detail.Count ; InfoX.Decim:=V_PGI.OkDecV ;
InfoX.Temps:=0 ; InfoX.Unique:=FALSE ;
For i:=1 To Fliste.RowCount-1 Do
  BEGIN
  Solde:=Arrondi(Valeur(Fliste.Cells[ab_montantSaisi,i]),V_PGI.OkDecV) ;
  Delta:=Solde/10 ;
  If Solde<>0 Then
    BEGIN
    InfoX.NbD:=ConstruitListe(LM,LP,LB) ;
    if LettrageApproche(Solde,Delta,LM,LP,InfoX)<>0 then
      begin
      for j:=0 to InfoX.NbD-1 do
        if LP[j]<>0 then
          begin
          TOBL:=TOBMvt.Detail[LB[j]] ; If TOBL=NIL Then Continue ;
          TOBL.PutValue('E_BANQUEPREVI',FListe.Cells[ab_Compte,i]) ;
          AffichePendantCalcul(i,TobL.GetValue('E_DEBIT'),TobL.GetValue('E_CREDIT')) ;
          end;
      end;
    END ;
  END ;
end;

procedure TFAffBQEED.AnnuleBanquePrevi ;
Var i : Integer ;
    TOBL : TOB ;
BEGIN
For i:=0 To TobMvt.Detail.Count-1 Do
  BEGIN
  TOBL:=TOBMvt.Detail[i] ; If TOBL=NIL Then Continue ;
  If TOBL.GetValue('E_BANQUEPREVI')<>'' Then TOBL.PutValue('E_BANQUEPREVI','') ;
  END ;
END ;

procedure TFAffBQEED.BCalculClick(Sender: TObject);
begin
CalculFait:=TRUE ;
If HM.Execute(3,Caption,'')<>mrYes Then Exit ;
If HM.Execute(4,Caption,'')=mrYes Then AnnuleBanquePrevi;
ReinitMontant ;
GoCalcul ;
If Not FiniCalcul(0) Then
 If Not FiniCalcul(10) Then
  If Not FiniCalcul(30) Then
    If Not FiniCalcul(50) Then
     If Not FiniCalcul(110) Then
       If Not FiniCalcul(130) Then
        If Not FiniCalcul(200) Then ;
AfficheTotauxCalcul(0) ;
end;

Procedure TFAffBQEED.ValideLeLot ;
BEGIN
TobMvt.UpdateDB(TRUE) ;
END ;

procedure TFAffBQEED.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    TOBL : TOB ;
    CodeLot : String ;
begin
If Action=taCreat Then
  BEGIN
  CodeLot:=E_NOMLOT.Text ;
  If SaisieCodelot(CodeLot) Then
    BEGIN
    E_NOMLOT.Text:=CodeLot ;
    If HM.Execute(5,Caption,'')<>mrYes Then Exit ;
    For i:=0 To TobMvt.Detail.Count-1 Do BEGIN TOBL:=TobMvt.Detail[i] ; TOBL.PutValue('E_NOMLOT',E_NOMLOT.Text) ; END ;
    If Transactions(ValideLeLot,5)<>oeOk Then HM.Execute(7,Caption,'') ;
    END ;
  END ;
(**)
end;

procedure TFAffBQEED.FListe2MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
Var ACol,ARow : Integer ;
begin
FListe2.MouseToCell(X,Y,ACol,ARow) ;
If ARow=0 Then Exit ;
if ((ssLeft in Shift)) then BEGIN Fliste2.BeginDrag(True) ; OkDrop:=FALSE ; NbFListe2Selected:=FListe2.NbSelected ; END ;
end;

procedure TFAffBQEED.FListe2DragOver(Sender, Source: TObject; X,Y: Integer; State: TDragState; var Accept: Boolean);
begin
Accept:=False ;
end;

Procedure TFAffBQEED.UpdateBanquePrevi(Cpt : String) ;
Var i,j : Integer ;
    TOBL : TOB ;
BEGIN
Fliste2.SortEnabled:=FALSE ;
If FListe2.NbSelected>0 Then
  BEGIN
  For i:=0 To FListe2.RowCount-1 Do If FListe2.IsSelected(i) Then
    BEGIN
    j:=StrToInt(Fliste2.Cells[Fliste2.ColCount-1,i]) ;
    TOBL:=TobMvt.Detail[j] ;
    If TOBL<>NIL Then TOBL.PutValue('E_BANQUEPREVI',Cpt) ;
    Fliste2.Cells[0,i]:=Cpt ;
    END ;
  END Else
  BEGIN
  j:=StrToInt(Fliste2.Cells[Fliste2.ColCount-1,FListe2.Row]) ;
  TOBL:=TobMvt.Detail[j] ;
  If TOBL<>NIL Then TOBL.PutValue('E_BANQUEPREVI',Cpt) ;
  Fliste2.Cells[0,Fliste2.Row]:=Cpt ;
  END ;
AfficheTotauxCalcul(0) ;
Fliste2.SortEnabled:=TRUE ;
FListe2.Refresh ;
END ;

procedure TFAffBQEED.FListe2EndDrag(Sender, Target: TObject; X,Y: Integer);
Var ACol,ARow : integer ;
    Cpt : String ;
    TC : TControl ;
begin
// Maj à faire
TC:=TControl(Target) ; If (TC=NIL) Or (TC.Name<>FListe.Name) Then OkDrop:=FALSE ;
If OkDrop Then
  BEGIN
  FListe.MouseToCell(X,Y,ACol,ARow) ;
  Cpt:=Fliste.Cells[ab_Compte,ARow] ;
  UpdateBanquePrevi(Cpt) ;
  OkDrop:=FALSE ;
  END  ;
end;

procedure TFAffBQEED.FListeDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
Var ACol,ARow : integer ;
begin
Accept:=True ;
If NbFListe2Selected>0 Then FListe.DragCursor:=crMultiDrag Else FListe.DragCursor:=crDrag ;
FListe.MouseToCell(X,Y,ACol,ARow) ;
if ARow<FListe.FixedRows then BEGIN Accept:=False ; Exit ; END ;
If ARow>Fliste.RowCount-1 then BEGIN Accept:=False ; Exit ; END ;
Fliste.Row:=ARow ;
end;

procedure TFAffBQEED.FListeDragDrop(Sender, Source: TObject; X,Y: Integer);
begin
OkDrop:=FALSE ;
If HM.Execute(8,Caption,'')=mrYes Then OkDrop:=TRUE ;
end;

Procedure TFAffBQEED.TraiteFListe2(What : Integer) ;
Var OkOk : Boolean ;
    i : Integer ;
    Cpt : String ;
BEGIN
Cpt:=FListe2.Cells[ab_Compte,FListe2.ROW] ;
If What<2 Then FListe2.ClearSelected ; If What=0 Then Exit ;
For i:=1 To Fliste2.RowCount-1 Do
  BEGIN
  If What=1 Then
    BEGIN
    FListe2.FlipSelection(i)  ;
    END Else
    BEGIN
    OkOk:=FListe2.Cells[ab_Compte,i]=Cpt ;
    If OkOk Then
      BEGIN
      If What=2 Then
        BEGIN
        If FListe2.IsSelected(i) Then FListe2.FlipSelection(i) ;
        END Else
        BEGIN
        If (Not FListe2.IsSelected(i)) Then FListe2.FlipSelection(i) ;
        END ;
      END ;
    END ;
  END ;
END ;

procedure TFAffBQEED.bSelectAllClick(Sender: TObject);
begin
TraiteFListe2(1) ;
end;

procedure TFAffBQEED.BSelectBqeClick(Sender: TObject);
begin
TraiteFListe2(2) ;
end;

procedure TFAffBQEED.bPoubelleClick(Sender: TObject);
begin
If HM.Execute(9,Caption,'')<>mrYes Then Exit ;
UpdateBanquePrevi('') ;
end;

procedure TFAffBQEED.BDeselectBqeClick(Sender: TObject);
begin
TraiteFListe2(2) ;
end;

procedure TFAffBQEED.BDeselectAllClick(Sender: TObject);
begin
TraiteFListe2(0) ;
end;

procedure TFAffBQEED.GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var OnGrasse : Boolean ;
    OkCol : Boolean ;
begin
if (ARow<FListe.FixedRows) then Exit ;
OnGrasse:=FALSE ; OkCol:=FALSE ;
If Action=taCreat Then
  BEGIN
  OkCol:=(ACol=ab_PourcentCalcul) Or (ACol=ab_MontantCalcul) Or (ACol=ab_SoldeApres) ;
  OnGrasse:=CalculFait ;
  END ;
If OnGrasse And OkCol then Canvas.Font.Style:=[fsBold] ;
end ;

procedure TFAffBQEED.SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
var R : TRect ;
begin
Canvas.Brush.Color := Fliste.FixedColor ;
Canvas.Brush.Style := bsBDiagonal ;
Canvas.Pen.Color   := Fliste.FixedColor ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psClear ;
Canvas.Pen.Width   := 1 ;
R:=Fliste.CellRect(ACol, ARow) ;
Canvas.Rectangle(R.Left, R.Top, R.Right+1, R.Bottom+1) ;
end ;

procedure TFAffBQEED.PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var OnGrise : Boolean ;
    OkCol : Boolean ;
begin
if (ARow<FListe.FixedRows) then Exit ;
OnGrise:=FALSE ; OkCol:=FALSE ;
If Action=taCreat Then
  BEGIN
  OkCol:=(ACol=ab_PourcentCalcul) Or (ACol=ab_MontantCalcul) Or (ACol=ab_SoldeApres) ;
  OnGrise:=Not CalculFait ;
  END ;
if OnGrise And OkCOl then SetGridGrise(ACol, ARow, Canvas);
end ;


procedure TFAffBQEED.BSauveVentilClick(Sender: TObject);
begin
ClickSauveVentil ;
end;

end.
