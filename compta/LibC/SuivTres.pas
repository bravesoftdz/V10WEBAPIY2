Unit SuivTres;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Spin, StdCtrls, ExtCtrls, Hctrls, Grids, Buttons, ComCtrls, Menus, HDB,
  hmsgbox, DB,
  AGLInit,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  Hqry, HEnt1, Mask, Hcompte, Ent1, SaisUtil,
  HSysMenu, HTB97, HPanel, UiUtil,UObjFiltres, ADODB {SG6 17/11/04 Gestion des filtres FQ 14976} ;

Procedure SuiviTresoParPeriode ;

Type Tablo=Array[1..31] Of TDateTime ;
Type TTotal = Array [1..31] of Double ;
Type CString=Class
         Compte : String ;
      END ;

type
  TFSuiTre = class(TForm)
    Pages: TPageControl;
    PCritere: TTabSheet;
    Bevel1: TBevel;
    TCpte: THLabel;
    PPeriode: TTabSheet;
    Bevel2: TBevel;
    FListe: THGrid;
    Cpte: TEdit;
    TQualPie: THLabel;
    QualPie: THValComboBox;
    TDev: THLabel;
    Dev: THValComboBox;
    TEtab: THLabel;
    Etab: THValComboBox;
    PFiltres: TToolWindow97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    FindDialog: TFindDialog;
    FPaie: THMultiValComboBox;
    TFPaie: THLabel;
    Tcol: TLabel;
    NbCol: TSpinEdit;
    Perio: THValComboBox;
    TPerio: TLabel;
    QCpte: THQuery;
    MsgBox: THMsgBox;
    BPerio: TToolbarButton97;
    Pdate: TPanel;
    Ptop: TPanel;
    Pbouton1: TPanel;
    BVal1: TToolbarButton97;
    BFer1: TToolbarButton97;
    Baide1: TToolbarButton97;
    TFP1: TLabel;
    FP1: TMaskEdit;
    TFP2: TLabel;
    FP2: TMaskEdit;
    TFP3: TLabel;
    FP3: TMaskEdit;
    TFP4: TLabel;
    FP4: TMaskEdit;
    TFP5: TLabel;
    FP5: TMaskEdit;
    TFP6: TLabel;
    FP6: TMaskEdit;
    TFP7: TLabel;
    FP7: TMaskEdit;
    TFP8: TLabel;
    FP8: TMaskEdit;
    TFP9: TLabel;
    FP9: TMaskEdit;
    TFP10: TLabel;
    FP10: TMaskEdit;
    TFP11: TLabel;
    FP11: TMaskEdit;
    TFP12: TLabel;
    FP12: TMaskEdit;
    TFP13: TLabel;
    FP13: TMaskEdit;
    TFP14: TLabel;
    FP14: TMaskEdit;
    TFP15: TLabel;
    FP15: TMaskEdit;
    TFP16: TLabel;
    FP16: TMaskEdit;
    TFP17: TLabel;
    FP17: TMaskEdit;
    TFP18: TLabel;
    FP18: TMaskEdit;
    TFP19: TLabel;
    FP19: TMaskEdit;
    TFP20: TLabel;
    FP20: TMaskEdit;
    FP21: TMaskEdit;
    TFP21: TLabel;
    TFP22: TLabel;
    FP22: TMaskEdit;
    TFP23: TLabel;
    FP23: TMaskEdit;
    TFP24: TLabel;
    FP24: TMaskEdit;
    Bevel3: TBevel;
    Cpte1: THCpteEdit;
    TPerio1: TLabel;
    Perio1: THValComboBox;
    QSom: THQuery;
    QEcr: THQuery;
    HMTrad: THSystemMenu;
    BGL: TToolbarButton97;
    PopZ: TPopupMenu;
    BMP: TToolbarButton97;
    BGene: TToolbarButton97;
    RgAnalyse: TRadioGroup;
    PAffichage: TTabSheet;
    Bevel4: TBevel;
    RgMp: TRadioGroup;
    CAfdev: TRadioGroup;
    RgDat: TRadioGroup;
    MEDat: TMaskEdit;
    LDat: TLabel;
    CbChoix: TRadioGroup;
    CbMp: THValComboBox;
    CbJouvre: TCheckBox;
    ParamFont: TFontDialog;
    BFont: TToolbarButton97;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BRechercher: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BReduire: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    Dock971: TDock97;
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure CpteKeyPress(Sender: TObject; var Key: Char);
    procedure BChercheClick(Sender: TObject);
    procedure BPerioClick(Sender: TObject);
    procedure BFer1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure PerioChange(Sender: TObject);
    procedure CbChoixClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
    procedure DevChange(Sender: TObject);
    procedure NbColChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BGeneClick(Sender: TObject);
    procedure BMPClick(Sender: TObject);
    procedure RgAnalyseClick(Sender: TObject);
    procedure RgDatClick(Sender: TObject);
    procedure BVal1Click(Sender: TObject);
    procedure BFontClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure Baide1Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    ObjFiltre:TObjFiltre; //SG6 17/11/04 Gestion des filtres FQ 14976 
    WMinX,WMinY : Integer ;
    FNomFiltre : String ;
    FirstFind : boolean;
    OldCon    : Longint ;
    LQPie,LEtab,LDev,LModP,LeAndEcr,LeAndGene : String ;
    LRgMp : Byte ;
    TabD,TabJ : Tablo ;
    OldBrush     : TBrush ;
    OldPen       : TPen ;
    Ladev : RDevise ;
    CritQp,CritMp,CritEtab,CritDev : String ;
    ParPerio : Boolean ;
    LaDat : String ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure AffectelesVars ;
    Function  FaitWherePourGene : Boolean ;
    Procedure AutoriseLesDates ;
    Procedure FaitWherePourTypePiece ;
    Procedure FaitWherePourEtab ;
    Procedure FaitWherePourDevise ;
    Procedure FaitWherePourModePaie ;
    Function  FabriqueLaRequeteSomDeb : String ;
    Function  FabriqueLaRequeteSomCre : String ;
    Function  FabriqueLaRequeteEcr : String ;
    Procedure RunLeCalcul(Typ : String) ;
    Procedure InitLeGrille ;
    Procedure RempliTabloJour ;
    Procedure FermeChoixDate ;
    Function  FaitTotalCompte(Var T: TTotal ; St : String ; Ind : Byte) : Boolean ;
  public
    { Déclarations publiques }
  end;


implementation


Uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPTypeCons,
  {$ENDIF MODENT1}
  PrintDBG, Filtre, HStatus, CalCole, CpteUtil, UtilEdt, CritEdt,
  CPGeneraux_TOM,
  uTOFCPGLGENE,
  FichComm ;

{$R *.DFM}

Procedure SuiviTresoParPeriode ;
var FSuiTre : TFSuiTre ;
    PP : THPanel ;
    Composants : TControlFiltre; //SG6   Gestion des Filtes 17/11/04   FQ 14976
BEGIN
PP:=FindInsidePanel ;
FSuiTre:=TFSuiTre.Create(Application) ;

//SG6 17/11/04 Gestion des Filtres      FQ 14976
Composants.PopupF   := FSuiTre.POPF;
Composants.Filtres  := FSuiTre.FFILTRES;
Composants.Filtre   := FSuiTre.BFILTRE;
Composants.PageCtrl := FSuiTre.Pages;
FSuiTre.ObjFiltre := TObjFiltre.Create(Composants, '');

FSuiTre.FNomFiltre:='SUIVITRESO' ;
if PP=Nil then
   BEGIN
    Try
     FSuiTre.ShowModal ;
    Finally
     FSuiTre.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FSuiTre,PP) ;
   FSuiTre.Show ;
   END ;
END ;

procedure TFSuiTre.FormResize(Sender: TObject);
begin
if Pdate.Visible then
   BEGIN
   Pdate.Left:=(Width div 2)-(Pdate.Width div 2) ;
   Pdate.Top:=(Height div 2)-(Pdate.Height div 2) ;
   END ;
end;

procedure TFSuiTre.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFSuiTre.BCreerFiltreClick(Sender: TObject);
begin NewFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFSuiTre.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFSuiTre.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre(FNomFiltre,FFiltres) ; end;

procedure TFSuiTre.BRenFiltreClick(Sender: TObject);
begin RenameFiltre(FNomFiltre,FFiltres) ; end;

procedure TFSuiTre.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres,Pages) ; end;

procedure TFSuiTre.FFiltresChange(Sender: TObject);
begin LoadFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFSuiTre.BRechercherClick(Sender: TObject);
begin FirstFind:=true; FindDialog.Execute ; end;

procedure TFSuiTre.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
LQPie:='' ; LEtab:='' ; LDev:='' ; LModP:='' ;
FNomFiltre:='' ; PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
OldBrush:=TBrush.Create ; OldPen:=TPen.Create ;
end;

procedure TFSuiTre.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FreeAndNil(ObjFiltre); //SG6 17/11/04 Gestion des filtres FQ 14976
FListe.VidePile(True) ; OldBrush.Free ; OldPen.Free ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFSuiTre.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFSuiTre.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFSuiTre.FormShow(Sender: TObject);
begin
ObjFiltre.FFI_TABLE:=FNomFiltre; //SG6 17/11/04 Gestion des filtres FQ 14976
ObjFiltre.Charger;               

PDate.Visible:=False ; Etab.ItemIndex:=0 ;
Dev.ItemIndex:=0 ; QualPie.ItemIndex:=0 ; Perio.ItemIndex:=0 ; FPaie.Text:=Traduirememoire('<<Tous>>') ;
Perio1.ItemIndex:=0 ; RgMp.ItemIndex:=0 ; Pages.ActivePage:=Pages.Pages[0] ;
if Not V_PGI.Halley then V_PGI.DateEntree:=Encodedate(1997,12,31) ;// A Virer dans Halley
MEDat.Text:=DateToStr(V_PGI.DateEntree) ;
DevChange(Nil) ; CbChoixClick(Nil) ; RgAnalyseClick(Nil) ;
RgDatClick(Nil) ; NbColChange(Nil) ;
ParamFont.Font.Style:=ParamFont.Font.Style+[fsBold] ;
ParamFont.Font.Name:='MS Sans Serif' ; ParamFont.Font.Size:=8 ;
ParamFont.Font.Color:=clWindowText ; ;
PositionneEtabUser(Etab, False); // 15090
end;

Procedure TFSuiTre.AffectelesVars ;
BEGIN
LQPie:=QualPie.Value ; LDev:=Dev.Value ; LEtab:=Etab.Value ; LModP:=FPaie.Text ;
LRgMp:=RgMp.ItemIndex ; BMP.Enabled:=(LRgMp<>0) ; BGene.Enabled:=(LRgMp<>2) ;
FaitWherePourTypePiece ; FaitWherePourEtab ; FaitWherePourDevise ; FaitWherePourModePaie ;
END ;

procedure TFSuiTre.CpteKeyPress(Sender: TObject; var Key: Char);
begin
if Not(Key in ['0'..'z',':',';',#8,#9,#13,VH^.Cpta[fbGene].Cb]) then Key:=#0 ;
end;

Function TFSuiTre.FaitWherePourGene : Boolean ;
Var St : String ;
    i : Byte ;
BEGIN
Result:=True ;
If CAfDev.ItemIndex=1 Then // GP le 21/06/99
  BEGIN
  If Dev.ItemIndex=0 Then
    BEGIN
    MsgBox.Execute(15,'','') ; Result:=False ; Pages.ActivePage:=Pages.Pages[0] ; Dev.SetFocus ; Exit ;
    END ;
  END ;
if CbChoix.ItemIndex=1 then
   BEGIN
   if Cpte.Text='' then
      BEGIN
      MsgBox.Execute(0,'','') ; Result:=False ; Exit ;
      END ;
   St:=Cpte.Text ;
   END else
   BEGIN
   if Cpte1.Text='' then
      BEGIN
      MsgBox.Execute(0,'','') ; Result:=False ; Exit ;
      END ;
   St:=Cpte1.Text ;
   END ;
LeAndGene:=AnalyseCompte(St,fbGene,False,False) ;
if LeAndGene<>'' then
   BEGIN
   QCpte.Close ;
   if QCpte.Sql.Count>2 then QCpte.Sql.Delete(QCpte.Sql.Count-1) ;
   QCpte.Sql.Add('AND '+LeAndGene) ;
   LeAndEcr:=LeAndGene ;
   While Pos('G_GENERAL',LeAndEcr)>0 do
      BEGIN
      i:=Pos('G_GENERAL',LeAndEcr) ; Delete(LeAndEcr,i,9) ; Insert('E_CONTREPARTIEGEN',LeAndEcr,i) ;
      END ;
   END ;
END ;

procedure TFSuiTre.BChercheClick(Sender: TObject);
Var i : Byte ;
begin
QCpte.Close ; QCpte.Sql.Clear ;
QCpte.Sql.Add('Select G_GENERAL From GENERAUX Where (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI")') ;
if Not FaitWherePourGene then Exit ;
ChangeSql(QCpte) ; AffectelesVars ;
if Not ParPerio then RempliTabloJour ;
QEcr.Close ; QEcr.Sql.Clear ; QEcr.Sql.Add(FabriqueLaRequeteEcr) ; ChangeSQL(QEcr) ; QEcr.Prepare ;
InitLeGrille ;
RunLeCalcul('RECETTE') ;
if FListe.RowCount>2 then
   BEGIN
   FListe.RowCount:=FListe.RowCount+1 ;
   for i:=0 to FListe.ColCount-1 do FListe.Cells[i,FListe.RowCount-1]:='*' ;
   FListe.RowCount:=FListe.RowCount+1 ;
   END ;
RunLeCalcul('DEPENSE') ;
end;

procedure TFSuiTre.BPerioClick(Sender: TObject);
begin
Pages.Enabled:=False ; PFiltres.Enabled:=False ; FListe.Enabled:=False ;
HPB.Enabled:=False ; OldCon:=HelpContext ; HelpContext:=7625100 ; 
Pdate.Left:=(Width div 2)-(Pdate.Width div 2) ; Pdate.Top:=(Height div 2)-(Pdate.Height div 2) ;
AutoriseLesDates ; PerioChange(Perio) ; PDate.Visible:=True ;
PDate.BringToFront ;
end;

Procedure TFSuiTre.FermeChoixDate ;
BEGIN
Pages.Enabled:=True ; PFiltres.Enabled:=True ; FListe.Enabled:=True ;
HPB.Enabled:=True ; Pdate.Visible:=False ;
HelpContext:=OldCon ;
END ;

procedure TFSuiTre.BVal1Click(Sender: TObject);
begin FermeChoixDate ; PerioChange(Perio) ; end;

procedure TFSuiTre.BFer1Click(Sender: TObject);
begin FermeChoixDate ; end;

procedure TFSuiTre.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin if Pdate.Visible then CanClose:=False ; end;

Procedure TFSuiTre.AutoriseLesDates ;
Var i : Integer ;
BEGIN
for i:=1 to 24 do
   BEGIN
   TLabel(FindComponent('TFP'+IntToStr(i))).Visible:=(i<((NbCol.Value*2)+1)) ;
   TMaskEdit(FindComponent('FP'+IntToStr(i))).Visible:=(i<((NbCol.Value*2)+1)) ;
   END ;
END ;

procedure TFSuiTre.NbColChange(Sender: TObject);
begin
if Not ParPerio then Exit ;
TMaskEdit(FindComponent('FP'+IntToStr(NbCol.Value*2))).Text:=DateToStr(V_PGI.DateEntree) ;
PerioChange(Perio) ;
end;

procedure TFSuiTre.PerioChange(Sender: TObject);
Var a,m,j,JMax : Word ;
    FinDat,DebDat : TDateTime ;
    i,k,Choix : Byte ;
begin
if THValComboBox(Sender).Name='Perio' then Perio1.ItemIndex:=Perio.ItemIndex
                                      else Perio.ItemIndex:=Perio1.ItemIndex ;
FillChar(TabD,SizeOf(TabD),0) ;
k:=NbCol.Value*2 ; Choix:=Perio.ItemIndex ;
if Perio.ItemIndex<0 then Exit ; 
FinDat:=FinDeMois(StrToDate(TMaskEdit(FindComponent('FP'+IntToStr(k))).Text)) ;
DebDat:=DebutDeMois(PlusMois(FinDat,-Choix)) ;
TabD[k]:=StrToDate(TMaskEdit(FindComponent('FP'+IntToStr(k))).Text) ;
Case Choix of
  0,1,2,
  3,4,5 : for i:=k Downto 1 do
             BEGIN
             if i Mod(2)=0 then
                BEGIN
                if i<>k then TabD[i]:=FinDat ;
                if i>1 then TabD[i-1]:=DebDat ;
                FinDat:=DebDat-1 ;
                DebDat:=DebutDeMois(PlusMois(FinDat,-Choix)) ;
                END ;
             END ;
  6     : for i:=k Downto 1 do
             BEGIN
             if i Mod(2)=0 then
                BEGIN
                if i<>k then TabD[i]:=FinDat ;
                DecodeDate(TabD[i],a,m,j) ;
                JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                if i>1 then
                  BEGIN
                  if j<=15 then TabD[i-1]:=(TabD[i]-15)+1 else TabD[i-1]:=(TabD[i]-(JMax-15))+1 ;
                  END ;
                if i>1 then FinDat:=TabD[i-1]-1 ;
               END ;
             END ;
  7     : for i:=k Downto 1 do
            BEGIN
            if i Mod(2)=0 then
               BEGIN
               if i<>k then TabD[i]:=FinDat ;
               DebDat:=TabD[i]-6 ; FinDat:=DebDat-1 ;
               if i>1 then TabD[i-1]:=DebDat ;
               END ;
            END ;
  End ;
for i:=1 to 24 do
   TMaskEdit(FindComponent('FP'+InttoStr(i))).Text:=DateToStr(TabD[i]) ;
end;

procedure TFSuiTre.CbChoixClick(Sender: TObject);
begin
Cpte1.Visible:=(CbChoix.ItemIndex=0) ; Cpte.Visible:=(CbChoix.ItemIndex=1) ; ;
if Cpte1.Visible then
   BEGIN
   TCpte.Caption:=MsgBox.Mess[1] ; TCpte.FocusControl:=Cpte1 ;
   Cpte.Text:='' ;
   END else
   BEGIN
   TCpte.Caption:=MsgBox.Mess[2] ; TCpte.FocusControl:=Cpte ;
   Cpte1.Text:='' ;
   END ;
end;

Procedure TFSuiTre.FaitWherePourTypePiece ;
Var Typiece : SetttTypePiece ;
BEGIN
if(LQPie='') Or (LQPie='TOU') then BEGIN CritQp:='' ; Exit ; END ;
Typiece:=WhatTypeEcr(LQPie,True,cbUnchecked) ;
CritQp:=WhereSupp('E_',Typiece) ;
END ;

Procedure TFSuiTre.FaitWherePourEtab ;
BEGIN
if LEtab='' then BEGIN CritEtab:='' ; Exit ; END ;
CritEtab:=' AND (E_ETABLISSEMENT="'+LEtab+'")' ;
END ;

Procedure TFSuiTre.FaitWherePourDevise ;
BEGIN
if LDev='' then BEGIN CritDev:='' ; Exit ; END ;
CritDev:=' AND (E_DEVISE="'+LDev+'")' ;
END ;

Procedure TFSuiTre.FaitWherePourModePaie ;
Var St : String ;
BEGIN
CritMp:='' ;
if(LModP='') Or (LModP=Traduirememoire('<<Tous>>')) then Exit ;
While LModP<>'' do
   BEGIN
   St:=ReadTokenSt(LModP) ;
   if St='' then Continue ;
   CritMp:=CritMp+'OR E_MODEPAIE="'+St+'" ' ;
   END ;
if CritMp<>'' then
   BEGIN
   CritMp:=Copy(CritMp,4,Length(CritMp)-3) ; CritMp:=' AND( '+CritMp+')' ;
   END ;
END ;

Function TFSuiTre.FabriqueLaRequeteEcr : String ;
Var St,StDatF,StDatD : String ;
BEGIN
Result:='' ;
if RgAnalyse.ItemIndex=0 then BEGIN StDatF:=UsDateTime(TabD[NbCol.Value*2]) ; StDatD:=UsDateTime(TabD[1]) ; END
                         else BEGIN StDatF:=UsDateTime(TabJ[NbCol.Value]) ; StDatD:=UsDateTime(TabJ[1]) ; END ;
Case LRgMp of
     0 : BEGIN
         St:='SELECT DISTINCT E_GENERAL FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte '+CritQp+CritMp+CritEtab+CritDev+
             ' AND E_ECRANOUVEAU="N" AND '+LaDat+'>="'+StDatD+'" '+
             ' AND '+LaDat+'<="'+StDatF+'" ORDER BY E_GENERAL';
         END ;
     1 : BEGIN
         St:='SELECT DISTINCT E_GENERAL, E_MODEPAIE FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte '+CritQp+CritMp+CritEtab+CritDev+
             ' AND E_ECRANOUVEAU="N" AND '+LaDat+'>="'+StDatD+'" '+
             ' AND '+LaDat+'<="'+StDatF+'" ORDER BY E_GENERAL,E_MODEPAIE';
         END ;
     2 : BEGIN
         St:='SELECT DISTINCT E_MODEPAIE FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte '+CritQp+CritMp+CritEtab+CritDev+
             ' AND E_ECRANOUVEAU="N" AND '+LaDat+'>="'+StDatD+'" '+
             ' AND '+LaDat+'<="'+StDatF+'" ORDER BY E_MODEPAIE';
         END ;
     End ;
Result:=St ;
END ;

Function TFSuiTre.FabriqueLaRequeteSomDeb : String ;
Var St : String ;
BEGIN
Result:='' ;
Case LRgMp of
     0 : BEGIN
         St:='SELECT E_GENERAL, SUM(E_DEBIT) As S1, '+
             'SUM(E_DEBITDEV) As S1DEV FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte AND '+LaDat+'>=:D1 AND '+LaDat+'<=:D2 '+
             CritQp+CritMp+CritEtab+CritDev+
             ' AND E_GENERAL=:Cte1 AND E_ECRANOUVEAU="N" GROUP BY E_GENERAL '+
             ' HAVING SUM(E_DEBIT)<>0 AND SUM(E_DEBITDEV)<>0';
         END ;
     1 : BEGIN
         St:='SELECT E_GENERAL, E_MODEPAIE, SUM(E_DEBIT) As S1, '+
             'SUM(E_DEBITDEV) As S1DEV FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte AND '+LaDat+'>=:D1 AND '+LaDat+'<=:D2 '+
             CritQp+CritMp+CritEtab+CritDev+
             ' AND E_GENERAL=:Cte1 AND E_MODEPAIE=:Mp AND E_ECRANOUVEAU="N" GROUP BY E_GENERAL,E_MODEPAIE '+
             ' HAVING SUM(E_DEBIT)<>0 AND SUM(E_DEBITDEV)<>0';
         END ;
     2 : BEGIN
         St:='SELECT E_MODEPAIE,E_GENERAL, SUM(E_DEBIT) As S1, '+
             'SUM(E_DEBITDEV) As S1DEV FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte AND '+LaDat+'>=:D1 AND '+LaDat+'<=:D2 '+
             CritQp+CritMp+CritEtab+CritDev+
             ' AND E_MODEPAIE=:Mp AND E_ECRANOUVEAU="N" GROUP BY E_MODEPAIE, E_GENERAL '+
             ' HAVING SUM(E_DEBIT)<>0 AND SUM(E_DEBITDEV)<>0';
         END ;
     End ;
Result:=St ;
END ;

Function TFSuiTre.FabriqueLaRequeteSomCre : String ;
Var St : String ;
BEGIN
Result:='' ;
Case LRgMp of
     0 : BEGIN
         St:='SELECT E_GENERAL, SUM(E_CREDIT) As S1, '+
             'SUM(E_CREDITDEV) As S1DEV FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte AND '+LaDat+'>=:D1 AND '+LaDat+'<=:D2 '+
             CritQp+CritMp+CritEtab+CritDev+
             ' AND E_GENERAL=:Cte1 AND E_ECRANOUVEAU="N" GROUP BY E_GENERAL '+
             ' HAVING SUM(E_CREDIT)<>0 AND SUM(E_CREDITDEV)<>0';
         END ;
     1 : BEGIN
         St:='SELECT E_GENERAL, E_MODEPAIE, SUM(E_CREDIT) As S1, '+
             'SUM(E_CREDITDEV) As S1DEV FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte AND '+LaDat+'>=:D1 AND '+LaDat+'<=:D2 '+
             CritQp+CritMp+CritEtab+CritDev+
             ' AND E_GENERAL=:Cte1 AND E_MODEPAIE=:Mp AND E_ECRANOUVEAU="N" GROUP BY E_GENERAL,E_MODEPAIE '+
             ' HAVING SUM(E_CREDIT)<>0 AND SUM(E_CREDITDEV)<>0';
         END ;
     2 : BEGIN
         St:='SELECT E_MODEPAIE,E_GENERAL, SUM(E_CREDIT) As S1, '+
             'SUM(E_CREDITDEV) As S1DEV FROM ECRITURE '+
             'WHERE E_CONTREPARTIEGEN=:Cte AND '+LaDat+'>=:D1 AND '+LaDat+'<=:D2 '+
             CritQp+CritMp+CritEtab+CritDev+
             ' AND E_MODEPAIE=:Mp AND E_ECRANOUVEAU="N" GROUP BY E_MODEPAIE, E_GENERAL '+
             ' HAVING SUM(E_CREDIT)<>0 AND SUM(E_CREDITDEV)<>0';
         END ;
     End ;
Result:=St ;
END ;

Procedure TFSuiTre.InitLeGrille ;
Var i,j : Integer ;
BEGIN
FListe.VidePile(True) ; FListe.ColCount:=NbCol.Value+1 ;
FListe.ColWidths[0]:=140 ; FListe.RowHeights[0]:=20 ;
j:=1 ;
for i:=1 to FListe.ColCount-1 do
    BEGIN
    FListe.ColWidths[i]:=155 ;
    if ParPerio then
       BEGIN
       FListe.Cells[i,0]:=MsgBox.Mess[14]+' '+DateToStr(TabD[j])+' '+MsgBox.Mess[7]+' '+DateToStr(TabD[j+1]) ;
       j:=j+2 ;
       END else
       BEGIN
       FListe.Cells[i,0]:=MsgBox.Mess[14]+' '+DateToStr(TabJ[i])+' '+MsgBox.Mess[7]+' '+DateToStr(TabJ[i]) ;
       END ;
    END ;
END ;

Procedure TFSuiTre.RunLeCalcul(Typ : String) ;
Var i,j : Integer ;
    TabDouble,TotCpte : TTotal ;
    PasleTitre : Boolean ;
    Stob : CString ;
BEGIN
QCpte.Close ; QEcr.Close ; QSom.Close ; QSom.Sql.Clear ;
if Typ='RECETTE' then QSom.Sql.Add(FabriqueLaRequeteSomCre)
                 else QSom.Sql.Add(FabriqueLaRequeteSomDeb) ;
ChangeSql(QSom) ; QSom.Prepare ; FillChar(TabDouble,SizeOf(TabDouble),0) ;
QCpte.Open ;

// Cherche tous les G_GENERAL de nature BQE ou CAI
While Not QCpte.EOF do
   BEGIN
   QEcr.Close ; PasLeTitre:=True ; FillChar(TotCpte,SizeOf(TotCpte),0) ;
   QEcr.ParamByName('Cte').AsString:=QCpte.Fields[0].AsString ;
   QEcr.Open ;

// Cherche tous les E_GENERAL pour la période choisit dont la contrepartie est
// le résultat de QCpte
   While Not QEcr.Eof do
     BEGIN
     QSom.Close ;
     QSom.ParamByName('Cte').AsString:=QCpte.Fields[0].AsString ;
     Case LRgMp of
       0: QSom.ParamByName('Cte1').AsString:=QEcr.Fields[0].AsString ;
       1: BEGIN
          QSom.ParamByName('Cte1').AsString:=QEcr.Fields[0].AsString ;
          QSom.ParamByName('Mp').AsString:=QEcr.Fields[1].AsString ;
          END ;
       2: QSom.ParamByName('Mp').AsString:=QEcr.Fields[0].AsString ;
      End ;
     if PasLetitre then  // Affiche le Compte de tréso de QEcr
        BEGIN
        PasLeTitre:=False ;
        Fliste.Cells[0,FListe.RowCount-1]:=MsgBox.Mess[RgMp.ItemIndex+3]+QCpte.Fields[0].AsString ;
        for i:=1 to FListe.ColCount-1 do FListe.Cells[i,FListe.RowCount-1]:='-' ;
        Fliste.RowCount:=Fliste.RowCount+1 ;
        END ;
     j:=1 ;
     for i:=1 to NbCol.Value do
        BEGIN
        QSom.Close ;
        if ParPerio then
           BEGIN
           QSom.ParamByName('D1').AsDateTime:=(TabD[j]) ;
           QSom.ParamByName('D2').AsDateTime:=(TabD[j+1]) ;
           END else
           BEGIN
           QSom.ParamByName('D1').AsDateTime:=(TabJ[i]) ;
           QSom.ParamByName('D2').AsDateTime:=(TabJ[i]) ;
           END ;
// Cherche les sommes E_DEBIT ou E_CREDIT pour la période
        QSom.Open ;
        if QSom.Eof then BEGIN j:=j+2 ; Continue ; END ;
        if Fliste.Cells[0,Fliste.RowCount-1]='' then
           Case LRgMp of
             0: Fliste.Cells[0,Fliste.RowCount-1]:=QSom.Fields[0].AsString ;
             1: BEGIN
                if(QSom.Fields[1].AsString='') or (IsFieldNull(QSom,'E_MODEPAIE'))then
                   Fliste.Cells[0,Fliste.RowCount-1]:=QSom.Fields[0].AsString+' : '+MsgBox.Mess[10]else
                   Fliste.Cells[0,Fliste.RowCount-1]:=QSom.Fields[0].AsString+' : '+CbMp.Items[CbMp.Values.Indexof(QSom.Fields[1].AsString)] ;
                END ;
             2: BEGIN
                Stob:=CString.Create ; Stob.Compte:=QSom.Fields[1].AsString ;
                if(QSom.Fields[0].AsString='') or (IsFieldNull(QSom,'E_MODEPAIE'))then
                   Fliste.Cells[0,Fliste.RowCount-1]:=MsgBox.Mess[10] else
                   Fliste.Cells[0,Fliste.RowCount-1]:=CbMp.Items[CbMp.Values.Indexof(QSom.Fields[0].AsString)] ;
                FListe.Objects[0,Fliste.RowCount-1]:=Stob ;
                END ;
           End ;
        if CAfdev.ItemIndex=0 then
           BEGIN
           TabDouble[i]:=TabDouble[i]+QSom.FindField('S1').AsFloat ;
           TotCpte[i]:=TotCpte[i]+QSom.FindField('S1').AsFloat ;
           FListe.Cells[i,Fliste.RowCount-1]:=StrFMontant(QSom.FindField('S1').AsFloat,0,Ladev.Decimale,'',True) ;
           END else
           BEGIN
           TabDouble[i]:=TabDouble[i]+QSom.FindField('S1DEV').AsFloat ;
           TotCpte[i]:=TotCpte[i]+QSom.FindField('S1DEV').AsFloat ;
           FListe.Cells[i,Fliste.RowCount-1]:=StrFMontant(QSom.FindField('S1DEV').AsFloat,0,Ladev.Decimale,Ladev.Symbole,True) ;
           END ;
        j:=j+2 ;
        END ;
     if Fliste.Cells[0,Fliste.RowCount-1]<>'' then Fliste.RowCount:=Fliste.RowCount+1 ;
     QEcr.Next ;
     END ;
     if CbChoix.ItemIndex=1 then
        BEGIN
        if FaitTotalCompte(TotCpte,QCpte.Fields[0].AsString,11) then FListe.RowCount:=FListe.RowCount+1 ;
        END ;
   QCpte.Next ;
   END ;
if Fliste.RowCount>2 then
   BEGIN
   if Typ='RECETTE' then FaitTotalCompte(TabDouble,'',8)
                    else FaitTotalCompte(TabDouble,'',9) ;
   END ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFSuiTre.FListeDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState) ;
var Text      : array[0..255] of Char;
    F         : TAlignment ;
BEGIN
OldBrush.assign(Fliste.Canvas.Brush) ;
OldPen.assign(Fliste.Canvas.Pen) ;
StrPCopy(Text,Fliste.Cells[Col,Row]);

if (FListe.Objects[0,Row]<>Nil) And (CString(FListe.Objects[0,Row]).Compte='$')
   And (Row<>0) then
   BEGIN
   FListe.Canvas.Font.Name:=ParamFont.Font.Name ;
   FListe.Canvas.Font.Size:=ParamFont.Font.Size ;
   FListe.Canvas.Font.Style:=ParamFont.Font.Style ;
   FListe.Canvas.Font.Color:=ParamFont.Font.Color ;
   END else
   BEGIN
   FListe.Canvas.Font.Name:='MS Sans Serif' ;
   FListe.Canvas.Font.Size:=8 ; FListe.Canvas.Font.Color:=clWindowText ;
   FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsBold,fsItalic,fsUnderLine,fsStrikeOut] ;
   END ;

if (gdFixed in State) then
  BEGIN
  Fliste.Canvas.Brush.Color:=Fliste.FixedColor ;
  Fliste.Canvas.Font.Color:=Fliste.Font.Color ;
  if (Row=0) And (Col=0) then F:=taCenter
                         else F:=taLeftJustify ;
  END else
  BEGIN
  F:=taRightJustify ;
  if(gdSelected in State) then
    BEGIN
    Fliste.Canvas.Brush.Color:=clActiveCaption ;
    Fliste.Canvas.Font.Color:=clCaptionText ;
    END else
    BEGIN
    Fliste.Canvas.Brush.Color:=clWindow ;
    if (FListe.Objects[0,Row]<>Nil) And (CString(FListe.Objects[0,Row]).Compte='$') then
         Fliste.Canvas.Font.Color:=ParamFont.Font.Color
    else Fliste.Canvas.Font.Color:=clWindowText ;
    END ;
  END ;

if(FListe.Cells[Col,Row]<>'*') And (FListe.Cells[Col,Row]<>'-')then
   BEGIN
   Case F of
        taRightJustify : ExtTextOut(Fliste.Canvas.Handle,Rect.Right-Fliste.Canvas.TextWidth(Fliste.Cells[Col,Row])-3,
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
        taCenter       : ExtTextOut(Fliste.Canvas.Handle,Rect.Left+((Rect.Right-Rect.Left-Fliste.canvas.TextWidth(Fliste.Cells[Col,Row])) div 2),
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
        else             ExtTextOut(Fliste.Canvas.Handle,Rect.Left+2,
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
   End ;
   END  ;
if(FListe.Cells[Col,Row]='*') then
   BEGIN
   Fliste.Canvas.Brush.Color:=Fliste.FixedColor ;
   Fliste.Canvas.Brush.Style:=bsSolid ;
   Fliste.Canvas.Pen.Color:=Fliste.FixedColor ;
   Fliste.Canvas.Pen.Mode:=pmCopy ;
   Fliste.Canvas.Pen.Style:=psClear ;
   Fliste.Canvas.Pen.Width:=1 ;
   Fliste.Canvas.Rectangle(Rect.Left,Rect.Top,Rect.Right+1,Rect.Bottom+1) ;
   DrawEdge(Fliste.Canvas.Handle, Rect, BDR_RAISEDOUTER, BF_BOTTOMRIGHT);
   if Not (gdFixed in State) then
      BEGIN
      Dec(Rect.Bottom); Dec(Rect.Right);
      DrawEdge(Fliste.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
      Inc(Rect.Top); Inc(Rect.Left);
      END ;
   DrawEdge(Fliste.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT or BF_MIDDLE);
   END ;

if(FListe.Cells[Col,Row]='-') then
   BEGIN
   Fliste.Canvas.Brush.Color:=Fliste.FixedColor ;
   Fliste.Canvas.Brush.Style:=bsBDiagonal ;
   Fliste.Canvas.Pen.Color:=Fliste.FixedColor ;
   Fliste.Canvas.Pen.Mode:=pmCopy ;
   Fliste.Canvas.Pen.Style:=psClear ;
   Fliste.Canvas.Pen.Width:=1 ;
   Fliste.Canvas.Rectangle(Rect.Left,Rect.Top,Rect.Right+1,Rect.Bottom+1) ;
   END ;

if((gdfixed in State) and Fliste.Ctl3D) then
   BEGIN
   DrawEdge(Fliste.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
   DrawEdge(Fliste.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
   END;
Fliste.Canvas.Brush.Assign(OldBrush) ; Fliste.Canvas.Pen.Assign(OldPen) ;
end;

procedure TFSuiTre.DevChange(Sender: TObject);
begin
if(Dev.Value=V_PGI.DevisePivot) or (Dev.Value='') then Ladev.Decimale:=V_PGI.OkDecV
else BEGIN
     Ladev.Code:=Dev.Value ; GetInfosDevise(Ladev) ;
     END ;
end;

procedure TFSuiTre.FListeDblClick(Sender: TObject);
Var ACol,ARow,i,j : Integer ;
    Cpte,Cpte1,ModP,St,StMp : String ;
    ACritEdt : ClassCritEdt;
begin
if Not BGL.Enabled then Exit ;
ACritEdt := ClassCritEdt.Create;
Fillchar(ACritEdt.CritEdt,SizeOf(ACritEdt.CritEdt),#0);
ACol:=Fliste.Col ; ARow:=Fliste.Row ;
if(FListe.Cells[ACol,ARow]='') or (FListe.Cells[ACol,ARow]='*') or
  (FListe.Cells[ACol,ARow]='-') or (FListe.Objects[0,ARow]<>Nil) then Exit ;
j:=1 ;  ModP:='' ; Cpte:='' ;
for i:=1 to ACol-1 do j:=j+2 ;
if ParPerio then
   BEGIN
   ACritEdt.CritEdt.Date1:=TabD[j] ; ACritEdt.CritEdt.Date2:=TabD[j+1] ;
   END else
   BEGIN
   ACritEdt.CritEdt.Date1:=TabJ[j] ; ACritEdt.CritEdt.Date2:=TabJ[j] ;
   END ;
ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ; ACritEdt.CritEdt.DateFin:=ACritEdt.CritEdt.Date2 ;
ACritEdt.CritEdt.NatureEtat:=neGL ;
Case LRgMp of
     0:Cpte:=Fliste.Cells[0,ARow] ;
     1:BEGIN
       Cpte:=Trim(Copy(Fliste.Cells[0,ARow],1,Pos(':',Fliste.Cells[0,ARow])-1)) ;
       St:=Copy(Fliste.Cells[0,ARow],Pos(':',Fliste.Cells[0,ARow])+2,Length(Fliste.Cells[0,ARow])) ;
       ModP:=CbMp.Values[CbMp.Items.IndexOf(St)] ;
       END ;
     2:BEGIN
       St:=Copy(Fliste.Cells[0,ARow],1,Length(Fliste.Cells[0,ARow])) ;
       if St<>MsgBox.Mess[10] then ModP:=CbMp.Values[CbMp.Items.IndexOf(St)]
                              else ModP:='' ;
       Cpte:=CString(FListe.Objects[0,ARow]).Compte ;
       END ;
 End ;
for i:=ARow DownTo 1 do BEGIN j:=i ; if FListe.Cells[ACol,i]='-' then Break ; END ;
Cpte1:=Copy(FListe.Cells[0,j],Pos(':',FListe.Cells[0,j])+1,Length(FListe.Cells[0,j])) ;
ACritEdt.CritEdt.GL.ForceNonCentralisable:=TRUE ;
ACritEdt.CritEdt.Cpt1:=Cpte ; ACritEdt.CritEdt.Cpt2:=Cpte ;
if ModP<>'' then StMp:=' AND (E_MODEPAIE="'+ModP+'") ' else StMp:=' ' ;
St:=' AND (E_GENERAL="'+Cpte+'") ' ;
ACritEdt.CritEdt.QualifPiece:=LQPie ;
ACritEdt.CritEdt.DeviseSelect:=LDev ;
ACritEdt.CritEdt.Etab:=LEtab ;
ACritEdt.CritEdt.SQLPLUS:=' AND '+LeAndEcr+St+StMp ;
TheData := ACritEdt;
CPLanceFiche_CPGLGene('');
TheData := nil;
ACritEdt.Free;
end;

procedure TFSuiTre.BGeneClick(Sender: TObject);
Var ACol,ARow : Integer ;
    Cpte : String ;
begin
ACol:=Fliste.Col ; ARow:=Fliste.Row ;
if(FListe.Cells[ACol,ARow]='*') or (FListe.Cells[ACol,ARow]='-') or
  (FListe.Objects[0,ARow]<>Nil) then Exit ;
Case LRgMp of
     0:Cpte:=Fliste.Cells[0,ARow] ;
     1:Cpte:=Trim(Copy(Fliste.Cells[0,ARow],1,Pos(':',Fliste.Cells[0,ARow])-1)) ;
   End ;
FicheGene(Nil,'',Cpte,taConsult,0);
end;

procedure TFSuiTre.BMPClick(Sender: TObject);
Var ModP,St : String ;
    ACol,ARow : Integer ;
begin
ACol:=Fliste.Col ; ARow:=Fliste.Row ;
if(FListe.Cells[ACol,ARow]='*') or (FListe.Cells[ACol,ARow]='-') or
  (FListe.Objects[0,ARow]<>Nil) or (FListe.Cells[0,ARow]=MsgBox.Mess[10])then Exit ;
Case LRgMp of
     1:BEGIN
       St:=Copy(Fliste.Cells[0,ARow],Pos(':',Fliste.Cells[0,ARow])+2,Length(Fliste.Cells[0,ARow])) ;
       if St=MsgBox.Mess[10] then Exit ;
       ModP:=CbMp.Values[CbMp.Items.IndexOf(St)] ;
       END ;
     2:BEGIN
       if Fliste.Cells[0,ARow]=MsgBox.Mess[10] then Exit ;
       ModP:=CbMp.Values[CbMp.Items.IndexOf(Fliste.Cells[0,ARow])] ;
       END ;
   End ;
FicheModePaie_AGL(ModP) ;
end;

procedure TFSuiTre.RgAnalyseClick(Sender: TObject);
begin
Case RgAnalyse.ItemIndex of
     0: BEGIN
        ParPerio:=True ; Tcol.Caption:=MsgBox.Mess[12] ; Perio.Enabled:=True ; BPerio.Enabled:=True ;
        MEDat.Enabled:=False ; NbCol.MaxValue:=12 ; NbCol.Value:=4 ;
        END ;
     1: BEGIN
        ParPerio:=False ; Tcol.Caption:=MsgBox.Mess[13] ; Perio.Enabled:=False ; BPerio.Enabled:=False ;
        MEDat.Enabled:=True ; MEDat.Text:=DateToStr(V_PGI.DateEntree) ; NbCol.MaxValue:=31 ; NbCol.Value:=4 ;
        END ;
  End ;
end;

Procedure TFSuiTre.RempliTabloJour ;
Var i,k : Byte ;
    Dat : TDateTime ;
BEGIN
FillChar(TabJ,SizeOf(TabJ),0) ;
k:=NbCol.Value ;
if CbJouvre.State=cbChecked then
   BEGIN
   for i:=k DownTo 1 do
     BEGIN
     if i=k then TabJ[i]:=StrToDate(MEDat.Text)
            else TabJ[i]:=TabJ[i+1]-1 ;
     END ;
   END else
   BEGIN
   Dat:=StrToDate(MEDat.Text) ;
   for i:=k DownTo 1 do
     BEGIN
     if (DayOfWeek(Dat)=1) Or (DayOfWeek(Dat)=7) then
        While(DayOfWeek(Dat)=1) Or (DayOfWeek(Dat)=7) do Dat:=Dat-1 ;
     TabJ[i]:=Dat ; Dat:=Dat-1 ;
     END ;
   END ;
END ;

procedure TFSuiTre.RgDatClick(Sender: TObject);
begin
Case RgDat.ItemIndex of
     0 : LaDat:='E_DATECOMPTABLE' ;
     1 : LaDat:='E_DATEVALEUR' ;
   End ;
end;

Function TFSuiTre.FaitTotalCompte(Var T: TTotal ; St : String ; Ind : Byte) : Boolean ;
Var X : CString ;
    i : Byte ;
BEGIN
Result:=False ;
for i:=1 to NbCol.Value do
    BEGIN
    if T[i]<>0 then
        BEGIN
        if CAfdev.ItemIndex=0 then
           FListe.Cells[i,Fliste.RowCount-1]:=StrFMontant(T[i],0,Ladev.Decimale,'',True)
        else
           FListe.Cells[i,Fliste.RowCount-1]:=StrFMontant(T[i],0,Ladev.Decimale,Ladev.Symbole,True) ;
        Result:=True ;
        END ;
    END ;
if Result then
   BEGIN
   Fliste.Cells[0,Fliste.RowCount-1]:=MsgBox.Mess[Ind]+St ;
   X:=CString.Create ; X.Compte:='$' ;
   Fliste.Objects[0,Fliste.RowCount-1]:=X ;
   END ;
END ;

procedure TFSuiTre.BFontClick(Sender: TObject);
begin if ParamFont.Execute then FListe.Invalidate ; end;

procedure TFSuiTre.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSuiTre.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(FListe,Nil,Caption,'') ;
end;

procedure TFSuiTre.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;

end;

procedure TFSuiTre.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFSuiTre.Baide1Click(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSuiTre.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=vk_F9 then BChercheClick(nil); //SG6 17/11/04 Gestion des filtres FQ 14976 
end;

end.
