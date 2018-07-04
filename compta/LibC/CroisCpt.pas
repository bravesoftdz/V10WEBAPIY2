{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 26/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CroisCpt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, Grids, Hctrls, StdCtrls, Buttons,
{$IFDEF EAGLCLIENT}
  UTob, UtileAGL,  
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  PrintDBG,
{$ENDIF}
  ExtCtrls, Ent1, HEnt1, Menus, ComCtrls,
  UObjFiltres, {SG6 04/01/05 Gestion Filtres V6 FQ 15145}
  HTB97, HPanel, UiUtil;

Procedure ParamCroiseCompte ;

type
  TFCroisCpt = class(TForm)
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BFerme: TToolbarButton97;
    BValider: TToolbarButton97;
    BImprimer: TToolbarButton97;
    FListe: THGrid;
    HMTrad: THSystemMenu;
    Pages: TPageControl;
    PCritere: TTabSheet;
    TypeCroisement: TRadioGroup;
    TBudJal: TLabel;
    BudJal: THValComboBox;
    CBGenere: TCheckBox;
    PFiltres: TToolWindow97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    BFiltre: TToolbarButton97;
    TCbGen: TLabel;
    CbGen: THValComboBox;
    TCbGen_: TLabel;
    TCbSec: TLabel;
    HM: THMsgBox;
    HLabel3: THLabel;
    TypVue: THValComboBox;
    CbGen_: THValComboBox;
    CbSec: THValComboBox;
    CbSec_: THValComboBox;
    TCbSec_: TLabel;
    Bevel1: TBevel;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    Label1: TLabel;
    Dock: TDock97;
    procedure BFermeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TypVueChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BudJalChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    procedure BImprimerClick(Sender: TObject);
    procedure TypeCroisementClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    //SG6 04/01/05 Gestion Filtres V6 FQ 15145
    ObjFiltre : TObjFiltre;
    WMinX,WMinY : Integer ;
    Vue : String ;
    LCpteG,LCpteS : HTStringList ;
    MemoJal,MemoVue ,GeneAxe : String ;
    Modifier,CatExiste : Boolean ;
    LFourCpteG,LFourCpteS : HTStringList ;
    CpteBudG,CpteBudS : String ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure RempliListeCompte(St : String ; Li : HTStringList ; i : Byte) ;
    Procedure InitFListe ;
    Procedure RempliFListe ;
    Procedure RempliColRow(Li : HTStringList ; LesRow : Boolean) ;
    Procedure NettoieLeGrille ;
    Procedure ChargeInfoCroiser ;
    Procedure InverseChoix ;
    Function  EstControle : boolean ;
    procedure MultiGenereCroise ( CodeGen,CodeSect : String ) ;
    procedure GenereCroisGene ;
    Function  QuelCode : String ;
    Procedure RempliFourchetteCpte(StG,StS : String) ;
  public
    { Déclarations publiques }
    SecAtt : string; {JP 17/11/05 : FQ 16317}
    GenAtt : string; {JP 17/11/05 : FQ 16317}
    function IsCpteSectAttente(Cpte, Sect : string) : Boolean;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus;

Procedure ParamCroiseCompte ;
var FCroisCpt : TFCroisCpt ;
    PP : THPanel ;
BEGIN
FCroisCpt:=TFCroisCpt.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FCroisCpt.ShowModal ;
    Finally
     FCroisCpt.Free ;
    End ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FCroisCpt,PP) ;
   FCroisCpt.Show ;
   END ;
SourisNormale ;
END ;

procedure TFCroisCpt.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCroisCpt.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFCroisCpt.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  PrintDBGrid(Caption, '', '', '', FListe);
{$ELSE}
  PrintDBGrid(FListe,Nil,Caption,'') ;
{$ENDIF}
end;

procedure TFCroisCpt.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  //SG6 04/01/2005 Gestion Filtres V6 FQ 15145
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'CROISCPT');

PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=Height ; GeneAxe:='' ; CpteBudG:='' ;
CpteBudS:='' ;
LCpteG:=HTStringList.Create ;
LCpteS:=HTStringList.Create ;
LFourCpteG:=HTStringList.Create ;
LFourCpteS:=HTStringList.Create ;
end;

procedure TFCroisCpt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG6 04/01/05 Gestion Filtres V6 FQ 15145
ObjFiltre.Free;

LCpteG.Free ; LCpteS.Free ; LFourCpteG.Free ; LFourCpteS.Free ;
if Parent is THPanel then Action:=caFree ;
end ;

procedure TFCroisCpt.BFermeClick(Sender: TObject);
begin
Close ;
//SG6 04/01/05 Vide le panel
if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFCroisCpt.FormShow(Sender: TObject);
Var i : Integer ;
begin
//SG6 04/01/05 Gestion Filtres V6 FQ 15145
ObjFiltre.Charger;

CatExiste:=FALSE ; For i:=1 To MaxCatBud Do If VH^.CatBud[i].Code<>'' Then CatExiste:=TRUE ;
MemoJal:='' ; MemoVue:='' ;  Modifier:=False ;
if Not CatExiste then TypeCroisement.Enabled:=False ;

// RR 01/09/2005 FQ 16518
// cette notion n'a plus lieu d'être.
// maintenant c'est la notion de buissness, donc le packavancé est en standard
(*if ((EstSerie(S3)) or (EstSerie(S5))) then
  BEGIN
  TypeCroisement.Visible:=False ;
  Caption:=HM.Mess[4] ; UpdateCaption(Self) ;
  END ;*)
CBGenere.Checked:=False ; CBGenere.Enabled:=False ;
If TypVue.Values.Count>0 Then TypVue.Value:=TypVue.Values[0] ;
If BudJal.Values.Count>0 Then BudJal.Value:=BudJal.Values[0] ;
end;

procedure TFCroisCpt.TypVueChange(Sender: TObject);
begin
Vue:=TypVue.Value ;
if Vue=MemoVue then Exit ;
MemoVue:=Vue ; BChercheClick(Nil) ;
end;

function TFCroisCpt.EstControle : boolean ;
BEGIN
Result:=False ; if BUDJAL.Value='' then Exit ;
if VH^.JalCtrlBud=BUDJAL.Value then Result:=True ;
END ;

procedure TFCroisCpt.BudJalChange(Sender: TObject);
Var QLoc : TQuery ;
    StG,StS,GBud,GSect : String ;
begin
if MemoJal=BudJal.Value then Exit ;
MemoJal:=BudJal.Value ;
StG:='' ; StS:='' ; GeneAxe:='' ; GBud:='' ; GSect:='' ;
//QLoc:=OpenSql('Select BJ_BUDGENES,BJ_BUDGENES2,BJ_BUDSECTS,BJ_BUDSECTS2,BJ_GENEATTENTE, BJ_SECTATTENTE, BJ_AXE from BUDJAL Where BJ_BUDJAL="'+BudJal.Value+'"',True) ;
//QLoc:=OpenSql('Select BJ_BUDGENES,BJ_BUDGENES2,BJ_BUDSECTS,BJ_BUDSECTS2,BJ_GENEATTENTE, BJ_SECTATTENTE, BJ_AXE from BUDJAL Where '+QuelCode+'="'+BudJal.Value+'"',True) ;
QLoc:=OpenSql('Select * from BUDJAL Where '+QuelCode+'="'+BudJal.Value+'" ORDER BY BJ_BUDJAL',True) ;
if Not QLoc.Eof then
  BEGIN
  if QLoc.FindField('BJ_BUDGENES2').AsString<>'' then StG:=Trim(QLoc.FindField('BJ_BUDGENES').AsString)+Trim(QLoc.FindField('BJ_BUDGENES2').AsString)
                                                 else StG:=Trim(QLoc.FindField('BJ_BUDGENES').AsString) ;
  if QLoc.FindField('BJ_BUDSECTS2').AsString<>'' then StS:=Trim(QLoc.FindField('BJ_BUDSECTS').AsString)+Trim(QLoc.FindField('BJ_BUDSECTS2').AsString)
                                                 else StS:=Trim(QLoc.FindField('BJ_BUDSECTS').AsString) ;
  GBud:=QLoc.FindField('BJ_GENEATTENTE').AsString ;
  GSect:=QLoc.FindField('BJ_SECTATTENTE').AsString ;
  GeneAxe:=QLoc.FindField('BJ_AXE').AsString ;
  StG:=StG+GBud+';' ; StS:=StS+GSect+';' ;
  END ;
Ferme(QLoc) ;
RempliFourchetteCpte(StG,StS) ;
CpteBudG:=StG ; CpteBudS:=StS ;

  {JP 17/11/05 : FQ 16317 : Mémorisation du compte et de la section d'attente}
  GenAtt := GBud;
  SecAtt := GSect;

if EstControle then BEGIN CBGenere.Enabled:=True  ; CBGenere.Checked:=True ; END
               else BEGIN CBGenere.Enabled:=False ; CBGenere.Checked:=False ; END ;
end;

Procedure TFCroisCpt.RempliListeCompte(St : String ; Li : HTStringList ; i : Byte) ;
Var StTemp : String ;
    j : Integer ;
BEGIN
Li.Clear ;
if ((i=0) And (CbGen.Value='')) Or ((i=1) And (CbSec.Value='')) then
   BEGIN
   While St<>'' do BEGIN StTemp:=ReadTokenSt(St) ; if StTemp<>'' then Li.Add(StTemp) ; END ;
   END else
   BEGIN
   Case i of
        0 : BEGIN
            if CbGen.ItemIndex>CbGen_.ItemIndex then BEGIN HM.Execute(3,'','') ; CbGen.SetFocus ; Exit ; END ;
            for j:=CbGen.ItemIndex to CbGen.Items.Count-1 do
                BEGIN
                Li.Add(CbGen.Values[j]) ;
                if j+1>CbGen_.ItemIndex then Break ;
                END ;
            END ;
        1 : BEGIN
            if CbSec.ItemIndex>CbSec_.ItemIndex then BEGIN HM.Execute(3,'','') ; CbSec.SetFocus ; Exit ; END ;
            for j:=CbSec.ItemIndex to CbSec.Items.Count-1 do
                BEGIN
                Li.Add(CbSec.Values[j]) ;
                if j+1>CbSec_.ItemIndex then Break ;
                END ;
            END ;
      End ;
   END ;
END ;

procedure TFCroisCpt.BChercheClick(Sender: TObject);
begin
RempliListeCompte(CpteBudG,LCpteG,0) ; RempliListeCompte(CpteBudS,LCpteS,1) ;
if (LCpteG.Count>0) And (LCpteS.Count>0) then RempliFListe ;
end;

Procedure TFCroisCpt.NettoieLeGrille ;
Var i : Integer ;
BEGIN
FListe.VidePile(False) ;
for i:=0 to FListe.ColCount-1 do FListe.Cells[i,0]:='' ;
FListe.ColCount:=2 ;
END ;

Procedure TFCroisCpt.RempliFListe ;
BEGIN
NettoieLeGrille ;
FListe.Cells[0,0]:=TypVue.Text ;
if Vue='GS' then
   BEGIN
   if LCpteG.Count<=0 then FListe.RowCount:=2 else FListe.RowCount:=LCpteG.Count+1 ;
   if LCpteS.Count<=0 then FListe.ColCount:=2 else FListe.ColCount:=LCpteS.Count+1 ;
   InitFListe ; RempliColRow(LCpteG,True) ; RempliColRow(LCpteS,False) ;
   END else
   BEGIN
   if LCpteS.Count<=0 then FListe.RowCount:=2 else FListe.RowCount:=LCpteS.Count+1 ;
   if LCpteG.Count<=0 then FListe.ColCount:=2 else FListe.ColCount:=LCpteG.Count+1 ;
   InitFListe ; RempliColRow(LCpteS,True) ; RempliColRow(LCpteG,False) ;
   END ;
ChargeInfoCroiser ;
END ;

Procedure TFCroisCpt.InitFListe ;
Var i,j : Integer ;
BEGIN
for i:=1 to FListe.ColCount-1 do
    BEGIN
    FListe.ColAligns[i]:=taCenter ;
    for j:=1 to FListe.RowCount-1 do FListe.Cells[i,j]:='-' ;
    END ;
END ;

Procedure TFCroisCpt.RempliColRow(Li : HTStringList ; LesRow : Boolean) ;
Var i : Integer ;
BEGIN
if Li.Count<=0 then Exit ;
if LesRow then for i:=1 to FListe.RowCount-1 do FListe.Cells[0,i]:=Li.Strings[i-1]
          else for i:=1 to FListe.ColCount-1 do FListe.Cells[i,0]:=Li.Strings[i-1] ;
END ;

Procedure TFCroisCpt.ChargeInfoCroiser ;
Var QLoc : TQuery ;
    ic,ir : integer ;
    ChampCol,ChampRow : String ;
BEGIN
if Vue='GS' then BEGIN ChampCol:='CX_SECTION' ; ChampRow:='CX_COMPTE' ; END
            else BEGIN ChampCol:='CX_COMPTE' ; ChampRow:='CX_SECTION' ; END ;
QLoc:=OpenSql('Select '+ChampCol+', '+ChampRow+',CX_INFO From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+BudJal.Value+'"',True) ;
While Not QLoc.Eof do
   BEGIN
   ic:=FListe.Rows[0].IndexOf(QLoc.Fields[0].AsString) ;
   ir:=FListe.Cols[0].IndexOf(QLoc.Fields[1].AsString) ;
   if ((ir>=0) and (ic>=0)) then FListe.Cells[ic,ir]:=QLoc.Fields[2].AsString ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

procedure TFCroisCpt.MultiGenereCroise ( CodeGen,CodeSect : String ) ;
Var QBG,QBS,QG,QS,QCH : TQuery ;
    BGCompteRub,BGExcluRub,BSCompteRub,BSExcluRub : String ;
    WG,WS,SQLG,SQLS,CG,CS : String ;
    fb              : TFichierBase ;
    TG,TS           : HTStrings ;
    i,j             : integer ;
    Find            : boolean ;
    k : Integer ;
BEGIN
{Retrouver les cptes gene du compte budget}
QBG:=OpenSQL('SELECT BG_COMPTERUB, BG_EXCLURUB from BUDGENE Where BG_BUDGENE="'+CodeGen+'"',True) ;
if QBG.EOF then BEGIN Ferme(QBG) ; Exit ; END else
   BEGIN
   BGCompteRub:=QBG.Fields[0].AsString ; BGExcluRub:=QBG.Fields[1].AsString ;
   Ferme(QBG) ;
   END ;
{Retrouver les sections gene de la section budget}
QBS:=OpenSQL('SELECT BS_SECTIONRUB, BS_EXCLURUB from BUDSECT Where BS_BUDSECT="'+CodeSect+'"',True) ;
if QBS.EOF then BEGIN Ferme(QBS) ; Exit ; END else
   BEGIN
   BSCompteRub:=QBS.Fields[0].AsString ; BSExcluRub:=QBS.Fields[1].AsString ;
   Ferme(QBS) ;
   END ;
{Lire les généraux}
WG:=AnalyseCompte(BGCompteRub,fbGene,False,False) ;
if WG<>'' then SQLG:='Select G_GENERAL from GENERAUX Where '+WG else Exit ;
WG:=AnalyseCompte(BGExcluRUB,fbGene,True,False) ;
if WG<>'' then SQLG:=SQLG+' AND '+WG ;
{Lire les sections}
fb:=AxeToFb(GeneAxe) ; WS:=AnalyseCompte(BSCompteRub,fb,False,False) ;
if WS<>'' then SQLS:='Select S_SECTION from SECTION Where '+WS else Exit ;
WS:=AnalyseCompte(BSExcluRUB,fb,True,False) ;
if WS<>'' then SQLS:=SQLS+' AND '+WS ;
{Charger les généraux}
TG:=HTStringList.Create ;
QG:=OpenSQL(SQLG,True) ;
While Not QG.EOF do
   BEGIN
   CG:=QG.Fields[0].AsString ;
   if TG.IndexOf(CG)<0 then TG.Add(CG) ;
   QG.Next ;
   END ;
Ferme(QG) ;
{Charger les sections}
TS:=HTStringList.Create ;
QS:=OpenSQL(SQLS,True) ;
While Not QS.EOF do
   BEGIN
   CS:=QS.Fields[0].AsString ;
   if TS.IndexOf(CS)<0 then TS.Add(CS) ;
   QS.Next ;
   END ;
Ferme(QS) ;
CS:=VH^.Cpta[fb].Attente ; if ((CS<>'') and (TS.IndexOf(CS)<0)) then TS.Add(CS) ;
{Création des enregistrements}
k:=0 ;
for i:=0 to TG.Count-1 do
    BEGIN
    for j:=0 to TS.Count-1 do
       BEGIN
       if (i*j)-k>1000 then BEGIN k:=i*j ; CommitTrans ; BeginTrans ; END ;
       CG:=TG[i] ; CS:=TS[j] ;
       QCH:=OpenSQL('Select * from CROISCPT Where CX_TYPE="GEN" AND CX_JAL="'+BudJal.Value+'" AND CX_COMPTE="'+CG+'" AND CX_SECTION="'+CS+'"',True) ;
       Find:=Not QCH.EOF ;
       Ferme(QCH) ;
       if Not Find then
          BEGIN
          ExecuteSql('Insert into CROISCPT (CX_TYPE,CX_JAL,CX_COMPTE,CX_SECTION,CX_INFO)'+
                     'Values ("GEN","'+BudJal.Value+'","'+CG+'","'+CS+'","X")') ;
          END ;
       END ;
    END ;
TG.Free ; TS.Free ;
END ;

procedure TFCroisCpt.GenereCroisGene ;
Var i,j : integer ;
    CodeGen,CodeSect : String ;
BEGIN
InitMove(FListe.ColCount-1,'') ;
ExecuteSql('Delete From CROISCPT Where CX_TYPE="GEN" And CX_JAL="'+BudJal.Value+'"') ;
for i:=1 to FListe.ColCount-1 do
    BEGIN
    MoveCur(False) ;
    for j:=1 to FListe.RowCount-1 do
        BEGIN
        if FListe.Cells[i,j]='X' then
           BEGIN
           if Vue='GS' then BEGIN CodeGen:=FListe.Cols[0].Strings[j] ; CodeSect:=FListe.Rows[0].Strings[i] ; END
                       else BEGIN CodeGen:=FListe.Rows[0].Strings[i] ; CodeSect:=FListe.Cols[0].Strings[j] ; END ;
           MultiGenereCroise(CodeGen,CodeSect) ;
           END ;
        END ;
    END ;
FiniMove ;
END ;

{---------------------------------------------------------------------------------------}
procedure TFCroisCpt.BValiderClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  i, j  : Integer;
  CodeGen,
  CodeSect,
  StCat : string;
  QLoc  : TQuery;
  SQL   : string;
begin
  if (FListe.Cells[1,0]='') or (FListe.Cells[0,1]='') then Exit ;
  if not Modifier then Exit ;
  if HM.Execute(0,'','')<>mrYes then Exit ;

  BeginTrans ;
  {On commence par vider la table des précédents croisements}
  if (CbGen.Value = '') and (CbSec.Value = '') then
    ExecuteSql('DELETE FROM CROISCPT WHERE CX_TYPE = "BUD" AND CX_JAL = "' + BudJal.Value + '"')

  else begin
    {S'il y a un filtre sur les sections ou les comptes, on ne supprime que les croisements concernés}
    Sql := 'DELETE FROM CROISCPT WHERE CX_TYPE = "BUD" AND CX_JAL = "' + BudJal.Value + '" AND CX_COMPTE="';
    for i := 1 to FListe.ColCount - 1 do begin
      for j := 1 to FListe.RowCount - 1 do begin
        {Selon que l'on est en Général par section ou Section par général, les sections et les généraux sont
         soit en colonne soit en ligne}
        if Vue = 'GS' then begin CodeGen := FListe.Cells[0, j]; CodeSect := FListe.Cells[i, 0]; end
                      else begin CodeGen := FListe.Cells[i, 0]; CodeSect := FListe.Cells[0, j]; end;
        ExecuteSQL(SQL + CodeGen + '" AND CX_SECTION = "' + CodeSect + '"');
      end;
    end;
  end;

  {Récupération de la catégorie budgétaire ou du Budget, en fonction de l'option}
  StCat := '';
  QLoc := OpenSql('SELECT BJ_CATEGORIE FROM BUDJAL WHERE ' + QuelCode + ' = "' + BudJal.Value + '" ORDER BY BJ_BUDJAL', True);
  if not QLoc.Eof then StCat := QLoc.FindField('BJ_CATEGORIE').AsString;
  Ferme(QLoc);

  InitMove((FListe.ColCount - 1) * (FListe.RowCount - 1), '');
  for i:=1 to FListe.ColCount - 1 do begin
    for j:=1 to FListe.RowCount - 1 do begin
      MoveCur(False) ;
      {JP 18/11/05 : Je ne vois pas trop l'intérêt de transactions par paquets, surtout qu'au premier
                     CommitTrans, la suppression des précédants croisements sera validée !!
      if (i * j) - k > 1000 then begin k := i * j; CommitTrans; BeginTrans; end;}

      {Selon que l'on est en Général par section ou Section par général, les sections et les généraux sont
       soit en colonne soit en ligne}
      if Vue = 'GS' then begin CodeGen := FListe.Cells[0, j]; CodeSect := FListe.Cells[i, 0]; end
                    else begin CodeGen := FListe.Cells[i, 0]; CodeSect := FListe.Cells[0, j]; end;
      {if Vue = 'GS' then begin CodeGen := FListe.Cols[0].Strings[j]; CodeSect := FListe.Rows[0].Strings[i]; end
                    else begin CodeGen := FListe.Rows[0].Strings[i]; CodeSect := FListe.Cols[0].Strings[j]; end;}

      if (FListe.Cells[i, j] = 'X') or
         {JP 17/11/05 : FQ 16317 : on ne génère plus automatiquement les enregistrements pour la dernière
                        colonne et la dernière ligne, mais si on est sur la section ou le compte d'attente.
                        (i = FListe.ColCount - 1) or (j = FListe.RowCount - 1)}
          IsCpteSectAttente(CodeGen, CodeSect) then begin
        {Création de l'enregistrement : seuls figurent dans la tables les croisements}
        ExecuteSql('INSERT INTO CROISCPT (CX_TYPE, CX_JAL, CX_COMPTE, CX_SECTION, CX_INFO, CX_CATEGORIE) '+
                   'VALUES ("BUD", "' + BudJal.Value + '", "' + CodeGen + '","' + CodeSect + '", "X", "' + StCat + '")');
      end;
    end;
  end;

  FiniMove ;
  {Gestion des croisements sur la comptabilité général}
  if CBGenere.Checked then GenereCroisGene ;
  CommitTrans ;
  Modifier := False ;
end;

procedure TFCroisCpt.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin if (Key=VK_SPACE) And (Shift=[]) then InverseChoix ; end;

procedure TFCroisCpt.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin if (ssCtrl in Shift) And (Button=mbLeft) then InverseChoix ; end;

Procedure TFCroisCpt.InverseChoix ;
Var i,j : Integer ;
    GRect : TGridRect ;
BEGIN
GRect:=FListe.Selection ; Modifier:=True ;
for i:=GRect.Left to GRect.Right do
 for j:=GRect.Top to GRect.Bottom do
     if FListe.Cells[i,j]='-' then FListe.Cells[i,j]:='X' else FListe.Cells[i,j]:='-' ;
END ;

Function TFCroisCpt.QuelCode : String ;
BEGIN
If TypeCroisement.ItemIndex=0 Then Result:=' BJ_BUDJAL' Else Result:=' BJ_CATEGORIE' ;
END ;

procedure TFCroisCpt.TypeCroisementClick(Sender: TObject);
begin
If (Not CatExiste) And (TypeCroisement.ItemIndex=1) Then BEGIN TypeCroisement.ItemIndex:=0 ; Exit ; END ;
If TypeCroisement.ItemIndex=0 Then
   BEGIN
   TBudJal.Caption:=HM.Mess[1] ;
   BudJal.Datatype:='ttBudJalSansCat' ; BudJal.ItemIndex:=0 ; BudJal.Reload ;
   END Else
   BEGIN
   TBudJal.Caption:=HM.Mess[2] ;
   BudJal.Datatype:='ttCatJalBud' ; BudJal.ItemIndex:=0 ; BudJal.Reload ;
   END ;
if EstControle then BEGIN CBGenere.Enabled:=True  ; CBGenere.Checked:=True ; END
               else BEGIN CBGenere.Enabled:=False ; CBGenere.Checked:=False ; END ;
NettoieLeGrille ; InitFliste ;
end;

Procedure TFCroisCpt.RempliFourchetteCpte(StG,StS : String) ;
Var StGen,StSec : String ;
    QLoc : TQuery ;
BEGIN
LFourCpteG.Clear ; LFourCpteS.Clear ;
CbGen.Values.Clear ; CbGen.Items.Clear ;
CbGen_.Values.Clear ; CbGen_.Items.Clear ;
CbSec.Values.Clear ; CbSec.Items.Clear ;
CbSec_.Values.Clear ; CbSec_.Items.Clear ;
CbGen.Values.Add('') ;
CbGen.Items.Add(TraduireMemoire('<<Tous>>')) ;
CbSec.Values.Add('') ;
CbSec.Items.Add(TraduireMemoire('<<Tous>>')) ;
While StG<>'' do
   BEGIN
   StGen:=ReadTokenSt(StG) ;
   if StGen<>'' then
      BEGIN
      QLoc := OpenSQL('SELECT BG_BUDGENE,BG_LIBELLE FROM BUDGENE WHERE BG_BUDGENE="'+StGen+'"', True);
      if Not QLoc.Eof then
         BEGIN
         CbGen.Values.Add(QLoc.Fields[0].AsString) ;
         CbGen.Items.Add(QLoc.Fields[1].AsString) ;
         END ;
      END ;
      Ferme(QLoc);
   END ;
CbGen_.Values.Assign(CbGen.Values) ; CbGen_.Items.Assign(CbGen.Items) ;

While StS<>'' do
   BEGIN
   StSec:=ReadTokenSt(StS) ;
   if StSec<>'' then
      BEGIN
      QLoc := OpenSQL('SELECT BS_BUDSECT,BS_LIBELLE FROM BUDSECT WHERE BS_BUDSECT="'+StSec+'" AND BS_AXE="'+GeneAxe+'"', True);
      if Not QLoc.Eof then
         BEGIN
         CbSec.Values.Add(QLoc.Fields[0].AsString) ;
         CbSec.Items.Add(QLoc.Fields[1].AsString) ;
         END ;
      END ;
      Ferme(QLoc) ;
   END ;
CbSec_.Values.Assign(CbSec.Values) ; CbSec_.Items.Assign(CbSec.Items) ;

CbGen.ItemIndex:=0 ; CbGen_.ItemIndex:=0 ;
CbSec.ItemIndex:=0 ; CbSec_.ItemIndex:=0 ;
END ;

{---------------------------------------------------------------------------------------}
procedure TFCroisCpt.FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  //SG6 04/01/05 Gestion des Filtres V6 FQ 15145
  if Key = VK_F9 then BChercheClick(nil);
end;

{JP 17/11/05 : FQ 16317 : régarde si le compte ou la section est d'attente
{---------------------------------------------------------------------------------------}
function TFCroisCpt.IsCpteSectAttente(Cpte, Sect : string): Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (Cpte = GenAtt) or (Sect = SecAtt);
end;

end.
