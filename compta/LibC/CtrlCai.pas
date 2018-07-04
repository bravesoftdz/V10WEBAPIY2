{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/08/2004
Modifié le ... :   /  /    
Description .. : - CA - 12/08/2004 - FQ 1368 : Visualisation du graphe en 
Suite ........ : fonction du tri dans la liste. Attention : désormais la 
Suite ........ : visualisation du graphe ne pourra se faire que si la liste est 
Suite ........ : triée par date.
Mots clefs ... : 
*****************************************************************}
unit CtrlCai ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1, HCompte,
  ExtCtrls, StdCtrls, Mask, Hctrls, ComCtrls, DB, DBTables, Teengine,
  Series, Buttons, HEnt1, HQry, Grids, DBGrids, HDebug, UObjFiltres {SG6 12/01/05 Gestion Filtres V6},
  hmsgbox, SaisUtil, Menus, HDB, HSysMenu, ParamDat, HTB97, HPanel, UiUtil ;

Procedure ControlCaisse(fb : TfichierBase ; ZoomT : TZoomTable ; Cpt : String ) ;

type
  TFCtrlCai = class(TForm)
    Q: THQuery;
    Pages: TPageControl;
    Standards: TTabSheet;
    TFGen: THLabel;
    Label7: TLabel;
    HLabel4: THLabel;
    HLabel6: THLabel;
    Cpt: THCpteEdit;
    FDateCompta2: TMaskEdit;
    FDateCompta1: TMaskEdit;
    FExercice: THValComboBox;
    HPB: TToolWindow97;
    FListe: THGrid;
    QDAT: TDateTimeField;
    QDEB: TFloatField;
    QCRE: TFloatField;
    QSOL: TFloatField;
    MessErreur: THMsgBox;
    BGraph: TToolbarButton97;
    Bevel2: TBevel;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    BChercher: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BReduire: TToolbarButton97;
    CBDebDat: TComboBox;
    HMTrad: THSystemMenu;
    BMenuZoom: TToolbarButton97;
    BGL: TToolbarButton97;
    BGene: TToolbarButton97;
    PopZ: TPopupMenu;
    FSigne: THValComboBox;
    Label1: TLabel;
    FEtab: THValComboBox;
    HLabel7: THLabel;
    FDevises: THValComboBox;
    Label2: TLabel;
    FValide: TCheckBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    Dock: TDock97;
    Formateur: THNumEdit;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock971: TDock97;
    procedure FExerciceChange(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BGraphClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CptExit(Sender: TObject);
    procedure FDevisesChange(Sender: TObject);
    procedure BGeneClick(Sender: TObject);
    procedure BGLClick(Sender: TObject);
    procedure FDateCompta1KeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure OnSortGrid ( Sender : TObject);
  private
    //SG6 12/01/05 Gestion Filtres V6
    ObjFiltre : TObjFiltre;
    fb             : TFichierBase ;
    ZoomTable      : TZoomTable ;
    CptEntree,
    Dev,Symbole    : String3 ;
    Decimale       : Integer ;
    SoldeDebiteur  : Boolean ;
    Quedal         : Boolean ;
    RDev : RDevise ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure PrepareLaSQL ;
    Function  ExecuteLaSQL : Boolean ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses PrintDBG, Graph, CritEdt, UtilEdt,
     CPGeneraux_TOM,
     QRGLGen ;

Procedure ControlCaisse(fb : TfichierBase ; ZoomT : TZoomTable ; Cpt : String ) ;
Var G  : TFCtrlCai ;
    PP : THPanel ;
begin
G:=TFCtrlCai.Create(Application) ;
G.fb:=fb ; G.CptEntree:=Cpt ; G.ZoomTable:=ZoomT ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     G.ShowModal ;
    Finally
     G.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(G,PP) ;
   G.Show ;
   END ;
end ;

procedure TFCtrlCai.FExerciceChange(Sender: TObject);
begin
if FExercice.ItemIndex<>0 then ExoToDates(FExercice.Value,FDateCompta1,FDateCompta2) ;
end;

Function StAxe(fb : TFichierBase) : String3 ;
BEGIN StAxe:='A'+IntToStr(ord(fb)+1) ; END ;

Procedure TFCtrlCai.PrepareLaSQL ;
Var FiltreEtab,FiltreExo : Boolean ;
    SUM_Montant,StCalc,P : String ;
BEGIN
P:='E_' ; if fb in [fbAxe1..fbAxe5] then P:='Y_' ;
Q.Close ;
Quedal:=False ;
FiltreEtab:=FEtab.ItemIndex>0 ;
FiltreExo:=FExercice.ItemIndex>0 ;
SoldeDebiteur:=(FSigne.Value='POS') ;
Q.SQL.Clear ;
StCalc:='SELECT '+P+'DATECOMPTABLE AS DAT, ' ; Q.SQL.Add(StCalc) ;
if (FDevises.Value=V_PGI.DevisePivot) or (FDevises.Value='')
   then SUM_Montant:='SUM('+P+'DEBIT) AS DEB, SUM('+P+'CREDIT) AS CRE, SUM('+P+'DEBIT-'+P+'CREDIT) AS SOL '
   else SUM_Montant:='SUM('+P+'DEBITDEV) AS DEB, SUM('+P+'CREDITDEV) AS CRE, SUM('+P+'DEBITDEV-'+P+'CREDITDEV) AS SOL ' ;
Case fb Of
  fbGene, fbAux  : StCalc:=SUM_MONTANT+'FROM ECRITURE ' ;
  fbAxe1..fbAxe5 : StCalc:=SUM_MONTANT+'FROM ANALYTIQ ' ;
 end ;
Q.SQL.Add(StCalc) ;
Case Fb Of
  FbGene         : StCalc:='WHERE E_GENERAL="'  +Cpt.text+'" ' ;
  FbAux          : StCalc:='WHERE E_AUXILIAIRE='+Cpt.text+'" ' ;
  fbAxe1..fbAxe5 : StCalc:='WHERE Y_SECTION='+Cpt.text+'" AND Y_AXE="'+StAxe(fb)+'" ' ;
 end ;
Q.SQL.Add(StCalc) ;
StCalc:='AND '+P+'DATECOMPTABLE>="'+UsDate(FDateCompta1)+'" AND '+P+'DATECOMPTABLE<="'+UsDate(FDateCompta2)+'" ' ;
Q.SQL.Add(StCalc) ;
Case FValide.State of
  cbUnchecked : BEGIN StCalc:='AND '+P+'VALIDE="-" ' ; Q.SQL.Add(StCalc) ; END ;
  cbChecked   : BEGIN StCalc:='AND '+P+'VALIDE="X" ' ; Q.SQL.Add(StCalc) ; END ;
 end ;
if FiltreEtab then
   BEGIN StCalc:='AND '+P+'ETABLISSEMENT="'+FEtab.Value+'" ' ; Q.SQL.Add(StCalc) ; END ;
if FiltreExo then
   BEGIN StCalc:='AND '+P+'EXERCICE="'+FExercice.Value+'" ' ; Q.SQL.Add(StCalc) ; END ;
{StCalc:=TraduitNatureEcr(FNatureEcr.Value, P+'QUALIFPIECE', true, cbUnchecked(*AvecRevision.State*)) ;
if StCalc<>'' then Q.SQL.Add(StCalc) ;
}
StCalc:='GROUP BY '+P+'DATECOMPTABLE ORDER BY E_DATECOMPTABLE ' ; Q.SQL.Add(StCalc) ;
ChangeSql(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
END ;

Function TFCtrlCai.ExecuteLaSQL : Boolean ;
var Dat         : TdateTime ;
    Deb,Cre,
    Sol,Cum     : Double ;
BEGIN
CBDebDat.Items.Clear ;
CBDebDat.Items.Add('');
Result:=Q.EOF ; Cum:=0 ;
While not Q.EOF Do
  BEGIN
  Dat:=Q.FindField('DAT').AsDateTime ;
  Deb:=Q.FindField('DEB').AsFloat ;
  Cre:=Q.FindField('CRE').AsFloat ;
  Sol:=Q.FindField('SOL').AsFloat ;
  CBDebDat.Items.Add(DatetoStr(Dat)) ;
  FListe.Cells[0,FListe.RowCount-1]:=DateToStr(Dat) ;
  if Deb<>0 then FListe.Cells[1,FListe.RowCount-1]:=StrFMontant(Deb,0,Decimale,Symbole,True) ;
  if Cre<>0 then FListe.Cells[2,FListe.RowCount-1]:=StrFMontant(Cre,0,Decimale,Symbole,True) ;
  if Sol<>0 then
     BEGIN
     if SoldeDebiteur then FListe.Cells[3,FListe.RowCount-1]:=StrFMontant(Sol,0,Decimale,Symbole,True)
                      else FListe.Cells[3,FListe.RowCount-1]:=StrFMontant(Sol*(-1),0,Decimale,Symbole,True) ;
     END ;
  Cum:=Cum+Sol ;
  if Cum<0 then FListe.Cells[4,FListe.RowCount-1]:=PrintSolde(0,-Cum,Decimale,Symbole,False)
           else FListe.Cells[4,FListe.RowCount-1]:=PrintSolde(Cum,0,Decimale, Symbole,False) ;
  FListe.RowCount:=FListe.RowCount+1 ; Q.Next ;
  END ;
if Not Result then  FListe.RowCount:=FListe.RowCount-1 ;
Q.Close ;
END ;

procedure TFCtrlCai.BChercherClick(Sender: TObject);
begin
if Cpt.Text='' then BEGIN MessErreur.Execute(1,'','') ; Exit ; END ;
if Cpt.ExisteH<=0 then BEGIN MessErreur.Execute(2,'','') ; Exit ; END ;
FListe.VidePile(False) ;
PrepareLaSQL ;
if ExecuteLaSQL then
   BEGIN
   Quedal:=True ; //MessErreur.Execute(0,'','');
   FListe.SortEnabled:=False ;
   END else
   BEGIN
   FListe.SortEnabled:=True ; FListe.SortGrid(0,True) ;
   END ;
end;

procedure TFCtrlCai.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Case fb Of
  fbGene : BEGIN
           Cpt.ZoomTable:=tzGeneral ;
           if ZoomTable<>tzImmo then Cpt.ZoomTable:=ZoomTable ;
           END ;
  fbAux  : BEGIN
           Cpt.ZoomTable:=tzTiers ;
           if ZoomTable<>tzImmo then Cpt.ZoomTable:=ZoomTable ;
           END ;
  fbAxe1 : Cpt.ZoomTable:=tzSection ;
  fbAxe2 : Cpt.ZoomTable:=tzSection2 ;
  fbAxe3 : Cpt.ZoomTable:=tzSection3 ;
  fbAxe4 : Cpt.ZoomTable:=tzSection4 ;
  fbAxe5 : Cpt.ZoomTable:=tzSection5 ;
 end ;
Cpt.Text:=CptEntree ;
FExercice.Value:=EXRF(VH^.Entree.Code) ;
FSigne.Value:='POS' ; FEtab.ItemIndex:=0 ; FDevises.ItemIndex:=0 ;
//FNatureEcr.ItemIndex:=0 ; Pages.ActivePage:=Standards ;
if ZoomTable=tzGCaisse then
   BEGIN
   FDevises.Visible:=False ; Label2.Visible:=False ;
   END ;
FDevisesChange(Nil) ; ChangeMask(Formateur,Decimale,Symbole);
Quedal:=True ; FListe.SortEnabled:=False ; Formateur.Visible:=False ; Pages.ActivePage:=Pages.Pages[0] ;
//SG6 12/01/05 Gestion Filtre V6
ObjFiltre.Charger;
//SG6 112/01/05 Gestion restriction utilisateur etablissement FQ 15233
PositionneEtabUser(FEtab);
if Cpt.Text<>'' then BChercherClick(Nil) ;
end;

procedure TFCtrlCai.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFCtrlCai.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFCtrlCai.BGraphClick(Sender: TObject);
var LDate : TStrings;
    i : integer;
begin
  if Quedal then Exit
          else
  begin
    if FListe.SortDesc then  // on trie CBDebDat par ordre descendant
    begin
      LDate := TStringList.Create;
      LDate.Add('');
      for i:=(CBDebDat.Items.Count - 1) downto 1  do
        LDate.Add(CBDebDat.Items[i]);
      VisuGraph('CTRLCAI'+IntToStr(Integer(fb)),Caption,FListe,LDate,0,1,FListe.RowCount-1,TRUE,[4],NIL,Nil) ;
      LDate.Free;
    end
    else
      VisuGraph('CTRLCAI'+IntToStr(Integer(fb)),Caption,FListe,CBDebDat.Items,0,1,FListe.RowCount-1,TRUE,[4],NIL,Nil) ;
  end;
end;

procedure TFCtrlCai.WMGetMinMaxInfo(var MSG: Tmessage);
begin with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; end;

procedure TFCtrlCai.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
//SG6 12/01/05 Gestion des Filtres V6
Composants.PopupF   := POPF;
Composants.Filtres  := FFILTRES;
Composants.Filtre   := BFILTRE;
Composants.PageCtrl := Pages;
ObjFiltre := TObjFiltre.Create(Composants,'CTRLCAI');

WMinX:=Width ;
WMinY:=Height ;
end;

procedure TFCtrlCai.CptExit(Sender: TObject);
begin if Cpt.Text<>'' then Cpt.Text:=BourreLaDonc(Cpt.Text,fb) ; end;

procedure TFCtrlCai.FDevisesChange(Sender: TObject);
begin
Dev:=FDevises.Value ;
if (Dev=V_PGI.DevisePivot) or (Dev='') then
   BEGIN
   Dev:=V_PGI.DevisePivot ; Decimale:=V_PGI.OkDecV ; Symbole:='' ;
   END else
   BEGIN
   RDev.Code:=Dev ; GetInfosDevise(RDev) ;
   Decimale:=RDev.Decimale ; Symbole:=RDev.Symbole ;
   END ;
end;

procedure TFCtrlCai.BGeneClick(Sender: TObject);
begin
if Cpt.Text='' then Exit ;
FicheGene(Nil,'',Cpt.Text,taConsult,0);
end;

procedure TFCtrlCai.BGLClick(Sender: TObject);
Var Crit : TCritEdt ;
    St : String ;
begin
//if FExercice.ItemIndex<=0 then Exit ; 
if FListe.Cells[0,1]='' then Exit ;
Fillchar(Crit,SizeOf(Crit),#0) ;
Crit.Date1:=StrToDate(CBDebDat.Items[1]) ; Crit.Date2:=StrToDate(CBDebDat.Items[CBDebDat.Items.Count-1]) ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neGL ;
InitCritEdt(Crit) ;
Crit.GL.ForceNonCentralisable:=TRUE ;
Crit.Cpt1:=Cpt.Text ; Crit.Cpt2:=Cpt.Text ;
St:=' (E_GENERAL="'+Cpt.Text+'") ' ;
//Crit.QualifPiece:=FNatureEcr.Value ;
Crit.DeviseSelect:=Dev ;
Crit.Etab:=FEtab.Value ;
Crit.SQLPLUS:='AND '+St ;
GLGeneralZoom(Crit) ;
end;

procedure TFCtrlCai.FDateCompta1KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFCtrlCai.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

procedure TFCtrlCai.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG6 12/01/05 Gestion Filtre V6
FreeAndnil(ObjFiltre);
if Parent is THPanel then Action:=caFree ;
end;

procedure TFCtrlCai.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFCtrlCai.OnSortGrid(Sender: TObject);
begin
  { CA - 12/08/2004 - Pas d'affichage du graphe si on n'est pas trié sur la date : sinon pas de sens ! }
  BGraph.Enabled := (FListe.SortedCol = 0);
end;



end.
