{***********UNITE*************************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : TreeView des rubriques comptables ou Budgétaires.
Suite ........ : Permet la visualisation, la modification, la création et la 
Suite ........ : suppression de rubriques.
Mots clefs ... : RUBRIQUES;COMPTA;BUDGET;
*****************************************************************}
unit TreeRubriques;

interface
                          
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, ComCtrls,hctrls,dbtables, ImgList, ent1, Hent1, Utob,
  HTB97, Menus,general,Tiers,section,budgene,budsect,HQuickRP, ColMemo, lookup, filtre,HPanel,
  Uiutil,HMsgBox,UtilDiv, Grids  ;

Procedure LanceTreeViewRubrique(LaRub : String ; Action : TActionFiche ; Budget : Boolean ) ;
procedure LanceTreeToutesRubriques(Action : TActionFiche ; Budget : Boolean ) ;

type
  TTreeRub = class(TForm)
    PetitesImages: TImageList;
    Grandesimages: TImageList;
    TreeView1: TTreeView;
    Splitter1: TSplitter;
    ListView1: TListView;
    Pages: TPageControl;
    PCritere: TTabSheet;
    Bevel1: TBevel;
    TRB_RUBRIQUE: THLabel;
    TRB_LIBELLE: TLabel;
    TRB_SIGNERUB: THLabel;
    TRB_TYPERUB: THLabel;
    TRB_CODEABREGE: THLabel;
    RB_RUBRIQUE: TEdit;
    RB_LIBELLE: TEdit;
    RB_SIGNERUB: THValComboBox;
    RB_TYPERUB: THValComboBox;
    RB_CODEABREGE: TEdit;
    Dock973: TDock97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    Dock974: TDock97;
    PanelBouton: TToolWindow97;
    BReduire: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BOuvrir: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    RB_RUBRIQUE2: TEdit;
    Label1: TLabel;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    CbCaption: TComboBox;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    TTCOMPTE: TToolWindow97;
    LEDCOMPTE: TLabel;
    EDCOMPTE: TEdit;
    EDEXCLUS: TEdit;
    LEDEXCLUS: TLabel;
    LEDCOMPTE2: TLabel;
    EDCOMPTE2: TEdit;
    LEDEXCLUS2: TLabel;
    EDEXCLUS2: TEdit;
    POPFO: TPopupMenu;
    BARREOUTILS: TMenuItem;
    POPFORUB: TMenuItem;
    POPFOCOMPTE: TMenuItem;
    POPFOAFF: TMenuItem;
    POPFOGRDICO: TMenuItem;
    POPFOPTTICO: TMenuItem;
    POPFOLISTE: TMenuItem;
    POPFODETAIL: TMenuItem;
    N2: TMenuItem;
    POPFOLVVIEW: TMenuItem;
    POPFOTVVIEW: TMenuItem;
    TTRUBRIQUE: TToolWindow97;
    CBNOMRUB: TComboBox;
    Label2: TLabel;
    ToolbarButton971: TToolbarButton97;
    TTLISTECOMPTE: TToolWindow97;
    FLISTE: THGrid;
    POPFOLISTECOMPTE: TMenuItem;
    RB_CLASSERUB: THValComboBox;
    TRB_CLASSERUB: THLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TreeView1Changing(Sender: TObject; Node: TTreeNode;
      var AllowChange: Boolean);
    procedure POPFOChangeStyleAff(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GoSelectionRubrique(Sender: TObject);
    procedure RB_RUBRIQUEDblClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure POPFOTVAfficheElement(Sender: TObject);
    procedure POPFOLVAfficheOuOuvre(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure POPFOBarreClick(Sender: TObject);
    procedure POPFOPopup(Sender: TObject);
    procedure POPFOLVAffiche(Sender: TObject);
    procedure POPFOTVAfficheOuOuvre(Sender: TObject);
    procedure EDCOMPTEDblClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
  private
    { Déclarations privées }
    LaTob,Laliste: TOB ;
    LaRub,FNomFiltre : String ;
    LeStyle : integer ;
    Action : TActionFiche ;
    Budget,AfficheOK: boolean ;
    NodeDeDebut: TTreeNode ;
    procedure OuvreLaForme ;
    procedure FermeLaForme ;
    function CreerLaTob(TobParent:Tob;Nom,Libelle : string;Image: integer) : TOB ;
    function AjoutEnfantNode(LaTobNode : Tob; nom : string;Inclus,Err : boolean) : Tob ;
    function AjoutEnfantFich(LaTobNode : Tob; Nom,Libelle : string;Inclus : boolean) : Tob ;
    procedure AjoutEnfantCpte(LaTobNode : Tob;  Compte,Exclus  : string;OnTableLibre : boolean;Lefb: TFichierBase) ;
    procedure AjoutSousNode(LaTobNode: Tob;Chaine: String;Inclus,OnTableLibre: boolean;Lefb: TFichierBase);
    function IsDansParent(LaTobNode: Tob;UneRub: String) : boolean ;
    procedure MetsBrancheFausse(LaTobNode: Tob;Indice: integer) ;
    procedure InitTree ;
    function IsLegende(Node: TTreeNode) : boolean ;  public
    procedure PutTobTree(Node: TTreeNode;TheTobNode: Tob)  ;
    procedure CreeOuAjoutLaListe(Ajout: boolean;Rub: string) ;
    function RecupereLeNom(Node: TTreeNode): string ;
    procedure AfficheZoom(Nom,typeaff: string);
    procedure RecupRubriqueEnSt(Compte1: string; ECompte: TEdit);
    Procedure AfficheLibelle(STypeRub: string; RubDeRub: boolean) ;
    procedure MetsAJourLaRubrique(Nom: string;TypeMaj: integer) ;
    procedure OnChangeTV(Node: TTreeNode);
  public
    { Déclarations publiques }
  end;

implementation
{$R *.DFM}

Uses Rubrique,Ed_Tools ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Enleve un mot dans une chaine
Mots clefs ... : CHAINE;
*****************************************************************}
procedure EnleveMotDansChaine(Mot: string ;var Chaine: string) ;
var P,L: integer ;
begin
P:=pos(Mot,Chaine) ; L:=length(Mot) ;
if (P<>0) AND (L<>0) then  delete(Chaine,P,L) ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Fonction d'agrandissement/reduction d'un multicritère
Mots clefs ... : MULTICRITERE;AGRANDISSEMENT;REDUCTION;
*****************************************************************}
Procedure ChangeListeCritT ( FF : TForm ; Plus : boolean ) ;
Var i   : integer ;
    C,G : TControl ;
begin
G:=Nil ;
for i:=0 to FF.ControlCount-1 do
    begin
    C:=FF.Controls[i] ;
    if C is TPageControl then
       begin
       if Plus then TPageControl(C).Align:=AlNone else TPageControl(C).Align:=AlTop ;
       TPageControl(C).Visible:=not Plus ;
       end else
       begin
       if ((C is TTreeView) Or (uppercase(C.Name)='TV')) then G:=C ;
       end ;
    end ;
TControl(FF.FindComponent('BAgrandir')).Visible:=Not Plus ;
TControl(FF.FindComponent('BReduire')).Visible:=Plus ;
if G<>Nil then if TTreeView(G).CanFocus then TTreeView(G).SetFocus ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Ajout une chaine avec le séparateur "\"
Mots clefs ... : SLASH;CHAINE;
*****************************************************************}
function SlashSep(const Path, S: String): String;
begin
  if AnsiLastChar(Path)^ <> '\' then
    Result := Path + '\' + S
  else
    Result := Path + S;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Enleve les crochets au début et a la fin d'une chaine entre crochets
Mots clefs ... : CHAINE;
*****************************************************************}
function NomSansLesDeuxCrochets(st: string): string ;
var i: integer ;
begin
i:=pos(']',st)-2 ;
if i<>0 then Result:=copy(st,2,i) else result:='' ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Renvoi un entier correspondant au type de rubriques associées
Mots clefs ... : RUBRIQUES;
*****************************************************************}
function QuelTypeAffichage(TypeAff: String) : integer ;
begin
result:=0 ;
if TypeAff=traduirememoire('Rubrique')          then result:=1 ;
if Typeaff=traduirememoire('Budget général')    then result:=2 ;
if Typeaff=traduirememoire('Section budgetaire')then result:=3 ;
if Typeaff=traduirememoire('Section analytique')then result:=4 ;
if Typeaff=traduirememoire('Compte général')    then result:=5 ;
if Typeaff=traduirememoire('Compte auxiliaire') then result:=6 ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /
Description .. : Lance un treeview sur les rubriques comptables ou budgetaires
Mots clefs ... : RUBRIQUE;COMPTA;BUDGET;
*****************************************************************}
procedure LanceTreeViewRubrique(LaRub : String ; Action : TActionFiche ; Budget : Boolean ) ;
var FTreeRub: TTreeRub;
    PP       : THPanel ;
begin
FTreeRub:=TTreeRub.Create(Application) ;
FTreeRub.LaRub:=LaRub ;
FTreeRub.Budget:=Budget ;
FTreeRub.Action:=Action ;
PP:=FindInsidePanel ;
if PP=Nil then
  begin try FTreeRub.ShowModal ;  finally FTreeRub.Free ;  end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(FTreeRub,PP) ;
  FTreeRub.Show ;
  end ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Lance un treeview sur toutes les rubriques comptables ou budgetaires
Mots clefs ... : RUBRIQUE;COMPTA;BUDGET;
*****************************************************************}
procedure LanceTreeToutesRubriques(Action : TActionFiche ; Budget : Boolean ) ;
var FTreeRub: TTreeRub;
begin
LanceTreeViewRubrique(W_W,Action,Budget) ;
end ;
//***
//
//   GESTION DE LA TOB
//
//***
procedure TTreeRub.CreeOuAjoutLaListe(Ajout: boolean;Rub: string) ;
var Q: TQuery; Sql,NomRub,Compte1,Exclus1,RubDeRub: string ;
begin
if not Ajout then
  begin
  if LaListe<>nil then LaListe.Free ;
  LaListe:=TOB.Create('RUBRIQUE',nil,-1) ;
  end ;
Sql:='SELECT * FROM RUBRIQUE' ;
if budget then Sql:=Sql+' WHERE RB_NATRUB="BUD"' else Sql:=Sql+' WHERE RB_NATRUB="CPT"' ;
if not Ajout then
  begin
  if RB_RUBRIQUE.Text<>''         then Sql:=Sql+' AND RB_RUBRIQUE>="'+RB_RUBRIQUE.Text+'" ' ;
  if RB_RUBRIQUE2.Text<>''        then Sql:=Sql+' AND RB_RUBRIQUE<="'+RB_RUBRIQUE2.Text+'" ' ;
  if RB_LIBELLE.Text<>''          then Sql:=Sql+' AND RB_LIBELLE LIKE "'+RB_LIBELLE.Text+'%" ' ;
  if RB_CODEABREGE.Text<>''       then Sql:=Sql+' AND RB_CODEABREGE LIKE "'+RB_CODEABREGE.Text+'%" ' ;
  if RB_CLASSERUB.Text<>''        then Sql:=Sql+' AND RB_CLASSERUB LIKE "'+RB_CLASSERUB.Value+'%" ' ;
//  if CHRUBDERUB.State=cbUnchecked then Sql:=Sql+' AND RB_RUBDERUB="-" ' ;
//  if CHRUBDERUB.State=cbChecked	  then Sql:=Sql+' AND RB_RUBDERUB="X" ';
//  if not CHRUBTABLIB.Checked      then Sql:=Sql+' AND RB_TABLELIBRE<>"X" ' else Sql:=Sql+' AND RB_TABLELIBRE="X" ';
  if RB_SIGNERUB.Text<>''         then Sql:=Sql+' AND RB_SIGNERUB="'+RB_SIGNERUB.Value+'" ' ;
  if RB_TYPERUB.Text<>''          then Sql:=Sql+' AND RB_TYPERUB="'+RB_TYPERUB.Value+'" ' ;
  end else
     Sql:=Sql+' AND RB_RUBRIQUE="'+Rub+'" ' ;
Q:=OpenSql(Sql,True) ;
if not Q.eof then LaListe.LoadDetailDB('RUBRIQUE','','',Q,True,not Ajout) ;
Ferme(Q) ;
end ;

procedure TTreeRub.OuvreLaForme ;
var  i: integer ;    TT : TQRProgressForm ;
begin
CreeOuAjoutLaListe(False,'') ;
if LaRub<>W_W then LaTob:=AjoutEnfantNode(nil,LaRub,True,False)
else
  begin
  LaTob:=CreerlaTob(nil,'','',0) ;
  TT:=DebutProgressForm (Self,'Explorateur de rubriques','Préparation des données ',LaListe.Detail.Count,True,True) ;
  for i:=0 to LaListe.Detail.count-1 do
    begin
    if TT.Canceled then break ;
    AjoutEnfantNode(LaTob,LaListe.Detail[i].GetValue('RB_RUBRIQUE'),True,False);
    TT.Value:=i ;
    TT.SubText:='Traitement de la rubrique : '+IntToStr(i) ;
    end ;
  TT.Free;
  end ;
LeStyle:=0 ;
end;

procedure TTreeRub.FermeLaForme ;
begin
LaTob.Free ;
LaListe.Free ;
end;

function TTreeRub.CreerLaTob(TobParent:Tob;Nom,Libelle : string;Image: integer) : TOB ;
begin
Result:=TOB.Create('$VIRTUAL',TobParent,-1) ;
if Result<>nil then
  begin
  Result.AddChampSup('RUBRIQUE',True) ; Result.PutValue('RUBRIQUE',Nom) ;
  Result.AddChampSup('LIBELLE',True) ;  Result.PutValue('LIBELLE',Libelle) ;
  Result.AddChampSup('IMAGE',True) ;    Result.PutValue('IMAGE',Image) ;
  end ;
end ;

function TTreeRub.AjoutEnfantFich(LaTobNode : Tob; Nom,libelle : string;Inclus : boolean) : Tob ;
var Image: integer ;
begin
if inclus then Image:=2 else Image:=5 ;
Result:=CreerLaTob(LaTobNode,Nom,LIbelle,Image) ;
end ;

procedure TTreeRub.AjoutEnfantCpte(LaTobNode : Tob; Compte,Exclus : string;OnTableLibre : boolean;Lefb: TFichierBase) ;
var Image: integer ; Sql,Where,notWhere: string ; Q: TQuery ;
begin
//if inclus then Image:=2 else Image:=5 ;
Where:=AnalyseCompte(Compte,Lefb,False,OntableLibre) ;
notWhere:=AnalyseCompte(Exclus,Lefb,True,OntableLibre) ;
case Lefb of
  fbBudgen             : Sql:='SELECT BG_GENERAL,BS_LIBELLE FROM BG_BUDGENE ' ;
  fbBudSec1..fbBudSec5 : Sql:='SELECT BS_SECTION,BS_LIBELLE FROM BS_BUDSECT ' ;
  fbAxe1..fbAxe5       : Sql:='SELECT S_SECTION,S_LIBELLE FROM SECTION ' ;
  fbGene               : Sql:='SELECT G_GENERAL,G_LIBELLE FROM GENERAUX ' ;
  fbAux                : Sql:='SELECT T_AUXILIAIRE,T_LIBELLE FROM TIERS ' ;
  end ;
if (Where<>'') or (NotWhere<>'') then Sql:=Sql+' WHERE ' ;
Sql:=Sql+Where ;
if (Where<>'') and (NotWhere<>'') then Sql:=Sql+' AND ' ;
Sql:=Sql+NotWhere ;
Q:=OpenSql(Sql,True) ;
while not Q.Eof do
  begin
  AjoutEnfantFich(LaTobNode,Q.Fields[0].AsString,Q.Fields[1].AsString,True) ;
  Q.Next ;
  end ;
ferme(Q);
end ;


function TTreeRub.AjoutEnfantNode(LaTobNode : Tob; nom : string;Inclus,Err : boolean) : Tob ;
var Image: integer ;RubDeRub,OnTableLibre: boolean; NomRub,Libelle,Compte,Exclus: string ;NewTobNode,TobRub: Tob ; LeFb: TFichierBase ;
begin
NewTobNode:=nil ; Result:=nil ;
if Nom='' then Exit ;
TobRub:=LaListe.FindFirst(['RB_RUBRIQUE'],[Nom],True) ;
if TobRub=nil then
  begin
  CreeOuAjoutLaListe(True,Nom) ;
  TobRub:=LaListe.FindFirst(['RB_RUBRIQUE'],[Nom],True) ;
  end ;
if TobRub<>nil then
  begin
  NomRub:=TobRub.GetValue('RB_RUBRIQUE') ;
  Libelle:=TobRub.GetValue('RB_LIBELLE') ;
  RubDeRub:=(TobRub.GetValue('RB_CLASSERUB')='RDR') ;
//  RubDeRub:=(TobRub.GetValue('RB_RUBDERUB')='X') ;
  Compte:=TobRub.GetValue('RB_COMPTE1')+';'+TobRub.GetValue('RB_COMPTE2') ;
  Exclus:=TobRub.GetValue('RB_EXCLUSION1')+';'+TobRub.GetValue('RB_EXCLUSION2') ;
  OnTableLibre:=(TobRub.GetValue('RB_CLASSERUB')='TLI') ;
//  OnTableLibre:=(TobRub.GetValue('RB_TABLELIBRE')='X') ;
  LeFb:=QuelEstLefb(budget,TobRub.GetValue('RB_TYPERUB'),TobRub.GetValue('RB_AXE')) ;
  if inclus then Image:=0 else Image:=3 ;
  NewTobNode:=CreerLaTob(LaTobNode,NomRub,Libelle,Image) ;
  if not Err then
    begin
    AjoutSousNode(NewTobNode,Compte,True,OnTableLibre,Lefb) ;
    AjoutSousNode(NewTobNode,Exclus,False,OnTableLibre,Lefb) ;
    end ;
  end ;
Result:=NewTobNode ;
end ;

procedure TTreeRub.AjoutSousNode(LaTobNode: Tob;Chaine: String;Inclus,OnTableLibre: boolean;Lefb: TFichierBase);
var Err: boolean; UneRub: string ; NewFichNode: Tob ;
begin
NewFichNode:=nil ;
while Chaine<>'' do
  begin
  UneRub:=ReadTokenStV(Chaine) ;
  if UneRub<>'' then
    begin
    Err:=IsDansParent(LaTobNode,UneRub) ;
    NewFichNode:=AjoutEnfantNode(LaTobNode,UneRub,Inclus,Err) ;
    if Err and (NewFichNode<>nil) then MetsBrancheFausse(NewFichNode,6) ;
    end ;
  end ;
end ;

function TTreeRub.IsDansParent(LaTobNode: Tob;UneRub: String): boolean ;
begin
if LaTobNode=nil then Result:=False
else if LaTobNode.GetValue('RUBRIQUE')=UneRub then Result:=True
     else Result:=IsDansParent(LaTobNode.Parent,UneRub) ;
end ;

procedure TTreeRub.MetsBrancheFausse(LaTobNode: Tob;Indice: integer) ;
begin
if LaTobNode<>nil then
  begin
  MetsBrancheFausse(LaTobNode.Parent,Indice) ;
  LaTobNode.PutValue('Image',Indice) ;
  end ;
end ;
//***
//
//   GESTION DE LA FORME
//
//***
procedure TTreeRub.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FermeLaForme ;
//RegSaveToolbarPos(Self,'TREERUBRIQUE') ;
if Parent is THPanel then Action:=caFree ;
end ;

procedure TTreeRub.InitTree ;
var Node,ChildNode: TTreeNode ;
begin
TreeView1.Items.Clear ;
Node:=TreeView1.items.AddChild(nil,TraduireMemoire('Legende')); Node.ImageIndex:=0;
Node.ImageIndex:=0;   Node.SelectedIndex:=1;
TreeView1.Selected:=TreeView1.Items[0] ;
ChildNode:=TreeView1.items.AddChild(Node,TraduireMemoire('Rubrique fermer'));
ChildNode.ImageIndex:=0;   ChildNode.SelectedIndex:=0;
ChildNode:=TreeView1.items.AddChild(Node,TraduireMemoire('Rubrique sélectionée'));
ChildNode.ImageIndex:=1;   ChildNode.SelectedIndex:=1;
ChildNode:=TreeView1.items.AddChild(Node,TraduireMemoire('Comptes de la rubrique'));
ChildNode.ImageIndex:=2;   ChildNode.SelectedIndex:=2;
ChildNode:=TreeView1.items.AddChild(Node,TraduireMemoire('Rubrique except. fermer'));
ChildNode.ImageIndex:=3;   ChildNode.SelectedIndex:=3;
ChildNode:=TreeView1.items.AddChild(Node,TraduireMemoire('Rubrique except. sélectionnée'));
ChildNode.ImageIndex:=4;   ChildNode.SelectedIndex:=4;
ChildNode:=TreeView1.items.AddChild(Node,TraduireMemoire('Comptes excpt. de la rubrique'));
ChildNode.ImageIndex:=5;   ChildNode.SelectedIndex:=5;
ChildNode:=TreeView1.items.AddChild(Node,TraduireMemoire('Erreur'));
ChildNode.ImageIndex:=6;   ChildNode.SelectedIndex:=6;
NodeDeDebut:=ChildNode ;
end ;

procedure TTreeRub.PutTobTree(Node: TTreeNode;TheTobNode: Tob)  ;
var i: integer ; ChildNode: TTreeNode ;
begin
if TheTobNode<>nil then
  begin
  ChildNode:=TreeView1.items.AddChild(Node,'['+TheTobNode.GetValue('RUBRIQUE')+'] '+TheTobNode.GetValue('LIBELLE'));
  ChildNode.ImageIndex:=TheTobNode.GetValue('IMAGE');
  if ChildNode.ImageIndex<>6 then ChildNode.SelectedIndex:=TheTobNode.GetValue('IMAGE')+1
                             else ChildNode.SelectedIndex:=TheTobNode.GetValue('IMAGE') ;
  for i:=0 to TheTobNode.Detail.Count-1 do PutTobTree(ChildNode,TheTobNode.Detail[i])  ;
  end;
end ;

function TTreeRub.IsLegende(Node: TTreeNode) : boolean ;
var i: integer ;
begin
i:=0 ;
Result:=False ;
if NodeDeDebut=nil then result:=True ;
while result=False do
  begin
  if TreeView1.Items[i]=Node then Result:=True ;
  if TreeView1.Items[i]=NodeDeDebut then break ;
  inc(i) ;
  end ;
end ;

procedure TTreeRub.RecupRubriqueEnSt(Compte1: string; ECompte: TEdit);
var  St,St1,separator: string ; j: integer ;
begin
St:=Compte1 ; St1:='' ; ECompte.Text:='' ;  separator:='' ;
While St<>'' do
  begin
  St1:=ReadTokenSt(St) ;
  if St1<>'' then
    begin
    ECompte.Text:=ECompte.Text+separator ;
    for j:=1 to Length(St1)do begin ECompte.Text:=ECompte.Text+St1[j] ; separator:=';' ; end ;
    end ;
  end ;
end ;


Procedure TTreeRub.AfficheLibelle(STypeRub: string; RubDeRub: boolean) ;
Var St : String[3] ;
    A,B,C : Char ;
BEGIN
St:=STypeRub ; A:=St[1] ; C:=St[2] ; B:=St[3] ;
if Action<>taConsult then BDelete.Enabled:=True ;
EDCOMPTE.Enabled:=True ;
if RubDeRub then
  begin
  LEdCompte.Caption:=CbCaption.Items[9] ; LEdExclus.Caption:=CbCaption.Items[10] ;
  LEdCompte2.Caption:=LEdCompte.Caption ; LEdExclus2.Caption:=LEdExclus.Caption ;
  end else
  begin
  Case A of
    'A':BEGIN LEdCompte.Caption:=CbCaption.Items[0] ; LEdExclus.Caption:=CbCaption.Items[1] ; END ;
    'B':BEGIN LEdCompte.Caption:=CbCaption.Items[2] ; LEdExclus.Caption:=CbCaption.Items[3] ; END ;
    'T':BEGIN LEdCompte.Caption:=CbCaption.Items[4] ; LEdExclus.Caption:=CbCaption.Items[5] ; END ;
    'G':BEGIN LEdCompte.Caption:=CbCaption.Items[6] ; LEdExclus.Caption:=CbCaption.Items[7] ; END ;
    end ;
  LEdCompte2.Caption:=LEdCompte.Caption ; LEdExclus2.Caption:=LEdExclus.Caption ;
  if C='/' then
    begin
    EDCOMPTE2.Enabled:=True ;
    Case B of
      'A':BEGIN LEdCompte2.Caption:=CbCaption.Items[0] ; LEdExclus2.Caption:=CbCaption.Items[1] ; END ;
      'B':BEGIN LEdCompte2.Caption:=CbCaption.Items[2] ; LEdExclus2.Caption:=CbCaption.Items[3] ; END ;
      'T':BEGIN LEdCompte2.Caption:=CbCaption.Items[4] ; LEdExclus2.Caption:=CbCaption.Items[5] ; END ;
      'G':BEGIN LEdCompte2.Caption:=CbCaption.Items[6] ; LEdExclus2.Caption:=CbCaption.Items[7] ; END ;
      end ;
    end ;
  end;
EDEXCLUS.Enabled:=EDCOMPTE.Enabled ;   LEdCompte.Enabled:=EDCOMPTE.Enabled ;   LEdExclus.Enabled:=EDCOMPTE.Enabled ;
EDEXCLUS2.Enabled:=EDCOMPTE2.Enabled ; LEdCompte2.Enabled:=EDCOMPTE2.Enabled ; LEdExclus2.Enabled:=EDCOMPTE2.Enabled ;
end ;

procedure TTreeRub.TreeView1Changing(Sender: TObject; Node: TTreeNode;  var AllowChange: Boolean);
begin
OnChangeTV(Node) ;
end;

procedure TTreeRub.OnChangeTV(Node: TTreeNode);
var Q: TQuery; i: integer ; ListItem: TListItem;  TobRub: Tob; OnTableLibre: boolean; LeFb: TFichierBase ;
TypeCpte,Sql,Where,NotWhere,Compte,Exclus,Compte1,Exclus1,Compte2,Exclus2: string ;
begin
// Cache les controles liés aux rubriques
BDelete.Enabled:=False ;
EDCOMPTE.Enabled:=False ; EDEXCLUS.Enabled:=EDCOMPTE.Enabled ;
LEdCompte.Enabled:=EDCOMPTE.Enabled ; LEdExclus.Enabled:=EDCOMPTE.Enabled ;
EDCOMPTE2.Enabled:=False ; EDEXCLUS2.Enabled:=EDCOMPTE.Enabled ;
LEdCompte2.Enabled:=EDCOMPTE2.Enabled ; LEdExclus2.Enabled:=EDCOMPTE.Enabled ;
CBNOMRUB.Text:=RecupereLeNom(Node) ;
if CBNOMRUB.Items.Count=0 then CBNOMRUB.Items.InsertObject(0,CBNOMRUB.Text,Node) ;
for i:=0 to CBNOMRUB.Items.Count-1 do if  uppercase(CBNOMRUB.Text)=uppercase(CBNOMRUB.Items[i]) then Break ;
if i=CBNOMRUB.Items.Count then CBNOMRUB.Items.InsertObject(0,CBNOMRUB.Text,Node) ;
EDCOMPTE.Text:='' ; EDEXCLUS.Text:='' ; EDCOMPTE2.Text:='' ; EDEXCLUS2.Text:='' ;
ListView1.items.Clear ;
if IsLegende(Node) then Exit ;
TobRub:=LaListe.FindFirst(['RB_RUBRIQUE'],[NomSansLesDeuxCrochets(Node.text)],True) ;
if TobRub<>nil then
  begin
  Compte1:=TobRub.GetValue('RB_COMPTE1') ;    Compte2:=TobRub.GetValue('RB_COMPTE2') ;
  Exclus1:=TobRub.GetValue('RB_EXCLUSION1') ; Exclus2:=TobRub.GetValue('RB_EXCLUSION2') ;
  AfficheLibelle(TobRub.GetValue('RB_TYPERUB'),(TobRub.GetValue('RB_CLASSERUB')='RDR') ) ;
//  AfficheLibelle(TobRub.GetValue('RB_TYPERUB'),(TobRub.GetValue('RB_RUBDERUB')='X') ) ;
  RecupRubriqueEnSt(Compte1,EDCOMPTE) ;   RecupRubriqueEnSt(Exclus1,EDEXCLUS) ;
  RecupRubriqueEnSt(Compte2,EDCOMPTE2) ;  RecupRubriqueEnSt(Exclus2,EDEXCLUS2) ;
  if Node.Count<>0 then
    begin
    for i:=0 to Node.Count-1 do
      begin
      ListItem:=ListView1.items.Add ;
      ListItem.Caption:=Node.item[i].Text ;
      if Node.item[i].ImageIndex<3 then ListItem.ImageIndex:=0 else ListItem.ImageIndex:=3 ;
      ListItem.SubItems.Add(traduirememoire('Rubrique')) ;
      end;
    end else
    begin
    Compte:=Compte1+';'+Compte2 ;
    Exclus:=Exclus1+';'+Exclus2 ;
    OnTableLibre:=(TobRub.GetValue('RB_CLASSERUB')='TLI') ;
//    OnTableLibre:=(TobRub.GetValue('RB_TABLELIBRE')='X') ;
    LeFb:=QuelEstLefb(budget,TobRub.GetValue('RB_TYPERUB'),TobRub.GetValue('RB_AXE')) ;
    Where:=AnalyseCompte(Compte,Lefb,False,OntableLibre) ;
    notWhere:=AnalyseCompte(Exclus,Lefb,True,OntableLibre) ;
    case Lefb of
      fbBudgen             : begin Sql:='SELECT BG_GENERAL,BS_LIBELLE FROM BG_BUDGENE ' ; TypeCpte:=traduirememoire('Budget général') ; end ;
      fbBudSec1..fbBudSec5 : begin Sql:='SELECT BS_SECTION,BS_LIBELLE FROM BS_BUDSECT ' ; TypeCpte:=traduirememoire('Section budgetaire') ; end ;
      fbAxe1..fbAxe5       : begin Sql:='SELECT S_SECTION,S_LIBELLE FROM SECTION ' ;      TypeCpte:=traduirememoire('Section analytique') ; end ;
      fbGene               : begin Sql:='SELECT G_GENERAL,G_LIBELLE FROM GENERAUX ' ;     TypeCpte:=traduirememoire('Compte général') ; end ;
      fbAux                : begin Sql:='SELECT T_AUXILIAIRE,T_LIBELLE FROM TIERS ' ;     TypeCpte:=traduirememoire('Compte auxiliaire') ; end ;
      end ;
    if (Where<>'') or (NotWhere<>'') then Sql:=Sql+' WHERE ' ;
    Sql:=Sql+Where ;
    if (Where<>'') and (NotWhere<>'') then Sql:=Sql+' AND ' ;
    Sql:=Sql+NotWhere ;
    Q:=OpenSql(Sql,True) ;
    while not Q.Eof do
      begin
      ListItem:=ListView1.items.Add ;
      ListItem.Caption:='['+Q.Fields[0].AsString+'] '+Q.Fields[1].AsString ;
      ListItem.SubItems.Add(TypeCpte) ;
      ListItem.ImageIndex:=2 ;
      Q.Next ;
      end ;
    ferme(Q);
    end ;
  end ;
end;


procedure TTreeRub.POPFOChangeStyleAff(Sender: TObject);
begin
ListView1.ViewStyle:=TViewStyle(TmenuItem(Sender).Tag) ;
TmenuItem(Sender).Checked:=True ;
end;

procedure TTreeRub.POPFOTVAfficheElement(Sender: TObject);
var Node: TTreeNode ;
begin
Node:=TreeView1.Selected ;
if IsLegende(Node) then Exit ;
AfficheZoom(TreeView1.Selected.Text,traduirememoire('Rubrique')) ;
end;

procedure TTreeRub.POPFOLVAffiche(Sender: TObject);
begin
AfficheZoom(ListView1.ItemFocused.Caption,ListView1.ItemFocused.SubItems.Strings[0]) ;
end;

procedure TTreeRub.POPFOLVAfficheOuOuvre(Sender: TObject);
var nom: string ;i: integer ;
begin
if ListView1.ItemFocused.SubItems.Strings[0]=traduirememoire('Rubrique')
then begin for i:=0 to TreeView1.Selected.Count-1 do if TreeView1.Selected.Item[i].Text=ListView1.ItemFocused.Caption then begin TreeView1.Selected:=TreeView1.Selected.Item[i]; break ; end ; end
else AfficheZoom(ListView1.ItemFocused.Caption,ListView1.ItemFocused.SubItems.Strings[0]) ;
end;

procedure TTreeRub.POPFOTVAfficheOuOuvre(Sender: TObject);
var  Node: TTreeNode ;
begin
Node:=TreeView1.Selected ;
if IsLegende(Node) then Exit ;
if (Node.ImageIndex<>2) and (Node.ImageIndex<>5) and (Node.Count=0) then
   AfficheZoom(TreeView1.Selected.Text,traduirememoire('Rubrique')) ;
end;


procedure TTreeRub.AfficheZoom(Nom,typeaff: string);
var NomTmp: string ;
begin
NomTmp:=NomSansLesDeuxCrochets(Nom) ;
case QuelTypeAffichage(TypeAff) of
  1 : begin ParametrageRubrique(NomTmp,Action,Budget) ; MetsAJourLaRubrique(NomTmp,3) ; end ;
  2 : FicheBudgene(nil,'',NomTmp,Action,0) ;
  3 : FicheBudsect(Nil,'',NomTmp,Action,0) ;
  4 : FicheSection(Nil,'',NomTmp,Action,0) ;
  5 : FicheGene(nil,'',NomTmp,Action,0) ;
  6 : FicheTiers(Nil,'',NomTmp,Action,0) ;
  end ;
end;

procedure TTreeRub.BChercheClick(Sender: TObject);
var i: integer ; TT : TQRProgressForm ;
begin
OuvreLaForme ;
InitTree ; CBNOMRUB.Items.Clear ;
if (LaRub=W_W) then
  begin
  TT:=DebutProgressForm (Self,'Explorateur de rubriques','Préparation de l''éxplorateur',LaTob.Detail.Count,True,True) ;
  for i:=0 to LaTob.Detail.Count-1 do
    begin
    if TT.Canceled then break ;
    PutTobTree(nil,LaTob.Detail[i]);
    TT.Value:=i ;
    TT.SubText:='Traitement de la rubrique : '+IntToStr(i) ;
    end ;
  TT.Free;
  end else
    PutTobTree(nil,LaTob) ;
if action<>taConsult then begin BInsert.Enabled:=True ; end ; 
end;

procedure TTreeRub.FormShow(Sender: TObject);
begin
FNomFiltre:='TREERUB' ;  
ChargeFiltre(FNomFiltre,FFiltres,Pages) ;
InitTree ;
end;

procedure TTreeRub.GoSelectionRubrique(Sender: TObject);
begin
TreeView1.Selected:=TTreeNode(CBNOMRUB.items.Objects[CBNOMRUB.ItemIndex]) ;
TreeView1.SetFocus ;
end;

function TTreeRub.RecupereLeNom(Node: TTreeNode): string ;
begin
if Node.Parent<>nil then
  begin
  Result:=RecupereLeNom(Node.Parent) ;
  Result:=SlashSep(Result,Node.Text) ;
  end
  else
    Result:=Node.Text;
end ;

procedure TTreeRub.RB_RUBRIQUEDblClick(Sender: TObject);
begin
LookupList(TControl(Sender),'','RUBRIQUE','RB_RUBRIQUE','RB_LIBELLE','','RB_RUBRIQUE', True,0)  ;
end;

procedure TTreeRub.BAgrandirClick(Sender: TObject);
begin
ChangeListeCritT(Self,True) ;
end;

procedure TTreeRub.BReduireClick(Sender: TObject);
begin
ChangeListeCritT(Self,False) ;
end;


procedure TTreeRub.BCreerFiltreClick(Sender: TObject);
begin
NewFiltre(FNomFiltre,FFiltres,Pages) ;
end;

procedure TTreeRub.BSaveFiltreClick(Sender: TObject);
begin
SaveFiltre(FNomFiltre,FFiltres,Pages) ;
end;

procedure TTreeRub.BDelFiltreClick(Sender: TObject);
begin
DeleteFiltre(FNomFiltre,FFiltres) ;
end;

procedure TTreeRub.BRenFiltreClick(Sender: TObject);
begin
RenameFiltre(FNomFiltre,FFiltres) ;
end;

procedure TTreeRub.BNouvRechClick(Sender: TObject);
begin
VideFiltre(FFiltres,Pages) ;
end;

procedure TTreeRub.FFiltresChange(Sender: TObject);
begin
LoadFiltre(FNomFiltre,FFiltres,Pages) ;
end;

procedure TTreeRub.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TTreeRub.FormCreate(Sender: TObject);
begin
//RegLoadToolbarPos(Self,'TREERUBRIQUE') ;
end;

procedure TTreeRub.POPFOBarreClick(Sender: TObject);
begin
TMenuItem(Sender).Checked:=not TMenuItem(Sender).Checked ;
TTRubrique.Visible:=POPFORUB.Checked ;
TTCompte.Visible :=POPFOCOMPTE.Checked ;
TTLISTECompte.Visible :=POPFOLISTECOMPTE.Checked ;
end;

procedure TTreeRub.POPFOPopup(Sender: TObject);
var PopListView,PopTreeView,Test: boolean ; Node: TTreeNode ; TypAff: string ;
begin
PopListView:=False ;
PopTreeView:=False ;
POPFORUB.Checked    :=TTRubrique.Visible ;
POPFOCOMPTE.Checked:=TTCompte.Visible ;
POPFOLISTECOMPTE.Checked:=TTListeCompte.Visible ;
Test:=(Sender is TTreeView) ;
if TreeView1.Focused then
  begin
  Node:=TreeView1.Selected ;
  if IsLegende(Node) then POPFOTVVIEW.Enabled:=False else POPFOTVVIEW.Enabled:=True ;
  PopTreeView:=True ;
  end ;
if ListView1.Focused then
  begin
  TypAff:='' ;  POPFOLVVIEW.Enabled:=True;
  if ListView1.ItemFocused<>nil then TypAff:=ListView1.ItemFocused.SubItems.Strings[0] ;
  case QuelTypeAffichage(TypAff) of
    1 : POPFOLVVIEW.Caption:=TraduireMemoire('&Voir la rubrique') ;
    2 : POPFOLVVIEW.Caption:=TraduireMemoire('&Voir le compte budgetaire') ;
    3 : POPFOLVVIEW.Caption:=TraduireMemoire('&Voir la section budgetaire ') ;
    4 : POPFOLVVIEW.Caption:=TraduireMemoire('&Voir la section analytique') ;
    5 : POPFOLVVIEW.Caption:=TraduireMemoire('&Voir le compte général') ;
    6 : POPFOLVVIEW.Caption:=TraduireMemoire('&Voir le compte auxiliaire') ;
    else POPFOLVVIEW.Enabled:=False ;
    end ;
  PopListView:=True ;
  end ;
POPFOLVVIEW.Visible:=PopListView ;
POPFOAFF.Visible:=PopListView ;
POPFOTVVIEW.Visible:=PopTreeView ;
N2.Visible:=(PopTreeView or PopListView) ;
end;

procedure TTreeRub.EDCOMPTEDblClick(Sender: TObject);
var NomTmp: string ;
begin
NomTmp:=NomSansLesDeuxCrochets(TreeView1.Selected.Text) ;
ParametrageRubrique2(NomTmp,Action,Budget) ;
MetsAJourLaRubrique(NomTmp,3) ;
end;

// TypeMaj : 1:Ajout, 2:Suppr, 3:Modif
procedure TTreeRub.MetsAJourLaRubrique(Nom: string;TypeMaj: integer) ;
var TobRub: Tob ;
begin
TobRub:=nil ;
if Action=taConsult then Exit ;
//if CHRUBDERUB.State=cbUnchecked then
if RB_CLASSERUB.Value='RDR' then
  begin
  case TypeMaj of
  3: begin
     TobRub:=LaListe.FindFirst(['RB_RUBRIQUE'],[Nom],True) ;
     if TobRub<>nil then TobRub.Free ;
     CreeOuAjoutLaListe(True,Nom) ;
     end;
  2: TreeView1.items.Delete(TreeView1.Selected) ;
  1: begin
     TobRub:=AjoutEnfantNode(LaTob,Nom,True,False) ;
     PutTobTree(nil,TobRub);
     end ;
  end ;
  OnChangeTV(TreeView1.Selected) ;
  end
  else
    BChercheClick(nil) ;
end ;

procedure TTreeRub.BInsertClick(Sender: TObject);
var NomTmp,NomTmp2: string ;
begin
NomTmp:='' ;
NomTmp2:=ParametrageRubrique(NomTmp,Action,Budget,True) ;
MetsAJourLaRubrique(NomTmp2,1) ;
end;

procedure TTreeRub.BDeleteClick(Sender: TObject);
var Nom,NomRubrique: string ;Q: TQuery ; Cpt1,Exl1,Cpt2,Exl2: string ;
begin
Nom:=NomSansLesDeuxCrochets(TreeView1.Selected.Text) ;
if HShowMessage(TraduireMemoire('1;Suppression de rubriques;Etes-vous sur de vouloir supprimer la rubrique : %%;W;YN;N;N'),Nom,'')=mrYes then
  begin
  //Suppression en cascade !
//  Q:=OpenSql('SELECT * FROM RUBRIQUE WHERE RB_RUBDERUB="X" AND'
  Q:=OpenSql('SELECT * FROM RUBRIQUE WHERE RB_CLASSERUB="RDR" AND'
            +'(RB_COMPTE1 LIKE "%'+Nom+'%" OR  RB_EXCLUSION1 LIKE "%'+Nom+'%" '
            +'OR RB_COMPTE2 LIKE "%'+Nom+'%" OR  RB_EXCLUSION2 LIKE "%'+Nom+'%") ',False) ;
  while not Q.Eof do
    begin
    Q.Edit ;
    Cpt1:=Q.FindField('RB_COMPTE1').AsString ;    EnleveMotDansChaine(Nom+',',Cpt1) ; Q.FindField('RB_COMPTE1').AsString:=Cpt1 ;
    Exl1:=Q.FindField('RB_EXCLUSION1').AsString ; EnleveMotDansChaine(Nom+',',Exl1) ; Q.FindField('RB_EXCLUSION1').AsString:=Exl1 ;
    Cpt2:=Q.FindField('RB_COMPTE2').AsString ;    EnleveMotDansChaine(Nom+',',Cpt2) ; Q.FindField('RB_COMPTE2').AsString:=Cpt2 ;
    Exl2:=Q.FindField('RB_EXCLUSION2').AsString ; EnleveMotDansChaine(Nom+',',Exl2) ; Q.FindField('RB_EXCLUSION2').AsString:=Exl2 ;
    Q.Post ;
    Q.Next ;
    end ;
  ferme(Q) ;
  // Suppression de la rubrique
  ExecuteSql('DELETE FROM RUBRIQUE WHERE RB_RUBRIQUE="'+Nom+'"') ;
  MetsAJourLaRubrique(Nom,2) ;
  end ;
end;

procedure TTreeRub.ToolbarButton971Click(Sender: TObject);
var Tobres: tob ;
begin
POPFOLISTECOMPTE.Checked:=True ;
TTLISTECompte.Visible :=POPFOLISTECOMPTE.Checked ;
tobres:=AnalyseCompteRubrique(NomSansLesDeuxCrochets(TreeView1.Selected.Text)) ;
if tobres<>nil then
  begin
  TobRes.PutGridDetail(Fliste,False,False,'',true) ;
  TobRes.free ;
  end ;
end;

end.
