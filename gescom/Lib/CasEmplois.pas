unit CasEmplois;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Mask, Hctrls, Grids, ComCtrls, HSysMenu, UTOB, EntGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,HPdfPrev,UtileAGL,
{$ELSE}
      Db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$IFDEF V530}
      EdtDoc,
{$ELSE}
      EdtRDoc,
{$ENDIF}
{$ENDIF}
  NomenUtil, Menus, Buttons, Math, M3FP, HTB97,
  UtilArticle, AglInitGc, TarifUtil, TntGrids ;

procedure Entree_CasEmploi(Parms: array of variant; nb: integer);

type
  TFCasEmplois = class(TForm)
    PANChoix: TPanel;
    GA_CODEARTICLE: THCritMaskEdit;
    Label1: TLabel;
    GA_ARTICLE: TEdit;
    GNE_NOMENCLATURE: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    HMTrad: THSystemMenu;
    TV_NLIG: TTreeView;
    G_NLIG: THGrid;
    PopupG_NLIG: TPopupMenu;
    PopupTV_NLIG: TPopupMenu;
    PopG_NLIG_A: TMenuItem;
    PopG_NLIG_N: TMenuItem;
    PopTV_NLIG_O: TMenuItem;
    PopTV_NLIG_F: TMenuItem;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    bPrint: TToolbarButton97;
    BInfos: TToolbarButton97;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure G_NLIGRowEnter(Sender: TObject; Ou: Integer;
                             var Cancel: Boolean; Chg: Boolean);
    procedure GA_CODEARTICLEElipsisClick(Sender: TObject);
    procedure GA_CODEARTICLEExit(Sender: TObject);
    procedure PopG_NLIG_AClick(Sender: TObject);
    procedure PopG_NLIG_NClick(Sender: TObject);
    procedure PopTV_NLIG_NClick(Sender: TObject);
    procedure PopTV_NLIG_FClick(Sender: TObject);
    procedure PopTV_NLIG_OClick(Sender: TObject);
    procedure TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
                             Shift: TShiftState; X, Y: Integer);
    procedure bPrintClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    TobNomen : TOB;
    TOBArt : TOB;
    Old_CODEARTICLE : String;
    function ChercheLigne (Valeur : string) : integer;
  public
    { Déclarations publiques }
  end;

var
  FCasEmplois: TFCasEmplois;
  Nomen_Open : boolean;
  Row_Cur, Col_Cur : integer;

implementation

uses NomenLig;

{$R *.DFM}

//=============================================================================
//  Point d'entree : parametre 0 indique NOMENCLATURE=OUI (defaut) ou NOMENCLATURE=NON
//                   parametre 1 indique l'article à traiter
//=============================================================================
procedure Entree_CasEmploi(Parms: array of variant; nb: integer);
var
    st1, st2 : string;
    Var_Appel : string;
begin
Nomen_Open := True;
Var_Appel := string(Parms[0]);
st1 := Copy(Var_Appel, 0, Pos('=',Var_Appel) - 1);
if st1 = 'NOMENCLATURE' then
    begin
    st2 := Copy(Var_Appel, Pos('=',Var_Appel) + 1, 255);
    if st2 = 'NON' then Nomen_Open := False;
    end;
FCasEmplois := TFCasEmplois.Create(Application);
FCasEmplois.GA_CODEARTICLE.Text := string(Parms[1]);
FCasEmplois.ShowModal;
end;

//=============================================================================
procedure TFCasEmplois.FormCreate(Sender: TObject);
begin
TobNomen := TOB.Create('', nil, -1);
TOBArt := TOB.Create('ARTICLE', nil, -1);
end;

//=============================================================================
procedure TFCasEmplois.FormShow(Sender: TObject);
var
    Cancel : boolean;
begin
Old_CODEARTICLE := '';
G_NLIG.Cells[0, 0] := 'Nomenclature';
G_NLIG.Cells[1, 0] := 'Libellé';
G_NLIG.Cells[2, 0] := 'Quantité';
G_NLIG.Titres.Text := 'Nomenclature';
G_NLIG.Titres.Add('Libellé');
G_NLIG.Titres.Add('Quantité');
G_NLIG.ColAligns[2] := taRightJustify;
HMTrad.ResizeGridColumns (G_NLIG) ;
//Old_CODEARTICLE := GA_CODEARTICLE.Text;
if GA_CODEARTICLE.Text <> '' then GA_CODEARTICLEExit(Sender);
Cancel := False;
G_NLIGRowEnter(Sender, 1, Cancel, False);
Row_Cur := 1;
Col_Cur := 1;
end;

//=============================================================================
//  Recherche d'une ligne de la grille en fonction d'une valeur chaine de caracteres
//=============================================================================
function TFCasEmplois.ChercheLigne (Valeur : string) : integer;
var
    i_ind1 : integer ;
    st1, st2 : string;
begin
Result := 0;
for i_ind1 := 0 to G_NLIG.RowCount - 1 do
    begin
    st1 := G_NLIG.Cells[0, i_ind1];
    st2 := Trim(Copy(Valeur, 0, Pos('(', Valeur) - 1));
    if Pos(st2, st1) <> 0 then
        begin
        Result := i_ind1;
        Exit;
        end;
    end;
end;

//=============================================================================
//  gestion des evenements sur la forme
//=============================================================================
procedure TFCasEmplois.bPrintClick(Sender: TObject);
Var TL : TList ;
    TT : TStrings ;
    SQL : String ;
begin
TL:=TList.Create ;
TT:=TStringList.Create ;
SQL:='SELECT * FROM NOMENLIG ' +
     'LEFT JOIN ARTICLE ON GNL_ARTICLE=GA_ARTICLE ' +
     'WHERE GNL_ARTICLE="' + GA_ARTICLE.Text + '" ';
TT.Add(SQL) ; TL.Add(TT) ;
LanceDocument('E','ART','NCA',TL,Nil,True,False) ;
TT.Free ; TL.Free ;
end;

procedure TFCasEmplois.G_NLIGRowEnter(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
begin
//--------------------------------------------------------------------------
//  Verification qu'on n'entre pas sur une ligne vide au dela de la premiere possible
//--------------------------------------------------------------------------
if Ou > TobNomen.detail.count + 1 then
    begin
    G_NLIG.Row := TobNomen.detail.count + 1;
    G_NLIG.Col := 1;
    Ou := Min(Ou, G_NLIG.Row);
    end;
//--------------------------------------------------------------------------
//  affichage synchrone du treeview avec la grid
//--------------------------------------------------------------------------
TV_NLIG.HideSelection := False;
//TV_NLIG.Selected := ChercheNode(TV_NLIG, nil, G_NLIG.Cells[0, Ou]); // DBR Fiche 10995
TV_NLIG.Selected := ChercheNode(TV_NLIG, nil, ou + 1); // DBR Fiche 10995
//--------------------------------------------------------------------------
//  Mise à jour du Hint
//--------------------------------------------------------------------------
G_NLIG.Hint := G_NLIG.Cells[1, G_NLIG.Row];
//--------------------------------------------------------------------------
//  gestion des options popup
//--------------------------------------------------------------------------
if Nomen_Open then
    PopG_NLIG_N.Enabled := True
    else
    PopG_NLIG_N.Enabled := False;
end;

procedure TFCasEmplois.GA_CODEARTICLEElipsisClick(Sender: TObject);
var
    Q : TQuery;
begin
GA_CODEARTICLE.DataType := 'GCARTICLE';
DispatchRecherche (GA_CODEARTICLE, 1, '',
                   'GA_CODEARTICLE=' + Trim (Copy (GA_CODEARTICLE.Text, 1, 18)), '');
if GA_CODEARTICLE.Text <> '' then
    begin
    Q := OpenSQL('Select * from ARTICLE Where GA_ARTICLE="' + GA_CODEARTICLE.Text + '"', True,-1,'',true);
    if Not Q.EOF then
        begin
        TOBArt.SelectDB ('', Q);
        GA_CODEARTICLE.Text := TOBArt.GetValue('GA_CODEARTICLE');
        GA_ARTICLE.Text := TOBArt.GetValue('GA_ARTICLE');
        GA_CODEARTICLEExit(Sender);
        end;
    Ferme(Q);
    end;
end;

procedure TFCasEmplois.GA_CODEARTICLEExit(Sender: TObject);
var RechArt : T_RechArt ;
    Q : TQuery;
    TF1 : TField;
    OkArt   : Boolean ;
    i_ind1 : integer;
begin
// si le code article n'a pas changé, on sort
if Old_CODEARTICLE = GA_CODEARTICLE.Text then Exit;
//-----------------------------------------------------------------
//  Recherche du code article saisi
//-----------------------------------------------------------------
if Old_CODEARTICLE <> '' then
    begin
    OkArt:=False ;
    RechArt := TrouverArticle (GA_CODEARTICLE, TOBArt);
    case RechArt of
           traOk : begin
    //-----------------------------------------------------------------
    //  Code article trouve, on regarde s'il a des nomenclatures associées
    //  si oui et plusieurs, ouverture d'une fenetre de selection
    //  si oui et une seule, on recupere directement son code
    //-----------------------------------------------------------------
                   Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_ARTICLE="' +
                                GA_CODEARTICLE.Text + '"', True);
                   if Q.RecordCount > 1 then
                       GNE_NOMENCLATURE.Text := AGLLanceFiche('GC','GCNOMEN_MUL','GNE_ARTICLE=' +
                                                              GA_CODEARTICLE.Text,'','')
                   else
                       if Q.RecordCount = 1 then
                           begin
                           TF1 := Q.FindField('GNE_NOMENCLATURE');
                           GNE_NOMENCLATURE.Text := TF1.AsString ;
                           end;
                   Ferme(Q);
                   OkArt:=True ;
                   end;
        traAucun : begin
    //-----------------------------------------------------------------
    // Recherche sur code via LookUp ou Recherche avancée
    //-----------------------------------------------------------------
                   DispatchRecherche (GA_CODEARTICLE, 1, '',
                                      'GA_CODEARTICLE=' + Trim (Copy (GA_CODEARTICLE.Text, 1, 18)), '');
                   if GA_CODEARTICLE.Text <> '' then OkArt := True;
                   end ;
       traGrille : begin
    //-----------------------------------------------------------------
    // Forcement objet dimension avec saisie obligatoire
    //-----------------------------------------------------------------
                   if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                       OkArt := True;
                   end;
    end ; // Case
    if (OkArt) then
        GA_ARTICLE.Text := TOBArt.GetValue('GA_ARTICLE');
    end;
//-----------------------------------------------------------------
// Charge les cas d'emploi de l'article choisi
//-----------------------------------------------------------------
CasEmploi(GA_CODEARTICLE.Text, '', True, TobNomen);
for i_ind1 := 1 to G_NLIG.RowCount - 1 do
    G_NLIG.Rows[i_ind1].Clear;
AffecteLibelle(TobNomen, 'GNL_NOMENCLATURE');
TobNomen.PutGridDetail(G_NLIG, False, False, 'GNL_SOUSNOMEN;GNL_LIBELLE;GNL_QTE');
AfficheTreeView(TV_NLIG, TobNomen, 'LIBELLE');
TV_NLIG.TopItem.Text := GA_CODEARTICLE.Text;
TV_NLIG.FullExpand;
Old_CODEARTICLE := GA_CODEARTICLE.Text;
end;

procedure TFCasEmplois.PopG_NLIG_AClick(Sender: TObject);
var
    st1 : string;
begin
st1 := CodeArticleUnique2(G_NLIG.Cells[0, G_NLIG.Row], '');
AGLLanceFiche('GC','GCARTICLE', '', st1, 'ACTION=CONSULTATION')
end;

procedure TFCasEmplois.PopG_NLIG_NClick(Sender: TObject);
var st_Article, st_Nomen : string;
begin
TFNomenLig.Create(Application);
st_Article := TOB(G_NLIG.Objects[0, G_NLIG.Row]).GetValue('GNL_ARTICLE');
st_Nomen := TOB(G_NLIG.Objects[0, G_NLIG.Row]).GetValue('GNL_SOUSNOMEN');
Entree_NomenLig(['CASEMPLOIS=NON', st_Article, st_Nomen], 3);
end;

procedure TFCasEmplois.PopTV_NLIG_NClick(Sender: TObject);
begin
PopG_NLIG_NClick(Sender);
end;

procedure TFCasEmplois.PopTV_NLIG_FClick(Sender: TObject);
begin
TV_NLIG.FullCollapse;
TV_NLIG.TopItem.Expand(False);
end;

procedure TFCasEmplois.PopTV_NLIG_OClick(Sender: TObject);
begin
TV_NLIG.FullExpand;
end;

//=============================================================================
//  Evenements du treeview
//=============================================================================
procedure TFCasEmplois.TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
                                      Shift: TShiftState; X, Y: Integer);
var
    tn1 : TTreeNode;
    i_ind1 : integer;
begin
case Button of
    mbLeft  : begin
              tn1 := TV_NLIG.GetNodeAt(X, Y);
              if tn1 = nil then Exit;
//--------------------------------------------------------------------------
//  le noeud selectionné n'est pas de niveau 1 donc pas accessible directement
//  on positionne sur le parent de niveau 1
//--------------------------------------------------------------------------
              if tn1.Level > 1 then
                  begin
                  tn1 := FindParentNiv1(tn1);
                  TV_NLIG.Selected := tn1;
                  end;
//--------------------------------------------------------------------------
//  affichage synchrone de la grid avec le treeview
//--------------------------------------------------------------------------
              i_ind1 := ChercheLigne(tn1.Text);
              if i_ind1 <> 0 then
                  begin
                  G_NLIG.Col := 0;
                  G_NLIG.Row := i_ind1;
                  G_NLIG.SetFocus;
                  Exit;
                  end;
              end;
    mbRight : begin
              end;
    end;
end;

/////////////////////////////////////////////////////////////////////////////
procedure InitCasEmploi();
begin
 RegisterAglProc( 'Entree_CasEmplois', False , 2, Entree_CasEmploi);
end;

procedure TFCasEmplois.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

Initialization
InitCasEmploi();

end.
