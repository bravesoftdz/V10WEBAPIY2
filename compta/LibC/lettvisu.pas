unit LettVisu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, ExtCtrls, Grids, Buttons, SaisUtil, LettUtil, SaisComm,
  Ent1, DBCtrls,
{$IFNDEF IMP}
  Saisie,
  SaisBor,
{$ENDIF}
{$IFDEF VER150}
   Variants,
 {$ENDIF}
  hmsgbox, Menus, Filtre, HSysMenu, HTB97,
  UtilSais,
  LettAuto, ComCtrls ;

type
  TFLettVisu = class(TForm)
    POutils: TPanel;
    BValideSelect: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    PEntete: TPanel;
    H_COMPTE: THLabel;
    GD: THGrid;
    PSepare: TPanel;
    GC: THGrid;
    H_NBSOL: TLabel;
    H_NUMSOL: TLabel;
    CJOURNAL: THValComboBox;
    CNATUREPIECE: THValComboBox;
    CMODEPAIE: THValComboBox;
    CDevise: THValComboBox;
    E_NBSOL: TLabel;
    E_NUMSOL: TLabel;
    BZoomPiece: TToolbarButton97;
    HMessVisuL: THMsgBox;
    POPS: TPopupMenu;
    HMTrad: THSystemMenu;
    PCDetail: TPageControl;
    TSDetail: TTabSheet;
    TSDetail1: TTabSheet;
    H_JOURNAL: TLabel;
    E_JOURNAL: TLabel;
    H_NUMEROPIECE: TLabel;
    E_NUMEROPIECE: TLabel;
    H_DATEECHEANCE: TLabel;
    E_DATEECHEANCE: TLabel;
    H_REFLIBRE: TLabel;
    E_REFLIBRE: TLabel;
    H_NATUREPIECE: TLabel;
    E_NATUREPIECE: TLabel;
    E_LIBELLE: TLabel;
    H_LIBELLE: TLabel;
    H_MODEPAIE: TLabel;
    E_MODEPAIE: TLabel;
    E_TOTAL: THNumEdit;
    H_TOTAL: TLabel;
    BSepare: TBevel;
    H_JOURNAL_: TLabel;
    H_NUMEROPIECE_: TLabel;
    H_DATEECHEANCE_: TLabel;
    H_REFLIBRE_: TLabel;
    E_REFLIBRE_: TLabel;
    E_DATEECHEANCE_: TLabel;
    E_NUMEROPIECE_: TLabel;
    E_JOURNAL_: TLabel;
    H_NATUREPIECE_: TLabel;
    H_LIBELLE_: TLabel;
    H_MODEPAIE_: TLabel;
    H_TOTAL_: TLabel;
    E_TOTAL_: THNumEdit;
    E_MODEPAIE_: TLabel;
    E_LIBELLE_: TLabel;
    E_NATUREPIECE_: TLabel;
    Label1: TLabel;
    e_naturepcl: TLabel;
    Label3: TLabel;
    e_refrelevepcl: TLabel;
    Label7: TLabel;
    e_affairepcl: TLabel;
    Label2: TLabel;
    e_dateecheancepcl: TLabel;
    Label4: TLabel;
    e_reflibrepcl: TLabel;
    Label8: TLabel;
    e_refexternepcl: TLabel;
    Label5: TLabel;
    E_CONTREPARTIEGENPCL: TLabel;
    E_ETABLISSEMENTPCL: TLabel;
    Label6: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BValideSelectClick(Sender: TObject);
    procedure MontreDetail(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure GDEnter(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    procedure DetailLigne ( Lig : integer ; Debit : boolean ) ;
    procedure SwapGrids ;
    procedure ActiveGrid ;
    procedure CentreGrids ;
    procedure Bouge(Button: TNavigateBtn) ;
    procedure OToO ( OSource,ODest : TOBM ) ;
    procedure PrepareSelect ( G : THGrid ) ;
    procedure RemplirGrids ;
    procedure ChargeSolution ;
    procedure ClickZoomPiece ;
    procedure ClickValide ;
  public
    //LG*
    BoUneGrille  : boolean;
    GLD,GLC      : THGrid ;
    TLAD,TLAC    : TList ;
    NbSol,CurSol : integer ;
    Client       : boolean ;
    DEV          : RDEVISE ;
    RL           : RLETTR ;
    Titre        : String ;
    CurGrid      : integer ;
  end;

function VisuSolutionsLettrage ( GD,GC : THGrid ; TLAD,TLAC : TList ; TEC : T_LETTECHANGE; UneGrille : boolean = false ) : integer ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  HEnt1;

{$R *.DFM}

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/11/2001
Modifié le ... :   /  /
Description .. : LG Modif : ajout d'un nouveau parametre FBoModePCL. Si
Suite ........ : FBoUneGrille:=true tous les mvt. sotn affichés dans la m. grille (
Suite ........ : celle des debits )
Mots clefs ... :
*****************************************************************}
function VisuSolutionsLettrage ( GD,GC : THGrid ; TLAD,TLAC : TList ; TEC : T_LETTECHANGE ; UneGrille : boolean = false) : integer ;
Var X  : TFLettVisu ;
    ii : integer ;
BEGIN
//LG*
Result:=-1 ;
X:=TFLettVisu.Create(Application) ;
 Try
  X.BoUneGrille := UneGrille;
  //LG*
  if X.BoUneGrille then X.GLC:=GD else X.GLC:=GC ;
  X.GLD:=GD ; X.TLAD:=TLAD ; X.TLAC:=TLAC ;
  X.NbSol:=TEC.NbSol ; X.Client:=TEC.Client ;
  X.DEV:=TEC.DEV ; X.RL:=TEC.RLL ; X.Titre:=TEC.Titre ;
  ii:=X.ShowModal ;
  if ii=mrYes then Result:=X.CurSol ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

{================================= VISU ======================================}
procedure TFLettVisu.Bouge(Button: TNavigateBtn) ;
BEGIN
Case Button of
   NbFirst : if CurSol=1 then Exit else CurSol:=1 ;
   NbLast  : if CurSol=NbSol then Exit else CurSol:=NbSol ;
   NbPrior : if CurSol=1 then Exit else Dec(CurSol) ;
   NbNext  : if CurSol=NbSol then Exit else Inc(CurSol) ;
   END ;
ChargeSolution ;
END ;

procedure TFLettVisu.ActiveGrid ;
BEGIN
if GD.Focused then BEGIN CurGrid:=1 ; GD.Font.Color:=clWindowText ; GC.Font.Color:=clGray ; END ;
if GC.Focused then BEGIN CurGrid:=2 ; GC.Font.Color:=clWindowText ; GD.Font.Color:=clGray ; END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Lauretn GENDREAU
Créé le ...... : 12/04/2002
Modifié le ... :   /  /    
Description .. : on ne peut pas changer de grille si on en a qu'une !
Mots clefs ... : 
*****************************************************************}
procedure TFLettVisu.SwapGrids ;
BEGIN
if BoUneGrille then exit ;
if GD.Focused then GC.SetFocus else if GC.Focused then GD.SetFocus ;
END ;

procedure TFLettVisu.DetailLigne ( Lig : integer ; Debit : boolean ) ;
Var i : integer ;
    O : TOBM ;
    V : Variant ;
    C : TControl ;
    Nom : String ;

 procedure AfficheMontant ( X,XD : double ; vDebit : boolean = true);
  begin
   if vDebit then
   BEGIN
    if RL.LettrageDevise then AfficheLeSolde(THNumEdit(C),X,0) else  // affiche un D dans la case
     BEGIN
     FormatLettrage(THNumEdit(C),CDevise,O.GetMvt('E_DEVISE')) ;
     AfficheLeSolde(THNumEdit(C),XD,0) ;
     END ;
   END else
   BEGIN
    if RL.LettrageDevise then AfficheLeSolde(THNumEdit(C),0,X) else // affiche un C dans la case
     BEGIN
     FormatLettrage(THNumEdit(C),CDevise,O.GetMvt('E_DEVISE')) ;
     AfficheLeSolde(THNumEdit(C),0,XD) ;
     END ;
   END;
  end;

BEGIN
if Debit then O:=GetO(GD,Lig) else O:=GetO(GC,Lig) ;
for i:=0 to TSDetail.ControlCount-1 do
    BEGIN
    C:=TSDetail.Controls[i] ; if C.Tag=0 then Continue ;
    if ((Debit) and (C.Name[Length(C.Name)]='_')) then Continue ;
    if ((Not Debit) and (C.Name[Length(C.Name)]<>'_')) then Continue ;
    if O=Nil then
       BEGIN
       if C is THNumEdit then
          BEGIN
          THNumEdit(C).Value:=0 ; FormatLettrage(THNumEdit(C),CDEvise,DEV.Code) ;
          END else if C is TLabel then TLabel(C).Caption:='' ;
       END else
       BEGIN
       Nom:=C.Name ; if Not Debit then Nom:=Copy(Nom,1,Length(Nom)-1) ; if Nom<>'E_TOTAL' then V:=O.GetMvt(Nom) ;
       if Nom='E_JOURNAL' then BEGIN CJOURNAL.Value:=V ; TLabel(C).Caption:=CJOURNAL.Text ; END else
       if Nom='E_NATUREPIECE' then BEGIN CNATUREPIECE.Value:=V ; TLabel(C).Caption:=CNATUREPIECE.Text ; END else
       if Nom='E_NUMEROPIECE' then TLabel(C).Caption:=VarAsType(V,VarString) else
       if Nom='E_LIBELLE' then TLabel(C).Caption:=VarAsType(V,VarString) else
       if Nom='E_DATEECHEANCE' then TLabel(C).Caption:=DateToStr(VarAsType(V,VarDate)) else
       if Nom='E_MODEPAIE' then BEGIN CMODEPAIE.Value:=V ; TLabel(C).Caption:=CMODEPAIE.Text ; END else
       if Nom='E_REFLIBRE' then TLabel(C).Caption:=VarAsType(V,VarString) else
       if Nom='E_TOTAL' then
          BEGIN
          //LG* modif pour la gestion d'une seule grille -> la case E_TOTAL affiche les debits te les credits
          if BoUneGrille then
             BEGIN
              if O.GetMvt('E_DEBIT') <> 0 then AfficheMontant(O.GetMvt('E_DEBIT'),O.GetMvt('E_DEBITDEV')) else AfficheMontant(O.GetMvt('E_CREDIT')*(-1),O.GetMvt('E_CREDITDEV'),false);
             END else
           if Debit then AfficheMontant(O.GetMvt('E_DEBIT'),O.GetMvt('E_DEBITDEV')) else AfficheMontant(O.GetMvt('E_CREDIT'),O.GetMvt('E_CREDITDEV'),false);
          END else TLabel(C).Caption:=VarAsType(V,VarString) ;
       END ;
    END ;
//LG* 26/12/2001 gestion des nouveaux controles
if BoUneGrille and (O<>nil) then
 BEGIN
  e_naturepcl.Caption:=RechDom('ttNaturePiece', O.GetMvt('E_NATUREPIECE') ,false) ; e_dateecheancepcl.Caption:=DateToStr(O.GetMvt('E_DATEECHEANCE')) ;
  e_refrelevepcl.Caption:=O.GetMvt('E_REFRELEVE') ; e_contrepartiegenpcl.Caption:=O.GetMvt('E_CONTREPARTIEGEN') ;
  e_etablissementpcl.Caption:=RechDom('ttEtablissement', O.GetMvt('E_ETABLISSEMENT') ,false) ; e_reflibrepcl.Caption:=O.GetMvt('E_REFLIBRE') ;
  e_refexternepcl.Caption:=O.GetMvt('E_REFEXTERNE') ; e_affairepcl.Caption:=O.GetMvt('E_AFFAIRE') ; e_dateecheancepcl.Caption:=O.GetMvt('E_DATEECHEANCE') ;
 END;
END ;

{================================= INITS, DEFAUTS =====================================}
procedure TFLettVisu.OToO ( OSource,ODest : TOBM ) ;
BEGIN
//LG* 25/03/2002 modif pour le TOBM en TOBM
 ODest.Dupliquer( OSource , true , true );
END ;

procedure TFLettVisu.PrepareSelect ( G : THGrid ) ;
BEGIN G.ColWidths[G.ColCount-1]:=0 ; END ;

procedure TFLettVisu.CentreGrids ;
BEGIN
//LG*
TSDetail.TabVisible:=false ; TSDetail1.TabVisible:=false ;
if BoUneGrille then
 BEGIN // positionnement des panels
  PCDetail.activepage:=TSDetail1 ; PCDetail.Height:=60 ;
 END
  else BEGIN PCDetail.activepage:=TSDetail ;  END ;

if BoUneGrille then exit;
GD.Width:=((PEntete.Width-1)-PSepare.Width) Div 2+1 ;
END ;

{================================== CHARGEMENTS ======================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 17/12/2001
Modifié le ... :   /  /    
Description .. : LG modif : utilisation de la fonction de CTOBVersTHGrid ( 
Suite ........ : les duex liste param sont differentes)
Mots clefs ... : 
*****************************************************************}
procedure TFLettVisu.RemplirGrids ;
Var i,Lig : integer ;
    OBM   : TOBM ;
    CD    : String3 ;
BEGIN
Lig:=0 ;
for i:=0 to TLAD.Count-1 do
    BEGIN
    if P_LAS(TLAD[i]).S[CurSol]<>slPasSolution then
       BEGIN
       inc(Lig) ; //GD.Rows[Lig].Assign(GLD.Rows[i+1]) ; LG* suppression de la ligne -> utilisation de la fct CTOBVersTHGrid
       OBM:=TOBM.Create(EcrGen,'',True) ; OToO(GetO(GLD,i+1),OBM) ;
       GD.Objects[GD.ColCount-1,GD.RowCount-1]:=OBM ;
       GD.Row:=Lig ;
       CTOBVersTHGrid(GD,OBM,RL.LettrageDevise);
       GD.RowCount:=GD.RowCount+1 ;
       if Not RL.LettrageDevise then BEGIN CD:=OBM.GetMvt('E_DEVISE') ; if CD<>V_PGI.DevisePivot then AjouteDevise(CDevise,CD) ; END ;
       END ;
    END ;
Lig:=0 ;

for i:=0 to TLAC.Count-1 do
    BEGIN
    if P_LAS(TLAC[i]).S[CurSol]<>slPasSolution then
       BEGIN
       inc(Lig) ; //GC.Rows[Lig].Assign(GLC.Rows[i+1]) ;
       OBM:=TOBM.Create(EcrGen,'',True) ; OToO(GetO(GLC,i+1),OBM) ;
       if BoUneGrille then
        BEGIN // en mode pcl une seule grille -> on met les credit dans GD et non pas GC
         GD.Objects[GD.ColCount-1,GC.RowCount-1]:=OBM ; GD.Row:=Lig ;
         CTOBVersTHGrid(GD,OBM,RL.LettrageDevise) ; GD.RowCount:=GD.RowCount+1 ;
         if Not RL.LettrageDevise then BEGIN CD:=OBM.GetMvt('E_DEVISE') ; if CD<>V_PGI.DevisePivot then AjouteDevise(CDevise,CD) ; END ;
        END
         else
        BEGIN
         GC.Objects[GC.ColCount-1,GC.RowCount-1]:=OBM ; GC.Row:=Lig ;
         CTOBVersTHGrid(GC,OBM,RL.LettrageDevise) ; GC.RowCount:=GC.RowCount+1 ;
         if Not RL.LettrageDevise then BEGIN CD:=OBM.GetMvt('E_DEVISE') ; if CD<>V_PGI.DevisePivot then AjouteDevise(CDevise,CD) ; END ;
        END;
       END ;
    END ;
END ;

procedure TFLettVisu.ChargeSolution ;
BEGIN
E_NumSol.Caption:=IntToStr(CurSol) ;
GC.VidePile(True) ; GD.VidePile(True) ; GC.RowCount:=2 ; GC.RowCount:=2 ;
RemplirGrids ;
GD.SetFocus ;
DetailLigne(1,True) ; DetailLigne(1,False) ;
END ;

{================================ BARRE OUTILS =======================================}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/03/2002
Modifié le ... :   /  /    
Description .. : //LG* 20/03/2002 plus de message de confirmation en 
Suite ........ : mode pcl
Mots clefs ... : 
*****************************************************************}
procedure TFLettVisu.ClickValide ;
BEGIN
if ctxPCL in V_PGI.PGIContexte then ModalResult:=mrYes
else if HMessVisuL.Execute(0,'','')=mrYes then BEGIN ModalResult:=mrYes ; END ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 18/08/2004
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour 
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices 
Suite ........ : decalees
Mots clefs ... : 
*****************************************************************}
{$IFNDEF EAGLCLIENT}
procedure TFLettVisu.ClickZoomPiece ;
{$IFNDEF IMP}
Var O  : TOBM ;
    M  : RMVT ;
    P  : RParFolio ;
{$ENDIF}
BEGIN
{$IFNDEF IMP}
if ((GD.Focused) or (CurGrid=1)) then O:=GetO(GD,GD.Row) else
 if ((GC.Focused) or (CurGrid=2)) then O:=GetO(GC,GC.Row) else Exit ;
if O=Nil then Exit ;
M:=OBMToIdent(O,False) ;
if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>'')) then
   BEGIN
   FillChar(P, Sizeof(P), #0) ;
   P.ParPeriode:=O.GetMvt('E_DATECOMPTABLE') ;
   P.ParCodeJal:=O.GetMvt('E_JOURNAL') ;
   P.ParNumFolio:=IntToStr(O.GetMvt('E_NUMEROPIECE')) ;
   P.ParNumLigne:=O.GetMvt('E_NUMLIGNE') ;
   ChargeSaisieFolio(P, taConsult) ;
   END else
   BEGIN
   LanceSaisie(Nil,taConsult,M) ;
   END ;
{$ENDIF}
END ;
{$ELSE}
procedure TFLettVisu.ClickZoomPiece ;
Var O  : TOBM ;
    M  : RMVT ;
BEGIN
{$IFNDEF IMP}
if ((GD.Focused) or (CurGrid=1)) then O:=GetO(GD,GD.Row) else
 if ((GC.Focused) or (CurGrid=2)) then O:=GetO(GC,GC.Row) else Exit ;
if O=Nil then Exit ;
M:=OBMToIdent(O,False) ;
if ((M.ModeSaisieJal='-') or (M.ModeSaisieJal='')) then
   BEGIN
    LanceSaisie(Nil,taConsult,M) ;
   END ;
{$ENDIF}
END ;
{$ENDIF}

procedure TFLettVisu.BValideSelectClick(Sender: TObject);
begin ClickValide ; end;

procedure TFLettVisu.BAbandonClick(Sender: TObject);
begin Close ; end;

procedure TFLettVisu.BFirstClick(Sender: TObject);
begin Bouge(NbFirst) ; end;

procedure TFLettVisu.BPrevClick(Sender: TObject);
begin Bouge(NbPrior) ; end;

procedure TFLettVisu.BNextClick(Sender: TObject);
begin Bouge(NbNext) ; end;

procedure TFLettVisu.BLastClick(Sender: TObject);
begin Bouge(NbLast) ; end;

procedure TFLettVisu.BZoomPieceClick(Sender: TObject);
begin ClickZoomPiece ; end;

{============================ METHODES DE LES GRIDS ==================================}
procedure TFLettVisu.MontreDetail(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var G : THGrid ;
begin
G:=THGrid(Sender) ; DetailLigne(Ou,G.Name='GD') ;
end;

procedure TFLettVisu.GDEnter(Sender: TObject);
begin ActiveGrid ; end;

{============================== METHODES DE LA FORM ==================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/11/2001
Modifié le ... : 29/11/2001
Description .. : LG Modif :
Suite ........ : - prise en compte de la nouvelle ListeParam LETTPCL
Suite ........ : - Gestion des panels details
Suite ........ : - gestion de l'affichage de la grille credit
Suite ........ : fonction deplacé du formcreate
Mots clefs ... : 
*****************************************************************}
procedure TFLettVisu.FormShow(Sender: TObject);
begin
//LG*
if BoUneGrille then BEGIN GD.ListeParam:='LETTVISU'; GC.ListeParam:='LETTVISU' ; END
else if Client then BEGIN GD.ListeParam:='LETCLIDEB' ; GC.ListeParam:='LETCLICRE' ; END
else BEGIN GD.ListeParam:='LETFOUDEB' ; GC.ListeParam:='LETFOUCRE' ; END ;
PrepareSelect(GD) ; PrepareSelect(GC) ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
if NbSol<=1 then
   BEGIN
   BFirst.Enabled:=False ; BPrev.Enabled:=False ;
   BNext.Enabled:=False ; BLast.Enabled:=False ;
   END ;
AjouteDevise(CDevise,V_PGI.DevisePivot) ; CentreGrids ;
if RL.LettrageDevise then
   BEGIN
   ChangeFormatDevise(Self,DEV.Decimale,DEV.Symbole) ;
   ChangeMask(E_TOTAL,V_PGI.OkDecV,V_PGI.SymbolePivot) ; ChangeMask(E_TOTAL_,V_PGI.OkDecV,V_PGI.SymbolePivot) ;
   END else
   BEGIN
   ChangeFormatPivot(Self) ;
   ChangeMask(E_TOTAL,DEV.Decimale,DEV.Symbole) ; ChangeMask(E_TOTAL_,DEV.Decimale,DEV.Symbole) ;
   END ;
H_COMPTE.Caption:=Titre ; E_NbSol.Caption:=IntToStr(NbSol) ;
CurSol:=1 ; ChargeSolution ;
CurGrid:=1 ; ActiveGrid ;
//LG*
GC.Visible:=not BoUneGrille ; PSepare.Visible:=not BoUneGrille;
if BoUneGrille then BEGIN GD.Align:=AlClient; HMTrad.ResizeGridColumns(GD) ; END ;
end;

procedure TFLettVisu.FormClose(Sender: TObject; var Action: TCloseAction);
begin
GC.VidePile(True) ; GD.VidePile(True) ; PurgePopup(POPS) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/11/2001
Modifié le ... :   /  /
Description .. : deplacement des fonctions des gestions des grilles vers le
Suite ........ : formshow
Mots clefs ... :
*****************************************************************}
procedure TFLettVisu.FormCreate(Sender: TObject);
begin
//LG*
GD.VidePile(True) ; GC.VidePile(True) ;
GD.RowCount:=2 ; GC.RowCount:=2 ; GD.TypeSais:=tsLettrage ; GC.TypeSais:=tsLettrage ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/12/2002
Modifié le ... :   /  /    
Description .. : -17/12/2002 - fiche 10748 - le F10 se repercutait dans les 
Suite ........ : fenetre suivante
Mots clefs ... : 
*****************************************************************}
procedure TFLettVisu.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide,OkG : boolean ;
begin
OkG:=((GD.Focused) or (GC.Focused)) ; Vide:=(Shift=[]) ;
Case Key of
   VK_TAB           : if OKG then SwapGrids ;
   VK_ESCAPE        : if Vide then Close ;
   VK_RIGHT,VK_LEFT : BEGIN Key:=0 ; SwapGrids ; END ;
   VK_F3            : if ((Shift=[]) and (BPrev.Enabled)) then Bouge(NbPrior) else
                       if ((Shift=[ssShift]) and (BFirst.Enabled)) then Bouge(NbFirst) ;
   VK_F4            : if ((Shift=[]) and (BNext.Enabled)) then Bouge(NbNext) else
                       if ((Shift=[ssShift]) and (BLast.Enabled)) then Bouge(NbLast) ;
   VK_F10           : begin Key:=0 ; ClickValide ; end ;
   END ;

end;

procedure TFLettVisu.POPSPopup(Sender: TObject);
begin
InitPopup(Self) ; 
end;

procedure TFLettVisu.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
