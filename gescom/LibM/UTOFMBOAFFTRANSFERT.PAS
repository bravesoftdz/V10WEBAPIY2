{***********UNITE*************************************************
Auteur  ...... : Julien D
Créé le ...... : 27/02/2001
Modifié le ... : 31/05/2002
Description .. : Source TOF de la FICHE : PR_AFFTRANSFERT ()
Mots clefs ... : TOF;PR_AFFTRANSFERT
*****************************************************************}
Unit UTOFMBOAFFTRANSFERT ;

Interface

uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
      db,dbTables,FE_main,Mul,
{$ENDIF}
      forms,sysutils,
      ComCtrls,Hpanel, Math
      ,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,M3FP,HTB97,Dialogs, AGLInitGC, ExtCtrls, Hqry,LicUtil
      ,HDimension,
      Facture,FactUtil,Menus,
      UTofOptionEdit,StockUtil,UtilDimArticle,HStatus;

Type
  TOF_MBOAFFTRANSFERT = Class (TOF)
  private
    Action : TActionFiche;
    LesColSel, LesColSai : array[1..2] of string ;
    Sel_Fixed,Sel_Art,Sel_CodeArt,Sel_Lib,Sel_QteAAffect,
    Sel_DepDest,Sel_QteAffect,Sel_DepLib,
    Sai_Fixed,Sai_DepDest,Sai_Art,Sai_CodeArt,
    Sai_StockIni,Sai_StockMini,Sai_StockMaxi,Sai_QteVte,Sai_StockALF,
    Sai_Manq,Sai_Surp,Sai_Prop,Sai_StockFin,Sai_DepLib,Sai_StockNet : integer;
    CodeSEL, CodeSAI : array of string;
    FFieldList,FTitre : string;
    mode, NbColSel, NbColSai,LigneSAI : integer;
    TOBPROP, TOBSel, TOBSai, TOBDimensions : TOB ;
    SPLITTER : TSplitter;
    G_SEL,G_SAI : THGRID;
    BINVERSE,BOUVRIR : TToolbarButton97;
    POPZ: TPopupMenu;
    mnart,mnetabl: TMenuItem;
    DepotEmetteur,CodePtrf : string;
    Valider : boolean;
    procedure CreerSplitter;
    procedure AffecterLesBoutons;
    procedure AffecterLesPopUp;
    procedure CreerLesColonnesDesGrids;
    procedure AffecterLesGrids;
    procedure InfoArticle;
    procedure OnResize(Sender: TObject);
    // Evenement lié aux Boutons
    procedure BINVERSEClick(Sender: TObject);
    procedure BOUVRIRClick(Sender: TObject);
    // Evenement lié aux PopUp
    procedure mnartClick(Sender: TObject);
    procedure mnetablClick(Sender: TObject);
    // Evenement du Grid
    procedure G_SELRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_SELRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_SELDessineTriangle (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure G_SAIElipsisClick(Sender: TObject);
    procedure G_SAIDblClick(Sender: TObject);
    procedure G_SAICellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_SAIRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_SAIRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_SAIDessineTriangle (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    // Manipulation du Grid
    procedure InitialiseGrille;
    procedure ChargeGrille;
    procedure AffectGrilleSEL(G : THGrid ; Entete,EntLib :boolean ;
                              CodeChp: array of string; CalculRowCount : Boolean);
    procedure AffectGrilleSAI(G : THGrid; Entete,EntLib :boolean;
                              CodeChp: array of string; CalculRowCount : Boolean);
    procedure VoirArticle;
    procedure VoirEtabliss;
    function  GetNumLigne(Depot : string) : integer;
    // Calcul
    function  CalculQteAAffecter(StArticle : string): double;
    function  CalculQteAffecter(DepotDest : string; var i_ind:integer): double;
    procedure TraiterProposition (ACol, ARow : integer);
    function  VerifDispo (TOBL : TOB; var QteDiff: Double) : Boolean;
    // Dimensions
    procedure ChargeTOBDimensions(DepotDest,ArticleGenerique : string; Mode:integer);
    procedure DimPourUnDepot(DepotDest,ArticleGenerique : string; TOBPropTemp : TOB);
    procedure TraiteLesDim();
    procedure TraiteLesGen(DepotDest,ArticleGenerique : string);
    // TOB
    function  GetTOBLigne (ARow : integer) : TOB ;
  public
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

// libellés des messages
const TexteMessage: array[1..1] of string 	= (
      {1}  'Le stock disponible est insuffisant.'
      );
const RepCodeEmett:string='...';

Implementation

procedure TOF_MBOAFFTRANSFERT.OnUpdate ;
var F : TFMul ;
    i_ind : integer ;
    QQ, QDepotEmetteur : TQuery ;
    StSQL : String;
begin
inherited;
InitialiseGrille;
F:=TFMul(Ecran);
if TOBSai<>Nil then begin TOBSai.Free; TOBSai:=Nil; end;
TOBSai := TOB.Create ('', Nil, -1) ;
if TOBPROP=Nil then
  begin
  TOBPROP := TOB.Create ('', Nil, -1) ;
  StSQL := 'SELECT GTL_ARTICLE,GTL_CODEARTICLE,GTL_CODEPTRF,GTL_DEPOTDEST,' +
  'GTL_MANQUANT,GTL_PROPOSITION,GTL_QTEMAX,GTL_QTEMIN,GTL_QTEVTE,' +
  'GTL_STATUTART,GTL_STOCKALF,GTL_STOCKINITIAL,GTL_STOCKMAXI,GTL_STOCKMINI,' +
  'GTL_SURPLUS,GTL_UTILISATEUR,' +
  '(GTL_STOCKINITIAL - GTL_PREPACLI - GTL_RESERVECLI + GTL_RESERVEFOU) AS GTL_STOCKNET ' +
  'FROM PROPTRANSFLIG WHERE GTL_CODEPTRF="'+ CodePtrf +'" ' +
  'ORDER BY GTL_CODEPTRF,GTL_DEPOTDEST,GTL_ARTICLE';
  QQ := OpenSQL(StSQL, True) ;
  if not QQ.EOF then TOBPROP.LoadDetailDB('PROPTRANSFLIG', '', '', QQ, FAlse) else TOBPROP := nil;
  Ferme (QQ) ;
  end;
F.Pages.Visible:= False;  F.PFiltres.Visible:= False;
F.BAgrandir.Visible := False; F.BReduire.Visible := False;
F.BParamListe.Visible := False;
if (TOBPROP = nil) or (TOBPROP.Detail.count = 0) then
    BEGIN
    G_SEL.VidePile(False); G_SAI.VidePile(False);
    G_SEl.Enabled:= False; G_SAI.Enabled:= False; BINVERSE.Enabled:=False;
    for i_ind:=0 to POPZ.Items.Count-1 do POPZ.Items[i_ind].Enabled:=False;
    exit;
    END else
    BEGIN
    if DepotEmetteur='' then
      begin
      QDepotEmetteur := OpenSQL('SELECT GTE_DEPOTEMET FROM PROPTRANSFENT Where GTE_CODEPTRF="'+TOBProp.detail[0].GetValue('GTL_CODEPTRF')+'"', True) ;
      if not QDepotEmetteur.Eof then DepotEmetteur:=QDepotEmetteur.FindField('GTE_DEPOTEMET').AsString
      else DepotEmetteur:='Inconnu';
      Ferme(QDepotEmetteur);
      end;
    G_SEl.Enabled:= True; G_SAI.Enabled:= True; BINVERSE.Enabled:=True;
    for i_ind:=0 to POPZ.Items.Count-1 do POPZ.Items[i_ind].Enabled:=True;
    ChargeGrille;
    END;
end ;

function TOF_MBOAFFTRANSFERT.CalculQteAAffecter(StArticle : string): double;
var TOBL : TOB;
begin
Result:=0;
TOBL:=TOBPROP.FindFirst(['GTL_ARTICLE','GTL_DEPOTDEST'],[StArticle,RepCodeEmett],False);
if TOBL<>nil then Result:= Valeur(TOBL.GetValue('GTL_SURPLUS'))-Valeur(TOBL.GetValue('GTL_PROPOSITION'));
end;

function TOF_MBOAFFTRANSFERT.CalculQteAffecter(DepotDest : string; var i_ind:integer): double;
var Qte : Double;
    StatutArt : String;
begin
Qte:=0;
while (i_ind<TOBPROP.Detail.Count) and
      (DepotDest=TOBPROP.Detail[i_ind].GetValue('GTL_DEPOTDEST')) do
  begin
  StatutArt:=TOBPROP.Detail[i_ind].GetValue('GTL_STATUTART');
  if (StatutArt='GEN') OR (StatutArt='UNI') then Qte:=Qte+Valeur(TOBPROP.Detail[i_ind].GetValue('GTL_PROPOSITION'));
  i_ind:=i_ind+1; MoveCur(False);
  end;
result:=Qte;
end;

procedure TOF_MBOAFFTRANSFERT.CreerSplitter;
begin
SPLITTER:=TSplitter.Create(THPanel(GetControl('PSUPPORT')));
SPLITTER.Name:='Splitter';
SPLITTER.Parent:=THPanel(GetControl('PSUPPORT')) ;
SPLITTER.AutoSnap:=False;
SPLITTER.Beveled:=True;
SPLITTER.ResizeStyle:=rsUpdate;
SPLITTER.Left:=THGRID(GetControl('G_SEL')).Left ;
SPLITTER.Align:=THGRID(GetControl('G_SEL')).Align;
SPLITTER.Width:=THGRID(GetControl('G_SEL')).Width;
SPLITTER.Height:=3;
SPLITTER.Cursor:=crVSplit;
SPLITTER.Color:=clActiveCaption;
end ;

procedure TOF_MBOAFFTRANSFERT.AffecterLesBoutons;
begin
BINVERSE:=TToolbarButton97(GetControl('BINVERSE'));
BINVERSE.OnClick:=BINVERSECLick;
BOUVRIR:=TToolbarButton97(GetControl('BOUVRIR'));
BOUVRIR.OnClick:=BOUVRIRCLick;
end;

procedure TOF_MBOAFFTRANSFERT.AffecterLesPopUp;
var i_mnart,i_mnetabl : integer;
begin
POPZ:=TPopupMenu(GetControl('POPZ'));
i_mnart:=-1; i_mnetabl:=-1 ;
if POPZ.Items[0].Name='mnart' then i_mnart:=0
    else if POPZ.Items[1].Name='mnart' then i_mnart:=1;
if POPZ.Items[0].Name='mnetabl' then i_mnetabl:=0
    else if POPZ.Items[1].Name='mnetabl' then i_mnetabl:=1;
if i_mnart>=0 then
    begin
    mnart:=POPZ.Items[i_mnart];
    mnart.OnClick:=mnartClick;
    end;
if i_mnetabl>=0 then
    begin
    mnetabl:=POPZ.Items[i_mnetabl];
    mnetabl.OnClick:=mnetablClick;
    end;
end;

procedure TOF_MBOAFFTRANSFERT.CreerLesColonnesDesGrids;
begin
LesColSel[1]:='FIXED;GTL_ARTICLE;GTL_CODEARTICLE;LIBELLE;QTEAAFFECTER';
LesColSel[2]:='FIXED;GTL_DEPOTDEST;DEPOTLIBELLE;QTEAFFECTER';
LesColSai[1]:='FIXED;GTL_DEPOTDEST;DEPOTLIBELLE;GTL_STOCKINITIAL;GTL_STOCKNET;GTL_STOCKMINI;GTL_STOCKMAXI;GTL_QTEVTE;GTL_STOCKALF;GTL_MANQUANT;GTL_SURPLUS;GTL_PROPOSITION;STOCKFINAL';
LesColSai[2]:='FIXED;GTL_ARTICLE;GTL_CODEARTICLE;GTL_STOCKINITIAL;GTL_STOCKNET;GTL_STOCKMINI;GTL_STOCKMAXI;GTL_QTEVTE;GTL_STOCKALF;GTL_MANQUANT;GTL_SURPLUS;GTL_PROPOSITION;STOCKFINAL';
end;

procedure TOF_MBOAFFTRANSFERT.AffecterLesGrids;
begin
G_SEL:=THGRID(GetControl('G_SEL'));
G_SEL.OnRowEnter:=G_SELRowEnter ;
G_SEL.OnRowExit:=G_SELRowExit ;
G_SEL.PostDrawCell:= G_SELDessineTriangle;

G_SAI:=THGRID(GetControl('G_SAI'));
G_SAI.OnElipsisClick:=G_SAIElipsisClick  ;
G_SAI.OnDblClick:=G_SAIDblClick ;
G_SAI.OnCellExit:=G_SAICellExit ;
G_SAI.OnRowEnter:=G_SAIRowEnter ;
G_SAI.OnRowExit:=G_SAIRowExit ;
G_SAI.PostDrawCell:= G_SAIDessineTriangle;
end;

procedure TOF_MBOAFFTRANSFERT.OnArgument (S : String ) ;
var Critere,ChampMul,ValMul : string;
    x : integer;
begin
inherited ;
mode:=1;
Repeat
  Critere:=uppercase(Trim(ReadTokenSt(S))) ;
  if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         if ChampMul='ACTION' then
           begin if ValMul='MODIFICATION' then Action := taModif else Action := taConsult; end
         else if ChampMul='CODEDEPOTEMET' then DepotEmetteur:=ValMul
         else if ChampMul='CODEPTRF' then CodePtrf:=ValMul
         ;
         end;
      end;
until  Critere='';
if LaTOB<>nil then
  begin
  TOBPROP := TOB.Create('', Nil, -1);
  TOBPROP.Dupliquer(LaTOB, True, True);
  LaTOB.Free;
  LaTOB:=nil;
  end;
CreerSplitter;
AffecterLesBoutons;
AffecterLesPopUp;
CreerLesColonnesDesGrids;
AffecterLesGrids;
TFMul(Ecran).OnResize := OnResize;
end;

procedure TOF_MBOAFFTRANSFERT.OnResize(Sender: TObject);
begin
  G_SEL.Height := TFMul(Ecran).Height div 3;
  TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SEL) ;
  TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SAI) ;
end;

procedure TOF_MBOAFFTRANSFERT.BINVERSEClick(Sender: TObject);
var i_ind : integer;
    FF : string ;
begin
if mode=1 then mode:=2 else mode:=1 ;
LigneSAI:=0;
for i_ind:=0 to G_SEL.ColCount do
    BEGIN
    FF:=G_SEL.ColFormats[i_ind];
    FillChar(FF, SizeOf(G_SEL.ColFormats[i_ind]),#0);
    G_SEL.ColFormats[i_ind]:=FF;
    G_SEL.ColLengths[i_ind]:=0;
    END;
for i_ind:=0 to G_SAI.ColCount do
    BEGIN
    FF:=G_SAI.ColFormats[i_ind];
    FillChar(FF, SizeOf(G_SAI.ColFormats[i_ind]),#0);
    G_SAI.ColFormats[i_ind]:=FF;
    G_SAI.ColLengths[i_ind]:=0;
    END;
G_SEL.VidePile(False);
G_SAI.VidePile(False);
InitialiseGrille;
ChargeGrille;
TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SEL); TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SAI) ;
end;

procedure TOF_MBOAFFTRANSFERT.mnartClick(Sender: TObject);
begin
VoirArticle;
end;

procedure TOF_MBOAFFTRANSFERT.mnetablClick(Sender: TObject);
begin
VoirEtabliss;
end;

procedure TOF_MBOAFFTRANSFERT.G_SELRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ACol : integer;
    Desc : Boolean;
begin
if LigneSAI<>Ou then
  begin
  LigneSAI:=Ou;
  G_SAI.VidePile(True);
  G_SEL.InvalidateRow(ou) ;
  AffectGrilleSAI(G_SAI,False,False,CodeSEL,True);
  ACol:=G_SAI.SortedCol ;
  Desc:=False;
  if ACol <> -1 then G_Sai.SortGrid(ACol,Desc);
  if Mode=1 then InfoArticle;
  end;
end;

procedure TOF_MBOAFFTRANSFERT.G_SELRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
G_SEL.InvalidateRow(ou) ;
end;

procedure TOF_MBOAFFTRANSFERT.G_SAIElipsisClick(Sender: TObject);
begin

end;

procedure TOF_MBOAFFTRANSFERT.ChargeTOBDimensions(DepotDest,ArticleGenerique:string; Mode:integer);
var TOBTemp : TOB;
    DepotDestinataire : String;
    i : integer;
begin
if TOBDimensions<>nil then
  begin TOBDimensions.free; TOBDimensions:=Nil; end;
TOBDimensions:=TOB.Create ('', Nil, -1);
DepotDestinataire:=DepotDest;

TOBTemp:= TOB.Create ('', Nil, -1) ;
TOBTemp.Dupliquer(TOBPROP, True, True);

// Chargement dans le cas du multi-dépôts
if Mode=1 then
  begin
  For i:=1 to G_SAI.RowCount-1 do
    begin
    DepotDestinataire:= G_SAI.Cells[Sai_DepDest,i];
    DimPourUnDepot(DepotDestinataire,ArticleGenerique,TOBTemp);
    end;
  end
else DimPourUnDepot(DepotDestinataire,ArticleGenerique,TOBTemp);

TOBTemp.Free;
end;

procedure TOF_MBOAFFTRANSFERT.DimPourUnDepot(DepotDest,ArticleGenerique : string; TOBPropTemp : TOB);
var TOBUneDim,TOBDepotEmetteur,TOBD : TOB;
    Etablissement : string;
begin
TOBUneDim:=TOBProp.FindFirst(['GTL_DEPOTDEST','GTL_CODEARTICLE','GTL_STATUTART'],[DepotDest,ArticleGenerique,'DIM'],False);
While TOBUneDim <> nil do
  begin
  TOBD:=TOB.Create ('', TOBDimensions,-1) ;
  TOBD.Dupliquer(TOBUneDim, False, True);
  TOBDepotEmetteur:=TOBPropTemp.FindFirst(['GTL_DEPOTDEST','GTL_ARTICLE'],[RepCodeEmett,TOBD.GetValue('GTL_ARTICLE')],False);
  if TOBDepotEmetteur<>nil then
    begin
    TOBD.AddChampSup('STOCKFINAL', False);
    TOBD.PutValue('STOCKFINAL',Valeur(TOBD.GetValue('GTL_STOCKINITIAL'))
                  +Valeur(TOBD.GetValue('GTL_PROPOSITION')));
                  //+Valeur(TOBD.GetValue('GTL_STOCKALF'))); //MODIF 01/07/2003
    if DepotDest='...' then Etablissement:=DepotEmetteur else Etablissement:=DepotDest;
    TOBD.AddChampSup('GDE_DEPOT', False); TOBD.PutValue('GDE_DEPOT',Etablissement);

    TOBD.AddChampSup('DEP_STOCKINITIAL', False); TOBD.PutValue('DEP_STOCKINITIAL',TOBDepotEmetteur.GetValue('GTL_STOCKINITIAL'));
    TOBD.AddChampSup('DEP_STOCKMINI', False); TOBD.PutValue('DEP_STOCKMINI',TOBDepotEmetteur.GetValue('GTL_STOCKMINI'));
    TOBD.AddChampSup('DEP_STOCKMAXI', False); TOBD.PutValue('DEP_STOCKMAXI',TOBDepotEmetteur.GetValue('GTL_STOCKMAXI'));
    TOBD.AddChampSup('DEP_MANQUANT', False); TOBD.PutValue('DEP_MANQUANT',TOBDepotEmetteur.GetValue('GTL_MANQUANT'));
    TOBD.AddChampSup('DEP_SURPLUS', False); TOBD.PutValue('DEP_SURPLUS',TOBDepotEmetteur.GetValue('GTL_SURPLUS'));
    TOBD.AddChampSup('DEP_PROPOSITION', False); TOBD.PutValue('DEP_PROPOSITION',TOBDepotEmetteur.GetValue('GTL_PROPOSITION'));
    TOBD.AddChampSup('DEP_STOCKFINAL', False);
    TOBD.PutValue('DEP_STOCKFINAL',Valeur(TOBDepotEmetteur.GetValue('GTL_STOCKINITIAL'))
                                  +Valeur(TOBDepotEmetteur.GetValue('GTL_PROPOSITION')));
                                  //+Valeur(TOBDepotEmetteur.GetValue('GTL_STOCKALF')));//MODIF 01/07/2003
    TOBD.AddChampSup('DEP_QTEMIN', False); TOBD.PutValue('DEP_QTEMIN',TOBDepotEmetteur.GetValue('GTL_QTEMIN'));
    TOBD.AddChampSup('DEP_QTEMAX', False); TOBD.PutValue('DEP_QTEMAX',TOBDepotEmetteur.GetValue('GTL_QTEMAX'));
    TOBD.AddChampSup('DEP_QTEVTE', False); TOBD.PutValue('DEP_QTEVTE',TOBDepotEmetteur.GetValue('GTL_QTEVTE'));
    //Modif 30/05/2002
    TOBD.AddChampSup('DEP_STOCKALF', False); TOBD.PutValue('DEP_STOCKALF',TOBDepotEmetteur.GetValue('GTL_STOCKALF'));
    end;
  TOBUneDim:= TOBProp.FindNext(['GTL_DEPOTDEST','GTL_CODEARTICLE','GTL_STATUTART'],[DepotDest,ArticleGenerique,'DIM'],False);
  end;
end;

procedure TOF_MBOAFFTRANSFERT.TraiteLesDim();
var i : integer;
    TOBEmet,TOBRecept : TOB;
    CodeArt,DepotDest : string;
begin
For i:=0 to TOBDimensions.Detail.Count-1 do
  begin
  if TOBDimensions.Detail[i].IsOneModifie then
    begin
    DepotDest := TOBDimensions.Detail[i].GetValue('GTL_DEPOTDEST');
    CodeArt := TOBDimensions.Detail[i].GetValue('GTL_ARTICLE');
    TOBEmet:=TOBPROP.FindFirst(['GTL_DEPOTDEST','GTL_ARTICLE'],['...',CodeArt],False);
    if TOBEmet<>nil then
      begin
      TOBEmet.PutValue('GTL_PROPOSITION',TOBDimensions.Detail[i].GetValue('DEP_PROPOSITION'));
      TOBEmet.PutValue('GTL_MANQUANT',TOBDimensions.Detail[i].GetValue('DEP_MANQUANT'));
      TOBEmet.PutValue('GTL_SURPLUS',TOBDimensions.Detail[i].GetValue('DEP_SURPLUS'));
      end;
    TOBRecept:=TOBPROP.FindFirst(['GTL_DEPOTDEST','GTL_ARTICLE'],[DepotDest,CodeArt],False);
    if TOBRecept<>nil then
      begin
      TOBRecept.PutValue('GTL_PROPOSITION',TOBDimensions.Detail[i].GetValue('GTL_PROPOSITION'));
      TOBRecept.PutValue('GTL_MANQUANT',TOBDimensions.Detail[i].GetValue('GTL_MANQUANT'));
      end;
    end;
  end;
end;

procedure TOF_MBOAFFTRANSFERT.TraiteLesGen(DepotDest,ArticleGenerique : string);
var TOBDim,TOBGen,TOBGenEmet,TOBSai : TOB;
    SumProp,SumManq,DiffQte : double;
    LigneDest : integer;
begin
SumProp:=0; SumManq:=0; DiffQte:=0;
//Cumul des articles dimensionnés pour mettre à jour la ligne générique
TOBDim:=TOBProp.FindFirst(['GTL_DEPOTDEST','GTL_CODEARTICLE','GTL_STATUTART'],[DepotDest,ArticleGenerique,'DIM'],False);
While TOBDim<>nil do
  begin
  SumProp:=SumProp+TOBDim.GetValue('GTL_PROPOSITION');
  SumManq:=SumManq+TOBDim.GetValue('GTL_MANQUANT');
  TOBDim:= TOBProp.FindNext(['GTL_DEPOTDEST','GTL_CODEARTICLE','GTL_STATUTART'],[DepotDest,ArticleGenerique,'DIM'],False);
  end;
//MAJ de la ligne générique
TOBGen:=TOBProp.FindFirst(['GTL_DEPOTDEST','GTL_CODEARTICLE','GTL_STATUTART'],[DepotDest,ArticleGenerique,'GEN'],False);
if TOBGen<>nil then
  begin
  DiffQte:=SumProp - TOBGen.GetValue('GTL_PROPOSITION');
  TOBGen.PutValue('GTL_PROPOSITION',SumProp);
  TOBGen.PutValue('GTL_MANQUANT',SumManq);
  end;
if Mode=1 then
     LigneDest:=GetNumLigne(DepotDest)
else LigneDest:=G_SAI.Row;
//MAJ de l'affichage de la ligne générique récépteur
G_SAI.Cells[Sai_Prop,LigneDest]:=FormatFloat(G_SAI.ColFormats[Sai_Prop],SumProp);
G_SAI.Cells[Sai_Manq,LigneDest]:=FormatFloat(G_SAI.ColFormats[Sai_Manq],SumManq);
G_SAI.Cells[Sai_StockFin,LigneDest]:=FormatFloat(G_SAI.ColFormats[Sai_StockFin],Valeur(G_SAI.Cells[Sai_StockIni,LigneDest])+SumProp);
TOBSai:=GetTOBLigne(LigneDest);
if TOBSai<>nil then TOBSai.PutValue('GTL_PROPOSITION',SumProp);
//MAJ de la ligne générique du dépôt émetteur
TOBGenEmet:=TOBProp.FindFirst(['GTL_DEPOTDEST','GTL_CODEARTICLE','GTL_STATUTART'],[RepCodeEmett,ArticleGenerique,'GEN'],False);
if TOBGenEmet<>nil then
  begin
  TOBGenEmet.PutValue('GTL_PROPOSITION',TOBGenEmet.GetValue('GTL_PROPOSITION')-DiffQte);
  TOBGenEmet.PutValue('GTL_MANQUANT',TOBGenEmet.GetValue('GTL_MANQUANT')-DiffQte);
  TOBGenEmet.PutValue('GTL_SURPLUS',TOBGenEmet.GetValue('GTL_SURPLUS')-DiffQte);
  end;
if Mode=1 then
  begin
  G_SAI.Cells[Sai_Prop,1]:=FormatFloat(G_SAI.ColFormats[Sai_Prop],TOBGenEmet.GetValue('GTL_PROPOSITION'));
  G_SAI.Cells[Sai_Manq,1]:=FormatFloat(G_SAI.ColFormats[Sai_Manq],TOBGenEmet.GetValue('GTL_MANQUANT'));
  G_SAI.Cells[Sai_Surp,1]:=FormatFloat(G_SAI.ColFormats[Sai_Surp],TOBGenEmet.GetValue('GTL_SURPLUS'));
  G_SAI.Cells[Sai_StockFin,1]:=FormatFloat(G_SAI.ColFormats[Sai_StockFin],Valeur(G_SAI.Cells[Sai_StockIni,1])+TOBGenEmet.GetValue('GTL_PROPOSITION'));
  TOBSai:=GetTOBLigne(1);
  if TOBSai<>nil then TOBSai.PutValue('GTL_PROPOSITION',TOBGenEmet.GetValue('GTL_PROPOSITION'));
  end
else
  begin
  G_SEL.Cells[Sel_QteAffect,G_SEL.Row]:=FormatFloat(G_SEL.ColFormats[Sel_QteAffect],Valeur(G_SEL.Cells[Sel_QteAffect,G_SEL.Row])+DiffQte);
  G_SEL.Cells[Sel_QteAffect,1]:=FormatFloat(G_SEL.ColFormats[Sel_QteAffect],Valeur(G_SEL.Cells[Sel_QteAffect,1])-DiffQte);
  end;
end;

function TOF_MBOAFFTRANSFERT.GetNumLigne(Depot : string) : integer;
var i : integer;
begin
Result:=1 ;
for i:=1 to G_SAI.RowCount-1 do
  begin
  if G_SAI.Cells[Sai_DepDest,i]=Depot then begin Result:=i; break; end;
  end;
end;

function TOF_MBOAFFTRANSFERT.GetTOBLigne (ARow : integer) : TOB ;
begin
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBSAI.Detail.Count)) then Exit ;
Result:=TOB(G_SAI.OBjects[0,ARow]);
end;

procedure TOF_MBOAFFTRANSFERT.VoirArticle;
var TOBL : TOB;
    RefArt : string;
BEGIN
TOBL := GetTOBLigne (G_SAI.Row); if TOBL=nil then exit;
RefArt:=TOBL.GetValue('GTL_ARTICLE');
AglLanceFiche('MBO','ARTICLE','',RefArt,'ACTION=CONSULTATION');
END;

procedure TOF_MBOAFFTRANSFERT.VoirEtabliss;
var TOBL : TOB;
    Etabliss : string;
BEGIN
TOBL := GetTOBLigne(G_SAI.Row); if TOBL=nil then exit;
Etabliss:=TOBL.GetValue('GTL_DEPOTDEST');
if Etabliss=RepCodeEmett then Etabliss:=DepotEmetteur;
V_PGI.DispatchTT(18,taConsult,Etabliss,'','');
END;


procedure TOF_MBOAFFTRANSFERT.G_SAIDblClick(Sender: TObject);
var TOBL : TOB;
    ArticleGen,CodeDepot,StatutArt,AuDessus,TempDepot : string;
    Top,Left,Height,Width,i : Integer ;
    CelluleEcran : Tpoint ;
begin
TOBL := GetTOBLigne (G_Sai.Row); if TOBL=nil then exit;
TempDepot:='';
ArticleGen:=TOBL.GetValue('GTL_CODEARTICLE');
CodeDepot:=TOBL.GetValue('GTL_DEPOTDEST');
// Cas du dépôt Emetteur
StatutArt:=TOBL.GetValue('GTL_STATUTART');
if (StatutArt='GEN') then
  begin
  ChargeTOBDimensions(CodeDepot,ArticleGen,Mode);
  CelluleEcran:=RetourneCoordonneeCellule(1,G_SAI)  ;
  if CelluleEcran.y >= 372 then AuDessus:='X' else AuDessus:='-' ;
  Top:=CelluleEcran.y ;
  Left:=CelluleEcran.x ;
  Height:=G_SAI.Height-G_SAI.RowHeights[0]-G_SAI.RowHeights[1] ;
  Width:=G_SAI.Width-G_SAI.ColWidths[0] ;
  // Traiter le cas du dépôt émetteur en interdisant la modification dans l'objet dimension
  TheTOB:=TobDimensions ;
  if (Action<>taModif) then
    AglLanceFiche('MBO','TRANSFERTSAIDIM','','', 'GA_CODEARTICLE='+ArticleGen+';ACTION=CONSULT;TOP='+IntToStr(Top)+';LEFT='+IntToStr(Left)+';OU='+AuDessus+';TYPEPARAM=PSA;HEIGTH='+IntToStr(Height)+';WIDTH='+IntToStr(Width)+';CODEDEPOT='+CodeDepot+';MULTIDEPOT='+IntToStr(Mode))
  else
    AglLanceFiche('MBO','TRANSFERTSAIDIM','','', 'GA_CODEARTICLE='+ArticleGen+';ACTION=SAISIE;TOP='+IntToStr(Top)+';LEFT='+IntToStr(Left)+';OU='+AuDessus+';TYPEPARAM=PSA;HEIGTH='+IntToStr(Height)+';WIDTH='+IntToStr(Width)+';CODEDEPOT='+CodeDepot+';MULTIDEPOT='+IntToStr(Mode)) ;
  TOBDimensions:=TheTOB;
  TheTOB:=nil;
  if TOBDimensions<>nil then
    begin
    if TOBDimensions.IsOneModifie then
      begin
      TraiteLesDim();
      if Mode=1 then
        begin
        For i:=0 to TOBDimensions.Detail.Count-1 do
          begin
          if TOBDimensions.Detail[i].IsOneModifie then
            begin
            CodeDepot:=TOBDimensions.Detail[i].GetValue('GTL_DEPOTDEST');
            if (TempDepot<>CodeDepot) and (CodeDepot<>RepCodeEmett) then TraiteLesGen(CodeDepot,ArticleGen);
            TempDepot:=CodeDepot;
            end;
          end;
        end
      else TraiteLesGen(CodeDepot,ArticleGen);
      end;
    end;
    TobDimensions.Free; TobDimensions:=nil;
  end;
end;

function TOF_MBOAFFTRANSFERT.VerifDispo (TOBL : TOB; var QteDiff: Double) : Boolean;
var QteDispo : Double;
    RefArt : string;
    TOBD : TOB;
BEGIN
Result:=False;
RefArt:=TOBL.GetValue('GTL_ARTICLE');
TOBD:=TOBProp.FindFirst(['GTL_ARTICLE','GTL_DEPOTDEST'],[RefArt,RepCodeEmett],False);
if TOBD=Nil then exit;
QteDispo:=Valeur(TOBD.GetValue('GTL_STOCKINITIAL'))+Valeur(TOBD.GetValue('GTL_PROPOSITION'));
if QteDispo < QteDiff then
    BEGIN
    if QteDispo > 0 then QteDiff:=QteDispo else QteDiff:=0;
    exit;
    END;
TOBD.PutValue('GTL_PROPOSITION', TOBD.GetValue('GTL_PROPOSITION')-QteDiff);
TOBD.PutValue('GTL_MANQUANT', TOBD.GetValue('GTL_MANQUANT')-QteDiff);
TOBD.PutValue('GTL_SURPLUS', TOBD.GetValue('GTL_SURPLUS')-QteDiff);
if Mode=2 then
  begin
  G_SEL.Cells[Sel_QteAffect,G_SEL.Row]:=FormatFloat(G_SEL.ColFormats[Sel_QteAffect],Valeur(G_SEL.Cells[Sel_QteAffect,G_SEL.Row])+QteDiff);
  G_SEL.Cells[Sel_QteAffect,1]:=FormatFloat(G_SEL.ColFormats[Sel_QteAffect],Valeur(G_SEL.Cells[Sel_QteAffect,1])-QteDiff);
  end;
Result:=True;
END;

procedure TOF_MBOAFFTRANSFERT.TraiterProposition(ACol, ARow : integer);
var TOBL,TOBP,TOBLD : TOB;
    St,StatutArt : String;
    QteProp, QteAnc, QteDiff : double ;
BEGIN
TOBL := GetTOBLigne(ARow); if TOBL=nil then exit;
StatutArt:=TOBL.GetValue('GTL_STATUTART');
QteAnc:=TOBL.GetValue('GTL_PROPOSITION');
QteProp:=Valeur(G_SAI.Cells [ACol, ARow]) ;
QteDiff:= QteProp - QteAnc;
if (QteProp < 0) or (StatutArt<>'UNI') or (QteDiff=0) then
  begin
  St:=FormatFloat(G_SAI.ColFormats[Sai_Prop], QteAnc);
  G_SAI.Cells[ACol, ARow]:=St;
  end
else
  begin
  if Not VerifDispo(TOBL, QteDiff) then
    begin
    if QteProp <> 0 then
        HShowMessage('0;'+TFMul(Ecran).Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    St:=FormatFloat(G_SAI.ColFormats[Sai_Prop], QteAnc);
    G_SAI.Cells [ACol, ARow]:=St;
    end
  else
    begin
    St:=FormatFloat(G_SAI.ColFormats[Sai_Prop], QteProp);
    G_SAI.Cells[ACol, ARow]:=St;
    G_SAI.Cells[Sai_Manq, ARow]:=FormatFloat(G_SAI.ColFormats[Sai_Manq],Valeur(G_SAI.Cells[Sai_Manq, ARow])-QteDiff);
    G_SAI.Cells[Sai_StockFin, ARow]:=FormatFloat(G_SAI.ColFormats[Sai_StockFin],Valeur(G_SAI.Cells[Sai_StockFin, ARow])+QteDiff);
    // MAJ de la ligne du dépôt émetteur
    if mode=1 then
      begin
      G_SAI.Cells[Sai_Prop, 1]:=FormatFloat(G_SAI.ColFormats[Sai_Prop],Valeur(G_SAI.Cells[Sai_Prop,1])-QteDiff);
      G_SAI.Cells[Sai_Manq, 1]:=FormatFloat(G_SAI.ColFormats[Sai_Manq],Valeur(G_SAI.Cells[Sai_Manq,1])-QteDiff);
      G_SAI.Cells[Sai_Surp, 1]:=FormatFloat(G_SAI.ColFormats[Sai_Surp],Valeur(G_SAI.Cells[Sai_Surp,1])-QteDiff);
      G_SAI.Cells[Sai_StockFin, 1]:=FormatFloat(G_SAI.ColFormats[Sai_StockFin],Valeur(G_SAI.Cells[Sai_StockFin,1])-QteDiff);
      TOBLD := GetTOBLigne(ARow);
      if TOBLD<>nil then TOBLD.PutValue('GTL_PROPOSITION', TOBLD.GetValue('GTL_PROPOSITION')-QteDiff);
      end;
    TOBL.PutValue('GTL_PROPOSITION', St);
    TOBP:=TOBProp.FindFirst(['GTL_DEPOTDEST','GTL_ARTICLE','GTL_STATUTART'],
          [TOBL.GetValue('GTL_DEPOTDEST'),TOBL.GetValue('GTL_ARTICLE'),StatutArt],
          False); //True
    if TOBP<>Nil then
      begin
      TOBP.PutValue('GTL_PROPOSITION', St);
      TOBP.PutValue('GTL_MANQUANT', TOBP.GetValue('GTL_MANQUANT')-QteDiff);
      end;
    end;
  end;
END;


procedure TOF_MBOAFFTRANSFERT.G_SAICellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if ACol = Sai_Prop then TraiterProposition(ACol, ARow);
end;

procedure TOF_MBOAFFTRANSFERT.G_SAIRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
G_SAI.InvalidateRow(ou) ;
if Mode=2 then InfoArticle;
end;

procedure TOF_MBOAFFTRANSFERT.G_SAIRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
G_SAI.InvalidateRow(ou) ;
end;


procedure TOF_MBOAFFTRANSFERT.G_SELDessineTriangle (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < G_SEL.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_SEL.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_SEL.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_SEL.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

procedure TOF_MBOAFFTRANSFERT.G_SAIDessineTriangle (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
BEGIN
If Arow < G_SAI.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_SAI.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_SAI.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_SAI.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
END;

procedure TOF_MBOAFFTRANSFERT.InitialiseGrille;
var i,j,Dec, FixedWidthSel, FixedWidthSai : integer ;
    St,Nam,stAl,stA,CH, FF, FPerso : string ;
    NomList,FRecordSource,FLien,FSortBy,FLargeur,FAlignement,FParams,tt,NC : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;
begin
LigneSAI:=0;

Sel_Fixed:=0; Sel_Art:=-1; Sel_CodeArt:=-1; Sel_Lib:=-1; Sel_QteAAffect:=-1;
Sel_DepDest:=-1; Sel_QteAffect:=-1; Sel_DepLib:=-1;

Sai_Fixed:=0; Sai_DepDest:=-1; Sai_Art:=-1; Sai_CodeArt:=-1;
Sai_StockIni:=-1; Sai_StockMini:=-1; Sai_StockMaxi:=-1; Sai_QteVte:=-1; Sai_StockALF:=-1;
Sai_Manq:=-1; Sai_Surp:=-1; Sai_Prop:=-1; Sai_StockFin :=-1; Sai_DepLib:=1;
Sai_StockNet:=-1;

FixedWidthSel:=10; FixedWidthSai:=10;

NomList:=TFMul(Ecran).Q.Liste;
ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,FPerso,OkTri,OkNumCol);

St:=LesColSel[mode] ;
NbColSel:=0;
While St<> '' do
    BEGIN
    ReadTokenSt(St) ;
    inc(NbColSel);
    END;
G_SEL.ColCount:=NbColSel;
St:=LesColSel[mode] ;
for i:=0 to G_SEL.ColCount-1 do
    BEGIN
    if i>1 then  G_SEL.ColWidths[i]:=100 else G_SEL.ColWidths[0]:=FixedWidthSel;
    Nam:=ReadTokenSt(St) ;
    if Nam='GTL_ARTICLE' then Sel_Art:=i
    else if Nam='GTL_CODEARTICLE' then Sel_CodeArt:=i
    else if Nam='LIBELLE' then Sel_Lib:=i
    else if Nam='QTEAAFFECTER' then Sel_QteAAffect:=i
    else if Nam='GTL_DEPOTDEST' then Sel_DepDest:=i
    else if Nam='DEPOTLIBELLE' then Sel_DepLib:=i
    else if Nam='QTEAFFECTER' then Sel_QteAffect:=i
    ;
    CH:=FFieldList ;
    StAl:=FAlignement;
{$IFDEF EAGLCLIENT}
    for j:=0 to TFMul(Ecran).Fliste.ColCount - 1 do
        BEGIN
        StA:=ReadTokenSt(StAl);
        TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if ReadTokenSt(CH)=Nam then
            BEGIN
            G_SEL.ColAligns[i]:=TFMul(Ecran).Fliste.ColAligns[j] ;
            G_SEL.ColWidths[i]:=TFMul(Ecran).Fliste.ColWidths[j] ;
            if OkLib then G_SEL.ColFormats[i]:='CB=' + Get_Join(Nam)
                     else if (Dec<>0) or (Sep) then G_SEL.ColFormats[i]:=FF ;
            break;
            END;
        END;
    END ;
{$ELSE}
    for j:=0 to TFMul(Ecran).Fliste.Columns.Count - 1 do
        BEGIN
        StA:=ReadTokenSt(StAl);
        TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if ReadTokenSt(CH)=Nam then
            BEGIN
            G_SEL.ColAligns[i]:=TFMul(Ecran).Fliste.Columns.Items[j].Field.Alignment;
            G_SEL.ColWidths[i]:=TFMul(Ecran).Fliste.Columns.Items[j].Width;
            if OkLib then G_SEL.ColFormats[i]:='CB=' + Get_Join(Nam)
                     else if (Dec<>0) or (Sep) then G_SEL.ColFormats[i]:=FF ;
            break;
            END;
        END;
    END ;
{$ENDIF}
St:=LesColSai[mode];
NbColSai:=0;
While St<> '' do
    BEGIN
    ReadTokenSt(St) ;
    inc(NbColSai);
    END;
G_SAI.ColCount:=NbColSai;
St:=LesColSai[mode];
for i:=0 to G_SAI.ColCount-1 do
    BEGIN
    if i>1 then  G_SAI.ColWidths[i]:=100 else G_SAI.ColWidths[0]:=FixedWidthSai;
    Nam:=ReadTokenSt(St) ;
    if Action=taModif then
      begin
      if Nam='GTL_DEPOTDEST' then Begin Sai_DepDest:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_STOCKINITIAL' then Begin Sai_StockIni:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_STOCKMINI' then Begin Sai_StockMini:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_STOCKMAXI' then Begin Sai_StockMaxi:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_QTEVTE' then Begin Sai_QteVte:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_STOCKALF' then Begin Sai_StockALF:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_STOCKNET' then Begin Sai_StockNet:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_MANQUANT' then Begin Sai_Manq:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_SURPLUS' then Begin Sai_Surp:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_ARTICLE' then Begin Sai_Art:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_CODEARTICLE' then Begin Sai_CodeArt:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='STOCKFINAL' then begin Sai_StockFin:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='DEPOTLIBELLE' then begin Sai_DepLib:=i; G_SAI.ColLengths[i]:=-1; end
      else if Nam='GTL_PROPOSITION' then Sai_Prop:=i
      ;
      end
    else
      begin
      if Nam='GTL_DEPOTDEST' then Sai_DepDest:=i
      else if Nam='GTL_STOCKINITIAL' then Sai_StockIni:=i
      else if Nam='GTL_STOCKMINI' then Sai_StockMini:=i
      else if Nam='GTL_STOCKMAXI' then Sai_StockMaxi:=i
      else if Nam='GTL_QTEVTE' then Sai_QteVte:=i
      else if Nam='GTL_STOCKALF' then Sai_StockALF:=i
      else if Nam='GTL_STOCKNET' then Sai_StockNet:=i
      else if Nam='GTL_MANQUANT' then Sai_Manq:=i
      else if Nam='GTL_SURPLUS' then Sai_Surp:=i
      else if Nam='GTL_ARTICLE' then Sai_Art:=i
      else if Nam='GTL_CODEARTICLE' then Sai_CodeArt:=i
      else if Nam='STOCKFINAL' then Sai_StockFin:=i
      else if Nam='DEPOTLIBELLE' then Sai_DepLib:=i
      else if Nam='GTL_PROPOSITION' then Sai_Prop:=i
      ;
      end;
    CH:=FFieldList ;
    StAl:=FAlignement;
{$IFDEF EAGLCLIENT}
    for j:=0 to TFMul(Ecran).Fliste.ColCount - 1 do
        BEGIN
        StA:=ReadTokenSt(StAl);
        TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if ReadTokenSt(CH)=Nam then
            BEGIN
            G_SAI.ColAligns[i]:=TFMul(Ecran).Fliste.ColAligns[j] ;
            G_SAI.ColWidths[i]:=TFMul(Ecran).Fliste.ColWidths[j] ;
            if OkLib then G_SAi.ColFormats[i]:='CB=' + Get_Join(Nam)
                     else if (Dec<>0) or (Sep) then G_SAI.ColFormats[i]:=FF ;
            break;
            END;
        END;
{$ELSE}
    for j:=0 to TFMul(Ecran).Fliste.Columns.Count - 1 do
        BEGIN
        StA:=ReadTokenSt(StAl);
        TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if ReadTokenSt(CH)=Nam then
            BEGIN
            G_SAI.ColAligns[i]:=TFMul(Ecran).Fliste.Columns.Items[j].Field.Alignment;
            G_SAI.ColWidths[i]:=TFMul(Ecran).Fliste.Columns.Items[j].Width;
            if OkLib then G_SAi.ColFormats[i]:='CB=' + Get_Join(Nam)
                     else if (Dec<>0) or (Sep) then G_SAI.ColFormats[i]:=FF ;
            break;
            END;
        END;
{$ENDIF}
    END ;
//Formate les colonnes qui ne sont pas dans la liste PR_AFFTRANSFERT
if mode=1 then
  begin
  if Sel_Lib<>-1 then
    begin
    G_SEL.ColWidths[Sel_Lib]:=200;
    G_SEL.ColAligns[Sel_Lib]:=taLeftJustify;
    end;
  if Sel_Art<>-1 then G_Sel.ColWidths[Sel_Art]:=0;
  if (Sel_QteAAffect<>-1) and (Sai_Prop<>-1) then
    begin
    G_SEL.ColAligns[Sel_QteAAffect]:=G_SAI.ColAligns[Sai_Prop];
    G_SEL.ColWidths[Sel_QteAAffect]:=G_SAI.ColWidths[Sai_Prop];
    G_SEL.ColFormats[Sel_QteAAffect]:=G_SAI.ColFormats[Sai_Prop];
    end;
  if Sai_DepDest<>-1 then G_Sai.ColWidths[Sai_DepDest]:=0;
  if Sai_DepLib<>-1 then
    begin
    G_Sai.ColWidths[Sai_DepLib]:=200;
    G_Sai.ColAligns[Sai_DepLib]:=taLeftJustify;
    end;
  end
else
  begin
  if (Sel_QteAffect<>-1) and (Sai_Prop<>-1) then
    begin
    G_Sel.ColAligns[Sel_QteAffect]:=G_SAI.ColAligns[Sai_Prop];
    G_Sel.ColWidths[Sel_QteAffect]:=G_SAI.ColWidths[Sai_Prop];
    G_Sel.ColFormats[Sel_QteAffect]:=G_SAI.ColFormats[Sai_Prop];
    end;
  if Sel_DepDest<>-1 then G_Sel.ColWidths[Sel_DepDest]:=0;
  if Sel_DepLib<>-1 then
    begin
    G_Sel.ColWidths[Sel_DepLib]:=200;
    G_Sel.ColAligns[Sel_DepLib]:=taLeftJustify;
    end;
  if Sai_Art<>-1 then G_Sai.ColWidths[Sai_Art]:=0;
  end;
if (Sai_Prop<>-1) then
begin
  if (Sai_StockFin<>-1) then
  begin
    G_SAI.ColAligns[Sai_StockFin]:=G_SAI.ColAligns[Sai_Prop];
    G_SAI.ColWidths[Sai_StockFin]:=G_SAI.ColWidths[Sai_Prop];
    G_SAi.ColFormats[Sai_StockFin]:=G_SAI.ColFormats[Sai_Prop];
  end;
  if (Sai_StockNet<>-1) then
  begin
    G_SAI.ColAligns[Sai_StockNet]:=G_SAI.ColAligns[Sai_Prop];
    G_SAI.ColWidths[Sai_StockNet]:=G_SAI.ColWidths[Sai_Prop];
    G_SAi.ColFormats[Sai_StockNet]:=G_SAI.ColFormats[Sai_Prop];
  end;
end;
AffecteGrid(G_SEL,taConsult) ;
G_SEL.ColWidths[0]:=FixedWidthSel;
AffecteGrid(G_SAI,Action) ;
G_SAI.ColWidths[0]:=FixedWidthSai;
end;

procedure TOF_MBOAFFTRANSFERT.ChargeGrille;
Var st, code, libel, CH : string ;
    ACol, icol : integer;
    TOBTemp, TOBTempD, TOBD1 : TOB;
begin
CH:=FFieldList ;
TOBTemp:=TOB.Create('', Nil,-1);
Code:=ReadTokenSt(CH); Libel:=ReadTokenSt(FTitre);
while Code<>'' do
  begin
  TOBTempD:=TOB.Create('', TOBTemp,-1);
  TOBTempD.AddChampSup('CODE',False); TOBTempD.PutValue('CODE',Code) ;
  TOBTempD.AddChampSup('LIBELLE',False); TOBTempD.PutValue('LIBELLE',Libel) ;
  Code:=ReadTokenSt(CH); Libel:=ReadTokenSt(FTitre);
  end;
st:= LesColSEL[mode];
ACol:=0; icol:=0;
SetLength(CodeSEL,G_SEL.ColCount);
While st <> '' do
    BEGIN
    code:=ReadTokenSt(st);
    TOBD1:=TOBTemp.FindFirst(['CODE'],[code],False); //True
    if TOBD1 <> nil then Libel:=TOBD1.GetValue('LIBELLE') else Libel:='';
    G_SEL.Cells[ACol, 0]:= Libel;
    inc(ACol);
    if TOBD1 <> nil then begin CodeSEL[icol]:=Code; inc(icol); end;
    END;
CodeSEL:= Copy(CodeSEL, 0, icol);
st:= LesColSAI[mode];
ACol:=0; icol:=0;
SetLength(CodeSAI,G_SAI.ColCount);
While st <> '' do
    BEGIN
    code:=ReadTokenSt(st);
    TOBD1:=TOBTemp.FindFirst(['CODE'],[code],False); //True
    if TOBD1 <> nil then Libel:=TOBD1.GetValue('LIBELLE') else Libel:='';
    G_SAI.Cells[ACol, 0]:= Libel;
    inc(ACol);
    if TOBD1 <> nil then begin CodeSAI[icol]:=Code; inc(icol); end;
    END;
if Mode=1 then
  begin
  if Sel_Lib<>-1 then G_SEL.Cells[Sel_Lib, 0]:= 'Désignation';
  if Sel_QteAAffect<>-1 then G_SEL.Cells[Sel_QteAAffect, 0]:= 'Qté à affecter';
  if Sai_DepLib<>-1 then G_Sai.Cells[Sai_DepLib, 0]:= 'Etablissement';
  end
else
  begin
  if Sel_QteAffect<>-1 then G_SEL.Cells[Sel_QteAffect, 0]:= 'Qté affectée';
  if Sel_DepLib<>-1 then G_SEL.Cells[Sel_DepLib, 0]:= 'Etablissement';
  end;
if Sai_StockFin<>-1 then G_SAI.Cells[Sai_StockFin, 0]:= 'Stock final';
if Sai_StockALF<>-1 then G_SAI.Cells[Sai_StockALF, 0]:= 'Stock transit';
if Sai_StockNet<>-1 then G_SAI.Cells[Sai_StockNet, 0]:= 'Stock net';
CodeSAI:= Copy(CodeSAI, 0, icol);
TOBTemp.Free;
AffectGrilleSEL(G_SEL,False,False,CodeSEL,True);
AffectGrilleSAI(G_SAI,False,False,CodeSEL,True);
InfoArticle;
end;

procedure TOF_MBOAFFTRANSFERT.InfoArticle;
var StInfo,StatutArt : string;
    TOBL : TOB;
begin
TOBL:=GetTOBLigne(G_SAI.Row); if TOBL=nil then exit;
StatutArt:=TOBL.GetValue('GTL_STATUTART');
if StatutArt='GEN' then StInfo:='Article dimensionné : double cliquer pour voir le détail'
else StInfo:='';
SetControlCaption('TINFOARTICLE',StInfo);
end;

procedure TOF_MBOAFFTRANSFERT.AffectGrilleSEL(G : THGrid ; Entete,EntLib :boolean ;
                          CodeChp: array of string; CalculRowCount : Boolean);
var TOBT, TOBTD,  TOBRefD : TOB;
    TabValeur : array of variant;
    i_ind, icode : integer;
    StatutArt,StArticle,CodeArticle,CodeDepotDest,DepotLibelle : string;
    TOBL : TOB;
BEGIN
if TOBSel<>nil then begin TOBSel.Free; TOBSel:=nil; end;
TOBSel := TOB.Create ('', Nil, -1) ;
INITMOVE(TOBPROP.Detail.count,'');
while i_ind < TOBPROP.Detail.count do
  begin
  MoveCur(False);
  StatutArt:=TOBPROP.Detail[i_ind].GetValue('GTL_STATUTART');
  if (StatutArt='UNI') or (StatutArt='GEN') then
    begin
    StArticle:=TOBPROP.Detail[i_ind].GetValue('GTL_ARTICLE');
    CodeArticle:=TOBPROP.Detail[i_ind].GetValue('GTL_CODEARTICLE');
    CodeDepotDest:=TOBPROP.Detail[i_ind].GetValue('GTL_DEPOTDEST');
    if mode=1 then
      begin
      i_ind:=i_ind+1;
      if TOBSel.FindFirst(['GTL_ARTICLE'],[StArticle],False) <> Nil then continue;
      TOBL:=TOB.Create ('', TOBSel,-1) ;
      TOBL.AddChampSup('GTL_ARTICLE', False); TOBL.PutValue('GTL_ARTICLE', StArticle);
      TOBL.AddChampSup('GTL_CODEARTICLE', False); TOBL.PutValue('GTL_CODEARTICLE', CodeArticle);
      TOBL.AddChampSup('LIBELLE', False); TOBL.PutValue('LIBELLE', RechDom('GCARTICLE',StArticle,False));
      TOBL.AddChampSup('QTEAAFFECTER', False); TOBL.PutValue('QTEAAFFECTER', CalculQteAAffecter(StArticle));
      end
    else
      begin
      if TOBSel.FindFirst(['GTL_DEPOTDEST'],[CodeDepotDest],False) <> Nil then
        begin i_ind:=i_ind+1; continue; end;
      TOBL:=TOB.Create ('', TOBSel,-1) ;
      TOBL.AddChampSup('GTL_DEPOTDEST', False); TOBL.PutValue('GTL_DEPOTDEST', CodeDepotDest);
      if CodeDepotDest=RepCodeEmett then DepotLibelle:='Dépôt : '+RechDom('GCDEPOT',DepotEmetteur,False)
      else DepotLibelle:=RechDom('GCDEPOT',CodeDepotDest,False);
      TOBL.AddChampSup('DEPOTLIBELLE', False); TOBL.PutValue('DEPOTLIBELLE', DepotLibelle);
      TOBL.AddChampSup('QTEAFFECTER', False); TOBL.PutValue('QTEAFFECTER', CalculQteAffecter(CodeDepotDest,i_ind));
      end;
    end
  else i_ind:=i_ind+1;
  end;
FINIMOVE;
TOBT:=TOB.Create ('', Nil, -1);
SetLength(TabValeur,Length(CodeChp));
INITMOVE(TOBSel.Detail.count,'');
for i_ind:=0 to TOBSel.Detail.count - 1 do
    BEGIN
    MoveCur(False);
    TOBRefD:= TOBSel.Detail[i_ind];
    for icode:=Low(CodeChp) to High(CodeChp) do
        BEGIN
        if CodeChp[icode]<>'' then TabValeur[icode]:=TOBRefD.GetValue(CodeChp[icode]);
        END;
    if TOBT.FindFirst(CodeChp,TabValeur,False) <> Nil then continue; //True
    TOBTD:=TOB.Create ('', TOBT,-1) ;
    TOBTD.Dupliquer(TOBRefD, False, True);
    END;
TOBT.PutGridDetail(G,Entete,EntLib,LesColSEL[mode],CalculRowCount);
FINIMOVE;
TOBT.Free;
end;


procedure TOF_MBOAFFTRANSFERT.AffectGrilleSAI(G : THGrid; Entete,EntLib :boolean;
                          CodeChp: array of string; CalculRowCount : Boolean);
var TOBT, TOBTD,  TOBRefD : TOB;
    TabValeur : array of variant;
    icode : integer;
    StName,StatutArt,CodeDepotDest,DepotLibelle : string;
    StockFinal: double;
begin
SetLength(TabValeur,Length(CodeChp));
StName:=LesColSEL[mode];

TOBT:=TOB.Create ('', Nil, -1);
while StName <> '' do TOBT.AddChampSup(ReadTokenSt(StName), False);
TOBT.GetLigneGrid(G_SEL,G_SEL.Row,LesColSEL[mode]);
for icode:=Low(CodeChp) to High(CodeChp) do TabValeur[icode]:=TOBT.GetValue(CodeChp[icode]);
TOBT.Free;

if TOBSAI<>nil then
  begin TOBSAI.free; TOBSAI:=Nil; end;
TOBSAI:=TOB.Create ('', Nil, -1);
TOBRefD:= TOBProp.FindFirst(CodeChp,TabValeur,False); //True
While TOBRefD <> nil do
    BEGIN
    StatutArt:=TOBRefD.GetValue('GTL_STATUTART');
    if (StatutArt='UNI') or (StatutArt='GEN') then
      begin
      TOBTD:=TOB.Create ('', TOBSAI,-1) ;
      TOBTD.Dupliquer(TOBRefD, False, True);
      if Mode=1 then
        begin
        CodeDepotDest:=TOBTD.GetValue('GTL_DEPOTDEST');
        if CodeDepotDest=RepCodeEmett then DepotLibelle:='Dépôt : '+RechDom('GCDEPOT',DepotEmetteur,False)
        else DepotLibelle:=RechDom('GCDEPOT',CodeDepotDest,False);
        TOBTD.AddChampSup('DEPOTLIBELLE', False); TOBTD.PutValue('DEPOTLIBELLE', DepotLibelle);
        end;
      StockFinal := Valeur(TOBTD.GetValue('GTL_STOCKINITIAL'))
                   +Valeur(TOBTD.GetValue('GTL_PROPOSITION'))
                   +Valeur(TOBTD.GetValue('GTL_STOCKALF')); //Modif 31/05/2002
      TOBTD.AddChampSup('STOCKFINAL', False); TOBTD.PutValue('STOCKFINAL',StockFinal);
      end;
    TOBRefD:= TOBProp.FindNext(CodeChp,TabValeur,False);  //true
    END;
TOBSAI.PutGridDetail(G,Entete,EntLib,LesColSAI[mode],CalculRowCount);
end;

procedure TOF_MBOAFFTRANSFERT.BOUVRIRClick(Sender: TObject);
begin
if Action=taModif then Valider:=True;
TFMul(ecran).Close;
end;

procedure TOF_MBOAFFTRANSFERT.OnClose ;
begin
  Inherited ;
SPLITTER.Free;
if Valider then TheTOB:=TOBPROP
           else BEGIN if TOBProp<>Nil then TOBProp.Free ; TheTOB:=Nil ; END ;
if TOBSel<>Nil then TOBSel.Free;
if TOBSai<>Nil then TOBSai.Free;
end ;


Initialization
  registerclasses ( [ TOF_MBOAFFTRANSFERT ] ) ;
end.
