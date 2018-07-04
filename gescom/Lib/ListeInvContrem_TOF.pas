{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 22/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCLISTEINVCONTREM ()
Mots clefs ... : TOF;GCLISTEINVCONTREM
*****************************************************************}
Unit ListeInvContrem_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,M3FP,
{$IFNDEF EAGLCLIENT}
     db,FE_Main,
     dbtables,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,Menus,windows,
     sysutils,Dialogs,
     ComCtrls,EntGC,
     HCtrls,HTB97,
     HEnt1,Hpanel,HDimension,
     HMsgBox,HSysMenu,ParamSoc,
     UTOF,UTOB,SaisUtil ;


function GCLanceFiche_ListeInvContre(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_GCLISTEINVCONTREM = Class (TOF)

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    GInv : THGrid;
    TotQtePhy,TotQteSai,TotQteEcart : THNumEdit;
    bDetailPrix,bChercher : TToolbarButton97;
    bZoomArticle,bZoomMvt : TMenuItem;
    CBEmplacement : THValCombobox;
    PCumul : THPanel;
    HMTrad: THSystemMenu;
    TobListeInv,TobArticle,TobEmplacement : TOB;
    ColPhy,ColQteSai,ColEcart : integer;
    FDialog : TFindDialog;
    FFirstFind : boolean;

    //Grid
    procedure GInvRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GInvDblClick(Sender: TObject);
    procedure GridColumnWidthsChanged(Sender: TObject);

    procedure EtudieColsListes ;
    procedure RowEnter(ARow : integer);
    //Boutons
    procedure BZoomArticleClick(Sender: TObject);
    procedure BZoomMvtClick(Sender: TObject);
    procedure bDetailPrixClick(Sender: TObject);
    procedure TDetailPrixClose(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FDialogFind(Sender: TObject);
    procedure ZoomCatalogue;
    //Form
    procedure ActualisePrix(TobInv : TOB);
    procedure CBEmplacementChange(Sender: TObject);
    procedure MajPied;
    //Tob
    procedure ChargeTobListe(CodeListe : string);
    function  GetDimensions(Article : string) : string;




  end ;

Implementation

function GCLanceFiche_ListeInvContre(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_GCLISTEINVCONTREM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINVCONTREM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINVCONTREM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINVCONTREM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINVCONTREM.OnArgument (S : String ) ;
begin
  Inherited ;
///Initialisation
FDialog := TFindDialog.Create(nil);
FDialog.OnFind := FDialogFind;
bChercher := TToolbarButton97(GetControl('BCHERCHER'));
BChercher.OnClick := BChercherClick;
BZoomArticle := TMenuItem(GetControl('mnZoomArticle'));
BZoomMvt := TMenuItem(GetControl('mnZoomMvt'));
PCumul := THPanel(GetControl('PCUMUL'));
TotQtePhy   := THNumEdit(GetControl('TOTQTEPHY')) ;
TotQteSai := THNumEdit(GetControl('TOTQTESAI')) ;
TotQteEcart := THNumEdit(GetControl('TOTQTEECART')) ;
bDetailPrix := TToolbarButton97(GetControl('BDETAILPRIX'));
CBEmplacement := THValCombobox(GetControl('GIM_EMPLACEMENT'));
bDetailPrix.OnClick := bDetailPrixClick;
TToolWindow97(GetControl('TDetailPrix')).OnClose := TDetailPrixClose;
BZoomArticle.OnClick := BZoomArticleClick;
BZoomMvt.OnClick := BZoomMvtClick;
CBEmplacement.OnChange := CBEmplacementChange;
//Grid
GInv := THGrid(GetControl('GINV')) ;
GInv.OnColumnWidthsChanged := GridColumnWidthsChanged;
GInv.OnRowEnter := GInvRowEnter;
GInv.OnDblClick := GInvDblClick;
// Tob
TobListeInv := TOB.Create('LISTEINVENT',nil,-1) ;
TobArticle := TOB.Create('',nil,-1) ;
TobEmplacement := TOB.Create('',nil,-1) ;//contient les filles de TobListeInv ki ne sont pas sur l'emplacem sélectionné
// Affectation
GInv.ListeParam:='GCLISTEINVLIGCON';
EtudieColsListes;
AffecteGrid(GInv, taConsult);
ChargeTobListe(S);
SetControlCaption('TCODELISTE',TobListeInv.GetValue('GIE_CODELISTE'));
SetControlCaption('TLIBELLE',TobListeInv.GetValue('GIE_LIBELLE'));
SetControlCaption('TDEPOT',RechDom('GCDEPOT',TobListeInv.GetValue('GIE_DEPOT'),False));
SetControlCaption('TDATE',TobListeInv.GetValue('GIE_DATEINVENTAIRE'));
TobListeInv.PutGridDetail(GInv,False,False,GInv.Titres[0]);
HMTrad.ResizeGridColumns(GInv);
MajPied;
RowEnter(1);
end ;

procedure TOF_GCLISTEINVCONTREM.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINVCONTREM.GInvRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
RowEnter(Ou);
end;

procedure TOF_GCLISTEINVCONTREM.GInvDblClick(Sender: TObject);
begin
ZoomCatalogue;
end;

procedure TOF_GCLISTEINVCONTREM.GridColumnWidthsChanged(Sender: TObject);
var Coord : TRect;
begin
Coord:=GInv.CellRect(ColPhy,0);
TotQtePhy.Left:=Coord.Left + 1;
TotQtePhy.Width:=GInv.ColWidths[ColPhy] + 1;

Coord:=GInv.CellRect(ColQteSai,0);
TotQteSai.Left:=Coord.Left + 1;
TotQteSai.Width:=GInv.ColWidths[ColQteSai] + 1;

Coord:=GInv.CellRect(ColEcart,0);
TotQteEcart.Left:=Coord.Left + 1;
TotQteEcart.Width:=GInv.ColWidths[ColEcart] + 1;
end;

procedure TOF_GCLISTEINVCONTREM.BZoomArticleClick(Sender: TObject);
begin
ZoomCatalogue;
end;

procedure TOF_GCLISTEINVCONTREM.BZoomMvtClick(Sender: TObject);
var TobInv : TOB;
    stWhere : string;
begin
TobInv := TOB(GInv.Objects[0,GInv.Row]);
if TobInv = nil then exit;
stWhere := ' ((GL_TIERS="' + TobInv.GetValue('GIM_CLIENT') + '" '
           + 'AND GL_FOURNISSEUR="' + TobInv.GetValue('GIM_FOURNISSEUR') + '") '
           + 'OR (GL_TIERS="' + TobInv.GetValue('GIM_FOURNISSEUR') + '" '
           + 'AND GL_FOURNISSEUR="' + TobInv.GetValue('GIM_CLIENT') + '")) '
           + 'AND GL_REFCATALOGUE="' + TobInv.GetValue('GIM_REFERENCE') + '" '
           + 'AND GL_DEPOT="' + TobListeInv.GetValue('GIE_DEPOT') + '" '
           + 'AND GL_VIVANTE="X" AND GL_QTERESTE > 0'; {DBR NEWPIECE QTERESTE > 0}
AGLLanceFiche('GC','GCINVCONLIG_MUL','','',stWhere);
end;

procedure TOF_GCLISTEINVCONTREM.bDetailPrixClick(Sender: TObject);
var okVisu : boolean;
    TobInv : TOB;
begin
okVisu := not(TToolWindow97(GetControl('TDETAILPRIX')).Visible);
SetControlVisible('TDETAILPRIX',okVisu);
if okVisu then
   begin
   TobInv := TOB(GInv.Objects[0,GInv.Row]);
   ActualisePrix(TobInv);
   end;
end;

procedure TOF_GCLISTEINVCONTREM.TDetailPrixClose(Sender: TObject);
begin
bDetailPrix.Down := TToolWindow97(GetControl('TDETAILPRIX')).Visible;
end;

procedure TOF_GCLISTEINVCONTREM.CBEmplacementChange(Sender: TObject);
var iInd : integer;
    TobE,TobL : TOB;
    st : string;
begin
st := CBEmplacement.Value;
for iInd := TobEmplacement.Detail.Count -1 downto 0 do
    begin
    TobE := TobEmplacement.Detail[iInd];
    if (TobE.GetValue('GIM_EMPLACEMENT') = st) or (st = '') then
       TobE.ChangeParent(TobListeInv,-1);
    end;
if st <> '' then
   begin
   for iInd := TobListeInv.Detail.Count -1 downto 0 do
       begin
       TobL := TobListeInv.Detail[iInd];
       if TobL.GetValue('GIM_EMPLACEMENT') <> st then
          TobL.ChangeParent(TobEmplacement,-1);
       end;
   end;
GInv.VidePile(False);
TobListeInv.PutGridDetail(GInv,False,False,GInv.Titres[0]);
MajPied;
RowEnter(1);
end;

procedure TOF_GCLISTEINVCONTREM.EtudieColsListes ;
Var NomCol,LesCols,DecQte : String ;
    icol  : integer ;
begin
DecQte := '';
for iCol := 1 to GetParamSoc('SO_DECQTE') do DecQte := DecQte + '0';
TotQtePhy.Masks.PositiveMask := '#,##0.' + DecQte;
TotQteSai.Masks.PositiveMask := '#,##0.' + DecQte;
TotQteEcart.Masks.PositiveMask := '#,##0.' + DecQte;

//Grids des lignes de doc
LesCols := GInv.Titres[0];
icol:=1 ;     ///**
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol='GIM_QTEPHOTOINV' then
    begin
      ColPhy:=icol;
      GInv.ColFormats[icol] := '#,##0.' + DecQte;
    end
    else if NomCol='GIM_INVENTAIRE' then
    begin
      ColQteSai:=icol;
      GInv.ColFormats[icol] := '#,##0.' + DecQte;
    end
    else if NomCol='(ECART)' then
    begin
      ColEcart:=icol;
      GInv.ColFormats[icol] := '#,##0.' + DecQte;
    end;
    Inc(icol) ;
Until ((LesCols='') or (NomCol=''));
end;


procedure TOF_GCLISTEINVCONTREM.ChargeTobListe(CodeListe : string);
var iInd : integer;
    TobI : TOB;
begin
TobListeInv.SelectDB('"'+CodeListe+'"',nil);
TobListeInv.LoadDetailFromSQL('SELECT LISTEINVLIGCONTREM.*,GA_STATUTART,'
         + 'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,'
         + 'GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5 FROM LISTEINVLIGCONTREM '
         + 'LEFT JOIN ARTICLE ON GIM_ARTICLE=GA_ARTICLE WHERE GIM_CODELISTE="'+CodeListe+'"');
for iInd := 0 to TobListeInv.Detail.Count -1 do
    begin
    TobI := TobListeInv.Detail[iInd];
    TobI.AddChampSupValeur('(ECART)',
       Arrondi(TobI.GetValue('GIM_INVENTAIRE')-TobI.GetValue('GIM_QTEPHOTOINV'),V_PGI.OkDecQ),False);
    TobI.AddChampSupValeur('(CODEARTICLE)',Trim(Copy(TobI.GetValue('GIM_ARTICLE'),1,18)),False);
    end;

end;

procedure TOF_GCLISTEINVCONTREM.ActualisePrix(TobInv : TOB);
begin
if TobInv = nil then
    begin
    SetControlText('GIM_DPAART',StrS(0,V_PGI.OkDecV));
    SetControlText('GIM_DPA',StrS(0,V_PGI.OkDecV));
    SetControlText('GIM_DPR',StrS(0,V_PGI.OkDecV));
    SetControlText('GIM_DPASAIS',StrS(0,V_PGI.OkDecV));
    SetControlText('GIM_DPRSAIS',StrS(0,V_PGI.OkDecV));
    exit;
    end;
SetControlText('GIM_DPAART',StrS(TobInv.GetValue('GIM_DPAART'),V_PGI.OkDecV));
SetControlText('GIM_DPA',StrS(TobInv.GetValue('GIM_DPA'),V_PGI.OkDecV));
SetControlText('GIM_DPR',StrS(TobInv.GetValue('GIM_DPR'),V_PGI.OkDecV));
SetControlText('GIM_DPASAIS',StrS(TobInv.GetValue('GIM_DPASAIS'),V_PGI.OkDecV));
SetControlText('GIM_DPRSAIS',StrS(TobInv.GetValue('GIM_DPRSAIS'),V_PGI.OkDecV));
end;

procedure TOF_GCLISTEINVCONTREM.RowEnter(ARow : integer);
var TobL : TOB;
    stDim : string;
begin
TobL := TOB(GInv.Objects[0,ARow]);
if TToolWindow97(GetControl('TDETAILPRIX')).Visible then ActualisePrix(TobL);
if TobL = nil then
   begin
   SetControlVisible('PDIMENSION',False);
   exit;
   end;
stDim := GetDimensions(TobL.GetValue('GIM_ARTICLE'));
SetControlText('TDIMENSION',stDim);
SetControlVisible('PDIMENSION',(stDim <> ''));
end;

procedure TOF_GCLISTEINVCONTREM.MajPied;
begin
TotQtePhy.Value := TobListeInv.Somme('GIM_QTEPHOTOINV',[''],[''],False);
TotQteSai.Value := TobListeInv.Somme('GIM_INVENTAIRE',[''],[''],False);
TotQteEcart.Value := TobListeInv.Somme('(ECART)',[''],[''],False);
PCumul.Caption := '     Total   ('+intToStr(TobListeInv.detail.Count)+' références)';
end;

function TOF_GCLISTEINVCONTREM.GetDimensions(Article : string) : string;
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
    TobL : TOB;
begin
result := '' ; StDim:='';
if Article = '' then exit;
TobL := TobListeInv.FindFirst(['GIM_ARTICLE'],[Article],False);
if TobL = nil then exit;
if TobL.GetValue('GA_STATUTART') <> 'DIM' then exit;
for i_indDim := 1 to MaxDimension do
    begin
    GrilleDim := TobL.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
    CodeDim := TobL.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
    LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
    if LibDim <> '' then
       if StDim='' then StDim:=StDim + LibDim
       else StDim := StDim + ' - ' + LibDim;
    end;
Result := stDim;
end;

procedure TOF_GCLISTEINVCONTREM.ZoomCatalogue;
var TobInv : TOB;
begin
TobInv := TOB(GInv.Objects[0,GInv.Row]);
if TobInv = nil then exit;
AGLLanceFiche('GC','GCCATALOGU_SAISI3','',TobInv.GetValue('GIM_REFERENCE')
          + ';' + TobInv.GetValue('GIM_FOURNISSEUR'),'ACTION=CONSULTATION;MONOFICHE');
end;

procedure TOF_GCLISTEINVCONTREM.FDialogFind(Sender: TObject);
begin
Rechercher(GInv,FDialog,FFirstFind);
end;

procedure TOF_GCLISTEINVCONTREM.BChercherClick(Sender: TObject);
begin
if GInv.RowCount<=2 then exit;
FFirstFind := true;
FDialog.Execute;
end;

// Procédure pour l'appel de la saisie d'inventaire depuis le script
procedure AGLSaisieInvContr(Parms : Array of Variant; Nb : Integer);
begin
if Parms[0] = '' then exit;
AGLLanceFiche('GC','GCLISTEINVCONTREM','','',Parms[0]);
end;

Initialization
registerclasses ( [ TOF_GCLISTEINVCONTREM ] ) ;
RegisterAGLProc('SaisieInvContr',false,1,AGLSaisieInvContr);
end.

