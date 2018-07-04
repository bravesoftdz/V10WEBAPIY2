{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 24/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCINVCONTREM ()
                 Inventaire de contremarque
Mots clefs ... : TOF;GCINVCONTREM
*****************************************************************}
Unit InvContrem_TOF ;

Interface

Uses StdCtrls,
     Controls,Windows,
     Classes, buttons,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,FE_Main,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,HTB97,Vierge,Dialogs,
     HMsgBox,HSysMenu,Hpanel,HDimension,
     UTOF,UTOB,Ed_Tools,Ent1,Menus,SaisUtil,ParamSoc,EntGC ;

procedure EntreeSaisieInvContrem;

Type
  TOF_GCINVCONTREM = Class (TOF)

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    private
    GInv : THGrid;
    DPA,DPR,TotQtePhy,TotQteSai,TotQteEcart : THNumEdit;
    bValideInv,bSupprimer,bInserer,bSelectAll,bComplement,bDetailPrix,bChercher : TToolbarButton97;
    bForcer,bSaisi,bNonSaisi,bNext,bZoomArticle,bZoomMvt : TMenuItem;
    PCumul : THPanel;
    TDetailPrix : TToolWindow97;
    HMTrad: THSystemMenu;
    stDepot,CellText : string;
    ColSaisi,ColPhy,ColEcart,ColQteSai, ColDPA, ColDPR : integer;
    TobDispoContrem,TobArticle,TobCatalogue : TOB;
    EventsLogCreate,EventsLogUpdate,EventsLogDelete : TStringList;
    FFirstFind : boolean;
    FindLigne: TFindDialog;

    //Grids
    procedure GInvRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GInvRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GInvCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GInvCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GridColumnWidthsChanged(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ToutLiberer ;
    //boutons
    procedure bForcerClick(Sender: TObject);
    procedure bSaisiClick(Sender: TObject);
    procedure bNonSaisiClick(Sender: TObject);
    procedure bNextClick(Sender: TObject);
    procedure bSupprimerClick(Sender: TObject);
    procedure bInsererClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure bComplementClick(Sender: TObject);
    procedure bValideInvClick(Sender: TObject);
    procedure bDetailPrixClick(Sender: TObject);
    procedure BZoomArticleClick(Sender: TObject);
    procedure BZoomMvtClick(Sender: TObject);
    procedure bRechInCatalogueElipsisClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    //Fonctions liées aux boutons
    procedure TDetailPrixClose(Sender: TObject);
    procedure SupprimeLaLigne(ARow : integer);
    procedure AjouteUneReference;
    procedure ActualisePrix(TobDC : TOB);
    procedure FindLigneFind(Sender: TObject);
    //Fonctions liées au grid
    procedure EtudieColsListes ;
    function  ValideLaCellule(ACol,ARow : integer) : boolean;
    function  ValideLaQte (TobDC : TOB ; QteSaisie : double) : boolean;
    procedure RowEnter(ARow : integer);
    //Affichage
    function  GetDimensions(Article : string) : string;
    procedure MajPiedCumul;
    //Tob
    procedure TobPutEcart(TobDC : TOB);
    function  TobSelectDBContrem(stReq : string) : TOB;
    function  TobCreateContrem(TobMere : TOB) : TOB;
    procedure AddLesChampsSup(TobDC : TOB);
    procedure InitLesChampsSup(TobDC : TOB);
    procedure ControleExistence(TobNew : TOB);
    procedure MajTobPrixSaisi(ARow : integer);
    function  GetPrixArticle(TobDC : TOB ; ChampPrix : string) : double;
    function  GetLibelleArticle(TobDC : TOB) : string;
    //Maj Table  / Création
    procedure UpdateDispocontrem;
    procedure ValideLaListe;
    //Recherche
    function  RechRangDispoContrem : integer;
    function  GetCondition(stPref,stInf,stSup : string) : string;
    function  GetCodeListe(RacineCode : string) : string;

  end ;

const
	// libellés des messages
	TexteMessage: array[1..18] of string 	= (
          {1}        'Le stock inventorié va être forcé à la valeur du stock ordinateur, sur toutes les lignes non saisies !'
          {2}       ,'La saisie est incomplète. Continuer?'
          {3}       ,'Le stock va être mis à jour. Continuer?'
          {4}       ,'Le stock inventorié va être remis à zéro, sur toutes les lignes saisies !'
          {5}       ,'Cette référence existe déjà dans la liste'
          {6}       ,'Liste validée'
          {7}       ,'Vous devez renseigner le code de liste d''inventaire que le traitement va générer'
          {8}       ,'La liste de saisie en cours va être perdue. Continuer ?'
          {9}       ,'Problème rencontré. La mise à jour n''a pu être effectuée.'
          {10}      ,'Liste non validée'
          {11}      ,'Vous devez renseigner l''intitulé de la liste d''inventaire.'
          {12}      ,'Confirmez vous l''abandon de la saisie ?'
          {13}      ,'La mise à jour s''est correctement effectuée.'
          {14}      ,'Etes-vous sûr de vouloir supprimer cette ligne?'
          {15}      ,'Vous devez saisir les quantités inventoriées.'
          {16}      ,'Voulez vous créer une référence sans code client?'
          {17}      ,'La quantité inventoriée est supérieure à la sommme des quantités réservées et en préparation du client.#13 L''excédent sera disponible pour tous les clients.'
          {18}      ,'La quantité inventoriée ne peut être négative'
                );

const stChampDim : string = 'GA_STATUTART,'
               + 'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,'
               + 'GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5 ';

Implementation

procedure EntreeSaisieInvContrem;
begin
AGLLanceFiche('GC','GCINVCONTREM','','','');
end;

procedure TOF_GCINVCONTREM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCINVCONTREM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCINVCONTREM.OnUpdate ;
var QDC : TQuery;
    stWhere : string;
    iInd : integer;
    TobDC : TOB;
begin
  Inherited ;
TobDC := TobDispoContrem.FindFirst(['(SAISI)'],['X'],false);
if TobDC <> nil then
    if PGIAsk(TraduireMemoire(TexteMessage[8]),TraduireMemoire(Ecran.Caption))<>mrYes then exit;
GInv.VidePile(false) ; TobDispoContrem.ClearDetail ;
TobArticle.ClearDetail ; TobCatalogue.ClearDetail ;
stDepot := GetControlText('GQC_DEPOT');
stWhere := 'WHERE GQC_DEPOT="'+stDepot+'" ';
stWhere := stWhere + GetCondition('GQC','ARTICLE','ARTICLE')
                   + GetCondition('GQC','REFERENCE','REFERENCE')
                   + GetCondition('GQC','EMPLACEMENT','EMPLACEMENT')
                   + GetCondition('GQC','FOURNISSEUR','FOURNISSEUR')
                   + GetCondition('GQC','CLIENT','CLIENT');
QDC := OpenSQL('SELECT * FROM DISPOCONTREM ' + stWhere ,True);
if not QDC.Eof then
   TobDispoContrem.LoadDetailDB('DISPOCONTREM','','',QDC,True);
Ferme(QDC);
for iInd := 0 to TobDispoContrem.Detail.Count -1 do
    begin
    TobDC := TobDispoContrem.Detail[iInd];
    AddLesChampsSup(TobDC);
    InitLesChampsSup(TobDC);
    end;
TobArticle.LoadDetailFromSQL('SELECT DISTINCT(GA_ARTICLE),'+stChampDim
    + 'FROM DISPOCONTREM LEFT JOIN ARTICLE ON GQC_ARTICLE=GA_ARTICLE ' + stWhere);
TobCatalogue.LoadDetailFromSQL('SELECT DISTINCT(GCA_ARTICLE),GCA_REFERENCE,GCA_LIBELLE,GCA_TIERS,GCA_DPA '
    + 'FROM DISPOCONTREM LEFT JOIN CATALOGU ON (GQC_REFERENCE=GCA_REFERENCE AND GQC_FOURNISSEUR=GCA_TIERS) '+ stWhere);
TobDispoContrem.PutGridDetail(GInv,False,False,GInv.Titres[0]);
HMTrad.ResizeGridColumns(GInv);
MajPiedCumul;
RowEnter(1);
end ;

procedure TOF_GCINVCONTREM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCINVCONTREM.OnArgument (S : String ) ;
var iCol : integer;
begin
  Inherited ;
///Initialisation
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
BZoomArticle := TMenuItem(GetControl('mnZoomArticle'));
BZoomMvt := TMenuItem(GetControl('mnZoomMvt'));
bForcer := TMenuItem(GetControl('mnForcer'));
bSaisi := TMenuItem(GetControl('mnSaisi'));
bNonSaisi := TMenuItem(GetControl('mnNonSaisi'));
bNext := TMenuItem(GetControl('mnNext'));
BZoomArticle.OnClick := BZoomArticleClick;
BZoomMvt.OnClick := BZoomMvtClick;
bForcer.OnClick := bForcerClick;
bSaisi.OnClick := bSaisiClick;
bNonSaisi.OnClick := bNonSaisiClick;
bNext.OnClick := bNextClick;
THEdit(GetControl('GQC_REFERENCE')).OnElipsisClick := bRechInCatalogueElipsisClick;
THEdit(GetControl('GQC_REFERENCE_')).OnElipsisClick := bRechInCatalogueElipsisClick;
PCumul := THPanel(GetControl('PCUMUL'));
DPA  := THNumEdit(GetControl('DPASAISI')) ;
DPR  := THNumEdit(GetControl('DPRSAISI')) ;
TotQtePhy   := THNumEdit(GetControl('TOTQTEPHY')) ;
TotQteSai := THNumEdit(GetControl('TOTQTESAI')) ;
TotQteEcart := THNumEdit(GetControl('TOTQTEECART')) ;
bValideInv := TToolbarButton97(GetControl('BVALIDEINV'));
bDetailPrix := TToolbarButton97(GetControl('BDETAILPRIX'));
bSupprimer := TToolbarButton97(GetControl('BSUPPRIMER'));
bInserer := TToolbarButton97(GetControl('BINSERER'));
bSelectAll := TToolbarButton97(GetControl('BSELECTALL'));
bComplement := TToolbarButton97(GetControl('BCOMPLEMENT'));
BChercher := TToolbarButton97(GetControl('BCHERCHER'));
bDetailPrix.OnClick := bDetailPrixClick;
bValideInv.OnClick := bValideInvClick;
bSupprimer.OnClick := bSupprimerClick;
bSelectAll.OnClick := bSelectAllClick;
bComplement.OnClick := bComplementClick;
BChercher.OnClick:=BChercherClick;
FindLigne:=TFindDialog.Create(TFVierge(Ecran));
FindLigne.OnFind:=FindLigneFind;
bInserer.OnClick := bInsererClick;
TDetailPrix := TToolWindow97(GetControl('TDETAILPRIX'));
TDetailPrix.OnClose := TDetailPrixClose;
EventsLogUpdate := TStringList.Create;
EventsLogDelete := TStringList.Create;
EventsLogCreate := TStringList.Create;
SetControlText('GQC_DEPOT',GetParamSoc('SO_GCDEPOTDEFAUT'));
//Grid
GInv := THGrid(GetControl('GINV')) ;
GInv.ListeParam:='GCINVCONTREM';
GInv.OnRowEnter := GInvRowEnter ; GInv.OnRowExit := GInvRowExit;
GInv.OnCellEnter := GInvCellEnter ; GInv.OnCellExit := GInvCellExit;
GInv.OnColumnWidthsChanged := GridColumnWidthsChanged;
EtudieColsListes;
for iCol := 1 to GInv.ColCount-1 do  ///**
    if (iCol <> ColQteSai) and (iCol <> ColDPA) and (iCol <> ColDPR) then GInv.ColLengths[iCol] := -1;
GInv.ColTypes[ColSaisi]:='B' ;
GInv.ColFormats[ColSaisi]:=IntToStr(Ord(csCheckbox));
HMTrad.ResizeGridColumns(GInv);
// Tob
TobDispoContrem := TOB.Create('',nil,-1) ;
TobArticle := TOB.Create('',nil,-1) ; TobCatalogue := TOB.Create('',nil,-1);
GInv.Col := ColQteSai;
//variables
stDepot := GetControlText('GQC_DEPOT');
end ;

procedure TOF_GCINVCONTREM.OnClose ;
begin
if TobDispoContrem.FindFirst(['(SAISI)'],['X'],false) <> nil then
    if PGIAsk(TraduireMemoire(TexteMessage[12]),TraduireMemoire(Ecran.Caption))<>mrYes  then
       begin
       AfficheError:=False;
       LastError:=1;
       end;
  Inherited ;
if LastError<> 0 then AfficheError:=true else ToutLiberer ;
end ;

procedure TOF_GCINVCONTREM.ToutLiberer ;
begin
TobDispoContrem.Free;
TobArticle.Free; TobCatalogue.Free;
EventsLogCreate.Free;
EventsLogUpdate.Free;
EventsLogDelete.Free;
FindLigne.Free;
end;

///GRID
procedure TOF_GCINVCONTREM.EtudieColsListes ;
Var NomCol,LesCols : String ;
    icol  : integer ;
begin
//Grids des lignes de doc
LesCols := GInv.Titres[0];
icol:=1 ;     ///**
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol='GQC_DPA' then ColDPA:=icol else
    if NomCol='GQC_DPR' then ColDPR:=icol else
    if NomCol='GQC_PHYSIQUE' then ColPhy:=icol else
    if NomCol='(QTESAISI)' then ColQteSai:=icol else
    if NomCol='(ECART)' then ColEcart:=icol else
    if NomCol='(SAISI)' then ColSaisi:=icol;
    Inc(icol) ;
Until ((LesCols='') or (NomCol=''));
end;

///TOB
function TOF_GCINVCONTREM.GetCondition(stPref,stInf,stSup : string) : string;
var st,stSQLInf,stSQLSup : string;
begin
st := ''; result := '';
//borne inférieure
if stInf <> '' then st := GetControlText('GQC_'+stInf);
if st <> '' then
   if stSup = '' then stSQLInf := stPref + '_' + stInf + '="' + st + '"'
   else stSQLInf := stPref + '_' + stInf + '>="' + st + '"';
//borne supérieure
if stSup <> '' then st := GetControlText('GQC_'+stSup+'_');
if st <> '' then
   begin
   if stSup = 'ARTICLE' then stSQLSup := stPref + '_' + stSup + '<="' + format('%-18.18s%16.16s',[st,StringOfChar('Z',16)]) + '"'
   else stSQLSup := stPref + '_' + stSup + '<="' + st + '"';
   if stSQLInf <> '' then stSQLSup := ' AND ' + stSQLSup;
   end;

if (stSQLInf = '') and (stSQLSup = '') then exit;
result := ' AND ' + stSQLInf + stSQLSup;
end;

///Ecran
procedure TOF_GCINVCONTREM.ActualisePrix(TobDC : TOB);
begin
if TobDC = nil then
    begin
    SetControlText('DPAART',StrS(0,V_PGI.OkDecV));
    SetControlText('GQC_DPA',StrS(0,V_PGI.OkDecV));
    SetControlText('GQC_DPR',StrS(0,V_PGI.OkDecV));
    SetControlText('DPASAISI',StrS(0,V_PGI.OkDecV));
    SetControlText('DPRSAISI',StrS(0,V_PGI.OkDecV));
    exit;
    end;
SetControlText('DPAART',StrS(GetPrixArticle(TobDC,'DPA'),V_PGI.OkDecV));
SetControlText('GQC_DPA',StrS(TobDC.GetValue('GQC_DPA'),V_PGI.OkDecV));
SetControlText('GQC_DPR',StrS(TobDC.GetValue('GQC_DPR'),V_PGI.OkDecV));
SetControlText('DPASAISI',StrS(TobDC.GetValue('DPASAISI'),V_PGI.OkDecV));
SetControlText('DPRSAISI',StrS(TobDC.GetValue('DPRSAISI'),V_PGI.OkDecV));
end;

function TOF_GCINVCONTREM.GetPrixArticle(TobDC : TOB ; ChampPrix : string) : double;
var TobR : TOB;
begin
Result := 0;
if TobDC = nil then exit;
TobR:=TobCatalogue.FindFirst(['GCA_REFERENCE','GCA_TIERS'],
        [TobDC.GetValue('GQC_REFERENCE'),TobDC.GetValue('GQC_FOURNISSEUR')],False);
if TobR <> nil then
    if TobR.FieldExists('GCA_'+ChampPrix) then Result := TobR.GetValue('GCA_'+ChampPrix);
end;

function TOF_GCINVCONTREM.GetLibelleArticle(TobDC : TOB) : string;
var TobR : TOB;
begin
Result := '';
if TobDC = nil then exit;
TobR:=TobCatalogue.FindFirst(['GCA_REFERENCE','GCA_TIERS'],
     [TobDC.GetValue('GQC_REFERENCE'),TobDC.GetValue('GQC_FOURNISSEUR')],False);
if TobR <> nil then Result := TobR.GetValue('GCA_LIBELLE');
end;

function TOF_GCINVCONTREM.GetDimensions(Article : string) : string;
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
    TobArt : TOB;
begin
result := '' ; StDim:='';
if Article = '' then exit;
TobArt := TobArticle.FindFirst(['GA_ARTICLE'],[Article],False);
if TobArt = nil then exit;
if TobArt.GetValue('GA_STATUTART') <> 'DIM' then exit;
for i_indDim := 1 to MaxDimension do
    begin
    GrilleDim := TOBArt.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
    CodeDim := TOBArt.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
    LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
    if LibDim <> '' then
       if StDim='' then StDim:=StDim + LibDim
       else StDim := StDim + ' - ' + LibDim;
    end;
Result := stDim;
end;

procedure TOF_GCINVCONTREM.TobPutEcart(TobDC : TOB);
begin
TobDC.PutValue('(ECART)',
    Arrondi(TobDC.GetValue('(QTESAISI)')-TobDC.GetValue('GQC_PHYSIQUE'),V_PGI.OkDecQ));
end;

function TOF_GCINVCONTREM.TobCreateContrem(TobMere : TOB) : TOB;
var TobDC : TOB;
begin
TobDC := TOB.Create('DISPOCONTREM',TobMere,-1);
AddLesChampsSup(TobDC);
Result := TobDC;
end;

function TOF_GCINVCONTREM.TobSelectDBContrem(stReq : string) : TOB;
var TobDC : TOB;
    QQ : TQuery;
begin
QQ := OpenSQL(stReq,true);
if not QQ.Eof then
   begin
   TobDC := TobCreateContrem(nil);
   TobDC.SelectDB('',QQ);
   InitLesChampsSup(TobDC);
   Result := TobDC;
   end else Result := nil;
Ferme(QQ);
end;


procedure TOF_GCINVCONTREM.AddLesChampsSup(TobDC : TOB);
begin
TobDC.AddChampSup('(CODEARTICLE)',False);
TobDC.AddChampSupValeur('(ECART)',Arrondi(0,V_PGI.OkDecQ),False);
TobDC.AddChampSupValeur('(QTESAISI)',Arrondi(0,V_PGI.OkDecQ),False);
TobDC.AddChampSupValeur('(QTECLIENT)',Arrondi(0,V_PGI.OkDecQ),False);
TobDC.AddChampSupValeur('(SAISI)','-',False);
TobDC.AddChampSupValeur('DPASAISI',0,False);
TobDC.AddChampSupValeur('DPRSAISI',0,False);
//TobDC.AddChampSupValeur('RESERVECLI', 0, False);
end;

procedure TOF_GCINVCONTREM.InitLesChampsSup(TobDC : TOB);
var stArt : string;
begin
TobDC.PutValue('DPASAISI',TobDC.GetValue('GQC_DPA'));
TobDC.PutValue('DPRSAISI',TobDC.GetValue('GQC_DPR'));
TobDC.PutValue('(QTECLIENT)',TobDC.GetValue('GQC_RESERVECLI')+TobDC.GetValue('GQC_PREPACLI'));
stArt := TobDC.GetValue('GQC_ARTICLE');
if stArt <> '' then TobDC.PutValue('(CODEARTICLE)',Trim(Copy(stArt,1,18)));
TobPutEcart(TobDC);
end;

function TOF_GCINVCONTREM.RechRangDispoContrem : integer;
var QRang : TQuery;
begin
Result := 1;
QRang := OpenSQL('SELECT MAX(GQC_RANG) AS MAXRANG FROM DISPOCONTREM',True);
if not QRang.Eof then Result := QRang.FindField('MAXRANG').AsInteger+1;
Ferme(QRang);
end;

procedure TOF_GCINVCONTREM.GInvRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
RowEnter(Ou);
end;

procedure TOF_GCINVCONTREM.GInvRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if TDetailPrix.Visible then MajTobPrixSaisi(Ou);
end;

procedure TOF_GCINVCONTREM.GInvCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
CellText := GInv.Cells[GInv.Col,GInv.Row];
end;

procedure TOF_GCINVCONTREM.GInvCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if CellText = GInv.Cells[ACol,ARow] then exit;
if not ValideLaCellule(ACol,ARow) then
   begin
   GInv.Col := ACol; GInv.Row := ARow;
   end;
end;

procedure TOF_GCINVCONTREM.GridColumnWidthsChanged(Sender: TObject);
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

procedure TOF_GCINVCONTREM.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_RETURN : if Screen.ActiveControl=GInv then Key:=VK_DOWN ;
   VK_DELETE : begin
               if ((Screen.ActiveControl=GInv) and (Shift=[ssCtrl])) then
                   begin Key := 0 ; SupprimeLaLigne (GInv.Row) ; end;
               end ;
   VK_INSERT : begin
               if ((Screen.ActiveControl=GInv) and (Shift=[ssCtrl])) then
                   begin Key:=0 ; AjouteUneReference ; end ;
               end;
   VK_F10    : begin Key:=0 ; bValideInvClick(nil) ; end ;
   VK_F5     : if (Screen.ActiveControl=GInv) then begin Key:=0 ; BZoomArticleClick(nil); end;
  {Alt+S} 83 : if Shift=[ssAlt]  then begin Key:=0 ; bSaisiClick(nil); end ;
   end;
end;

procedure TOF_GCINVCONTREM.TDetailPrixClose(Sender: TObject);
begin
bDetailPrix.Down := TDetailPrix.Visible;
MajTobPrixSaisi(GInv.Row);
end;


function TOF_GCINVCONTREM.ValideLaCellule(ACol,ARow : integer) : boolean;
var TobDC : TOB;
    ValCell : double;
begin
Result := True;
TobDC := TOB(GInv.Objects[0,ARow]);
if TobDC = nil then exit;
ValCell := Valeur(GInv.Cells[ACol,ARow]);
if ACol = ColQteSai then
    begin
    GInv.Cells[ACol,ARow] := StrS(ValCell,V_PGI.OkDecQ);
    if not ValideLaQte(TobDC,ValCell) then begin Result := False ; exit ; end;
    if (TobDC.GetValue('(SAISI)')='-') and (ValCell <> 0) then
       begin
       GInv.Cells[ColSaisi,ARow] := 'X';
       TobDC.PutValue('(SAISI)','X');
       end;
    GInv.Cells[ColEcart,ARow] := Strs(ValCell-TobDC.GetValue('GQC_PHYSIQUE'),V_PGI.OkDecQ);
    TobDC.PutValue('(QTESAISI)',ValCell);
    TobDC.PutValue('(ECART)',Valeur(GInv.Cells[ColEcart,ARow]));
    end else
 if Acol = ColDPA then
       begin
       TobDC.PutValue('DPASAISI',ValCell) ;
       GInv.Cells[ACol,ARow] := StrS(ValCell,V_PGI.OkDecV);
       end
else if Acol = ColDPR then
       begin
       TobDC.PutValue('DPRSAISI',ValCell);
       GInv.Cells[ACol,ARow] := StrS(ValCell,V_PGI.OkDecV);
       end;
MajPiedCumul;
end;

function TOF_GCINVCONTREM.ValideLaQte (TobDC : TOB ; QteSaisie : double) : boolean;
begin
Result := True;
if TobDC = nil then exit;
if QteSaisie < 0 then
    begin
    PGIBox(TraduireMemoire(TexteMessage[18]),TraduireMemoire(Ecran.Caption)) ;
    Result := False;
    exit;
    end;
if TobDC.GetValue('GQC_CLIENT') = '' then exit;
if QteSaisie > TobDC.GetValue('(QTECLIENT)') then
    PGIBox(TraduireMemoire(TexteMessage[17]),TraduireMemoire(Ecran.Caption)) ;
end;

procedure TOF_GCINVCONTREM.RowEnter(ARow : integer);
var TobDC : TOB;
    stInfo,stDim : string;
begin
TobDC := TOB(GInv.Objects[0,ARow]);
if TToolWindow97(GetControl('TDETAILPRIX')).Visible then ActualisePrix(TobDC);
if TobDC = nil then
   begin
   SetControlVisible('PINFO',False);
   exit;
   end;
stInfo := '';
if TobDC.GetValue('GQC_EMPLACEMENT') <> '' then
   stInfo := 'Emplacement : ' + RechDom('GCEMPLACEMENT',TobDC.GetValue('GQC_EMPLACEMENT'),False);
stDim := GetDimensions(TobDC.GetValue('GQC_ARTICLE'));
if stDim <> '' then
   begin
   stDim := 'Dimensions article : ' + stDim;
   if stInfo <> '' then stInFo := stInfo + '  -  ' + stDim
   else stInFo := stDim;
   end;
SetControlText('TINFO',stInfo);
SetControlVisible('PINFO',(stInfo <> ''));
end;

procedure TOF_GCINVCONTREM.bForcerClick(Sender: TObject);
var iInd : integer;
begin
if TobDispoContrem.Detail.Count = 0 then exit;
for iInd := 0 to TobDispoContrem.Detail.Count -1 do
    TobDispoContrem.Detail[iInd].PutValue('(SAISI)','X');
TobDispoContrem.PutGridDetail(GInv,False,False,GInv.Titres[0]);
end;

procedure TOF_GCINVCONTREM.bSaisiClick(Sender: TObject);
var TobDC : TOB;
    ARow : integer;
begin
ARow := GInv.Row;
TobDC := TOB(GInv.Objects[0,ARow]);
if TobDC = nil then exit;
TobDC.PutValue('(SAISI)','X');
GInv.Cells[ColSaisi,ARow] := 'X';
end;

procedure TOF_GCINVCONTREM.bNonSaisiClick(Sender: TObject);
var TobDC : TOB;
    ARow : integer;
begin
ARow := GInv.Row;
TobDC := TOB(GInv.Objects[0,ARow]);
if TobDC = nil then exit;
if TobDC.GetValue('(QTESAISI)') <> 0 then exit;
TobDC.PutValue('(SAISI)','-');
GInv.Cells[ColSaisi,ARow] := '-';
end;

procedure TOF_GCINVCONTREM.bNextClick(Sender: TObject);
var iRow : integer;
    TobDC : TOB;
begin
for iRow := 1 to GInv.RowCount -1 do
  begin
  TobDC := TOB(GInv.Objects[0,iRow]);
  if TobDC = nil then continue;
  if TobDC.GetValue('(SAISI)') <> 'X' then
    begin
    if ValideLaCellule(GInv.Col,GInv.Row) then GInv.Row := iRow;
    break;
    end;
  end;

end;

procedure TOF_GCINVCONTREM.bSupprimerClick(Sender: TObject);
begin
SupprimeLaLigne(GInv.Row);
end;

procedure TOF_GCINVCONTREM.bInsererClick(Sender: TObject);
begin
AjouteUneReference;
end;

procedure TOF_GCINVCONTREM.bSelectAllClick(Sender: TObject);
var iInd : integer;
    TobDC : TOB;
begin
if TobDispoContrem.Detail.Count = 0 then
   begin
   bSelectAll.Down := (not bSelectAll.Down);
   exit;
   end;
if bSelectAll.Down then
  begin
  if PGIAsk(TraduireMemoire(TexteMessage[1]),TraduireMemoire(Ecran.Caption))<>mrYes  then
     begin bSelectAll.Down:=False ; exit ; end;
  for iInd := 0 to TobDispoContrem.Detail.Count -1 do
      begin
      TobDC := TobDispoContrem.Detail[iInd];
      if TobDC.GetValue('(SAISI)') = 'X' then continue;
      TobDC.PutValue('(QTESAISI)',TobDC.GetValue('GQC_PHYSIQUE'));
      TobPutEcart(TobDC);
      TobDC.PutValue('(SAISI)','X');
      end;
  end else
  begin
  if PGIAsk(TraduireMemoire(TexteMessage[4]),TraduireMemoire(Ecran.Caption))<>mrYes then
     begin bSelectAll.Down:=True ; exit ; end;
  for iInd := 0 to TobDispoContrem.Detail.Count -1 do
      begin
      TobDC := TobDispoContrem.Detail[iInd];
      if TobDC.GetValue('(SAISI)') <> 'X' then continue;
      TobDC.PutValue('(QTESAISI)',Arrondi(0,V_PGI.OkDecQ));
      TobPutEcart(TobDC);
      TobDC.PutValue('(SAISI)','-');
      end;
  end;
TobDispoContrem.PutGridDetail(GInv,False,False,GInv.Titres[0]);
MajPiedCumul;
end;

procedure TOF_GCINVCONTREM.bComplementClick(Sender: TObject);
var ARow : integer;
    LigneVide : boolean;
begin
ARow := GInv.Row;
LigneVide := (TOB(GInv.Objects[0,ARow]) = nil);
bNonSaisi.Visible := (not LigneVide) and (Valeur(GInv.Cells[ColQteSai,ARow]) = 0) and (GInv.Cells[ColSaisi,ARow] = 'X');
bSaisi.Visible := (not LigneVide) and (GInv.Cells[ColSaisi,ARow] = '-');
end;

procedure TOF_GCINVCONTREM.bDetailPrixClick(Sender: TObject);
var TobDC : TOB;
begin
TDetailPrix.Visible := not(TDetailPrix.Visible);
if TDetailPrix.Visible then
   begin
   TobDC := TOB(GInv.Objects[0,GInv.Row]);
   ActualisePrix(TobDC);
   end;
end;

procedure TOF_GCINVCONTREM.bValideInvClick(Sender: TObject);
var TobDC,TOBJnal : TOB;
    JNalCode,MsgCode,stCodeInv,stLibInv : string;
    NumEvt : integer;
    io : TIOErr ;
    QQ : TQuery;
    EventsLog : TStringList;
begin
if TobDispoContrem.Detail.Count = 0 then exit;
if not ValideLaCellule(GInv.Col,GInv.Row) then exit;
TobDC := TobDispoContrem.FindFirst(['(SAISI)'],['X'],false);
if TobDC = nil then //Aucune saisie
   begin
   PGIBox(TraduireMemoire(TexteMessage[15]),TraduireMemoire(Ecran.Caption)) ;
   exit ;
   end;
stCodeInv := Trim(GetControlText('CODELISTE'));
if stCodeInv = '' then
   begin
   PGIBox(TraduireMemoire(TexteMessage[7]),TraduireMemoire(Ecran.Caption)) ;
   SetFocusControl('CODELISTE') ;
   exit ;
   end;
stLibInv :=Trim(GetControlText('LIBELLECODELISTE'));
if stLibInv = '' then
   begin
   PGIBox(TraduireMemoire(TexteMessage[11]),TraduireMemoire(Ecran.Caption)) ;
   SetFocusControl('LIBELLECODELISTE') ;
   exit ;
   end;
TobDC := TobDispoContrem.FindFirst(['(SAISI)'],['-'],false);
if TobDC <> nil then
   begin
   if PGIAsk(TraduireMemoire(TexteMessage[2]),TraduireMemoire(Ecran.Caption))<>mrYes then exit;
   end else if PGIAsk(TraduireMemoire(TexteMessage[3]),TraduireMemoire(Ecran.Caption))<>mrYes then exit;
if TDetailPrix.Visible then
   begin
   MajTobPrixSaisi(GInv.Row);
   TDetailPrix.Visible := False;
   bDetailPrix.Down := False;
   end;

io:=Transactions(UpdateDispocontrem,5) ;
JNalCode := 'OK';
EventsLog := TStringList.Create;
Case io of
    oeOk       : begin
                 JNalCode := 'OK';
                 MsgCode := TexteMessage[13];
                 PGIBox(MsgCode,TraduireMemoire(Ecran.Caption));
                 end;
    oeUnknown  : begin MessageAlerte(TexteMessage[9]) ; JNalCode := 'ERR';  MsgCode := TexteMessage[9] ;  end ;
    oeSaisie   : begin MessageAlerte(TexteMessage[9]) ; JNalCode := 'ERR';  MsgCode := TexteMessage[9] ;  end ;
    end ;
// Journal des evenements --------
EventsLog.Add(stCodeInv + ' ' + stLibInv + ' - ' + RechDom('GCDEPOT',stDepot,False));

if JNalCode = 'ERR' then
   begin
   EventsLog.Add(TexteMessage[10] + ' : ' + MsgCode);
   end else
   begin
   EventsLog.Add(TexteMessage[6]);
   EventsLog.Add('');
   EventsLog.Text := EventsLog.Text + EventsLogCreate.Text + EventsLogUpdate.Text + EventsLogDelete.Text;
   end;
NumEvt := 0;
TOBJnal := TOB.Create('JNALEVENT', nil, -1);
TOBJnal.PutValue('GEV_TYPEEVENT', 'INV');
TOBJnal.PutValue('GEV_LIBELLE', Ecran.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', true);
if not QQ.EOF then NumEvt := QQ.Fields[0].AsInteger;
Ferme(QQ);
inc(NumEvt);
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
TOBJnal.PutValue('GEV_ETATEVENT', JNalCode);
TOBJnal.PutValue('GEV_BLOCNOTE', EventsLog.Text);
TOBJnal.InsertDB(nil);
TOBJnal.Free;
if io <> oeOk then exit;
SetFocusControl('GQC_ARTICLE');
GInv.VidePile(false);
GInv.Refresh;
GInv.Col := ColQteSai;
TobDispoContrem.ClearDetail;
SetControlText('CODELISTE','');SetControlText('LIBELLECODELISTE','');
MajPiedCumul;
end;

procedure TOF_GCINVCONTREM.BZoomArticleClick(Sender: TObject);
var TobDC : TOB;
begin
TobDC := TOB(GInv.Objects[0,GInv.Row]);
if TobDC = nil then exit;
AGLLanceFiche('GC','GCCATALOGU_SAISI3','',TobDC.GetValue('GQC_REFERENCE')
          + ';' + TobDC.GetValue('GQC_FOURNISSEUR'),'ACTION=CONSULTATION;MONOFICHE');
end;

procedure TOF_GCINVCONTREM.BZoomMvtClick(Sender: TObject);
var TobDC : TOB;
    stWhere : string;
begin
TobDC := TOB(GInv.Objects[0,GInv.Row]);
if TobDC = nil then exit;
stWhere := ' ((GL_TIERS="' + TobDC.GetValue('GQC_CLIENT') + '" '
           + 'AND GL_FOURNISSEUR="' + TobDC.GetValue('GQC_FOURNISSEUR') + '") '
           + 'OR (GL_TIERS="' + TobDC.GetValue('GQC_FOURNISSEUR') + '" '
           + 'AND GL_FOURNISSEUR="' + TobDC.GetValue('GQC_CLIENT') + '")) '
           + 'AND GL_REFCATALOGUE="' + TobDC.GetValue('GQC_REFERENCE') + '" '
           + 'AND GL_DEPOT="' + TobDC.GetValue('GQC_DEPOT') + '" '
           + 'AND GL_VIVANTE="X" AND GL_QTERESTE > 0'; {DBR NEWPIECE QTERESTE > 0}
AGLLanceFiche('GC','GCINVCONLIG_MUL','','',stWhere);
end;

procedure TOF_GCINVCONTREM.bRechInCatalogueElipsisClick(Sender: TObject);
var RefCata,ControlName : string;
begin
ControlName := THEdit(Sender).Name;
RefCata := GetControlText(ControlName);
RefCata:=AGLLanceFiche('GC','GCMULCATALOG_RECH','','','GCA_REFERENCE='+Trim(RefCata)) ;
RefCata:=uppercase(Trim(ReadTokenSt(RefCata)));
if RefCata<>'' then SetControlText(ControlName,RefCata);
end;

procedure TOF_GCINVCONTREM.UpdateDispocontrem;
var TobDC,TobR,TobNew,TobN,TobModif,TobDelete : TOB;
    iInd,iRang : integer;
    Reste,QteInv,QteMax : double;
begin
ValideLaListe;
if V_PGI.IoError <> oeOk then exit;
TobNew := TOB.Create('',nil,-1); TobDelete := TOB.Create('',nil,-1);
EventsLogCreate.Clear; EventsLogUpdate.Clear; EventsLogDelete.Clear;
iRang := RechRangDispoContrem;
for iInd := TobDispoContrem.Detail.Count -1 downto 0 do
    begin
    TobDC := TobDispoContrem.Detail[iInd];
    if TobDC.GetValue('(SAISI)') = '-' then begin TobDC.Free ; Continue ; end;
    QteInv := TobDC.GetValue('(QTESAISI)');
    QteMax := TobDC.GetValue('(QTECLIENT)');
    //si toutes les qtés sont à 0 => suppression de l'enreg dans table DispoContrem
    if (QteInv = 0) and (TobDC.GetValue('GQC_RESERVECLI') = 0)
       and (TobDC.GetValue('GQC_RESERVEFOU') = 0) and (TobDC.GetValue('GQC_PREPACLI') = 0) then
        begin
        TobDC.ChangeParent(TobDelete,-1);
        continue;
        end;
    TobDC.PutValue('GQC_DPA',TobDC.GetValue('DPASAISI'));
    TobDC.PutValue('GQC_DPR',TobDC.GetValue('DPRSAISI'));
    if QteInv > QteMax then
        begin
        if TobDC.GetValue('GQC_PHYSIQUE') < QteMax then TobDC.PutValue('GQC_PHYSIQUE',QteMax);
        Reste := QteInv - TobDC.GetValue('GQC_PHYSIQUE');
        TobR := TobNew.FindFirst(['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],
             [stDepot,TobDC.GetValue('GQC_REFERENCE'),TobDC.GetValue('GQC_FOURNISSEUR'),''],False);
        if TobR <> nil then TobR.PutValue('GQC_PHYSIQUE',TobR.GetValue('GQC_PHYSIQUE')+Reste)
        else begin
             TobN := TobCreateContrem(TobNew);
             TobN.Dupliquer(TobDC,False,True);
             TobModif := TobN;
             with TobModif do
               begin
               PutValue('GQC_CLIENT','');
               PutValue('GQC_PHYSIQUE',Reste);
               PutValue('GQC_RESERVECLI',0) ; PutValue('GQC_RESERVEFOU',0);
               PutValue('GQC_PREPACLI',0);
               PutValue('GQC_RANG',iRang);
               end;
             inc(iRang);
             end;
        end else TobDC.PutValue('GQC_PHYSIQUE',TobDC.GetValue('(QTESAISI)'));
    end;
//faire le Delete avant le TobNew.insert
if TobDelete.Detail.Count > 0 then
   begin
   EventsLogDelete.Add(intToStr(TobDelete.Detail.Count) + ' référence(s) supprimée(s)');
   TobDelete.DeleteDB(True);
   end;
ControleExistence(TobNew); //Vérif de l'existence des réf dans la tob dispocontrem ou dans la table
if TobNew.Detail.Count > 0 then
   begin
   EventsLogCreate.Add(intToStr(TobNew.Detail.Count) + ' référence(s) créée(s)');
   TobNew.InsertOrUpdateDB(True);
   end;
if TobDispoContrem.Detail.Count > 0 then
   begin
   EventsLogUpdate.Add(intToStr(TobDispoContrem.Detail.Count) + ' référence(s) modifiée(s)');
   TobDispoContrem.InsertOrUpdateDB(True);
   end;
TobNew.Free;
TobDelete.Free;
end;

procedure TOF_GCINVCONTREM.ValideLaListe;
var TobListeInv,TobDC,TobLigneInv : TOB;
    stCodeListe{,stEmplacement} : string;
    iInd : integer;
begin
//Entête
stCodeListe := GetCodeListe(GetControlText('CODELISTE'));
//stEmplacement := GetParamSoc('SO_GCEMPLACEMENTDEFAUT');
TobListeInv := TOB.Create('LISTEINVENT',nil,-1);
with TobListeInv do
    begin
    InitValeurs;
    PutValue('GIE_CODELISTE',stCodeListe);
    PutValue('GIE_LIBELLE',GetControlText('LIBELLECODELISTE'));
    PutValue('GIE_NATUREPIECEG','INV');
    PutValue('GIE_DEPOT', stDepot);
    PutValue('GIE_DATEINVENTAIRE', V_PGI.DateEntree);
    PutValue('GIE_DATECREATION', Date);
    PutValue('GIE_DATEMODIF', Date);
    PutValue('GIE_CREATEUR', V_PGI.User);
    PutValue('GIE_UTILISATEUR', V_PGI.User);
    PutValue('GIE_VALIDATION', 'X');
    PutValue('GIE_CONTREMARQUE','X');
    end;
//Lignes
for iInd := 0 to TobDispoContrem.Detail.Count -1 do
    begin
    TobDC := TobDispoContrem.Detail[iInd];
    if TobDC.GetValue('(SAISI)') = '-' then Continue;
    TobLigneInv := TOB.Create('LISTEINVLIGCONTREM',TobListeInv,-1);
    with TobLigneInv do
        begin
        InitValeurs;
        PutValue('GIM_CODELISTE', stCodeListe);
        PutValue('GIM_ARTICLE',TobDC.GetValue('GQC_ARTICLE'));
        PutValue('GIM_DEPOT',STDepot);
        PutValue('GIM_SAISIINV','X');
        PutValue('GIM_INVENTAIRE',TobDC.GetValue('(QTESAISI)'));
        PutValue('GIM_QTEPHOTOINV',TobDC.GetValue('GQC_PHYSIQUE'));
        //PutValue('GIM_EMPLACEMENT',stEmplacement);
        PutValue('GIM_DPA',TobDC.GetValue('GQC_DPA'));
        PutValue('GIM_DPR',TobDC.GetValue('GQC_DPR'));
        PutValue('GIM_DPAART',GetPrixArticle(TobDC,'DPA'));
        PutValue('GIM_DPASAIS',TobDC.GetValue('DPASAISI'));
        PutValue('GIM_DPRSAIS',TobDC.GetValue('DPRSAISI'));
        PutValue('GIM_FOURNISSEUR',TobDC.GetValue('GQC_FOURNISSEUR'));
        if TobDC.GetValue('GQC_CLIENT') = '' then PutValue('GIM_CLIENT','NON_DEFINI')
        else PutValue('GIM_CLIENT',TobDC.GetValue('GQC_CLIENT'));
        PutValue('GIM_REFERENCE',TobDC.GetValue('GQC_REFERENCE'));
        PutValue('GIM_LIBELLE',GetLibelleArticle(TobDC));
        end;
    end;
TobListeInv.InsertDB(nil,True);
TobListeInv.Free;
end;

procedure TOF_GCINVCONTREM.MajPiedCumul;
begin
TotQtePhy.Value := TobDispoContrem.Somme('GQC_PHYSIQUE',[''],[''],False);
TotQteSai.Value := TobDispoContrem.Somme('(QTESAISI)',[''],[''],False);
TotQteEcart.Value := TobDispoContrem.Somme('(ECART)',[''],[''],False);
PCumul.Caption := '     Total   ('+intToStr(TobDispoContrem.detail.Count)+' références)';
end;

procedure TOF_GCINVCONTREM.SupprimeLaLigne(ARow : integer);
var TobDC : TOB;
begin
if TobDispoContrem.Detail.Count = 0 then exit;
if PGIAsk(TraduireMemoire(TexteMessage[14]),TraduireMemoire(Ecran.Caption))<>mrYes then exit;
TobDC := TOB(GInv.Objects[0,ARow]);
if TobDC <> nil then TobDC.free;
//réf volontairement laissé dans tob cata et article
with GInv do
    begin
    if RowCount = 2 then RowCount := 3;
    CacheEdit ; SynEnabled := False;
    DeleteRow (ARow);
    MontreEdit; SynEnabled := True;
    end;
MajPiedCumul;
if not TDetailPrix.Visible then exit;
TobDC := TOB(GInv.Objects[0,ARow]);
if TobDC <> nil then ActualisePrix(TobDC)
else begin TDetailPrix.Visible := False ; bDetailPrix.Down := False; end;
end;

procedure TOF_GCINVCONTREM.AjouteUneReference;
var st,stSQL,stRef,stFour,stArt : string;
    Retour,ChampMul,ValMul,RefCli : string;
    ipos : integer;
    okAjout : boolean;
    TobTemp,TobT,TobDC,TobR,TobA,TobC : TOB;
    QQ : TQuery;
begin
TobDC := nil; okAjout := True;
st := AglLanceFiche ('GC','GCMULCATALOG_RECH','','','');
stRef := Trim(ReadTokenSt(st)); stFour := Trim(ReadTokenSt(st)) ; stArt := Trim(ReadTokenSt(st));
if (stRef = '') or (stFour = '') then exit;
stSQL := 'WHERE GQC_DEPOT="'+stDepot+'" '
       + 'AND GQC_REFERENCE="'+stRef+'" '
       + 'AND GQC_FOURNISSEUR="'+stFour+'"';
TobTemp := TOB.Create('',nil,-1);
QQ := OpenSQL('SELECT * FROM DISPOCONTREM ' + stSQL,True);
if not QQ.Eof then TobTemp.LoadDetailDB('DISPOCONTREM','','',QQ,True);
Ferme(QQ);
if TobTemp.Detail.Count > 0 then
    begin
    Retour := AglLanceFiche ('GC','GCDISPOCONTR_MUL','GQC_DEPOT='+stDepot
                +';GQC_FOURNISSEUR='+stFour+';GQC_REFERENCE='+stRef,'','DISABLED_DEPOT');
    if Retour <> '' then  //Validation d'une réf de GCDISPOCONTR_MUL
        begin
        Repeat
           st:=Trim(ReadTokenSt(Retour));
           if st<>'' then
              begin
              ipos:=pos('=',st);
              if ipos<>0 then
                 begin
                 ChampMul:=uppercase(copy(st,1,ipos-1));
                 ValMul:=copy(st,ipos+1,length(st));
                 if ChampMul='GQC_CLIENT' then RefCli := ValMul;
                 end;
              end;
        until Retour='';
        TobT := TobTemp.FindFirst(['GQC_CLIENT'],[RefCli],False);
        if TobT <> nil then
            begin
            TobDC := TOB.Create('',nil,-1);
            TobDC.Dupliquer(TobT,False,True);
            AddLesChampsSup(TobDC);
            InitLesChampsSup(TobDC);
            end else okAjout := False; // pb
        end else      //sortie sans select de réf  : possibilité de créer une nlle réf sans code client si pas déjà existante dans DispoContrem
        begin
        if TobTemp.FindFirst(['GQC_CLIENT'],[''],False) = nil then
            begin
            if PGIAsk(TraduireMemoire(TexteMessage[16]),TraduireMemoire(Ecran.Caption))<>mrYes then okAjout := False;
            end else okAjout := False;
        end;
    end;
TobTemp.Free;
if not okAjout then exit;
if TobDC = nil then //ajout de la réf du catalogue
   begin
   TobDC := TobCreateContrem(nil);
   with TobDC do
       begin
       //InitValeurs;
       PutValue('GQC_RANG',RechRangDispoContrem);
       PutValue('GQC_DEPOT',stDepot);
       PutValue('GQC_REFERENCE',stRef);
       PutValue('GQC_FOURNISSEUR',stFour);
       PutValue('GQC_ARTICLE',stArt);
       PutValue('(CODEARTICLE)',Trim(Copy(stArt,1,18)));
       //PutValue('GQC_EMPLACEMENT',GetParamSoc('SO_GCEMPLACEMENTDEFAUT'));
       end;
   end;

///recherche existence enreg dans tobs
/// ... dispocontrem
TobR := TobDispoContrem.FindFirst(['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],
     [stDepot,stRef,stFour,TobDC.GetValue('GQC_CLIENT')],False);
if TobR <> nil then
   begin
   PGIBox(TraduireMemoire(TexteMessage[5]),TraduireMemoire(Ecran.Caption));
   TobDC.Free ;
   exit;
   end else TobDC.ChangeParent(TobDispoContrem,-1);
/// ... article
TobR := TobArticle.FindFirst(['GA_ARTICLE'],[stArt],False);
if TobR = nil then
   begin
   QQ := OpenSQL('SELECT GA_ARTICLE,'+stChampDim//+',GA_DPA,GA_DPR '
                             + 'FROM ARTICLE WHERE GA_ARTICLE="' + stArt + '"',True);
   if not QQ.Eof then
       begin
       TobA := TOB.Create('',TobArticle,-1);
       TobA.SelectDB('',QQ);
       end;
   Ferme(QQ);
   end;
/// ... catalogue
TobR := TobCatalogue.FindFirst(['GCA_REFERENCE','GCA_TIERS'],[stRef,stFour],False);
if TobR = nil then
   begin
   QQ := OpenSQL('SELECT GCA_ARTICLE,GCA_REFERENCE,GCA_LIBELLE,GCA_TIERS,GCA_DPA '
              + 'FROM CATALOGU WHERE GCA_REFERENCE="'+ stRef +'" AND GCA_TIERS="'+ stFour +'"',True);
   if not QQ.Eof then
       begin
       TobC := TOB.Create('',TobCatalogue,-1);
       TobC.SelectDB('',QQ);
       end;
   Ferme(QQ);
   end;

with GInv do
  begin
  if RowCount >= 2 then
     begin
     if TOB(Objects[0,RowCount-1]) <> nil then
        begin
        CacheEdit; SynEnabled := false;
        InsertRow(RowCount);
        MontreEdit; SynEnabled := true;
        end;
     end;
  TobDC.PutLigneGrid(GInv, RowCount-1, false, false, Titres[0]);
  end;
MajPiedCumul;
end;

function TOF_GCINVCONTREM.GetCodeListe(RacineCode : string) : string;
var NextCode : integer;
begin
NextCode := 1;
while ExisteSQL('SELECT GIE_CODELISTE FROM LISTEINVENT WHERE GIE_CODELISTE="'
      + RacineCode + Format('%.6d',[NextCode]) + '"') do inc(NextCode);
Result := RacineCode + Format('%.6d',[NextCode]);
end;

procedure TOF_GCINVCONTREM.MajTobPrixSaisi(ARow : integer);
var TobDC : TOB;
    dDPA,dDPR : double;
begin
TobDC := TOB(GInv.Objects[0,ARow]);
if TobDC = nil then exit;
dDPA := Valeur(GetControlText('DPASAISI'));
dDPR := Valeur(GetControlText('DPRSAISI'));
TobDC.PutValue('DPASAISI',dDPA) ;
TobDC.PutValue('DPRSAISI',dDPR) ;
GInv.Cells[ColDPA,ARow] := strs(dDPA,V_PGI.OkDecV);
GInv.Cells[ColDPR,ARow] := strs(dDPR,V_PGI.OkDecV);
end;

procedure TOF_GCINVCONTREM.ControleExistence(TobNew : TOB);
var iInd : integer;
    TobND,TobR,TobDC : TOB;
    stSQL,stRef,stFour : string;
    dQte : double;
begin
for iInd := TobNew.Detail.Count -1 downto 0 do
    begin
    TobND := TobNew.Detail[iInd];
    stRef :=  TobND.GetValue('GQC_REFERENCE');
    stFour := TobND.GetValue('GQC_FOURNISSEUR');
    dQte :=   TobND.getValue('GQC_PHYSIQUE');
    //Controle existence dans Tob DispoContrem
    TobR := TobDispoContrem.FindFirst(['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],
                                 [stDepot,stRef,stFour,''],False);
    if TobR <> nil then
        begin
        TobR.PutValue('GQC_PHYSIQUE',TobR.GetValue('GQC_PHYSIQUE')+dQte);
        TobND.Free;
        continue;
        end else  //sinon recherche de l'existence dans la table
        begin
        stSQL := 'SELECT * FROM DISPOCONTREM WHERE GQC_DEPOT="'+ stDepot +'" '
            +'AND GQC_REFERENCE="'+ stRef + '" AND GQC_FOURNISSEUR="'+ stFour +'" '
            +'AND GQC_CLIENT=""';
        TobDC := TobSelectDBContrem(stSQL);
        if TobDC <> nil then
           begin
           TobDC.PutValue('GQC_PHYSIQUE',dQte+TobDC.GetValue('GQC_PHYSIQUE'));
           TobDC.ChangeParent(TobDispoContrem,-1);
           TobND.Free;
           continue;
           end;
        end;
    end;
end;

// Recherche
procedure TOF_GCINVCONTREM.BChercherClick(Sender: TObject);
begin
if GInv.RowCount < 3 then Exit;
FFirstFind := true;
FindLigne.Execute;
end;

procedure TOF_GCINVCONTREM.FindLigneFind(Sender: TObject);
begin
Rechercher(GInv, FindLigne, FFirstFind);
end;

Initialization
  registerclasses ( [ TOF_GCINVCONTREM ] ) ;
end.


