{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCMOUVSTKEXCONTR ()
                 Mouvements exceptionnels de stock de contremarque
Mots clefs ... : TOF;GCMOUVSTKEXCONTR
*****************************************************************}
Unit MouvStkExContr_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,Menus,Windows,
     sysutils,Ent1,Messages,HCtrls,HSysMenu,Hpanel,
     ComCtrls, ParamSoc,EntGC,
     HEnt1,SaisUtil,HMsgBox,Vierge,Ed_Tools,HTB97,UTOF,UTOB,
     FactUtil,FactComm,FactGrp,Facture,StockUtil,FactPiece, FactTOB,uEntCommun ;



function GCLanceFiche_MouvStkExContr(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_GCMOUVSTKEXCONTR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    GDisp, GInfosLigne : THGrid;
    bZoomReference,bZoomDispo : TMenuItem;
    // CBEmplacement : THValCombobox;
    bValideMouv,bSupprimer,bInserer : TToolbarButton97;
    HMTrad: THSystemMenu;
    TobDispoContrem,TobArticle,TobCatalogue : TOB;
    TobPieceGen : TOB;
    TobEEX,TobSEX : TOB;
    ColSaisi,ColSens,ColMotif : integer;
    Depot, CellText : string;
    PPInfosLigneEEX,PPInfosLigneSEX : TStrings ;
    MsgJnalEventEEX,MsgJnalEventSEX : TStringList;
    //Grid
    procedure GDispRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GDispRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GDispCellExit(Sender: TObject; var ACol,ARow : Integer; var Cancel : Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    //Boutons
    procedure bSupprimerClick(Sender: TObject);
    procedure bInsererClick(Sender: TObject);
    procedure bValideMouvClick(Sender: TObject);
    procedure BZoomReferenceClick(Sender: TObject);
    procedure bZoomDispoClick(Sender: TObject);
    procedure bRechInCatalogueElipsisClick(Sender: TObject);
    //Initialisations
    procedure ToutAllouer ;
    procedure ToutLiberer ;
    procedure AddLesSupDispo(TobDC : TOB);
    procedure InitLesSupDispo(TobDC : TOB);
    //Form
    procedure ActualiseLePied(TobDC : TOB);
    function  GetCondition(stInf,stSup : string) : string;
    //Fonctions liées au grid
    procedure RowEnter(ARow : integer);
    procedure EtudieColsListes ;
    function  ValideLaSaisie (ARow : integer) : boolean;
    procedure SupprimeLaLigne(ARow : integer);
    procedure AjouteUneReference;
    procedure TraiteLaSaisie;
    Function  SortDeLaLigne : boolean ;
    //Pièces
    function  AssocieCatalogueDispoContreM(TobDispoCM : TOB) : TOB;
    procedure GenereLesPieces(TobSource,TobES : TOB ; JNal : TStringList ; NaturePiece : string);
    procedure AssocieLignePiece(TobPiece : TOB);
    function  CreerTobPiece(Nature,CodeTiers : string) : TOB;
    procedure CreerTobLigne(TobPiece,TobDC : TOB ; NumLigne : integer);
    function  GetTobLignePieceGen(TobDC : TOB) : TOB;
    procedure EpureLesLignes;
    procedure CodesCataToCodesLigne(TOBL, TOBCata, TOBArt: TOB);
    function  GenereCompteRendu : string;
    procedure AddEvent(MsgJnalEvent : TStringList ; TobPiece : TOB);
    procedure MAJJnalEvent(MsgJnalEvent : TStringList ; MsgCode : string);

  end ;

const
	// libellés des messages
	TexteMessage: array[1..14] of string 	= (
          {1}        'La quantité en stock est supérieure à la sommme des quantités réservées et en préparation du client.#13 L''excédent sera disponible pour tous les clients.'
          {2}       ,'Vous ne pouvez pas enregistrer une pièce sans mouvements'
          {3}       ,'Cet article est en rupture'
          {4}       ,'ATTENTION : Le stock n''a pu être mis à jour. Le traitement a été annulé.#13'
          {5}       ,'Problèmes rencontrés en validation de pièces.#13Le traitement a été annulé.'
          {6}       ,'Cette référence existe déjà dans la liste'
          {7}       ,'Les quantités négatives ne sont pas autorisées'
          {8}       ,'Abandonner la saisie ?'
          {9}       ,''
          {10}      ,''
          {11}      ,''
          {12}      ,'Confirmez vous l''abandon de la saisie ?'
          {13}      ,''
          {14}      ,'Etes-vous sûr de vouloir supprimer cette ligne?'
              );


const
NatureEntree = 'EEX';
NatureSortie = 'SEX';

Implementation

function GCLanceFiche_MouvStkExContr(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_GCMOUVSTKEXCONTR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCMOUVSTKEXCONTR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCMOUVSTKEXCONTR.OnUpdate ;
var QDC : TQuery;
    stWhere : string;
    iInd,NumLigne : integer;
    TobDC : TOB;
begin
// JT eQualité 10823
  if TestDepotOblig then
  begin
    if GetControlText('GQC_DEPOT') = '' then
    begin
      PGIBox('Le dépôt doit être renseigné',Ecran.Caption);
      SetFocusControl('GQC_DEPOT') ;
      exit;
    end;
  end;
// fin JT eQualité 10823
if Not SortDeLaLigne then exit;
   Inherited ;
if TobDispoContrem.Somme('(SAISI)',[''],[''],False) <> 0 then
   if PGIAsk(TraduireMemoire(TexteMessage[12]),TraduireMemoire(Ecran.Caption))<>mrYes then exit;
GDisp.VidePile(false) ; TobDispoContrem.ClearDetail ;
TobArticle.ClearDetail ; TobCatalogue.ClearDetail ;
TobPieceGen.ClearDetail;
Depot := GetControlText('GQC_DEPOT');
TobPieceGen.PutValue('GP_DEPOT',Depot);
stWhere := 'WHERE GQC_DEPOT="'+Depot+'" ';
stWhere := stWhere + GetCondition('GQC_ARTICLE','GQC_ARTICLE')
                   + GetCondition('GQC_REFERENCE','GQC_REFERENCE')
                   + GetCondition('GQC_EMPLACEMENT','GQC_EMPLACEMENT')
                   + GetCondition('GQC_FOURNISSEUR','GQC_FOURNISSEUR')
                   + GetCondition('GQC_CLIENT','GQC_CLIENT');
QDC := OpenSQL('SELECT * FROM DISPOCONTREM ' + stWhere + ' ORDER BY GQC_REFERENCE,GQC_FOURNISSEUR,GQC_CLIENT',True,-1, '', True);
if not QDC.Eof then
   TobDispoContrem.LoadDetailDB('DISPOCONTREM','','',QDC,True);
Ferme(QDC);
TobArticle.LoadDetailFromSQL('SELECT * '
    + 'FROM ARTICLE WHERE GA_ARTICLE IN (SELECT DISTINCT(GQC_ARTICLE) FROM DISPOCONTREM ' + stWhere + ')');
TobCatalogue.LoadDetailFromSQL('SELECT * '
    + 'FROM CATALOGU WHERE GCA_ARTICLE IN (SELECT DISTINCT(GQC_ARTICLE) FROM DISPOCONTREM '+ stWhere + ')');
NumLigne := 0;
for iInd := 0 to TobDispoContrem.Detail.Count-1 do
   begin
   TobDC := TobDispoContrem.Detail[iInd];
   AddLesSupDispo(TobDC);
   InitLesSupDispo(TobDC);
   inc(NumLigne);
   CreerTobLigne(TobPieceGen,TobDC,NumLigne);
   end;
TobDispoContrem.PutGridDetail(GDisp,False,False,GDisp.Titres[0]);
HMTrad.ResizeGridColumns(GDisp);
if TobDispoContrem.Detail.Count>0 then
   begin
   GDisp.Enabled:=True;
   for iInd := 1 to GDisp.RowCount -1 do GDisp.Cells[ColSaisi,iInd] := '';
   RowEnter(1);
   end else
   begin
   GDisp.Enabled:=False;
   end;
end ;


procedure TOF_GCMOUVSTKEXCONTR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCMOUVSTKEXCONTR.OnArgument (S : String ) ;
var iCol,iInd : integer;
    CodeL : string;
begin
  Inherited ;
///Initialisation
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
bZoomReference := TMenuItem(GetControl('mnZoomReference'));
bZoomReference.OnClick := BZoomReferenceClick;
bZoomDispo := TMenuItem(GetControl('mnZoomDispo'));
bZoomDispo.OnClick := BZoomDispoClick;
bValideMouv := TToolbarButton97(GetControl('BVALIDEMOUV'));
bSupprimer := TToolbarButton97(GetControl('BSUPPRIMER'));
bInserer := TToolbarButton97(GetControl('BINSERER'));
bValideMouv.OnClick := bValideMouvClick;
bSupprimer.OnClick := bSupprimerClick;
bInserer.OnClick := bInsererClick;
THEdit(GetControl('GQC_REFERENCE')).OnElipsisClick := bRechInCatalogueElipsisClick;
THEdit(GetControl('GQC_REFERENCE_')).OnElipsisClick := bRechInCatalogueElipsisClick;
SetControlText('GQC_DEPOT',GetParamSoc('SO_GCDEPOTDEFAUT'));
//Grid
GInfosLigne := THGrid(GetControl('INFOSLIGNE'));
GDisp := THGrid(GetControl('GDISP')) ;
GDisp.ListeParam:='GCMOUVSTKEXCON';
GDisp.OnRowEnter := GDispRowEnter;
GDisp.OnRowExit := GDispRowExit;
GDisp.OnCellExit := GDispCellExit;
EtudieColsListes;
for iCol := 1 to GDisp.ColCount-1 do  ///**
    if (iCol <> ColSaisi) and (iCol <> ColSens) and (iCol <> ColMotif) then GDisp.ColLengths[iCol] := -1;
GDisp.ColFormats[ColSens] := 'CB=GCSENSPIECE| AND CO_CODE<>"MIX"';
GDisp.ColFormats[ColMotif] := 'CB=GCMOTIFMOUVEMENT';
GDisp.ColTypes[ColSaisi]:='R';
GDisp.ColFormats[ColSaisi]:='##,##';
HMTrad.ResizeGridColumns(GDisp);
ToutAllouer ;
PPInfosLigneEEX.Clear ; PPInfosLigneSEX.Clear ;
for iInd:=1 to 8 do
    begin
    CodeL:=GetInfoParPiece(NatureEntree,'GPP_IFL'+IntToStr(iInd)) ;
    if CodeL<>'' then PPInfosLigneEEX.Add(RechDom('GCINFOLIGNE',CodeL,False)+';'+RechDom('GCINFOLIGNE',CodeL,True)) ;
    CodeL:=GetInfoParPiece(NatureSortie,'GPP_IFL'+IntToStr(iInd)) ;
    if CodeL<>'' then PPInfosLigneSEX.Add(RechDom('GCINFOLIGNE',CodeL,False)+';'+RechDom('GCINFOLIGNE',CodeL,True)) ;
    end;
TobPieceGen := CreerTobPiece(NatureEntree,'');
CellText := '';
GDisp.Enabled:=False;
end ;

procedure TOF_GCMOUVSTKEXCONTR.OnClose ;
begin
if TobDispoContrem.Somme('(SAISI)',[''],[''],False) <> 0 then
    if PGIAsk(TraduireMemoire(TexteMessage[12]),TraduireMemoire(Ecran.Caption))<>mrYes  then
       begin
       AfficheError:=False;
       LastError:=1;
       end;
  Inherited ;
if LastError<> 0 then AfficheError:=true else ToutLiberer ;
end ;


///////// Grid
procedure TOF_GCMOUVSTKEXCONTR.GDispRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
RowEnter(Ou);
CellText := GDisp.Cells[ColSens,Ou]+';'+GDisp.Cells[ColSaisi,Ou];
end;

procedure TOF_GCMOUVSTKEXCONTR.GDispRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if CellText = GDisp.Cells[ColSens,Ou]+';'+GDisp.Cells[ColSaisi,Ou] then exit;
if not ValideLaSaisie(Ou) then
   begin
   GDisp.Col := ColSaisi;
   GDisp.Row := Ou;
   end;
end;

procedure TOF_GCMOUVSTKEXCONTR.GDispCellExit(Sender: TObject; var ACol,ARow : Integer; var Cancel : Boolean);
var TOBL, TobDC : TOB;
begin
if ACol = ColSaisi then GDisp.Cells[ACol,ARow] := StrF00(Valeur(GDisp.Cells[ACol,ARow]),V_PGI.OkDecV);
TobDC := TOB(GDisp.Objects[0,ARow]);
if TobDC = nil then exit;
TOBL := GetTobLignePieceGen(TobDC);
if TOBL = nil then exit;
if ACol = ColSens then
   begin
   TobDC.PutValue('(SENS)',GDisp.CellValues[ACol,ARow]);
   if TobDC.GetValue('(SENS)') = 'ENT' then TOBL.PutValue('GL_NATUREPIECEG',NatureEntree)
   else TOBL.PutValue('GL_NATUREPIECEG',NatureSortie);
   end;
if ACol = ColMotif then  TOBL.PutValue('GL_MOTIFMVT',GDisp.CellValues[ACol,ARow]);
end;

procedure TOF_GCMOUVSTKEXCONTR.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_RETURN : if Screen.ActiveControl=GDisp then Key:=VK_TAB ;
   VK_DELETE : begin
               if ((Screen.ActiveControl=GDisp) and (Shift=[ssCtrl])) then
                   begin Key := 0 ; SupprimeLaLigne (GDisp.Row) ; end;
               end ;
   VK_INSERT : begin
               if ((Screen.ActiveControl=GDisp) and (Shift=[ssCtrl])) then
                   begin Key:=0 ; AjouteUneReference ; end ;
               end;
   VK_F10    : begin Key:=0 ; bValideMouvClick(nil) ; end ;
   VK_F5     : if (Screen.ActiveControl=GDisp) then begin Key:=0 ; bZoomReferenceClick(nil); end;
   end;
end;

procedure TOF_GCMOUVSTKEXCONTR.RowEnter(ARow : integer);
var TobDC : TOB;
begin
TobDC := TOB(GDisp.Objects[0,ARow]);
if TobDC = nil then exit;
ActualiseLePied(TobDC);
end;


Function TOF_GCMOUVSTKEXCONTR.SortDeLaLigne : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
if GDisp.CanFocus then
   begin
   Result:=False ;
   ACol:=GDisp.Col ; ARow:=GDisp.Row ; Cancel:=False ;
   GDispCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
   GDispRowExit(Nil,ARow,Cancel,False) ; if Cancel then Exit ;
   end;
Result:=True ;
end;

///////// Boutons
procedure TOF_GCMOUVSTKEXCONTR.bSupprimerClick(Sender: TObject);
begin
SupprimeLaLigne(GDisp.Row);
end;

procedure TOF_GCMOUVSTKEXCONTR.bInsererClick(Sender: TObject);
begin
AjouteUneReference;
end;

procedure TOF_GCMOUVSTKEXCONTR.bValideMouvClick(Sender: TObject);
var io : TIOErr ;
    JNalCode : string;
    ACol, ARow : integer;
begin
if TobDispoContrem.Detail.Count = 0 then exit;
if Not SortDeLaLigne then exit;
ACol := GDisp.Col ; ARow := GDisp.Row;
if Copy(GDisp.ColFormats[ACol],1,3)='CB=' then PostMessage(GDisp.ValCombo.Handle, WM_KEYDOWN, VK_TAB, 0);
Application.processMessages;
TobEEX.ClearDetail ; TobSEX.ClearDetail ;
MsgJnalEventEEX.Clear ; MsgJnalEventSEX.Clear ;
if not ValideLaSaisie(ARow) then exit;
if TobDispoContrem.Somme('(SAISI)',[''],[''],False) = 0 then //Aucune saisie
   begin
   PGIBox(TraduireMemoire(TexteMessage[2]),TraduireMemoire(Ecran.Caption)) ;
   exit ;
   end;
io:=Transactions(TraiteLaSaisie,5) ;
Case io of
    oeOk       : begin
                 JNalCode := 'OK';
                 PGIBox(GenereCompteRendu,TraduireMemoire(Ecran.Caption));
                 end;
    oeStock    : BEGIN MessageAlerte(TexteMessage[4]) ; JNalCode := 'ERR';  Exit ; END ;
    oeUnknown  : begin MessageAlerte(TexteMessage[5]) ; JNalCode := 'ERR';  end ;
    oeSaisie   : begin MessageAlerte(TexteMessage[5]) ; JNalCode := 'ERR';  end ;
    end ;
// Journal des evenements --------
MAJJnalEvent(MsgJnalEventEEX,JNalCode);
MAJJnalEvent(MsgJnalEventSEX,JNalCode);
if io <> oeOk then exit;
SetFocusControl('GQC_ARTICLE');
GDisp.VidePile(false);
GDisp.Refresh;
GDisp.Col := ColSaisi;
TobDispoContrem.ClearDetail;
end;

procedure TOF_GCMOUVSTKEXCONTR.BZoomReferenceClick(Sender: TObject);
var TobDC : TOB;
begin
TobDC := TOB(GDisp.Objects[0,GDisp.Row]);
if TobDC = nil then exit;
AGLLanceFiche('GC','GCCATALOGU_SAISI3','',TobDC.GetValue('GQC_REFERENCE')
          + ';' + TobDC.GetValue('GQC_FOURNISSEUR'),'ACTION=CONSULTATION;MONOFICHE');
end;

procedure TOF_GCMOUVSTKEXCONTR.bZoomDispoClick(Sender: TObject);
var TobDC : TOB;
begin
TobDC := TOB(GDisp.Objects[0,GDisp.Row]);
if TobDC = nil then exit;
AGLLanceFiche('GC','GCDISPOCONTR_MUL','GQC_REFERENCE=' + TobDC.GetValue('GQC_REFERENCE') + ';'
           + 'GQC_FOURNISSEUR=' + TobDC.GetValue('GQC_FOURNISSEUR') + ';'
           + 'GQC_CLIENT=' + TobDC.GetValue('GQC_CLIENT') + ';'
           + 'GQC_DEPOT=' + Depot,'','');
end;

procedure TOF_GCMOUVSTKEXCONTR.bRechInCatalogueElipsisClick(Sender: TObject);
var RefCata,ControlName : string;
begin
ControlName := THEdit(Sender).Name;
RefCata := GetControlText(ControlName);
RefCata:=AGLLanceFiche('GC','GCMULCATALOG_RECH','','','GCA_REFERENCE='+Trim(RefCata)) ;
RefCata:=uppercase(Trim(ReadTokenSt(RefCata)));
if RefCata<>'' then SetControlText(ControlName,RefCata);
end;

///Initialisation
procedure TOF_GCMOUVSTKEXCONTR.ToutAllouer ;
begin
TobDispoContrem := TOB.Create('',nil,-1) ;
TobSEX := TOB.Create('',nil,-1) ; TobEEX := TOB.Create('',nil,-1) ;
TobArticle := TOB.Create('',nil,-1) ; TobCatalogue := TOB.Create('',nil,-1);
//TobPieceGen := TOB.Create('',nil,-1) ;
MsgJnalEventEEX := TStringList.Create ; MsgJnalEventSEX := TStringList.Create;
PPInfosLigneEEX := TStringList.Create ; PPInfosLigneSEX := TStringList.Create ;
end;

procedure TOF_GCMOUVSTKEXCONTR.ToutLiberer ;
begin
TobDispoContrem.Free;  TobEEX.Free ; TobSEX.Free;
TobArticle.Free; TobCatalogue.Free ; TobPieceGen.Free;
MsgJnalEventEEX.Free ; MsgJnalEventSEX.Free;
PPInfosLigneEEX.Free ; PPInfosLigneSEX.Free;
end;

procedure TOF_GCMOUVSTKEXCONTR.AddLesSupDispo(TobDC : TOB);
begin
TobDC.AddChampSupValeur('(SAISI)',0,False);
TobDC.AddChampSupValeur('(SENS)','ENT',False);
TobDC.AddChampSup('(QTECLIENT)',False);
TobDC.AddChampSup('(CODEARTICLE)',False);
TobDC.AddChampSup('(LIBELLE)',False);
TobDC.AddChampSupValeur('UTILISE', '-', False);
TobDC.AddChampSupValeur('RESERVECLI', 0, False);
end;

procedure TOF_GCMOUVSTKEXCONTR.InitLesSupDispo(TobDC : TOB);
var TobCata : TOB;
    IdentArt : string;
begin
if TobDC = nil then exit;
IdentArt := TobDC.GetValue('GQC_ARTICLE');
TobDC.PutValue('(QTECLIENT)',TobDC.GetValue('GQC_RESERVECLI') + TobDC.GetValue('GQC_PREPACLI'));
if IdentArt <> '' then TobDC.PutValue('(CODEARTICLE)',Copy(IdentArt,1,18));
TobCata := TobCatalogue.FindFirst(['GCA_REFERENCE','GCA_TIERS'],
           [TobDC.GetValue('GQC_REFERENCE'),TobDC.GetValue('GQC_FOURNISSEUR')],False);
if TobCata <> nil then TobDC.PutValue('(LIBELLE)', TobCata.GetValue('GCA_LIBELLE'));
end;

///Form
procedure TOF_GCMOUVSTKEXCONTR.ActualiseLePied(TobDC : TOB);
var TOBL,TobCata,TobArt,TobTemp : TOB;
begin
if TobDC = nil then exit;
TOBL := GetTobLignePieceGen(TobDC);
if TOBL = nil then exit;
TobCata := TobCatalogue.FindFirst(['GCA_REFERENCE','GCA_TIERS'],
           [TobDC.GetValue('GQC_REFERENCE'),TobDC.GetValue('GQC_FOURNISSEUR')],False);
if TobCata = nil then exit;
TobTemp := TOB.Create('', TobCata,-1);  //Nécessaire
TobTemp.Dupliquer(TobDC,True,True);    //   pour InfoLigne
TobArt := TobArticle.FindFirst(['GA_ARTICLE'],[TobDC.GetValue('GQC_ARTICLE')],False);
if TobDC.GetValue('(SENS)') = 'ENT' then MontreInfosLigne(TOBL,TOBArt,TobCata,Nil,GInfosLigne,PPInfosLigneEEX)
else MontreInfosLigne(TOBL,TOBArt,TobCata,Nil,GInfosLigne,PPInfosLigneSEX) ;
TobTemp.Free;
end;

function TOF_GCMOUVSTKEXCONTR.GetCondition(stInf,stSup : string) : string;
var st,stSQLInf,stSQLSup : string;
begin
st := ''; result := '';
//borne inférieure
if stInf <> '' then st := GetControlText(stInf);
if st <> '' then
   if stSup = '' then stSQLInf := stInf + '="' + st + '"'
   else stSQLInf := stInf + '>="' + st + '"';
//borne supérieure
if stSup <> '' then st := GetControlText(stSup+'_');
if st <> '' then
   begin
   if pos('ARTICLE',stSup) > 0 then stSQLSup := stSup + '<="' + format('%-18.18s%16.16s',[st,StringOfChar('Z',16)]) + '"'
   else stSQLSup := stSup + '<="' + st + '"';
   if stSQLInf <> '' then stSQLSup := ' AND ' + stSQLSup;
   end;

if (stSQLInf = '') and (stSQLSup = '') then exit;
result := ' AND ' + stSQLInf + stSQLSup;
end;

///Fonctions liées au grid
procedure TOF_GCMOUVSTKEXCONTR.EtudieColsListes ;
Var NomCol,LesCols : String ;
    icol  : integer ;
begin
//Grids des lignes de doc
LesCols := GDisp.Titres[0];
icol:=1 ;     ///**
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol='(SAISI)' then ColSaisi:=icol
    else if NomCol='(SENS)' then ColSens:=icol
       else if NomCol='(MOTIF)' then ColMotif:=icol;
    Inc(icol) ;
Until ((LesCols='') or (NomCol=''));
end;

function TOF_GCMOUVSTKEXCONTR.ValideLaSaisie (ARow : integer) : boolean;
var TobDC,TOBL : TOB;
    NewQtePhy,CellValue : double;
begin
Result := True;
TobDC := TOB(GDisp.Objects[0,ARow]);
if TobDC = nil then exit;
TOBL := GetTobLignePieceGen(TobDC);
if TOBL = nil then exit;
CellValue := Valeur(GDisp.Cells[ColSaisi,ARow]);
if CellValue <> 0 then
   begin
   if CellValue < 0 then
      begin
      PGIBox(TraduireMemoire(TexteMessage[7]),TraduireMemoire(Ecran.Caption));
      result := False;
      exit;
      end;
   if TobDC.GetValue('(SENS)')='ENT' then NewQtePhy := TobDC.GetValue('GQC_PHYSIQUE') + CellValue
   else NewQtePhy := TobDC.GetValue('GQC_PHYSIQUE') - CellValue ;
   if NewQtePhy < 0 then
      begin
      PGIBox(TraduireMemoire(TexteMessage[3]),TraduireMemoire(Ecran.Caption));
      result := False;
      exit;
      end;
   if (TobDC.GetValue('GQC_CLIENT') <> '') and (NewQtePhy > TobDC.GetValue('(QTECLIENT)')) then
        PGIBox(TraduireMemoire(TexteMessage[1]),TraduireMemoire(Ecran.Caption));
   TobDC.PutValue('UTILISE','X');
   TOBL.PutValue('UTILISE','X');
   end else
   begin
   TobDC.PutValue('UTILISE','-');
   TOBL.PutValue('UTILISE','-');
   end;

TobDC.PutValue('(SAISI)',Cellvalue);
TOBL.PutValue('GL_QTEFACT',CellValue);
TOBL.PutValue('GL_QTESTOCK',CellValue);
TOBL.PutValue('GL_QTERESTE',CellValue); { NEWPIECE }
end;

procedure TOF_GCMOUVSTKEXCONTR.SupprimeLaLigne(ARow : integer);
var TobDC,TOBL : TOB;
begin
if TobDispoContrem.Detail.Count = 0 then exit;
if PGIAsk(TraduireMemoire(TexteMessage[14]),TraduireMemoire(Ecran.Caption))<>mrYes then exit;
TobDC := TOB(GDisp.Objects[0,ARow]);
if TobDC <> nil then
   begin
   TOBL := GetTobLignePieceGen(TobDC);
   if TOBL <> nil then TOBL.Free;
   TobDC.Free;
   end;
with GDisp do
    begin
    if RowCount = 2 then RowCount := 3;
    CacheEdit ; SynEnabled := False;
    DeleteRow (ARow);
    MontreEdit; SynEnabled := True;
    end;
end;

procedure TOF_GCMOUVSTKEXCONTR.AjouteUneReference;
var st,RefCata,RefFour,RefCli : string;
    Retour,ChampMul,ValMul : string;
    ipos : integer;
    TobT,TobDC,TobR,TobA,TobC : TOB;
    QQ : TQuery;
begin
TobDC := nil;
Retour := AglLanceFiche ('GC','GCDISPOCONTR_MUL','GQC_DEPOT='+Depot,'','DISABLED_DEPOT');
if Retour = '' then exit;
Repeat
   st:=Trim(ReadTokenSt(Retour));
   if st<>'' then
      begin
      ipos:=pos('=',st);
      if ipos<>0 then
         begin
         ChampMul:=uppercase(copy(st,1,ipos-1));
         ValMul:=copy(st,ipos+1,length(st));
         if ChampMul='GQC_REFERENCE' then RefCata := ValMul
         else if ChampMul='GQC_FOURNISSEUR' then RefFour := ValMul
         else if ChampMul='GQC_CLIENT' then RefCli := ValMul;
         end;
      end;
until Retour='';

TobT := TobDispoContrem.FindFirst(['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],
              [Depot,RefCata,RefFour,RefCli],False);
if TobT <> nil then
   begin
   PGIBox(TraduireMemoire(TexteMessage[6]),TraduireMemoire(Ecran.Caption));
   exit;
   end;
QQ := OpenSQL('SELECT * FROM DISPOCONTREM WHERE GQC_DEPOT="'+Depot+'" '
              +'AND GQC_REFERENCE="'+RefCata+'" AND GQC_FOURNISSEUR="'+RefFour+'" '
              +'AND GQC_CLIENT="'+RefCli+'"', True,-1, '', True);
if not QQ.Eof then
   begin
   TobDC := TOB.Create('',TobDispoContrem,-1);
   TobDC.SelectDB('',QQ);
   AddLesSupDispo(TobDC);
   end;
Ferme(QQ);
if TobDC = nil then exit;
/// ... article
TobR := TobArticle.FindFirst(['GA_ARTICLE'],[TobDC.GetValue('GQC_ARTICLE')],False);
if TobR = nil then
   begin
   QQ := OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="' + TobDC.GetValue('GQC_ARTICLE') + '"',True,-1, '', True);
   if not QQ.Eof then
       begin
       TobA := TOB.Create('',TobArticle,-1);
       TobA.SelectDB('',QQ);
       end;
   Ferme(QQ);
   end;
/// ... catalogue
TobR := TobCatalogue.FindFirst(['GCA_REFERENCE','GCA_TIERS'],[RefCata,RefFour],False);
if TobR = nil then
   begin
   QQ := OpenSQL('SELECT * '
              + 'FROM CATALOGU WHERE GCA_REFERENCE="'+ RefCata +'" AND GCA_TIERS="'+ RefFour +'"',True,-1, '', True);
   if not QQ.Eof then
       begin
       TobC := TOB.Create('',TobCatalogue,-1);
       TobC.SelectDB('',QQ);
       end;
   Ferme(QQ);
   end;
InitLesSupDispo(TobDC);
CreerTobLigne(TobPieceGen,TobDC,0);
with GDisp do
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
  TobDC.PutLigneGrid(GDisp, RowCount-1, false, false, Titres[0]);
  Cells[ColSaisi,RowCount-1] := '';
  end;
end;

procedure TOF_GCMOUVSTKEXCONTR.TraiteLaSaisie;
var iInd : integer;
    TobDC,TobDispoEntrees,TobDispoSorties : TOB;
begin
TobDispoEntrees := TOB.Create('',nil,-1);
TobDispoSorties := TOB.Create('',nil,-1);
TobDispoContrem.Detail.Sort('GQC_FOURNISSEUR;GQC_REFERENCE');
for iInd := TobDispoContrem.Detail.Count -1 downto 0 do
    begin
    TobDC := TobDispoContrem.Detail[iInd];
    if TobDC.GetValue('(SENS)') = 'ENT' then TobDC.ChangeParent(TobDispoEntrees,-1)
    else TobDC.ChangeParent(TobDispoSorties,-1);
    end;
TobDispoEntrees.ChangeParent(TobDispoContrem,-1);
TobDispoSorties.ChangeParent(TobDispoContrem,-1);
EpureLesLignes;
if V_PGI.IOError = oeOk then GenereLesPieces(TobDispoEntrees,TobEEX,MsgJnalEventEEX,NatureEntree);
if V_PGI.IOError = oeOk then GenereLesPieces(TobDispoSorties,TobSEX,MsgJnalEventSEX,NatureSortie);
end;


////// Pièces
function TOF_GCMOUVSTKEXCONTR.AssocieCatalogueDispoContreM(TobDispoCM : TOB) : TOB;
var TobDC,TobCata,TobCatalog,TobSansClient : TOB;
    iInd : integer;
    RefFour,RefCata,RefCli : string;
    QQ : TQuery;
begin
TobCatalog := TOB.Create('',nil,-1) ;
TobCatalog.Dupliquer(TobCatalogue,True,True);
for iInd := TobDispoCM.Detail.Count -1 downto 0 do
    begin
    TobDC := TobDispoCM.Detail[iInd];
    if TobDC.GetValue('UTILISE') = '-' then continue;
    RefFour := TobDC.getValue('GQC_FOURNISSEUR');
    RefCata  := TobDC.getValue('GQC_REFERENCE');
    RefCli  := TobDC.getValue('GQC_CLIENT');
    TobCata := TobCatalog.FindFirst(['GCA_TIERS','GCA_REFERENCE'],[RefFour,RefCata],False);
    if TobCata = nil then continue;
    TobDC.ChangeParent(TobCata,-1);
    TobCata.AddChampSupValeur('UTILISE','X',False); //Nécessaire pour ValideLesLignes
    if (RefCli <> '') and (TobCata.FindFirst(['GQC_FOURNISSEUR','GQC_REFERENCE','GQC_CLIENT'],[RefFour,RefCata,''],False) = nil) then
       begin
       TobSansClient := TobDispoContrem.FindFirst(['GQC_FOURNISSEUR','GQC_REFERENCE','GQC_CLIENT'],[RefFour,RefCata,''],True);
       if TobSansClient <> nil then TobSansClient.ChangeParent(TobCata,-1)
       else begin //Recherche dans la table
            QQ := OpenSQL('SELECT * FROM DISPOCONTREM WHERE GQC_FOURNISSEUR="'+RefFour+'"'
                                       + 'AND GQC_REFERENCE="'+RefCata+'" AND GQC_CLIENT=""',True,-1, '', True);
            if not QQ.Eof then
               begin
               TobSansClient := TOB.Create('DISPOCONTREM',TobCata,-1);
               TobSansClient.SelectDB('',QQ);
               TobSansClient.AddChampSupValeur('RESERVECLI',0,False);
               end;
            Ferme(QQ);
            end;
       end;
    end;
for iInd := TobCatalog.Detail.Count -1 downto 0 do
    begin
    TobCata := TobCatalog.Detail[iInd];
    if TobCata.Detail.Count  <= 0 then TobCata.Free;
    end;
Result:=TobCatalog;
end;

procedure TOF_GCMOUVSTKEXCONTR.GenereLesPieces(TobSource,TobES : TOB ; JNal : TStringList ; NaturePiece : string);
var iInd,NbPiece,NumPiece : integer;
    TobDC,TobCatalog : TOB;
    TobPiece,TobTemp : TOB;
    RefFour : string;
begin
RefFour := ''; TobTemp := TOB.Create('',nil,-1) ;
for iInd := 0 to TobSource.Detail.Count -1 do  //Création des tobs pièce
    begin
    TobDC := TobSource.Detail[iInd];
    if TobDC.GetValue('UTILISE') = '-' then continue;
    if RefFour <> TobDC.getValue('GQC_FOURNISSEUR') then
       begin
       RefFour := TobDC.getValue('GQC_FOURNISSEUR');    //Regroup par tiers
       TobPiece := CreerTobPiece(NaturePiece,RefFour);
       TobPiece.ChangeParent(TobES,-1);
       AssocieLignePiece(TobPiece);
       end;
    end;
TobCatalog := AssocieCatalogueDispoContreM(TobSource);   //pour utilisation des fonctions des pièces
NbPiece := TobES.Detail.Count;
if V_PGI.IOError = oeOk then
    if NbPiece > 0 then
       begin
       NumPiece := SetNumberAttribution (TobES.Detail[0].GetValue('GP_NATUREPIECEG'),TobES.Detail[0].GetValue('GP_SOUCHE'), TobES.Detail[0].GetValue('GP_DATEPIECE'), NbPiece);
       if NumPiece <= 0 then begin V_PGI.IOError := OeSaisie ; exit ; end;
       for iInd := 0 to NbPiece -1 do
           begin
           TobPiece := TobES.Detail[iInd];
           if not SetDefinitiveNumber(TobPiece,nil,nil,nil,nil,nil,nil,nil,nil,NumPiece) then   
              begin V_PGI.IOError := OeSaisie ; exit ; end;
           inc(NumPiece);
           AddEvent(JNal,TobPiece);
           ReaffecteDispoContreM(TobTemp,TobPiece,TobCatalog);
           if V_PGI.IOError = oeOk then ValideLesLignes(TobPiece,TobTemp,TobCatalog,nil,nil,nil,nil,False,False);
           if V_PGI.IOError = oeOk then ValideLesCatalogues(TobPiece,TobCatalog);
           end;
       if V_PGI.IOError = oeOk then TobES.InsertDB(nil,True);
       end;
TobTemp.Free; TobCatalog.Free;
end;

procedure TOF_GCMOUVSTKEXCONTR.AssocieLignePiece(TobPiece : TOB);
var iInd: integer;
    TOBL : TOB;
    Nature : string;
begin
Nature := TobPiece.GetValue('GP_NATUREPIECEG');
for iInd := TobPieceGen.Detail.Count -1 downto 0 do
    begin
    TOBL := TobPieceGen.Detail[iInd];
    if (Nature = TOBL.GetValue('GL_NATUREPIECEG')) and (TobPiece.GetValue('GP_TIERS') = TOBL.GetValue('GL_TIERS')) then
       begin
       PieceVersLigne(TobPiece,TOBL);
       TOBL.ChangeParent(TobPiece,-1);
       end;
    end;
NumeroteLignesGC(Nil,TOBPiece) ;
end;

function TOF_GCMOUVSTKEXCONTR.CreerTobPiece(Nature,CodeTiers : string) : TOB;
var CleDoc : R_CleDoc ;
    DatePiece : TDateTime;
    iPeriode,iSemaine  : integer;
begin
CleDoc.NaturePiece := Nature;
DatePiece := V_PGI.DateEntree;
Result := CreerTOBPieceVide (CleDoc,CodeTiers,'','','', True,False);
if Result <> nil then
   begin
   AddLesSupEntete (Result);
   Result.PutValue('GP_DATEPIECE',DatePiece);
   Result.PutValue('GP_DEPOT',Depot);
   if GetInfoParPiece(Nature,'GPP_ACTIONFINI')='TRA' then Result.PutValue('GP_VIVANTE','X')
   else Result.PutValue('GP_VIVANTE','-');
   iPeriode := GetPeriode(DatePiece); iSemaine := NumSemaine(DatePiece);
   Result.PutValue ('GP_PERIODE',iPeriode);
   Result.PutValue ('GP_SEMAINE',isemaine);
   Result.PutValue ('GP_CONTREMARQUE','X');
   end;
end;

procedure TOF_GCMOUVSTKEXCONTR.CreerTobLigne(TobPiece,TobDC : TOB ; NumLigne : integer);
Var TobLigne,TobCata,TobRefArt : TOB;
    RefTiers,RefFour : string;
    QteSaisie : double;
begin
QteSaisie := TobDC.GetValue('(SAISI)');
RefTiers := TobDC.GetValue('GQC_FOURNISSEUR');
RefFour := TobDC.GetValue('GQC_CLIENT');
TobLigne := NewTOBLigne(TOBPiece,0);
InitLigneVide(TobPiece,TobLigne,nil,nil,NumLigne,QteSaisie);
TobCata := TobCatalogue.FindFirst(['GCA_REFERENCE','GCA_TIERS'],[TobDC.GetValue('GQC_REFERENCE'),RefTiers],False);
TobRefArt := TobArticle.FindFirst(['GA_ARTICLE'],[TobDC.GetValue('GQC_ARTICLE')],False);
CodesCataToCodesLigne(TobLigne,TobCata,TobRefArt);
CatalogueVersLigne(TobPiece,TobCata,TobLigne,nil,TobRefArt);
TobLigne.PutValue('GL_NUMLIGNE',NumLigne);
TobLigne.PutValue('GL_FOURNISSEUR',RefFour);
TobLigne.PutValue('GL_TIERS',RefTiers);
TobLigne.PutValue('GL_QTEFACT',QteSaisie);
TobLigne.PutValue('GL_QTESTOCK',QteSaisie);
TobLigne.PutValue('GL_QTERESTE',QteSaisie);  { NEWPIECE }
TobLigne.PutValue('GL_PERIODE', TobPiece.GetValue('GP_PERIODE')) ;
TobLigne.PutValue('GL_SEMAINE', TobPiece.GetValue('GP_SEMAINE')) ;
{Prix}
TobLigne.PutValue('GL_PUHT',0) ;     TobLigne.PutValue('GL_PUHTDEV',0) ;
TobLigne.PutValue('GL_PUTTC',0) ;    TobLigne.PutValue('GL_PUTTCDEV',0) ;
TobLigne.PutValue('GL_PUHTNET',0) ;  TobLigne.PutValue('GL_PUHTNETDEV',0);
TobLigne.PutValue('GL_PUTTCNET',0) ; TobLigne.PutValue('GL_PUTTCNETDEV',0);

TobLigne.AddChampSupValeur('UTILISE','-',False);
end;

function TOF_GCMOUVSTKEXCONTR.GetTobLignePieceGen(TobDC : TOB) : TOB;
begin
Result := nil;
if TobDC = nil then exit;
Result := TobPieceGen.FindFirst(['GL_TIERS','GL_REFARTSAISIE','GL_FOURNISSEUR'],
    [TobDC.GetValue('GQC_FOURNISSEUR'),TobDC.GetValue('GQC_REFERENCE'),TobDC.GetValue('GQC_CLIENT')],false);
end;

procedure TOF_GCMOUVSTKEXCONTR.EpureLesLignes;
var iInd : integer;
    TOBL : TOB;
begin
for iInd := TobPieceGen.Detail.Count -1 downto 0 do
    begin
    TOBL := TobPieceGen.Detail[iInd];
    if TOBL.GetValue('UTILISE') = '-' then TOBL.Free;
    end;
end;

procedure TOF_GCMOUVSTKEXCONTR.CodesCataToCodesLigne(TOBL, TOBCata, TOBArt: TOB);
begin
if TOBL = nil then exit;
if TOBCata = nil then exit;
TOBL.PutValue('GL_ENCONTREMARQUE','X') ;
TOBL.PutValue('GL_REFARTSAISIE',TOBCata.GetValue('GCA_REFERENCE')) ;
TOBL.PutValue('GL_ARTICLE',TOBCata.GetValue('GCA_ARTICLE')) ;
TOBL.PutValue('GL_REFCATALOGUE',TOBCata.GetValue('GCA_REFERENCE')) ;
TOBL.PutValue('GL_TYPEREF','CAT') ;
if TobArt <> Nil then
   begin
   TOBL.PutValue('GL_CODEARTICLE',TOBArt.GetValue('GA_CODEARTICLE')) ;
   TOBL.PutValue('GL_TYPEREF','ART') ;
   end;
end;

function TOF_GCMOUVSTKEXCONTR.GenereCompteRendu : string;
var NbP : integer;
begin
Result := 'Le traitement s''est correctement effectué.';
NbP := TobEEX.Detail.Count;
if NbP > 0 then
   begin
   Result := Result + '#13#10' + IntToStr(NbP) + ' pièce(s) générée(s) de nature '
           + RechDom('GCNATUREPIECEG',TobEEx.Detail[0].GetValue('GP_NATUREPIECEG'),False)
           + ' du N° ' + IntToStr(TobEEx.Detail[0].GetValue('GP_NUMERO'))
           + ' au N° ' + IntToStr(TobEEx.Detail[NbP-1].GetValue('GP_NUMERO')) ;
   end;
NbP := TobSEX.Detail.Count;
if NbP > 0 then
   begin
   Result := Result + '#13#10' + IntToStr(NbP) + ' pièce(s) générée(s) de nature '
           + RechDom('GCNATUREPIECEG',TobSEX.Detail[0].GetValue('GP_NATUREPIECEG'),False)
           + ' du N° ' + IntToStr(TobSEX.Detail[0].GetValue('GP_NUMERO'))
           + ' au N° ' + IntToStr(TobSEX.Detail[NbP-1].GetValue('GP_NUMERO')) ;
   end;
end;

procedure TOF_GCMOUVSTKEXCONTR.AddEvent(MsgJnalEvent : TStringList ; TobPiece : TOB);
begin
MsgJnalEvent.Add(RechDom('GCNATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG'),False)
   + ' N° ' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ', Tiers ' + TOBPiece.GetValue('GP_TIERS'));
end;

procedure TOF_GCMOUVSTKEXCONTR.MAJJnalEvent(MsgJnalEvent : TStringList ; MsgCode : string);
var QQ : TQuery;
    TobJnal : TOB;
    NumEvent : integer;
begin
if MsgJnalEvent.Text = '' then exit;
NumEvent := 0;
TOBJnal := TOB.Create('JNALEVENT', nil, -1);
TOBJnal.PutValue('GEV_TYPEEVENT', 'GEN');
TOBJnal.PutValue('GEV_LIBELLE', Ecran.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', true,-1, '', True);
if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
Ferme(QQ);
inc(NumEvent);
TOBJnal.PutValue('GEV_NUMEVENT', NumEvent);
TOBJnal.PutValue('GEV_ETATEVENT', MsgCode);
TOBJnal.PutValue('GEV_BLOCNOTE', MsgJnalEvent.Text);
TOBJnal.InsertDB(nil);
TOBJnal.Free;
end;

Initialization
  registerclasses ( [ TOF_GCMOUVSTKEXCONTR ] ) ;
end.
