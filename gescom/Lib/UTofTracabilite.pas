{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TRACABILITE ()
Mots clefs ... : TOF;TRACABILITE
*****************************************************************}
Unit UTofTracabilite;

Interface

Uses StdCtrls,
     Controls,Windows,
     Classes,M3FP,Buttons,Vierge,
{$IFDEF EAGLCLIENT}
     eFiche,MaineAGL,
{$ELSE}
     db,FE_Main,
     dbtables,Fiche,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,ParamDBG,LicUtil,Ent1,Menus,UTofOptionEdit,TarifUtil,HDimension,LookUp,AGLInitGC,
     UTOF,UTOB,EntGC,HTB97, Facture, FactUtil,HPanel,HSysMenu;

Type
  TOF_TRACABILITE = Class (TOF)

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    GENT, GSOR, GSTO : THGrid;
    HMTrad: THSystemMenu;
    BParamListe,BImprimer,BZoom : TToolbarButton97;
    BZoomDispo,BZoomClient,BZoomFour,BZoomPiece : TBitBtn;
    PEtat : RPARETAT;
    TobLigneA, TobLigneV, TobStock : TOB; // Tobs des mouvements
    TobAfficheA, TobAfficheV, TobAfficheS : TOB; // Tobs d'affichage
    IdentArticle, OldArticle : string;
    ListePiece,ListeStock : string;
    TableStock, TableLigne, PrefStock, PrefLigne, Reference : string;
    stTitresStock, stTitresLigne : string;
    iEntree, iCpte : integer;
    ColRef, ColTiers, ColQteFact, ColQtePhy, ColReserv, ColPrepa: integer;

    /// Boutons
    procedure BZoomClick(Sender: TObject);
    procedure BZoomDispoClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BZoomClientClick(Sender: TObject);
    procedure BZoomFourClick(Sender: TObject);
    procedure ParamListeClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    /// Grid
    procedure GDblClick(Sender: TObject);
    procedure GridColumnWidthsChanged(Sender: TObject);
    procedure InitialiseForm;
    procedure RefreshGrids;
    procedure EtudieColsListes ;
    procedure AppelLaPiece(VenteAchat : string);
    procedure AddChampListe(var NomListe,NomChamp : string);
    procedure InitCumul(Grid : THGrid ; ColQte : integer);
    ///Tob
    procedure ChargeTobStock;
    procedure ChargeTobMouvement;
    procedure ActualiseGENT;
    procedure ActualiseGSOR;
    procedure ActualiseGSTO;
    ///Fiche
    function  RecupCriteres(stPrefixe : string) : string;
    procedure ArticleElipsisClick(Sender : TObject);
    procedure ArticleExit(Sender : TObject);
    procedure RechercheDimensions;
    function  RechercheLibelleDim(TobArt : TOB) : string;
    ///Impression
    procedure ChargeLigneImpression(TobLigne, TobTmp : TOB);
    function  PrepareImpression : integer;

  end ;

const iLot   : integer = 1;
const iSerie : integer = 2;

const
// libellés des messages
TexteMessage: array[1..2] of string = (
          {1}  'Vous devez renseigner un article'
          {2} ,'Vous devez renseigner un numéro'
              );

Implementation

procedure TOF_TRACABILITE.OnArgument (S : String ) ;
var stArg : string;
begin
  Inherited ;
//***Récup des tables et des préfixes
ListePiece := 'GL_NATUREPIECEG;GL_NUMERO;GL_NUMLIGNE;GL_SOUCHE;GL_INDICEG;GL_VIVANTE;'
             +'GL_ARTICLE;GL_DEPOT;GL_DATEPIECE;T_AUXILIAIRE;GL_TIERS;T_LIBELLE';
ListeStock := 'GQ_ARTICLE;GQ_DEPOT;GQ_EMPLACEMENT';
stArg := Trim(S);
if stArg = 'LOT' then iEntree := iLot
   else if stArg = 'SERIE' then iEntree := iSerie;
if iEntree = iLot then
    begin
    TableStock := 'DISPOLOT' ; TableLigne := 'LIGNELOT';
    Reference :=  'NUMEROLOT';
    ListePiece := ListePiece + ';GLL_QUANTITE';
    ListeStock := ListeStock + ';GQL_DATELOT;GQL_PHYSIQUE';
    end else
if iEntree = iSerie then
    begin
    TableStock := 'DISPOSERIE' ; TableLigne := 'LIGNESERIE';
    Reference :=  'IDSERIE';
    end;
PrefLigne := TableToPrefixe(TableLigne);
PrefStock := TableToPrefixe(TableStock);
ListePiece := ListePiece+ ';'+PrefLigne+'_'+Reference;
ListeStock := ListeStock+ ';'+PrefStock+'_'+Reference;
//Affectation
GENT:=THGrid(GetControl('GENT'));
GSOR:=THGrid(GetControl('GSOR'));
GSTO:=THGrid(GetControl('GSTO'));
GENT.OnDblClick := GDblClick;
GSOR.OnDblClick := GDblClick;
GSTO.OnDblClick := GDblClick;
THEdit(GetControl('CODEARTICLE')).OnElipsisClick := ArticleElipsisClick;
THEdit(GetControl('CODEARTICLE')).OnDblClick := ArticleElipsisClick;
THEdit(GetControl('CODEARTICLE')).OnExit := ArticleExit;
//Boutons
BParamListe := TToolbarButton97(GetControl('BParamListe'));
BImprimer := TToolbarButton97(GetControl('BImprimer'));
BZoom := TToolbarButton97(GetControl('BZOOM'));
BZoomClient := TBitBtn(GetControl('BZOOMCLIENT'));
BZoomFour := TBitBtn(GetControl('BZOOMFOUR'));
BZoomPiece := TBitBtn(GetControl('BZOOMPIECE'));
BZoomDispo := TBitBtn(GetControl('BZOOMDISPO'));
BParamListe.OnClick := ParamListeClick;
BImprimer.OnClick := BImprimerClick;
BZoom.OnClick := BZoomClick;
BZoomClient.OnClick := BZoomClientClick;
BZoomFour.OnClick := BZoomFourClick;
BZoomDispo.OnClick := BZoomDispoClick;
BZoomPiece.OnClick := BZoomPieceClick;
////////Initialisation
BParamListe.Visible := (V_PGI.PassWord = CryptageSt(DayPass(Date)));
//Grids
GENT.ListeParam := 'GCTRACABILITELIG';    GSOR.ListeParam := 'GCTRACABILITELIG';
if iEntree = iLot then GSTO.ListeParam := 'GCTRACABILITELOT'
else GSTO.ListeParam := 'GCTRACABILITESERI';
EtudieColsListes ;
InitialiseForm;
HMTrad.ResizeGridColumns(GENT) ; HMTrad.ResizeGridColumns(GSOR);
HMTrad.ResizeGridColumns(GSTO);
IdentArticle := '' ; OldArticle := '';
//Tobs
TobLigneA := TOB.Create('',nil,-1) ; TobAfficheA := TOB.Create('',nil,-1);
TobLigneV := TOB.Create('',nil,-1) ; TobAfficheV := TOB.Create('',nil,-1);
TobStock := TOB.Create('',nil,-1)  ; TobAfficheS := TOB.Create('',nil,-1);
//Panel cumul
GENT.OnColumnWidthsChanged := GridColumnWidthsChanged;
GSOR.OnColumnWidthsChanged := GridColumnWidthsChanged;
GSTO.OnColumnWidthsChanged := GridColumnWidthsChanged;
if iEntree = iLot then
   begin
   InitCumul(GENT,ColQteFact); InitCumul(GSOR,ColQteFact);
   InitCumul(GSTO,ColQtePhy);
   end;
//Etat
PEtat.Tip:='E'; PEtat.Nat:='GZJ'; PEtat.Modele:='GZJ'; PEtat.Titre:=Ecran.Caption;
PEtat.Apercu:=True; PEtat.DeuxPages:=False; PEtat.First:=True;
PEtat.stSQL := 'GZJ_UTILISATEUR = "'+ V_PGI.USer +'"';
//Menu zoom
PopZoom97(TToolbarButton97(GetControl('BZOOM')),TPopupMenu(GetControl('POPMENU'))) ;
end ;

procedure TOF_TRACABILITE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_TRACABILITE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_TRACABILITE.OnUpdate ;
var stCondition : string;
begin
  Inherited ;
GENT.VidePile(false) ;    GSOR.VidePile(false) ;    GSTO.VidePile(false);
TobLigneA.ClearDetail ;   TobLigneV.ClearDetail ;   TobStock.ClearDetail ;
TobAfficheA.ClearDetail ; TobAfficheV.ClearDetail ; TobAfficheS.ClearDetail ;
THNumEdit(GetControl('NCUMULENT')).Value := 0 ; THNumEdit(GetControl('NCUMULSOR')).Value := 0;
THNumEdit(GetControl('NCUMULSTO')).Value := 0 ;
if THEdit(GetControl('CODEARTICLE')).Focused then ArticleExit(GetControl('CODEARTICLE'));
{if GetControlText('CODEARTICLE')='' then
   begin
   SetFocusControl('CODEARTICLE');
   HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';W;O;O;O;','','');
   exit;
   end; }
if GetControlText('REFERENCE')='' then
   begin
   SetFocusControl('REFERENCE');
   HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[2]+';W;O;O;O;','','');
   exit;
   end;
if iEntree = iLot then stCondition := '" AND GA_LOT="X"' else stCondition := '" AND GA_NUMEROSERIE="X"';

if IdentArticle<>'' then
  if not ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'
                   + IdentArticle + stCondition) then exit;

ChargeTobMouvement; ChargeTobStock;
ActualiseGENT ; ActualiseGSOR ; ActualiseGSTO;
end ;

procedure TOF_TRACABILITE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_TRACABILITE.OnClose ;
begin
  Inherited ;
TobLigneA.Free ;
TobAfficheA.Free;
TobLigneV.Free ; TobAfficheV.Free;
TobStock.Free  ; TobAfficheS.Free;
end ;

////////////////////////////////////////////////////////////////////////////////
//Grid
procedure TOF_TRACABILITE.GDblClick(Sender: TObject);
var Onglet : TTabSheet;
    TOBD : TOB;
begin
Onglet := GetActiveTabSheet('PAGECONTROL');
if Onglet.Name = 'TABSTOCK' then
   begin
   if TobAfficheS.Detail.Count < GSTO.Row then exit;
   TOBD := Tob(GSTO.Objects[0,GSTO.Row]);
   if TOBD = nil then exit;
   AGLLanceFiche('GC','GCDISPO','',TOBD.GetValue('GQ_ARTICLE') +';'+
   TOBD.GetValue('GQ_DEPOT')+';-;01/01/1900','ACTION=CONSULTATION') ;
   exit;
   end;
if (Onglet.Name = 'TABACHAT') or (Onglet.Name = 'TABVENTE') then AppelLaPiece(Onglet.Name);
end;

procedure TOF_TRACABILITE.BZoomClick(Sender: TObject);
var Onglet : TTabSheet;
begin
Onglet := GetActiveTabSheet('PAGECONTROL');
SetControlEnabled('BZOOMDISPO',(Onglet.Name = 'TABSTOCK'));
SetControlEnabled('BZOOMCLIENT',(Onglet.Name = 'TABVENTE'));
SetControlEnabled('BZOOMFOUR',(Onglet.Name = 'TABACHAT'));
SetControlEnabled('BZOOMPIECE',(Onglet.Name <> 'TABSTOCK'));
PopZoom97(TToolbarButton97(GetControl('BZOOM')),TPopupMenu(GetControl('POPMENU'))) ;
end;

procedure TOF_TRACABILITE.ParamListeClick(Sender: TObject);
var Liste : string;
    Onglet : TTabSheet;
begin
Liste := '';
Onglet := GetActiveTabSheet('PAGECONTROL');
if Onglet.Name = 'TABSTOCK' then Liste := GSTO.ListeParam
else Liste := GENT.ListeParam;
if Liste = '' then exit;
{$IFDEF EAGLCLIENT}
ParamListe(Liste,nil) ;
{$ELSE}
ParamListe(Liste,nil,nil) ;
{$ENDIF}
EtudieColsListes;
InitialiseForm;
HMTrad.ResizeGridColumns(GENT); HMTrad.ResizeGridColumns(GSOR);
TobLigneA.ClearDetail ;   TobLigneV.ClearDetail ;   TobStock.ClearDetail ;
TobAfficheA.ClearDetail ; TobAfficheV.ClearDetail ; TobAfficheS.ClearDetail ;
ChargeTobMouvement; ChargeTobStock;
ActualiseGENT ; ActualiseGSOR ; ActualiseGSTO;
end;

procedure TOF_TRACABILITE.BImprimerClick(Sender: TObject);
begin
if (TobAfficheA.Detail.Count = 0) and (TobAfficheS.Detail.Count = 0)
   and (TobAfficheV.Detail.Count = 0) then exit;
SetControlText('LIBELLEARTICLE',RechDom('GCARTICLE',IdentArticle,false) + '  '
                 + GetControlText('TLIBELLEDIM'));
EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('PAGECONTROL')),PEtat.stSQL,PEtat.Titre,PrepareImpression);
ExecuteSQL('DELETE FROM GCTMPTRACABILITE WHERE ' + PEtat.stSQL);
end;

////////////////////////////////////////////////////////////////////////////////
//Tob
procedure TOF_TRACABILITE.ChargeTobStock;
var i_ind : integer;
    TobD : TOB;
    stSQL : string;
begin
stSQL := 'SELECT ' + StringReplace(ListeStock,';',',',[rfReplaceAll])
  + ' FROM ' + TableStock + ' LEFT JOIN DISPO ON GQ_ARTICLE='
  + PrefStock + '_ARTICLE AND GQ_DEPOT=' + PrefStock + '_DEPOT '
  + RecupCriteres(PrefStock) + ' AND GQ_CLOTURE="-"';
TobStock.LoadDetailFromSQL(stSQL,True);
if TobStock.Detail.count <= 0 then exit;
TobStock.Detail[0].AddChampSup('(DEPOT)',true);
TobStock.Detail[0].AddChampSup('(EMPLACEMENT)',true);
for i_ind := 0 to TobStock.Detail.Count -1 do
    begin
    TobD := TobStock.Detail[i_ind];
    TobD.PutValue('(DEPOT)',RechDom('GCDEPOT',TobD.GetValue('GQ_DEPOT'),false));
    TobD.PutValue('(EMPLACEMENT)',RechDom('GCEMPLACEMENT',TobD.GetValue('GQ_EMPLACEMENT'),false));
    end;
end;

procedure TOF_TRACABILITE.ChargeTobMouvement;
var TobMvt, TobD : TOB;
    i_ind : integer;
    stSQL : string;
begin
TobMvt := TOB.Create('',nil,-1);
stSQL := 'SELECT ' + StringReplace(ListePiece,';',',',[rfReplaceAll])
  + ' FROM ' + TableLigne + ' LEFT JOIN LIGNE ON '
  + ' GL_NATUREPIECEG='+PrefLigne+'_NATUREPIECEG AND GL_NUMLIGNE='+PrefLigne+'_NUMLIGNE'
  + ' AND GL_SOUCHE='+PrefLigne+'_SOUCHE AND GL_INDICEG='+PrefLigne+'_INDICEG'
  + ' AND GL_NUMERO='+PrefLigne+'_NUMERO LEFT JOIN TIERS ON GL_TIERS=T_TIERS'
  + RecupCriteres('GL') + ' AND (GL_DATEPIECE>="' + USDateTime(strToDate(GetControlText('DATEPIECE'))) + '"'
  + ' AND GL_DATEPIECE<="' + USDateTime(strToDate(GetControlText('DATEPIECE_'))) + '")';
TobMvt.LoadDetailFromSQL(stSQL,True);
if TobMvt.Detail.Count = 0 then exit;
TobMvt.Detail[0].AddChampSup('(NATUREPIECE)',True);
TobMvt.Detail[0].AddChampSup('(DEPOT)',True);
TobMvt.Detail[0].AddChampSup('(REFERENCE)',True);
TobMvt.Detail[0].AddChampSup('(QUANTITE)',True);
//suppression des pièces non vivantes si elles ont été transformées
for i_ind := TobMvt.Detail.Count -1 downto 0 do
    begin
    TobD := TobMvt.Detail[i_ind];
    if (TobD.GetValue('GL_VIVANTE') = '-') and
       (GetInfoParPiece(TobD.GetValue('GL_NATUREPIECEG'),'GPP_ACTIONFINI') = 'TRA') then TobD.Free
    else begin
         TobD.PutValue('(NATUREPIECE)',RechDom('GCNATUREPIECEG',TobD.GetValue('GL_NATUREPIECEG'),false));
         TobD.PutValue('(DEPOT)',RechDom('GCDEPOT',TobD.GetValue('GL_DEPOT'),false));
         TobD.PutValue('(REFERENCE)',TobD.GetValue(PrefLigne+'_'+Reference));
         if TobD.FieldExists(PrefLigne+'_QUANTITE') then
            TobD.PutValue('(QUANTITE)',TobD.GetValue(PrefLigne+'_QUANTITE'));
         if GetInfoParPiece(TobD.GetValue('GL_NATUREPIECEG'),'GPP_SENSPIECE')='ENT' then
              TobD.ChangeParent(TobLigneA,-1)
         else if GetInfoParPiece(TobD.GetValue('GL_NATUREPIECEG'),'GPP_SENSPIECE')='SOR' then
              TobD.ChangeParent(TobLigneV,-1);
         end;
    end;
end;

////////////////////////////////////////////////////////////////////////////////
//Grid
procedure TOF_TRACABILITE.EtudieColsListes ;
Var NomCol,LesCols : String ;
    icol, ichamp : integer ;
begin
//Grids des lignes de doc
stTitresLigne:=GENT.Titres[0] ; 
LesCols := stTitresLigne;
icol:=0 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then AddChampListe(ListePiece,NomCol);
        end;
    if NomCol='(REFERENCE)' then ColRef:=icol+1 else
    if NomCol='GL_TIERS'    then ColTiers:=icol+1 else
    if NomCol='(QUANTITE)'  then ColQteFact:=icol+1;
    Inc(icol) ;
Until ((LesCols='') or (NomCol=''));

//Grid du stock
stTitresStock:=GSTO.Titres[0] ; 
LesCols := stTitresStock;
icol:=0 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then AddChampListe(ListeStock,NomCol);
        end;
    if NomCol=PrefStock+'_PHYSIQUE' then ColQtePhy:=icol+1 else
    if NomCol=PrefStock+'_ENPREPACLI' then ColPrepa:=icol+1 else
    if NomCol=PrefStock+'_ENRESERVECLI' then ColReserv:=icol+1;
    Inc(icol) ;
Until ((LesCols='') or (NomCol=''));
end ;

procedure TOF_TRACABILITE.RefreshGrids;
var Onglet : TTabSheet;
begin
Onglet := GetActiveTabSheet('PAGECONTROL');
if Onglet.Name = 'TABACHAT' then ActualiseGENT
else if Onglet.Name = 'TABVENTE' then ActualiseGSOR
else ActualiseGSTO;
end;

procedure TOF_TRACABILITE.InitialiseForm;
begin
if iEntree = iLot then
   begin
   Ecran.Caption := TraduireMemoire('Traçabilité des numéros de lot');
   SetControlProperty('CODEARTICLE','Plus','AND GA_LOT="X"');
   SetControlCaption('TREFERENCE','N° de lot  de');
   SetControlText('LOTSERIE',TraduireMemoire('Lot'));
   GENT.Cells[ColRef,0] := TraduireMemoire('Numéro de lot');
   GSOR.Cells[ColRef,0] := TraduireMemoire('Numéro de lot');
   end else
if iEntree = iSerie then
   begin
   Ecran.Caption := TraduireMemoire('Traçabilité des numéros de série');
   SetControlProperty('CODEARTICLE','Plus','AND GA_NUMEROSERIE="X"');
   SetControlCaption('TREFERENCE','N° de série  de');
   SetControlText('LOTSERIE',TraduireMemoire('Série'));
   GENT.Cells[ColRef,0] := TraduireMemoire('Numéro de série');
   GSOR.Cells[ColRef,0] := TraduireMemoire('Numéro de série');
   GENT.ColWidths[ColQteFact] := 0 ; GSOR.ColWidths[ColQteFact] := 0;
   end;
UpdateCaption(Ecran);
SetControlVisible('PSTOCKLOT',(iEntree = iLot));
SetControlVisible('PSTOCKSERIE',(iEntree = iSerie));
SetControlVisible('PCUMULENT',(iEntree = iLot));
SetControlVisible('PCUMULSTO',(iEntree = iLot));
SetControlVisible('PCUMULSOR',(iEntree = iLot));
GENT.Cells[ColTiers,0] := TraduireMemoire('Code fournisseur');
GSOR.Cells[ColTiers,0] := TraduireMemoire('Code client');
GSTO.ColTypes[ColReserv] := 'B'; GSTO.ColFormats[ColReserv] :=intToStr(integer(csCoche));
GSTO.ColTypes[ColPrepa] := 'B';  GSTO.ColFormats[ColPrepa] :=intToStr(integer(csCoche));
SetControlText('TLIBELLEDIM','');
end;

procedure TOF_TRACABILITE.ActualiseGENT;
var i_ind : integer;
    stNatureP, stTiers : string;
    TobD, TobA : TOB;
begin
//Achat
TobAfficheA.ClearDetail ; GENT.VidePile(false);
if TobLigneA.Detail.Count = 0 then exit;
stNatureP := GetControlText('NATUREPIECEGACH');
stTiers := GetControlText('FOURNISSEUR');
if (stNatureP='') and (stTiers='') then TobAfficheA.Dupliquer(TobLigneA,true,true)
else begin
     for i_ind := 0 to TobLigneA.Detail.Count -1 do
         begin
         TobD := TobLigneA.Detail[i_ind];
         if (stNatureP <> '') and ( stNatureP <> TobD.GetValue('GL_NATUREPIECEG')) then continue;
         if (stTiers <> '')  and (stTiers <> Copy(TobD.GetValue('GL_TIERS'),1,Length(stTiers))) then continue;
         TobA := TOB.Create('',TobAfficheA,-1);
         TobA.Dupliquer(TobD,false,true);
         end;
     end;
TobAfficheA.PutGridDetail(GENT,false,false,stTitresLigne);
if iEntree = iLot then
   THNumEdit(GetControl('NCUMULENT')).Value := TobAfficheA.Somme('(QUANTITE)',[''],[''],true);
end;

procedure TOF_TRACABILITE.ActualiseGSOR;
var i_ind : integer;
    stNatureP, stTiers : string;
    TobD,TobA : TOB;
begin
//Vente
TobAfficheV.ClearDetail ; GSOR.VidePile(false) ;
if TobLigneV.Detail.Count = 0 then exit;
stNatureP := GetControlText('NATUREPIECEGVTE');
stTiers := GetControlText('CLIENT');
if (stNatureP='') and (stTiers='') then TobAfficheV.Dupliquer(TobLigneV,true,true)
else begin
     for i_ind := 0 to TobLigneV.Detail.Count -1 do
         begin
         TobD := TobLigneV.Detail[i_ind];
         if (stNatureP <> '') and (stNatureP <> TobD.GetValue('GL_NATUREPIECEG')) then continue;
         if (stTiers <> '')  and (stTiers <> Copy(TobD.GetValue('GL_TIERS'),1,Length(stTiers))) then continue;
         TobA := TOB.Create('',TobAfficheV,-1);
         TobA.Dupliquer(TobD,false,true);
         end;
     end;
TobAfficheV.PutGridDetail(GSOR,false,false,stTitresLigne);
if iEntree = iLot then
   THNumEdit(GetControl('NCUMULSOR')).Value := TobAfficheV.Somme('(QUANTITE)',[''],[''],true);
end;

procedure TOF_TRACABILITE.ActualiseGSTO;
var i_ind : integer;
    stEmplacement, stReserve, stPrepa : string;
    EnReserve,EnPrepa : TCheckBoxState;
    TobD,TobA : TOB;
begin
//Stock
TobAfficheS.ClearDetail ; GSTO.VidePile(false);
if TobStock.Detail.Count = 0 then exit;
stEmplacement := GetControlText('EMPLACEMENT');
EnReserve := TCheckbox(GetControl('ENRESERVECLI')).State;
if EnReserve = cbChecked then stReserve := 'X'
   else if EnReserve = cbunChecked then stReserve := '-'
        else stReserve := '';
EnPrepa := TCheckbox(GetControl('ENPREPACLI')).State;
if EnPrepa = cbChecked then stPrepa := 'X'
   else if EnPrepa = cbunChecked then stPrepa := '-'
        else stPrepa := '';
for i_ind := 0 to TobStock.Detail.Count -1 do
    begin
    TobD := TobStock.Detail[i_ind];
    if iEntree=iLot then
       begin
       if stEmplacement <> Copy(TobD.GetValue('GQ_EMPLACEMENT'),1,Length(stEmplacement)) then continue;
       end else
    if iEntree=iSerie then
       begin
       if (stReserve <> '') and (stReserve <> TobD.GetValue('GQS_ENRESERVECLI')) then continue;
       if (stPrepa <> '')   and (stPrepa <> TobD.GetValue('GQS_ENPREPACLI')) then continue;
       end;
    TobA := TOB.Create('',TobAfficheS,-1);
    TobA.Dupliquer(TobD,false,true);
    end;
TobAfficheS.PutGridDetail(GSTO,false,false,stTitresStock);
if iEntree = iLot then
   THNumEdit(GetControl('NCUMULSTO')).Value := TobAfficheS.Somme(PrefStock+'_PHYSIQUE',[''],[''],true);
end;

procedure TOF_TRACABILITE.ChargeLigneImpression(TobLigne,TobTmp : TOB);
var i_ind, iOrdre : integer;
    TOBD,TOBT : TOB;
begin
for i_ind := 0  to TobLigne.Detail.Count -1 do
    begin
    TOBT := TOB.Create('GCTMPTRACABILITE',TobTmp,-1);
    TOBD := TobLigne.Detail[i_ind];
    TOBT.PutValue('GZJ_UTILISATEUR',V_PGI.USer);
    inc(iCpte);
    TOBT.PutValue('GZJ_COMPTEUR',iCpte);
    TOBT.PutValue('GZJ_ARTICLE',TOBD.GetValue('GL_ARTICLE'));
    TOBT.PutValue('GZJ_LOTSERIE',TOBD.GetValue(PrefLigne+'_'+Reference));
    TOBT.PutValue('GZJ_DEPOT',TOBD.GetValue('GL_DEPOT'));
    TOBT.PutValue('GZJ_DATE',TOBD.GetValue('GL_DATEPIECE'));
    TOBT.PutValue('GZJ_NATUREPIECEG',TOBD.GetValue('GL_NATUREPIECEG'));
    if GetInfoParPiece(TOBD.GetValue('GL_NATUREPIECEG'),'GPP_SENSPIECE') = 'ENT' then iOrdre := 0 else iOrdre := 2;
    TOBT.PutValue('GZJ_NUMORDRE',iOrdre);
    TOBT.PutValue('GZJ_NUMERO',TOBD.GetValue('GL_NUMERO'));
    TOBT.PutValue('GZJ_TIERS',TOBD.GetValue('GL_TIERS'));
    TOBT.PutValue('GZJ_RAISONSOCIALE',TOBD.GetValue('T_LIBELLE'));
    if TOBD.FieldExists(PrefLigne+'_QUANTITE') then
       TOBT.PutValue('GZJ_QUANTITE',valeur(TOBD.GetValue(PrefLigne+'_QUANTITE')));
    end;
end;

////////////////////////////////////////////////////////////////////////////////
//Fiche
function TOF_TRACABILITE.RecupCriteres(stPrefixe : string) : string;
var stPrefRef : string;
begin
if stPrefixe = 'GL' then stPrefRef := PrefLigne else stPrefRef := stPrefixe;
if IdentArticle='' then result:=' WHERE '+stPrefRef+'_ARTICLE<>""'
                   else result:=' WHERE '+stPrefRef+'_ARTICLE="'+IdentArticle+'"';
if GetControlText('DEPOT') <> '' then
   result := result + ' AND ' + stPrefixe + '_DEPOT="' + GetControlText('DEPOT') + '"';
if GetControlText('REFERENCE') <> '' then
   result := result + ' AND ' + stPrefRef + '_' + Reference + '>="' + GetControlText('REFERENCE') + '"';
if GetControlText('REFERENCE_') <> '' then
   result := result + ' AND ' + stPrefRef + '_' + Reference + '<="' + GetControlText('REFERENCE_') + '"';
end;

Procedure TOF_TRACABILITE.ArticleElipsisClick(Sender : TObject);
var stWhere,CodeArt, StatutArt : string;
begin
CodeArt := GetControlText('CODEARTICLE');
StatutArt := '' ;  IdentArticle := '' ; stWhere := 'GA_LOT="X"';
DispatchRecherche(THCritMaskEdit(GetControl('CODEARTICLE')), 1, stWhere, 'GA_CODEARTICLE=' + CodeArt + ';RETOUR_CODEARTICLE=X', '');
RechercheDimensions;
OldArticle := GetControlText('CODEARTICLE');
end;

Procedure TOF_TRACABILITE.ArticleExit(Sender : TObject);
begin
if GetControlText('CODEARTICLE') = OldArticle then exit;
RechercheDimensions;
OldArticle := GetControlText('CODEARTICLE');
end;

procedure TOF_TRACABILITE.RechercheDimensions;
var CodeArt, StatutArt : string;
    QArt : TQuery;
    TobArt : TOB;
begin
CodeArt := GetControlText('CODEARTICLE');  IdentArticle := CodeArt ;
QArt :=  OpenSQL('SELECT GA_ARTICLE,GA_STATUTART FROM ARTICLE WHERE GA_CODEARTICLE="' + CodeArt
     + '" AND GA_STATUTART<>"DIM"',True);
if not QArt.Eof then
   begin
   IdentArticle := QArt.FindField('GA_ARTICLE').AsString;
   StatutArt := QArt.FindField('GA_STATUTART').AsString;
   end;
Ferme(QArt);
if StatutArt = 'GEN' then
   begin
   TobArt := TOB.Create('',nil,-1);
   if ChoisirDimension(IdentArticle,TobArt) then
      begin
      SetControlText('TLIBELLEDIM',RechercheLibelleDim(TobArt));
      IdentArticle := TobArt.GetValue('GA_ARTICLE');
      end;
   TobArt.Free;
   end else SetControlText('TLIBELLEDIM','');
end;

function TOF_TRACABILITE.RechercheLibelleDim(TobArt : TOB) : string;
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
begin
StDim:='';
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

////////////////////////////////////////////////////////////////////////////////
procedure TOF_TRACABILITE.AddChampListe(var NomListe,NomChamp : string);
var TokenListe,st : string;
begin
TokenListe := NomListe;
Repeat
    st := Trim(ReadTokenSt(TokenListe));
    if st = NomChamp then exit;
Until (TokenListe = '') or (st = '');
NomListe := NomListe + ';' + NomChamp;
end;

function TOF_TRACABILITE.PrepareImpression : integer;
var TobTemp,TOBT, TOBS : TOB;
    i_ind : integer;
begin
//////Nom table à renseigner
result := 0;
ExecuteSQL('DELETE FROM GCTMPTRACABILITE WHERE ' + PEtat.stSQL);
TobTemp := TOB.Create('',nil,-1);
iCpte := 0;
//Chargement des mouvements
ChargeLigneImpression(TobAfficheA,TobTemp);
ChargeLigneImpression(TobAfficheV,TobTemp);

//Chargement du stock
for i_ind := 0  to TobAfficheS.Detail.Count -1 do
    begin
    TOBT := TOB.Create('GCTMPTRACABILITE',TobTemp,-1);
    TOBS := TobAfficheS.Detail[i_ind];
    TOBT.InitValeurs();
    TOBT.PutValue('GZJ_UTILISATEUR',V_PGI.USer);
    inc(iCpte);
    TOBT.PutValue('GZJ_COMPTEUR',iCpte);
    TOBT.PutValue('GZJ_ARTICLE',TOBS.GetValue('GQ_ARTICLE'));
    TOBT.PutValue('GZJ_LOTSERIE',TOBS.GetValue(PrefStock+'_'+Reference));
    TOBT.PutValue('GZJ_DEPOT',TOBS.GetValue('GQ_DEPOT'));
    if TOBS.FieldExists(PrefStock+'_PHYSIQUE') then
       TOBT.PutValue('GZJ_QUANTITE',valeur(TOBS.GetValue(PrefStock+'_PHYSIQUE')));
    TOBT.PutValue('GZJ_NUMORDRE',1);
    end;

TobTemp.InsertOrUpdateDB(false);
TobTemp.Free;
end;

procedure TOF_TRACABILITE.AppelLaPiece(VenteAchat : string);
var TOBD : TOB;
    CleDoc : R_CleDoc;
begin
if VenteAchat = 'TABACHAT' then
   begin
   if TobAfficheA.Detail.Count < GENT.Row then exit;
   TOBD := Tob(GENT.Objects[0,GENT.Row]);
   end else if VenteAchat = 'TABVENTE' then
               begin
               if TobAfficheV.Detail.Count < GSOR.Row then exit;
               TOBD := Tob(GSOR.Objects[0,GSOR.Row]);
               end else exit;
if TOBD = nil then exit;
CleDoc.NaturePiece:=TOBD.GetValue('GL_NATUREPIECEG');
CleDoc.Souche:=TOBD.GetValue('GL_SOUCHE');
CleDoc.DatePiece:=TOBD.GetValue('GL_DATEPIECE');
CleDoc.NumeroPiece:=TOBD.GetValue('GL_NUMERO');
CleDoc.Indice:=TOBD.GetValue('GL_INDICEG');
SaisiePiece(CleDoc,taConsult)
end;

////////////////////////////////////////////////////////////////////////////////

procedure AGLRefreshGrid( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit;
if (TOTOF is TOF_TRACABILITE) then TOF_TRACABILITE(TOTOF).RefreshGrids else exit;
end ;

procedure TOF_TRACABILITE.BZoomDispoClick(Sender: TObject);
var Onglet : TTabSheet;
    TOBD : TOB;
begin
Onglet := GetActiveTabSheet('PAGECONTROL');
if Onglet.Name = 'TABSTOCK' then
   begin
   if TobAfficheS.Detail.Count < GSTO.Row then exit;
   TOBD := Tob(GSTO.Objects[0,GSTO.Row]);
   if TOBD = nil then exit;
   AGLLanceFiche('GC','GCDISPO','',TOBD.GetValue('GQ_ARTICLE') +';'+
   TOBD.GetValue('GQ_DEPOT')+';-;01/01/1900','ACTION=CONSULTATION') ;
   exit;
   end;
end;

procedure TOF_TRACABILITE.BZoomPieceClick(Sender: TObject);
begin
AppelLaPiece(TPageControl(GetControl('PAGECONTROL')).ActivePage.Name);
end;

procedure TOF_TRACABILITE.BZoomClientClick(Sender: TObject);
var TOBD : TOB;
begin
if TobAfficheV.Detail.Count < GSOR.Row then exit;
TOBD := Tob(GSOR.Objects[0,GSOR.Row]);
if TOBD = nil then exit;
if (TOBD.GetValue('T_AUXILIAIRE') <> null) and (TOBD.GetValue('T_AUXILIAIRE') <> '') then
   AGLLanceFiche('GC','GCTIERS','',TOBD.GetValue('T_AUXILIAIRE'),'ACTION=CONSULTATION;MONOFICHE') ;
end;

procedure TOF_TRACABILITE.BZoomFourClick(Sender: TObject);
var TOBD : TOB;
begin
if TobAfficheA.Detail.Count < GENT.Row then exit;
TOBD := Tob(GENT.Objects[0,GENT.Row]);
if TOBD = nil then exit;
if (TOBD.GetValue('T_AUXILIAIRE') <> null) and (TOBD.GetValue('T_AUXILIAIRE') <> '') then
   AGLLanceFiche('GC','GCFOURNISSEUR','',TOBD.GetValue('T_AUXILIAIRE'),'ACTION=CONSULTATION;MONOFICHE');
end;

procedure TOF_TRACABILITE.InitCumul(Grid : THGrid ; ColQte : integer);
var Coord : TRect;
    NCumul : THNumEdit;
    PCumul : THPanel;
    Suffixe : string;
begin
suffixe := Copy(Grid.Name,2,Length(Grid.Name));
NCumul := THNumEdit(GetControl('NCUMUL'+suffixe));
PCumul := THPanel(GetControl('PCUMUL'+suffixe));
NCumul.ParentColor := True;
NCumul.Font.Style := PCumul.Font.Style;
NCumul.Font.Size := PCumul.Font.Size;
NCumul.Masks.PositiveMask := Grid.ColFormats[ColQte];
NCumul.Ctl3D := False;
NCumul.Top := -1;
Coord := Grid.CellRect (ColQte, 0);
NCumul.Left := Coord.Left;
NCumul.Width := Grid.ColWidths[ColQte] + 1;
NCumul.Height := PCumul.Height;
NCumul.Enabled := False;
end;

procedure TOF_TRACABILITE.GridColumnWidthsChanged(Sender: TObject);
var Coord : TRect;
    Grid : THGrid;
    NCumul : THNumEdit;
    PCumul : THPanel;
    Suffixe : string;
    ColCli : integer;
begin
Grid := THGrid(Sender);
suffixe := Copy(Grid.Name,2,Length(Grid.Name));
NCumul := THNumEdit(GetControl('NCUMUL'+suffixe));
PCumul := THPanel(GetControl('PCUMUL'+suffixe));
if suffixe='STO' then ColCli := ColQtePhy else ColCli := ColQteFact;
Grid := THGrid(Sender);
if PCumul.ControlCount<=0 then exit ;
Coord:=Grid.CellRect(ColCli,0);
NCumul.Left:=Coord.Left + 1;
NCumul.Width:=Grid.ColWidths[ColCli] + 1;
end;


Initialization
  registerclasses ( [ TOF_TRACABILITE ] ) ;
  RegisterAglProc( 'AGLRefreshGrid', TRUE , 0, AGLRefreshGrid );
end.

