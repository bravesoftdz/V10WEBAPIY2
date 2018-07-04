{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/11/2001
Modifié le ... : 29/11/2001
Description .. : Source TOF de la TABLE : SAISIESERIE ()
Suite ........ : Saisie des numéros de série liés aux pièces
Mots clefs ... : TOF;SAISIESERIE;SERIE
*****************************************************************}
Unit SaisieSerie_TOF ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
      MaineAgl,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,Fe_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF,UTOB,Vierge,AglInit,Math,Windows,SaisUtil,LicUtil,StockUtil,ExtCtrls,
     HTB97,Hpanel,EntGc, graphics,Grids;

function Entree_SaisiSerie(TOBLigne,TOBP,TOBSerie,TobSerie_O,TobSerieReliquat : TOB; ActionF : TActionFiche) : boolean;

Type
  TOF_SAISIESERIE = Class (TOF)
    private
    BDroite,BGauche,BBas,BHaut,BSelectAll,BReliquat : TToolbarButton97;
    colRang : integer ;
    NbRowsInit : integer ;

    NbRowsPlus : integer ;
    SerStkCol : integer ;
    SerLigCol : integer ;
    SerRelCol : integer ;
    GStock: THGRID ;
    GLigne: THGRID ;
    GCompo: THGRID ;
    GReliq: THGRID ;
    SPLITTER : TSplitter;
    LesColStock,LesColLigne,LesColCompo,LesColReliq : string ;
    TOBSerLig,TOBSerLigRel,TOBSerStk,TOBSerStkArt : TOB ;
    TobTemp, TobTempRel : TOB;
    EntreeSerie,StockVide : Boolean;
    GereLot,GereReliquat : boolean;
    // Evènements du Grid
    procedure GligneCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GligneCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GligneRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GligneRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GCompoRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GCompoRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GStockDblClick(Sender: TObject);
    procedure GStockBeforeFlip( Sender : TObject ; ARow : Longint ; Var Cancel : Boolean ) ;
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    {$ENDIF}
    // Actions liées au Grid
    procedure EtudieColsListe ;
    procedure AfficheGrilleStock ;
    procedure ChargeGrille ;
    Function  GetTOBSStock ( ARow : integer) : TOB ;
    Function  GetTOBSLigne ( ARow : integer) : TOB ;
    Function  LigVide(Row : integer) : Boolean ;
    Procedure InitRow (Row : integer) ;
    Procedure CreerTOBPLigne (ARow : integer);
    Procedure CreerTOBStock (TOBSL : TOB);
    Procedure LigneVersSerie (TOBSL : TOB);
    Function  SortDeLaLigne : boolean ;
    procedure SupprimeLigne(ARow : integer);
    // Manipulation des Champs GRID
    procedure TraiterLigneSerie (ACol, ARow : integer; var Cancel : Boolean);
    // Evenements des boutons
    procedure BDroiteClick(Sender: TObject);
    procedure BGaucheClick(Sender: TObject);
    procedure BBasClick(Sender: TObject);
    procedure BHautClick(Sender: TObject);
    procedure BSelectAllClick(Sender: TObject);
    procedure BReliquatClick(Sender: TObject);
    procedure CreateSplitter;
    public
    Action   : TActionFiche ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

procedure ReaffecterLesSerie (TOBSerie,TobSerieReliquat : TOB ) ;
function GetQteConvert(TOBL,TOBLN : TOB ; Qte : double ; VersPiece : boolean = false) : double;

Implementation

var TOBL,TOBPlat,TOBSer,TOBSer_O,TobSerRel : TOB;
    QteUnitaire,QteReliquat,QteTotalCompose,QteTotalRelCompose : double;

function Entree_SaisiSerie(TOBLigne,TOBP,TOBSerie,TobSerie_O,TobSerieReliquat : TOB; ActionF : TActionFiche) : boolean;
var Arg,Retour : string;
    i_ind : integer;
    QteCompose, QteRelCompose : double;
    TobPlatD : TOB;
begin
result := false;
if (TOBLigne=Nil) then exit;
if (TOBSerie=Nil) then exit;
if GetInfoParPiece(TOBLigne.GetValue('GL_NATUREPIECEG'),'GPP_SENSPIECE')='MIX' then exit ;
TOBL:=TOBLigne;
TOBPlat:=TOBP;
QteUnitaire := 0 ; QteTotalCompose := 0 ; QteReliquat:=0 ; QteTotalRelCompose := 0;
if TobPlat <> nil then
   begin
   for i_ind := 0 to TobPlat.Detail.Count -1 do
       begin
       TobPlatD := TobPlat.Detail[i_ind];
       QteCompose := GetQteConvert(TOBL,TobPlatD,TobPlatD.GetValue('GLN_QTE'));
       QteRelCompose := GetQteConvert(TOBL,TobPlatD,TobPlatD.GetValue('QTERELIQUAT'));
       if (Frac(QteCompose) <> 0) or (Frac(QteRelCompose) <> 0) then
          begin
          PGIBox('La conversion de la quantité de l''un des composants ne permet pas d''obtenir un nombre entier d''article en quantité unitaire.#13'
          +'Il n''est donc pas possible d''affecter des numéros de série à cette nomenclature',
          'Affectation numéros de série');
          exit;
          end;
       QteTotalCompose := QteTotalCompose + QteCompose;
       QteTotalRelCompose := QteTotalRelCompose + QteRelCompose;
       end;
   end
else if TOBL.FieldExists('NUMEROLOT') then QteUnitaire:=TOBL.GetValue('QTELOT')
else  begin
      QteUnitaire := GetQteConvert(TOBL,nil,TOBL.GetValue('GL_QTESTOCK'));
      QteReliquat := GetQteConvert(TOBL,nil,TOBL.GetValue('GL_QTERELIQUAT'));
      if (Frac(QteUnitaire) <> 0) or (Frac(QteReliquat) <> 0) then
          begin
          PGIBox('La conversion de la quantité saisie ne permet pas d''obtenir un nombre entier d''article en quantité unitaire.#13'
          +'Il n''est donc pas possible d''affecter des numéros de série à cet article',
          'Affectation numéros de série');
          exit;
          end;
      end;
TOBSer_O := TobSerie_O;
TobSerRel := TOB.Create('', Nil,-1) ;
if TobSerieReliquat <> nil then TobSerRel.Dupliquer(TobSerieReliquat,True,True) ;
TOBSer := TOB.Create('', Nil,-1) ;
TOBSer.Dupliquer(TOBSerie,True,True) ;
Arg:=ActionToString(ActionF);
Retour:=AGLLanceFiche('GC','GCSAISISERIE','','',Arg);
if Retour='VALIDE' then
    BEGIN
    if ActionF<>taconsult then Begin ReaffecterLesSerie(TOBSerie,TobSerieReliquat); Result:=True; End ;
    END;
TOBSer.Free; TOBSerRel.Free;
end;

procedure ReaffecterLesSerie (TOBSerie,TobSerieReliquat : TOB ) ;
BEGIN
if TOBSer <> Nil then
   begin
   TOBSerie.ClearDetail;
   TOBSerie.Dupliquer(TOBSer,True,True) ;
   end;
if (TobSerieReliquat <> nil) and (TOBSerRel <> nil) then
   begin
   TobSerieReliquat.ClearDetail;
   TobSerieReliquat.Dupliquer(TOBSerRel,True,True)  ;
   end;
END;

function GetQteConvert(TOBL,TOBLN : TOB ; Qte : double ; VersPiece : boolean = false) : double;
var VenteAchat,QualQte,QualQteN : string;
    UnitePiece : double;
    TOBM : TOB;
begin
if TOBLN <> nil then
VenteAchat := GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT');
if VenteAchat='ACH' then
   begin
   QualQte:=TOBL.GetValue('GL_QUALIFQTEACH');
   if TOBLN <> nil then QualQteN := TOBLN.GetValue('GLN_QUALIFQTEACH');
   end else
if (VenteAchat='AUT') or (VenteAchat='TRF') then
   begin
   QualQte:=TOBL.GetValue('GL_QUALIFQTESTO');
   if TOBLN <> nil then QualQteN := TOBLN.GetValue('GLN_QUALIFQTESTO');
   end else
   begin
   QualQte:=TOBL.GetValue('GL_QUALIFQTEVTE');
   if TOBLN <> nil then QualQteN := TOBLN.GetValue('GLN_QUALIFQTEVTE');
   end;
TOBM:=VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'],['PIE',QualQte],False) ;
UnitePiece:=0 ; if TOBM<>Nil then UnitePiece:=TOBM.GetValue('GME_QUOTITE') ;
if UnitePiece=0 then UnitePiece:=1.0 ;
TOBM:=VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'],['PIE',QualQteN],False) ;
if TOBM<>Nil then
   if TOBM.GetValue('GME_QUOTITE') <> 0 then UnitePiece := UnitePiece*TOBM.GetValue('GME_QUOTITE') ;
if VersPiece then result:=Arrondi(Qte/UnitePiece, 6)
else result:=Arrondi(Qte*UnitePiece, 6);
end;

{==============================================================================================}
{================================= Evénement de la TOF ========================================}
{==============================================================================================}
procedure TOF_SAISIESERIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SAISIESERIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SAISIESERIE.OnUpdate ;
Var OkSerie : boolean ;
    i_ind : integer;
    QteAffecte,QtePiece : double;
begin
  Inherited ;
OkSerie:=True;
if Not SortDeLaLigne then Begin TFVierge(Ecran).ModalResult:=0; Exit ; end;
QteAffecte := TOBSerLig.Detail.Count;
if TobPlat <> nil then
   begin
   QtePiece := QteTotalCompose ; QteAffecte := QteAffecte + TobTemp.Detail.Count ;
   end else QtePiece := QteUnitaire;
if QtePiece = 0 then    //cas de la qté à 0 si reliquat
   if TOBSerLig.Detail.Count=1 then
      if TobSerLig.Detail[0].GetValue('GLS_IDSERIE')='' then
         begin
         TobSerLig.Detail[0].Free;
         QteAffecte := QteAffecte -1;
         end;

if QteAffecte > QtePiece then
   begin
   PGIBox('Le nombre de numéros affectés est supérieur '#13'à la quantité saisie',Ecran.Caption);
   Ecran.ModalResult := 0;
   Exit;
   end;
if QteAffecte < QtePiece then OkSerie:=False;
if TOBSerLig.FindFirst(['GLS_IDSERIE'],[''],true) <> nil then OkSerie := False;
if TobPlat <> nil then
   if TOBTemp.FindFirst(['GLS_IDSERIE'],[''],true) <> nil then OkSerie := False;
if Not OkSerie then
   begin
   PGIBox('La liste des numéros de série est incomplète',Ecran.Caption);
   Ecran.ModalResult := 0;
   Exit;
   end;

if GereReliquat then
    begin
    QteAffecte := TobSerLigRel.Detail.Count;
    if TobPlat <> nil then
       begin
       QtePiece := QteTotalRelCompose;
       QteAffecte := QteAffecte + TobTempRel.Detail.Count ;
       end else QtePiece := QteReliquat;
    if QteAffecte > QtePiece then
       begin
       PGIBox('Le nombre de numéros affectés en reliquat est supérieur '#13'à la quantité saisie',Ecran.Caption);
       Ecran.ModalResult := 0;
       Exit;
       end;
    if QteAffecte < QtePiece then
       begin
       PGIBox('La liste des numéros affectés en reliquat est incomplète',Ecran.Caption);
       Ecran.ModalResult := 0;
       Exit;
       end;
    end;
for i_ind := TobTemp.Detail.Count -1 downto 0 do
    TobTemp.Detail[i_ind].ChangeParent(TobSerLig,-1);
for i_ind := TobTempRel.Detail.Count -1 downto 0 do
    TobTempRel.Detail[i_ind].ChangeParent(TobSerLigRel,-1);
TFVierge(Ecran).Retour:='VALIDE' ;
end ;

procedure TOF_SAISIESERIE.OnLoad ;
begin
  Inherited ;
ChargeGrille;
end ;

procedure TOF_SAISIESERIE.OnArgument (S : String ) ;
Var St,NatPiece,Sens,PiecePreced : string ;
    EstAvoir : Boolean ;
    i : integer;
begin
  Inherited ;
St:=S ;
i:=Pos('ACTION=',St) ;
if i>0 then
   BEGIN
   System.Delete(St,1,i+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='CREATION' then BEGIN Action:=taCreat ; END ;
   if St='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if St='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;

BDroite:=TToolbarButton97(GetControl('BDROITE'));
BDroite.OnClick:=BDroiteCLick;
BGauche:=TToolbarButton97(GetControl('BGAUCHE'));
BGauche.OnClick:=BGaucheCLick;
BBas:=TToolbarButton97(GetControl('BBAS'));
BBas.OnClick:=BBasCLick;
BHaut:=TToolbarButton97(GetControl('BHAUT'));
BHaut.OnClick:=BHautCLick;
BSelectAll:=TToolbarButton97(GetControl('BSELECTALL'));
BSelectAll.OnClick:=BSelectAllCLick;
BReliquat:=TToolbarButton97(GetControl('BRELIQUAT'));
BReliquat.OnClick:=BReliquatCLick;
if Action=taConsult then
   begin
   BDroite.Enabled:=False;
   BGauche.Enabled:=False;
   BBas.Enabled:=False;
   BHaut.Enabled:=False;
   BSelectAll.Enabled:=False;
   BReliquat.Enabled:=False;
   end;

GStock:=THGRID(GetControl('GSTOCK'));
GStock.ListeParam:='GCSERIESTOCK' ;
Gligne:=THGRID(GetControl('GLIGNE'));
Gligne.ListeParam:='GCSERIELIGNE' ;
GCompo:=THGRID(GetControl('GCOMPO'));
GCompo.ListeParam:='GCSERIECOMPO' ;
GReliq:=THGRID(GetControl('GRELIQUAT'));
GReliq.ListeParam:='GCSERIELIGNE' ;

colRang := 1 ; NbRowsInit := 50 ; NbRowsPlus := 20 ;
SerStkCol := 0 ; SerLigCol := 0 ; SerRelCol := 0;
EtudieColsListe;

TOBSerStk:=TOB.Create('',Nil,-1) ;
TOBSerStkArt:=TOB.Create('',Nil,-1) ;
TobTemp := TOB.Create('',nil,-1);
TobTempRel := TOB.Create('',nil,-1);

SetControlText('HDepot',RechDom('GCDEPOT', TOBL.GetValue('GL_DEPOT'), false));

GereLot := TOBL.FieldExists('NUMEROLOT');
GereReliquat := false;
PiecePreced := TOBL.GetValue('GL_PIECEPRECEDENTE');
ReadTokenSt(PiecePreced) ; st := ReadTokenSt(PiecePreced);
if (st <> '') and (TOBSerRel.Detail.Count>0) then
   if GetInfoParpiece(st,'GPP_NUMEROSERIE')='X' then
       GereReliquat := (TOBL.GetValue('GL_QTERELIQUAT') <> 0) or
                       (TOBSerRel.Detail[TOBL.GetValue('GL_INDICESERIE')-1].Detail.Count>0);

if GereLot then SetControlText('HLot',TOBL.GetValue('NUMEROLOT'))
else begin SetControlText('HLot',''); SetControlCaption('TLOT',''); end;

if TOBPlat=Nil then
   begin
   SetControlText('HArticle',TOBL.GetValue('GL_CODEARTICLE')+'  '+TOBL.GetValue('GL_LIBELLE'));
   SetControlVisible('TCOMPO',False);
   end else
   begin
   SetControlProperty('TCOMPO','Height',Ecran.Height);
   SetControlText('COMPOSE',TOBL.GetValue('GL_CODEARTICLE'));
   SetControlText('LIBCOMPOSE',TOBL.GetValue('GL_LIBELLE'));
   end;

if GereReliquat then
    begin
    SetControlVisible('BRELIQUAT',True);
    SetControlVisible('PRELIQUAT',True);
    CreateSplitter;
    GReliq.Cells[SerRelCol,0] := 'Reliquat';
    end else GLIGNE.Align := alClient;

Gligne.OnCellEnter:=GligneCellEnter ;
Gligne.OnCellExit:=GligneCellExit ;
Gligne.OnRowEnter:=GligneRowEnter ;
Gligne.OnRowExit:=GligneRowExit ;
GStock.OnDblClick := GStockDblClick ;  GStock.OnBeforeFlip := GStockBeforeFlip;
GCompo.OnRowEnter:=GCompoRowEnter ; GCompo.OnRowExit:=GCompoRowExit ;
AffecteGrid(GStock,taConsult) ;
AffecteGrid(Gligne,Action) ;
AffecteGrid(GCompo,taConsult) ;
AffecteGrid(GReliq,taConsult) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GStock) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(Gligne) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GCompo) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GReliq) ;
SetControlText('QTESTOCK',FloatToStr(QteUnitaire));
NatPiece:=TOBL.GetValue('GL_NATUREPIECEG');
Sens:=GetInfoParPiece(NatPiece,'GPP_SENSPIECE') ;
EstAvoir:=(GetInfoParPiece(NatPiece,'GPP_ESTAVOIR')='X') ;
if Sens='ENT' then EntreeSerie:=Not (EstAvoir) else EntreeSerie:=EstAvoir;
if EntreeSerie then
   begin
   SetControlVisible('GSTOCK',False);
   SetControlVisible('TSERIESTOCK',False);
   SetControlText('QUANTITE',FloatToStr(QteUnitaire));
   SetControlProperty('QUANTITE','MAXVALUE',QteUnitaire);
   SetControlVisible('BSELECTALL',False);
   end else
   begin
   SetControlVisible('TATTRIBUTION',False);
   SetControlVisible('PREFIXE',False);
   SetControlVisible('SUFFIXE',False);
   SetControlVisible('TPREFIXE',False);
   SetControlVisible('TSUFFIXE',False);
   SetControlVisible('NUMDEPART',False);
   SetControlVisible('TNUMDEPART',False);
   end;
end ;

procedure TOF_SAISIESERIE.OnClose ;
begin
  Inherited ;
GStock.VidePile(False) ;
Gligne.VidePile(False) ;
TOBSerStk.Free ; TOBSerStkArt.Free ;
TobTemp.Free;
TobTempRel.Free;
if GereReliquat then
   if bReliquat.Visible and bReliquat.Down then SPLITTER.Free;
end ;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}
procedure TOF_SAISIESERIE.GligneCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
BEGIN

END;

procedure TOF_SAISIESERIE.GligneCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
BEGIN
if Action=taConsult then Exit ;
if ACol = SerLigCol then TraiterLigneSerie (ACol, ARow, Cancel) ;
END;

procedure TOF_SAISIESERIE.GligneRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
ARow := Min (Ou, TOBSerLig.Detail.Count + 1);
if (ARow = TOBSerLig.Detail.Count + 1) AND (not LigVide (ARow - 1)) then
    BEGIN
    CreerTOBPLigne (ARow);
    END;
if Ou > TOBSerLig.Detail.Count then
    BEGIN
    Gligne.Row := TOBSerLig.Detail.Count;
    END;
END;

procedure TOF_SAISIESERIE.GligneRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN

END;

procedure TOF_SAISIESERIE.GCompoRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN

END;

procedure TOF_SAISIESERIE.GCompoRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
if not SortDeLaLigne then GCompo.Row := Ou else ChargeGrille;
END;

procedure TOF_SAISIESERIE.GStockDblClick(Sender: TObject);
var ARow,i_ind : integer;
    Cancel : boolean;
    TOBSK : TOB;
BEGIN
if QteUnitaire <= StrToFloat(GetControlText('QTEAFFECT')) then
   if GereReliquat then GLigne.RowCount := GLigne.RowCount + 1 else exit;
ARow := GStock.Row;
Cancel:=False;
for i_ind:=1 to GLigne.RowCount do
   begin
   if LigVide (i_ind) then
      begin
      if (i_ind = TOBSerLig.Detail.Count + 1) then CreerTOBPLigne (i_ind);
      GLigne.Cells [SerLigCol, i_ind] := GStock.Cells [SerStkCol, ARow];
      TraiterLigneSerie(SerLigCol,i_ind,Cancel);
      If Cancel then Begin GLigne.Cells [SerLigCol, i_ind] := ''; break; end;
      break;
      end;
   end;
GStock.AllSelected := false;
for i_ind:=1 to GStock.RowCount do
   begin
   TOBSK:=GetTOBSStock(i_ind);
   if TOBSK<>Nil then
      if TOBSK.FieldExists('SELECTED') then
         if TOBSK.GetValue('SELECTED') then GStock.FlipSelection(i_ind);
   end;
END;

procedure TOF_SAISIESERIE.GStockBeforeFlip( Sender : TObject ; ARow : Longint ; Var Cancel : Boolean ) ;
begin
if TobSerStk.Detail.Count <= 0 then exit;
if TobSerStk.Detail[ARow-1].FieldExists('SELECTED') then
   TobSerStk.Detail[ARow-1].PutValue('SELECTED', not GStock.isSelected(ARow));
end;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
procedure TOF_SAISIESERIE.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp: integer ;
BEGIN
LesColCompo:=GCompo.Titres[0] ; //ReadTokenSt(LesColCompo);//supp de la fixedrow
LesCols:=GStock.Titres[0] ;
LesColStock:=LesCols ; icol:=1 ; //ReadTokenSt(LesColStock);
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        BEGIN
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
            BEGIN
            if NomCol='GQS_IDSERIE'   then SerStkCol:=icol ;
            END;
        END;
    Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
GStock.ColLengths[SerStkCol]:=35;

LesCols:=GLigne.Titres[0] ;
LesColLigne:=LesCols ; icol:=1 ;  //ReadTokenSt(LesColLigne);
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        BEGIN
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
            BEGIN
            if NomCol='GLS_IDSERIE'   then SerLigCol:=icol ;
            END;
        END;
    Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
GLigne.ColLengths[SerLigCol]:=35;

LesCols:=GReliq.Titres[0] ;
LesColReliq:=LesCols ; icol:=1 ;  //ReadTokenSt(LesColReliq);
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        BEGIN
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
            BEGIN
            if NomCol='GLS_IDSERIE'   then SerRelCol:=icol ;
            END;
        END;
    Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
GReliq.ColLengths[SerRelCol]:=35;
END;

procedure TOF_SAISIESERIE.AfficheGrilleStock ;
begin
if not EntreeSerie then
   begin
   TOBSerStk.Detail.Sort('GQS_IDSERIE');
   TOBSerStk.PutGridDetail(GStock,False,False,LesColStock,True);
   if TOBSerStk.Detail.Count>0 then
      begin
      GStock.RowCount:=TOBSerStk.Detail.Count+1;
      SetControlEnabled('BDROITE',(Action<>taConsult));
      end else
      begin
      GStock.RowCount:=TOBSerStk.Detail.Count+2 ;
      StockVide:=True;
      SetControlEnabled('BDROITE',False);
      end;
   end;
end;

procedure TOF_SAISIESERIE.ChargeGrille ;
Var i_ind : integer ;
    TOBS,TOBSA,TOBP,TOBS_O,TOBSL : TOB;
    RefUnique,Depot,NumSerie : string ;
BEGIN
if TOBPLat=Nil then
   begin
   RefUnique:=TOBL.GetValue('GL_ARTICLE') ;
   end else
   begin
   TOBPLat.PutGridDetail(GCompo,False,False,LesColCompo,false);
   GCompo.RowCount:=TOBPLat.Detail.Count+1 ;
   i_ind:=GCompo.Row-1;
   TOBP := TOBPLat.Detail[i_ind];
   RefUnique:=TOBP.GetValue('GLN_ARTICLE') ;
   SetControlText('HArticle',TOBP.GetValue('GLN_CODEARTICLE')+'  '+
                   TOBP.GetValue('GLN_LIBELLE'));
   QteUnitaire := GetQteConvert(TOBL,TOBP,TOBP.GetValue('GLN_QTE'));
   QteReliquat := GetQteConvert(TOBL,TOBP,TOBP.GetValue('QTERELIQUAT'));
   SetControlText('QTESTOCK',FloatToStr(QteUnitaire));
   end;
Depot:=TOBL.GetValue('GL_DEPOT') ;
TOBSerStkArt.ClearDetail ; TOBSerStk.ClearDetail;
TOBSerStkArt.LoadDetailFromSQL('SELECT * FROM DISPOSERIE WHERE GQS_ARTICLE="'+RefUnique+'"',True);
for i_ind:=0 to TOBSerStkArt.Detail.Count-1 do
    begin
    TOBSA:=TOBSerStkArt.Detail[i_ind];
    if TOBSA.GetValue('GQS_DEPOT')<>Depot then continue;
    if GereLot then
      if TOBSA.GetValue('GQS_NUMEROLOT')<>TOBL.GetValue('NUMEROLOT') then continue;
    if (TOBSA.GetValue('GQS_ENRESERVECLI')='X') or (TOBSA.GetValue('GQS_ENPREPACLI')='X') then continue;
    NumSerie:=TOBSA.GetValue('GQS_IDSERIE');
    if TOBSer.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[RefUnique,NumSerie],True)<>Nil then continue ;
    if TOBTemp.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[RefUnique,NumSerie],True)<>Nil then continue ;
    if TOBTempRel.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[RefUnique,NumSerie],True)<>Nil then continue ;
    TOBS:=TOB.Create('DISPOSERIE',TOBSerStk,-1);
    TOBS.Dupliquer(TOBSA,True,True) ;
    TOBS.AddChampSupValeur('SELECTED',false,false);
    end;

if TOBSer_O <> nil then
   begin
//   TOBS_O := TOBSer_O.FindFirst(['GLS_ARTICLE','GLS_DEPOT','GLS_TENUESTOCK'],[RefUnique,Depot,'X'],True);
   TOBS_O := TOBSer_O.FindFirst(['GLS_ARTICLE','GLS_TENUESTOCK'],[RefUnique,'X'],True);
   While TOBS_O <> nil do
       begin
       if TOBSer.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[RefUnique,TOBS_O.GetValue('GLS_IDSERIE') ],True) = nil then
          CreerTOBStock(TOBS_O);
//          TOBS_O := TOBSer_O.FindNext(['GLS_ARTICLE','GLS_DEPOT','GLS_TENUESTOCK'],[RefUnique,Depot,'X'],True);
          TOBS_O := TOBSer_O.FindNext(['GLS_ARTICLE','GLS_TENUESTOCK'],[RefUnique,'X'],True);
       end;
   end;

AfficheGrilleStock ;
i_ind:=TOBL.GetValue('GL_INDICESERIE');
if TOBSer.Detail.Count>=i_ind then
   begin
   TOBSerLig:=TOBSer.Detail[i_ind-1];
   if (TOBSerRel.Detail.Count>0) then TobSerLigRel:=TOBSerRel.Detail[i_ind-1] else TobSerLigRel:=Nil;
   if (GereLot) or (GereReliquat) or (TobPlat <> nil) then
      begin
      for i_ind := TobTemp.Detail.Count -1 downto 0 do
          begin
          if TobTemp.Detail[i_ind].GetValue('GLS_IDSERIE') = '' then
               TobTemp.Detail[i_ind].Free
          else TobTemp.Detail[i_ind].ChangeParent(TOBSerLig,-1);
          end;
      for i_ind := TobTempRel.Detail.Count -1 downto 0 do
          TobTempRel.Detail[i_ind].ChangeParent(TOBSerLigRel,-1);
      for i_ind := TOBSerLig.Detail.Count -1 downto 0 do
          begin
          TOBSL := TobSerLig.Detail[i_ind];
          if GereLot then
             if TOBSL.GetValue('GLS_NUMEROLOT') <> TOBL.GetValue('NUMEROLOT') then
                 TOBSL.ChangeParent(TobTemp,-1);
          if TobPlat <> nil then
             begin
             if TOBSL.GetValue('GLS_ARTICLE') <> RefUnique then
                 if  TOBSL.GetValue('GLS_IDSERIE') = '' then TOBSL.Free
                  else TOBSL.ChangeParent(TobTemp,-1);
             end;
          end;
      if TobSerLigRel<>Nil then
         begin
         for i_ind := TobSerLigRel.Detail.Count -1 downto 0 do
             begin
             if TobSerLigRel.Detail[i_ind].GetValue('GLS_ARTICLE') <> RefUnique then
                TobSerLigRel.Detail[i_ind].ChangeParent(TobTempRel,-1);
             end;
          end;
      end;
   SetControlText('QTEAFFECT',intToStr(TobSerLig.Detail.Count));
   end else
   begin
   TOBSerLig:=TOB.Create('',TOBSer,-1) ;TOBSerLigRel:=TOB.Create('',TOBSerRel,-1) ;
   SetControlText('QTEAFFECT','0');
   end;
GLigne.VidePile(False);
if Tobserlig.Detail.Count <= 0 then CreerTOBPLigne(1);
TOBSerLig.PutGridDetail(GLigne,False,False,LesColLigne,false);
if GereReliquat then
   begin
   GReliq.VidePile(False);
   TobSerLigRel.PutGridDetail(GReliq,False,False,LesColReliq);
   end;
GLigne.RowCount:=Max(TOBSerLig.Detail.Count+1,trunc(QteUnitaire)+1) ;
END;

Function TOF_SAISIESERIE.GetTOBSStock ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBSerStk.Detail.Count)) then Exit ;
Result:=TOBSerStk.Detail[ARow-1] ;
END ;

Function TOF_SAISIESERIE.GetTOBSLigne ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBSerLig.Detail.Count)) then Exit ;
Result:=TOBSerLig.Detail[ARow-1] ;
END ;

Function TOF_SAISIESERIE.LigVide(Row : integer) : Boolean ;
BEGIN
Result:=True ;
if (GLigne.Cells[SerLigCol,Row]<>'') then result:= False ;
END ;

Procedure TOF_SAISIESERIE.InitRow (Row : integer) ;
var Col : integer ;
    TOBSL : TOB;
BEGIN
TOBSL:=GetTOBSLigne(Row) ;
if TOBSL<>Nil then
   begin
   TOBSL.InitValeurs ;
   LigneVersSerie(TOBSL);
   end;
for Col:=0 to GLigne.ColCount do GLigne.cells[Col,Row]:='';
END ;

Procedure TOF_SAISIESERIE.CreerTOBPLigne (ARow : integer);
BEGIN
if ARow <> TOBSerLig.Detail.Count + 1 then exit;
TOB.Create ('LIGNESERIE', TOBSerLig, ARow-1) ;
InitRow (ARow) ;
END;

Procedure TOF_SAISIESERIE.CreerTOBStock (TOBSL : TOB);
var TOBS : TOB;
BEGIN
if TOBSL = nil then exit;
TOBS:=TOB.Create('DISPOSERIE',TOBSerStk,-1);
TOBS.PutValue('GQS_ARTICLE',TOBSL.GetValue('GLS_ARTICLE')) ;
TOBS.PutValue('GQS_IDSERIE',TOBSL.GetValue('GLS_IDSERIE')) ;
TOBS.PutValue('GQS_DEPOT',TOBL.GetValue('GL_DEPOT')) ;
TOBS.PutValue('GQS_NUMEROLOT',TOBSL.GetValue('GLS_NUMEROLOT')) ;
TOBS.AddChampSupValeur('SELECTED',false,false);
END;

Procedure TOF_SAISIESERIE.LigneVersSerie (TOBSL : TOB);
var RefUnique : string;
BEGIN
TOBSL.PutValue('GLS_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
TOBSL.PutValue('GLS_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
TOBSL.PutValue('GLS_NUMERO',TOBL.GetValue('GL_NUMERO'));
TOBSL.PutValue('GLS_INDICEG',TOBL.GetValue('GL_INDICEG'));
TOBSL.PutValue('GLS_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE'));
//TOBSL.PutValue('GLS_DEPOT',TOBL.GetValue('GL_DEPOT'));
if TobPlat <> nil then
     RefUnique := TOBPLat.Detail[GCompo.Row-1].GetValue('GLN_ARTICLE')
else RefUnique := TOBL.GetValue('GL_ARTICLE');
TOBSL.PutValue('GLS_ARTICLE',RefUnique);
if GereLot then TOBSL.PutValue('GLS_NUMEROLOT',TOBL.GetValue('NUMEROLOT'));
if EntreeSerie then
   begin
   if TOBL.GetValue('GL_TENUESTOCK')='X' then TOBSL.PutValue('GLS_TENUESTOCK','X')
                                        else TOBSL.PutValue('GLS_TENUESTOCK','-');
   end else TOBSL.PutValue('GLS_TENUESTOCK','-');
END;

Function TOF_SAISIESERIE.SortDeLaLigne : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
ACol:=GLigne.Col ; ARow:=GLigne.Row ; Cancel:=False ;
GLigneCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
GLigneRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
Result:=True ;
END ;

procedure TOF_SAISIESERIE.SupprimeLigne(ARow : integer);
Var QteAff : Double ;
begin
QteAff:=StrToFloat(GetControlText('QTEAFFECT'));
if QteAff > 0 then QteAff:=QteAff-1 ;
SetControlText('QTEAFFECT',FloatToStr(QteAff));
if GLigne.RowCount > 2 then
   begin
   GLigne.CacheEdit ; GLigne.SynEnabled := False;
   GLigne.DeleteRow (ARow);
   GLigne.MontreEdit; GLigne.SynEnabled := True;
   end else CreerTOBPLigne(ARow);
if QteUnitaire > GLigne.RowCount-1 then
   GLigne.RowCount := GLigne.RowCount + 1;
end;

{==============================================================================================}
{============================ Manipulation des Champs GRID ====================================}
{==============================================================================================}
procedure TOF_SAISIESERIE.TraiterLigneSerie (ACol, ARow : integer; var Cancel : Boolean);
var TOBSL,TOBP : TOB;
    NumSerie,St1,RefUnique : string ;
    TenueStk : Boolean ;
    QteAff : Double ;
BEGIN
TOBSL := GetTOBSLigne (ARow); if TOBSL=nil then exit;
NumSerie := UpperCase(Trim(GLigne.Cells [ACol, ARow]));
GLigne.Cells [ACol, ARow] := NumSerie;
TenueStk:=(TOBSL.GetValue('GLS_TENUESTOCK')='X') ;
if NumSerie='' then
    BEGIN
    {NEWPIECE}
    if (QteUnitaire < StrToFloat(GetControlText('QTEAFFECT'))) and (TOBSL.GetValue('GLS_RANG') = 0) then
       begin
       if (TenueStk) and ( not EntreeSerie) then CreerTOBStock (TOBSL);
       AfficheGrilleStock ;
       TOBSL.Free ; SupprimeLigne(ARow);
       end else GLigne.Cells [ACol, ARow]:= TOBSL.GetValue ('GLS_IDSERIE') ;
    exit ;
    END;
if TOBSL.GetValue ('GLS_IDSERIE')<>NumSerie then
    BEGIN
    {NEWPIECE}
    if TOBSL.GetValue('GLS_RANG') = 1 then
       begin
       GLigne.Cells [ACol, ARow]:= TOBSL.GetValue ('GLS_IDSERIE') ;
       PGIBox('Ce numéro a été affecté sur une autre pièce,#13vous ne pouvez plus le remettre en stock',Ecran.Caption);
       exit;
       end;
    RefUnique:=TOBSL.GetValue('GLS_ARTICLE');
    if (TOBSer.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[RefUnique,NumSerie],True)<>Nil) or
       (TOBSerRel.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[RefUnique,NumSerie],True)<>Nil) then
        BEGIN
        St1:='0;?caption?;Ce numéro de série est déjà affecté;X;O;O;O;';
        HShowMessage(St1, Ecran.Caption, '');
        Cancel:=True;
        Exit;
        END;
    if GereLot then if TobTemp.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[RefUnique,NumSerie],True)<>Nil then
        BEGIN
        St1:='0;?caption?;Ce numéro de série est déjà affecté sur un autre lot;X;O;O;O;';
        HShowMessage(St1, Ecran.Caption, '');
        Cancel:=True;
        Exit;
        END;
    TOBP:=TOBSerStk.FindFirst(['GQS_ARTICLE','GQS_IDSERIE'],[RefUnique,NumSerie],True) ;
    if TOBP<>Nil then
       begin
       if not EntreeSerie then
          begin
          TOBSL.PutValue('GLS_TENUESTOCK','X') ;
          TOBP.Free;
          end else
          begin
          St1:='0;?caption?;Ce numéro de série existe déjà dans ce dépôt;X;O;O;O;';
          HShowMessage(St1, Ecran.Caption, '');
          Cancel:=True;
          Exit;
          end;
       end else
       begin
       if TOBSerStkArt.FindFirst(['GQS_ARTICLE','GQS_IDSERIE'],[RefUnique,NumSerie],True)<>Nil then
          begin
          St1:='0;?caption?;Ce numéro de série existe dans un autre dépôt;X;O;O;O;';
          HShowMessage(St1, Ecran.Caption, '');
          Cancel:=True;
          Exit;
          end;
       if not EntreeSerie then TOBSL.PutValue('GLS_TENUESTOCK','-') ;
       end ;
    if (TenueStk) and ( not EntreeSerie) then CreerTOBStock (TOBSL);
    AfficheGrilleStock ;
    if TOBSL.GetValue('GLS_IDSERIE')='' then
       begin
       QteAff:=StrToFloat(GetControlText('QTEAFFECT'))+1 ;
       SetControlText('QTEAFFECT',FloatToStr(QteAff));
       end;
    TOBSL.PutValue('GLS_IDSERIE',NumSerie) ;
    END;
END;

{==============================================================================================}
{================================ Evenements des boutons ======================================}
{==============================================================================================}
procedure TOF_SAISIESERIE.BDroiteClick(Sender: TObject);
Var Pref,Cpte,Suff,St : string ;
    NumDebut,NbSerie,NbSerieAffect,NbMax,i_ind : integer;
    Cancel : boolean;
    TOBS : TOB;
begin
if Action=taConsult then Exit;
if EntreeSerie then
   begin
   Pref:=GetControlText('PREFIXE');
   Suff:=GetControlText('SUFFIXE');
   NumDebut:=StrToInt (GetControlText('NUMDEPART'));
   NbSerie:=GLigne.RowCount-1 ;
   Cancel:=False; NbSerieAffect:=0; NbMax:=StrToInt(GetControlText('QUANTITE'));
   for i_ind:=1 to NbSerie do
       begin
       if LigVide (i_ind) then inc(NbSerieAffect);
       if NbSerieAffect >= NbMax then break;
       end;
   if NbSerieAffect=0 then exit;
   St:=IntToStr(NbSerieAffect+NumDebut-1);   NbSerieAffect:=0;
   for i_ind:=1 to NbSerie do
       begin
       if NbSerieAffect >= NbMax then break;
       if LigVide (i_ind) then
          begin
          if (i_ind = TOBSerLig.Detail.Count + 1) then CreerTOBPLigne (i_ind);
          Cpte:=Format('%.'+IntToStr(length(St))+'d',[NumDebut]);
          GLigne.Cells [SerLigCol, i_ind] := Pref+Cpte+Suff;
          TraiterLigneSerie(SerLigCol,i_ind,Cancel);
          If Cancel then Begin GLigne.Cells [SerLigCol, i_ind] := ''; break; end;
          inc(NumDebut);  inc(NbSerieAffect);
          end;
       end;
  end else
   begin
   TOBS := TobSerStk.FindFirst(['SELECTED'],[True],true);
   while TOBS <> nil do
       begin
       if QteUnitaire <= StrToFloat(GetControlText('QTEAFFECT')) then
          if GereReliquat then GLigne.RowCount := GLigne.RowCount + 1 else break;
       St:=TOBS.GetValue('GQS_IDSERIE');
       NbSerie:=GLigne.RowCount ;
       Cancel:=False;
       for i_ind:=1 to NbSerie do
           begin
           if LigVide (i_ind) then
              begin
              if (i_ind = TOBSerLig.Detail.Count + 1) then CreerTOBPLigne (i_ind);
              GLigne.Cells [SerLigCol, i_ind] := St;
              TraiterLigneSerie(SerLigCol,i_ind,Cancel);
              If Cancel then Begin GLigne.Cells [SerLigCol, i_ind] := ''; break; end;
              break;
              end;
           end;
       TOBS := TobSerStk.FindNext(['SELECTED'],[True],true);
       end;
   BSelectAll.Down := False ; GStock.AllSelected := false;
   end;
end;

procedure TOF_SAISIESERIE.BGaucheClick(Sender: TObject);
var TOBSL : TOB;
    ARow : integer;
begin
if Action=taConsult then Exit;
ARow := GLigne.Row;
TOBSL := GetTOBSLigne (ARow);
if TOBSL=nil then exit;
if ARow > TOBSerLig.Detail.Count then exit;
if (not EntreeSerie) then
  begin
  {NEWPIECE}
  if TOBSL.GetValue('GLS_RANG') = 1 then
    begin
    PGIBox('Ce numéro a été affecté sur une autre pièce,#13vous ne pouvez plus le remettre en stock',Ecran.Caption);
    exit;
    end;
  if TOBSL.GetValue('GLS_TENUESTOCK') = 'X' then CreerTOBStock (TOBSL);
  end;
AfficheGrilleStock ;
TOBSL.Free;
SupprimeLigne(ARow);
end;

procedure TOF_SAISIESERIE.BSelectAllClick(Sender: TObject);
begin
if Action=taConsult then Exit;
GStock.AllSelected := not GStock.AllSelected ;
if not bSelectAll.Down then TobSerStk.PutValueAllFille('SELECTED',false);
end;

procedure TOF_SAISIESERIE.BReliquatClick(Sender: TObject);
begin
if Action=taConsult then Exit;
SetControlVisible('PRELIQUAT', not GetControlVisible('PRELIQUAT'));
if GetControlVisible('PRELIQUAT') then
    begin
    GLIGNE.Align := alTop;
    GLIGNE.Height := 140;
    CreateSplitter;
    end else
    begin
    Splitter.Free;
    GLIGNE.Align := alClient;
    end;

end;

procedure TOF_SAISIESERIE.CreateSplitter;
begin
SPLITTER := TSplitter.Create(THPanel(GetControl('PLIGNE')));
SPLITTER.Name := 'Splitter';
SPLITTER.Parent := THPanel(GetControl('PLIGNE')) ;
SPLITTER.AutoSnap := True;
SPLITTER.Beveled := True;
SPLITTER.ResizeStyle := rsUpdate;
SPLITTER.Left := THGRID(GetControl('GLIGNE')).Left ;
SPLITTER.Align := THGRID(GetControl('GLIGNE')).Align;
SPLITTER.Width := THGRID(GetControl('GLIGNE')).Width;
SPLITTER.Height := 3;
SPLITTER.Cursor := crVSplit;
SPLITTER.Color := clActiveCaption;
end;

procedure TOF_SAISIESERIE.BBasClick(Sender: TObject);
Var St : string ;
    ARow : integer;
    QteAff : Double ;
    TOBSL : TOB;
    Cancel : boolean;
begin
if Action=taConsult then Exit;
ARow := GLigne.Row ;
TraiterLigneSerie (SerLigCol, ARow, Cancel);
if Cancel then exit;
QteAff:= TobSerLigRel.Detail.Count;
if QteAff >= QteReliquat then exit;
St:=GLigne.Cells [SerLigCol, ARow];
if St = '' then exit;
TOBSL:=GetTOBSLigne(ARow);
if TOBSL = nil then exit;
TOBSL.ChangeParent(TOBSerLigRel,-1);

if QteAff > 0 then GReliq.RowCount := GReliq.RowCount+1;
GReliq.Cells [SerRelCol,GReliq.RowCount-1] := St;
SupprimeLigne(ARow);
end;

procedure TOF_SAISIESERIE.BHautClick(Sender: TObject);
var i_ind : integer;
    St : string ;
    ARow : integer;
    QteAff,QteAffRel : Double ;
begin
if Action=taConsult then Exit;
QteAff := StrToFloat(GetControlText('QTEAFFECT'));
ARow := GReliq.Row ;
QteAffRel:=TobSerLigRel.Detail.Count;//StrToFloat(GetControlText('QTEAFFECTREL'));
if QteAffRel = 0 then exit;
if  QteAff >= QteUnitaire then GLigne.RowCount := GLigne.RowCount + 1;
St:=GReliq.Cells [SerRelCol, ARow];
for i_ind:=1 to GLigne.RowCount do
   begin
   if LigVide (i_ind) then
      begin
      if (i_ind = TOBSerLig.Detail.Count) then
         begin
         TobSerLig.Detail[i_ind-1].Dupliquer(TOBSerLigRel.Detail[ARow-1],true,true);
         TOBSerLigRel.Detail[ARow-1].Free;
         end else
      if (i_ind = TOBSerLig.Detail.Count + 1) then
         TOBSerLigRel.Detail[ARow-1].ChangeParent(TOBSerLig,i_ind-1);
      GLigne.Cells [SerLigCol, i_ind] := St;
      QteAff := QteAff + 1;
      SetControlText('QTEAFFECT',FloatToStr(QteAff));
      break;
      end;
   end;
if GReliq.RowCount > 2 then
   begin
   GReliq.CacheEdit ; GReliq.SynEnabled := False;
   GReliq.DeleteRow (ARow);
   GReliq.MontreEdit; GReliq.SynEnabled := True;
   end else GReliq.Cells [SerRelCol, ARow] := '';
end;


Initialization
  registerclasses ( [ TOF_SAISIESERIE ] ) ;
end.
