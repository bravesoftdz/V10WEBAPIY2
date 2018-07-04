unit UtofAffectStock;

interface

uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_main,Mul,
{$ENDIF}
      forms,sysutils,
      ComCtrls,Hpanel, Math
      ,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,M3FP,HTB97,Dialogs, AGLInitGC, ExtCtrls, Hqry,LicUtil
      ,HDimension,
      Facture,FactUtil,Menus,AssistAffectStock
      ,UTofOptionEdit,StockUtil,UtilPGI, UtilDispGC,uEntCommun;

Type
     TOF_AffectStock = Class (TOF)
     private
        LesColSel, LesColSai : array[1..2] of string ;
        CodeSEL, CodeSAI : array of string;
        FFieldList : string;
        mode, NbColSel, NbColSai : integer;
        TOBAff, TOBSAI, TOBStk : TOB ;
        SPLITTER : TSplitter;
        G_SEL : THGRID ;
        G_SAI : THGRID;
        PSAICUMUL : THPanel;
        CUMQTESTOCK : THNumEdit;
        CUMQTEAFF : THNumEdit;
        BINVERSE : TToolbarButton97;
        BAFFECTE : TToolbarButton97;
        BAFFECTEALL : TToolbarButton97;
        BASSIST: TToolbarButton97;
        BIMPRIMER: TToolbarButton97;
        POPZ: TPopupMenu;
        mnart: TMenuItem;
        mnpiece: TMenuItem;
        // Manipulation du Grid
        procedure InitialiseGrille;
        procedure ChargeGrille;
        procedure AffectGrilleSEL(G : THGrid ; Entete,EntLib :boolean ;
                                  CodeChp: array of string; CalculRowCount : Boolean);
        procedure AffectGrilleSAI(G : THGrid; Entete,EntLib :boolean;
                                  CodeChp: array of string; CalculRowCount : Boolean);
        function  QTEAFFSomme : Extended;
        procedure VoirArticle;
        procedure VoirPiece;
      // Evenement du Grid
        procedure G_SELDblClick(Sender: TObject);
        procedure G_SELCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_SELCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_SELRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_SELRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_SAIElipsisClick(Sender: TObject);
        procedure G_SAIDblClick(Sender: TObject);
        procedure G_SAICellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_SAICellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_SAIRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_SAIRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_SAIColumnWidthsChanged(Sender: TObject);
        procedure DessineCellSEL (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
        procedure DessineCellSAI(ACol,ARow : Longint; Canvas : TCanvas;  AState: TGridDrawState);
        // Manipulation des Champs GRID
        procedure AffichePied (ACol, ARow : integer);
        Function  GetTOBLigne ( ARow : integer) : TOB ;
        function  VerifDispo (TOBL : TOB; var QteNec : Double) : Boolean;
        procedure TraiterQteAff (ACol, ARow : integer);
        procedure Affecteuneligne(ARow : integer);
        // Evenement lié aux Boutons
        procedure BINVERSEClick(Sender: TObject);
        procedure BAFFECTEClick(Sender: TObject);
        procedure BAFFECTEALLClick(Sender: TObject);
        procedure BASSISTClick(Sender: TObject);
        procedure BIMPRIMERClick(Sender: TObject);
        procedure mnartClick(Sender: TObject);
        procedure mnpieceClick(Sender: TObject);
        // impression
        procedure Imprime ;
     public
        procedure OnArgument (Arguments : String ) ; override;
        procedure OnUpdate ; override;
        procedure Onclose  ; override ;
        function PrepareImpression : integer ;

     END ;

Var PEtat : RPARETAT;

Const Sel_Num      : integer = 0 ;
      Sel_Ind      : integer = 0 ;
      Sel_Tie      : integer = 0 ;
      Sel_Lib      : integer = 0 ;
      Sel_Dat      : integer = 0 ;
      Sel_Art      : integer = 8 ;
      Sel_Ali      : integer = 8 ;
      Sel_Dep      : integer = 8 ;
      Sel_Dli      : integer = 8 ;
      Sel_Qst      : integer = 8 ;
      Sel_Qaf      : integer = 8 ;
      Sel_Tht      : integer = 8 ;

Const Sai_Num      : integer = 0 ;
      Sai_Ind      : integer = 0 ;
      Sai_Tie      : integer = 0 ;
      Sai_Lib      : integer = 0 ;
      Sai_Dat      : integer = 0 ;
      Sai_Art      : integer = 8 ;
      Sai_Ali      : integer = 8 ;
      Sai_Dep      : integer = 8 ;
      Sai_Dli      : integer = 8 ;
      Sai_Qst      : integer = 8 ;
      Sai_Qaf      : integer = 8 ;
      Sai_Tht      : integer = 8 ;

const
// libellés des messages
TexteMessage: array[1..3] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'Le stock disponible est insuffisant.'
          {3} ,''
              );


implementation

Procedure TOF_AffectStock.OnArgument (Arguments : String ) ;
var i_mnart,i_mnpiece : integer;
begin
inherited ;
mode:=1;
PEtat.Tip:='E'; PEtat.Nat:='GZF'; PEtat.Modele:='GZF'; PEtat.Titre:=TFmul(Ecran).Caption;
PEtat.Apercu:=True; PEtat.DeuxPages:=False; PEtat.First:=True;
PEtat.stSQL:='';
SPLITTER:=TSplitter.Create(THPanel(GetControl('PSUPPORT')));
//SPLITTER:=TSplitter.Create(Ecran);
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

if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    BEGIN
    TToolbarButton97(GetControl('BParamListe')).Visible:=True ;
    END else
    BEGIN
    TToolbarButton97(GetControl('BParamListe')).Visible:=False;
    END;
BINVERSE:=TToolbarButton97(GetControl('BINVERSE'));
BINVERSE.OnClick:=BINVERSECLick;
BAFFECTE:=TToolbarButton97(GetControl('BAFFECTE'));
BAFFECTE.OnClick:=BAFFECTECLick;
BAFFECTEALL:=TToolbarButton97(GetControl('BAFFECTEALL'));
BAFFECTEALL.OnClick:=BAFFECTEALLCLick;
BASSIST:=TToolbarButton97(GetControl('BASSIST'));
BASSIST.OnClick:=BASSISTCLick;
BIMPRIMER:=TToolbarButton97(GetControl('Bimprimer')) ;
BIMPRIMER.OnClick:=BIMPRIMERCLick;

POPZ:=TPopupMenu(GetControl('POPZ'));
i_mnart:=-1; i_mnpiece:=-1 ;
if POPZ.Items[0].Name='mnart' then i_mnart:=0
    else if POPZ.Items[1].Name='mnart' then i_mnart:=1;
if POPZ.Items[0].Name='mnpiece' then i_mnpiece:=0
    else if POPZ.Items[1].Name='mnpiece' then i_mnpiece:=1;
if i_mnart>=0 then
    begin
    mnart:=POPZ.Items[i_mnart];
    mnart.OnClick:=mnartClick;
    end;
if i_mnpiece>=0 then
    begin
    mnpiece:=POPZ.Items[i_mnpiece];
    mnpiece.OnClick:=mnpieceClick;
    end;

LesColSel[1]:='FIXED;GP_NATUREPIECEG;GP_NUMERO;GP_INDICEG;GP_TIERS;T_LIBELLE;GP_DATEPIECE;GP_TOTALHT' ;
LesColSel[2]:='FIXED;GL_ARTICLE;GL_CODEARTICLE;GA_LIBELLE;GL_DEPOT' ;
LesColSai[1]:='FIXED;GL_CODEARTICLE;GA_LIBELLE;GL_DEPOT;GL_DATELIVRAISON;GL_QTESTOCK;QTEAFF' ;
LesColSai[2]:='FIXED;GP_NUMERO;GP_INDICEG;GP_TIERS;T_LIBELLE;GP_DATEPIECE;GL_DATELIVRAISON;GL_QTESTOCK;QTEAFF' ;

PSAICUMUL:=THPANEL(GetControl('PSAICUMUL'));

G_SEL:=THGRID(GetControl('G_SEL'));
G_SEL.OnDblClick:=G_SELDblClick ;
G_SEL.OnCellEnter:=G_SELCellEnter ;
G_SEL.OnCellExit:=G_SELCellExit ;
G_SEL.OnRowEnter:=G_SELRowEnter ;
G_SEL.OnRowExit:=G_SELRowExit ;
G_SEL.PostDrawCell:= DessineCellSEL;
G_SEL.ColCount:=7;
G_SEL.ColWidths[0]:=5;

G_SAI:=THGRID(GetControl('G_SAI'));
G_SAI.OnElipsisClick:=G_SAIElipsisClick  ;
G_SAI.OnDblClick:=G_SAIDblClick ;
G_SAI.OnCellEnter:=G_SAICellEnter ;
G_SAI.OnCellExit:=G_SAICellExit ;
G_SAI.OnRowEnter:=G_SAIRowEnter ;
G_SAI.OnRowExit:=G_SAIRowExit ;
G_SAI.OnColumnWidthsChanged:=G_SAIColumnWidthsChanged;
G_SAI.PostDrawCell:= DessineCellSAI;
G_SAI.ColCount:=6;
G_SAI.ColWidths[0]:=5;
SetControlVisible('Fliste',False);
THLabel(GetControl('TDIMENSION')).Visible:=False

end;

procedure TOF_AffectStock.OnUpdate;
var F : TFMul ;
    i_ind : integer ;
    Select,Where, Where2,St,Depot : string ;
    QQ : TQuery ;
    TOBL : TOB;
    Qte,RatioVA : Double ;
BEGIN
inherited;
InitialiseGrille;
F:=TFMul(Ecran);
Where := RecupWhereCritere(F.Pages) ;
//i_pos:=pos('@$@',Where);
//if i_pos >=0 then Where:=Copy(Where,0,i_pos-1) + 'CC' + Copy(Where,i_pos + 3,length(Where)-i_pos);
if TobStk<>Nil then begin TobStk.Free; TobStk:=Nil; end;
if TOBAff<>Nil then begin TOBAff.Free; TOBAff:=Nil; end;
if TOBSAI<>Nil then begin TOBSAI.Free; TOBSAI:=Nil; end;
TobStk := TOB.Create ('', Nil, -1) ;
TobAff := TOB.Create ('', Nil, -1) ;
//QQ := OpenSQL('SELECT * FROM GCAFFECTSTOCK ' + Where + ' ORDER BY GP_SOUCHE,GP_NUMERO', True) ;
Select:='SELECT gp_naturepieceg, gp_souche, gp_numero, gp_indiceg, gp_datepiece,gp_etablissement,'+
        'gp_totalht,gp_tiers,  t_libelle, gl_naturepieceg,gl_article,gl_codearticle,gl_libelle,'+
        'gl_datelivraison, gl_qtestock, gl_depot,gl_tenuestock,gl_qualifqtevte,gl_qualifqtesto,'+
        'gq_physique, gq_prepacli,ga_libelle, ga_statutart,ga_grilledim1,ga_grilledim2,ga_grilledim3,'+
        'ga_grilledim4, ga_grilledim5,ga_codedim1,ga_codedim2,ga_codedim3,ga_codedim4,ga_codedim5  '+
        'FROM piece, ligne LEFT Join DISPO on gq_article=gl_article and gl_depot=gq_depot and gq_cloture<>"X" ,tiers,article  ';
Where2:='gp_naturepieceg=gl_naturepieceg and gp_souche=gl_souche and gp_numero=gl_numero and '+
        'gp_indiceg=gl_indiceg and gp_etatvisa<>"ATT" and gp_vivante="X" and ga_article=gl_article and '+
        'gl_typearticle="MAR" and t_tiers=gp_tiers and gp_naturepieceg="CC" ';
if Where='' then Where:='WHERE '+Where2 else Where:=Where+' and '+where2;
QQ := OpenSQL(Select + Where + ' ORDER BY GP_SOUCHE,GP_NUMERO', True) ;
if not QQ.EOF then TOBAff.LoadDetailDB('GCAFFECTSTOCK', '', '', QQ, FAlse) else TOBAff := nil;
Ferme (QQ) ;
if (TOBAff = nil) or (TOBAff.Detail.count = 0) then
    BEGIN
    G_SEL.VidePile(False); G_SAI.VidePile(False);
    G_SEl.Enabled:= False; G_SAI.Enabled:= False; BINVERSE.Enabled:=False;
    for i_ind:=0 to POPZ.Items.Count-1 do POPZ.Items[i_ind].Enabled:=False;
    exit;
    END else
    BEGIN
    G_SEl.Enabled:= True; G_SAI.Enabled:= True; BINVERSE.Enabled:=True;
    for i_ind:=0 to POPZ.Items.Count-1 do POPZ.Items[i_ind].Enabled:=True;
    TOBAff.Detail[0].AddChampSup('LIEN', True);
    for i_ind:=0 to TOBAff.Detail.count-1 do
        BEGIN
        TOBAff.Detail[i_ind].PutValue('LIEN', i_ind);
        St:=TOBAff.Detail[i_ind].GetValue('GL_ARTICLE');
        Depot:=TOBAff.Detail[i_ind].GetValue('GL_DEPOT');
        if TobStk.FindFirst(['ARTICLE','DEPOT'],[St,Depot],False) <> Nil then continue;
        TOBL:=TOB.Create ('', TobStk,-1) ;
        TOBL.AddChampSup('ARTICLE', False);
        TOBL.AddChampSup('DEPOT', False);
        TOBL.AddChampSup('QTEDISPO', False);
        TOBL.PutValue('ARTICLE', St);
        TOBL.PutValue('DEPOT', Depot);
        Qte:=0;
        if (TOBAff.Detail[i_ind].GetValue('GQ_PHYSIQUE')= null) or
           (TOBAff.Detail[i_ind].GetValue('GQ_PREPACLI')= null) then
            BEGIN
            TOBL.PutValue('QTEDISPO', Qte);
            TOBAff.Detail[i_ind].PutValue('GQ_PHYSIQUE', 0);
            TOBAff.Detail[i_ind].PutValue('GQ_PREPACLI', 0);
            END else
            Qte:=TOBAff.Detail[i_ind].GetValue('GQ_PHYSIQUE')-
                 TOBAff.Detail[i_ind].GetValue('GQ_PREPACLI');
            RatioVA:=GetRatio(TOBAff.Detail[i_ind],Nil,trsStock) ; if RatioVA=0 then RatioVA:=1 ;
            Qte:=Qte*RatioVA;
            TOBL.PutValue('QTEDISPO', Qte);
        END;
    St:=FormatFloat(G_SAI.ColFormats[Sai_Qaf], 0);
    TobAff.Detail[0].AddChampSupValeur('QTEAFF', St, True);
    ChargeGrille;
    END;
END;

Procedure TOF_AffectStock.Onclose  ;
begin
inherited ;
SPLITTER.Free;
CUMQTESTOCK.Free; CUMQTEAFF.Free;
if TOBSAI<>Nil then TOBSAI.Free;
if TobStk<>Nil then TobStk.Free;
if TOBAff<>Nil then TOBAff.Free;
end ;

{==============================================================================================}
{================================= Manipulation du Grid =========================================}
{==============================================================================================}
procedure TOF_AffectStock.InitialiseGrille;
var i,j,Dec, FixedWidthSel, FixedWidthSai : integer ;
    St,Nam,stAl,stA,CH, FF, FPerso : string ;
    NomList,FRecordSource,FLien,FSortBy,FLargeur,FAlignement,FParams : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;
    FTitre,tt,NC : Hstring;
BEGIN
Sel_Num := 0 ; Sel_Ind := 0 ; Sel_Tie := 0 ; Sel_Lib := 0 ; Sel_Dat := 0 ;
Sel_Art := 8 ; Sel_Ali := 8 ; Sel_Dep := 8 ; Sel_Dli := 8 ; Sel_Qst := 8 ;
Sel_Qaf := 8 ;

Sai_Num := 0 ; Sai_Ind := 0 ; Sai_Tie := 0 ; Sai_Lib := 0 ; Sai_Dat := 0 ;
Sai_Art := 8 ; Sai_Ali := 8 ; Sai_Dep := 8 ; Sai_Dli := 8 ; Sai_Qst := 8 ;
Sai_Qaf := 8 ;
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
    if Nam='GP_NUMERO' then Sel_Num:=i
    else if Nam='GP_INDICEG' then Sel_Ind:=i
    else if Nam='GP_TIERS' then Sel_Tie:=i
    else if Nam='T_LIBELLE' then Sel_Lib:=i
    else if Nam='GP_DATEPIECE' then Sel_Dat:=i
    else if Nam='GL_CODEARTICLE' then Sel_Art:=i
    else if Nam='GA_LIBELLE' then Sel_Ali:=i
    else if Nam='GL_DEPOT' then Sel_Dep:=i
    else if Nam='GL_DATELIVRAISON' then Sel_Dli:=i
    else if Nam='GL_QTESTOCK' then Sel_Qst:=i
    else if Nam='GL_QTEAFF' then Sel_Qaf:=i
    else if Nam='GP_TOTALHT' then Sel_Tht:=i
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
    if Nam='GP_NUMERO' then Begin Sai_Num:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GP_INDICEG' then Begin Sai_Ind:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GP_TIERS' then Begin Sai_Tie:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='T_LIBELLE' then Begin Sai_Lib:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GP_DATEPIECE' then Begin Sai_Dat:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GL_CODEARTICLE' then Begin Sai_Art:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GA_LIBELLE' then Begin Sai_Ali:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GL_DEPOT' then Begin Sai_Dep:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GL_DATELIVRAISON' then Begin Sai_Dli:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='GL_QTESTOCK' then Begin Sai_Qst:=i; G_SAI.ColLengths[i]:=-1; end
    else if Nam='QTEAFF' then Sai_Qaf:=i
    else if Nam='GP_TOTALHT' then Begin Sai_Tht:=i; G_SAI.ColLengths[i]:=-1; end
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
            G_SAI.ColAligns[i]:=TFMul(Ecran).Fliste.ColAligns[j] ;
            G_SAI.ColWidths[i]:=TFMul(Ecran).Fliste.ColWidths[j] ;
            if OkLib then G_SAi.ColFormats[i]:='CB=' + Get_Join(Nam)
                     else if (Dec<>0) or (Sep) then G_SAI.ColFormats[i]:=FF ;
            if i=Sai_Qst then THNumEdit(GetControl('QTEDISPO')).Decimals:=Dec ;
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
            if i=Sai_Qst then THNumEdit(GetControl('QTEDISPO')).Decimals:=Dec ;
            break;
            END;
        END;
{$ENDIF}
    END ;
G_SAI.ColAligns[Sai_Qaf]:=G_SAI.ColAligns[Sai_Qst];
G_SAI.ColWidths[Sai_Qaf]:=G_SAI.ColWidths[Sai_Qst];
G_SAi.ColFormats[Sai_Qaf]:=G_SAI.ColFormats[Sai_Qst];
THNumEdit(GetControl('QTEDISPO')).Masks.PositiveMask:=G_SAI.ColFormats[Sai_Qst];

AffecteGrid(G_SEL,taConsult) ;
TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SEL) ;
G_SEL.ColWidths[1]:=0;
G_SEL.ColWidths[0]:=FixedWidthSel;TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SEL) ;
AffecteGrid(G_SAI,taModif) ;
TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SAI) ;
G_SAI.ColWidths[0]:=FixedWidthSai;TFMul(Ecran).Hmtrad.ResizeGridColumns(G_SAI) ;
//TFMul(Ecran).OnKeyDown:=FormKeyDown ;

END;

procedure TOF_AffectStock.ChargeGrille;
Var st, code, libel, CH : string ;
    i_ind, ACol, icol : integer;
    TOBTemp, TOBTempD, TOBD1 : TOB;
    Coord : TRect;
BEGIN
While PSAICUMUL.ControlCount>0 do PSAICUMUL.Controls[0].Free ;
CUMQTESTOCK:=THNumEdit.Create(TFMul(Ecran));
CUMQTESTOCK.Parent:=PSAICUMUL;
CUMQTESTOCK.ParentColor:=True;
CUMQTESTOCK.Font.Style:=PSAICUMUL.Font.Style;
CUMQTESTOCK.Font.Size:=PSAICUMUL.Font.Size;
CUMQTESTOCK.Masks.PositiveMask:=G_SAI.ColFormats[Sai_Qst];
CUMQTESTOCK.Ctl3D:=False; CUMQTESTOCK.Top:=1;
Coord:=G_SAI.CellRect(Sai_Qst,0);
CUMQTESTOCK.Left:=Coord.Left + 1;
CUMQTESTOCK.Width:=G_SAI.ColWidths[Sai_Qst] + 1;
CUMQTESTOCK.Height:=PSAICUMUL.Height;

CUMQTEAFF:=THNumEdit.Create(TFMul(Ecran));
CUMQTEAFF.Parent:=PSAICUMUL;
CUMQTEAFF.ParentColor:=True;
CUMQTEAFF.Font.Style:=PSAICUMUL.Font.Style;
CUMQTEAFF.Font.Size:=PSAICUMUL.Font.Size;
CUMQTEAFF.Masks.PositiveMask:=G_SAI.ColFormats[Sai_Qaf];
CUMQTEAFF.Ctl3D:=False; CUMQTEAFF.Top:=1;
Coord:=G_SAI.CellRect(Sai_Qaf,0);
CUMQTEAFF.Left:=Coord.Left + 1;
CUMQTEAFF.Width:=G_SAI.ColWidths[Sai_Qaf] + 1;
CUMQTEAFF.Height:=PSAICUMUL.Height;

CH:=FFieldList ;
TOBTemp:=TOB.Create('', Nil,-1);
{$IFDEF EAGLCLIENT}
for i_ind:=0 to TFMul(Ecran).Fliste.ColCount - 1 do
    BEGIN
    TOBTempD:=TOB.Create('', TOBTemp,-1);
    TOBTempD.AddChampSup('CODE',False);
    TOBTempD.AddChampSup('LIBELLE',False);
    TOBTempD.PutValue('CODE',ReadTokenSt(CH)) ;
    TOBTempD.PutValue('LIBELLE',TFMul(Ecran).Fliste.Cells[i_ind,0]) ;
    END;
{$ELSE}
for i_ind:=0 to TFMul(Ecran).Fliste.Columns.Count - 1 do
    BEGIN
    TOBTempD:=TOB.Create('', TOBTemp,-1);
    TOBTempD.AddChampSup('CODE',False);
    TOBTempD.AddChampSup('LIBELLE',False);
    TOBTempD.PutValue('CODE',ReadTokenSt(CH)) ;
    TOBTempD.PutValue('LIBELLE',TFMul(Ecran).Fliste.Columns.Items[i_ind].Title.Caption) ;
    END;
{$ENDIF}
st:= LesColSEL[mode];
ACol:=0; icol:=0;
SetLength(CodeSEL,G_SEL.ColCount);
While st <> '' do
    BEGIN
    code:=ReadTokenSt(st);
    TOBD1:=TOBTemp.FindFirst(['CODE'],[code],True);
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
    TOBD1:=TOBTemp.FindFirst(['CODE'],[code],True);
    if TOBD1 <> nil then Libel:=TOBD1.GetValue('LIBELLE') else Libel:='';
    G_SAI.Cells[ACol, 0]:= Libel;
    inc(ACol);
    if TOBD1 <> nil then begin CodeSAI[icol]:=Code; inc(icol); end;
    END;
G_SAI.Cells[Sai_Qaf, 0]:= 'Qté affectée';
CodeSAI:= Copy(CodeSAI, 0, icol);
TOBTemp.Free;
AffectGrilleSEL(G_SEL,False,False,CodeSEL,True);
AffectGrilleSAI(G_SAI,False,False,CodeSEL,True);
G_SAI.Col:=Sai_Qaf;
END;

procedure TOF_AffectStock.AffectGrilleSEL(G : THGrid ; Entete,EntLib :boolean ;
                                          CodeChp: array of string; CalculRowCount : Boolean);
var TOBT, TOBTD,  TOBRefD : TOB;
    Valeur : array of variant;
    i_ind, icode : integer;
BEGIN
TOBT:=TOB.Create ('', Nil, -1);
SetLength(Valeur,Length(CodeChp));
for i_ind:=0 to TOBAff.Detail.count - 1 do
    BEGIN
    TOBRefD:= TOBAff.Detail[i_ind];
    for icode:=Low(CodeChp) to High(CodeChp) do
        BEGIN
        if CodeChp[icode]<>'' then Valeur[icode]:=TOBRefD.GetValue(CodeChp[icode]);
        END;
    if TOBT.FindFirst(CodeChp,Valeur,True) <> Nil then continue;
    TOBTD:=TOB.Create ('', TOBT,-1) ;
    TOBTD.Dupliquer(TOBRefD, False, True);
    END;
TOBT.PutGridDetail(G,Entete,EntLib,LesColSEL[mode],CalculRowCount);
TOBT.Free;
END;

procedure TOF_AffectStock.AffectGrilleSAI(G : THGrid; Entete,EntLib :boolean;
                                          CodeChp: array of string; CalculRowCount : Boolean);
var TOBT, TOBTD,  TOBRefD : TOB;
    Valeur : array of variant;
    icode, i_ind : integer;
    StName : string;
BEGIN
SetLength(Valeur,Length(CodeChp));
StName:=LesColSEL[mode];

TOBT:=TOB.Create ('', Nil, -1);
while StName <> '' do TOBT.AddChampSup(ReadTokenSt(StName), False);
TOBT.GetLigneGrid(G_SEL,G_SEL.Row,LesColSEL[mode]);
for icode:=Low(CodeChp) to High(CodeChp) do Valeur[icode]:=TOBT.GetValue(CodeChp[icode]);
TOBT.Free;

if TOBSAI<>nil then
    BEGIN
    for i_ind:=0 to TOBSAI.Detail.Count - 1 do
        BEGIN
        TOBTD:=TOBAff.FindFirst(['LIEN'],[TOBSAI.Detail[i_ind].GetValue('LIEN')],False);
        if TOBTD<>Nil then TOBTD.PutValue('QTEAFF', TOBSAI.Detail[i_ind].GetValue('QTEAFF'));
        END;
    TOBSAI.free; TOBSAI:=Nil ;
    END;
TOBSAI:=TOB.Create ('', Nil, -1);
TOBRefD:= TOBAff.FindFirst(CodeChp,Valeur,True);
While TOBRefD <> nil do
    BEGIN
    TOBTD:=TOB.Create ('', TOBSAI,-1) ;
    TOBTD.Dupliquer(TOBRefD, False, True);
    TOBRefD:= TOBAff.FindNext(CodeChp,Valeur,True);
    END;
TOBSAI.PutGridDetail(G,Entete,EntLib,LesColSAI[mode],CalculRowCount);
CUMQTESTOCK.Value:=TOBSAI.Somme('GL_QTESTOCK',[''],[''],False);
CUMQTEAFF.Value:=QTEAFFSomme;
AffichePied (G_SAI.Col, G_SAI.Row);
END;

function TOF_AffectStock.QTEAFFSomme : Extended;
var i_row : Integer;
BEGIN
Result:=0;
for i_row:=1 to G_SAI.RowCount - 1 do Result:=Result + Valeur(G_SAI.Cells[Sai_Qaf, i_row]);
END;

procedure TOF_AffectStock.VoirArticle;
var TOBL : TOB;
    RefArt : string;
BEGIN
TOBL := GetTOBLigne (G_SAI.Row); if TOBL=nil then exit;
RefArt:=TOBL.GetValue('GL_ARTICLE');
{$IFNDEF GPAO}
  DispatchTTArticle(TaConsult,RefArt,'ACTION=CONSULTATION','');
{$ELSE}
  V_PGI.DispatchTT(7, taConsult, RefArt, '', '');
{$ENDIF GPAO}
END;

procedure TOF_AffectStock.VoirPiece;
var TOBL : TOB;
    CleDoc : R_CleDoc;
BEGIN
TOBL := GetTOBLigne (G_SAI.Row); if TOBL=nil then exit;
CleDoc.NaturePiece:=Uppercase(TOBL.GetValue('GP_NATUREPIECEG'));
CleDoc.Souche:=Uppercase(TOBL.GetValue('GP_SOUCHE'));
CleDoc.DatePiece:=TOBL.GetValue('GP_DATEPIECE');
CleDoc.NumeroPiece:=TOBL.GetValue('GP_NUMERO');
CleDoc.Indice:=TOBL.GetValue('GP_INDICEG');
SaisiePiece(CleDoc,taConsult)
END;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}
procedure TOF_AffectStock.G_SELDblClick(Sender: TObject);
BEGIN
if Mode=1 then VoirPiece else VoirArticle;
END;

procedure TOF_AffectStock.G_SELCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
BEGIN
END;

procedure TOF_AffectStock.G_SELCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
BEGIN
END;

procedure TOF_AffectStock.G_SELRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ACol : integer;
    Desc : Boolean;
BEGIN
G_SEL.InvalidateRow(ou) ;
AffectGrilleSAI(G_SAI,False,False,CodeSEL,True);
ACol:=G_SAI.SortedCol ;
Desc:=False;
if ACol <> -1 then G_Sai.SortGrid(ACol,Desc);
END;

procedure TOF_AffectStock.G_SELRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
G_SEL.InvalidateRow(ou) ;
END;

procedure TOF_AffectStock.G_SAIElipsisClick(Sender: TObject);
BEGIN
END;

procedure TOF_AffectStock.G_SAIDblClick(Sender: TObject);
BEGIN
if Mode=2 then VoirPiece else VoirArticle;
END;

procedure TOF_AffectStock.G_SAICellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
BEGIN
AffichePied (G_SAI.Col, G_SAI.Row);
END;

procedure TOF_AffectStock.G_SAICellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
BEGIN
if ACol = SAI_Qaf then TraiterQteAff (ACol, ARow);
END;

procedure TOF_AffectStock.G_SAIRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
G_SAI.InvalidateRow(ou) ;
END;

procedure TOF_AffectStock.G_SAIRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
G_SAI.InvalidateRow(ou) ;
END;

procedure TOF_AffectStock.G_SAIColumnWidthsChanged(Sender: TObject);
var Coord : TRect;
BEGIN
if PSAICUMUL.ControlCount<=0 then exit ;
Coord:=G_SAI.CellRect(Sai_Qst,0);
CUMQTESTOCK.Left:=Coord.Left + 1;
CUMQTESTOCK.Width:=G_SAI.ColWidths[Sai_Qst] + 1;

Coord:=G_SAI.CellRect(Sai_Qaf,0);
CUMQTEAFF.Left:=Coord.Left + 1;
CUMQTEAFF.Width:=G_SAI.ColWidths[Sai_Qaf] + 1;
END;

procedure TOF_AffectStock.DessineCellSEL (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
BEGIN
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
END;

procedure TOF_AffectStock.DessineCellSAI(ACol,ARow : Longint; Canvas : TCanvas;  AState: TGridDrawState);
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

{==============================================================================================}
{============================ Manipulation des Champs GRID ====================================}
{==============================================================================================}
procedure TOF_AffectStock.AffichePied (ACol, ARow : integer);
var TOBD,TOBL : TOB;
    RefArt,Depot,St : string;
    Dim : Array of string;
    i_ind: integer;
    Qte : double;
    GrilleDim,CodeDim,LibDim : String ;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
RefArt:=TOBL.GetValue('GL_ARTICLE');
Depot:=TOBL.GetValue('GL_DEPOT');
TOBD:=TOBStk.FindFirst(['ARTICLE','DEPOT'],[RefArt,Depot],False);
if TOBD=Nil then Qte:=0
            else Qte:=TOBD.GetValue('QTEDISPO') ;
if TOBL.GetValue('GL_TENUESTOCK')<>'X' then
    BEGIN
    THNumEdit(GetControl('QTEDISPO')).Visible:=False ;
    THLabel(GetControl('TDISPO')).Visible:=True ;
    END else
    BEGIN
    THLabel(GetControl('TDISPO')).Visible:=False ;
    THNumEdit(GetControl('QTEDISPO')).Visible:=True ;
    THNumEdit(GetControl('QTEDISPO')).Value:=Qte;
    END;
if TOBL.GetValue('GA_STATUTART')<>'DIM' then
    THLabel(GetControl('TDIMENSION')).Visible:=False
    else
    BEGIN
    SetLength(dim,MaxDimension);
    St:='';
    for i_ind:=1 to MaxDimension do
        BEGIN
        GrilleDim:=TOBL.GetValue ('GA_GRILLEDIM'+IntToStr(i_ind)) ;
        CodeDim:=TOBL.GetValue ('GA_CODEDIM'+IntToStr(i_ind)) ;
        LibDim:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
        if LibDim<>'' then Dim[i_ind - 1]:=LibDim else Dim[i_ind - 1]:='' ;
        if Dim[i_ind - 1]<>'' then
            if St='' then St:=Dim[i_ind -1]
                     else St:=St + ' - ' + Dim[i_ind - 1];
        END;
    THLabel(GetControl('TDIMENSION')).Visible:=True;
    THLabel(GetControl('TDIMENSION')).Caption:=St;
    END;
END;

Function TOF_AffectStock.GetTOBLigne (ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBSAI.Detail.Count)) then Exit ;
Result:=TOB(G_SAI.OBjects[0,ARow]);
END ;

function TOF_AffectStock.VerifDispo (TOBL : TOB; var QteNec: Double) : Boolean;
var QteDispo : Double;
    RefArt,Depot : string;
    TOBD : TOB;
BEGIN
Result:=False;
if TOBL.GetValue('GL_TENUESTOCK')<>'X' then begin Result:=True; exit; end;
RefArt:=TOBL.GetValue('GL_ARTICLE');
Depot:=TOBL.GetValue('GL_DEPOT');
TOBD:=TOBStk.FindFirst(['ARTICLE','DEPOT'],[RefArt,Depot],False);
if TOBD=Nil then exit;
QteDispo:=TOBD.GetValue('QTEDISPO');
if QteDispo < QteNec then
    BEGIN
    if QteDispo > 0 then QteNec:=QteDispo else QteNec:=0;
    exit;
    END;
TOBD.PutValue('QTEDISPO', QteDispo-QteNec);
Result:=True;
END;

procedure TOF_AffectStock.TraiterQteAff (ACol, ARow : integer);
var TOBL : TOB;
    St : String;
    QteAff, QteStk, QteAnc, QteNec : double;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
QteStk:=TOBL.GetValue ('GL_QTESTOCK');
QteAnc:=TOBL.GetValue ('QTEAFF');
QteAff:=Valeur(G_SAI.Cells [ACol, ARow]) ;
QteNec:= QteAff - QteAnc;
if (QteAff < 0) or (QteAff>QteStk) then
    BEGIN
    St:=FormatFloat(G_SAI.ColFormats[Sai_Qaf], QteAnc);
    G_SAI.Cells [ACol, ARow]:=St;
    END else
    BEGIN
    if Not VerifDispo(TOBL, QteNec) then
        BEGIN
        if QteAff <> 0 then
            HShowMessage('0;'+TFMul(Ecran).Caption+';'+TexteMessage[2]+';W;O;O;O;','','') ;
        St:=FormatFloat(G_SAI.ColFormats[Sai_Qaf], QteAnc);
        G_SAI.Cells [ACol, ARow]:=St;
        END else
        BEGIN
        St:=FormatFloat(G_SAI.ColFormats[Sai_Qaf], QteAff);
        G_SAI.Cells [ACol, ARow]:=St;
        TOBL.PutValue ('QTEAFF', St);
        END;
    END;
CUMQTEAFF.Value:=QTEAFFSomme;
END;

procedure TOF_AffectStock.Affecteuneligne(ARow : integer);
var TOBL : TOB;
    St : String;
    QteAff, QteStk, QteAnc, QteNec : double;
begin
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
QteStk:=TOBL.GetValue ('GL_QTESTOCK');
QteAnc:=TOBL.GetValue ('QTEAFF');
QteAff:=QteStk ;
QteNec:= QteAff - QteAnc;
if (QteAff < 0) or (QteStk=0) then exit;
if Not VerifDispo(TOBL, QteNec) then
    BEGIN
    VerifDispo(TOBL, QteNec);
    END;
St:=FormatFloat(G_SAI.ColFormats[Sai_Qaf], QteNec+QteAnc);
G_SAI.Cells [Sai_Qaf, ARow]:=St;
TOBL.PutValue ('QTEAFF', St);
CUMQTEAFF.Value:=QTEAFFSomme;
end;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TOF_AffectStock.BINVERSEClick(Sender: TObject);
var i_ind : integer;
    FF : string ;
begin
if mode=1 then mode:=2 else mode:=1 ;
CUMQTESTOCK.Free;
CUMQTEAFF.Free;
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
G_SEL.ValCombo.Free; G_SEL.ValCombo:=Nil;
G_SAI.ValCombo.Free; G_SAI.ValCombo:=Nil;
G_SEL.VidePile(False);
G_SAI.VidePile(False);
InitialiseGrille;
ChargeGrille;
end;

procedure TOF_AffectStock.BAFFECTEClick(Sender: TObject);
begin
Affecteuneligne(G_SAI.Row);
AffichePied (G_SAI.Col,G_SAI.Row);
end;

procedure TOF_AffectStock.BAFFECTEALLClick(Sender: TObject);
var i_ind : integer ;
begin
for i_ind:=1 to G_SAI.RowCount-1 do Affecteuneligne(i_ind);
AffichePied (G_SAI.Col,G_SAI.Row);
end;

procedure TOF_AffectStock.BASSISTClick(Sender: TObject);
var Cancel : Boolean ;
begin
G_SELRowEnter(Nil, G_SEL.Row, Cancel, False);
Assist_AffectStock (TOBAff, TOBStk, G_SAI.ColFormats[Sai_Qaf]);
TOBSai.free; TOBSai:=Nil;
G_SELRowEnter(Nil, G_SEL.Row, Cancel, False);
end;

procedure TOF_AffectStock.BIMPRIMERClick(Sender: TObject);
var Cancel : Boolean ;
begin
G_SELRowEnter(Nil, G_SEL.Row, Cancel, False);
Imprime ;
TOBSai.free; TOBSai:=Nil;
G_SELRowEnter(Nil, G_SEL.Row, Cancel, False);
end;

procedure TOF_AffectStock.mnartClick(Sender: TObject);
BEGIN
VoirArticle;
END;

procedure TOF_AffectStock.mnpieceClick(Sender: TObject);
BEGIN
VoirPiece;
END;

{==============================================================================================}
{=================================== Impression ===============================================}
{==============================================================================================}
function TOF_AffectStock.PrepareImpression : integer ;
var i_ind : integer;
    TOBGZF, TOBD, TOBL : TOB;
BEGIN
Result:=0;
ExecuteSQL('DELETE FROM GCTMPAFFECTSTK WHERE GZF_UTILISATEUR = "'+V_PGI.USer+'"');
TOBGZF:=TOB.Create('',Nil,-1);
for i_ind:=0 to TOBAff.Detail.Count-1 do
    BEGIN
    TOBL:=TOBAff.Detail[i_ind] ;
    TOBD:=TOB.Create('GCTMPAFFECTSTK',TOBGZF,-1);
    TOBD.PutValue('GZF_UTILISATEUR',V_PGI.USer);
    TOBD.PutValue('GZF_COMPTEUR',i_ind);
    TOBD.PutValue('GZF_NATUREPIECEG',TOBL.GetValue('GP_NATUREPIECEG'));
    TOBD.PutValue('GZF_SOUCHE',TOBL.GetValue('GP_SOUCHE'));
    TOBD.PutValue('GZF_NUMERO',TOBL.GetValue('GP_NUMERO'));
    TOBD.PutValue('GZF_INDICEG',TOBL.GetValue('GP_INDICEG'));
    TOBD.PutValue('GZF_DATEPIECE',TOBL.GetValue('GP_DATEPIECE'));
    TOBD.PutValue('GZF_ETABLISSEMENT',TOBL.GetValue('GP_ETABLISSEMENT'));
    TOBD.PutValue('GZF_TIERS',TOBL.GetValue('GP_TIERS'));
    TOBD.PutValue('GZF_LIBELLETIERS',TOBL.GetValue('T_LIBELLE'));
    TOBD.PutValue('GZF_TOTALHT',TOBL.GetValue('GP_TOTALHT'));
    TOBD.PutValue('GZF_ARTICLE',TOBL.GetValue('GL_ARTICLE'));
    TOBD.PutValue('GZF_CODEARTICLE',TOBL.GetValue('GL_CODEARTICLE'));
    TOBD.PutValue('GZF_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
    TOBD.PutValue('GZF_DATELIVRAISON',TOBL.GetValue('GL_DATELIVRAISON'));
    TOBD.PutValue('GZF_DEPOT',TOBL.GetValue('GL_DEPOT'));
    TOBD.PutValue('GZF_TENUESTOCK',TOBL.GetValue('GL_TENUESTOCK'));
    TOBD.PutValue('GZF_STATUTART',TOBL.GetValue('GA_STATUTART'));
    TOBD.PutValue('GZF_GRILLEDIM1',TOBL.GetValue('GA_GRILLEDIM1'));
    TOBD.PutValue('GZF_GRILLEDIM2',TOBL.GetValue('GA_GRILLEDIM2'));
    TOBD.PutValue('GZF_GRILLEDIM3',TOBL.GetValue('GA_GRILLEDIM3'));
    TOBD.PutValue('GZF_GRILLEDIM4',TOBL.GetValue('GA_GRILLEDIM4'));
    TOBD.PutValue('GZF_GRILLEDIM5',TOBL.GetValue('GA_GRILLEDIM5'));
    TOBD.PutValue('GZF_CODEDIM1',TOBL.GetValue('GA_CODEDIM1'));
    TOBD.PutValue('GZF_CODEDIM2',TOBL.GetValue('GA_CODEDIM2'));
    TOBD.PutValue('GZF_CODEDIM3',TOBL.GetValue('GA_CODEDIM3'));
    TOBD.PutValue('GZF_CODEDIM4',TOBL.GetValue('GA_CODEDIM4'));
    TOBD.PutValue('GZF_CODEDIM5',TOBL.GetValue('GA_CODEDIM5'));
    TOBD.PutValue('GZF_QTESTOCK',TOBL.GetValue('GL_QTESTOCK'));
    TOBD.PutValue('GZF_PHYSIQUE',TOBL.GetValue('GQ_PHYSIQUE'));
    TOBD.PutValue('GZF_PREPACLI',TOBL.GetValue('GQ_PREPACLI'));
    TOBD.PutValue('GZF_QTEAFF',TOBL.GetValue('QTEAFF'));
    END;
TOBGZF.InsertOrUpdateDB(True);
TOBGZF.Free;
END;

procedure TOF_AffectStock.Imprime ;
BEGIN
EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('Pages')),PEtat.stSQL,PEtat.Titre, PrepareImpression);
ExecuteSQL('DELETE FROM GCTMPAFFECTSTK WHERE GZF_UTILISATEUR = "'+V_PGI.USer+'"');
END;


Initialization
registerclasses([TOF_AffectStock]);

end.
