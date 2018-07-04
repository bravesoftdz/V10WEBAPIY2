unit UTofPrixRevient;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Hpanel, Math
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,ParamSoc,
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,
{$ENDIF}
      M3FP,HTB97,Dialogs, AGLInitGC;

Type
     TOF_PRIXREVIENT = Class (TOF)
     private
        LesColArt, LesColFou, LesColDep : string ;
        G_ART : THGRID ;
        G_FOU : THGRID ;
        G_DEP : THGRID ;
        AArt,AFou,ADEP,ATYP, AVAL : integer ;              
        FArt,FFou,FDEP,FTYP, FVAL : integer ;
        DArt,DFou,DDEP,DTYP, DVAL : integer ;
        PART  : THPanel ;
        PFOU  : THPanel ;
        PDEP  : THPanel ;
        RBART : TRadioButton;
        RBFOU : TRadioButton;
        RBDEP : TRadioButton;
        BNewLine: TToolbarButton97;
        BDelLine: TToolbarButton97;
        BChercher: TToolbarButton97;
        FindLigne: TFindDialog;
        TobPR, TobToDelete : TOB ;
        procedure BNewLineClick(Sender: TObject);
        procedure BDelLineClick(Sender: TObject);
        procedure BChercherClick(Sender: TObject);
        procedure FindLigneFind(Sender: TObject);
        procedure RBARTClick(Sender: TObject);
        procedure RBFOUClick(Sender: TObject);
        procedure RBDEPClick(Sender: TObject);
        procedure G_ARTElipsisClick(Sender: TObject);
        procedure G_ARTCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_ARTCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_ARTRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_ARTRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_FOUElipsisClick(Sender: TObject);
        procedure G_FOUCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_FOUCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_FOURowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_FOURowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_DEPElipsisClick(Sender: TObject);
        procedure G_DEPCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_DEPCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_DEPRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_DEPRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure DessineCellArt (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
        procedure DessineCellFou (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
        procedure DessineCellDep (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
        // Actions liées au grid
        procedure ChargeGrille;
        Function  GetTOBLigne ( ARow : integer) : TOB ;
        procedure InitRow (Row : integer; GS : THGrid) ;
        Procedure CreerTOBLigne (ARow : integer; GS : THGrid);
        Function  LigVide( Row : integer; GS : THGRID) : Boolean ;
        procedure InsertLigne (ARow : Longint; GS : THGrid) ;
        procedure SupprimeLigne (ARow : Longint; GS : THGrid) ;
        Procedure ChercheArticle(GS : THGRID);
        Procedure ChercheTiers(GS : THGRID);
        Function  SortDeLaLigne : boolean ;
        Function  GrilleModifie : Boolean;
        Function  QuestionModifEnCours : Integer;
        //anipulation des Champs GRID
        procedure TraiterDepot (ACol, ARow : integer; GS : THGrid);
        procedure TraiterType (ACol, ARow : integer; GS : THGrid);
        procedure TraiterValeur (ACol, ARow : integer; GS : THGrid);
        procedure TraiterArticle (ACol, ARow : integer; GS : THGrid);
        procedure TraiterTiers(ACol, ARow : integer; GS : THGrid);
        // Validation
        procedure ValidePrix;
        procedure VerifTOB;
     public
        Action   : TActionFiche ;
        FindDebut : Boolean;
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure Onclose  ; override ;

     END ;

const colRang=1 ;
      NbRowsInit = 50 ;
      NbRowsPlus = 20 ;


function  CalculPrixRevient(Article, Tiers, Depot : string;  DPA : double) : Double;
function  WherePrixRevient(Article, Tiers, Depot : string) : String;


implementation

Procedure TOF_PRIXREVIENT.OnArgument (Arguments : String ) ;
var i,NbCol : integer ;
    St,Nam, DecPrix : string ;
begin
inherited ;
St:=Arguments ;
Action:=taModif ;
i:=Pos('ACTION=',St) ;
if i>0 then
   BEGIN
   System.Delete(St,1,i+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='CREATION' then BEGIN Action:=taCreat ; END ;
   if St='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if St='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;

St:=Arguments ;

DecPrix := '';
for i := 1 to GetParamSoc('SO_DECPRIX') do DecPrix := DecPrix + '0';

NbCol:=7;
LesColArt:='FIXED;PRV_RANG;PRV_ARTICLE;PRV_TIERS;PRV_DEPOT;PRV_TYPECALCUL;PRV_VALEUR' ;
LesColFou:='FIXED;PRV_RANG;PRV_TIERS;PRV_ARTICLE;PRV_DEPOT;PRV_TYPECALCUL;PRV_VALEUR' ;
LesColDep:='FIXED;PRV_RANG;PRV_DEPOT;PRV_ARTICLE;PRV_TIERS;PRV_TYPECALCUL;PRV_VALEUR' ;

BNewLine:=TToolbarButton97(GetControl('BNEWLINE'));
BNewLine.OnClick:=BNewLineClick;
BDelLine:=TToolbarButton97(GetControl('BDELLINE'));
BDelLine.OnClick:=BDelLineClick;
BChercher:=TToolbarButton97(GetControl('BCHERCHER'));
BChercher.OnClick:=BChercherClick;
FindLigne:=TFindDialog.Create(Ecran);
FindLigne.OnFind:=FindLigneFind ;

RBART:=TRadioButton(GetControl('RBART'));
RBART.OnClick:=RBARTClick;
RBFOU:=TRadioButton(GetControl('RBFOU'));
RBFOU.OnClick:=RBFOUClick;
RBDEP:=TRadioButton(GetControl('RBDEP'));
RBDEP.OnClick:=RBDEPClick;

PART:=THPanel(GetControl('PARTICLE'));
PFOU:=THPanel(GetControl('PFOUR'));
PDEP:=THPanel(GetControl('PDEPOT'));
PART.Align:=alClient;
PFOU.Align:=alClient;PFOU.Visible:=RBFOU.Checked;
PDEP.Align:=alClient;PDEP.Visible:=RBDEP.Checked;

G_ART:=THGRID(GetControl('G_ART'));
G_ART.OnElipsisClick:=G_ARTElipsisClick  ;
G_ART.OnDblClick:=G_ARTElipsisClick ;
G_ART.OnCellEnter:=G_ARTCellEnter ;
G_ART.OnCellExit:=G_ARTCellExit ;
G_ART.OnRowEnter:=G_ARTRowEnter ;
G_ART.OnRowExit:=G_ARTRowExit ;
G_ART.PostDrawCell:= DessineCellArt;
G_ART.ColCount:=NbCol;
G_ART.ColAligns[colRang]:=taCenter;
G_ART.ColTypes[colRang]:='I' ;
G_ART.ColWidths[1]:=0;   // rang
G_ART.ColWidths[0]:=15;

G_FOU:=THGRID(GetControl('G_FOU'));
G_FOU.OnElipsisClick:=G_FOUElipsisClick  ;
G_FOU.OnDblClick:=G_FOUElipsisClick ;
G_FOU.OnCellEnter:=G_FOUCellEnter ;
G_FOU.OnCellExit:=G_FOUCellExit ;
G_FOU.OnRowEnter:=G_FOURowEnter ;
G_FOU.OnRowExit:=G_FOURowExit ;
G_FOU.PostDrawCell:= DessineCellFou;
G_FOU.ColCount:=NbCol;
G_FOU.ColAligns[colRang]:=taCenter;
G_FOU.ColTypes[colRang]:='I' ;
G_FOU.ColWidths[1]:=0;   // rang
G_FOU.ColWidths[0]:=15;

G_DEP:=THGRID(GetControl('G_DEP'));
G_DEP.OnElipsisClick:=G_DEPElipsisClick  ;
G_DEP.OnDblClick:=G_DEPElipsisClick ;
G_DEP.OnCellEnter:=G_DEPCellEnter ;
G_DEP.OnCellExit:=G_DEPCellExit ;
G_DEP.OnRowEnter:=G_DEPRowEnter ;
G_DEP.OnRowExit:=G_DEPRowExit ;
G_DEP.PostDrawCell:= DessineCellDep;
G_DEP.ColCount:=NbCol;
G_DEP.ColAligns[colRang]:=taCenter;
G_DEP.ColTypes[colRang]:='I' ;
G_DEP.ColWidths[1]:=0;   // rang
G_DEP.ColWidths[0]:=15;

St:=LesColArt ;
for i:=0 to G_ART.ColCount-1 do
   BEGIN
   if i>1 then  G_ART.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='PRV_DEPOT' then BEGIN G_ART.ColFormats[i]:='CB=GCDEPOT'; ADEP:=i; END
   else if Nam='PRV_ARTICLE' then AArt:=i
   else if Nam='PRV_TIERS' then AFou:=i
   else if Nam='PRV_TYPECALCUL' then BEGIN G_ART.ColFormats[i]:='CB=GCTYPEPRIXREVIENT'; ATYP:=i; END
   else if Nam='PRV_VALEUR' then
        BEGIN G_ART.ColFormats[i]:='#,##0.' + DecPrix;G_ART.ColAligns[i]:=taRightJustify ;AVAL:=i; END
   ;
   END ;
St:=LesColFou ;
for i:=0 to G_FOU.ColCount-1 do
   BEGIN
   if i>1 then  G_FOU.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='PRV_DEPOT' then BEGIN G_FOU.ColFormats[i]:='CB=GCDEPOT'; FDEP:=i; END
   else if Nam='PRV_ARTICLE' then FArt:=i
   else if Nam='PRV_TIERS' then FFou:=i
   else if Nam='PRV_TYPECALCUL' then BEGIN G_FOU.ColFormats[i]:='CB=GCTYPEPRIXREVIENT'; FTYP:=i; END
   else if Nam='PRV_VALEUR' then
        BEGIN G_FOU.ColFormats[i]:='#,##0.' + DecPrix;G_FOU.ColAligns[i]:=taRightJustify ; FVAL:=i; END
   ;
   END ;
St:=LesColDep ;
for i:=0 to G_DEP.ColCount-1 do
   BEGIN
   if i>1 then  G_DEP.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='PRV_DEPOT' then BEGIN G_DEP.ColFormats[i]:='CB=GCDEPOT'; DDEP:=i; END
   else if Nam='PRV_ARTICLE' then DArt:=i
   else if Nam='PRV_TIERS' then DFou:=i
   else if Nam='PRV_TYPECALCUL' then BEGIN G_DEP.ColFormats[i]:='CB=GCTYPEPRIXREVIENT'; DTYP:=i; END
   else if Nam='PRV_VALEUR' then
        BEGIN G_DEP.ColFormats[i]:='#,##0.' + DecPrix;G_DEP.ColAligns[i]:=taRightJustify ; DVAL:=i; END
   ;
   END ;

AffecteGrid(G_ART,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_ART) ;
AffecteGrid(G_FOU,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_FOU) ;
AffecteGrid(G_DEP,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_DEP) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;

if ctxAffaire in V_PGI.PGIContexte then SetControlVisible ('RBDEP',False);
end;

procedure TOF_PRIXREVIENT.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
    GS : THGrid ;
BEGIN
GS:=Nil;
FocusGrid := False;
ARow:=0;
if(Screen.ActiveControl = G_Art) then
    BEGIN
    FocusGrid := True;
    GS := G_ART;
    ARow := G_Art.Row;
    END else if(Screen.ActiveControl = G_FOU) then
        BEGIN
        FocusGrid := True;
        GS := G_FOU;
        ARow := G_Fou.Row;
        END else if (Screen.ActiveControl = G_DEP) then
            BEGIN
            FocusGrid := True;
            GS := G_Dep;
            ARow := G_Dep.Row;
            END ;
Case Key of
    VK_F5 : if FocusGrid then GS.OnElipsisClick(Sender);
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if FocusGrid then
                    BEGIN
                    Key := 0;
                    InsertLigne (ARow, GS);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow, GS) ;
                    END ;
                END;
    END;
END;

Procedure TOF_PRIXREVIENT.OnLoad  ;
BEGIN
inherited ; 
Transactions (ChargeGrille, 1);
END;

Procedure TOF_PRIXREVIENT.OnUpdate  ;
BEGIN
inherited ;
Transactions (ValidePrix, 2);
TobPR.SetAllModifie(False);
END;

Procedure TOF_PRIXREVIENT.Onclose  ;
var Fermer : boolean ;
    St : string ;
begin
inherited ;
Fermer:=True ;
if GrilleModifie then
    BEGIN
    St:='1;?caption?;Confirmez-vous l''abandon de la saisie ?;Q;YN;Y;N;';
    if HShowMessage(St, Ecran.Caption, '') <> mrYes then Fermer:=False ;
    END;
if Fermer then
    BEGIN
    TobPR.free ; TobPR:=nil;
    TobToDelete.free ;
    FindLigne.Destroy;
    END else
    BEGIN
    AfficheError:=False;
    LastError:=1; LastErrorMsg:='Non Fermeture' ;
    END;
end ;

{==============================================================================================}
{================================= Evènements Entête ==========================================}
{==============================================================================================}
procedure TOF_PRIXREVIENT.RBARTClick(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
BEGIN
if PART.Visible then exit;
i_Rep := QuestionModifEnCours ;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValidePrix, 2);;
               Case ioerr of
                   oeOk : ;
                   oeUnknown : BEGIN MessageAlerte('Erreur') ; END ;
                   oeSaisie : BEGIN MessageAlerte('Erreur') ; END ;
                   END ;
               TobPR.free ; TobPR:=nil;
               TobToDelete.free ;
               PART.Visible:=True;
               PFOU.Visible:=False;
               PDEP.Visible:=False;
               Transactions (ChargeGrille, 1) ;
               END;
    mrNo     : BEGIN
               TobPR.free ; TobPR:=nil;
               TobToDelete.free ;
               PART.Visible:=True;
               PFOU.Visible:=False;
               PDEP.Visible:=False;
               Transactions (ChargeGrille, 1) ;
               END;
    mrCancel : BEGIN
               RBART.Checked:=PART.Visible;
               RBFOU.Checked:=PFOU.Visible;
               RBDEP.Checked:=PDEP.Visible;
               END;
    END ;
END;

procedure TOF_PRIXREVIENT.RBFOUClick(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
BEGIN
if PFOU.Visible then exit;
i_Rep := QuestionModifEnCours ;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValidePrix, 2);;
               Case ioerr of
                   oeOk : ;
                   oeUnknown : BEGIN MessageAlerte('Erreur') ; END ;
                   oeSaisie : BEGIN MessageAlerte('Erreur') ; END ;
                   END ;
               TobPR.free ; TobPR:=nil;
               TobToDelete.free ;
               PART.Visible:=False;
               PFOU.Visible:=True;
               PDEP.Visible:=False;
               Transactions (ChargeGrille, 1) ;
               END;
    mrNo     : BEGIN
               TobPR.free ; TobPR:=nil;
               TobToDelete.free ;
               PART.Visible:=False;
               PFOU.Visible:=True;
               PDEP.Visible:=False;
               Transactions (ChargeGrille, 1) ;
               END;
    mrCancel : BEGIN
               RBART.Checked:=PART.Visible;
               RBFOU.Checked:=PFOU.Visible;
               RBDEP.Checked:=PDEP.Visible;
               END;
    END ;
END;

procedure TOF_PRIXREVIENT.RBDEPClick(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
BEGIN
if PDEP.Visible then exit;
i_Rep := QuestionModifEnCours ;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValidePrix, 2);;
               Case ioerr of
                   oeOk : ;
                   oeUnknown : BEGIN MessageAlerte('Erreur') ; END ;
                   oeSaisie : BEGIN MessageAlerte('Erreur') ; END ;
                   END ;
               TobPR.free ; TobPR:=nil;
               TobToDelete.free ;
               PART.Visible:=False;
               PFOU.Visible:=False;
               PDEP.Visible:=True;
               Transactions (ChargeGrille, 1) ;
               END;
    mrNo     : BEGIN
               TobPR.free ; TobPR:=nil;
               TobToDelete.free ;
               PART.Visible:=False;
               PFOU.Visible:=False;
               PDEP.Visible:=True;
               Transactions (ChargeGrille, 1) ;
               END;
    mrCancel : BEGIN
               RBART.Checked:=PART.Visible;
               RBFOU.Checked:=PFOU.Visible;
               RBDEP.Checked:=PDEP.Visible;
               END;
    END ;
END;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}
procedure TOF_PRIXREVIENT.G_ARTElipsisClick(Sender: TObject);
begin
if G_ART.Col = AART then
    BEGIN
    ChercheArticle (G_ART);
    END ;
if G_ART.Col = AFOU then
    BEGIN
    ChercheTiers(G_ART);
    END ;
end;


procedure TOF_PRIXREVIENT.G_ARTCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
//    if (G_ART.Col <> AART) AND (G_ART.Cells [AART,G_ART.Row] = '') then BEGIN G_ART.Col := AART;Cancel:=true; END;
    if (G_ART.Col <> AART) AND (G_ART.Cells [AART,G_ART.Row] = '') then BEGIN G_ART.Col := AART; END;
    G_ART.ElipsisButton:=((G_ART.Col=AART) or (G_ART.Col=AFOU)) ;
    end ;
end;

procedure TOF_PRIXREVIENT.G_ARTCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if ACol = AART then TraiterArticle (ACol, ARow, G_ART) else
    if ACol = AFOU then TraiterTiers (ACol, ARow, G_ART) else
        if ACol = ADEP then TraiterDepot (ACol, ARow, G_ART) else
            if ACol = ATYP then TraiterType (ACol, ARow, G_ART) else
                if ACol = AVAL then
                    BEGIN
                    G_Art.Cells[ACol,ARow]:=StrF00(Valeur(G_Art.Cells[ACol,ARow]),2) ;
                    TraiterValeur (ACol, ARow, G_ART) ;
                    END;
end;


procedure TOF_PRIXREVIENT.G_ARTRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
G_ART.InvalidateRow(ou) ;
if Ou >= G_ART.RowCount - 1 then G_ART.RowCount := G_ART.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TobPR.detail.count + 1);
if (ARow = TobPR.detail.count + 1) AND (not LigVide (ARow - 1, G_ART)) then
    BEGIN
    CreerTOBligne (ARow, G_ART);
    END;
if Ou > TobPR.detail.count then
    BEGIN
    G_ART.Row := TobPR.detail.count;
    END;
end;

procedure TOF_PRIXREVIENT.G_ARTRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
G_ART.InvalidateRow(ou) ;
if LigVide (Ou, G_ART) Then G_ART.Row := Min (G_ART.Row,Ou);
end;

procedure TOF_PRIXREVIENT.G_FOUElipsisClick(Sender: TObject);
begin
if G_FOU.Col = FART then
    BEGIN
    ChercheArticle (G_FOU);
    END ;
if G_FOU.Col = FFOU then
    BEGIN
    ChercheTiers(G_FOU);
    END ;
end;


procedure TOF_PRIXREVIENT.G_FOUCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_FOU.Col <> FFOU) AND (G_FOU.Cells [FFOU,G_FOU.Row] = '') then BEGIN G_FOU.Col := FFOU;Cancel:=true; END;
    G_FOU.ElipsisButton:=((G_FOU.Col=FART) or (G_FOU.Col=FFOU)) ;
    end ;
end;

procedure TOF_PRIXREVIENT.G_FOUCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if ACol = FART then TraiterArticle (ACol, ARow, G_FOU) else
    if ACol = FFOU then TraiterTiers (ACol, ARow, G_FOU) else
        if ACol = FDEP then TraiterDepot (ACol, ARow, G_FOU) else
            if ACol = FTYP then TraiterType (ACol, ARow, G_FOU) else
                if ACol = FVAL then
                    BEGIN
                    G_FOU.Cells[ACol,ARow]:=StrF00(Valeur(G_FOU.Cells[ACol,ARow]),2) ;
                    TraiterValeur (ACol, ARow, G_FOU) ;
                    END;
end;


procedure TOF_PRIXREVIENT.G_FOURowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
G_FOU.InvalidateRow(ou) ;
if Ou >= G_FOU.RowCount - 1 then G_FOU.RowCount := G_FOU.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TobPR.detail.count + 1);
if (ARow = TobPR.detail.count + 1) AND (not LigVide (ARow - 1, G_FOU)) then
    BEGIN
    CreerTOBligne (ARow, G_FOU);
    END;
if Ou > TobPR.detail.count then
    BEGIN
    G_FOU.Row := TobPR.detail.count;
    END;
end;

procedure TOF_PRIXREVIENT.G_FOURowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
G_FOU.InvalidateRow(ou) ;
if LigVide (Ou, G_FOU) Then G_FOU.Row := Min (G_FOU.Row,Ou);
end;

procedure TOF_PRIXREVIENT.G_DEPElipsisClick(Sender: TObject);
begin
if G_DEP.Col = DART then
    BEGIN
    ChercheArticle (G_DEP);
    END ;
if G_DEP.Col = DFOU then
    BEGIN
    ChercheTiers(G_DEP);
    END ;
end;


procedure TOF_PRIXREVIENT.G_DEPCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_DEP.Col <> DDEP) AND (G_DEP.CellValues [DDEP,G_DEP.Row] = '') then BEGIN G_DEP.Col := DDEP; Cancel:=true; END;
    G_DEP.ElipsisButton:=((G_DEP.Col=DART) or (G_DEP.Col=DFOU)) ;
    end ;
end;

procedure TOF_PRIXREVIENT.G_DEPCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if ACol = DART then TraiterArticle (ACol, ARow, G_DEP) else
    if ACol = DFOU then TraiterTiers (ACol, ARow, G_DEP) else
        if ACol = DDEP then TraiterDepot (ACol, ARow, G_DEP) else
            if ACol = DTYP then TraiterType (ACol, ARow, G_DEP) else
                if ACol = DVAL then
                    BEGIN
                    G_DEP.Cells[ACol,ARow]:=StrF00(Valeur(G_DEP.Cells[ACol,ARow]),2) ;
                    TraiterValeur (ACol, ARow, G_DEP) ;
                    END;
end;


procedure TOF_PRIXREVIENT.G_DEPRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
G_DEP.InvalidateRow(ou) ;
if Ou >= G_DEP.RowCount - 1 then G_DEP.RowCount := G_DEP.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TobPR.detail.count + 1);
if (ARow = TobPR.detail.count + 1) AND (not LigVide (ARow - 1, G_DEP)) then
    BEGIN
    CreerTOBligne (ARow, G_DEP);
    END;
if Ou > TobPR.detail.count then
    BEGIN
    G_DEP.Row := TobPR.detail.count;
    END;
end;

procedure TOF_PRIXREVIENT.G_DEPRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
G_DEP.InvalidateRow(ou) ;
if LigVide (Ou, G_DEP) Then G_DEP.Row := Min (G_DEP.Row,Ou);
end;

procedure TOF_PRIXREVIENT.DessineCellArt(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < G_ART.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_ART.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_ART.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_ART.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

procedure TOF_PRIXREVIENT.DessineCellFou(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < G_FOU.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_FOU.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_FOU.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_FOU.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

procedure TOF_PRIXREVIENT.DessineCellDep(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < G_DEP.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_DEP.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_DEP.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_DEP.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

{==============================================================================================}
{================================ Actions liées au Grid =======================================}
{==============================================================================================}
procedure TOF_PRIXREVIENT.ChargeGrille ;
var QQ : TQuery ;
    St : string ;
BEGIN
if RBART.Checked then St := 'SELECT * FROM PRIXREVIENT WHERE PRV_ARTICLE<>"" ' +
                            'ORDER BY PRV_ARTICLE, PRV_TIERS, PRV_DEPOT, PRV_RANG'
    else if RBFOU.Checked then St := 'SELECT * FROM PRIXREVIENT WHERE PRV_TIERS<>"" ' +
                                     'ORDER BY PRV_TIERS, PRV_ARTICLE, PRV_DEPOT, PRV_RANG'
        else St := 'SELECT * FROM PRIXREVIENT WHERE PRV_DEPOT<>"" ' +
                   'ORDER BY PRV_DEPOT, PRV_ARTICLE, PRV_TIERS, PRV_RANG' ;
QQ:=OpenSql(St, True) ;
TobPR:=tob.create('',Nil,-1) ;
TobPR.LoadDetailDB('PRIXREVIENT','','',QQ,false,true) ;
Ferme(QQ) ;
If TobPR.detail.count=0 then begin Tob.create ('PRIXREVIENT',TobPR,-1) ; end ;
if RBART.Checked then
    BEGIN
    TobPR.PutGridDetail(G_ART,True,True,LesColArt,True);
    G_ART.RowCount:=Max (NbRowsInit, G_ART.RowCount+1) ;
    END else if RBFOU.Checked then
        BEGIN
        TobPR.PutGridDetail(G_FOU,True,True,LesColFou,True);
        G_FOU.RowCount:=G_FOU.RowCount+1 ;
        G_FOU.RowCount:=Max (NbRowsInit, G_FOU.RowCount+1) ;
        END else if RBDEP.Checked then
            BEGIN
            TobPR.PutGridDetail(G_DEP,True,True,LesColDep,True);
            G_DEP.RowCount:=G_DEP.RowCount+1 ;
            G_DEP.RowCount:=Max (NbRowsInit, G_DEP.RowCount+1) ;
            END;
TobToDelete:=tob.create('',Nil,-1) ;
END;

Function TOF_PRIXREVIENT.GetTOBLigne ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBPR.Detail.Count)) then Exit ;
Result:=TOBPR.Detail[ARow-1] ;
END ;

Procedure TOF_PRIXREVIENT.InitRow (Row : integer; GS : THGrid) ;
var Col : integer ;
    TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(Row) ; if TOBL<>Nil then TOBL.InitValeurs ;
for Col:=0 to GS.ColCount do GS.cells[Col,Row]:='';
END ;

Procedure TOF_PRIXREVIENT.CreerTOBLigne (ARow : integer; GS : THGrid);
BEGIN
if ARow <> TOBPR.Detail.Count + 1 then exit;
TOB.Create ('PRIXREVIENT', TOBPR, ARow-1) ;
InitRow (ARow, GS) ;
END;

Function TOF_PRIXREVIENT.LigVide(Row : integer; GS : THGRID) : Boolean ;
var Col : integer ;
BEGIN
Result:=True ;
if PART.Visible then  Col:=AART
    else if PFOU.Visible then  Col:=FFOU
        else if PDEP.Visible then  Col:=DDEP
          else Exit;
if (GS.Cells[Col,Row]<>'') then result:= False ;
END ;

procedure TOF_PRIXREVIENT.InsertLigne (ARow : Longint; GS : THGrid) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigVide (ARow, GS) then exit;
if (ARow > TOBPR.Detail.Count) then Exit;
GS.CacheEdit; GS.SynEnabled := False;
TOB.Create ('PRIXREVIENT', TOBPR, ARow-1) ;
GS.InsertRow (ARow); GS.Row := ARow;
InitRow (ARow, GS) ;
GS.MontreEdit; GS.SynEnabled := True;
END;

procedure TOF_PRIXREVIENT.SupprimeLigne (ARow : Longint; GS : THGrid) ;
Var i_ind: integer;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBPR.Detail.Count) then Exit;
GS.CacheEdit; GS.SynEnabled := False;
GS.DeleteRow (ARow);
if (ARow = TOBPR.Detail.Count) then CreerTOBLigne (ARow + 1, GS);
if TOBPR.Detail[ARow-1].GetValue ('PRV_RANG') <> 0 then
    BEGIN
    i_ind := TobToDelete.Detail.Count;
    TOB.Create ('PRIXREVIENT', TobToDelete, i_ind) ;
    TobToDelete.Detail[i_ind].Dupliquer (TOBPR.Detail[ARow-1], False, True);
    END;
TOBPR.Detail[ARow-1].Free;
if GS.RowCount < NbRowsInit then GS.RowCount := NbRowsInit;
GS.MontreEdit; GS.SynEnabled := True;
END;

Procedure TOF_PRIXREVIENT.ChercheArticle(GS : THGRID);
Var ARTICLE : THCritMaskEdit;
    Coord : TRect;
BEGIN
Coord := GS.CellRect (GS.Col, GS.Row);
ARTICLE := THCritMaskEdit.Create (ECRAN);
ARTICLE.Parent := GS;
ARTICLE.Top := Coord.Top;
ARTICLE.Left := Coord.Left;
ARTICLE.Width := 3; ARTICLE.Visible := False;
ARTICLE.Text:= GS.Cells[GS.Col,GS.Row] ;
ARTICLE.DataType:='GCARTICLEGENERIQUE';
DispatchRechArt (ARTICLE, 1, '',
                 'GA_CODEARTICLE=' + Trim (Copy (ARTICLE.Text, 1, 18)), '');
if ARTICLE.Text <> '' then GS.Cells[GS.Col,GS.Row]:= Trim (Copy (ARTICLE.Text, 1, 18));
ARTICLE.Destroy;
END ;

Procedure TOF_PRIXREVIENT.ChercheTiers(GS : THGRID);
Var TIERS : THCritMaskEdit;
    Coord : TRect;
BEGIN
Coord := GS.CellRect (GS.Col, GS.Row);
TIERS := THCritMaskEdit.Create (ECRAN);
TIERS.Parent := GS;
TIERS.Top := Coord.Top;
TIERS.Left := Coord.Left;
TIERS.Width := 3; TIERS.Visible := False;
TIERS.Text:= GS.Cells[GS.Col,GS.Row] ;
TIERS.DataType:='GCTIERSFOURN';
DispatchRecherche (TIERS, 2, 'T_NATUREAUXI="FOU"','T_TIERS=' + Trim (TIERS.Text), '');
if TIERS.Text <> '' then GS.Cells[GS.Col,GS.Row]:= TIERS.Text;
TIERS.Destroy;
END ;

Function TOF_PRIXREVIENT.SortDeLaLigne : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
if RBART.Checked then
    BEGIN
    ACol:=G_Art.Col ; ARow:=G_Art.Row ; Cancel:=False ;
    G_ArtCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
    G_ArtRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
    END else if RBFOU.Checked then
        BEGIN
        ACol:=G_Fou.Col ; ARow:=G_Fou.Row ; Cancel:=False ;
        G_FouCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
        G_FouRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
        END else if RBDEP.Checked then
            BEGIN
            ACol:=G_Dep.Col ; ARow:=G_Dep.Row ; Cancel:=False ;
            G_DepCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
            G_DepRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
            END;
Result:=True ;
END ;

Function TOF_PRIXREVIENT.GrilleModifie : Boolean;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=(TOBPR.IsOneModifie) or (TobToDelete.IsOneModifie);
END;

Function TOF_PRIXREVIENT.QuestionModifEnCours : Integer;
var St : string ;
BEGIN
Result := mrNo;
if Action = taConsult then Exit;
if GrilleModifie then
    BEGIN
    St:='0;?CAPTION?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;C;';
    Result := HShowMessage(St, Ecran.Caption, '');
    END;
END;

{==============================================================================================}
{============================ Manipulation des Champs GRID ====================================}
{==============================================================================================}
procedure TOF_PRIXREVIENT.TraiterDepot (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := GS.CellValues [ACol, ARow];
if St <> '' then
    BEGIN
    TOBL.PutValue ('PRV_DEPOT', St);
    if (RBDEP.Checked) AND (TOBL.GetValue('PRV_TYPECALCUL')='') then
        BEGIN
        TOBL.PutValue ('PRV_TYPECALCUL', 'MON');
        GS.CellValues [DTYP, ARow]:=TOBL.GetValue('PRV_TYPECALCUL');
        END;
    END;
END;

procedure TOF_PRIXREVIENT.TraiterType (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := GS.CellValues [ACol, ARow];
TOBL.PutValue ('PRV_TYPECALCUL', St);
END;

procedure TOF_PRIXREVIENT.TraiterValeur (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
TOBL.PutValue ('PRV_VALEUR', Valeur (GS.Cells [ACol, ARow]));
END;

procedure TOF_PRIXREVIENT.TraiterArticle (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
    QQ : TQuery ;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := UpperCase(GS.Cells [ACol, ARow]);
GS.Cells [ACol, ARow] := St;
if (RBART.Checked) AND (St='') then
    BEGIN
    GS.Cells [ACol, ARow]:= TOBL.GetValue ('PRV_ARTICLE') ;
    exit ;
    END;
if St='' then BEGIN TOBL.PutValue ('PRV_ARTICLE', St); exit ; END ;

if TOBL.GetValue ('PRV_ARTICLE') <> St then
    BEGIN
    QQ := OpenSQL('Select * from ARTICLE Where GA_CODEARTICLE="' +
                   St + '" AND GA_STATUTART <> "DIM" ',True) ;
    if Not QQ.EOF then
        BEGIN
        TOBL.PutValue ('PRV_ARTICLE', St);
        if (RBART.Checked) AND (TOBL.GetValue('PRV_TYPECALCUL')='') then
            BEGIN
            TOBL.PutValue ('PRV_TYPECALCUL', 'MON');
            GS.CellValues [ATYP, ARow]:=TOBL.GetValue('PRV_TYPECALCUL');
            END;
        END else
        BEGIN
        GS.Col:=ACol; GS.Row:=ARow ;
        ChercheArticle (GS);
        END;
    Ferme (QQ);
    END;
END;

procedure TOF_PRIXREVIENT.TraiterTiers(ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
    QQ : TQuery ;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := GS.Cells [ACol, ARow];
if (RBFOU.Checked) AND (St='') then
    BEGIN
    GS.Cells [ACol, ARow]:= TOBL.GetValue ('PRV_TIERS') ;
    exit ;
    END;
if St='' then BEGIN TOBL.PutValue ('PRV_TIERS', St); exit ; END ;

if TOBL.GetValue ('PRV_TIERS') <> St then
    BEGIN
    QQ := OpenSQL('Select * from TIERS Where T_TIERS="' +
                   St + '" AND T_NATUREAUXI= "FOU" ',True) ;
    if Not QQ.EOF then
        BEGIN
        TOBL.PutValue ('PRV_TIERS', St);
        if (RBFOU.Checked) AND (TOBL.GetValue('PRV_TYPECALCUL')='') then
            BEGIN
            TOBL.PutValue ('PRV_TYPECALCUL', 'MON');
            GS.CellValues [FTYP, ARow]:=TOBL.GetValue('PRV_TYPECALCUL');
            END;
        END else
        BEGIN
        GS.Col:=ACol; GS.Row:=ARow ;
        ChercheTiers (GS);
        END;
    Ferme (QQ);
    END;
END;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TOF_PRIXREVIENT.BChercherClick(Sender: TObject);
begin
if RBART.Checked then
    if G_Art.RowCount < 3 then Exit;
if RBFOU.Checked then
    if G_Fou.RowCount < 3 then Exit;
if RBDEP.Checked then
    if G_Dep.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TOF_PRIXREVIENT.FindLigneFind(Sender: TObject);
begin
if RBART.Checked then
    Rechercher (G_Art, FindLigne, FindDebut) ;
if RBFOU.Checked then
    Rechercher (G_Fou, FindLigne, FindDebut) ;
if RBDEP.Checked then
    Rechercher (G_Dep, FindLigne, FindDebut) ;
end;

procedure TOF_PRIXREVIENT.BNewLineClick(Sender: TObject);
BEGIN
if(Screen.ActiveControl=G_Art) then  InsertLigne (G_Art.Row, G_ART)
    else if(Screen.ActiveControl=G_FOU) then InsertLigne (G_FOU.Row, G_FOU)
        else if (Screen.ActiveControl=G_DEP) then InsertLigne (G_DEP.Row, G_DEP);
        //else if G_DEP.focused then InsertLigne (G_DEP.Row, G_DEP);
end;

procedure TOF_PRIXREVIENT.BDelLineClick(Sender: TObject);
begin
if(Screen.ActiveControl = G_Art) then  SupprimeLigne (G_Art.Row, G_ART)
    else if(Screen.ActiveControl = G_FOU) then SupprimeLigne (G_FOU.Row, G_FOU)
        else if (Screen.ActiveControl = G_DEP) then SupprimeLigne (G_DEP.Row, G_DEP);
end;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}
procedure TOF_PRIXREVIENT.ValidePrix;
begin
if Not SortDeLaLigne then Exit ;
TobToDelete.DeleteDB (False);
VerifTOB;
TOBPR.InsertOrUpdateDB(False) ;
end;

procedure TOF_PRIXREVIENT.VerifTOB;
var i_ind : integer;
    QQ : TQuery ;
    MaxRang : Longint ;
    GS : THGrid ;
BEGIN
  QQ := OpenSQL('SELECT MAX(PRV_RANG) FROM PRIXREVIENT',true);
  if QQ.EOF then MaxRang := 1 else MaxRang := QQ.Fields[0].AsInteger + 1 ;
  Ferme(QQ) ;
  if PART.Visible then  GS:=G_Art
  else if PFOU.Visible then GS:=G_Fou
  else if PDEP.Visible then  GS:=G_Dep
  else exit ;
  // DBR - Fiche 10413 - Si on fait un Downto l'ordre est inversé. Le calcul du PR n'est plus le meme
  for i_ind := TOBPR.Detail.count - 1 Downto 0 do // Downto pour supprimer les lignes vides
(*    BEGIN
    if LigVide (i_ind + 1,GS) then
        BEGIN
        TOBPR.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBPR.Detail[i_ind].GetValue ('PRV_RANG') = 0 then
            BEGIN
            TOBPR.Detail[i_ind].PutValue ('PRV_RANG', MaxRang);
            Inc (MaxRang);
            END;
        END;
    END; *)
  BEGIN
    if LigVide (i_ind + 1,GS) then
    BEGIN
      TOBPR.Detail[i_ind].Free ;
    END
  end;
  for i_ind := 0 to TOBPR.Detail.count - 1 do // DBR puis dans l'ordre pour enregistrer
  BEGIN
    if TOBPR.Detail[i_ind].GetValue ('PRV_RANG') = 0 then
    BEGIN
      TOBPR.Detail[i_ind].PutValue ('PRV_RANG', MaxRang);
      Inc (MaxRang);
    END;
  END;
END;

{==============================================================================================}
{============================ Calcul du prix de revient =======================================}
{==============================================================================================}
function  WherePrixRevient(Article, Tiers, Depot : string) : String;
Var St_Where : string ;
BEGIN
St_Where := '(PRV_ARTICLE="'+ Article +
            '" and PRV_TIERS="'+Tiers+'" and PRV_DEPOT="'+Depot+'") ';
St_Where := St_Where + 'OR (PRV_ARTICLE="'+ Article +
            '" and PRV_TIERS="'+Tiers+'" and PRV_DEPOT="") ';
St_Where := St_Where + 'OR (PRV_ARTICLE="'+ Article +
            '" and PRV_TIERS="" and PRV_DEPOT="'+Depot+'") ';
St_Where := St_Where + 'OR (PRV_ARTICLE="" and PRV_TIERS="'+
            Tiers+'" and PRV_DEPOT="'+Depot+'") ';
St_Where := St_Where + 'OR (PRV_ARTICLE="'+  Article +
           '" and PRV_TIERS="" and PRV_DEPOT="") ';
St_Where := St_Where + 'OR (PRV_ARTICLE="" and PRV_TIERS="'+
            Tiers+'" and PRV_DEPOT="") ';
St_Where := St_Where + 'OR (PRV_ARTICLE="" and PRV_TIERS="'+
            '" and PRV_DEPOT="'+Depot+'") ';
St_Where := St_Where + 'OR (PRV_ARTICLE="" and PRV_TIERS="" and PRV_DEPOT="") ';
Result:=St_Where;
END;

function  CalculPrixRevient(Article, Tiers, Depot : string;  DPA : double) : Double;
var TobGr, TobLig : TOB;
    st1, CodeArt : string;
    Q : TQuery;
    i_ind : integer;
    PR, Valeur : double;
begin
//Modif 20/02/2003
PR := DPA;
Result := PR;
if (ctxMode in V_PGI.PGIContexte) then exit;
CodeArt:=Trim(copy(Article,1,18)) ;
st1 := 'Select * from PRIXREVIENT Where ' + WherePrixRevient(CodeArt, Tiers, Depot) +
       ' ORDER BY PRV_RANG';
Q := OpenSQL(st1, True);
if Q.EOF then begin Ferme(Q); Exit; end;
TobGr := TOB.Create('PRIXREVIENT', nil, -1);
TobGr.LoadDetailDB('PRIXREVIENT', '', '', Q, False);
Ferme(Q);
i_ind:=1;
TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,Tiers,Depot], TRUE) ;
if TobLig=Nil then
    BEGIN
    TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,Tiers,''], TRUE) ;
    i_ind:=2;
    END;
if TobLig=Nil then
    BEGIN
    TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,'',Depot], TRUE) ;
    i_ind:=3;
    END;
if TobLig=Nil then
    BEGIN
    TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['',Tiers,Depot], TRUE) ;
    i_ind:=4;
    END;
if TobLig=Nil then
    BEGIN
    TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,'',''], TRUE) ;
    i_ind:=5;
    END;
if TobLig=Nil then
    BEGIN
    TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['',Tiers,''], TRUE) ;
    i_ind:=6;
    END;
if TobLig=Nil then
    BEGIN
    TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['','',Depot], TRUE) ;
    i_ind:=7;
    END;
if TobLig=Nil then
    BEGIN
    TobLig := TobGr.FindFirst(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['','',''], TRUE) ;
    i_ind:=8;
    END;
While TobLig<>Nil do
    begin
    Valeur := TobLig.GetValue('PRV_VALEUR');
    if Trim(TobLig.GetValue('PRV_TYPECALCUL')) = 'POU' then
        begin
        PR := PR * (1 + (Valeur / 100.0));
        end else if Trim(TobLig.GetValue('PRV_TYPECALCUL')) = 'COE' then
            begin
            PR := (PR * Valeur);
            end else if Trim(TobLig.GetValue('PRV_TYPECALCUL')) = 'MON' then
                PR := PR + Valeur ;

    Case i_ind of
        1: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,Tiers,Depot], TRUE) ;
        2: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,Tiers,''], TRUE) ;
        3: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,'',Depot], TRUE) ;
        4: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['',Tiers,Depot], TRUE) ;
        5: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],[CodeArt,'',''], TRUE) ;
        6: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['',Tiers,''], TRUE) ;
        7: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['','',Depot], TRUE) ;
        8: TobLig := TobGr.FindNext(['PRV_ARTICLE','PRV_TIERS','PRV_DEPOT'],['','',''], TRUE) ;
        end;
    end;
TobGr.Free; 
Result := PR;
end;

Initialization
registerclasses([TOF_PRIXREVIENT]);

end.

