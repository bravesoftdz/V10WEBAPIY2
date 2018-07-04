unit UTofGcFamHier;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Hpanel, Math
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,FactUtil,Ent1,
{$IFDEF EAGLCLIENT}
     HPdfPrev,UtileAGL,Maineagl,eTablette,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,EdtEtat,FE_Main,Tablette,
{$ENDIF}
{$IFNDEF SANSCOMPTA}
      Ventil,
{$ENDIF}
      M3FP,HTB97,Dialogs, AGLInitGC;

Type
     TOF_GCFamHier = Class (TOF)
     private
        LesColArt, CodeFamSup, NvSup, NvSaisi,TypeAct : string ;
        imprim_etiq, valide_NV2,valide_NV3 : integer;
        GFAM : THGRID ;
        ATYP,ACOD,ALIB,AABR,ALBR: integer ;
        LesColArt2, NvSaisi2 : string ;
        GFAM2 : THGRID ;
        ATYP2,ACOD2,ALIB2,AABR2,ALBR2: integer ;
//        BNewLine: TToolbarButton97;
//        BDelLine: TToolbarButton97;
        BInsert: TToolbarButton97;
        BDelete: TToolbarButton97;
        BChercher: TToolbarButton97;
        BImprimer: TToolbarButton97;
        BValider: TToolbarButton97;
        FindLigne: TFindDialog;
        TobPR, TobToDelete: TOB ;
        TobPR2, TobToDelete2: TOB ;
        procedure BNewLineClick(Sender: TObject);
        procedure BDelLineClick(Sender: TObject);
        procedure BChercherClick(Sender: TObject);
        procedure BValiderClick(Sender: TObject);
        procedure FindLigneFind(Sender: TObject);
        procedure GFAMElipsisClick(Sender: TObject);
        procedure GFAMCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GFAMCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GFAMRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GFAMRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure DessineCellFam (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
        // Actions liées au grid
        procedure ChargeGrille;
        Function  GetTOBLigne ( ARow : integer) : TOB ;
        procedure InitRow (Row : integer; GS : THGrid) ;
        Procedure CreerTOBLigne (ARow : integer; GS : THGrid);
        Function  LigVide( Row : integer; GS : THGRID) : Boolean ;
        procedure InsertLigne (ARow : Longint; GS : THGrid) ;
        procedure SupprimeLigne (ARow : Longint; GS : THGrid) ;
        //anipulation des Champs GRID
        Procedure ChercheFamille(GS : THGRID);
        procedure TraiterLIB (ACol, ARow : integer; GS : THGrid);
        procedure TraiterABR (ACol, ARow : integer; GS : THGrid);
        procedure TraiterCODE (ACol, ARow : integer; GS : THGrid);
        // Validation
        procedure ValideFamille;
        Procedure AfficheFAM(CodeFam: string;GS : THGRID);
        function  ExisteCode(Code : String; Ligne : Integer) : boolean;
        procedure MajLesFamillesToDelete;
        procedure GFAM2ElipsisClick(Sender: TObject);
        procedure GFAM2CellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GFAM2CellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GFAM2RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GFAM2RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure DessineCellFam2 (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
        // Actions liées au grid
        procedure ChargeGrille2 (CodNv2 : string);
        Function  GetTOBLigne2 ( ARow : integer) : TOB ;
        procedure InitRow2 (Row : integer; GS : THGrid) ;
        Procedure CreerTOBLigne2 (ARow : integer; GS : THGrid);
        Function  LigVide2( Row : integer; GS : THGRID) : Boolean ;
        procedure InsertLigne2 (ARow : Longint; GS : THGrid) ;
        procedure SupprimeLigne2 (ARow : Longint; GS : THGrid) ;
        //anipulation des Champs GRID
        Procedure ChercheFamille2(GS : THGRID);
        procedure NV3TraiterLIB (ACol, ARow : integer; GS : THGrid);
        procedure NV3TraiterABR (ACol, ARow : integer; GS : THGrid);
        procedure NV3TraiterCODE (ACol, ARow : integer; GS : THGrid);
        // Validation
        procedure ValideFamille2;
        Procedure AfficheFAM2(CodeFam: string;GS : THGRID);
        function  ExisteCode2(Code : String; Ligne : Integer) : boolean;
        procedure MajLesFamillesToDelete2;
    procedure ZoneSuivanteOuOk(GS: THGrid; var ACol, ARow: Integer;
      var Cancel: boolean);
    function ZoneAccessible(GS: THGrid; ACol, ARow: Integer): boolean;
    procedure DellesLigne2;
     public
        Action   : TActionFiche ;
        FindDebut : Boolean;
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure FormKeyDown2(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure Onclose  ; override ;

     END ;


const colRang=1 ;
      NbRowsInit = 1;
      NbRowsPlus = 1 ;
      NbRowsInit2 = 2 ;
      NbRowsPlus2 = 1 ;


implementation

Procedure TOF_GCFamHier.OnArgument (Arguments : String ) ;
var i,NbCol : integer ;
    St,Critere,Nam,ChampMul,ValMul : string ;
begin
inherited ;
Action:=taModif ;
//ecran.width := ecran.Width - 40;
Repeat
  Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
  if Critere<>'' then
    begin
    i:=pos('=',Critere);
    if i<>0 then
      begin
      ChampMul:=copy(Critere,1,i-1);
      ValMul:=copy(Critere,i+1,length(Critere));
      if ChampMul='ACTION' then
       begin
       if ValMul='CREATION' then  Action:=taCreat
       else if ValMul='MODIFICATION' then Action:=taModif
       else if ValMul='CONSULTATION' then Action:=taConsult;
       end;
      if ChampMul='NIV1'    then NvSup    :=ValMul;
      if ChampMul='NIV2'    then NvSaisi  :=ValMul;
      if ChampMul='NIV3'    then NvSaisi2 :=ValMul;
      if ChampMul='TYPE'    then TypeAct  :=ValMul;
      if ChampMul='FAMNV1'  then CodeFamSup :=ValMul;
      end;
    end;
until  Critere='' ;
THValComboBox(GetControl('FAMNVSUP')).DataType:='GCFAMILLENIV'+Copy(NvSup,3,1);

// gestion de famille ouvrage NIV1 BATIPRIX
If NvSup = 'BO1' then
	Begin
  		THValComboBox(GetControl('FAMNVSUP')).DataType:='BTFAMILLEOUV1';

  end;
If NvSup = 'BP1' then
	Begin
  		THValComboBox(GetControl('FAMNVSUP')).DataType:='BTFAMILLEPARC1';

  end;

NbCol:=6;
LesColArt:='FIXED;CC_TYPE;CC_CODE;CC_LIBELLE;CC_ABREGE;CC_LIBRE' ;
LesColArt2:='FIXED;CC_TYPE;CC_CODE;CC_LIBELLE;CC_ABREGE;CC_LIBRE' ;

//BNewLine:=TToolbarButton97(GetControl('BNEWLINE'));
//BNewLine.OnClick:=BNewLineClick;
//BDelLine:=TToolbarButton97(GetControl('BDELLINE'));
//BDelLine.OnClick:=BDelLineClick;
BInsert:=TToolbarButton97(GetControl('BINSERT'));
BInsert.OnClick:=BNewLineClick;
BDelete:=TToolbarButton97(GetControl('BDELETE'));
BDelete.OnClick:=BDelLineClick;
BChercher:=TToolbarButton97(GetControl('BCHERCHER'));
BChercher.OnClick:=BChercherClick;
BImprimer:=TToolbarButton97(GetControl('BIMPRIMER'));
FindLigne:=TFindDialog.Create(Ecran);
FindLigne.OnFind:=FindLigneFind ;
BValider:=TToolbarButton97(GetControl('BValider'));
BValider.OnClick:=BValiderClick;

GFAM:=THGRID(GetControl('GFAM'));
GFAM.OnCellEnter:=GFAMCellEnter ;
GFAM.OnCellExit:=GFAMCellExit ;
GFAM.OnRowEnter:=GFAMRowEnter ;
GFAM.OnRowExit:=GFAMRowExit ;
GFAM.PostDrawCell:= DessineCellFam;
GFAM.OnElipsisClick:=GFAMElipsisClick  ;
GFAM.ColCount:=NbCol;
//
GFAM2:=THGRID(GetControl('GFAM2'));
GFAM2.OnCellEnter:=GFAM2CellEnter ;
GFAM2.OnCellExit:=GFAM2CellExit ;
GFAM2.OnRowEnter:=GFAM2RowEnter ;
GFAM2.OnRowExit:=GFAM2RowExit ;
GFAM2.PostDrawCell:= DessineCellFam2;
GFAM2.OnElipsisClick:=GFAM2ElipsisClick  ;
GFAM2.ColCount:=NbCol;


St:=LesColArt ;
GFAM.ColWidths[0]:=8;
for i:=0 to GFAM.ColCount-1 do
   BEGIN
//   if i>1 then  GFAM.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='CC_TYPE' then
     BEGIN
     GFAM.ColWidths[i]:=0;
     GFAM.ColLengths[i]:=-1;
     ATYP:=i;
     END
   else if Nam='CC_CODE' then
     BEGIN
     GFAM.ColWidths[i]:=50;
     GFAM.ColFormats[i]:='UPPER';
     GFAM.ColAligns[i]:=taLeftJustify ;
     GFAM.ColLengths[i]:=3;
     ACOD:=i;
     END
   else if Nam='CC_LIBELLE' then
     BEGIN
//     GFAM.ColWidths[i]:=150;
     GFAM.ColWidths[i]:=150;
     GFAM.ColAligns[i]:=taLeftJustify ;
     GFAM.ColLengths[i]:=35;
     ALIB:=i;
     END
   else if Nam='CC_ABREGE' then
     BEGIN
     GFAM.ColWidths[i]:=80;
     GFAM.ColAligns[i]:=taLeftJustify ;
     GFAM.ColLengths[i]:=17;
     AABR:=i;
     END
   else if Nam='CC_LIBRE' then
     BEGIN
     GFAM.ColWidths[i]:=0;
     GFAM.ColLengths[i]:=-1;
     ALBR:=i;
     END;
   END ;

St:=LesColArt2 ;
GFAM2.ColWidths[0]:=8;
for i:=0 to GFAM2.ColCount-1 do
   BEGIN
//   if i>1 then  GFAM2.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='CC_TYPE' then
     BEGIN
     GFAM2.ColWidths[i]:=0;
     GFAM2.ColLengths[i]:=-1;
     ATYP2:=i;
     END
   else if Nam='CC_CODE' then
     BEGIN
     GFAM2.ColWidths[i]:=50;
     GFAM2.ColFormats[i]:='UPPER';
     GFAM2.ColAligns[i]:=taLeftJustify ;
     GFAM2.ColLengths[i]:=3;
     ACOD2:=i;
     END
   else if Nam='CC_LIBELLE' then
     BEGIN
     GFAM2.ColWidths[i]:=150;
     GFAM2.ColAligns[i]:=taLeftJustify ;
     GFAM2.ColLengths[i]:=35;
     ALIB2:=i;
     END
   else if Nam='CC_ABREGE' then
     BEGIN
     GFAM2.ColWidths[i]:=80;
     GFAM2.ColAligns[i]:=taLeftJustify ;
     GFAM2.ColLengths[i]:=17;
     AABR2:=i;
     END
   else if Nam='CC_LIBRE' then
     BEGIN
     GFAM2.ColWidths[i]:=0;
     GFAM2.ColLengths[i]:=-1;
     ALBR2:=i;
     END;
   END ;

imprim_etiq := 0;
valide_Nv2 := 1;
valide_Nv3 := 1;
AffecteGrid(GFAM,Action) ;
//TFVierge(Ecran).Hmtrad.ResizeGridColumns(GFAM) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
AffecteGrid(GFAM2,Action) ;
//TFVierge(Ecran).Hmtrad.ResizeGridColumns(GFAM2) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown2 ;

If NvSup = 'BO1' then TFVierge(Ecran).Caption := 'Familles Ouvrages'
else TFVierge(Ecran).Caption := 'Familles Articles';
UpdateCaption(TFVierge(Ecran));

end;

procedure TOF_GCFamHier.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
    GS : THGrid ;
BEGIN
FocusGrid := True;
GS := GFAM;
ARow := GFAM.Row;
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

Procedure TOF_GCFamHier.OnLoad  ;
var  CC : THLabel ;
     F : TFVierge ;
     tmp : string;
     CB : THValComboBox;
BEGIN
inherited ;
F:=TFVierge(Ecran) ;
CB:=THValComboBox(GetControl('FAMNVSUP'));
if CodeFamSup='' then
  begin
//  if THValComboBox(GetControl('FAMNVSUP')).Items.Count<>0 then THValComboBox(GetControl('FAMNVSUP')).ItemIndex:=0;
  if CB.Items.Count<>0 then CB.ItemIndex:=0;
  CodeFamSup:=CB.Value;
  end
else CB.Value:=CodeFamSup;
if TypeAct = 'SAIART' then
  begin
  CB.Enabled :=False;
  CB.DataTypeParametrable :=False;
  end;
  if NvSup = 'BO1' then
  begin
		CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP'));
    CC.Caption:=RechDom('GCLIBOUVRAGE','BP'+Copy(NvSup,3,1),false);
    CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP2'));
    CC.Caption:=RechDom('GCLIBOUVRAGE','BP'+Copy(NvSaisi,3,1),false);
    CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP3'));
    CC.Caption:=RechDom('GCLIBOUVRAGE','BP'+Copy(NvSaisi2,3,1),false);
  end else if NvSup = 'BP1' then
  begin
		CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP'));
    CC.Caption:=RechDom('GCLIBFAMILLE','PA'+Copy(NvSup,3,1),false);
    CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP2'));
    CC.Caption:=RechDom('GCLIBFAMILLE','PA'+Copy(NvSaisi,3,1),false);
    CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP3'));
    CC.Caption:=RechDom('GCLIBFAMILLE','PA'+Copy(NvSaisi2,3,1),false);
  end else
  begin
    CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP'));
    CC.Caption:=RechDom('GCLIBFAMILLE','LF'+Copy(NvSup,3,1),false);
    CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP2'));
    CC.Caption:=RechDom('GCLIBFAMILLE','LF'+Copy(NvSaisi,3,1),false);
    CC:=THLabel(TFVierge(F).FindComponent('TFAMNVSUP3'));
    CC.Caption:=RechDom('GCLIBFAMILLE','LF'+Copy(NvSaisi2,3,1),false);
	end;
ChargeGrille;
// Permet d'afficher en gras le code de la première ligne du Grid
GFAM.Col:=ACOD;
if GFAM.Cells [ACOD,1] = '' then
  begin
  tmp := CodeFamSup+'***';
  GFAM2.Enabled := False;
  end else
  begin
  tmp := CodeFamSup+GFAM.Cells [ACOD,1];
  GFAM2.Enabled := True;
  GFAM.Col := ALIB;
  end;
GFAM.ElipsisButton:=(GFAM.Col=ACOD);
GFAM2.Col:=ACOD2;
GFAM2.ElipsisButton:=(GFAM2.Col=ACOD2);
ChargeGrille2 (tmp);
//
//GFAM2.SetFocus;
END;

Procedure TOF_GCFamHier.OnUpdate  ;
BEGIN
inherited ;
ValideFamille;
TobPR.SetAllModifie(False);
ValideFamille2;
TobPR2.SetAllModifie(False);
END;

Procedure TOF_GCFamHier.Onclose  ;
var Fermer : boolean ;
    St : string ;
begin
inherited ;
Fermer:=True ;
if (valide_Nv2 = 0) or (valide_Nv3 = 0) then
    BEGIN
    St:='1;?caption?;Confirmez-vous l''abandon de la saisie ?;Q;YN;Y;N;';
    if HShowMessage(St, Ecran.Caption, '') <> mrYes then Fermer:=False ;
    END;
if (Fermer) then
    BEGIN
    TobPR.free ; TobPR:=nil;
    TobToDelete.free ;
    TobPR2.free ; TobPR2:=nil;
    TobToDelete2.free ;
    FindLigne.Destroy;
    END else
    BEGIN
    AfficheError:=False;
    LastError:=1; LastErrorMsg:='Non Fermeture' ;
    END;
end ;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}
procedure TOF_GCFamHier.GFAMElipsisClick(Sender: TObject);
begin
if GFAM.Col = ACOD then
    BEGIN
    ChercheFamille (GFAM);
    END ;
end;


procedure TOF_GCFamHier.GFAMCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var tmp : string;
		Acol2 : integer;
    ARow2 : integer;
begin
if Action=taConsult then Exit ;
if GFAM.Cells [ACOD,GFAM.Row] = '' then valide_Nv2 := 0;
ZoneSuivanteOuOk (GFam,Acol,Arow,Cancel);
GFAM.ElipsisButton:=(GFAM.Col=ACOD);
if cancel then exit;
if GFAM.Cells [ACOD,GFAM.Row] = '' then
  begin
  tmp := CodeFamSup+'***';
  GFAM2.Enabled := False;
  end else
  begin
  tmp := CodeFamSup+GFAM.Cells [ACOD,GFAM.Row];
  GFAM2.Enabled := True;
//  GFAM.Col := ALIB;
  end;
if Not Cancel then
    BEGIN
//    if (GFAM.Col <> ACOD) AND (GFAM.Cells [ACOD,GFAM.Row] = '') then BEGIN GFAM.Col := ACOD; cancel := true; Exit; END;
    GFAM.ElipsisButton:=(GFAM.Col=ACOD);
    end ;
ChargeGrille2 (tmp);
end;

procedure TOF_GCFamHier.GFAMCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if ACol = ACOD then
  begin
  if ExisteCode(GFAM.Cells[ACOD,ARow],ARow) then
    begin
    PGIInfo('Ce code existe déjà',Ecran.Caption);
    GFAM.Cells[ACOD,ARow]:=''; Cancel:=True;
    end;
  if GFAM.Cells [ACOD,ARow] <> '' then GFAM2.Enabled := True;
  end;
if (ACol = ALIB) and (GFAM.Cells[AABR,ARow] = '') then
  begin
  GFAM.Cells[AABR,ARow]:= copy(GFAM.Cells[ALIB,ARow],1,17);
  end;
end;


procedure TOF_GCFamHier.GFAMRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
GFAM.InvalidateRow(ou) ;
if Ou >= GFAM.RowCount - 1 then GFAM.RowCount := GFAM.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TobPR.detail.count + 1);
if (ARow = TobPR.detail.count + 1) AND (not LigVide (ARow - 1, GFAM)) then
    BEGIN
    CreerTOBligne (ARow, GFAM);
    END;
if Ou > TobPR.detail.count then
    BEGIN
    GFAM.Row := TobPR.detail.count;
    END;
end;

procedure TOF_GCFamHier.GFAMRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if (GFAM.Row > 1) and (GFAM.Cells[ACOD,GFAM.Row -1] <> '') then
  BEGIN
  TraiterCODE (ACOD, GFAM.Row -1, GFAM);
  TraiterLIB (ALIB, GFAM.Row -1, GFAM);
  TraiterABR (AABR, GFAM.Row -1, GFAM);
  if TobPR.Detail[GFAM.Row-2].Modifie = True then
    begin
    ValideFamille;
    TobPR.SetAllModifie(False);
    end;
  END;
GFAM.InvalidateRow(ou) ;
if LigVide (Ou, GFAM) Then GFAM.Row := Min (GFAM.Row,Ou);
end;


procedure TOF_GCFamHier.DessineCellFam(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < GFAM.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GFAM.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GFAM.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GFAM.row) then
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
Procedure TOF_GCFamHier.ChercheFamille(GS : THGRID);
Var ARTICLE : THCritMaskEdit;
    Coord : TRect;
    xx_where : string;
    i_ind : integer;
BEGIN
xx_where := '';
{
for i_ind := 0 to  TobPR.detail.count-1 do
  begin
  if TobPR.Detail[i_ind].GetValue('CC_CODE') <> '' then
  begin
    if xx_where <> '' then xx_where := xx_where+' and';
    xx_where := xx_where+' CC_CODE <> "'+TobPR.Detail[i_ind].GetValue('CC_CODE')+'"';
    end;
  end;
}
for i_ind := 1 to  GS.RowCount-1 do
  begin
  if GS.Cells[ACOD,i_ind] <> '' then
  begin
    if xx_where <> '' then xx_where := xx_where+' and';
    xx_where := xx_where+' CC_CODE <> "'+GS.Cells[ACOD,i_ind]+'"';
    end;
  end;
Coord := GS.CellRect (GS.Col, GS.Row);
ARTICLE := THCritMaskEdit.Create (ECRAN);
ARTICLE.Parent := GS;
ARTICLE.Top := Coord.Top;
ARTICLE.Left := Coord.Left;
ARTICLE.Width := 3; ARTICLE.Visible := False;
ARTICLE.Text:= GS.Cells[GS.Col,GS.Row] ;
ARTICLE.DataType:='GCFAMHIER_MUL';
ARTICLE.Text := AGLLanceFiche('GC','GCFAMHIER_MUL','XX_WHERE='+xx_where,'',NvSaisi);
if ARTICLE.Text <> '' then GS.Cells[GS.Col,GS.Row]:= Trim (Copy (ARTICLE.Text, 1, 18));
ARTICLE.Destroy;
AfficheFAM(GS.CellValues [GS.Col,GS.Row],GS);
END ;

Procedure TOF_GCFamHier.AfficheFAM(CodeFam: string;GS : THGRID);
Var QQ : TQuery;
BEGIN
QQ:=OpenSQL('Select * from CHOIXCOD Where CC_TYPE= "'+NvSaisi+'" and CC_CODE="'+CodeFam+'"',True) ;
if Not QQ.EOF then
  begin
  GS.CellValues [ATYP,GS.Row] := QQ.Findfield('CC_TYPE').Asstring;
  GS.CellValues [ALIB,GS.Row] := QQ.Findfield('CC_LIBELLE').Asstring;
  GS.CellValues [AABR,GS.Row] := QQ.Findfield('CC_ABREGE').Asstring;
  GS.CellValues [ALBR,GS.Row] := QQ.Findfield('CC_LIBRE').Asstring;
  end;
Ferme (QQ);
END;


procedure TOF_GCFamHier.ChargeGrille ;
var QQ : TQuery ;
    St : string ;
BEGIN
St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "'+NvSaisi+'" AND CC_LIBRE like "%'+CodeFamSup+'%"';
QQ:=OpenSql(St, True) ;
TobPR:=tob.create('',Nil,-1) ;
if not QQ.Eof then TobPR.LoadDetailDB('CHOIXCOD','','',QQ,false,true) ;
Ferme(QQ) ;
If TobPR.detail.count=0 then begin Tob.create ('CHOIXCOD',TobPR,-1) ; end ;
TobPR.PutGridDetail(GFAM,True,True,LesColArt,True);
GFAM.RowCount:=Max (NbRowsInit, GFAM.RowCount+1) ;
TobToDelete:=tob.create('',Nil,-1) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GFAM) ;
GFAM.Col := ACOD;
GFAM.Row := 1;
GFAM.SetFocus;
END;

function TOF_GCFamHier.ExisteCode(Code : String; Ligne : Integer) : boolean;
var i : integer;
		QQ : Tquery;
begin
Result:=False;
if Code='' then exit;
For i:=1 to GFAM.RowCount-1 Do
  begin
  if (GFAM.Cells[ACOD,i]=Code) and (i<>Ligne) then
    begin Result:=True; Break; end;
  end;
  if NvSup = 'BO1' then
  begin
  	QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="BO2" AND CC_CODE="'+Code+'"',True);
  end else if NvSup = 'BP1' then
  begin
  	QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PA2" AND CC_CODE="'+Code+'"',True);
  end else
  begin
// verification de la non existance dans les autres familles
  	QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="FN2" AND CC_CODE="'+Code+'"',True);
  end;
  if not QQ.eof then result := true;
  ferme (QQ);
end;

Function TOF_GCFamHier.GetTOBLigne ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBPR.Detail.Count)) then Exit ;
Result:=TOBPR.Detail[ARow-1] ;
END ;

Procedure TOF_GCFamHier.InitRow (Row : integer; GS : THGrid) ;
var Col : integer ;
    TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(Row) ; if TOBL<>Nil then TOBL.InitValeurs ;
for Col:=0 to GS.ColCount do GS.cells[Col,Row]:='';
END ;

Procedure TOF_GCFamHier.CreerTOBLigne (ARow : integer; GS : THGrid);
BEGIN
if ARow <> TOBPR.Detail.Count + 1 then exit;
TOB.Create ('CHOIXCOD', TOBPR, ARow-1) ;
InitRow (ARow, GS) ;
END;

Function TOF_GCFamHier.LigVide(Row : integer; GS : THGRID) : Boolean ;
var Col : integer ;
BEGIN
Result:=True ;
Col:=ACOD ;
if (GS.Cells[Col,Row]<>'') then result:= False ;
END ;

procedure TOF_GCFamHier.InsertLigne (ARow : Longint; GS : THGrid) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigVide (ARow, GS) then exit;
if (ARow > TOBPR.Detail.Count) then Exit;
GS.CacheEdit; GS.SynEnabled := False;
TOB.Create ('CHOIXCOD', TOBPR, ARow-1) ;
GS.InsertRow (ARow); GS.Row := ARow;
InitRow (ARow, GS) ;
GS.MontreEdit; GS.SynEnabled := True;
END;

procedure TOF_GCFamHier.SupprimeLigne (ARow : Longint; GS : THGrid) ;
Var i_ind: integer;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBPR.Detail.Count) then Exit;
if (Arow=GS.rowcount) and (GS.Cells [ACOD,ARow] = '' ) then exit;
if (Arow < GS.rowcount) and (GS.Cells [ACOD,ARow] = '' ) and (GS.Cells [ACOD,ARow+1] = '' ) then exit;
GS.CacheEdit; GS.SynEnabled := False;
GS.DeleteRow (ARow);
if (ARow = TOBPR.Detail.Count) then CreerTOBLigne (ARow + 1, GS);
if TOBPR.Detail[ARow-1].GetValue ('CC_CODE') <> '' then
    BEGIN
    i_ind := TobToDelete.Detail.Count;
    TOB.Create ('CHOIXCOD', TobToDelete, i_ind) ;
    TobToDelete.Detail[i_ind].Dupliquer (TOBPR.Detail[ARow-1], False, True);
    END;
TOBPR.Detail[ARow-1].Free;
if GS.RowCount < NbRowsInit then GS.RowCount := NbRowsInit;
GS.MontreEdit; GS.SynEnabled := True;
//
//
END;

{==============================================================================================}
{============================ Manipulation des Champs GRID ====================================}
{==============================================================================================}

procedure TOF_GCFamHier.TraiterLIB (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := GS.CellValues [ACol, ARow];
if (St <> '') and (TOBL.GetValue('CC_LIBELLE') <> St) then
  TOBL.PutValue('CC_LIBELLE', St);
END;


procedure TOF_GCFamHier.TraiterABR (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    st : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := GS.CellValues [ACol, ARow];
if TOBL.GetValue ('CC_ABREGE') <>  st then TOBL.PutValue ('CC_ABREGE', st);
END;

procedure TOF_GCFamHier.TraiterCODE (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St, tmp, str : string;
    trouve : boolean;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := UpperCase(GS.Cells [ACol, ARow]);
GS.Cells [ACol, ARow] := St;
if St=''then exit ;
if TOBL.GetValue ('CC_CODE') <> st then
  begin
  TOBL.PutValue ('CC_TYPE', NvSaisi);
  TOBL.PutValue ('CC_CODE', st);
  end;
str := GS.Cells [ALBR, ARow];
trouve := False;
while str <> '' do
  begin
  tmp := ReadTokenSt(str) ;
  if tmp = CodeFamSup then trouve := True;
  end;
if trouve = False then
  begin
  if GS.Cells [ALBR, ARow] = '' then tmp := CodeFamSup
  else tmp := GS.Cells [ALBR, ARow]+';'+CodeFamSup;
  if TOBL.GetValue ('CC_LIBRE') <>  tmp then
    begin
    TOBL.PutValue ('CC_LIBRE', tmp);
    GS.Cells [ALBR, ARow] := tmp;
    end;
  end;
END;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TOF_GCFamHier.BChercherClick(Sender: TObject);
begin
if GFAM.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TOF_GCFamHier.BValiderClick(Sender: TObject);
begin
  TraiterCODE (ACOD, GFAM.Row, GFAM);
  TraiterLIB (ALIB, GFAM.Row , GFAM);
  TraiterABR (AABR, GFAM.Row, GFAM);
  if TobPR.Detail[GFAM.Row-1].Modifie = True then
    begin
    ValideFamille;
    TobPR.SetAllModifie(False);
    end;
  NV3TraiterCODE (ACOD2, GFAM2.Row, GFAM2);
  NV3TraiterLIB (ALIB2, GFAM2.Row, GFAM2);
  NV3TraiterABR (AABR2, GFAM2.Row, GFAM2);
  if TobPR2.Detail[GFAM2.Row-1].Modifie = True then
    begin
    ValideFamille2;
    TobPR2.SetAllModifie(False);
    end;
end;

procedure TOF_GCFamHier.FindLigneFind(Sender: TObject);
begin
if(Screen.ActiveControl=GFAM) then Rechercher (GFAM, FindLigne, FindDebut) ;
if(Screen.ActiveControl=GFAM2) then Rechercher (GFAM2, FindLigne, FindDebut) ;
end;

procedure TOF_GCFamHier.BNewLineClick(Sender: TObject);
BEGIN
if(Screen.ActiveControl=GFAM) then  InsertLigne (GFAM.Row, GFAM);
if(Screen.ActiveControl=GFAM2) then  InsertLigne2 (GFAM2.Row, GFAM2);
end;

procedure TOF_GCFamHier.BDelLineClick(Sender: TObject);
var tmp : string ;
begin
if(Screen.ActiveControl = GFAM) then
  begin
  DellesLigne2;
  MajLesFamillesToDelete2;
  SupprimeLigne (GFAM.Row, GFAM);
  MajLesFamillesToDelete;
  if GFAM.Cells [ACOD,GFAM.Row] = '' then
    begin
    tmp := CodeFamSup+'***';
    end else
    begin
    tmp := CodeFamSup+GFAM.Cells [ACOD,GFAM.Row];
    end;
    ChargeGrille2 (tmp);
  end else
      begin
      SupprimeLigne2 (GFAM2.Row, GFAM2);
      MajLesFamillesToDelete2;
    end;
end;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}
procedure TOF_GCFamHier.ValideFamille;
begin
//MajLesFamillesToDelete;
TobPR.InsertOrUpdateDB();
valide_Nv2 := 1;
end;

procedure TOF_GCFamHier.MajLesFamillesToDelete;
var i_ind : integer;
    stlibre,St, Nvlibre : string;
begin
  IF TobToDelete.detail.count > 0 THEN
  BEGIN
  for i_ind := 0 to TobToDelete.detail.count -1 do
    begin
      StLibre :=  TobToDelete.Detail[i_ind].GetValue('CC_LIBRE') ;
      St:=(ReadTokenSt(StLibre));
      While (St <>'') do
        begin
        if st <> CodeFamSup then
          begin
          if NvLibre <> '' then NvLibre := NvLibre+';';
          NvLibre := NvLibre+St;
          end;
        St:=(ReadTokenSt(StLibre));
        end;
        TobToDelete.Detail[i_ind].PutValue('CC_LIBRE',NvLibre) ;
    end;
  TobToDelete.InsertOrUpdateDB();
  END;
end;

// Gestion de la grille des familles niveau 3
procedure TOF_GCFamHier.FormKeyDown2(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
    GS : THGrid ;
BEGIN
FocusGrid := True;
GS := GFAM2;
ARow := GFAM2.Row;
Case Key of
    VK_F5 : if FocusGrid then GS.OnElipsisClick(Sender);
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if FocusGrid then
                    BEGIN
                    Key := 0;
                    InsertLigne2 (ARow, GS);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne2 (ARow, GS) ;
                    END ;
                END;
    END;
END;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}
procedure TOF_GCFamHier.GFAM2ElipsisClick(Sender: TObject);
begin
if GFAM2.Col = ACOD2 then
    BEGIN
    ChercheFamille2 (GFAM2);
    END ;
end;


procedure TOF_GCFamHier.GFAM2CellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if GFAM2.Cells [ACOD2,GFAM2.Row] = '' then valide_Nv3 := 0;
ZoneSuivanteOuOk (GFam2,Acol,Arow,Cancel);
if cancel then exit;
if Not Cancel then
    BEGIN
//    if (GFAM2.Col <> ACOD2) AND (GFAM2.Cells [ACOD2,GFAM2.Row] = '') then BEGIN GFAM2.Col := ACOD2; END;
    GFAM2.ElipsisButton:=(GFAM2.Col=ACOD2);
    end ;
end;

procedure TOF_GCFamHier.GFAM2CellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if ACol = ACOD2 then
  begin
  if ExisteCode2(GFAM2.Cells[ACOD2,ARow],ARow) then
    begin
    PGIInfo('Ce code existe déjà',Ecran.Caption);
    GFAM2.Cells[ACOD2,ARow]:=''; Cancel:=True;
    end;
  end;
if (ACol = ALIB2) and (GFAM2.Cells[AABR2,ARow] = '') then
  begin
  GFAM2.Cells[AABR2,ARow]:= copy(GFAM2.Cells[ALIB2,ARow],1,17);
  end;
end;


procedure TOF_GCFamHier.GFAM2RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
GFAM2.InvalidateRow(ou) ;
if Ou >= GFAM2.RowCount - 1 then GFAM2.RowCount := GFAM2.RowCount + NbRowsPlus2 ;;
ARow := Min (Ou, TobPR2.detail.count + 1);
if (ARow = TobPR2.detail.count + 1) AND (not LigVide2 (ARow - 1, GFAM2)) then
    BEGIN
    CreerTOBLigne2 (ARow, GFAM2);
    END;
if Ou > TobPR2.detail.count then
    BEGIN
    GFAM2.Row := TobPR2.detail.count;
    END;
end;

procedure TOF_GCFamHier.GFAM2RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if (GFAM2.Row-1 = 1) and (valide_Nv2 = 0) then
   begin
   TraiterCODE (ACOD, GFAM.Row, GFAM);
   TraiterLIB (ALIB, GFAM.Row , GFAM);
   TraiterABR (AABR, GFAM.Row , GFAM);
   if TobPR.Detail[GFAM.Row-1].Modifie = True then
     begin
     ValideFamille;
     TobPR.SetAllModifie(False);
     end;
   end;

if (GFAM2.Row > 1) and (GFAM2.Cells[ACOD2,GFAM2.Row -1] <> '') then
  BEGIN
  NV3TraiterCODE (ACOD2, GFAM2.Row -1, GFAM2);
  NV3TraiterLIB (ALIB2, GFAM2.Row -1, GFAM2);
  NV3TraiterABR (AABR2, GFAM2.Row -1, GFAM2);
  if TobPR2.Detail[GFAM2.Row-2].Modifie = True then
    begin
    ValideFamille2;
    TobPR2.SetAllModifie(False);
    end;
  END;
GFAM2.InvalidateRow(ou) ;
if LigVide2 (Ou, GFAM2) Then GFAM2.Row := Min (GFAM2.Row,Ou);
end;


procedure TOF_GCFamHier.DessineCellFam2(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < GFAM2.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GFAM2.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GFAM2.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GFAM2.row) then
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
Procedure TOF_GCFamHier.ChercheFamille2(GS : THGRID);
Var ARTICLE : THCritMaskEdit;
    Coord : TRect;
    xx_where : string;
    i_ind : integer;
BEGIN
xx_where := '';
{for i_ind := 0 to  TobPR2.detail.count-1 do
  begin
  if TobPR2.Detail[i_ind].GetValue('CC_CODE') <> '' then
  begin
    if xx_where <> '' then xx_where := xx_where+' and';
    xx_where := xx_where+' CC_CODE <> "'+TobPR2.Detail[i_ind].GetValue('CC_CODE')+'"';
    end;
  end;}
for i_ind := 1 to  GS.RowCount-1 do
  begin
  if GS.Cells[ACOD2,i_ind] <> '' then
  begin
    if xx_where <> '' then xx_where := xx_where+' and';
    xx_where := xx_where+' CC_CODE <> "'+GS.Cells[ACOD,i_ind]+'"';
    end;
  end;

Coord := GS.CellRect (GS.Col, GS.Row);
ARTICLE := THCritMaskEdit.Create (ECRAN);
ARTICLE.Parent := GS;
ARTICLE.Top := Coord.Top;
ARTICLE.Left := Coord.Left;
ARTICLE.Width := 3; ARTICLE.Visible := False;
ARTICLE.Text:= GS.Cells[GS.Col,GS.Row] ;
ARTICLE.Datatype:='GCFAMHIER_MUL';
ARTICLE.Text := AGLLanceFiche('GC','GCFAMHIER_MUL','XX_WHERE='+xx_where,'',NvSaisi2);
if ARTICLE.Text <> '' then GS.Cells[GS.Col,GS.Row]:= Trim (Copy (ARTICLE.Text, 1, 18));
ARTICLE.Destroy;
AfficheFAM2(GS.CellValues [GS.Col,GS.Row],GS);
END ;

Procedure TOF_GCFamHier.AfficheFAM2(CodeFam: string;GS : THGRID);
Var QQ : TQuery;
BEGIN
QQ:=OpenSQL('Select * from CHOIXCOD Where CC_TYPE= "'+NvSaisi2+'" and CC_CODE="'+CodeFam+'"',True) ;
if Not QQ.EOF then
  begin
  GS.CellValues [ATYP2,GS.Row] := QQ.Findfield('CC_TYPE').Asstring;
  GS.CellValues [ALIB2,GS.Row] := QQ.Findfield('CC_LIBELLE').Asstring;
  GS.CellValues [AABR2,GS.Row] := QQ.Findfield('CC_ABREGE').Asstring;
  GS.CellValues [ALBR2,GS.Row] := QQ.Findfield('CC_LIBRE').Asstring;
  end;
Ferme (QQ);
END;


procedure TOF_GCFamHier.ChargeGrille2 (CodNv2 : string);
var QQ : TQuery ;
    St : string ;
    Acol,Arow : integer;
    Cancel  : boolean;
BEGIN
St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "'+NvSaisi2+'" AND CC_LIBRE like "%'+CodNv2+'%"';
QQ:=OpenSql(St, True) ;
TobPR2:=tob.create('',Nil,-1) ;
if not QQ.Eof then TobPR2.LoadDetailDB('CHOIXCOD','','',QQ,false,true) ;
Ferme(QQ) ;
If TobPR2.detail.count=0 then begin Tob.create ('CHOIXCOD',TobPR2,-1) ; end ;
TobPR2.PutGridDetail(GFAM2,True,True,LesColArt2,True);
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GFAM2) ;
GFAM2.RowCount:=Max (NbRowsInit2, GFAM2.RowCount+1) ;
TobToDelete2:=tob.create('',Nil,-1) ;
GFAM2.CacheEdit; GFAM2.SynEnabled := False;
Acol := ACOD2;
Arow := GFAM2.Row;
ZoneSuivanteOuOk (GFAM2, Acol,ARow,cancel);
GFAM2.Row := ARow;
GFAM2.Col := ACol;
GFAM2.MontreEdit; GFAM2.SynEnabled := True;

END;

function TOF_GCFamHier.ExisteCode2(Code : String; Ligne : Integer) : boolean;
var i : integer;
		QQ : TQuery;
begin
Result:=False;
if Code='' then exit;
For i:=1 to GFAM2.RowCount-1 Do
  begin
  if (GFAM2.Cells[ACOD2,i]=Code) and (i<>Ligne) then
    begin Result:=True; Break; end;
  end;
  if NvSup = 'BO1' then
  begin
		QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="BO3" AND CC_CODE="'+Code+'"',True);
  end else if NvSup = 'BP1' then
  begin
		QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PA3" AND CC_CODE="'+Code+'"',True);
  end else
  begin
		QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="FN3" AND CC_CODE="'+Code+'"',True);
  end;
if not QQ.eof then result := true;
ferme (QQ); 
end;

Function TOF_GCFamHier.GetTOBLigne2 ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TobPR2.Detail.Count)) then Exit ;
Result:=TobPR2.Detail[ARow-1] ;
END ;

Procedure TOF_GCFamHier.InitRow2 (Row : integer; GS : THGrid) ;
var Col : integer ;
    TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne2(Row) ; if TOBL<>Nil then TOBL.InitValeurs ;
for Col:=0 to GS.ColCount do GS.cells[Col,Row]:='';
END ;

Procedure TOF_GCFamHier.CreerTOBLigne2 (ARow : integer; GS : THGrid);
BEGIN
if ARow <> TobPR2.Detail.Count + 1 then exit;
TOB.Create ('CHOIXCOD', TobPR2, ARow-1) ;
InitRow2 (ARow, GS) ;
END;

Function TOF_GCFamHier.LigVide2(Row : integer; GS : THGRID) : Boolean ;
var Col : integer ;
BEGIN
Result:=True ;
Col:=ACOD2 ;
if (GS.Cells[Col,Row]<>'') then result:= False ;
END ;

procedure TOF_GCFamHier.InsertLigne2 (ARow : Longint; GS : THGrid) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigVide2 (ARow, GS) then exit;
if (ARow > TobPR2.Detail.Count) then Exit;
GS.CacheEdit; GS.SynEnabled := False;
TOB.Create ('CHOIXCOD', TobPR2, ARow-1) ;
GS.InsertRow (ARow); GS.Row := ARow;
InitRow2 (ARow, GS) ;
GS.MontreEdit; GS.SynEnabled := True;
END;

procedure TOF_GCFamHier.DellesLigne2;
Var Indice,i_ind: integer;
begin
Indice := 0;
if TOBPR2.detail.count = 0 then exit;
repeat
if TobPR2.Detail[Indice].GetValue ('CC_CODE') <> '' then
   BEGIN
   i_ind := TobToDelete2.Detail.Count;
   TOB.Create ('CHOIXCOD', TobToDelete2, I_ind) ;
   TobToDelete2.Detail[i_ind].Dupliquer (TobPR2.Detail[Indice], False, True);
   END;
   TobPR2.Detail[indice].Free;
until TOBPr2.detail.count = 0;
end;

procedure TOF_GCFamHier.SupprimeLigne2 (ARow : Longint; GS : THGrid) ;
Var i_ind: integer;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TobPR2.Detail.Count) then Exit;
if (Arow=GS.rowcount) and (GS.Cells [ACOD,ARow] = '' ) then exit;
if (Arow < GS.rowcount) and (GS.Cells [ACOD,ARow] = '' ) and (GS.Cells [ACOD,ARow+1] = '' ) then exit;
GS.CacheEdit; GS.SynEnabled := False;
GS.DeleteRow (ARow);
if (ARow = TobPR2.Detail.Count) then CreerTOBLigne2 (ARow + 1, GS);
if TobPR2.Detail[ARow-1].GetValue ('CC_CODE') <> '' then
    BEGIN
    i_ind := TobToDelete2.Detail.Count;
    TOB.Create ('CHOIXCOD', TobToDelete2, i_ind) ;
    TobToDelete2.Detail[i_ind].Dupliquer (TobPR2.Detail[ARow-1], False, True);
    END;
TobPR2.Detail[ARow-1].Free;
if GS.RowCount < NbRowsInit2 then GS.RowCount := NbRowsInit2;
GS.MontreEdit; GS.SynEnabled := True;
END;

{==============================================================================================}
{============================ Manipulation des Champs GRID ====================================}
{==============================================================================================}

procedure TOF_GCFamHier.NV3TraiterLIB (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne2 (ARow); if TOBL=nil then exit;
St := GS.CellValues [ACol, ARow];
if (St <> '') and (TOBL.GetValue('CC_LIBELLE') <> St) then
  TOBL.PutValue('CC_LIBELLE', St);
END;


procedure TOF_GCFamHier.NV3TraiterABR (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    st : string;
BEGIN
TOBL := GetTOBLigne2 (ARow); if TOBL=nil then exit;
St := GS.CellValues [ACol, ARow];
if TOBL.GetValue ('CC_ABREGE') <>  st then TOBL.PutValue ('CC_ABREGE', st);
END;

procedure TOF_GCFamHier.NV3TraiterCODE (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St, tmp, str : string;
    trouve : boolean;
BEGIN

  TOBL := GetTOBLigne2 (ARow); if TOBL=nil then exit;

  St := UpperCase(GS.Cells [ACol, ARow]);

  GS.Cells [ACol, ARow] := St;

  if St=''then exit ;

  if TOBL.GetValue ('CC_CODE') <>  st then
  begin
    TOBL.PutValue ('CC_TYPE', NvSaisi2);
    TOBL.PutValue ('CC_CODE', st);
  end;

  str := GS.Cells [ALBR2, ARow];
  st := CodeFamSup+GFAM.Cells [ACOD,GFAM.Row];
  trouve := False;

  while str <> '' do
  begin
    tmp := ReadTokenSt(str) ;
    if tmp = st then trouve := True;
  end;

  if trouve = False then
  begin
    if GS.Cells [ALBR2, ARow] = '' then
      tmp := st
    else
      tmp := GS.Cells [ALBR2, ARow]+';'+st;
    if TOBL.GetValue ('CC_LIBRE') <>  tmp then
    begin
      TOBL.PutValue ('CC_LIBRE', tmp);
      GS.Cells [ALBR2, ARow] := tmp;
    end;
  end;
END;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}
procedure TOF_GCFamHier.ValideFamille2;
begin
MajLesFamillesToDelete2;
TobPR2.InsertOrUpdateDB();
valide_Nv3 := 1;
end;

procedure TOF_GCFamHier.MajLesFamillesToDelete2;
var i_ind : integer;
    stlibre,St, Nvlibre, tmp : string;
begin
IF TobToDelete2.detail.count > 0 THEN
  BEGIN
  for i_ind := 0 to TobToDelete2.detail.count -1 do
    begin
      StLibre :=  TobToDelete2.Detail[i_ind].GetValue('CC_LIBRE') ;
      St:=(ReadTokenSt(StLibre));
      While (St <>'') do
        begin
        tmp := CodeFamSup+GFAM.Cells [ACOD,GFAM.Row];
        if st <> tmp then
          begin
          if NvLibre <> '' then NvLibre := NvLibre+';';
          NvLibre := NvLibre+St;
          end;
        St:=(ReadTokenSt(StLibre));
        end;
        TobToDelete2.Detail[i_ind].PutValue('CC_LIBRE',NvLibre) ;
    end;
  TobToDelete2.InsertOrUpdateDB();
  END;
end;

procedure AGLOnChangeNvSup (Parms: array of variant; nb: integer) ;
var F : TForm ;
    TOTOF : TOF ;
    tmp : string;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_GCFamHier) then
  begin
  TOF_GCFamHier(TOTOF).CodeFamSup:=Parms[1];
  TOF_GCFamHier(TOTOF).ChargeGrille;
  if TOF_GCFamHier(TOTOF).GFAM.Cells [2,1] = '' then
    begin
    tmp := TOF_GCFamHier(TOTOF).CodeFamSup+'***';
    TOF_GCFamHier(TOTOF).GFAM2.Enabled := False;
  end  else
    begin
    tmp := TOF_GCFamHier(TOTOF).CodeFamSup+TOF_GCFamHier(TOTOF).GFAM.Cells [2,1];
    TOF_GCFamHier(TOTOF).GFAM2.Enabled := True;
    TOF_GCFamHier(TOTOF).GFAM.Col := 3;
    end;
  TOF_GCFamHier(TOTOF).ChargeGrille2 (tmp);
  end;
end ;


Function TOF_GCFamHier.ZoneAccessible ( GS: THGrid; ACol,ARow : Longint) : boolean ;
BEGIN
Result:=True ;
if GS.ColLengths[ACol]<0 then BEGIN Result:=False ; Exit ; END ;
if (ACol = ACOD) AND (GS.Cells [ACOD,ARow] <> '') then BEGIN result := false; Exit; END;
END ;

procedure TOF_GCFamHier.ZoneSuivanteOuOk ( GS:THGrid; Var ACol,ARow : Longint ; Var Cancel : boolean ) ;
Var Sens,ii,Lim : integer ;
    OldEna,ChgLig,ChgSens  : boolean ;
BEGIN
OldEna:=GS.SynEnabled ; GS.SynEnabled:=False ;
Sens:=-1 ; ChgLig:=(GS.Row<>ARow) ; ChgSens:=False ;
if GS.Row>ARow then Sens:=1 else if ((GS.Row=ARow) and (ACol<GS.Col)) then Sens:=1 ;
ACol:=GS.Col ; ARow:=GS.Row ; ii:=0 ;
While Not ZoneAccessible(GS,ACol,ARow)  do
   BEGIN
   Cancel:=True ; inc(ii) ; if ii>500 then Break ;
   if Sens=1 then
      BEGIN
      Lim:=GS.RowCount-1 ;
      if ((ACol=GS.ColCount-1) and (ARow>=Lim)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=-1 ; Continue ; ChgSens:=True ; END ;
         END ;
      if ChgLig then BEGIN ACol:=GS.FixedCols-1 ; ChgLig:=False ; END ;
      if ACol<GS.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=GS.FixedCols ; END ;
      END else
      BEGIN
      if ((ACol=GS.FixedCols) and (ARow=1)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=1 ; Continue ; END ;
         END ;
      if ChgLig then BEGIN ACol:=GS.ColCount ; ChgLig:=False ; END ;
      if ACol>GS.FixedCols then Dec(ACol) else BEGIN Dec(ARow) ; ACol:=GS.ColCount-1 ; END ;
      END ;
   END ;
GS.SynEnabled:=OldEna ;
END ;

Initialization
registerclasses([TOF_GCFamHier]);
RegisterAglProc('OnChangeNvSup', True , 1,AGLOnChangeNvSup ) ;
end.

