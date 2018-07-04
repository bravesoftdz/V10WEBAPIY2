{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCCONTREAVT ()
Mots clefs ... : TOF;GCCONTREAVT
*****************************************************************}
Unit UTOFGCCONTREAVT ;

Interface

Uses StdCtrls, extctrls, Controls, Classes, forms,
     windows, sysutils, ComCtrls, Graphics, ParamDBG, EntGC,
     UIUtil, HCtrls, HPanel, HEnt1, HMsgBox, HSysMenu, HTB97, UTOF, UTOB,  HDB,
{$IFDEF EAGLCLIENT}
     eFiche, eMul,
{$ELSE}
     DBTables, db, Mul,DBGrids,
{$ENDIF}
     Grids, HQry, LicUtil, FactUtil, Facture, math,AGLInit,
     SaisUtil,ActionContreM_Tof;

Type
  TOF_GCCONTREAVT = Class (TOF)
    G_General, G_Detail : THGrid;
    PCumul : TPanel;
//    PQTE, PCLINUM, PREFART : THPanel;
    PQTE : THPanel;
    Split : TSplitter;
    BCherche, BParamListe, BINVERSE : TToolbarButton97;
    GQC_REFERENCE, GQC_FOURNISSEUR, GQC_ARTICLE, GQC_CLIENT : THEdit;
    GP_TIERS, GP_NUMERO, GL_DATEPIECE, GL_DATEPIECE_ : THEdit;
    GQC_DEPOT : THValComboBox;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
  private
    LesColsGen, LesColsDet,LesChampsDispo,LesChampsLig : string;
    TobCont : TOB;
    NbLignesGridDown : integer;
    procedure GridRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GridRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BChercheClick(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure BINVERSEClick(Sender: TObject);
    procedure G_GeneralDblClick(Sender: TObject);
    procedure G_DetailDblClick(Sender: TObject);
    procedure EtudieColsListes;
    procedure ChargeMouvements;
// Initialisation
    procedure LoadLesTobLigneREFART;
    function  InitialiseEntete(G_SEL : THGrid) : string;
    procedure InitialiseGrille;

    function  RecupeCritereAvance : string;
    //Function  Evaluedate (St : String) : TDateTime;
    Procedure InverseTobs;
  end ;

var
   bGeneralOnTop : boolean;

Implementation

procedure TOF_GCCONTREAVT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCCONTREAVT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCCONTREAVT.OnUpdate ;
begin
  Inherited ;
//  ChargeMouvements;
end ;

procedure TOF_GCCONTREAVT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCCONTREAVT.OnArgument (S : String ) ;
begin
  Inherited ;
  bGeneralOnTop := True;
  G_General := THGrid(GetControl('G_GENERAL'));
  G_General.OnRowEnter := GridRowEnter;
  G_General.OnRowExit:=GridRowExit;
  G_General.OnDblClick:=G_GeneralDblClick;
  G_General.GetCellCanvas := DessineCell;;
  G_Detail := THGrid(GetControl('G_DETAIL'));
  G_Detail.OnDblClick:=G_DetailDblClick;
  G_Detail.GetCellCanvas := DessineCell;;
  PCumul := TPanel(GetControl('PCumul'));
  PQTE := THPanel(GetControl('PQTE'));
  SetControlVisible('PCritere', True);
  SetControlVisible('PComplement', True);
  BCherche := TToolbarButton97(GetControl('BCherche'));
  BCherche.OnClick := BChercheClick;
  BParamListe := TToolbarButton97(GetControl('BParamListe'));
  BParamListe.OnClick := BParamListeClick;
  BParamListe.Visible := (V_PGI.PassWord = CryptageSt(DayPass(Date))) ;
  BINVERSE := TToolbarButton97(GetControl('BINVERSE'));
  BINVERSE.OnClick := BINVERSEClick;
  GQC_REFERENCE := THEdit(GetControl('GQC_REFERENCE'));
  GQC_REFERENCE.ElipsisButton:=False;
  GQC_FOURNISSEUR := THEdit(GetControl('GQC_FOURNISSEUR'));
  GQC_ARTICLE := THEdit(GetControl('GQC_ARTICLE'));
  GQC_CLIENT := THEdit(GetControl('GQC_CLIENT'));
  GP_TIERS := THEdit(GetControl('GP_TIERS'));
  GP_NUMERO := THEdit(GetControl('GP_NUMERO'));
  GL_DATEPIECE := THEdit(GetControl('GL_DATEPIECE'));
  GL_DATEPIECE_ := THEdit(GetControl('GL_DATEPIECE_'));
  GQC_DEPOT := THValComboBox(GetControl('GQC_DEPOT'));

  SetControlVisible('Fliste',False);
  InitialiseGrille;

  PQTE.Visible := True;
  Split := TSplitter.Create(Ecran);
  Split.ResizeStyle:=rsUpdate;
  Split.Visible := True;
  Split.Beveled := True;
  Split.Height := 10;
  Split.AutoSnap := False;

  G_General.Parent := PQTE;
  G_Detail.Parent := PQTE;
  Split.Parent := PQTE;
  G_General.Align := alTop;
  Split.Align := alTop;
  G_Detail.Align := alClient;
  G_General.Height := Trunc((PQTE.Height - Split.Height) / 2);

  Split.Cursor := crVSplit;

  G_General.ListeParam := 'GCMULCONTREAVTGT';
  G_Detail.ListeParam := 'GCMULCONTREAVTLD';
  EtudieColsListes;
  AffecteGrid(G_General,taConsult) ;
  TFMul(Ecran).HMTrad.ResizeGridColumns(G_General);
  AffecteGrid(G_Detail,taConsult) ;
  TFMul(Ecran).HMTrad.ResizeGridColumns(G_Detail);

  TobCont := TOB.Create('Avct', nil, -1);
end ;

procedure TOF_GCCONTREAVT.OnClose ;
begin
  Inherited ;
  if TobCont<>Nil then begin TobCont.Free; TobCont:=Nil; end;
end ;

procedure TOF_GCCONTREAVT.GridRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
  if bGeneralOnTop then
      begin
      G_General.Cells [0, Ou] := '„';
      TOB(G_General.Objects[0, Ou]).PutGridDetail(G_Detail, False, False, LesColsDet);
      end
      else
      begin
      G_Detail.Cells [0, Ou] := '„';
      TOB(G_Detail.Objects[0, Ou]).PutGridDetail(G_General, False, False, LesColsGen);
      end;
end;

procedure TOF_GCCONTREAVT.GridRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
  if bGeneralOnTop then
      begin
      G_General.Cells [0, Ou] := ' ';
      G_Detail.VidePile(False);
      end
      else
      begin
      G_Detail.Cells [0, Ou] := ' ';
      G_General.VidePile(False);
      end;
end;

procedure TOF_GCCONTREAVT.BChercheClick(Sender: TObject);
var ARow : integer;
    bc : boolean;
    GS : THGrid;
begin
  inherited;
ChargeMouvements;
if bGeneralOnTop then GS:=G_General else GS:=G_Detail;
if GS.Row <=GS.RowCount then ARow:=GS.Row else ARow:=1;
if G_General.Enabled then GridRowEnter(GS,ARow,bc,False) ;
end;

procedure TOF_GCCONTREAVT.BParamListeClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
if bGeneralOnTop then
    begin
    if G_Detail.Focused then ParamListe(G_Detail.ListeParam,nil)
                        else ParamListe(G_General.ListeParam,nil);
    end else
    begin
    if G_General.Focused then ParamListe(G_General.ListeParam,nil)
                         else ParamListe(G_Detail.ListeParam,nil);
    Split.Align := alNone;
    bGeneralOnTop := not (bGeneralOnTop);
    end;
{$ELSE}
if bGeneralOnTop then
    begin
    if G_Detail.Focused then ParamListe(G_Detail.ListeParam,nil,nil)
                        else ParamListe(G_General.ListeParam,nil,nil);
    end else
    begin
    if G_General.Focused then ParamListe(G_General.ListeParam,nil,nil)
                         else ParamListe(G_Detail.ListeParam,nil,nil);
    Split.Align := alNone;
    bGeneralOnTop := not (bGeneralOnTop);
    end;
{$ENDIF}
InitialiseGrille;
G_General.ListeParam := 'GCMULCONTREAVTGT';
G_Detail.ListeParam := 'GCMULCONTREAVTLD';
EtudieColsListes;
ChargeMouvements;
G_General.OnRowEnter := GridRowEnter;
G_General.OnRowExit := GridRowExit;
G_Detail.OnRowEnter := Nil;
G_Detail.OnRowExit := Nil;
G_General.Row := 1;
{G_General.Align := alNone;
Split.Align := alNone;
G_Detail.Align := alNone;
G_General.Align := alTop;
//G_General.Height := Max(Trunc(PQTE.Height/2),(PQTE.Height-(G_Detail.DefaultRowHeight*(NbLignesGridDown+1))-12));
//G_General.Height := Min(Trunc(PQTE.Height/2),(G_General.DefaultRowHeight*(G_General.RowCount+2)+12));
Split.Align := alTop;
G_Detail.Align := alClient;  }
TFMul(Ecran).HMTrad.ResizeGridColumns(G_General);
TFMul(Ecran).HMTrad.ResizeGridColumns(G_Detail);
end;

procedure TOF_GCCONTREAVT.BINVERSEClick(Sender: TObject);
begin
inherited;
G_General.Cols[0].Clear;
G_Detail.Cols[0].Clear;
InverseTobs;
if not (bGeneralOnTop) then
    begin
    G_General.Align := alNone;
    Split.Align := alNone;
    G_Detail.Align := alNone;
    G_General.Align := alTop;
    //G_General.Height := Max(Trunc(PQTE.Height/2),(PQTE.Height-(G_Detail.DefaultRowHeight*(NbLignesGridDown+1))-12));
    G_General.Height := Min(Trunc(PQTE.Height/2),(G_General.DefaultRowHeight*(TobCont.Detail.Count+3)+12));
    Split.Align := alTop;
    G_Detail.Align := alClient;
    G_General.OnRowEnter := GridRowEnter;
    G_General.OnRowExit := GridRowExit;
    G_Detail.OnRowEnter := Nil;
    G_Detail.OnRowExit := Nil;
    G_General.ListeParam := 'GCMULCONTREAVTGT';
    G_Detail.ListeParam := 'GCMULCONTREAVTLD';
    EtudieColsListes;
    if TobCont.Detail.Count > 0 then
        begin
        G_General.RowCount := TobCont.Detail.Count + 1;
        TobCont.PutGridDetail (G_General, False, False, LesColsGen, False);
        G_General.SetFocus;
        G_General.Cells [0, 1] := '„';
        G_General.Cells [0, 0] := '';
        G_Detail.RowCount := TOB(G_General.Objects[0, 1]).Detail.Count + 1;
        TOB(G_General.Objects[0, 1]).PutGridDetail(G_Detail, False, False, LesColsDet);
        end;
    G_General.Row := 1;
    end
    else
    begin
    G_Detail.Align := alNone;
    Split.Align := alNone;
    G_General.Align := alNone;
    G_Detail.Align := alTop;
    //G_Detail.Height := Max(Trunc(PQTE.Height/2),(PQTE.Height-(G_General.DefaultRowHeight*(NbLignesGridDown+1))-12));
    G_Detail.Height := Min(Trunc(PQTE.Height/2),(G_Detail.DefaultRowHeight*(TobCont.Detail.Count+3)+12));
    Split.Align := alTop;
    G_General.Align := alClient;
    G_General.OnRowEnter := Nil;
    G_General.OnRowExit := Nil;
    G_Detail.OnRowEnter := GridRowEnter;
    G_Detail.OnRowExit := GridRowExit;
    G_General.ListeParam := 'GCMULCONTREAVTGD';
    G_Detail.ListeParam := 'GCMULCONTREAVTLT';
    EtudieColsListes;
    if TobCont.Detail.Count > 0 then
        begin
        G_Detail.RowCount := TobCont.Detail.Count + 1;
        TobCont.PutGridDetail (G_Detail, False, False, LesColsDet, False);
        G_Detail.SetFocus;
        G_Detail.Cells [0, 1] := '„';
        G_Detail.Cells [0, 0] := '';
        G_General.RowCount := TOB(G_Detail.Objects[0, 1]).Detail.Count + 1;
        TOB(G_Detail.Objects[0, 1]).PutGridDetail(G_General, False, False, LesColsGen);
        end;
    G_Detail.Row := 1;
    end;
TFMul(Ecran).HMTrad.ResizeGridColumns(G_General);
TFMul(Ecran).HMTrad.ResizeGridColumns(G_Detail);
bGeneralOnTop := not (bGeneralOnTop);
end;

procedure TOF_GCCONTREAVT.G_GeneralDblClick(Sender: TObject);
begin
end;

procedure TOF_GCCONTREAVT.G_DetailDblClick(Sender: TObject);
var TOBL : TOB;
    CleDoc : R_CleDoc;
    Pref,Retour : string;
BEGIN
TOBL := TOB(G_Detail.Objects[0, G_Detail.Row]); if TOBL=nil then exit;
Pref := 'GL_';
CleDoc.NaturePiece:=Uppercase(TOBL.GetValue(Pref + 'NATUREPIECEG'));
CleDoc.Souche:=Uppercase(TOBL.GetValue(Pref + 'SOUCHE'));
CleDoc.DatePiece:=TOBL.GetValue(Pref + 'DATEPIECE');
CleDoc.NumeroPiece:=TOBL.GetValue(Pref + 'NUMERO');
CleDoc.Indice:=TOBL.GetValue(Pref + 'INDICEG');
Retour:=GCLanceFiche_ActionContreM('GC','GCACTIONCONTREM','','',CleDoc.NaturePiece);
if Retour='CON' then SaisiePiece(CleDoc,taConsult)
else if Retour='MOD' then begin SaisiePiece(CleDoc,taModif); AglRefreshDB([LongInt(Ecran),'FListe'],2); end
else if Retour='PRE' then begin TransformePiece(CleDoc,'PRE'); AglRefreshDB([LongInt(Ecran),'FListe'],2); end
else if Retour='BLF' then begin TransformePiece(CleDoc,'BLF'); AglRefreshDB([LongInt(Ecran),'FListe'],2); end
else if Retour='BLC' then begin TransformePiece(CleDoc,'BLC'); AglRefreshDB([LongInt(Ecran),'FListe'],2); end
else if Retour='FAC' then begin TransformePiece(CleDoc,'FAC'); AglRefreshDB([LongInt(Ecran),'FListe'],2); end ;
end;

{Function TOF_GCCONTREAVT.Evaluedate (St : String) : TDateTime;
Var dd,mm,yy : Word;
begin
Result := 0; if St='' then Exit;
dd := StrToInt(Copy(St,1,2)); mm := StrToInt(Copy(St,3,2)); yy := StrToInt(Copy(St,5,4));
Result := Encodedate(yy,mm,dd);
end;}

Procedure TOF_GCCONTREAVT.InverseTobs;
var
   i_ind1, i_ind2 : integer;
   TobMere, TobTemp, TobTemp2, TobTemp3 : TOB;

   function VerifieExisteDeja (Mere, Fille : TOB) : TOB;
   begin
   Result := nil;
   if Fille.FieldExists('GL_NATUREPIECEG') then
      Result := Mere.FindFirst(['GL_NATUREPIECEG','GL_NUMERO','GL_SOUCHE','GL_INDICEG'],
                                [Fille.GetValue('GL_NATUREPIECEG'),
                                 Fille.GetValue('GL_NUMERO'),
                                 Fille.GetValue('GL_SOUCHE'),
                                 Fille.GetValue('GL_INDICEG')], False)
      else
      Result := Mere.FindFirst(['GQC_DEPOT','GQC_REFERENCE','GQC_ARTICLE','GQC_FOURNISSEUR','GQC_CLIENT'],
                                [Fille.GetValue('GQC_DEPOT'),
                                 Fille.GetValue('GQC_REFERENCE'),
                                 Fille.GetValue('GQC_ARTICLE'),
                                 Fille.GetValue('GQC_FOURNISSEUR'),
                                 Fille.GetValue('GQC_CLIENT')], False);
   end;

begin
NbLignesGridDown := 0;
TobMere := TOB.Create('', nil, -1);
for i_ind1 := 0 to TobCont.Detail.Count - 1 do
    begin
    TobTemp := TobCont.Detail[i_ind1];
    for i_ind2 := 0 to TobTemp.Detail.Count - 1 do
        begin
        TobTemp2 := VerifieExisteDeja (TobMere, TobTemp.Detail[i_ind2]);
        if TobTemp2 = nil then
           begin
           TobTemp2 := TOB.Create ('', TobMere, -1);
           TobTemp2.Dupliquer(TobTemp.Detail[i_ind2], False, True);
           end;
        TobTemp3 := VerifieExisteDeja (TobTemp2, TobTemp);
        if TobTemp3 = nil then
           begin
           TobTemp3 := TOB.Create ('', TobTemp2, -1);
           TobTemp3.Dupliquer(TobTemp, False, True);
           end;
        end;
    end;
TobCont.ClearDetail;
TobCont.Dupliquer(TobMere, True, True);
TobMere.Free;
for i_ind1 := 0 to TobCont.Detail.Count - 1 do
    begin
    TobTemp := TobCont.Detail[i_ind1];
    NbLignesGridDown := Max(NbLignesGridDown, TobTemp.Detail.Count);
    end;
end;

{========================================================================================}
{========================= Actions liées au Grid ========================================}
{========================================================================================}
procedure TOF_GCCONTREAVT.EtudieColsListes;
begin
LesColsGen := InitialiseEntete(G_General);
LesColsDet := InitialiseEntete(G_Detail);
end;

procedure TOF_GCCONTREAVT.DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
Var Coord : TRect;
    ST : string;
begin
if (ACol = 0) AND (ARow > 0) then
    begin
    Coord := G_General.CellRect (ACol, ARow);
    Canvas.Font.Name := 'Wingdings 3';
    Canvas.Font.Size := 10;
    Canvas.Font.Style := [fsBold];
    st := G_General.Cells [ACol, Arow];
    Canvas.TextOut ( (Coord.Left+Coord.Right) div 2 - Canvas.TextWidth(st) div 2,
                     (Coord.Top+ Coord.Bottom) div 2 - Canvas.TextHeight(st) div 2, st);
    end;
end;

function TOF_GCCONTREAVT.InitialiseEntete(G_SEL : THGrid) : string;
var
    St, StA, FF, Perso, NomList, LesCols, CH, Typ : string ;
    tal : TAlignment ;
    FFieldList,FRecordSource,FLien,FTitre,FSortBy,FLargeur,FAlignement,FParams,tt,NC : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;
    Dec, icol, Larg : integer;
begin
LesCols := '';
NomList := G_SEL.ListeParam;
ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,Perso,OkTri,OkNumCol);
G_SEL.Titres.Text := '';
G_SEL.ColCount:=50;
icol:=0; LesCols:='';
if G_SEL=G_General then LesChampsDispo:=FFieldList;
if G_SEL=G_Detail then LesChampsLig:=FFieldList;
G_SEL.Titres.Add(FFieldList) ;
While Ftitre<> '' do
    BEGIN
    StA:=ReadTokenSt(FAlignement);
    St:=ReadTokenSt(Ftitre);
    CH:=ReadTokenSt(FFieldList);
    Larg:=ReadTokenI(FLargeur);
    tal:=TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    if OkVisu then
        BEGIN
        G_SEL.Cells[icol,0]:=St ;
        G_SEL.ColAligns[icol]:=tal;
        G_SEL.ColWidths[icol]:=Larg*G_SEL.Canvas.TextWidth('W') ;;
        if OkLib then G_SEL.ColFormats[icol]:='CB=' + Get_Join(Ch);
        Typ:=ChampToType(Ch) ;
        if (Typ='INTEGER') or (Typ='SMALLINT') or (Typ='DOUBLE') then G_SEL.ColFormats[icol]:=FF ;
        if LesCols='' then LesCols:=CH else LesCols:=LesCols+';'+CH;
        inc (icol);
        END;
    END;
G_SEL.ColCount:=icol ;
Result := LesCols ;
end;

procedure TOF_GCCONTREAVT.InitialiseGrille;
begin
G_General.VidePile (False);
G_Detail.VidePile (False);
G_General.Cols[0].Clear;
G_Detail.Cols[0].Clear;
end;

procedure TOF_GCCONTREAVT.ChargeMouvements;
Var InverseGrille : boolean;
begin
if not (bGeneralOnTop) then begin BINVERSEClick(Self);InverseGrille:=True; end else InverseGrille:=False;
TobCont.ClearDetail;
LoadLesTobLigneREFART;
if TobCont.Detail.Count > 0 then
    begin
    G_General.Enabled:= True;
    G_General.RowCount := TobCont.Detail.Count + 1;
    TobCont.PutGridDetail (G_General, False, False, LesColsGen, False);
    G_General.SetFocus;
    G_General.Cells [0, 1] := '„';
    G_General.Cells [0, 0] := '';
    TOB(G_General.Objects[0, 1]).PutGridDetail(G_Detail, False, False, LesColsDet);
    if InverseGrille then BINVERSEClick(Self);
    end else
    begin
    InitialiseGrille;
    G_General.RowCount := 2;
    G_General.Enabled:= False;
    end;
end;

procedure TOF_GCCONTREAVT.LoadLesTobLigneREFART;
var stWhere, stWhereAnd, stWhereWhere,stWhereCri, stLesCols, stChamp, stSep : string;
    i_ind1, i_ind2 : integer;
    TSql : TQuery;
    TobTemp, TobTemp2 : TOB;
    LignePresente,OkTri,OkNumCol : boolean;
    Liste,FFieldList,FRecordSource,FLien,FTitre,FSortBy,FLargeur,FAlignement,FParams,tt,NC,Perso : string ;
begin
stLesCols := LesChampsDispo;
stSep := ReadTokenSt(stLesCols);
stSep := '';
while stLesCols <> '' do
    begin
    stChamp := stChamp + stSep + ReadTokenSt(stLesCols);
    stSep := ',';
    end;
stWhereCri:=RecupeCritereAvance;
stWhere := 'Select ' + stChamp + ' from DISPOCONTREM ' + stWhereCri;
if stWhereCri<>'' then
   begin
   end else
    stWhereWhere := '';
    stWhereAnd := ' and ';
   begin
   stWHereWhere := ' where ';
   stWhereAnd := '';
   end;
if GQC_REFERENCE.Text <> '' then
    begin
    stWhere := stWhere + stWhereWhere + stWhereAnd + 'GQC_REFERENCE="' + GQC_REFERENCE.Text + '"';
    stWhereWhere := '';
    stWhereAnd := ' and ';
    end;
if GQC_DEPOT.Value <> '' then
    begin
    stWhere := stWhere + stWhereWhere + stWhereAnd + 'GQC_DEPOT="' + GQC_DEPOT.Value + '"';
    stWhereWhere := '';
    stWhereAnd := ' and ';
    end;
if GQC_ARTICLE.Text <> '' then
    begin
    stWhere := stWhere + stWhereWhere + stWhereAnd + 'GQC_ARTICLE like "' + GQC_ARTICLE.Text + '%"';
    stWhereWhere := '';
    stWhereAnd := ' and ';
    end;
if GP_TIERS.Text <> '' then
    begin
    stWhere := stWhere + stWhereWhere + stWhereAnd + 'GQC_CLIENT="' + GP_TIERS.Text + '"';
    stWhereWhere := '';
    stWhereAnd := ' and ';
    end;
if GQC_FOURNISSEUR.Text <> '' then
    begin
    stWhere := stWhere + stWhereWhere + stWhereAnd + 'GQC_FOURNISSEUR="' + GQC_FOURNISSEUR.Text + '"';
    stWhereWhere := '';
    stWhereAnd := ' and ';
    end;

THSQLMemo(GetControl('Z_SQL')).Text := stWhere;

LignePresente := False;
TSql := OpenSql (stWhere, True);
if not TSql.Eof then
    begin
    TobCont.LoadDetailDB ('DISPOCONTREM', '', '', TSql, False);
    Ferme (TSql);
    Liste:='GCMULCONTREAVTLT';
    ChargeHListe(Liste,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,Perso,OkTri,OkNumCol);
    stLesCols := LesChampsLig;
    ReadTokenSt(stLesCols); ReadTokenSt(FFieldList);
    stChamp := 'GL_DATEPIECE,GL_NATUREPIECEG,GL_NUMERO,GL_LIBELLE,GL_QTESTOCK, ' +
                   'GL_TIERS,GL_FOURNISSEUR,GL_DEPOT,GL_INDICEG,GL_SOUCHE ';
    while stLesCols <> '' do
        begin
        stSep := ReadTokenSt(stLesCols);
        if pos(StSep,stChamp)=0 then stChamp := stChamp+','+stSep
        end;
    while FFieldList <> '' do
        begin
        stSep := ReadTokenSt(FFieldList);
        if pos(StSep,stChamp)=0 then stChamp := stChamp+','+stSep
        end;
    NbLignesGridDown := 0;
    for i_ind1 := 0 to TobCont.Detail.Count - 1 do
        begin
        TobTemp := TobCont.Detail[i_ind1];         
        stWhere := 'Select ' + stChamp + ' from LIGNE ' +
                   'where GL_DATEPIECE>="' + UsDateTime(StrToDate(GL_DATEPIECE.Text)) + '" ' +
                   'and GL_DATEPIECE<="' + UsDateTime(StrToDate(GL_DATEPIECE_.Text)) + '" ' +
                   'and ((GL_TIERS="' + TobTemp.GetValue('GQC_CLIENT') + '" ' +
                   'and GL_FOURNISSEUR="' + TobTemp.GetValue('GQC_FOURNISSEUR') + '") ' +
                   'or (GL_TIERS="' + TobTemp.GetValue('GQC_FOURNISSEUR') + '" ' +
                   'and GL_FOURNISSEUR="' + TobTemp.GetValue('GQC_CLIENT') + '")) ' +
                   'and GL_REFCATALOGUE="' + TobTemp.GetValue('GQC_REFERENCE') + '" ' +
                   'and GL_DEPOT="' + TobTemp.GetValue('GQC_DEPOT') + '" ' +
                   'and GL_VIVANTE="X" AND GL_QTERESTE > 0 order by GL_DATEPIECE'; {DBR NEWPIECE QTERESTE > 0}
        TSql := OpenSql (stWhere, True);
        if not TSql.Eof then
            begin
            TobTemp.LoadDetailDB('', '', '', TSql, False);
            LignePresente := True;
            NbLignesGridDown := Max(NbLignesGridDown, TobTemp.Detail.Count);
            for i_ind2 := 0 to TobTemp.Detail.Count - 1 do
                begin
                TobTemp2 := TobTemp.Detail[i_ind2];
                TobTemp2.PutValue('GL_TIERS',RechDom('GCTIERS',TobTemp2.GetValue('GL_TIERS'),False));
//                TobTemp2.PutValue('GL_NATUREPIECEG',GetInfoParPiece(TobTemp2.GetValue('GL_NATUREPIECEG'),'GPP_LIBELLE'));
//                TobTemp2.PutValue('GL_DEPOT',RechDom('GCDEPOT',TobTemp2.GetValue('GL_DEPOT'),False));
                end;
            end;
        Ferme(TSql);
        TobTemp.PutValue('GQC_CLIENT',RechDom('GCTIERSCLI',TobTemp.GetValue('GQC_CLIENT'),False));
        TobTemp.PutValue('GQC_FOURNISSEUR',RechDom('GCTIERSFOURN',TobTemp.GetValue('GQC_FOURNISSEUR'),False));
//        TobTemp.PutValue('GQC_DEPOT',RechDom('GCDEPOT',TobTemp.GetValue('GQC_DEPOT'),False));
        end;
    TobCont.PutGridDetail(G_General, False, False, LesColsGen);
    end else Ferme (TSql);

if not LignePresente then
    begin
    G_Detail.Visible := False;
    BINVERSE.Visible := False;
    Split.Visible := False;
    G_General.Align := alClient;
    Split.Align := alBottom;;
    end else
    begin
    G_Detail.Visible := True;
    BINVERSE.Visible := True;
    Split.Visible := True;
    G_General.Align := alTop;
    //G_General.Height := Max(Trunc(PQTE.Height/2),(PQTE.Height-(G_Detail.DefaultRowHeight*(NbLignesGridDown+1))-12));
    G_General.Height := Min(Trunc(PQTE.Height/2),(G_General.DefaultRowHeight*(G_General.RowCount+2)+12));
    G_Detail.Align := alClient;
    Split.Align := alTop;
    end;
end;

function TOF_GCCONTREAVT.RecupeCritereAvance : string;
Var StZ_C1,StZ_C2,StZ_C3,StRempl,Prefixe,stWhereCri : string;
    i_pos,i_fin,i_point : integer;
begin
StZ_C1:=GetControlText('Z_C1');
StZ_C2:=GetControlText('Z_C2');
StZ_C3:=GetControlText('Z_C3');
stWhereCri:=RecupWhereCritere(TPageControl(GetControl('PAGES')));
i_pos:=Pos('WHERE (AND',stWhereCri);
if i_pos>0 then stWhereCri:='Where ('+Copy(stWhereCri,i_pos+10,length(StWhereCri)) else stWhereCri:='';
if stWhereCri<>'' then
   begin
   if (StZ_C1<>'') and (Copy(StZ_C1,2,4)<>'GQC_') then
      begin
      i_Point:=Pos('.',StZ_C1);
      i_fin:=Pos('_',StZ_C1);
      Prefixe:=Copy(StZ_C1,i_Point+1,i_fin-i_point-1);
      if Prefixe='GDE' then StRempl:='GQC_DEPOT'
      else if Prefixe='GA' then StRempl:='GQC_ARTICLE'
       else if Prefixe='T' then StRempl:='GQC_FOURNISSEUR'
        else if Prefixe='GQL' then StRempl:='GQC_CLIENT'
         else if Prefixe='GEM' then StRempl:='GQC_EMPLACEMENT';
      stWhereCri:=FindEtReplace(stWhereCri,StZ_C1,StRempl,True) ;
      end;
   if (StZ_C2<>'') and (Copy(StZ_C2,2,4)<>'GQC_') then
      begin
      i_Point:=Pos('.',StZ_C2);
      i_fin:=Pos('_',StZ_C2);
      Prefixe:=Copy(StZ_C2,i_Point+1,i_fin-i_point-1);
      if Prefixe='GDE' then StRempl:='GQC_DEPOT'
      else if Prefixe='GA' then StRempl:='GQC_ARTICLE'
       else if Prefixe='T' then StRempl:='GQC_FOURNISSEUR'
        else if Prefixe='GQL' then StRempl:='GQC_CLIENT'
         else if Prefixe='GEM' then StRempl:='GQC_EMPLACEMENT';
      stWhereCri:=FindEtReplace(stWhereCri,StZ_C2,StRempl,True) ;
      end;
   if (StZ_C3<>'') and (Copy(StZ_C3,2,4)<>'GQC_') then
      begin
      i_Point:=Pos('.',StZ_C3);
      i_fin:=Pos('_',StZ_C3);
      Prefixe:=Copy(StZ_C3,i_Point+1,i_fin-i_point-1);
      if Prefixe='GDE' then StRempl:='GQC_DEPOT'
      else if Prefixe='GA' then StRempl:='GQC_ARTICLE'
       else if Prefixe='T' then StRempl:='GQC_FOURNISSEUR'
        else if Prefixe='GQL' then StRempl:='GQC_CLIENT'
         else if Prefixe='GEM' then StRempl:='GQC_EMPLACEMENT';
      stWhereCri:=FindEtReplace(stWhereCri,StZ_C3,StRempl,True) ;
      end;
   end;
Result:=stWhereCri;
end;

Initialization
  registerclasses ( [ TOF_GCCONTREAVT ] ) ;
end.
