unit UTofInitStock;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Hpanel,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB,AglInit,EntGC,SaisUtil,graphics,
      grids,windows,M3FP,HTB97,Dialogs, AGLInitGC, ExtCtrls, Hqry, LicUtil,
{$IFDEF EAGLCLIENT}
      MaineAgl,eMul,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,Mul,
{$ENDIF}
{$IFDEF BTP}
      lookUp,
{$ENDIF}
      UTofOptionEdit,HDimension, TarifUtil, UTofPrixRevient 
      ,FactUtil // JT eQualité 10823
      ;

Type
    TOF_InitStock = Class (TOF)
        procedure FormKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    private
        GIS : THGRID ;
        BImprimer: TToolbarButton97;
        BValider: TToolbarButton97;
        procedure DessineCellIs (ACol,ARow : Longint; Canvas : TCanvas;
                                      AState: TGridDrawState);
        function TestDepot : boolean; //JT eQualité 10823
    public
        TobInitStk : TOB;
        stColInitStock, stColListe, stColOblig, stDepot : String;
        FixedWidthArt : integer;
        ColDepot, ColQteInit, ColDPA, ColPMAP, ColDPR : integer;
{$IFDEF BTP}
        ColEmplacement,ColOk, ColQteMini,ColDateInit : integer;
{$ENDIF}
        d_dpa : double;
        procedure OnArgument (Arguments : String ) ; override;
        procedure OnClose ; override;
        procedure OnUpdate; override;
    // Initialisation
        procedure AffecteLibelleColonne;
        procedure DepileTob;
    // Impression
        procedure BImprimerClick(Sender: TObject);
        function PrepareImpression : integer ;
    // Actions liées au Grid
{$IFDEF BTP}
        procedure GISCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GISKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        function EmplacementExists (Emplacement : string; CurrTOB : tob) : integer;
        procedure GISElipsisClick(Sender: TObject);
{$ENDIF}
        procedure GISCellExit (Sender: TObject; var ACol, ARow: integer;
                               var Cancel: boolean);
        procedure GISDblClick (Sender : TObject);
        procedure GISRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                                   Chg: Boolean);
        procedure GISRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                                  Chg: Boolean);
        procedure BValiderClick (Sender : TObject);
    // Gestion des données
        procedure AfficheDim;
        procedure AjoutSelect (var stSelect : string; stChamp : string);
        procedure ChargeLibDim;
        procedure ChargeTob;
        Function  RecupRatioMesure(Cat, Mesure : String) : Double;
        Function  GetTob (Grid : THGrid; ARow, ARowMax : integer) : TOB ;
        procedure InitialiseGrille;
        procedure LoadLesTob;
    end;

Const NbDecVSup : integer = 2;

Var PEtat : RPARETAT;

implementation

procedure TOF_InitStock.DessineCellIs (ACol,ARow : Longint;
                                       Canvas : TCanvas;
                                       AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < GIS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GIS.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GIS.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = GIS.row) then
        begin
        Canvas.Brush.Color := clBlack ;
        Canvas.Pen.Color := clBlack ;
        Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ;
        Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
        Triangle[0].X:=Triangle[1].X-5 ;
        Triangle[0].Y:=Triangle[1].Y-5 ;
        Triangle[2].X:=Triangle[1].X-5 ;
        Triangle[2].Y:=Triangle[1].Y+5 ;
        Canvas.Polygon(Triangle) ;
        end;
    end;
end;

procedure TOF_InitStock.OnArgument (Arguments : String ) ;
begin
inherited ;
PEtat.Tip:='E';
PEtat.Nat:='GZI';
PEtat.Modele:='GZI';
PEtat.Titre:=TFmul(Ecran).Caption;
PEtat.Apercu:=True;
PEtat.DeuxPages:=False;
PEtat.First:=True;
PEtat.stSQL:='';

GIS := THGRID(GetControl('GIS'));
BValider := TToolBarButton97(GetControl('BOuvrir'));
BValider.OnClick := BValiderClick;
GIS.PostDrawCell := DessineCellIs;
GIS.OnDblClick := GISDblClick;
GIS.OnRowEnter := GISRowEnter;
GIS.OnRowExit := GISRowExit;
GIS.OnCellExit := GISCellExit;
{$IFDEF BTP}
GIS.OnCellEnter := GISCellEnter;
GIS.OnKeyDown := GISKeyDown;
GIS.OnElipsisClick := GISElipsisClick ;
{$ENDIF}
TFMul(Ecran).OnKeyDown := FormKeyDown;

BImprimer:=TToolbarButton97(GetControl('Bimprimer')) ;
BImprimer.OnClick:=BImprimerClick;

if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    begin
    TToolbarButton97(GetControl('BParamListe')).Visible:=True ;
    end else
    begin
    TToolbarButton97(GetControl('BParamListe')).Visible:=False;
    end;

stColOblig := 'GA_CODEARTICLE;GA_LIBELLE;GA_ARTICLE;GA_STATUTART;GA_GRILLEDIM1;GA_GRILLEDIM2;' +
              'GA_GRILLEDIM3;GA_GRILLEDIM4;GA_GRILLEDIM5;GA_CODEDIM1;GA_CODEDIM2;GA_CODEDIM3;' +
              'GA_CODEDIM4;GA_CODEDIM5;GA_DPA;GA_PMAP;GA_DPR;GA_PMRP;GA_QUALIFUNITESTO;GA_QUALIFUNITEVTE';

FixedWidthArt := 10;
THValComboBox (GetControl('_GDE_DEPOT')).ItemIndex := 0;
SetControlVisible('Fliste',False);
end;

procedure TOF_InitStock.OnClose;
begin
inherited;
DepileTob;
GIS.VidePile (False);
end;

procedure TOF_InitStock.OnUpdate;
begin
inherited;
InitialiseGrille;
SetControlText ('XX_WHERE', '');
end;

{========================================================================================}
{========================= Impression ===================================================}
{========================================================================================}
procedure TOF_InitStock.BImprimerClick(Sender: TObject);
begin
if TobInitStk=Nil then exit;

EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('Pages')),PEtat.stSQL,PEtat.Titre, PrepareImpression);
ExecuteSQL('DELETE FROM GCTMPINITSTOCK WHERE GZI_UTILISATEUR = "'+V_PGI.USer+'"');
end;

function TOF_InitStock.PrepareImpression : integer ;
var i_ind : integer;
    TobGZI, TobD, TobIS : TOB;
    stChamp, stLesCol, stCodeArticle : string;
begin
Result:=0;
stCodeArticle := '';
ExecuteSQL('DELETE FROM GCTMPINITSTOCK WHERE GZI_UTILISATEUR = "'+V_PGI.USer+'"');
TobGZI:=TOB.Create('',Nil,-1);
for i_ind:=0 to TobInitStk.Detail.Count-1 do
    begin
    TobIS:=TobInitStk.Detail[i_ind] ;
    TobD:=TOB.Create('GCTMPINITSTOCK',TobGZI,-1);
    TobD.PutValue('GZI_UTILISATEUR',V_PGI.USer);
    TobD.PutValue('GZI_COMPTEUR',i_ind);
    stLesCol := stColListe;
    stChamp := ReadTokenSt (stLesCol);
    while stChamp <> '' do
        begin
        if (pos ('GA_', stChamp) = 1) then
            begin
            TobD.PutValue ('GZI_' + Copy (stChamp, 4, Length(stChamp) - 3),
                           TobIS.GetValue (stChamp));
            end;
        stChamp := ReadTokenSt (stLesCol);
        end;

    TobD.PutValue ('GZI_DEPOT', stDepot);
    TobD.PutValue ('GZI_QTEINIT', TobIS.GetValue ('QTEINIT'));
    TobD.PutValue ('GZI_DATEINIT', TobIS.GetValue ('DATEINIT'));
    TobD.PutValue ('GZI_LIBDIM', TobIS.GetValue ('LIBDIM'));
    end;
TobGZI.InsertDB (Nil);
TobGZI.Free;
end;

{========================================================================================}
{========================= Actions liées au Grid ========================================}
{========================================================================================}
procedure TOF_InitStock.FormKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
    VK_RETURN :
        if Screen.ActiveControl=GIS then Key:=VK_TAB;
    VK_F10    :
        begin
        Key:=0;
        BValiderClick (Sender);
        end;
    end;
end;

{$IFDEF BTP}
procedure TOF_InitStock.GISKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var TOBIS : TOB;
begin
  if TobInitStk=Nil then exit;
  TobIS := GetTob (GIS, GIS.row, TobInitStk.Detail.Count);
  {Ctrl+O}
  if (Shift = [ssctrl]) and (Key = 79) then
  begin
    if (TOBIS.GetValue('QTEINIT')=0) and (TOBIS.GetValue('SELECTOK')='X') then
    begin
      TOBIS.Putvalue('SELECTOK','-');
    end else
    begin
      TOBIS.Putvalue('SELECTOK','X');
    end;
    TobIS.PutLigneGrid (GIS,GIS.row,false,false,
                        'FIXED;' + stColListe + ';QTEINIT;DATEINIT;QTEMINI;EMPLACEMENT;SELECTOK');
    Key := 0;
  end;
end;

procedure TOF_InitStock.GISCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  GIS.ElipsisButton := (GIS.col = ColEmplacement) or (GIS.col = ColDateInit) ;
end;

function TOF_InitStock.EmplacementExists (Emplacement : string; CurrTOB : tob) : integer;
var QQ : TQuery;
    Select : String;
    TheTOB : TOB;
begin
  result := 0;
  if Emplacement = '' then exit;
  Select := 'SELECT GEM_MONOARTICLE,GEM_EMPLACEOCCUPE FROM EMPLACEMENT WHERE GEM_DEPOT="'+ stDepot +
            '" AND GEM_EMPLACEMENT="'+Emplacement+'"';
  QQ := OpenSQL(Select,False);
  TRY
    if QQ.eof then
    begin
      result := 1
    end else
    begin
      if (QQ.FindField('GEM_MONOARTICLE').AsString = 'X') then
      begin
        if (QQ.FindField('GEM_EMPLACEOCCUPE').AsString = 'X') then
        begin
          result := 2
        end else
        begin
          TheTob := TobInitStk.FindFirst (['EMPLACEMENT'],[Emplacement],true);
          if (TheTOB <> nil) and (TheTOB <> CurrTob) then result := 2;
        end;
      end;
    end;
  FINALLY
    Ferme (QQ);
  END;
end;

procedure TOF_InitStock.GISElipsisClick(Sender: TObject);
var Result : boolean;
  HDATE: THCritMaskEdit;
  Coord: TRect;
begin
  if GiS.col = ColEmplacement then
  begin
    Result := LookupList(GIS, 'Liste des emplacements','EMPLACEMENT','GEM_EMPLACEMENT','GEM_LIBELLE',
                        'GEM_DEPOT="'+stDepot+'" AND GEM_EMPLACEOCCUPE<>"X"','GEM_EMPLACEMENT',False,-1)
  end else if GIS.col = ColDateInit then
  begin
    Coord := GiS.CellRect(GiS.Col, GiS.Row);
    HDATE := THCritMaskEdit.Create(GiS);
    HDATE.Parent := GiS;
    HDATE.Top := Coord.Top;
    HDATE.Left := Coord.Left;
    HDATE.Width := 3;
    HDATE.Visible := False;
    HDATE.OpeType := otDate;
    GetDateRecherche(TForm(HDATE.Owner), HDATE);
    if HDATE.Text <> '' then GiS.Cells[GiS.Col, GiS.Row] := HDATE.Text;
    HDATE.Free;
  end;
end;

{$ENDIF}

procedure TOF_InitStock.GISCellExit (Sender : TObject; var ACol, ARow: integer;
                                     var Cancel: boolean);
var TobIS : TOB;
    dValeur : double;
    stValeur : string;
    ErrEmplacement : integer;
Begin
if TobInitStk=Nil then exit;
TobIS := GetTob (GIS, ARow, TobInitStk.Detail.Count);
stValeur := FormatFloat (GIS.ColFormats[ACol], dValeur);
if (ACol = ColDPA) and (TobIS.GetValue ('GA_DPA')<>dValeur) then
    begin
		dValeur := Valeur (GIS.Cells [ACol, ARow]) ;
    TobIS.PutValue ('GA_DPA', dValeur);
    dValeur := CalculPrixRevient (TobIS.GetValue ('GA_ARTICLE'),
                                  '', stDepot, TobIS.GetValue ('GA_DPA'));
    TobIS.PutValue ('GA_DPR', dValeur);
    GIS.Cells[ColDPR, ARow] :=
        FormatFloat(GIS.ColFormats[ColDPR], dValeur);
    end;
if ACol = ColPMAP then
    begin
		dValeur := Valeur (GIS.Cells [ACol, ARow]) ;
    TobIS.PutValue ('GA_PMAP', dValeur);
    end;
if ACol = ColQteInit then
    begin
{$IFDEF BTP}
		dValeur := Valeur (GIS.Cells [ACol, ARow]) ;
    if dvaleur > 0 then
    begin
      TOBIS.Putvalue('SELECTOK','X');
    end else
    begin
      if TOBIS.GetValue('QTEINIT') > 0 then TOBIS.Putvalue('SELECTOK','-');
    end;
{$ENDIF}
    TobIS.PutValue ('QTEINIT', dValeur);
    end;
  if Acol = ColDateInit then
  begin
  	if IsValidDate(GIS.Cells [ACol, ARow]) then
    begin
    	TobIS.PutValue ('DATEINIT', StrToDate(GIS.Cells [ACol, ARow]));
    end;
  end;
{$IFDEF BTP}
  if Acol = ColQteMini then
  begin
		dValeur := Valeur (GIS.Cells [ACol, ARow]) ;
    TobIS.PutValue ('QTEMINI', dValeur);
  end;

if Acol = ColEmplacement then
begin
  ErrEmplacement := EmplacementExists(GIS.Cells [ACol, ARow],TOBIs);
  if ErrEmplacement = 0 then
  begin
      TOBIS.putValue('EMPLACEMENT',GIS.Cells [ACol, ARow]);
  end else
  begin
    if ErrEmplacement = 1 then
    begin
      PGIBox (traduireMemoire('L''emplacement n''existe pas'),ecran.caption);
    end else if ErrEmplacement = 2 then
    begin
      PGIBox (traduireMemoire('Cet emplacement ne peut pas être réutilisé'),ecran.caption);
    end;
    cancel := true;
  end;
end else
begin
{$ENDIF}
  GIS.Cells [ACol, ARow] := stValeur;
{$IFDEF BTP}
end;
TobIS.PutLigneGrid (GIS,Arow,false,false,'FIXED;' + stColListe + ';QTEINIT;DATEINIT;QTEMINI;EMPLACEMENT;SELECTOK');
{$ENDIF}
end;

procedure TOF_InitStock.GISDblClick(Sender: TObject);
var TobIS : TOB;
begin
if TobInitStk=Nil then exit;
TobIS := GetTob (GIS, GIS.Row, TobInitStk.Detail.Count);
if TobIS.GetValue ('GA_CODEARTICLE') <> '' then
{$IFDEF BTP}
    V_PGI.DispatchTT (7,taConsult,TobIS.GetValue ('GA_ARTICLE'),'','MONOFICHE');
{$ELSE}
    AglLanceFiche ('GC', 'GCARTICLE', '', TobIS.GetValue ('GA_ARTICLE'),
                   'ACTION=CONSULTATION;MONOFICHE');
{$ENDIF}
end;

procedure TOF_InitStock.GISRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                                            Chg: Boolean);
begin
  if TestDepot then //JT eQualité 10823
  begin
    GIS.InvalidateRow (ou);
    AfficheDim;
  end;
end;

procedure TOF_InitStock.GISRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                                           Chg: Boolean);
begin
GIS.InvalidateRow (ou);
end;

procedure TOF_InitStock.BValiderClick (Sender : TObject);
var TobDispo, TobArt, TobIS, TobD, TobA : TOB;
    iInd : integer;
    dValeur : double;
    FUS,FUV,Prix : Double ;
{$IFDEF BTP}
    Acol,Arow: integer;
    Cancel,Chg : boolean;
    TOBEmplac,TOBE : tob;
    QQ : TQuery;
    Select : string;
{$ENDIF}
begin
{$IFDEF BTP}
  Acol := GIS.col;
  Arow := GIS.row;
  GISCellExit (self,Acol,Arow,cancel);
  if cancel then exit;
{$ENDIF}
if TobInitStk=Nil then exit;
TobDispo := Tob.Create ('', Nil, -1);
TobArt := TOB.Create('', nil, -1);
TOBEmplac := TOB.create ('LES EMPLACEMENTS',nil,-1);
for iInd := 0 to TobInitStk.Detail.Count - 1 do
    begin
    TobIS := TobInitStk.Detail[iInd];

    dValeur :=  Valeur (TobIS.GetValue ('QTEINIT'));
{$IFDEF BTP}
    if TOBIS.GetValue('SELECTOK') = 'X' then
{$ELSE}
    if dValeur > 0 then
{$ENDIF}
        begin
        TobD := Tob.Create ('DISPO', TobDispo, -1);
        TobD.InitValeurs;
        TobD.PutValue ('GQ_ARTICLE', TobIS.GetValue ('GA_ARTICLE'));
        TobD.PutValue ('GQ_DEPOT', stDepot);
        TobD.PutValue ('GQ_PHYSIQUE', TobIS.GetValue ('QTEINIT'));
        TobD.PutValue ('GQ_DPA', TobIS.GetValue ('GA_DPA'));
        TobD.PutValue ('GQ_PMAP', TobIS.GetValue ('GA_PMAP'));
        TobD.PutValue ('GQ_DPR', TobIS.GetValue ('GA_DPR'));
        TobD.PutValue ('GQ_PMRP',
                       CalculPrixRevient (TobD.GetValue ('GQ_ARTICLE'),
                                          '', TobD.GetValue ('GQ_DEPOT'),
                                          TobD.GetValue ('GQ_PMAP')));
        TobD.PutValue ('GQ_UTILISATEUR', V_PGI.USer);
        TobD.PutValue ('GQ_STOCKINITIAL', TobIS.GetValue ('QTEINIT'));
        TobD.PutValue ('GQ_STOCKINV',     TobIS.GetValue ('QTEINIT'));
        TobD.PutValue ('GQ_DATEINV',     TobIS.GetValue ('DATEINIT'));
        //
        if TobIS.GetValue ('GA_PMAP') <> 0 then
        begin
        	TobD.PutValue ('GQ_PRIXINV', TobIS.GetValue ('GA_PMAP'));
        end else
        begin
        	TobD.PutValue ('GQ_PRIXINV', TobIS.GetValue ('GA_DPA'));
        end;
        //
        TobD.PutValue ('GQ_CLOTURE', '-');
{$IFDEF BTP}
        if TobIS.GetValue ('QTEMINI') > 0 then
        begin
          TobD.PutValue ('GQ_STOCKMIN', TobIS.GetValue ('QTEMINI'));
          TobD.PutValue ('GQ_STOCKMAX', 99999999);
        end;
        TobD.PutValue ('GQ_EMPLACEMENT', TobIS.GetValue ('EMPLACEMENT'));
        Select := 'SELECT * FROM EMPLACEMENT WHERE GEM_DEPOT="'+StDepot+'" AND '+
                  'GEM_EMPLACEMENT="'+ TobIS.GetValue ('EMPLACEMENT')+'"';
        QQ := OpenSql (Select,true);
        TOBE := TOB.Create ('EMPLACEMENT',TOBEmplac,-1);
        TOBE.SelectDB ('',QQ);
        ferme (QQ);
        if TOBE.GetValue('GEM_MONOARTICLE')='X' Then TOBE.PutValue('GEM_EMPLACEOCCUPE','X');
{$ENDIF}
        TobA := Tob.Create ('ARTICLE', TobArt, -1);
        TobA.InitValeurs;
        TobA.PutValue ('GA_ARTICLE', TobIS.GetValue ('GA_ARTICLE'));
        TobA.ChargeCle1;
        TobA.LoadDB;
        FUS := RecupRatioMesure('PIE', TobIS.GetValue('GA_QUALIFUNITESTO')); if FUS=0 then FUS:=1;
        FUV := RecupRatioMesure('PIE', TobIS.GetValue('GA_QUALIFUNITEVTE')); if FUV=0 then FUV:=1;
        Prix:=TobIS.GetValue('GA_DPA');
        Prix:=Arrondi(Prix*FUV/FUS,V_PGI.OkDecV+NbDecVSup);;
        TobA.PutValue('GA_DPA',Prix);
        Prix:=TobIS.GetValue('GA_PMAP');
        Prix:=Arrondi(Prix*FUV/FUS,V_PGI.OkDecV+NbDecVSup);;
        TobA.PutValue('GA_PMAP',Prix);
        Prix:=TobIS.GetValue('GA_DPR');
        Prix:=Arrondi(Prix*FUV/FUS,V_PGI.OkDecV+NbDecVSup);;
        TobA.PutValue('GA_DPR',Prix);
        TobA.PutValue ('GA_PMRP',
                       CalculPrixRevient (TobD.GetValue ('GQ_ARTICLE'),
                                          '', TobD.GetValue ('GQ_DEPOT'),
                                          TobD.GetValue ('GQ_PMAP')));
        end;
    end;
if TobDispo.Detail.Count > 0 then
    begin
    TobDispo.InsertDB (Nil, True);
    TobArt.UpdateDB;
{$IFDEF BTP}
    TOBEmplac.UpdateDB;
{$ENDIF}
    end;
TobDispo.Free;
TobArt.Free;
TOBEmplac.free;
ChargeTob;
end;

{========================================================================================}
{===================      Initialisation    =============================================}
{========================================================================================}
procedure TOF_InitStock.AffecteLibelleColonne;
Var st, code, libel, stChamp : string ;
    i_ind, ACol : integer;
    TOBTemp, TOBTempD, TobD1 : TOB;
begin
TOBTemp := TOB.Create('', Nil,-1);
stChamp := stColInitStock;
for i_ind := 0 to TFMul(Ecran).Fliste.Columns.Count - 1 do
    begin
    TOBTempD := TOB.Create('', TOBTemp,-1);
    TOBTempD.AddChampSup('CODE',False);
    TOBTempD.AddChampSup('LIBELLE',False);
    TOBTempD.PutValue('CODE',ReadTokenSt(stChamp));
    TOBTempD.PutValue('LIBELLE',TFMul(Ecran).Fliste.Columns.Items[i_ind].Title.Caption) ;
    end;
st := stColListe;
ACol := 1;
While (st <> '') and (ACol < GIS.ColCount) do
    begin
    code := ReadTokenSt(st);
    TobD1 := TOBTemp.FindFirst(['CODE'],[code],True);
    if TobD1 <> nil then Libel := TobD1.GetValue('LIBELLE') else Libel := '';
    GIS.Cells[ACol, 0] := Libel;
    inc(ACol);
    end;
TOBTemp.Free;
GIS.Cells[ColQteInit, 0] := 'Stock Initial';
GIS.Cells[ColDateInit, 0] := 'Date comptage';
{$IFDEF BTP}
GIS.Cells[ColQteMini, 0] := 'Qte mini.';
GIS.Cells[ColEmplacement, 0] := 'Emplacement';
GIS.Cells[ColOK, 0] := 'Valide';
{$ENDIF}
end;

procedure TOF_InitStock.DepileTob;
begin
if TobInitStk <> nil then
    begin
    TobInitStk.Free;
    TobInitStk := nil;
    end;
end;

{========================================================================================}
{================================= Gestion des Données ==================================}
{========================================================================================}
procedure TOF_InitStock.AfficheDim;
var TobIS : TOB;
begin
if TobInitStk=Nil then exit;
TobIS := GetTob (GIS, GIS.Row, TobInitStk.Detail.Count);
if TobIS.GetValue('GA_STATUTART')<>'DIM' then
    THPanel(GetControl('PPIED')).Visible := False
    else
    begin
    THPanel(GetControl('PPIED')).Visible := True;
    THLabel(GetControl('TDIMENSION')).Caption:= TobIS.GetValue ('LIBDIM');
    end;
end;

procedure TOF_InitStock.AjoutSelect (var stSelect : string; stChamp : string);
begin
if Pos (stChamp, stSelect) = 0 then
    begin
    if stSelect <> '' then stSelect := stSelect + ', ';
    stSelect := stSelect + stChamp;
    end;
end;

procedure TOF_InitStock.ChargeLibDim;
var TobIS : TOB;
    iIndex, i_indDim : integer;
    Dim : Array of string;
    GrilleDim,CodeDim,LibDim,St : string;
begin
TobInitStk.Detail[0].AddChampSup ('LIBDIM', True);
TobInitStk.Detail[0].AddChampSup ('QTEINIT', True);
TobInitStk.Detail[0].AddChampSup ('DATEINIT', True);
for iIndex := 0 to TobInitStk.Detail.Count - 1 do
    begin
    TobIS := TobInitStk.Detail[iIndex];

    TobIS.PutValue ('QTEINIT', 0.0);
    TobIS.PutValue ('DATEINIT', Now);

    if TobIS.GetValue('GA_STATUTART')='DIM' then
        begin
        SetLength (dim, MaxDimension);
        St:='';
        for i_indDim := 1 to MaxDimension do
            begin
            GrilleDim := TobIS.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
            CodeDim := TobIS.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
            LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
            if LibDim <> '' then
                if St='' then St := LibDim
                else St := St + ' - ' + LibDim;
            end;
        TobIS.PutValue ('LIBDIM', St);
        end;
    end;
end;

procedure TOF_InitStock.ChargeTob;
var TobTemp : TOB;
    stCol, stChamp : string;
    ind : integer ;
    FUS,FUV,Prix : Double ;
begin
DepileTob;
GIS.VidePile (False);
if not Testdepot then // JT eQualité 10823
  exit;
LoadLesTob;
if TobInitStk = Nil then exit;
if TobInitStk.Detail.Count > 0 then
    begin
    ChargeLibDim;
    for ind:=0 to TobInitStk.Detail.Count-1 do
      begin
      TobTemp:=TobInitStk.Detail[ind];
      FUS := RecupRatioMesure('PIE', TobTemp.GetValue('GA_QUALIFUNITESTO')); if FUS=0 then FUS:=1;
      FUV := RecupRatioMesure('PIE', TobTemp.GetValue('GA_QUALIFUNITEVTE')); if FUV=0 then FUV:=1;
      Prix:=TobTemp.GetValue('GA_DPA');
      Prix:=Arrondi(Prix*FUS/FUV,V_PGI.OkDecV+NbDecVSup);;
      TobTemp.PutValue('GA_DPA',Prix);
      Prix:=TobTemp.GetValue('GA_DPR');
      Prix:=Arrondi(Prix*FUS/FUV,V_PGI.OkDecV+NbDecVSup);;
      TobTemp.PutValue('GA_DPR',Prix);
      Prix:=TobTemp.GetValue('GA_PMAP');
      Prix:=Arrondi(Prix*FUS/FUV,V_PGI.OkDecV+NbDecVSup);;
      TobTemp.PutValue('GA_PMAP',Prix);
      Prix:=TobTemp.GetValue('GA_PMRP');
      Prix:=Arrondi(Prix*FUS/FUV,V_PGI.OkDecV+NbDecVSup);;
      TobTemp.PutValue('GA_PMRP',Prix);
{$IFDEF BTP}
      TobTemp.AddChampSupValeur ('QTEMINI', 0,False);
      TobTemp.AddChampSupValeur ('EMPLACEMENT', '',False);
      TobTemp.AddChampSupValeur ('SELECTOK', '-',False);
{$ENDIF}
      end;
    end else
    begin
    TobTemp := Tob.Create ('', TobInitStk, -1);
    stCol := stColListe;
    stChamp := ReadTokenSt (stCol);
    while stChamp <> '' do
        begin
        TobTemp.AddChampSup (stChamp, False);
        stChamp := ReadTokenSt (stCol);
        end;
    TobTemp.AddChampSup ('QTEINIT', False);
    TobTemp.AddChampSup ('DATEINIT', False);
    TobTemp.InitValeurs;
{$IFDEF BTP}
    TobTemp.AddChampSupValeur ('QTEMINI', 0,False);
    TobTemp.AddChampSupValeur ('EMPLACEMENT', '',False);
    TobTemp.AddChampSupValeur ('SELECTOK', '-',False);
{$ENDIF}
    end;
GIS.Cells [0, 0] := '';
{$IFDEF BTP}
TobInitStk.PutGridDetail (GIS, False, False,
                          'FIXED;' + stColListe + ';QTEINIT;DATEINIT;QTEMINI;EMPLACEMENT;SELECTOK', False);
{$ELSE}
TobInitStk.PutGridDetail (GIS, False, False,
                          'FIXED;' + stColListe + ';QTEINIT', False);
{$ENDIF}
GIS.RowCount := TobInitStk.Detail.Count + 1;
GIS.Row := 1;
GIS.Col := ColQteInit;
AfficheDim;
end;

Function TOF_InitStock.RecupRatioMesure(Cat, Mesure : String) : Double;
var TOBM : TOB;
    X : Double;
begin
TOBM := VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'], [Cat, Mesure], False);
X := 0; if TOBM <> nil then X := TOBM.GetValue('GME_QUOTITE'); if X = 0 then X := 1.0;
result := X;
end;

Function TOF_InitStock.GetTob (Grid : THGrid; ARow, ARowMax : integer) : TOB ;
begin
Result := Nil ;
if ((ARow <= 0) or (ARow > ARowMax)) then Exit ;
Result := TOB (Grid.OBjects [0,ARow]);
end ;

procedure TOF_InitStock.InitialiseGrille;
var iInd,j,Dec, nbColListe : integer ;
    st, stLesCols, stLesTitres, stChamp, Nam, stAl, stA, FF, FPerso : string ;
    NomList,FRecordSource,FLien,FSortBy,FLargeur,FAlignement,FParams : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul, OkTri,OkNumCol : boolean ;
		Ftitre,tt,NC : Hstring;

begin
ColDateInit := -1;
ColQteInit := 0;
ColDPA := 0;
ColPMAP := 0;
ColDPR := 0;
{$IFDEF BTP}
ColEmplacement:= -1;
ColQteMini := -1;
ColOk := -1;
{$ENDIF}
GIS.ColWidths[0] := FixedWidthArt;

NomList := TFMul(Ecran).Q.Liste;
ChargeHListe (NomList, FRecordSource, FLien, FSortBy, stColInitStock, FTitre,
              FLargeur, FAlignement, FParams, tt, NC, FPerso, OkTri, OkNumCol);

stLesCols := stColInitStock;
nbColListe := 0;
stColListe := '';
While stLesCols <> '' do
    begin
    stChamp := ReadTokenSt(stLesCols);
    if pos ('GA_', stChamp) = 1 then
        begin
        inc(nbColListe);
        if nbColListe > 1 then stColListe := stColListe + ';';
        stColListe := stColListe + stChamp;
        end;
    end;
{$IFDEF BTP}
GIS.ColCount := nbColListe + 6; // pour fixed, QteInit, DateInit,QteMini,Emplacement et OK
st := stColListe + ';QTEINIT;DATEINIT;';
{$ELSE}
GIS.ColCount := nbColListe + 2; // pour fixed, QteInit
st := stColListe + ';QTEINIT;EMPLACEMENT;SELECTOK';
{$ENDIF}

for iInd := 1 to GIS.ColCount-1 do
    begin
    Nam := ReadTokenSt (St);
    GIS.ColWidths[iInd]:=80;
    stLesCols := stColInitStock;
    stLesTitres := FTitre;
    StAl := FAlignement;
    for j := 0 to TFMul(Ecran).Fliste.Columns.Count - 1 do
        begin
        StA := ReadTokenSt(StAl);
        TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if (ReadTokenSt(stLesCols)=Nam) or (ReadTokenSt (stLesTitres)=Nam) then
            begin
            GIS.ColAligns[iInd] := TFMul(Ecran).Fliste.Columns.Items[j].Field.Alignment;
            GIS.ColWidths[iInd] := TFMul(Ecran).Fliste.Columns.Items[j].Width;
            if OkLib then GIS.ColFormats[iInd] := 'CB=' + Get_Join(Nam)
                     else if (Dec<>0) or (Sep) then GIS.ColFormats[iInd] := FF ;
            break;
            end;
        end;
    end ;

ColDepot := 1;
{$IFDEF BTP}
ColOk := GIS.colCount -1;
ColEmplacement := GIS.colCount -2;
ColQteMini := GIS.ColCount - 3;
ColDateInit := GIS.ColCount - 4;
ColQteInit := GIS.ColCount - 5;
ColDPR := GIS.ColCount - 6;
ColPMAP := GIS.ColCount - 7;
ColDPA := GIS.ColCount - 8;
{$ELSE}
ColQteInit := GIS.ColCount - 1;
ColDPA := GIS.ColCount - 4;
ColPMAP := GIS.ColCount - 3;
ColDPR := GIS.ColCount - 2;
{$ENDIF}
FF:='# ##0.';
for iInd:=1 to V_PGI.OkDecV+NbDecVSup do FF:=FF+'0';
GIS.ColFormats[ColDPA]:=FF ;
GIS.ColFormats[ColPMAP]:=FF ;
GIS.ColFormats[ColDPR]:=FF ;
{$IFDEF BTP}
GIS.colaligns[ColQteMini]:= taRightJustify;
GIS.ColFormats[ColQteMini]:=FF ;
GIS.ColWidths [ColQteMini]:=420 ;
//
GIS.colaligns[ColDateInit]:= taRightJustify;
GIS.ColTypes[ColDateInit]:='D' ;
GIS.ColFormats[ColDateInit]:=ShortdateFormat ;
GIS.ColWidths [ColDateInit]:=420 ;

GIS.ColWidths [ColEmplacement]:=350 ;

GIS.ColWidths [ColOK]:=100 ;
GIS.ColTypes [ColOK]:='B' ;
GIS.colaligns[ColOK]:= tacenter;
GIS.colformats[ColOK]:= inttostr(Integer(csCoche));
{$ENDIF}
for iInd := 1 To GIS.ColCount - 1 do
    begin
{$IFDEF BTP}
    if (iInd <> ColQteInit) and (iInd <> ColDPA) and (iInd <> ColDateInit)  and
       (iInd <> ColPMAP) and (iInd <> ColEmplacement) and (iInd <> ColQteMini) then
{$ELSE}
    if (iInd <> ColQteInit) and (iInd <> ColDPA) and (iInd <> ColPMAP) then
{$ENDIF}
        begin
        GIS.ColLengths [iInd]:=-1;
        end;
    end;

GIS.ColAligns[ColQteInit]:=GIS.ColAligns[ColPMAP];
GIS.ColWidths[ColQteInit]:=GIS.ColWidths[ColPMAP];
FF:='# ##0.';
for iInd:=1 to V_PGI.OkDecQ do FF:=FF+'0';
GIS.ColFormats[ColQteInit]:=FF;
AffecteGrid (GIS, taModif);
AffecteLibelleColonne;
ChargeTob;
TFMul(Ecran).Hmtrad.ResizeGridColumns(GIS) ;
end;

procedure TOF_InitStock.LoadLesTob;
var stWhere, stSelect, stLesCols, stChamp, stSelectCount : string;
    TSql : TQuery;
begin
stWhere := RecupWhereCritere (TFMul(Ecran).Pages);
stSelect := '';
stLesCols := stColInitStock;
stChamp := ReadTokenSt (stLesCols);
stDepot := GetControlText ('_GDE_DEPOT');
while stChamp <> '' do
    begin
    if stChamp <> '(1)' then
        begin
        if stSelect <> '' then stSelect := stSelect + ', ';
        stSelect := stSelect + stChamp;
        end;
    stChamp := ReadTokenSt (stLesCols);
    end;

stLesCols := stColOblig;
stChamp := ReadTokenSt (stLesCols);
while stChamp <> '' do
    begin
    AjoutSelect (stSelect, stChamp);
    stChamp := ReadTokenst (stLesCols);
    end;
if stWhere <> '' then stWhere := stWhere + ' AND '
else stWhere := ' WHERE ';
stSelect := 'SELECT ' + stSelect + ' FROM ARTICLE ' +
            stWhere + ' GA_STATUTART <> "GEN" AND GA_TENUESTOCK="X" AND NOT EXISTS ' +
            ' (SELECT GQ_ARTICLE FROM DISPO WHERE GQ_DEPOT="' + stDepot + '" AND ' +
            ' GQ_ARTICLE=ARTICLE.GA_ARTICLE) ORDER BY GA_ARTICLE';
stSelectCount := 'SELECT count (GA_ARTICLE) as nombre FROM ARTICLE ' +
            stWhere + ' GA_STATUTART <> "GEN" AND GA_TENUESTOCK="X" AND NOT EXISTS ' +
            ' (SELECT GQ_ARTICLE FROM DISPO WHERE GQ_DEPOT="' + stDepot + '" AND ' +
            ' GQ_ARTICLE=ARTICLE.GA_ARTICLE) ';

TSql := OpenSql (stSelectCount, True);
if not TSql.Eof then
    begin
    if TSql.FindField ('nombre').AsInteger > 2500 then
        begin
        HShowMessage ('0;?caption?;Selection trop importante (supérieure à 2500) ! veuillez la restreindre;W;O;O;O;','','');
        Ferme (TSql);
        end else
        begin
        Ferme (TSql);
        TSql := OpenSql (stSelect, True);
        if not TSql.Eof then
            begin
            TobInitStk := Tob.Create ('', Nil, -1);
            TobInitStk.LoadDetailDB ('ARTICLE', '', '', TSql, False);
            end;
        Ferme (TSql);
        end;
    end;
end;

function TOF_InitStock.TestDepot : boolean; //JT eQualité 10823
begin
  Result := True;
  if TestDepotOblig then
  begin
    if (GetControlText('_GDE_DEPOT') = '') and (GetControlText('XX_WHERE') = '')  then
    begin
      TFMul(Ecran).Hmtrad.ResizeGridColumns(GIS) ;
      PGIBox('Le dépôt doit être renseigné',Ecran.Caption);
      DepileTob;
      GIS.VidePile (False);
      SetFocusControl('_GDE_DEPOT') ;
      Result := False;
    end;
  end;
end;

Initialization
registerclasses([TOF_InitStock]);

end.
