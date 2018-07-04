unit UTofMBOPerEtablis;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,TarifUtil,DBTables,M3FP,Fe_Main,HDB,voirtob;

Type
    TOF_MBOPerEtablis = Class(TOF)
     private
        PerEtab : THGRID ;
        TobPerEtab: Tob ;
        Col_Mov,colEtab,colDebut,colFin,ColArrondi: Integer ;
        LesColonnes,CodePeriode,CodeArrondi: String ;
        DateDebut,DateFin:TDateTime ;

        procedure GSElipsisClick(Sender: TObject) ;
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSEnter(Sender: TObject);

     public
         Action   : TActionFiche ;
         DEV       : RDEVISE ;
         Procedure OnArgument(Arguments:string) ; override ;
         Procedure OnLoad ; override ;
         Procedure OnUpdate ; override ;
         Procedure OnClose ; override ;

         procedure ChargePerEtab ;
         procedure NouveauPerEtab ;
         procedure MiseAJourTarif ;

         procedure ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
         procedure CodeGras ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);

    END ;
    
implementation
uses UTomTarifPer ;

Procedure TOF_MBOPerEtablis.OnArgument(Arguments:string) ;
var St,S,NomCol,Critere,ChampMul,ValMul :String ;
x,i: Integer ;
Begin
inherited ;
St:=Arguments ;
Action:=taModif ;
Repeat
  Critere:=Trim(ReadTokenSt(Arguments)) ;
  if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       ChampMul:=copy(Critere,1,x-1) ;
       ValMul:=copy(Critere,x+1,length(Critere)) ;
       if ChampMul='DateDebut' then DateDebut:=StrToDate(ValMul) ;
       if ChampMul='DateFin' then DateFin:=StrToDate(ValMul) ;
       if ChampMul='CodePeriode' then CodePeriode:=ValMul ;
       if ChampMul='CodeArrondi' then CodeArrondi:=ValMul ;
       end ;
    end ;
until  Critere='' ;
LesColonnes:='MOV;GFP_ETABLISSEMENT;GFP_DATEDEBUT;GFP_DATEFIN;GFP_ARRONDI' ;
PerEtab:=THGRID(GetControl('PerEtab'));
PerEtab.OnCellEnter:=GSCellEnter ;
PerEtab.OnCellExit:=GSCellExit ;
PerEtab.OnRowExit:=GSRowExit ;
PerEtab.OnRowEnter:=GSRowEnter ;
PerEtab.OnEnter:=GSEnter ;
PerEtab.OnElipsisClick:=GSElipsisClick  ;
PerEtab.OnDblClick:=GSElipsisClick ;
PerEtab.GetCellCanvas:= CodeGras;
PerEtab.PostDrawCell:= ColTriangle;
PerEtab.ColCount:=1; i:=0;
S:=LesColonnes ;
Col_Mov:=-1; colEtab :=-1; colDebut:=-1;  colFin:=-1 ; colArrondi:=-1 ;
Repeat
   NomCol:=ReadTokenSt(S) ;
   if NomCol<>'' then
     begin
     if NomCol='MOV' then
       begin
       if i<>0 then PerEtab.ColCount:=PerEtab.ColCount+1;
       Col_Mov:=i; PerEtab.ColWidths[Col_Mov]:=10;
       end
       else if NomCol='GFP_ETABLISSEMENT' then
         begin
         if i<>0 then PerEtab.ColCount:=PerEtab.ColCount+1;
         colEtab:=i; PerEtab.ColWidths[colEtab]:=100; PerEtab.ColLengths[colEtab]:=-1;
         end
         else if NomCol='GFP_DATEDEBUT' then
           begin
           if i<>0 then PerEtab.ColCount:=PerEtab.ColCount+1;
           colDebut:=i; PerEtab.ColWidths[colDebut]:=50;
           end
           else if NomCol='GFP_DATEFIN' then
             begin
             if i<>0 then PerEtab.ColCount:=PerEtab.ColCount+1;
             colFin:=i; PerEtab.ColWidths[colFin]:=50;
             end
             else if NomCol='GFP_ARRONDI' then
               begin
               if i<>0 then PerEtab.ColCount:=PerEtab.ColCount+1;
               colArrondi:=i; PerEtab.ColWidths[colArrondi]:=50;
               end ;
       Inc(i);
       end;
Until ((St='') or (NomCol='')) ;
if Col_Mov<>-1 then PerEtab.FixedCols:=1;
AffecteGrid(PerEtab,taModif) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(PerEtab) ;
End ;

procedure TOF_MBOPerEtablis.OnLoad ;
Begin
ChargePerEtab ;
// Permet d'afficher en gras le code de la première ligne du Grid
PerEtab.SetFocus ;
End ;

procedure TOF_MBOPerEtablis.OnUpdate ;
Var i_ind:Integer ;
Etab:String ;
Begin
TobPerEtab.GetGridDetail(PerEtab,PerEtab.RowCount-1,'les Tarifs',LesColonnes) ;
for i_ind := TobPerEtab.Detail.Count - 1 downto 0 do
  begin
  If (TobPerEtab.Detail[i_ind].GetValue('GFP_DATEDEBUT')=DateDebut) and (TobPerEtab.Detail[i_ind].GetValue('GFP_DATEFIN')=DateFin)
     and (TobPerEtab.Detail[i_ind].GetValue('GFP_ARRONDI')=CodeArrondi)
         then
         begin
         if Etab='' then Etab:=Etab+'"'+TobPerEtab.Detail[i_ind].GetValue('GFP_ETABLISSEMENT')+'"'
          else Etab:=Etab+',"'+TobPerEtab.Detail[i_ind].GetValue('GFP_ETABLISSEMENT')+'"' ;
         TobPerEtab.Detail[i_ind].Free ;
         end ;
  end ;
TobPerEtab.InsertOrUpdateDB ;
ExecuteSQL('DELETE FROM TARIFPER WHERE GFP_ETABLISSEMENT in ('+Etab+')'
          +' and GFP_CODEPERIODE="'+CodePeriode+'"') ;
MiseAJourTarif ;
end ;


procedure TOF_MBOPerEtablis.OnClose ;
Begin
TobPerEtab.free ; TobPerEtab:=nil ;
End ;

{==============================================================================================}
{============================ Chargement et traitement ========================================}
{==============================================================================================}
procedure TOF_MBOPerEtablis.ChargePerEtab ;
var QPeriode: TQuery ;
SQL: String ;
j: Integer ;
Begin
j:=0 ;
SQL:='Select GFP_ETABLISSEMENT, GFP_DATEDEBUT, GFP_DATEFIN, GFP_ARRONDI from TARIFPER Where GFP_ETABLISSEMENT<>"..." And GFP_CODEPERIODE="'+CodePeriode+'"' ;
if ExisteSQL(SQL) then
  begin
  TobPerEtab:= TOB.Create('',NIL,-1);
  QPeriode:=OpenSQL(SQL,True) ;
  While not QPeriode.EOF do
    begin
    TOB.Create ('TARIFPER', TobPerEtab, j) ;
    TobPerEtab.Detail[j].AddChampSup('MOV',False) ;
    TobPerEtab.Detail[j].PutValue('GFP_CODEPERIODE',CodePeriode) ;
    TobPerEtab.Detail[j].PutValue('GFP_ETABLISSEMENT',QPeriode.Findfield('GFP_ETABLISSEMENT').AsString) ;
    TobPerEtab.Detail[j].PutValue('GFP_DATEDEBUT',QPeriode.Findfield('GFP_DATEDEBUT').AsDateTime ) ;
    TobPerEtab.Detail[j].PutValue('GFP_DATEFIN',QPeriode.Findfield('GFP_DATEFIN').AsDateTime ) ;
    TobPerEtab.Detail[j].PutValue('GFP_ARRONDI',QPeriode.Findfield('GFP_ARRONDI').AsString ) ;
    PerEtab.RowCount:=PerEtab.RowCount+1 ;
    j:=j+1 ;
    QPeriode.Next ;
    end ;
    ferme(QPeriode) ;
    if TobPerEtab.Detail.Count>0 then TobPerEtab.PutGridDetail(PerEtab,False,False,LesColonnes,True) ;
  end ;
NouveauPerEtab ;
end ;

procedure TOF_MBOPerEtablis.NouveauPerEtab ;
var QEtab,QPer: TQuery ;
j: Integer ;
TobTemp: Tob ;
Begin
j:=0 ;
QPer:=OpenSQL('SELECT * FROM TARIFPER WHERE GFP_CODEPERIODE="'+CodePeriode+'" and GFP_ETABLISSEMENT="..."',True );
QEtab:=OpenSQL('Select ET_ETABLISSEMENT, ET_ABREGE from ETABLISS',True) ;
if TobPerEtab=nil then TobPerEtab:= TOB.Create('_PERETABLIS',NIL,-1);
While not QEtab.EOF do
    begin
    TobTemp:=TobPerEtab.FindFirst(['GFP_ETABLISSEMENT'],[QEtab.Findfield('ET_ETABLISSEMENT').AsString],False) ;
    if TobTemp= nil then
      begin
      TOB.Create ('TARIFPER', TobPerEtab, j) ;
      TobPerEtab.Detail[j].AddChampSup('MOV',False) ;
      TobPerEtab.Detail[j].PutValue('GFP_CODEPERIODE',CodePeriode) ;
      TobPerEtab.Detail[j].PutValue('GFP_ETABLISSEMENT',QEtab.Findfield('ET_ETABLISSEMENT').AsString) ;
      TobPerEtab.Detail[j].PutValue('GFP_DATEDEBUT',DateDebut) ;
      TobPerEtab.Detail[j].PutValue('GFP_DATEFIN',DateFin) ;
      TobPerEtab.Detail[j].PutValue('GFP_ARRONDI',CodeArrondi) ;
      TobPerEtab.Detail[j].PutValue('GFP_CASCADE',QPer.Findfield('GFP_CASCADE').AsString) ;
      TobPerEtab.Detail[j].PutValue('GFP_PROMO',QPer.Findfield('GFP_PROMO').AsString) ;
      TobPerEtab.Detail[j].PutValue('GFP_LIBELLE',QPer.Findfield('GFP_LIBELLE').AsString) ;
      TobPerEtab.Detail[j].PutValue('GFP_DEMARQUE',QPer.Findfield('GFP_DEMARQUE').AsString) ;
      PerEtab.RowCount:=PerEtab.RowCount+1 ;
      end ;
    QEtab.Next ;
    j:=j+1 ;
    end ;
ferme(QEtab) ;
ferme(QPer) ;
if TobPerEtab.Detail.Count>0 then TobPerEtab.PutGridDetail(PerEtab,False,False,LesColonnes,True) ;
end ;

Procedure TOF_MBOPerEtablis.MiseAJourTarif ;
Var TarfMode,SQLTarif,SQLEtab: String;
TOBTarif,TOBPerEtab,TOBEtab :TOB ;
QTarif,QPerEtab: TQuery ;
i: Integer ;
begin
TOBTarif:=TOB.Create('' ,nil,-1) ;
TarfMode:=RechTarfMode(CodePeriode) ;
if (TarfMode<>'') and (ExisteSQL('Select GF_REMISE From tarif where GF_TARFMODE IN ('+TarfMode+')')) then
  if PGIAsk('Voulez-vous modifier les tarifs de cette période ?', Ecran.Caption) <> mrYes then exit
  else begin
  // Si tarif par etablissement
  TOBPerEtab:=TOB.Create('' ,nil,-1) ;
  SQLEtab:='Select GFP_ETABLISSEMENT,GFP_DATEDEBUT,GFP_DATEFIN,GFP_ARRONDI from TARIFPER Where GFP_ETABLISSEMENT<>"..." And GFP_CODEPERIODE="'+CodePeriode+'"' ;
  QPerEtab:=OpenSQL(SQLEtab,True) ;
  if Not QPerEtab.EOF then
  begin
  TOBPerEtab.LoadDetailDB('TARIFPER','','',QPerEtab,True,False) ;
  // TOBTarif pour traiter tous les cas
  SQLTarif:='Select GF_TARIF,GF_DEPOT,GF_DATEDEBUT,GF_DATEFIN,GF_DEMARQUE,GF_ARRONDI,GF_CASCADEREMISE from TARIF where GF_TARFMODE in ('+TarfMode+') AND GF_DEPOT<>""' ;
  QTarif:=OpenSql(SQLTarif,True) ;
  if Not QTarif.EOF then TobTarif.LoadDetailDB('TARIF','','',QTarif,True,False) ;
  Ferme(QTarif) ;
  // Mise A jour des champs
  for i:=0 to TOBTarif.Detail.count-1 do
    begin
    TOBEtab:=TOBPerEtab.FindFirst(['GFP_ETABLISSEMENT'],[TOBTarif.Detail[i].GetValue('GF_DEPOT')],False) ;
    if TOBEtab<>nil then
      begin
      TOBTarif.Detail[i].PutValue('GF_DATEDEBUT',TOBEtab.GetValue('GFP_DATEDEBUT'));
      TOBTarif.Detail[i].PutValue('GF_DATEFIN',TOBEtab.GetValue('GFP_DATEFIN'));
      TOBTarif.Detail[i].PutValue('GF_ARRONDI',TOBEtab.GetValue('GFP_ARRONDI'));
      end ;
    end ;
  end ;
  TOBTarif.UpdateDB ;
  Ferme(QPerEtab) ;
  end ;
end ;

{==============================================================================================}
{=============================== Evenement du Grid ============================================}
{==============================================================================================}
Procedure TOF_MBOPerEtablis.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
PerEtab.ElipsisButton := ((PerEtab.Col = colDebut) or (PerEtab.Col = colFin) or (PerEtab.Col = colArrondi)) ;
End;

Procedure TOF_MBOPerEtablis.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
if (ACol=ColDebut) and (PerEtab.Cells[ACol,Arow]='') then PerEtab.Cells[ACol,Arow]:=DateToStr(DateDebut) ;
if (ACol=ColFin) and (PerEtab.Cells[ACol,Arow]='') then PerEtab.Cells[ACol,Arow]:=DateToStr(DateFin) ;
End;

procedure TOF_MBOPerEtablis.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
PerEtab.InvalidateRow(Ou);
end ;

procedure TOF_MBOPerEtablis.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
PerEtab.InvalidateRow(Ou);
end ;

procedure TOF_MBOPerEtablis.GSEnter(Sender: TObject);
Var ACol,ARow : integer;
    Temp : Boolean;
begin
if Action=taConsult then Exit ;
ACol:=PerEtab.Col; ARow:=PerEtab.Row;
PerEtab.InvalidateRow(ARow);
GSCellEnter(PerEtab,ACol,ARow,Temp);
end;

procedure TOF_MBOPerEtablis.GSElipsisClick(Sender: TObject);
Var DATE,ARRONDI: THCritMaskEdit;
    Coord : TRect;
begin
     Inherited ;
if (PerEtab.Col = colDebut) or (PerEtab.Col = colFin) then
    BEGIN
    Coord := PerEtab.CellRect (PerEtab.Col, PerEtab.Row);
    DATE := THCritMaskEdit.Create (PerEtab);
    DATE.Parent := PerEtab;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate ;
    V_PGI.ParamDateProc ( DATE) ;
    if DATE.Text <> '' then PerEtab.Cells[PerEtab.Col,PerEtab.Row]:= DATE.Text ;
    DATE.Destroy;
    END;
if (PerEtab.Col = colArrondi) then
    BEGIN
    Coord := PerEtab.CellRect (PerEtab.Col, PerEtab.Row);
    ARRONDI := THCritMaskEdit.Create (PerEtab);
    ARRONDI.Parent := PerEtab;
    ARRONDI.Top := Coord.Top;
    ARRONDI.left := Coord.Left;
    ARRONDI.Width := 3; ARRONDI.Visible := False;
    ARRONDI.OpeType:=otString ;
    ARRONDI.DATATYPE:='GCCODEARRONDI' ;
    LookUpCombo(ARRONDI) ;
    if ARRONDI.Text <> '' then PerEtab.Cells[PerEtab.Col,PerEtab.Row]:= ARRONDI.Text ;
    ARRONDI.Destroy;
    END;
END;

{==============================================================================================}
{=============================== Actions liées au grid ========================================}
{==============================================================================================}
procedure TOF_MBOPerEtablis.ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas;
                                         AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < PerEtab.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = Col_Mov) then
  begin
  Arect:=PerEtab.CellRect(Acol,Arow) ;
  Canvas.Brush.Color := PerEtab.FixedColor;
  Canvas.FillRect(ARect);
    if (ARow = PerEtab.row) then
       begin
       Canvas.Brush.Color := clBlack ;
       Canvas.Pen.Color := clBlack ;
       Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
       Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
       Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
       if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
       end ;
  end;
end;

procedure TOF_MBOPerEtablis.CodeGras ( ACol,ARow : Longint; Canvas : TCanvas;
                                      AState: TGridDrawState);
begin
if (ACol = colEtab) and (ARow>0) then
  begin
  Canvas.Font.Style := [fsBold];
  end;
end;


Initialization
registerclasses([TOF_MBOPerEtablis]);
end.

