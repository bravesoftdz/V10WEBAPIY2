unit UTofTarifSaisiBase;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,TarifUtil,DBTables,M3FP,Fe_Main,HDB,UtilArticle,Paramsoc;

Type
    TOF_TarifSaisiBase = Class(TOF)
     private
        TarifBase : THGRID ;
        TobBase,TobBaseLigne,TOBTarif,TobTarifDim: Tob ;
        Col_Mov, colType,colDev, colPrix,  colArrondi, colCreer: Integer ;
        LesColonnes,CodeArticle,Etat: String ;
        TarifCreer: Boolean ;

        procedure GSElipsisClick(Sender: TObject) ;
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSEnter(Sender: TObject);
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure ChargeTarifBase ;
        procedure NouveauTarifBase ;
        procedure MAJPrixArrondi(Coef: Double) ;
        procedure FormateColSaisie (ACol,ARow : Longint ) ;
        procedure FormatePrix ;

     public
         Action   : TActionFiche ;
         DEV       : RDEVISE ;
         Procedure OnArgument(Arguments:string) ; override ;
         Procedure OnLoad ; override ;
         Procedure OnUpdate ; override ;
         Procedure OnClose ; override ;

         // MAJ Table
         Procedure MAJTarif(Arow:Integer; CodeTarifMode:String) ;
         Procedure MAJTarifDim ;
         Function  MAJTarifMode(Arow:Integer): String ;
         Procedure ValideTarif ;

         procedure ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
         procedure CodeGras ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);

    END ;
    
implementation

Procedure TOF_TarifSaisiBase.OnArgument(Arguments:string) ;
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
       if ChampMul='CodeArticle' then CodeArticle:=ValMul ;
       if ChampMul='Etat' then Etat:=ValMul ;
       end ;
    end ;
until  Critere='' ;
Ecran.Caption:='Tarif de base de l''article: '+CodeArticle ;
LesColonnes:='MOV;_TYPE;_DEV;_PRIX;_ARRONDI;_CREER' ;
TarifBase:=THGRID(GetControl('TARIFBASE'));
TarifBase.OnCellEnter:=GSCellEnter ;
TarifBase.OnCellExit:=GSCellExit ;
TarifBase.OnRowExit:=GSRowExit ;
TarifBase.OnRowEnter:=GSRowEnter ;
TarifBase.OnEnter:=GSEnter ;
TarifBase.OnElipsisClick:=GSElipsisClick  ;
TarifBase.OnDblClick:=GSElipsisClick ;
//TarifBase.OnClick:=GCBoiteCoche ;
TarifBase.GetCellCanvas:= CodeGras;
TarifBase.PostDrawCell:= ColTriangle;
TarifBase.ColCount:=1; i:=0;
S:=LesColonnes ;
Col_Mov:=-1; colType :=-1; colPrix:=-1;  colArrondi:=-1 ;
Repeat
   NomCol:=ReadTokenSt(S) ;
   if NomCol<>'' then
     begin
     if NomCol='MOV' then
       begin
       if i<>0 then TarifBase.ColCount:=TarifBase.ColCount+1;
       Col_Mov:=i; TarifBase.ColWidths[Col_Mov]:=10;
       end
       else if NomCol='_TYPE' then
         begin
         if i<>0 then TarifBase.ColCount:=TarifBase.ColCount+1;
         colType:=i; TarifBase.ColWidths[colType]:=150; TarifBase.ColLengths[colType]:=-1;
         end
         else if NomCol='_DEV' then
           begin
           if i<>0 then TarifBase.ColCount:=TarifBase.ColCount+1;
           colDev:=i; TarifBase.ColWidths[colDev]:=50; TarifBase.ColLengths[colDev]:=-1;
           end
           else if NomCol='_PRIX' then
             begin
             if i<>0 then TarifBase.ColCount:=TarifBase.ColCount+1;
             colPrix:=i; TarifBase.ColWidths[colPrix]:=50;
             TarifBase.ColLengths[colPrix]:=30;
             end
             else if NomCol='_ARRONDI' then
               begin
               if i<>0 then TarifBase.colCount:=TarifBase.ColCount+1;
               colArrondi:=i ;
               TarifBase.ColWidths[colArrondi]:=50;
               //TarifBase.ColFormats[colArrondi]:='CB=GCCODEARRONDI||' ;
               end
                else if NomCol='_CREER' then
                 begin
                 if i<>0 then TarifBase.colCount:=TarifBase.ColCount+1;
                 colCreer:=i ;
                 TarifBase.ColWidths[colCreer]:=150;
                 TarifBase.ColFormats[colCreer]:='CB=PGOUINON|''|' ;
                 TarifBase.ColLengths[colCreer]:=1;
                 end ;
       Inc(i);
       end;
Until ((St='') or (NomCol='')) ;
if Col_Mov<>-1 then TarifBase.FixedCols:=1;

AffecteGrid(TarifBase,taModif) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(TarifBase) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
End ;

procedure TOF_TarifSaisiBase.OnLoad ;
Begin
//if GetParamSoc('SO_GCPERBASETARIF')='' then begin TFVierge(Ecran).close; exit ; end ;
TarifCreer:=False ;
if Etat='nouveau' then NouveauTarifBase else ChargeTarifBase ;
// Permet d'afficher en gras le code de la première ligne du Grid
TarifBase.SetFocus ;
FormatePrix ;
TobTarif:=TOB.Create('TARIF',NIL,-1) ;
TobTarifDim:=TOB.Create('TARIF',NIL,-1) ;
End ;

procedure TOF_TarifSaisiBase.OnUpdate ;
Var CodeTarifMode: String ;
i:Integer ;
Begin
// Vérifier si tarif existant ou pas => modif ou création
TOBBase.GetGridDetail(TarifBase,TarifBase.RowCount-1,'les Tarifs',LesColonnes) ;
TarifCreer:=True ;
for i:=1 to TarifBase.rowcount-1 do
  begin
  if TarifBase.CellValues[ColCREER,i]='OUI' then
    begin
    If (TarifBase.Cells[ColPrix,i]='') or (TarifBase.Cells[ColPrix,i]='0') then continue ;
    CodeTarifMode:=MAJTarifMode(i) ;
    MAJTarif(i,CodeTarifMode) ;
    ValideTarif ;
    end ;
  end ;
end ;

procedure TOF_TarifSaisiBase.OnClose ;
Begin
if not TarifCreer then
   if HShowMessage('1;Tarif de base;Confirmer vous l''abandon de création des tarifs?;Q;YNC;N;C;','','')<>mrYes then
   begin
   LastError:=-1;
   exit;
   end ;
TobBase.free ; TobBase:=nil;
TOBTarif.free ; TobTarif:=nil;
TOBTarifDim.free ; TobTarifDim:=nil;
End ;

{==============================================================================================}
{=============================== MAJ des tables ===============================================}
{==============================================================================================}
procedure TOF_TarifSaisiBase.MAJTarif(Arow:Integer;CodeTarifMode:String) ;
Var Q:TQuery ;
Libelle: String ;
MaxTarif: Integer ;
begin
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
Q:=OpenSQL('Select * from TarifMode where GFM_TARFMODE="'+CodeTarifMode+'"',True) ;
if not Q.EOF then
  begin
  //TobTarif.PutValue('GF_TARIF',MaxTarif) ;
  if TobBase.Detail[Arow-1].GetValue('_CODETARF')=0 then TobTarif.PutValue('GF_TARIF',MaxTarif)
     else TobTarif.PutValue('GF_TARIF',TobBase.Detail[Arow-1].GetValue('_CODETARF')) ;
  TobTarif.PutValue('GF_TARFMODE',CodeTarifMode) ;
  TobTarif.PutValue('GF_ARTICLE',CodeArticle) ;
  //TobTarif.PutValue('GF_DEVISE',Q.FindField('GFM_DEVISE').AsString ) ;
  TobTarif.PutValue('GF_DEVISE',TobBase.Detail[Arow-1].GetValue('_DEV') ) ;
  TobTarif.PutValue('GF_DATEDEBUT',Q.FindField('GFM_DATEDEBUT').AsDateTime) ;
  TobTarif.PutValue('GF_DATEFIN',Q.FindField('GFM_DATEFIN').AsDateTime) ;
  //TobTarif.PutValue('GF_PRIXUNITAIRE',StrToFloat(TarifBase.Cells[ColPrix,ARow])) ;
  TobTarif.PutValue('GF_PRIXUNITAIRE',TobBase.Detail[Arow-1].GetValue('_PRIX')) ;
  TobTarif.PutValue('GF_BORNEINF', -999999);
  TobTarif.PutValue('GF_BORNESUP', 999999);
  TobTarif.PutValue('GF_QUANTITATIF', '-');
  TobTarif.PutValue('GF_REMISE', 0);
  //TobTarif.PutValue('GF_ARRONDI',Q.FindField('GFM_ARRONDI').AsString) ;
  //TobTarif.Putvalue('GF_ARRONDI',TarifBase.Cells[ColArrondi,ARow]) ;
  TobTarif.Putvalue('GF_ARRONDI',TobBase.Detail[Arow-1].GetValue('_ARRONDI')) ;
  TobTarif.PutValue('GF_DEMARQUE',Q.FindField('GFM_DEMARQUE').AsString) ;
  TobTarif.PutValue('GF_CALCULREMISE', '');
  TobTarif.PutValue('GF_MODECREATION', 'MAN');
  TobTarif.PutValue('GF_CASCADEREMISE', 'MIE');
  TobTarif.PutValue('GF_QUALIFPRIX', 'GRP');
  TobTarif.PutValue('GF_REGIMEPRIX', 'TTC') ;
  TobTarif.PutValue('GF_DEPOT',Q.FindField('GFM_ETABLISREF').AsString) ;
  TobTarif.PutValue('GF_SOCIETE',Q.FindField('GFM_ETABLISREF').AsString) ;
  libelle:=RechDom('GCTARIFTYPE1',Q.FindField('GFM_TYPETARIF').AsString,False)+' - '+RechDom('GCTARIFPERIODE1',Q.FindField('GFM_PERTARIF').AsString,False) ;
  TobTarif.PutValue('GF_LIBELLE',libelle) ;
  end ;
Ferme(Q) ;
if LaTob.GetValue('GA_DIMMASQUE')<>'' then MAJTarifDim ;
end ;

procedure TOF_TarifSaisiBase.MAJTarifDim ; // AC
var MaxTarif,j: Integer ;
QArtDim: TQuery ;
begin
j:=0 ;
TobTarifDim:= TOB.Create ('', nil, -1) ;
MaxTarif:=TOBTarif.GetValue ('GF_TARIF') + 1;// + TOBTarfArt.Detail.count;
begin
QArtDim:=OpenSql('Select GA_ARTICLE from Article where GA_CODEARTICLE="'+TRIM(copy(CodeArticle,1,18))+'" And GA_STATUTART="DIM"', True) ;
While not QArtDim.EOF do
 Begin
 TOB.Create ('TARIF', TobTarifDim, j) ;
 TobTarifDim.Detail[j].Dupliquer (TOBTarif, False, True);
 TobTarifDim.Detail[j].PutValue('GF_ARTICLE',QArtDim.FindField('GA_ARTICLE').AsString) ;
 TobTarifDim.Detail[j].PutValue('GF_TARIF',MaxTarif) ;
 QArtDim.next ;
 MaxTarif:=MaxTarif+1 ;
 j:=j+1 ;
 End ;
end ;
ferme(QArtDim) ;
end;

Function TOF_TarifSaisiBase.MAJTarifMode(Arow:Integer):String ;
var Q: TQuery ;
    CodeMaxInt: Integer ;
    TobType,TobPer,TobTarifMode: TOB ;
    CodeMax: String ;
begin
CodeMaxInt:=1 ;
TobTarifMode:=TOB.Create('TARIFMODE',NIL,-1) ;
TobType:=TOB.Create('TARIFTYPMODE',Nil,-1) ;
TobPer:=TOB.Create('TARIFPER',Nil,-1) ;
TOBType.SelectDB('"' + Copy(TarifBase.Cells[ColType,ARow],1,3) + '"',nil) ;
TobPer.SelectDB('"' + GetParamSoc('SO_GCPERBASETARIF')+ '"',nil) ;
Q:=OpenSql('Select * from TARIFMODE where GFM_TYPETARIF="'+Copy(TarifBase.Cells[ColType,ARow],1,3)+'" and GFM_PERTARIF="'+GetParamSoc('SO_GCPERBASETARIF')+'"',True) ;
if not Q.EOF then
  begin                                                                          
  TobTarifMode.SelectDB('',Q) ;
  Result:=TobTarifMode.GetValue('GFM_TARFMODE') ;
  ferme(Q) ;
  end else
  begin
  Q := OpenSQL ('SELECT MAX(GFM_TARFMODE) FROM TARIFMODE',True);
  if Not Q.EOF then
    begin
    CodeMax:=Q.Fields[0].AsString;
    if CodeMax <> '' then CodeMaxInt := StrToInt(CodeMax)+1;
    ferme(Q) ;
    end ;
  Result:=IntToStr(CodeMaxInt) ;
  TobTarifMode.PutValue('GFM_TARFMODE', IntToStr(CodeMaxInt));
  TobTarifMode.PutValue('GFM_TYPETARIF', TobType.GetValue('GFT_CODETYPE'));
  TobTarifMode.PutValue('GFM_PERTARIF', TobPer.GetValue('GFP_CODEPERIODE'));
  TobTarifMode.PutValue('GFM_LIBELLE', TobType.GetValue('GFT_LIBELLE')+'-'+TobPer.GetValue('GFP_LIBELLE'));
  TobTarifMode.PutValue('GFM_DATEDEBUT', TobPer.GetValue('GFP_DATEDEBUT'));
  TobTarifMode.PutValue('GFM_DATEFIN', TobPer.GetValue('GFP_DATEFIN'));
  TobTarifMode.PutValue('GFM_DEMARQUE', TobPer.GetValue('GFP_DEMARQUE'));
  TobTarifMode.PutValue('GFM_ARRONDI', TobPer.GetValue('GFP_ARRONDI'));
  TobTarifMode.PutValue('GFM_TARIFBASE', TobPer.GetValue('GFP_TARIFBASE'));
  TobTarifMode.PutValue('GFM_COEFFICIENT', TobType.GetValue('GFT_COEFFICIENT'));
  TobTarifMode.PutValue('GFM_DEVISE', TobType.GetValue('GFT_DEVISE'));
  TobTarifMode.PutValue('GFM_ETABLISREF', TobType.GetValue('GFT_ETABLISREF'));
  TobTarifMode.InsertOrUpdateDB ;
  end ;
TobType.Free; TobType:=Nil ;
TobPer.Free; TobPer:=Nil ;
TobTarifMode.Free; TobTarifMode:=Nil
end ;

Procedure TOF_TarifSaisiBase.ValideTarif ;
Var Erreur: Boolean ;
begin
TOBTarif.InsertOrUpdateDB(False) ;
if LaTob.GetValue('GA_DIMMASQUE')<>'' then TOBTarifDim.InsertOrUpdateDB(False) ;
end ;


{==============================================================================================}
{=============================== Evenement du Grid ========================================}
{==============================================================================================}
Procedure TOF_TarifSaisiBase.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
TarifBase.ElipsisButton := (TarifBase.Col = colArrondi) ;
//If TarifBase.Col=colCreer then BoiteCoche(Sender) ;
End;

Procedure TOF_TarifSaisiBase.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var ValArrondi: String ;
Prix: Double ;
Begin
If (ACol=ColPrix) or (ACol=ColArrondi) then TarifBase.CellValues[ColCREER,ARow]:='OUI' ;
Prix:=Valeur(TarifBase.Cells[ColPrix,Arow]);
ValArrondi:=TarifBase.CellValues [ColArrondi,Arow] ;
If ACol=ColPrix then TarifBase.Cells[ACol,Arow]:=FloatToStr(ArrondirPrix(ValArrondi, Prix)) ;
If ACol=ColArrondi then TarifBase.Cells[ColPrix,Arow]:=FloatToStr(ArrondirPrix(ValArrondi, Prix)) ;
DEV.Code :=TOBBase.Detail[ARow-1].GetValue('_DEV') ;
GetInfosDevise (DEV) ;
FormateColSaisie(Acol, Arow) ;
End;

procedure TOF_TarifSaisiBase.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
TarifBase.InvalidateRow(Ou);
end ;

procedure TOF_TarifSaisiBase.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
TarifBase.InvalidateRow(Ou);
//BoiteCoche(Sender) ;
end ;

procedure TOF_TarifSaisiBase.GSEnter(Sender: TObject);
Var ACol,ARow : integer;
    Temp : Boolean;
begin
if Action=taConsult then Exit ;
ACol:=TarifBase.Col; ARow:=TarifBase.Row;
//CreerBoite (Sender) ;
TarifBase.InvalidateRow(ARow);
GSCellEnter(TarifBase,ACol,ARow,Temp);
end;

procedure TOF_TarifSaisiBase.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
end ;

procedure TOF_TarifSaisiBase.GSElipsisClick(Sender: TObject);
Var ARRONDI: THCritMaskEdit;
    Coord : TRect;
begin
     Inherited ;
if (TarifBase.Col = colArrondi) then
    BEGIN
    Coord := TarifBase.CellRect (TarifBase.Col, TarifBase.Row);
    ARRONDI := THCritMaskEdit.Create (TarifBase);
    ARRONDI.Parent := TarifBase;
    ARRONDI.Top := Coord.Top;
    ARRONDI.left := Coord.Left;
    ARRONDI.Width := 3; ARRONDI.Visible := False;
    ARRONDI.OpeType:=otString ;
    ARRONDI.DATATYPE:='GCCODEARRONDI' ;
    LookUpCombo(ARRONDI) ;
    if ARRONDI.Text <> '' then TarifBase.Cells[TarifBase.Col,TarifBase.Row]:= ARRONDI.Text ;
    ARRONDI.Destroy;
    END;
END;

{==============================================================================================}
{=============================== Chargement des Tarifs de base ================================}
{==============================================================================================}
procedure TOF_TarifSaisiBase.NouveauTarifBase ;
var QType: TQuery ;
Coef: Double ;
i: Integer ;
Exist: Boolean ;
Begin
Exist:=False ;
if TOBBase=nil then TOBBase:= TOB.Create('les Tarifs',NIL,-1);
QType:=OpenSql('Select * from TARIFTYPMODE ',True) ;
if not QType.EOF then
   begin
   While not QType.EOF do
     Begin
       for i:=0 to TOBBase.Detail.Count-1 do
         begin
         if TOBBase.Detail[i].GetValue('_TYPE')=QType.Findfield('GFT_CODETYPE').AsString+'-'+QType.Findfield('GFT_LIBELLE').AsString then
           begin
           Exist:=True ;
           end ;
         end ;
         if not Exist then
           begin
           TOBBaseLigne:=Tob.create('un tarif',TOBBase,-1);
           TOBBaseLigne.AddChampSup('MOV',False) ;
           TOBBaseLigne.AddChampSup('_TYPE',False) ;
           TOBBaseLigne.PutValue('_TYPE',QType.Findfield('GFT_CODETYPE').AsString+'-'+QType.Findfield('GFT_LIBELLE').AsString);
           TOBBaseLigne.AddChampSup('_DEV',False) ;
           TOBBaseLigne.PutValue('_DEV',QType.Findfield('GFT_DEVISE').AsString) ;
           DEV.Code :=QType.Findfield('GFT_DEVISE').AsString ;
           GetInfosDevise (DEV) ;
           TOBBaseLigne.AddChampSup('_CREER',False) ;
           TOBBaseLigne.PutValue('_CREER','OUI') ;
           // Recup des prix et arrondi
           TOBBaseLigne.AddChampSup('_PRIX',False) ;
           TOBBaseLigne.AddChampSup('_ARRONDI',False) ;
           MAJPrixArrondi(QType.FindField('GFT_COEFFICIENT').AsFloat) ;
           // init Codetarif existant
           TOBBaseLigne.AddChampSup('_CODETARF',False) ;
           TOBBaseLigne.PutValue('_CODETARF',0) ;
           TarifBase.RowCount:=TarifBase.RowCount+1 ;
           end ;
         QType.next ;
         Exist:=False;
     End ;
     ferme(QType) ;
   end ;
if TOBBase.Detail.Count>0 then TOBBase.PutGridDetail(TarifBase,False,False,LesColonnes,True);
end ;

procedure TOF_TarifSaisiBase.ChargeTarifBase ;
var QType,QTarifMode,QTarif: TQuery ;
Coef: Double ;
SQL: String ;
Begin
SQL:='Select GF_TARIF, GF_TARFMODE, GF_PRIXUNITAIRE, GF_REMISE, GF_ARRONDI, GF_DEVISE from TARIF Where GF_ARTICLE="'+LaTob.GetVAlue('GA_ARTICLE')+'"' ;
if ExisteSQL(SQL) then
  begin
  TOBBase:= TOB.Create('les Tarifs',NIL,-1) ;
  QTarif:=OpenSQL(SQL,True) ;
  if not QTarif.EOF then
    begin
    While not QTarif.EOF do
      begin
      if QTarif.FindField('GF_TARFMODE').AsString<>'0' then
      begin
        QTarifMode:=OpenSql('Select * from TARIFMODE where GFM_TARFMODE='+QTarif.FindField('GF_TARFMODE').AsString ,True) ;
        if not QTarifMode.EOF then
           begin
             if QTarifMode.FindField('GFM_PERTARIF').AsString=GetParamSoc('SO_GCPERBASETARIF') then
               begin
               TOBBaseLigne:=Tob.create('un tarif',TOBBase,-1);
               TOBBaseLigne.AddChampSup('MOV',False) ;
               TOBBaseLigne.AddChampSup('_TYPE',False) ;
               TOBBaseLigne.PutValue('_TYPE',QTarifMode.Findfield('GFM_TYPETARIF').AsString+'-'+RechDom('GCTARIFTYPE1',QTarifMode.Findfield('GFM_TYPETARIF').AsString,False));
               TOBBaseLigne.AddChampSup('_DEV',False) ;
               TOBBaseLigne.PutValue('_DEV',QTarif.Findfield('GF_DEVISE').AsString) ;
               DEV.Code :=QTarif.Findfield('GF_DEVISE').AsString ;
               GetInfosDevise (DEV) ;
               TOBBaseLigne.AddChampSup('_CREER',False) ;
               TOBBaseLigne.PutValue('_CREER','NON') ;
               TOBBaseLigne.AddChampSup('_PRIX',False) ;
               TOBBaseLigne.AddChampSup('_ARRONDI',False) ;
               TOBBaseLigne.PutValue('_PRIX',QTarif.Findfield('GF_PRIXUNITAIRE').AsFloat) ;
               TOBBaseLigne.PutValue('_ARRONDI',QTarif.Findfield('GF_ARRONDI').AsString) ;
               // Recup CodeTarif
               TOBBaseLigne.AddChampSup('_CODETARF',False) ;
               TOBBaseLigne.PutValue('_CODETARF',QTarif.Findfield('GF_TARIF').AsFloat) ;
               TarifBase.RowCount:=TarifBase.RowCount+1 ;
               end ;
           End ;
           ferme(QTarifMode) ;
           QTarif.Next ;
         end else
         QTarif.Next ;
       end ;
      ferme(QTarif) ;
     end ;
    if TOBBase.Detail.Count>0 then TOBBase.PutGridDetail(TarifBase,False,False,LesColonnes,True) ;
  end ;
NouveauTarifBase ;
end ;

procedure TOF_TarifSaisiBase.MAJPrixArrondi(Coef: Double) ;
var PrixArticle, Prix: Double;
ValArrondi: String ;
begin
if LaTob<>nil then
  begin
  CodeArticle:=LaTob.GetValue('GA_ARTICLE') ;
  ValArrondi:=LaTob.GetValue('GA_ARRONDIPRIXTTC') ;
  PrixArticle:=LaTob.GetValue('GA_PVTTC') ;
  if Coef<>0 then Prix:=PrixArticle*Coef else Prix:=PrixArticle ;
  Prix:=ArrondirPrix(ValArrondi,Prix) ;
  TOBBaseLigne.PutValue('_PRIX',Prix) ;
  TOBBaseLigne.PutValue('_ARRONDI',ValArrondi) ;
  end ;
end ;
{==============================================================================================}
{=============================== Actions liées au grid ========================================}
{==============================================================================================}
procedure TOF_TarifSaisiBase.FormateColSaisie (ACol,ARow : Longint ) ;
Var st, Stc:String ;
Begin
St:=TarifBase.Cells[ACol,ARow] ; StC:=St ;
if ACol=ColPrix then StC:=StrF00(Valeur(St),DEV.Decimale) ;
TarifBase.Cells[ACol,ARow]:=StC ;
end ;

procedure TOF_TarifSaisiBase.FormatePrix ;
var i:Integer ;
begin
for i:=1 to TarifBase.RowCount-1 do
  begin
  DEV.Code :=TOBBase.Detail[i-1].GetValue('_DEV') ;
  GetInfosDevise (DEV) ;
  FormateColSaisie(ColPrix, i) ;
  end ;
end ;

procedure TOF_TarifSaisiBase.ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas;
                                         AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < TarifBase.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = Col_Mov) then
  begin
  Arect:=TarifBase.CellRect(Acol,Arow) ;
  Canvas.Brush.Color := TarifBase.FixedColor;
  Canvas.FillRect(ARect);
    if (ARow = TarifBase.row) then
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

procedure TOF_TarifSaisiBase.CodeGras ( ACol,ARow : Longint; Canvas : TCanvas;
                                      AState: TGridDrawState);
begin
if (ACol = colType) and (ARow>0) then
  begin
  Canvas.Font.Style := [fsBold];
  end;
end;


Initialization
registerclasses([TOF_TarifSaisiBase]);
end.
 