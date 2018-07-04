unit AnalysHT;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Grids, ExtCtrls, ComCtrls, Menus, hmsgbox, Buttons,
  HSysMenu, Ent1, Hent1, Filtre, SaisUtil, Mask, HDB, DB, {$IFNDEF DBXPRESS}dbtables,
  HTB97{$ELSE}uDbxDataSet{$ENDIF}, HStatus,
  ParamDat, DetaTva, HTB97 ;

Procedure AnalyseMvtHT ;

type
  TFAnalysHT = class(TForm)
    HMTrad: THSystemMenu;
    PanelBouton: TPanel;
    BReduire: TBitBtn;
    BAgrandir: TBitBtn;
    BRechercher: TBitBtn;
    Panel1: TPanel;
    BImprimer: TBitBtn;
    BOuvrir: TBitBtn;
    BFerme: TBitBtn;
    BAide: TBitBtn;
    BStop: TBitBtn;
    FindDialog: TFindDialog;
    HM: THMsgBox;
    Pages: TPageControl;
    ChoixEtat: TTabSheet;
    Bevel1: TBevel;
    RgChoix: TRadioGroup;
    FListe: THGrid;
    PFiltres: TPanel;
    BCherche: TBitBtn;
    FFiltres: THValComboBox;
    Exercice: THValComboBox;
    TExercice: TLabel;
    TEtablissement: TLabel;
    Etablissement: THValComboBox;
    TDateCpta: THLabel;
    DateCpta: THCritMaskEdit;
    TTDateCpta_: THLabel;
    DateCpta_: THCritMaskEdit;
    TDateEche: THLabel;
    DateEche: THCritMaskEdit;
    TTDateEche: THLabel;
    DateEche_: THCritMaskEdit;
    TDevise: TLabel;
    Devise: THValComboBox;
    QEcr: TQuery;
    CRegime: THValComboBox;
    CTva: THValComboBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    procedure BFermeClick(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ExerciceChange(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure DateCptaKeyPress(Sender: TObject; var Key: Char);
    procedure BImprimerClick(Sender: TObject);
  private
    FNomFiltre : String ;
    NatJal,UnWhere,LesDateEche : String ;
    LNumPiece,LGenCol,LGenChaPro : TStringList ;
    WMinX,WMinY : Integer ;
    LPiece : TList ;
    Stop : Boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas) ;
    Procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas) ;
    Procedure InitCriteres ;
    Procedure TitresColRow ;
    Function  UnWhereSupp : String ;
    Procedure InitLeWhere ;
    Procedure RempliLNumPiece ;
    Procedure RempliLGenCol ;
    Procedure RempliLGenChaPro ;
    Function  ChercheTauxTva(CodeTva,Regime : String) : Double ;
    Function  CreerLobjet(Q : TQuery ; Typ,Col,Row : Integer) : TAnalTva ;
    Procedure CopieLobjet(X,Y : TAnalTva) ;
    Procedure ChercheLesMvts ;
    Function  EstCpteTTC(StGen : String) : Boolean ;
    Function  EstCpteHT(StGen : String) : Boolean ;
    Function  OnTraiteLaPiece(St1 : String) : Boolean ;
    Function  DateEcheDansCrit(D : TDateTime) : Boolean ;
    Function  QuelCol(Regtva : String) : Integer ;
    Function  QuelRow(Tauxtva : String ; TypLig : Integer) : Integer ;
    Procedure TrouveAColARow(TypLig : Integer ; Var ACol,ARow : Integer) ;
    Function  TrouveLaSomme : Double ;
    Procedure FaitTotaux ;
    Procedure ProratLesSommes(TotTTC,TTC : Double ; Prorat : Boolean) ;
    Procedure NettoieLaGrille ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses PrintDBG ;

Procedure AnalyseMvtHT ;
var FAnalysHT : TFAnalysHT ;
BEGIN
FAnalysHT:=TFAnalysHT.Create(Application) ;
  Try
   FAnalysHT.FNomFiltre:='ANALYSHT' ;
   FAnalysHT.ShowModal ;
  Finally
   FAnalysHT.Free ;
  End ;
SourisNormale ;
END ;

procedure TFAnalysHT.FormCreate(Sender: TObject);
begin
LNumPiece:=TStringList.Create ; LGenCol:=TStringList.Create ; LGenChaPro:=TStringList.Create ;
LPiece:=TList.Create ; NatJal:='' ; UnWhere:='' ; LesDateEche:='' ;
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFAnalysHT.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFAnalysHT.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFAnalysHT.FormClose(Sender: TObject; var Action: TCloseAction);
begin
LNumPiece.Free ; LGenCol.Free ; LGenChaPro.Free ; NettoieLaGrille ;
VideListe(LPiece) ; LPiece.Free ;
end;

procedure TFAnalysHT.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFAnalysHT.BCreerFiltreClick(Sender: TObject);
begin NewFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFAnalysHT.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFAnalysHT.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre(FNomFiltre,FFiltres) ; end;

procedure TFAnalysHT.BRenFiltreClick(Sender: TObject);
begin RenameFiltre(FNomFiltre,FFiltres) ; end;

procedure TFAnalysHT.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres,Pages) ; end;

procedure TFAnalysHT.FFiltresChange(Sender: TObject);
begin LoadFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TFAnalysHT.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFAnalysHT.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFAnalysHT.FormShow(Sender: TObject);
begin
ChargeFiltre(FNomFiltre,FFiltres,Pages) ; InitCriteres ; TitresColRow ;
FListe.GetCellCanvas:=GetCellCanvas ; FListe.PostDrawCell:=PostDrawCell ;
end;

procedure TFAnalysHT.DateCptaKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFAnalysHT.BImprimerClick(Sender: TObject);
begin PrintDBGrid(FListe,Nil,Caption,'') ; end;

// *****************************************************************************

Function TFAnalysHT.CreerLobjet(Q : TQuery ; Typ,Col,Row : Integer) : TAnalTva ;
Var X : TAnalTva ;
BEGIN
X:=TAnalTva.Create ;
With X do
  BEGIN
  Journal:=Q.FindField('E_JOURNAL').AsString ; NumPiece:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ; CompteGen:=Q.FindField('E_GENERAL').AsString ;
  CompteAux:=Q.FindField('E_AUXILIAIRE').AsString ; Deb:=Q.FindField('E_DEBIT').AsFloat ;
  Cred:=Q.FindField('E_CREDIT').AsFloat ;  RegTva:=Q.FindField('E_REGIMETVA').AsString ;
  CodTva:=Q.FindField('E_TVA').AsString ; Couverture:=Q.FindField('E_COUVERTURE').AsFloat ;
  DebDev:=Q.FindField('E_DEBITDEV').AsFloat ; CredDev:=Q.FindField('E_CREDITDEV').AsFloat ;
  NumEche:=Q.FindField('E_NUMECHE').AsInteger ; DatCompta:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  TypCpte:=Typ ; ACol:=Col ; ARow:=Row ; SommeTotal:=TrouveLaSomme ;
  END ;
Result:=X ;
END ;

Procedure TFAnalysHT.CopieLobjet(X,Y : TAnalTva) ;
BEGIN
With X do
   BEGIN
   TypCpte:=Y.TypCpte ; ACol:=Y.ACol ; ARow:=Y.ARow ; Journal:=Y.Journal ;
   NumPiece:=Y.NumPiece ; NumLigne:=Y.NumLigne ; CompteGen:=Y.CompteGen ; CompteAux:=Y.CompteAux ;
   Deb:=Y.Deb ; Cred:=Y.Cred ; RegTva:=Y.RegTva ; CodTva:=Y.CodTva ; Couverture:=Y.Couverture ;
   DebDev:=Y.DebDev ; CredDev:=Y.CredDev ; NumEche:=Y.NumEche ; SommeTotal:=Y.SommeTotal ;
   DatCompta:=Y.DatCompta ;
   END ;
END ;

Procedure TFAnalysHT.NettoieLaGrille ;
Var i,j : Integer ;
BEGIN
for i:=1 to FListe.RowCount-1 do
    for j:=1 to FListe.ColCount-1 do
        BEGIN
        if FListe.Objects[j,i]<>Nil then
           BEGIN
           VideListe(TList(FListe.Objects[j,i])) ;
           TObject(FListe.Objects[j,i]).Free ;
           END ;
        FListe.Cells[j,i]:='' ; FListe.Objects[j,i]:=Nil ;
        END ;
FListe.RowCount:=FListe.FixedRows+1 ;
END ;

procedure TFAnalysHT.ExerciceChange(Sender: TObject);
begin
if Exercice.Value='' then BEGIN DateCpta.Text:=StDate1900 ; DateCpta_.Text:=StDate2099 ; END
                     else ExoToDates(Exercice.Value,DateCpta,DateCpta_) ;
end;

procedure TFAnalysHT.InitCriteres ;
Var D1,D2 : TDateTime ;
BEGIN
Exercice.ItemIndex:=0 ; Etablissement.ItemIndex:=0 ; Devise.ItemIndex:=0 ;
DateCpta.Text:=StDate1900 ; DateCpta_.Text:=StDate2099 ;
DateEche_.Text:=DateToStr(V_PGI.DateEntree) ;
QuelDateExo(V_PGI.DateEntree,D1,D2) ; DateEche.Text:=DateToStr(D1) ;
END ;

Procedure TFAnalysHT.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas) ;
BEGIN
if ACol=0 then
   BEGIN
   if (CTva.Items.IndexOf(FListe.Cells[0,ARow])>=0) Or (FListe.Cells[0,ARow]=HM.Mess[4]) Or
      (FListe.Cells[0,ARow]=HM.Mess[5]) Or (FListe.Cells[0,ARow]=HM.Mess[6]) Or
      (FListe.Cells[0,ARow]=HM.Mess[7]) Or (FListe.Cells[0,ARow]=HM.Mess[0])then
      FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsBold] ;
   END ;
END ;

Procedure TFAnalysHT.PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas) ;
Var R : TRect ;
BEGIN
if ((CTva.Items.IndexOf(FListe.Cells[0,ARow])>=0) Or (FListe.Cells[0,ARow]=HM.Mess[7])) And
   (ACol<>0) then
    BEGIN
    Canvas.Brush.Color := Fliste.FixedColor ; ;
    Canvas.Brush.Style := bsBDiagonal ;
    Canvas.Pen.Color   := Fliste.FixedColor ;
    Canvas.Pen.Mode    := pmCopy ;
    Canvas.Pen.Style   := psClear ;
    Canvas.Pen.Width   := 1 ;
    R:=FListe.CellRect(ACol,ARow) ;
    Canvas.Rectangle(R.Left,R.Top,R.Right+1,R.Bottom+1) ;
    END ;
END ;

Procedure TFAnalysHT.TitresColRow ;
Var i,j : Integer ;
BEGIN
FListe.ColCount:=CRegime.Items.Count+1 ;
FListe.RowCount:=(4*(CTva.Items.Count+1))+4 ;
FListe.ColWidths[0]:=150 ; FListe.Cells[0,0]:=HM.Mess[0] ;
for i:=1 to FListe.ColCount-1 do
    BEGIN
    FListe.ColWidths[i]:=100 ; FListe.Cells[i,0]:=CRegime.Items[i-1] ;
    FListe.ColAligns[i]:=taRightJustify ;
    END ;
for i:=0 to CTva.Items.Count-1 do
    BEGIN
    FListe.Cells[0,(4*i)+1]:=CTva.Items[i] ;
    FListe.Cells[0,(4*i)+2]:='    '+HM.Mess[1] ;
    FListe.Cells[0,(4*i)+3]:='    '+HM.Mess[2] ;
    FListe.Cells[0,(4*i)+4]:='    '+HM.Mess[3] ;
    END ;
j:=4*(CTva.Items.Count) ;
FListe.Cells[0,j+1]:=HM.Mess[7] ;
FListe.Cells[0,j+2]:='    '+HM.Mess[1] ;
FListe.Cells[0,j+3]:='    '+HM.Mess[2] ;
FListe.Cells[0,j+4]:='    '+HM.Mess[3] ;
FListe.Cells[0,j+5]:=HM.Mess[4] ;
FListe.Cells[0,j+6]:=HM.Mess[5] ;
FListe.Cells[0,j+7]:=HM.Mess[6] ;
END ;

procedure TFAnalysHT.BChercheClick(Sender: TObject);
begin
Stop:=False ;
NettoieLaGrille ; InitLeWhere ; RempliLNumPiece ; RempliLGenCol ; RempliLGenChaPro ;
TitresColRow ; BStop.Enabled:=True ; ChercheLesMvts ; BStop.Enabled:=False ;
Stop:=False ;
end;

Procedure TFAnalysHT.InitLeWhere ;
Var NatureQualif,LesDates,WhereSupplement : String ;
BEGIN
if RgChoix.ItemIndex=0 then
   BEGIN
   NatJal:='ACH' ;
   NatureQualif:='(E_NATUREPIECE="FF" Or E_NATUREPIECE="AF") And E_QUALIFPIECE="N" And E_ECRANOUVEAU="N"' ;
   END else
   BEGIN
   NatJal:='VTE' ;
   NatureQualif:='(E_NATUREPIECE="FC" Or E_NATUREPIECE="AC") And E_QUALIFPIECE="N" And E_ECRANOUVEAU="N"' ;
   END ;
LesDates:='E_DATECOMPTABLE>="'+UsDate(DateCpta)+'" And E_DATECOMPTABLE<= "'+UsDate(DateCpta_)+'" ' ;
LesDateEche:='And E_DATEECHEANCE>="'+UsDate(DateEche)+'" And E_DATEECHEANCE<="'+UsDate(DateEche_)+'" ' ;
WhereSupplement:=UnWhereSupp ;
UnWhere:=NatureQualif+' And '+LesDates+WhereSupplement ;
END ;

Procedure TFAnalysHT.RempliLNumPiece ;
Var St : String ;
    QLoc : TQuery ;
BEGIN
St:='Select Distinct E_NUMEROPIECE,E_JOURNAL From ECRITURE '+
    'Left Join JOURNAL On E_JOURNAL=J_JOURNAL '+
    'Left Join GENERAUX On E_GENERAL=G_GENERAL '+
    'Where '+UnWhere+LesDateEche+' And J_NATUREJAL="'+NatJal+'" '+
    'And (G_NATUREGENE="COC" Or G_NATUREGENE ="COF" Or G_NATUREGENE ="TIC" Or G_NATUREGENE ="TID") '+
    'Order By E_NUMEROPIECE';
QLoc:=OpenSql(St,True) ; LNumPiece.Clear ;
While Not QLoc.Eof do BEGIN LNumPiece.Add(QLoc.Fields[0].AsString+';'+QLoc.Fields[1].AsString) ; QLoc.Next ; END ;
Ferme(QLoc) ;
END ;

Procedure TFAnalysHT.RempliLGenCol ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select G_GENERAL From GENERAUX Where G_NATUREGENE ="COC" Or '+
              'G_NATUREGENE ="COF" Or G_NATUREGENE ="TIC" Or G_NATUREGENE ="TID"',True) ;
LGenCol.Clear ;
While Not QLoc.Eof do BEGIN LGenCol.Add(QLoc.Fields[0].AsString) ; QLoc.Next ; END ;
Ferme(QLoc) ;
END ;

Procedure TFAnalysHT.RempliLGenChaPro ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select G_GENERAL From GENERAUX Where G_NATUREGENE ="CHA" Or G_NATUREGENE ="PRO"',True) ;
LGenChaPro.Clear ;
While Not QLoc.Eof do BEGIN LGenChaPro.Add(QLoc.Fields[0].AsString) ; QLoc.Next ; END ;
Ferme(QLoc) ;
END ;

Function TFAnalysHT.UnWhereSupp : String ;
Var St : String ;
BEGIN
St:='' ;
if Exercice.Value<>'' then St:=St+' And E_EXERCICE="'+Exercice.Value+'" ' ;
if Etablissement.Value<>'' then St:=St+' And E_ETABLISSEMENT="'+Etablissement.Value+'" ' ;
if Devise.Value<>'' then St:=St+' And E_DEVISE="'+Devise.Value+'" ' ;
Result:=St ;
END ;

Function TFAnalysHT.ChercheTauxTva(CodeTva,Regime : String) : Double ;
Var QLoc : TQuery ;
    ChampTaux : String ;
BEGIN
if CodeTva='' then BEGIN Result:=0 ; Exit ; END ;
if RgChoix.ItemIndex=0 then ChampTaux:='TV_TAUXACH' else ChampTaux:='TV_TAUXVTE' ;
QLoc:=OpenSql('Select '+ChampTaux+' From TXCPTTVA Where TV_TVAOUTPF="'+VH^.DefCatTVA+'" And '+
              'TV_CODETAUX="'+CodeTva+'" And TV_REGIME="'+Regime+'"',True) ;
Result:=QLoc.Fields[0].AsFloat ; Ferme(QLoc) ;
END ;

Function TFAnalysHT.EstCpteTTC(StGen : String) : Boolean ;
BEGIN Result:=(LGenCol.IndexOf(StGen)>=0) ; END ;

Function TFAnalysHT.EstCpteHT(StGen : String) : Boolean ;
BEGIN Result:=(LGenChaPro.IndexOf(StGen)>=0) ; END ;

Function TFAnalysHT.OnTraiteLaPiece(St1 : String) : Boolean ;
Var St : String ;
BEGIN
St:='Select Count(Distinct E_AUXILIAIRE) From ECRITURE Where '+UnWhere+' And '+
    'E_NUMEROPIECE='+ReadTokenSt(St1)+' And E_JOURNAL="'+ReadTokenSt(St1)+'" '+
    'And E_AUXILIAIRE<>"" ' ;
QEcr.Close ; QEcr.Sql.Clear ; QEcr.Sql.Add(St) ; ChangeSql(QEcr) ; QEcr.Open ;
if QEcr.Fields[0].AsInteger>1 then Result:=False else Result:=True ;
END ;

Function TFAnalysHT.DateEcheDansCrit(D : TDateTime) : Boolean ;
BEGIN Result:=(D>=StrToDate(DateEche.Text)) And (D<=StrToDate(DateEche_.Text)) ; END ;

Function TFAnalysHT.QuelCol(Regtva : String) : Integer ;
BEGIN Result:=CRegime.Values.IndexOf(Regtva)+1 ; END ;

Function TFAnalysHT.QuelRow(Tauxtva : String ; TypLig : Integer) : Integer ;
Var i : Integer ;
    St : String ;
BEGIN
Result:=0 ;
if Tauxtva='' then St:=HM.Mess[7]
              else St:=CTva.Items[CTva.Values.IndexOf(Tauxtva)] ;
for i:=1 to FListe.RowCount-1 do if Trim(FListe.Cells[0,i])=St then BEGIN Result:=i ; Break ; END ;
Result:=Result+TypLig ;
END ;

Procedure TFAnalysHT.TrouveAColARow(TypLig : Integer ; Var ACol,ARow : Integer) ;
BEGIN
ACol:=QuelCol(QEcr.FindField('E_REGIMETVA').AsString) ;
ARow:=QuelRow(QEcr.FindField('E_TVA').AsString,TypLig) ;
END ;

Function TFAnalysHT.TrouveLaSomme : Double ;
Var ChampDeb,ChampCred : String ;
BEGIN
Result:=0 ;
if (Devise.Value='') Or (Devise.Value=V_PGI.DevisePivot) then
    BEGIN ChampDeb:='E_DEBIT' ; ChampCred:='E_CREDIT' ; END else
    BEGIN ChampDeb:='E_DEBITDEV' ; ChampCred:='E_CREDITDEV' ; END ;
Case RgChoix.ItemIndex of
     0 : Result:=QEcr.FindField(ChampCred).AsFloat-QEcr.FindField(ChampDeb).AsFloat ;
     1 : Result:=QEcr.FindField(ChampDeb).AsFloat-QEcr.FindField(ChampCred).AsFloat ;
   End ;
Result:=Abs(Result) ;
END ;

Procedure TFAnalysHT.FaitTotaux ;
Var i,j : Integer ;
    Ht,Ttc,Tva : Double ;
    LTotaHT,LTotalTTC,LTotalTVA : Integer ;
BEGIN
LTotaHT:=0 ; LTotalTTC:=0 ; LTotalTVA:=0 ;
for i:=1 to FListe.RowCount-1 do
    BEGIN
    if FListe.Cells[0,i]=HM.Mess[4] then LTotaHT:=i else
       if FListe.Cells[0,i]=HM.Mess[6] then LTotalTTC:=i else
          if FListe.Cells[0,i]=HM.Mess[5] then LTotalTVA:=i  ;
    END ;
for i:=1 to FListe.ColCount-1 do
    BEGIN
    Ht:=0 ; Ttc:=0 ; Tva:=0 ;
    for j:=1 to FListe.RowCount-1 do
        BEGIN
        if (Trim(FListe.Cells[0,j])=HM.Mess[1]) And (FListe.Cells[i,j]<>'') then Ht:=Ht+Valeur(FListe.Cells[i,j]) else
           if (Trim(FListe.Cells[0,j])=HM.Mess[3]) And (FListe.Cells[i,j]<>'') then Ttc:=Ttc+Valeur(FListe.Cells[i,j]) else
              if (Trim(FListe.Cells[0,j])=HM.Mess[2]) And (FListe.Cells[i,j]<>'') then Tva:=Tva+Valeur(FListe.Cells[i,j]) ;
        END ;
    if Ht<>0 then FListe.Cells[i,LTotaHT]:=StrfMontant(Ht,15,2,'',True) ;
    if Ttc<>0 then FListe.Cells[i,LTotalTTC]:=StrfMontant(Ttc,15,2,'',True) ;
    if Tva<>0 then FListe.Cells[i,LTotalTVA]:=StrfMontant(Tva,15,2,'',True) ;
    END ;
END ;

Procedure TFAnalysHT.ChercheLesMvts ;
Var i : Integer ;
    St : String ;
    St1 : String ;
    TTC,TotTTC,TotHT,TotTVA : Double ;
    ACol,ARow : Integer ;
    X : TAnalTva ;
BEGIN
InitMove(LNumPiece.Count,'') ;
for i:=0 to LNumPiece.Count-1 do
   BEGIN
   Application.ProcessMessages ;
   if Stop then BEGIN if HM.Execute(8,'','')=mrYes then Break else Stop:=False ; END ;
   MoveCur(False) ; St1:=LNumPiece.Strings[i] ; VideListe(LPiece) ;
   TotTTC:=0 ; TotHT:=0 ; TotTVA:=0 ; TTC:=0 ;
   if Not OnTraiteLaPiece(St1) then Continue ;
   St:='Select * from ECRITURE Where '+UnWhere+' And E_NUMEROPIECE='+ReadTokenSt(St1)+' And '+
       'E_JOURNAL="'+ReadTokenSt(St1)+'"' ;
   QEcr.Close ; QEcr.Sql.Clear ; QEcr.Sql.Add(St) ; ChangeSql(QEcr) ;
   QEcr.Open ;
   While not QEcr.Eof do
     BEGIN
     if EstCpteTTC(QEcr.FindField('E_GENERAL').AsString) then
        BEGIN
        TrouveAColARow(3,ACol,ARow) ; TotTTC:=TotTTC+TrouveLaSomme ;
        if DateEcheDansCrit(QEcr.FindField('E_DATEECHEANCE').AsDateTime) then
           BEGIN
           TTC:=TTC+TrouveLaSomme ; X:=CreerLobjet(QEcr,3,ACol,ARow) ; LPiece.Add(X) ;
           END ;
        END else
        if EstCpteHT(QEcr.FindField('E_GENERAL').AsString) then
           BEGIN
           TrouveAColARow(1,ACol,ARow) ; X:=CreerLobjet(QEcr,1,ACol,ARow) ; LPiece.Add(X) ;
           END else
           BEGIN
           TrouveAColARow(2,ACol,ARow) ; X:=CreerLobjet(QEcr,2,ACol,ARow) ; LPiece.Add(X) ;
           END ;
     QEcr.Next ;
     END ;
   ProratLesSommes(TotTTC,TTC,TotTTC<>TTC) ;
   END ;
QEcr.Close ; FaitTotaux ; FiniMove ;
END ;

Procedure TFAnalysHT.ProratLesSommes(TotTTC,TTC : Double ; Prorat : Boolean) ;
Var i : Integer ;
    Somme : Double ;
    X,Y : TAnalTva ;
    Li : TList ;
//TypCpte 1-->HT , 2 -->TVA , 3-->TTC
BEGIN
for i:=0 to LPiece.Count-1 do
    BEGIN
    Y:=TAnalTva(LPiece[i]) ; Somme:=Y.SommeTotal ;
    if ((Y.TypCpte=1) Or (Y.TypCpte=2)) And Prorat then Somme:=Somme*(TTC/TotTTC) ;
    FListe.Cells[Y.ACol,Y.ARow]:=StrfMontant(Valeur(FListe.Cells[Y.ACol,Y.ARow])+Somme,15,2,'',True) ;
    if FListe.Objects[Y.ACol,Y.ARow]=Nil then
       BEGIN
       Li:=TList.Create ; X:=TAnalTva.Create ; CopieLobjet(X,Y) ;
       Li.Add(X) ; FListe.Objects[Y.ACol,Y.ARow]:=Li ;
       END else
       BEGIN
       X:=TAnalTva.Create ; CopieLobjet(X,Y) ;
       TList(FListe.Objects[Y.ACol,Y.ARow]).Add(X) ;
       END ;
    END ;
END ;

procedure TFAnalysHT.BStopClick(Sender: TObject);
begin Stop:=True ; end;

procedure TFAnalysHT.FListeDblClick(Sender: TObject);
begin
if FListe.Objects[FListe.Col,FListe.Row]=Nil then Exit ;
if BStop.Enabled then Exit ;
MontreLeDetail(TList(FListe.Objects[FListe.Col,FListe.Row]),Pos('TTC',FListe.Cells[0,FListe.Row])>0) ;
end;

end.
