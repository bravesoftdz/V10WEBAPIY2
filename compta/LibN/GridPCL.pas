unit GridPcl ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, HPanel, Grids, Hctrls, UTob, HEnt1, Formule, HTB97, PrintDBG,
  ComCtrls, StdCtrls,HStatus, Buttons, HColor,QRGrid,CritEdt, Hcompte ,Ent1 ;

procedure LanceLiasse (ModeleName : String ; VoirDetail,VoirNom : Boolean ;
                       FormuleCol,FormatCol : Array of String ; Var Crit : tCritEdtPCL ; Demo : Boolean) ;
CONST  MAX_COL = 20 ;

type
  TFEtatLiasse = class(TForm)
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock971: TDock97;
    BImprimer: TToolbarButton97;
    Pages: TPageControl;
    PResult: TTabSheet;
    PDetails: TTabSheet;
    FListe: TStringGrid;
    PParams: TTabSheet;
    GParams: THGrid;
    FType: THValComboBox;
    FD: TFontDialog;
    CD: TColorDialog;
    FTestFont: TLabel;
    Panel1: TPanel;
    Label1: TLabel;
    FTitre: TEdit;
    BDelLigne: TBitBtn;
    BInsLigne: TBitBtn;
    bSauve: TToolbarButton97;
    bColor: TToolbarButton97;
    bPolice: TToolbarButton97;
    bDroite: TToolbarButton97;
    bCentre: TToolbarButton97;
    bGauche: TToolbarButton97;
    bFrame: TPaletteButton97;
    FDetail: TStringGrid;
    Cache: THCpteEdit;
    LibRub: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListeDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure BImprimerClick(Sender: TObject);
    procedure PagesChange(Sender: TObject);
    procedure bSauveClick(Sender: TObject);
    procedure GParamsRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GParamsRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
    procedure GParamsCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BDelLigneClick(Sender: TObject);
    procedure BInsLigneClick(Sender: TObject);
    procedure GParamsRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    procedure bPoliceClick(Sender: TObject);
    procedure bColorClick(Sender: TObject);
    procedure bGaucheClick(Sender: TObject);
    procedure bFrameChange(Sender: TObject);
    procedure GParamsDblClick(Sender: TObject);
    procedure GParamsKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private   { Déclarations privées }
    FModeleName : String ;
    GX,GY : Integer ;
    FCurrentTOB : TOB ;
    FTL : TStringList ;
    FVoirDetail,FVoirNom,FDetailCalcule : Boolean ;
    FFormuleCol,FFormatCol : Array [1..MAX_COL] of String ;
    FCurrentCol,FNbCol : Integer ;
    FLaTOB,FDetailTob : TOB ;
    Crit : tCritEdtPCL ;
    Demo : Boolean ;
    procedure Charge ;
    procedure ChargeDetail ;
    procedure CalculRub ;
    procedure CalculRubDetail ;
    function  LeCumul(Nom : String ; LaCol : Integer ; Var LesCptes : tStringList) : Variant ;
    function  LeCumulDetail(Nom : String ; LaCol : Integer) : Double ;
    Function  GetCellCumul ( St : String) : Variant ;
    procedure GetRubDetail (Nom : String ; Noms,Libelles : TStringList );
    procedure DecodeFont (St : string) ;
    Function  EncodeFont : string;
    Function  AddGras ( St : String ) : String ;
    Function  IsLigneVide (ARow : Integer) : Boolean ;
    Procedure GDrawCell(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;

  public    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses CalcOLE, QRBalGen ;

CONST POS_COL : INTEGER = 2 ;
      POS_LIBELLE : Integer = 1 ;
      POS_NOM : Integer = 0 ;
      Default_Style = 'Arial;8;;clBlack;clWhite;----;' ;

procedure LanceLiasse (ModeleName : String ; VoirDetail,VoirNom : Boolean ;
                       FormuleCol,FormatCol : Array of String ; Var Crit : tCritEdtPCL ; Demo : Boolean) ;
var FEtatLiasse: TFEtatLiasse;
    i : Integer ;
BEGIN
FEtatLiasse:=TFEtatLiasse.Create(Application) ;
FEtatLiasse.FModeleName:=ModeleName ;
FEtatLiasse.FVoirDetail:=VoirDetail ;
FEtatLiasse.FVoirNom:=VoirNom ;
FEtatLiasse.Crit:=Crit ;
FEtatLiasse.Demo:=Demo ;
for i:=1 to MAX_COL do BEGIN FEtatLiasse.FFormuleCol[i]:='' ; FEtatLiasse.FFormatCol[i]:='' ; END ;
for i:=Low(FormuleCol) to High(FormuleCol) do FEtatLiasse.FFormuleCol[i+1]:=FormuleCol[i] ;
for i:=Low(FormatCol) to High(FormatCol) do FEtatLiasse.FFormatCol[i+1]:=FormatCol[i] ;
FEtatLiasse.FNbCol:=High(FormuleCol)+1 ;
FEtatLiasse.ShowModal ;
FEtatLiasse.Free ;
END ;

procedure TFEtatLiasse.FormShow(Sender: TObject);
Var i : Integer ;
begin
Cache.ZoomTable:=tzRubCPTA ;
Pages.ActivePageIndex:=0 ;
if FVoirNom then BEGIN POS_NOM:=0 ; POS_LIBELLE:=1 ; POS_COL:=2 ; END
            else BEGIN POS_NOM:=-1 ; POS_LIBELLE:=0 ;POS_COL:=1 ;  END ;
if POS_LIBELLE>=0 then BEGIN FListe.Cells[POS_LIBELLE,0]:='Libellé' ; FListe.ColWidths[POS_LIBELLE]:=250 ; END ;
if POS_NOM>=0 then BEGIN FListe.Cells[POS_NOM,0]:='Nom cellules' ; FListe.ColWidths[POS_NOM]:=80 ; END ;
FListe.ColCount:=FNbCol+1+Ord(FVoirNom) ;
FDetail.ColCount:=FListe.ColCount ;
if POS_LIBELLE>=0 then FDetail.ColWidths[POS_LIBELLE]:=FListe.ColWidths[POS_LIBELLE] ;
if POS_NOM>=0 then FDetail.ColWidths[POS_NOM]:=FListe.ColWidths[POS_NOM] ;
For i:=1 To FNbCol Do
  BEGIN
  FListe.Cells[POS_COL+i-1,0]:=Crit.Col[i-1].StTitre ;
  END ;
Charge ;
If Demo Then
  BEGIN
  Pages.ActivePage:=PParams ;
//  PResult.TabVisible:=FALSE ;
  PDetails.TabVisible:=FALSE ;
  GParams.SetFocus ;
  PagesChange(Nil);
  END ;
end;

procedure TFEtatLiasse.FormDestroy(Sender: TObject);
begin
FTL.Free ;  FLaTOB.Free ; FDetailTOB.Free ;
end;

procedure TFEtatLiasse.FormCreate(Sender: TObject);
begin
FDetailCalcule:=FALSE ;
FTL:=TStringList.Create ;
FLaTOB:=TOB.Create('Les rubriques',Nil,-1) ;
FDetailTOB:=TOB.Create('Les rubriques detail',Nil,-1) ;
end;

procedure TFEtatLiasse.Charge;
Var i,nb,ih,kk : Integer ;
    First : Boolean ;
    T : TOB ;
    St,F1,FA : String ;
begin
FTL.LoadFromFile(FModeleName); FLaTOB.ClearDetail ;
Caption:=FTL[0] ; First:=TRUE ;
FTL.Delete(0) ; FTL.Delete(0) ;
FListe.RowCount:=2 ;
//FListe.Rows.BeginUpdate ;
For i:=0 to FTL.Count-1 do
   BEGIN
   St:=FTL[i] ;
   T:=TOB.Create('Une TOB',FLaTob,-1) ;
   T.AddChampSupValeur('TYPE',Trim(UpperCase(ReadTokenPipe(St,'|')))) ;
   T.AddChampSupValeur('ALIGN',Trim(UpperCase(ReadTokenPipe(St,'|')))) ;
   T.AddChampSupValeur('STYLE',Trim(UpperCase(ReadTokenPipe(St,'|')))) ;
   T.AddChampSupValeur('LIBELLE',Trim(ReadTokenPipe(St,'|'))) ;
   T.AddChampSupValeur('NOM',Trim(UpperCase(ReadTokenPipe(St,'|')))) ;
   T.AddChampSupValeur('BASE',Trim(UpperCase(ReadTokenPipe(St,'|')))) ;
   for kk:=1 to MAX_COL do
      BEGIN
      FA:='' ;
      if kk=1 then F1:=Trim(UpperCase(ReadTokenPipe(St,'|')))
              else FA:=Trim(UpperCase(ReadTokenPipe(St,'|'))) ;
      if FA='' then FA:=F1 ;
      T.AddChampSupValeur('FORMULE'+IntToStr(kk),FA) ;
      T.AddChampSupValeur('VALEUR'+IntToStr(kk),0.00) ;
      END ;
   ih:=ValeurI(ReadTokenPipe(St,'|')) ; if ih<=0 then ih:=18 ;   //RowHeight
   if (T.GetValue('TYPE')<>'I') and (T.GetValue('TYPE')<>'L') then
      BEGIN
      if Not First then FListe.RowCount:=FListe.RowCount+1 else First:=FALSE ;
      nb:=FListe.RowCount-1 ;
      END else Nb:=-1 ;
   T.AddChampSupValeur('LIGNE',nb) ;
   if nb>0 then
      BEGIN
      if POS_LIBELLE>=0 then FListe.Cells[POS_LIBELLE,nb]:=T.GetValue('LIBELLE') ;
      if POS_NOM>=0 then FListe.Cells[POS_NOM,nb]:=T.GetValue('NOM') ;
      FListe.Objects[0,nb]:=T ;
      FListe.RowHeights[nb]:=ih ;
      END ;
   END ;
//FListe.Rows.EndUpdate ;
CalculRub ;
end;


procedure TFEtatLiasse.CalculRub;
Var kk,i,j,nbl : Integer ;
    T : TOB ;
    okok : boolean ;
    tipe,St,StF,StFF : String ;
    X : Double ;
    VV : Variant ;
    StErr : String ;
    LesCptes : tStringList ;
begin
InitMove(1000,'') ;
for kk:=1 to 4 do
 For i:=0 to FLaTob.Detail.Count-1 do
   BEGIN
   MoveCur(FALSE) ;
   T:=FLaTOB.Detail[i] ; FCurrentTob:=T ;
   nbl:=T.GetValue('LIGNE') ; tipe:=T.GetValue('TYPE') ;
   if (tipe<>'C') then
       BEGIN
       for j:=1 to FNbCol do
           BEGIN
           FCurrentCol:=j ; okok:=FALSE ;
           St:=T.GetValue('FORMULE'+IntToStr(j)) ;
           StFF:=FFormuleCol[j] ; X:=0 ;
           if  (kk=1) and (St<>'')  and (Copy(St,1,1)<>'=') and (Copy(StFF,1,1)<>'=') then
                BEGIN
                StErr:='' ;
                LesCptes:=NIL ;
                VV:=LeCumul(St,j,LesCptes) ; OkOk:=TRUE ;
                If LesCptes<>NIL Then FListe.Objects[POS_COL+j-1,nbl]:=LesCptes ;
                If V_PGI.StOLEErr='' Then X:=VV Else
                  BEGIN
                  StErr:=VV ; X:=0 ;
                  END ;
                END ;
           if (kk=2) and (Copy(St,1,1)='=') and (Copy(StFF,1,1)<>'=') then
                BEGIN
                StFF:=St ; Delete(StFF,1,1) ; StFF:='{'+StFF+'}' ;
                StF:=GFormule(StFF,GetCellCumul,Nil,1) ;
                X:=Valeur(StF) ;
                okok:=TRUE ;
                END ;
           if (kk=3) and (St<>'') and (Copy(St,1,1)<>'=') and (Copy(StFF,1,1)='=') then
                BEGIN
                Delete(StFF,1,1) ; StFF:='{'+StFF+'}' ;
                StF:=GFormule(StFF,GetCellCumul,Nil,1) ;
                X:=Valeur(StF) ;
                okok:=TRUE ;
                END ;
           if (kk=4) and (Copy(St,1,1)='=') and (Copy(StFF,1,1)='=') then
                BEGIN
                Delete(StFF,1,1) ; StFF:='{'+StFF+'}' ;
                StF:=GFormule(StFF,GetCellCumul,Nil,1) ;
                X:=Valeur(StF) ;
                okok:=TRUE ;
                END ;

           if okok then
              BEGIN
              T.PutValue('VALEUR'+IntToStr(j),X) ;
              if nbl>0 then
                BEGIN
                If StErr<>'' Then FListe.Cells[POS_COL+j-1,nbl]:=StErr
                             Else FListe.Cells[POS_COL+j-1,nbl]:=FormatFloat(FFormatcol[j],X) ;
                END ;
              END ;
           END ;
       END ;
   END ;
 FiniMove ;
end;

function TFEtatLiasse.LeCumul(Nom : String ; LaCol : Integer ; Var LesCptes : tStringList) : Variant ;
Var Appel : String ;
begin
//FFormuleCol[LaCol]
//Result:=Get_Cumul('RUBRIQUE',Nom, FFormuleCol[LaCol], AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
If Demo Then Result:=Valeur(Nom)*100*LaCol Else
  BEGIN
  (*
  if GetCumul(Appel,Nom,'',LRubAno,Crit.Etab,Crit.Devise,Crit.Col[LaCol-1].Exo.Code,Crit.Col[LaCol-1].Date1,Crit.Col[LaCol-1].Date2,
              False,False,LesCpte,TResult,FALSE)<>CodError then
  *)
  Appel:='RUBRIQUE' ;
  LesCptes:=TStringList.Create ;
  Result:=Get_CumulPCL(Appel,Nom,Crit.StTypEcr,Crit.Etab,Crit.Devise,FFormuleCol[LaCol],'',LesCptes) ;
  END ;
end;

function TFEtatLiasse.LeCumulDetail(Nom : String ; LaCol : Integer) : Double ;
Var Appel : String ;
begin
//FFormuleCol[LaCol]
If Demo Then Result:=Valeur(Nom)*100*LaCol Else
  BEGIN
  (*
  if GetCumul(Appel,Nom,'',LRubAno,Crit.Etab,Crit.Devise,Crit.Col[LaCol-1].Exo.Code,Crit.Col[LaCol-1].Date1,Crit.Col[LaCol-1].Date2,
              False,False,LesCpte,TResult,FALSE)<>CodError then
  *)
  Appel:='RUBRIQUE' ;
  Result:=Get_Cumul(Appel,Nom,Crit.StTypEcr,Crit.Etab,Crit.Devise,FFormuleCol[LaCol],'') ;
  END ;
end;

function TFEtatLiasse.GetCellCumul(St: String): Variant;
Var T : TOB ;
    i : Integer ;
    StB : String ;
begin
if Copy(St,1,3)='COL' then
   BEGIN
   i:=ValeurI(St) ;
   if FCurrentTOB<>Nil then Result:=FCurrentTOB.GetValue('VALEUR'+IntToStr(i)) else result:=0 ;
   END else
if Copy(St,1,4)='BASE' then
   BEGIN
   if FCurrentTOB=Nil then BEGIN result:=0 ; exit ; END ;
   StB:=FCurrentTOB.GetValue('BASE') ;
   T:=FLaTob.FindFirst(['NOM'],[Stb],FALSE );
   if T<>Nil then Result:=T.GetValue('VALEUR'+IntToStr(ValeurI(St))) else result:=0 ;
   END else
   BEGIN
   T:=FLaTob.FindFirst(['NOM'],[St],FALSE );
   if T<>Nil then Result:=T.GetValue('VALEUR'+IntToStr(FCurrentCol)) else result:=0 ;
   END ;
end;

procedure TFEtatLiasse.GetRubDetail(Nom: String; Noms,Libelles: TStringList);
Var i,k : Integer ;
    LesCpte : tStringList ;
begin
Noms.Clear ; Libelles.Clear ;
///
k:=ValeurI(Nom) ;
//For i:=1 to k do
For i:=1 to FNbCol do
   BEGIN
   Noms.Add(Nom+' D'+IntToStr(i)) ;
   Libelles.Add(Nom+' Détail n°'+IntToStr(i)) ;
   END ;
end;

procedure TFEtatLiasse.CalculRubDetail;
Var kk,i,j,nbl : Integer ;
    T : TOB ;
    okok : boolean ;
    tipe,St,StF,StFF : String ;
    X : Double ;

begin
InitMove(1000,'') ;
for kk:=1 to 4 do
 For i:=0 to FDetailTob.Detail.Count-1 do
   BEGIN
   MoveCur(FALSE) ;
   T:=FDetailTob.Detail[i] ; FCurrentTob:=T ;
   nbl:=T.GetValue('LIGNE') ; tipe:=T.GetValue('TYPE') ;
   for j:=1 to FNbCol do
       BEGIN
       St:=T.GetValue('FORMULE'+IntToStr(j)) ;
       StFF:=FFormuleCol[j] ; X:=0 ;
       if (tipe='D') then
           BEGIN
           FCurrentCol:=j ; okok:=FALSE ;
           if  (kk=1) and (St<>'')  and (Copy(St,1,1)<>'=') and (Copy(StFF,1,1)<>'=') then
                BEGIN
                X:=LeCumulDetail(St,j) ; ///premier passage rubrique
                okok:=TRUE ;
                END ;
           if (kk=2) and (Copy(St,1,1)='=') and (Copy(StFF,1,1)<>'=') then
                BEGIN
                StFF:=St ; Delete(StFF,1,1) ; StFF:='{'+StFF+'}' ;
                StF:=GFormule(StFF,GetCellCumul,Nil,1) ;
                X:=Valeur(StF) ;
                okok:=TRUE ;
                END ;
           if (kk=3) and (St<>'') and (Copy(St,1,1)<>'=') and (Copy(StFF,1,1)='=') then
                BEGIN
                Delete(StFF,1,1) ; StFF:='{'+StFF+'}' ;
                StF:=GFormule(StFF,GetCellCumul,Nil,1) ;
                X:=Valeur(StF) ;
                okok:=TRUE ;
                END ;
           if (kk=4) and (Copy(St,1,1)='=') and (Copy(StFF,1,1)='=') then
                BEGIN
                Delete(StFF,1,1) ; StFF:='{'+StFF+'}' ;
                StF:=GFormule(StFF,GetCellCumul,Nil,1) ;
                X:=Valeur(StF) ;
                okok:=TRUE ;
                END ;

           if okok then
              BEGIN
              T.PutValue('VALEUR'+IntToStr(j),X) ;
              if nbl>0 then FDetail.Cells[POS_COL+j-1,nbl]:=FormatFloat(FFormatcol[j],X) ;
              END ;
           END else
           BEGIN
           if (tipe='F') or (tipe='R') or (tipe='L') then
             if (St<>'') or (Copy(StFF,1,1)='=') then
             BEGIN
             if nbl>0 then FDetail.Cells[POS_COL+j-1,nbl]:=FormatFloat(FFormatcol[j],T.GetValue('VALEUR'+IntToStr(j))) ;
             END ;
           END ;
       END ;
   END ;
 FiniMove ;
end;


procedure TFEtatLiasse.ChargeDetail;
Var i,nb,j,ih,ih1,kk : Integer ;
    OkDetail,First : Boolean ;
    T1,T2,T3 : TOB ;
    TN,TL : TStringList ;
    StF : String ;
begin
//Exit ;
if FDetailCalcule then exit ;
FDetailTOB.ClearDetail ;
First:=TRUE ; FDetailCalcule:=TRUE ;
FDetail.RowCount:=2 ;
TN:=TStringList.Create ; TL:=TStringList.Create ;
//FDetail.Rows.BeginUpdate ;
For i:=0 to FLaTob.Detail.Count-1 do
   BEGIN
   T1:=FLaTob.Detail[i] ; okdetail:=FALSE ;
   ih:=18 ; ih1:=T1.GetValue('LIGNE') ; if ih1>0 then ih:=FListe.RowHeights[ih1] ;
   if (T1.GetValue('TYPE')='R') or (T1.GetValue('TYPE')='L') then
      BEGIN
      GetRubDetail(T1.GetValue('FORMULE1'),TN,TL) ; okdetail:=TRUE ;
      for j:=0 to TN.Count-1 do
         BEGIN
         T3:=TOB.Create('Une TOB detail',FDetailTob,-1) ;
         T3.AddChampSupValeur('TYPE','D') ;
         T3.AddChampSupValeur('ALIGN',T1.GetValue('ALIGN')) ;
         T3.AddChampSupValeur('STYLE',T1.GetValue('STYLE')) ;
         T3.AddChampSupValeur('LIBELLE',TL[j]) ;
         T3.AddChampSupValeur('NOM',TN[j]) ;
         T3.AddChampSupValeur('BASE',T1.GetValue('BASE')) ;
         for kk:=1 to MAX_COL do
            BEGIN
            StF:=T1.GetValue('FORMULE'+IntToStr(kk)) ;
            if StF='' then T3.AddChampSupValeur('FORMULE'+IntToStr(kk),'')else
            if Copy(StF,1,1)<>'=' then T3.AddChampSupValeur('FORMULE'+IntToStr(kk),TN[j]) else
                                       T3.AddChampSupValeur('FORMULE'+IntToStr(kk),T1.GetValue('FORMULE'+IntToStr(kk))) ;
            T3.AddChampSupValeur('VALEUR'+IntToStr(kk),0.00) ;
            END ;
         if Not First then FDetail.RowCount:=FDetail.RowCount+1 else First:=FALSE ;
         nb:=FDetail.RowCount-1 ;
         T3.AddChampSupValeur('LIGNE',nb) ;
         if POS_LIBELLE>=0 then FDetail.Cells[POS_LIBELLE,nb]:=T3.GetValue('LIBELLE') ;
         if POS_NOM>=0 then FDetail.Cells[POS_NOM,nb]:=T3.GetValue('NOM') ;
         FDetail.Objects[0,nb]:=T3 ;
         FDetail.RowHeights[nb]:=ih ;
         END ;
      END ;
   T2:=TOB.Create('Une TOB',FDetailTob,-1) ;
   T2.Dupliquer(T1,FALSE,TRUE) ;
//   if OkDetail then
   T2.PutValue('STYLE',AddGras(T2.GetValue('STYLE'))) ;
   if T2.GetValue('TYPE')<>'I' then
      BEGIN
      if Not First then FDetail.RowCount:=FDetail.RowCount+1 else First:=FALSE ;
      nb:=FDetail.RowCount-1 ;
      END else Nb:=-1 ;
   T2.AddChampSupValeur('LIGNE',nb) ;
   if nb>0 then
      BEGIN
      if POS_LIBELLE>=0 then FDetail.Cells[POS_LIBELLE,nb]:=T2.GetValue('LIBELLE') ;
      if POS_NOM>=0 then FDetail.Cells[POS_NOM,nb]:=T2.GetValue('NOM') ;
      FDetail.Objects[0,nb]:=T2 ;
      FDetail.RowHeights[nb]:=ih ;
      END ;
   END ;
//FDetail.Rows.EndUpdate ;
TN.Free ; TL.Free ;
CalculRubDetail ;
end;


procedure TFEtatLiasse.FListeDrawCell(Sender: TObject; ACol, ARow: Integer;Rect: TRect; State: TGridDrawState);
var Text,Style   : String ;
    T      : TOB ;
    f      : TAlignment ;
    St     : String ;
    GG : TStringGrid ;
begin
GG:=TStringGrid(Sender) ;
//OldBrush.Assign(Canvas.Brush) ; OldPen.Assign(Canvas.Pen) ; OldFont.Assign(Canvas.Font) ;
T:=TOB(GG.Objects[0,ARow]) ; if T=Nil then exit ;
text:=GG.Cells[ACol,ARow] ; if Text=#0 then text:='' ;
Style:=T.GetValue('STYLE') ;
if Style='' then Style:=Default_Style ;
DecodeFont(Style) ;
GG.Canvas.Font:=FTestFont.Font ;
if ACol=POS_NOM then f:=taLeftJustify else
 if ACol<>POS_LIBELLE then f:=taRightJustify else
   BEGIN
   St:=T.GetValue('ALIGN') ; f:=taLeftJustify ;
   if St='D' then f:=taRightJustify ; if St='C' then f:=taCenter ;
   END ;
if ARow<GG.FixedRows then F:=taCenter ;
if (gdfixed in State) then GG.Canvas.Brush.Color:=GG.FixedColor else
    BEGIN
    if ((gdSelected in State) or (gdFocused in State)) then
        BEGIN
        GG.Canvas.Brush.Color:=clHighlight ;
        GG.Canvas.Font.Color:=clHighlightText;
        END else
        BEGIN
        GG.Canvas.Brush.Color:=FTestFont.Color ;
        END ;
    END ;
Case f of
     taRightJustify : ExtTextOut(GG.Canvas.Handle, Rect.Right - GG.canvas.TextWidth(Text) - 3, Rect.Top + 2,
        ETO_OPAQUE or ETO_CLIPPED,@Rect,PChar(Text),Length(Text),nil) ;
    taCenter       : ExtTextOut(GG.Canvas.Handle, Rect.Left +
        ((Rect.Right-Rect.Left-GG.canvas.TextWidth(Text)) div 2) , Rect.Top + 2,
        ETO_OPAQUE or ETO_CLIPPED,@Rect,PChar(Text),Length(Text),nil) ;
    else ExtTextOut(GG.Canvas.Handle, Rect.Left + 2, Rect.Top + 2,
                           ETO_OPAQUE or ETO_CLIPPED,@Rect,PChar(Text),Length(Text),nil) ;
    END ;
if Style<>'' then
   BEGIN
   GG.Canvas.Pen.Style:=psSolid ;
   GG.Canvas.Pen.Color:=clBlack ;
   if akTop in FTestFont.Anchors then BEGIN GG.Canvas.MoveTo(Rect.Left,Rect.Top) ; GG.Canvas.LineTo(Rect.Right+1,Rect.Top) ; END ;
//   if akLeft in FTestFont.Anchors then BEGIN GG.Canvas.MoveTo(Rect.Left,Rect.Top) ; GG.Canvas.LineTo(Rect.Left,Rect.Bottom+1) ;END ;
//   if akRight in FTestFont.Anchors then BEGIN GG.Canvas.MoveTo(Rect.Right,Rect.Top) ; GG.Canvas.LineTo(Rect.Right,Rect.Bottom+1) ; END ;
   if akBottom in FTestFont.Anchors then BEGIN GG.Canvas.MoveTo(Rect.Left,Rect.Bottom) ; GG.Canvas.LineTo(Rect.Right+1,Rect.Bottom) ; END ;
   END ;

//Canvas.Brush.Assign(OldBrush) ; Canvas.Pen.Assign(OldPen) ; Canvas.Font.Assign(OldFont) ;
end;

procedure TFEtatLiasse.BImprimerClick(Sender: TObject);
begin
PrintGrid([FListe,FDetail],caption);
end;

procedure TFEtatLiasse.PagesChange(Sender: TObject);
Var i,kk,ih : Integer ;
    TL : TStringList ;
    St : String ;
begin
bSauve.Visible:=Pages.ActivePage=PParams ;
if Pages.ActivePage=PDetails then
   BEGIN
   ChargeDetail ;
   END else
if Pages.ActivePage=PParams then
   BEGIN
   GParams.GetCellCanvas:=GDrawCell ;
   GParams.RowCount:=2 ; GParams.ColCount:=7+MAX_COL ;
   GParams.ColLengths[3]:=-1 ; GParams.ColWidths[3]:=0 ;
   GParams.ColLengths[2]:=-1 ; GParams.ColWidths[2]:=0 ;
   TL:=TStringList.Create ;
   TL.LoadFromFile(FModeleName);
   Ftitre.text:=TL[0] ; TL.Delete(0) ;
   for i:=1 to TL.Count-1 do
      BEGIN
      St:=TL[i] ;
      GParams.Cells[0,i]:=IntTostr(i) ;
      for kk:=1 to 6+MAX_COL do
        GParams.CellValues[kk,i]:=Trim(ReadTokenPipe(St,'|')) ;
      ih:=ReadTokenI(St) ;
      if ih>0 then GParams.RowHeights[i]:=ih ;

      GParams.RowCount:=GParams.RowCount+1 ;
      END ;
   TL.Free ;
   END ;
if GParams.RowCount>2 then GParams.RowCount:=GParams.RowCount-1 ;
GParams.FixedRows:=1 ;
end;

procedure TFEtatLiasse.bSauveClick(Sender: TObject);
Var i,kk : Integer ;
    TL : TStringList ;
    St : String ;
begin
TL:=TStringList.Create ;
TL.Add(Ftitre.text) ;
for i:=0 to GParams.RowCount-1 do
    BEGIN
    St:='' ;
    for kk:=1 to 6+MAX_COL do
      St:=St+Trim(GParams.CellValues[kk,i])+'|';
    St:=St+IntToStr(GParams.RowHeights[i])+'|' ;
    if Not IsLigneVide(i) then TL.Add(st) ;
    END ;
TL.SaveToFile(FModeleName);
TL.Free ;
end;

function TFEtatLiasse.EncodeFont: string;
begin
Result := FTestFont.Font.Name+';'+IntToStr(FTestFont.Font.Size)+';';
if fsBold in FTestFont.Font.Style then Result := Result + 'B';
if fsItalic in FTestFont.Font.Style then Result := Result + 'I';
if fsUnderline in FTestFont.Font.Style then Result := Result + 'U';
if fsStrikeOut in FTestFont.Font.Style then Result := Result + 'S';
Result :=Result+';'+ColorToString(FTestFont.Font.Color)+';'+ColorToString(FTestFont.Color)+';' ;
if akLeft in FTestFont.Anchors then Result:=Result+'X' else Result:=Result+'-' ;
if akTop in FTestFont.Anchors then Result:=Result+'X' else Result:=Result+'-' ;
if akRight in FTestFont.Anchors then Result:=Result+'X' else Result:=Result+'-' ;
if akBottom in FTestFont.Anchors then Result:=Result+'X' else Result:=Result+'-' ;
Result:=Result+';' ;
end;

procedure TFEtatLiasse.DecodeFont (St : string) ;
Var Frame,StS : String ;

begin
if St='' then St:=Default_Style ;
Try
FTestFont.Font.Name:=ReadTokenSt(St) ;
FTestFont.Font.Size:=ReadTokenI(St) ;
StS:=ReadTokenSt(St) ;
FTestFont.Font.Style:=[] ;
if Pos('B',StS)>0 then FTestFont.Font.Style:=FTestFont.Font.Style+[fsBold] ;
if Pos('I',StS)>0 then FTestFont.Font.Style:=FTestFont.Font.Style+[fsItalic] ;
if Pos('U',StS)>0 then FTestFont.Font.Style:=FTestFont.Font.Style+[fsUnderline] ;
if Pos('S',StS)>0 then FTestFont.Font.Style:=FTestFont.Font.Style+[fsStrikeOut] ;

FTestFont.Font.Color:=StringToColor(ReadTokenSt(St));
FTestFont.Color:=StringToColor(ReadTokenSt(St));
Frame:=ReadTokenSt(St) ; if Length(Frame)<4 then Frame:=Frame+'    ' ;
FTestFont.Anchors:=[] ;
if (Frame[1]='X') then FTestFont.Anchors:=FTestFont.Anchors+[akLeft] ;
if (Frame[2]='X') then FTestFont.Anchors:=FTestFont.Anchors+[akTop] ;
if (Frame[3]='X') then FTestFont.Anchors:=FTestFont.Anchors+[akRight] ;
if (Frame[4]='X') then FTestFont.Anchors:=FTestFont.Anchors+[akBottom] ;
except
End ;
end;


function TFEtatLiasse.IsLigneVide(ARow: Integer): Boolean;
Var i : Integer ;
begin
Result:=TRUE ;
For i:=4 to GParams.ColCount-1 do
  if Trim(GParams.Cells[i,ARow])<>'' then BEGIN Result:=FALSE ; Exit ; END ;
end;

procedure TFEtatLiasse.GParamsRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if (GParams.Row=GParams.RowCount-1)  and (Not IsLigneVide(GParams.Row)) then
  BEGIN
  GParams.RowCount:=GParams.RowCount+1 ;
  GParams.Cells[0,GParams.RowCount-1]:=IntToStr(GParams.RowCount-1) ;
  END ;
if GParams.Cells[2,GParams.Row]='G' then bGauche.Down:=TRUE else
  if GParams.Cells[2,GParams.Row]='D' then bDroite.Down:=TRUE else
                                             bCentre.Down:=TRUE ;
end;

procedure TFEtatLiasse.GParamsRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if (GParams.Row=GParams.RowCount-3) and (ou=GParams.RowCount-2)
   and (IsLigneVide(Ou+1)) then GParams.RowCount:=GParams.RowCount-1 ;
end;

procedure TFEtatLiasse.GDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
Var Style : String ;
    Rect : TRect ;
begin
if (ARow>0) and (ACol=4) then
   BEGIN
   Style:=Gparams.Cells[3,ARow] ;
   if Style='' then Style:=Default_Style ;
   DecodeFont(Style) ; Rect:=GParams.CellRect(ACol,ARow) ;
   GParams.Canvas.Font:=FTestFont.Font ;
   GParams.Canvas.Brush.Color:=FTestFont.Color ;
   GParams.Canvas.Pen.Style:=psSolid ;
   GParams.Canvas.Pen.Color:=clBlack ;
   if akTop in FTestFont.Anchors then BEGIN GParams.Canvas.MoveTo(Rect.Left,Rect.Top) ; GParams.Canvas.LineTo(Rect.Right+1,Rect.Top) ; END ;
//      if akLeft in FTestFont.Anchors then BEGIN GParams.Canvas.MoveTo(Rect.Left,Rect.Top) ; GParams.Canvas.LineTo(Rect.Left,Rect.Bottom+1) ;END ;
//      if akRight in FTestFont.Anchors then BEGIN GParams.Canvas.MoveTo(Rect.Right,Rect.Top) ; GParams.Canvas.LineTo(Rect.Right,Rect.Bottom+1) ; END ;
   if akBottom in FTestFont.Anchors then BEGIN GParams.Canvas.MoveTo(Rect.Left,Rect.Bottom) ; GParams.Canvas.LineTo(Rect.Right+1,Rect.Bottom) ; END ;
   END ;
end;

procedure TFEtatLiasse.GParamsCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if (GParams.Row=GParams.RowCount-1)  and (Not IsLigneVide(GParams.Row)) then
  BEGIN
  GParams.RowCount:=GParams.RowCount+1 ;
  GParams.Cells[0,GParams.RowCount-1]:=IntToStr(GParams.RowCount-1) ;
  END ;
end;

procedure TFEtatLiasse.BDelLigneClick(Sender: TObject);
Var i : Integer ;
begin
if GParams.RowCount<3 then exit ;
GParams.DeleteRow(GParams.Row) ;
for i:=1 to GParams.RowCount-1 do
  GParams.Cells[0,i]:=IntToStr(i) ;
end;

procedure TFEtatLiasse.BInsLigneClick(Sender: TObject);
Var i : Integer ;
begin
GParams.InsertRow(GParams.Row) ;
for i:=1 to GParams.RowCount-1 do
  GParams.Cells[0,i]:=IntToStr(i) ;
GParams.Row:=GParams.Row-1 ;
end;

procedure TFEtatLiasse.GParamsRowMoved(Sender: TObject; FromIndex,  ToIndex: Integer);
Var i : Integer ;
begin
for i:=1 to GParams.RowCount-1 do
  GParams.Cells[0,i]:=IntToStr(i) ;
end;

procedure TFEtatLiasse.bPoliceClick(Sender: TObject);
begin
With FD do
	begin
  DecodeFont(GParams.Cells[3,GParams.Row]) ;
	Font.Assign(FTestFont.Font);
	if Execute then
		  begin
		  FTestFont.Font.Assign(Font);
      GParams.Cells[3,GParams.Row]:=EncodeFont ;
      GParams.InvalidateRow(GParams.Row) ;
		  end;
  end;
end;

procedure TFEtatLiasse.bColorClick(Sender: TObject);
begin
With CD do
	begin
  DecodeFont(GParams.Cells[3,GParams.Row]) ;
	Color:=FTestFont.Color;
	if Execute then
		  begin
		  FTestFont.Color:=Color ;
      GParams.Cells[3,GParams.Row]:=EncodeFont ;
      GParams.InvalidateRow(GParams.Row) ;
		  end;
  end;
end;

procedure TFEtatLiasse.bGaucheClick(Sender: TObject);
begin
if bGauche.Down then  GParams.Cells[2,GParams.Row]:='G' else
  if bDroite.Down then GParams.Cells[2,GParams.Row]:='D' else
                        GParams.Cells[2,GParams.Row]:='C' ;
GParams.InvalidateRow(GParams.Row) ;
end;

procedure TFEtatLiasse.bFrameChange(Sender: TObject);
begin
DecodeFont(GParams.Cells[3,GParams.Row]) ;
FTestFont.Anchors:=[] ;
if bFrame.CurrentChoix in [1,5,7] then FTestFont.Anchors:=FTestFont.Anchors+[aktop] ;
if bFrame.CurrentChoix in [2,6,7] then FTestFont.Anchors:=FTestFont.Anchors+[akLeft] ;
if bFrame.CurrentChoix in [3,5,7] then FTestFont.Anchors:=FTestFont.Anchors+[akBottom] ;
if bFrame.CurrentChoix in [4,6,7] then FTestFont.Anchors:=FTestFont.Anchors+[akRight] ;
GParams.Cells[3,GParams.Row]:=EncodeFont ;
GParams.InvalidateRow(GParams.Row) ;
end;

function TFEtatLiasse.AddGras(St: String): String;
begin
DecodeFont (St) ;
FTestFont.Font.Style:=FTestFont.Font.Style+[fsBold] ;
result:=EncodeFont ;
end;

Procedure Dechiffre(i : Integer ; Var St,St1,St2,St3 : String) ;
Var i1,i2,j : Integer ;
    k1,k2 : Integer ;
    Ok2 : Boolean ;
BEGIN
Inc(i) ;
i1:=Pos('+',St) ; i2:=Pos('-',St) ;
St2:='' ; St1:='' ; St3:='' ;
If St='' Then Exit ; Ok2:=FALSE ;
If (i2>0) Or (i1>0) Then
  BEGIN
  k1:=0 ; For j:=i-1 Downto 1 Do If St[j] in ['+','-'] Then BEGIN k1:=j ; break ; END ;
  k2:=Length(St) ; For j:=i to k2 Do If St[j] in ['+','-'] Then BEGIN Ok2:=TRUE ; k2:=j ; break ; END ;
  If Not Ok2 Then k2:=0 ;
  For j:=1 To Length(St) Do
    BEGIN
    If j<=k1 Then St1:=St1+St[j] ;
    If (J>k1) And ((j<k2) Or (k2=0)) Then St2:=St2+St[j] ;
    If (j>=k2) And (k2>0) Then St3:=St3+St[j] ;
    END ;
  END Else St2:=St ;
END ;

procedure TFEtatLiasse.GParamsDblClick(Sender: TObject);
Var St,St1,St2,St3 : String ;
    i : Integer ;
begin
If GParams.Col>=7 Then
  BEGIN
  St:=GParams.Cells[GParams.Col,GParams.Row] ;
  If (St='') Or (Copy(St,1,1)<>'=') Then
    BEGIN
    If GParams.inplaceEditor<>NIL Then
      BEGIN
      i:=GParams.inplaceEditor.SelStart ;
      Dechiffre(i,St,St1,St2,St3) ;
      END ELse
      BEGIN
      i:=Length(St) ; Dechiffre(i,St,St1,St2,St3) ;
      END ;
    Cache.Text:=St2 ;
    if GChercheCompte(Cache,Nil) then
      BEGIN
      St:=St1+Cache.Text+St3 ;
      GParams.Cells[GParams.Col,GParams.Row]:=St ;
      If i>Length(St) Then i:=Length(St) ;
      GParams.inplaceEditor.SelStart:=i ;
      if GParams.Cells[POS_LIBELLE,GParams.Row]='' Then GParams.Cells[POS_LIBELLE,GParams.Row]:=LibRub.Caption ;
      END ;
    END ;
  END ;
end;

procedure TFEtatLiasse.GParamsKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : boolean ;
begin
OkG:=(Screen.ActiveControl=GParams) ; Vide:=(Shift=[]) ;
Case Key of
  VK_F5     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; GParamsDblClick(NIL) ; END ;
  END ;
end;

procedure TFEtatLiasse.FListeDblClick(Sender: TObject);
Var C,R,C1,R1 : longint ;
    St : String ;
    i : Integer ;
begin
Exit ;
FListe.MouseToCell(GX,GY,C,R) ;
if R<=0 then Exit ;
C1:=C ; R1:=R ;
C1:=POS_COL+C-1 ;
(*
Appel:='RUBRIQUE' ;
  Result:=Get_Cumul(Appel,Nom,Crit.StTypEcr,Crit.Etab,Crit.Devise,FFormuleCol[LaCol],'') ;
  FListe.Cells[POS_COL+i-1,0]:=Crit.Col[i-1].StTitre ;
i:=POS_COL+i-1,0
*)
St:=GParams.Cells[6+C1,R1] ; i:=R1-1 ;
If (St='') Or (Copy(St,1,1)<>'=') Then
  GetBalanceSimple('RUBRIQUE',Fliste.Cells[0,Fliste.Row],Crit.StTypEcr,Crit.Etab,Crit.Devise,
                   Crit.Col[i].Exo.Code,Crit.Col[i].Date1,Crit.Col[i].Date2,False) ;
end;

procedure TFEtatLiasse.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
GX:=X ; GY:=Y ;
end;


procedure TFEtatLiasse.FormClose(Sender: TObject; var Action: TCloseAction);
Var C,R : integer ;
begin
for R:=Fliste.FixedRows to FListe.RowCount-1 do
    for C:=0 to FListe.ColCount-1 do
    BEGIN
    if (FListe.Objects[C,R]<>Nil) then Fliste.Objects[C,R].Free ;
    FListe.Cells[C,R]:='' ; FListe.Objects[C,R]:=nil ;
    END ;
end;

end.
