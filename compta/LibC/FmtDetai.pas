unit FmtDetai;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, Spin, ComCtrls, Buttons, ExtCtrls, DB, {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} DBGrids,
  HSysMenu, hmsgbox, Hqry, Hctrls, HTB97,
  DBCtrls, Mask, DBCGrids, ent1, HEnt1;

Procedure ParamImportDetail(Nature,Import,Code,Lib,Prefixe : String) ;
Procedure ParamImportDetailLibre(Nature,Import,Code,Lib,Prefixe : String) ;

type
  TFFmtDetail = class(TForm)
    TFMTIMPDE: THTable;
    HM: THMsgBox;
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Outils: TToolbar97;
    BUp: TToolbarButton97;
    BDown: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    DockTop: TDock97;
    DockLeft: TDock97;
    DockRight: TDock97;
    Panel1: TPanel;
    FListe: THGrid;
    Ligne97: TToolbar97;
    FLigne: TSpinEdit;
    FChamps: THValComboBox;
    FDecimale: TSpinEdit;
    FLong: TSpinEdit;
    FDebut: TSpinEdit;
    Dock971: TDock97;
    PCorresp: TPanel;
    Panel3: TPanel;
    FCorresp: THGrid;
    Panel4: TPanel;
    BValCorresp: TBitBtn;
    BAnnulCorresp: TBitBtn;
    BAideCorresp: TBitBtn;
    BDroite: TToolbarButton97;
    BCorresp2: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    BCopy: TToolbarButton97;
    ToolbarSep972: TToolbarSep97;
    ToolbarSep973: TToolbarSep97;
    BPrint: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    HMTrad: THSystemMenu;
    OkLibelle: TCheckBox;
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure BDeleteClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure FLigneChange(Sender: TObject);
    procedure FChampsChange(Sender: TObject);
    procedure FDebutChange(Sender: TObject);
    procedure FLongChange(Sender: TObject);
    procedure FDecimaleChange(Sender: TObject);
    procedure BUpClick(Sender: TObject);
    procedure BDownClick(Sender: TObject);
    procedure BCorresp2Click(Sender: TObject);
    procedure BAnnulCorrespClick(Sender: TObject);
    procedure BValCorrespClick(Sender: TObject);
    procedure BCopyClick(Sender: TObject);
    procedure BDroiteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BPrintClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure OkLibelleClick(Sender: TObject);
  private { Déclarations privées }
    OkNewLigne,FModifie,Loading : Boolean ;
    FNat,FImp,FCode,FPrefixe : String ;
    LChamp : TStrings ;
    FLibre : Boolean ;
    procedure Charge ;
    Function  LigneVide(L : Integer) : Boolean ;
    function  TransChamp(NomChamp : String) : String ;
    function  TransPrefixe(CurNat : String) : String ;
    procedure MajDebLong(Diff : Integer) ;
    Function  EstRempliIE( GS : THGrid ; Lig : integer) : Boolean ;
    procedure GereNewLigneIE ( GS : THGrid ) ;
    Procedure GereLibLibre(OkAddLChamp : Boolean) ;
  public  { Déclarations publiques }
  end;

implementation

uses PrintDBG,TImpFic ;

{$R *.DFM}


Procedure ParamImportDetail(Nature,Import,Code,Lib,Prefixe : String) ;
var FFmtDetail: TFFmtDetail;
BEGIN
FFmtDetail:=TFFmtDetail.Create(Application) ;
Try
  FFmtDetail.Caption:=FFmtDetail.Caption+' '+Code+' '+Lib ;
  FFmtDetail.FNat:=Nature ;
  FFmtDetail.FImp:=Import ;
  FFmtDetail.FCode:=Code ;
  FFmtDetail.FPrefixe:=Prefixe ;
  FFmtDetail.FLibre:=FALSE ;
  FFmtDetail.ShowModal ;
  finally
  FFmtDetail.Free ;
  end ;
SourisNormale ;
END ;

Procedure ParamImportDetailLibre(Nature,Import,Code,Lib,Prefixe : String) ;
var FFmtDetail: TFFmtDetail;
BEGIN
FFmtDetail:=TFFmtDetail.Create(Application) ;
Try
  FFmtDetail.Caption:=FFmtDetail.Caption+' '+Code+' '+Lib ;
  FFmtDetail.FNat:=Nature ;
  FFmtDetail.FImp:=Import ;
  FFmtDetail.FCode:=Code ;
  FFmtDetail.FPrefixe:=Prefixe ;
  FFmtDetail.FLibre:=TRUE ;
  FFmtDetail.ShowModal ;
  finally
  FFmtDetail.Free ;
  end ;
SourisNormale ;
END ;


procedure TFFmtDetail.BFermeClick(Sender: TObject);
begin
Close ;
end;

Function TFFmtDetail.LigneVide(L : Integer) : boolean ;
BEGIN
Result:=(FListe.Cells[1,L]='') ;
END ;

procedure TFFmtDetail.BValiderClick(Sender: TObject);
Var l,nb : Integer ;
begin
Loading:=TRUE ;
ExecuteSQL('DELETE FROM FMTIMPDE WHERE ID_NATURE="'+FNat+'" AND ID_IMPORTATION="'+FImp+'" AND ID_CODE="'+FCode+'"') ;
TFMTIMPDE.Open ; nb:=0 ;
For L:=1 to FListe.RowCount-1 do
  if Not LigneVide(L) then
   BEGIN
   Inc(nb) ;
   TFMTIMPDE.Insert ;
   TFMTIMPDE.FindField('ID_NATURE').AsString:=FNat ;
   TFMTIMPDE.FindField('ID_IMPORTATION').AsString:=FImp ;
   TFMTIMPDE.FindField('ID_CODE').AsString:=FCode ;
   TFMTIMPDE.FindField('ID_NUMLIGNE').AsInteger:=StrToInt(FListe.Cells[0,L]) ;
   TFMTIMPDE.FindField('ID_NUMCHAMP').AsInteger:=nb ;
   FChamps.Libelle:=FListe.Cells[1,L] ;
   TFMTIMPDE.FindField('ID_CHAMP').AsString:=FChamps.Value ;
   TFMTIMPDE.FindField('ID_DEBUT').AsInteger:=StrToInt(FListe.Cells[2,L]) ;
   TFMTIMPDE.FindField('ID_LONGUEUR').AsInteger:=StrToInt(FListe.Cells[3,L]) ;
   TFMTIMPDE.FindField('ID_DECIMAL').AsInteger:=StrToInt(FListe.Cells[4,L]) ;
   TFMTIMPDE.FindField('ID_ALIGNEDROITE').AsString:=FListe.Cells[5,L] ;
   TFMTIMPDE.FindField('ID_CORRESP').AsString:=FListe.Cells[6,L] ;
   TFMTIMPDE.Post ;
   END ;
TFMTIMPDE.Close ;
Loading:=FALSE ; FModifie:=False ;
Close ;
end;

procedure TFFmtDetail.Charge ;
Var Q : TQuery ;
    L : Integer ;
BEGIN
Fliste.RowCount:=2 ;
Q:=OpenSQL('SELECT * FROM FMTIMPDE WHERE ID_NATURE="'+FNat+'" AND ID_IMPORTATION="'+FImp+'" AND ID_CODE="'+FCode+'"'+
           ' ORDER BY ID_NATURE, ID_IMPORTATION, ID_CODE, ID_NUMLIGNE, ID_NUMCHAMP',TRUE) ;
While Not Q.EOF do
   BEGIN
   L:=Q.FindField('ID_NUMCHAMP').AsInteger ;
   FListe.Cells[0,L]:=Q.FindField('ID_NUMLIGNE').AsString ;
   FChamps.Value:=Q.FindField('ID_CHAMP').AsString ;
   FListe.Cells[1,L]:=FChamps.Text ;
   FListe.Cells[2,L]:=Q.FindField('ID_DEBUT').AsString ;
   FListe.Cells[3,L]:=Q.FindField('ID_LONGUEUR').AsString ;
   FListe.Cells[4,L]:=Q.FindField('ID_DECIMAL').AsString ;
   FListe.Cells[5,L]:=Q.FindField('ID_ALIGNEDROITE').AsString ;
   FListe.Cells[6,L]:=Q.FindField('ID_CORRESP').AsString ;
   Q.Next ;
   Fliste.RowCount:=Fliste.RowCount+1 ;
   END ;
Ferme(Q) ; FModifie:=False ;
END ;

procedure TFFmtDetail.FormCreate(Sender: TObject);
begin
Fliste.TypeSais:=tsGene ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Loading:=FALSE ;
RegLoadToolbarPos(Self,'FmtImport') ;
end;

function TransTypeChamp(St : String) : String ;
var S : String ;
    p : integer ;
BEGIN
Result:='1' ;
S:=St ;
p:=Pos('VARCHAR(',S) ;
if p<>0 then
  BEGIN
  Result:=Copy(s,p+8,(Length(s)-1)-(p+7)) ;
  END else
if S='COMBO' then
  BEGIN
  Result:='3' ;
  END else
if S='BOOLEAN' then
  BEGIN
  Result:='1' ;
  END else
if S='DATE' then
  BEGIN
  Result:='10' ;
  END else
if S='DOUBLE' then
  BEGIN
  Result:='20' ;
  END else
if (S='INTEGER') or (S='SMALLINT') then
  BEGIN
  Result:='10' ;
  END else
if S='BLOB' then
  BEGIN
  Result:='70' ;
  END ;
END ;

Procedure TFFmtDetail.GereLibLibre(OkAddLChamp : Boolean) ;
Var Ind : tLibreChampExp ;
BEGIN
If FLibre Then For Ind:=_EB_Comptable To _EB_CreditEnCoursEchuNonRegle Do
  BEGIN
  FChamps.Values.Add(LibLibreExp[Ind,lCode]) ; FChamps.Items.Add(LibLibreExp[Ind,lLib]) ;
  If OkAddLChamp Then LChamp.Add('10') ;
  END ;
END ;

procedure TFFmtDetail.FormShow(Sender: TObject);
Var Q : TQuery ;
    SQL : String ;
    F : Boolean ;
begin
Loading:=TRUE ; FChamps.Clear ; FChamps.Values.Add('-') ;
if FImp='X' then
  BEGIN
  FCorresp.Cells[0,0]:=HM.Mess[7] ;
  BCopy.Hint:=HM.Mess[0] ;
  FChamps.Items.Add(HM.Mess[2])
  END else
  BEGIN
//FCorresp.Cells[0,0]:=HM.Mess[8] ;
  BCopy.Hint:=HM.Mess[1] ;
  FChamps.Items.Add(HM.Mess[3]) ;
  END ;
SQL:='Select DH_NOMCHAMP, DH_TYPECHAMP, DH_LIBELLE, DH_CONTROLE, DH_PREFIXE' ;
SQL:=SQL+' From DECHAMPS Where DH_PREFIXE="'+TransPrefixe(FNat)+'" and (DH_CONTROLE like "%L%" OR DH_CONTROLE like "%F%")' ;
SQL:=SQL+' Order by DH_LIBELLE' ;
Q:=OpenSQL(SQL,True) ;
LChamp:=TStringList.Create ; LChamp.Add('1') ; LChamp.Add('1') ; // Constante + XX_SENS
FChamps.Items.Add(HM.Mess[4]) ; FChamps.Values.Add('XX_SENS') ;
While Not Q.EOF do
    BEGIN
    FChamps.Values.Add(Q.Fields[0].AsString) ;
    FChamps.Items.Add(Q.Fields[2].AsString) ;
    LChamp.Add(TransTypeChamp(Q.Fields[1].AsString)) ;
    Q.Next ;
    END ;
Ferme(Q) ;
GereLibLibre(TRUE) ;
Charge ;
FListeRowEnter(Nil,1,F,TRUE);
Loading:=FALSE ; FModifie:=False ; OkNewLigne:=False ;
end;

Function TFFmtDetail.EstRempliIE( GS : THGrid ; Lig : integer) : Boolean ;
BEGIN
Result:=(GS.Cells[4,Lig]<>'') or (GS.Cells[5,Lig]<>'') ;
END ;

procedure TFFmtDetail.GereNewLigneIE ( GS : THGrid ) ;
BEGIN
if EstRempliIE(GS,GS.RowCount-1) then GS.RowCount:=GS.RowCount+1 else
   if Not EstRempliIE(GS,GS.RowCount-2) then GS.RowCount:=GS.RowCount-1 ;
END ;
procedure TFFmtDetail.FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
Loading:=True ;
FLigne.Value:=Round(Valeur(FListe.Cells[0,Ou])) ;
FChamps.Libelle:=FListe.Cells[1,Ou] ;
if (Round(Valeur(FListe.Cells[2,Ou]))=1) and (Ou>1) then FDebut.Value:=Round(Valeur(FListe.Cells[2,Ou]))
                                                    else FDebut.Value:=Round(Valeur(FListe.Cells[2,Ou-1])+Valeur(FListe.Cells[3,Ou-1])) ;
FLong.Value:=Round(Valeur(FListe.Cells[3,Ou])) ;
FDecimale.Value:=Round(Valeur(FListe.Cells[4,Ou])) ;
if FListe.Cells[5,ou]='X' then BDroite.Down:=TRUE else BDroite.Down:=FALSE ;
GereNewLigneIE(Fliste) ;
Loading:=False ; OkNewLigne:=False ;
end;

procedure TFFmtDetail.BDeleteClick(Sender: TObject);
Var C,R,Diff : Integer ;
    F : Boolean ;
begin
Diff:=-Round(Valeur(FListe.Cells[3,Fliste.Row])) ;
FListe.Cells[2,Fliste.Row+1]:=FListe.Cells[2,Fliste.Row] ;
For R:=FListe.Row to FListe.RowCount-2 do
   for C:=0 to FListe.ColCount-1 do FListe.Cells[C,R]:=FListe.Cells[C,R+1] ;
for C:=0 to FListe.ColCount-1 do FListe.Cells[C,FListe.RowCount-1]:='' ;
FListeRowEnter(Nil,FListe.Row,F,TRUE);
FModifie:=True ;
MajDebLong(Diff) ;
end;

procedure TFFmtDetail.BInsertClick(Sender: TObject);
Var C,R : Integer ;
    F : Boolean ;
begin
For R:=FListe.RowCount-1 downto FListe.Row+1 do
   for C:=0 to FListe.ColCount-1 do FListe.Cells[C,R]:=FListe.Cells[C,R-1] ;
for C:=0 to FListe.ColCount-1 do FListe.Cells[C,FListe.Row]:='' ;
FListeRowEnter(Nil,FListe.Row,F,TRUE);
FModifie:=True ; OkNewLigne:=True ;
end;

procedure TFFmtDetail.FLigneChange(Sender: TObject);
begin
if Loading then exit ;
FListe.Cells[0,FListe.Row]:=FLigne.Text ;
end;

procedure TFFmtDetail.FChampsChange(Sender: TObject);
var Diff : Integer ;
begin
if Loading then exit ;
if FListe.Row=FListe.RowCount-1 then FListe.RowCount:=Fliste.RowCount+1 ;
FListe.Cells[1,FListe.Row]:=FChamps.Text ;
if FListe.Cells[0,FListe.Row]='' then FListe.Cells[0,FListe.Row]:='1' ;
if FListe.Cells[2,FListe.Row]='' then
   BEGIN
   if FListe.Row=1 then FListe.Cells[2,FListe.Row]:='1'
                   else FListe.Cells[2,FListe.Row]:=FloatToStr(Valeur(FListe.Cells[2,FListe.Row-1])+Valeur(FListe.Cells[3,FListe.Row-1])) ;
   END ;
Diff:=0 ;
if OkNewLigne or (FListe.Cells[3,FListe.Row]='') then
  BEGIN
  OkNewLigne:=True ;
  FLong.Value:=Round(Valeur(LChamp[FChamps.ItemIndex])) ;
  Diff:=FLong.Value-Round(Valeur(FListe.Cells[3,FListe.Row])) ;
  FListe.Cells[3,FListe.Row]:=FloatToStr(Valeur(LChamp[FChamps.ItemIndex])) ;
  END ;
if FListe.Cells[4,FListe.Row]='' then FListe.Cells[4,FListe.Row]:='0' ;
if FListe.Cells[5,FListe.Row]='' then FListe.Cells[5,FListe.Row]:='-' ;
MajDebLong(Diff) ;
FModifie:=True ;
end;

procedure TFFmtDetail.MajDebLong(Diff : Integer) ;
var i : Integer ;
BEGIN
for i:=FListe.Row+1 to FListe.RowCount-1 do
  BEGIN
  if FListe.Cells[2,i]<>'' then FListe.Cells[2,i]:=IntToStr(Round(Valeur(FListe.Cells[2,i])+Diff)) ;
  END ;
FModifie:=True ;
END ;

procedure TFFmtDetail.FDebutChange(Sender: TObject);
var Diff : Integer ;
begin
if Loading then exit ;
if (FListe.Cells[2,FListe.Row]='') or (FDebut.Text='') then exit ;
Diff:=StrToInt(FDebut.Text)-StrToInt(FListe.Cells[2,FListe.Row]) ;
FListe.Cells[2,FListe.Row]:=FDebut.Text ;
MajDebLong(Diff) ;
end;

procedure TFFmtDetail.FLongChange(Sender: TObject);
var Diff : Integer ;
begin
if Loading then exit ;
if (FListe.Cells[3,FListe.Row]='') or (FLong.Text='') then exit ;
Diff:=StrToInt(FLong.Text)-StrToInt(FListe.Cells[3,FListe.Row]) ;
FListe.Cells[3,FListe.Row]:=FLong.Text ;
MajDebLong(Diff) ;
end;

procedure TFFmtDetail.FDecimaleChange(Sender: TObject);
begin
if Loading then exit ;
FListe.Cells[4,FListe.Row]:=FDecimale.Text ;
FModifie:=True ;
end;

procedure TFFmtDetail.BUpClick(Sender: TObject);
begin
if FListe.Row>Fliste.FixedRows then FListe.ExchangeRow(FListe.Row,FListe.Row-1) else Exit ;
FListe.Row:=FListe.Row-1;
FListe.Cells[2,FListe.Row]:=FListe.Cells[2,FListe.Row+1] ;
FListe.Cells[2,FListe.Row+1]:=IntToStr(Round(Valeur(FListe.Cells[2,FListe.Row]))+Round(Valeur(FListe.Cells[3,FListe.Row])))  ;
FModifie:=True ;
end;

procedure TFFmtDetail.BDownClick(Sender: TObject);
begin
if FListe.Row<FListe.RowCount-1 then FListe.ExchangeRow(FListe.Row,FListe.Row+1) else Exit ;
FListe.Row:=FListe.Row+1;
FListe.Cells[2,FListe.Row-1]:=FListe.Cells[2,FListe.Row] ;
FListe.Cells[2,FListe.Row]:=IntToStr(Round(Valeur(FListe.Cells[2,FListe.Row-1]))+Round(Valeur(FListe.Cells[3,FListe.Row-1]))) ;
FModifie:=True ;
end;

procedure TFFmtDetail.BCorresp2Click(Sender: TObject);
Var R : Integer ;
    St,St1,St2 : String ;
begin
PCorresp.Visible:=TRUE ;
FListe.Enabled:=FALSE ;
For R:=1 to FCorresp.RowCount-1 do
   BEGIN
   FCorresp.Cells[0,R]:='' ; FCorresp.Cells[1,R]:='' ;
   END ;
St:=FListe.Cells[6,FListe.Row] ; R:=0 ;
While St<>'' do
   BEGIN
   Inc(R) ;
   St1:=ReadTokenSt(St) ; FCorresp.Cells[0,R]:=St1 ;
   St2:=ReadTokenSt(St) ; FCorresp.Cells[1,R]:=St2 ;
   END ;
FModifie:=True ;
end;

procedure TFFmtDetail.BAnnulCorrespClick(Sender: TObject);
begin
PCorresp.Visible:=FALSE ;
FListe.Enabled:=TRUE ;
end;

procedure TFFmtDetail.BValCorrespClick(Sender: TObject);
Var R : Integer ;
    St : String ;
begin
St:='' ;
For R:=1 to FCorresp.RowCount-1 do
 if (Pos('#',FCorresp.Cells[0,R])<>0) then
   BEGIN
   St:=St+FCorresp.Cells[0,R] ;
   END else
   if ((FCorresp.Cells[0,R]<>'') And (FCorresp.Cells[1,R]<>'')) then
     BEGIN
     St:=St+FCorresp.Cells[0,R]+';'+FCorresp.Cells[1,R]+';' ;
     END ;
FListe.Cells[6,FListe.Row]:=St;
BAnnulCorrespClick(Nil);
end;

procedure TFFmtDetail.BDroiteClick(Sender: TObject);
begin
if Loading then exit ;
if BDroite.Down then FListe.Cells[5,FListe.Row]:='X' else FListe.Cells[5,FListe.Row]:='-' ;
Fmodifie:=True ;
end;

function TFFmtDetail.TransChamp(NomChamp : String) : String ;
var p : Integer ;
    Pref : String ;
BEGIN
Result:=NomChamp ;
p:=Pos('_',NomChamp) ;
if p<>0 then Pref:=Copy(NomChamp,1,p-1) ;
//if (Pref<>'IE') or (Pref<>'E') or
if (p=0) or (Pref='XX') then Exit ;
if (Pref='IE') and (FImp='-') then Pref:=FPrefixe ;
if ((FPrefixe='BE') or (FPrefixe='Y') or (FPrefixe='E')) and (FImp='X') then Pref:='IE' ;
Result:=Pref+Trim(Copy(NomChamp,p,length(NomChamp)-p+1)) ;
if (Pref='IE') and (FImp='X') then
  BEGIN
  if NomChamp=FPrefixe+'_NUMEROPIECE' then Result:='IE_NUMPIECE' ;
  if NomChamp='BE_BUDJAL' then Result:='IE_JOURNAL' ;
  if NomChamp='BE_BUDGENE' then Result:='IE_GENERAL' ;
  if NomChamp='BE_BUDSECT' then Result:='IE_SECTION' ;
  if NomChamp='BE_NATUREBUD' then Result:='IE_NATUREPIECE' ;
  END ;
if (Pref=FPrefixe) and (FImp='-') then
  BEGIN
  if NomChamp='IE_NUMPIECE' then Result:=FPrefixe+'_NUMEROPIECE' ;
  if FPrefixe='BE' then
    BEGIN
    if NomChamp='IE_JOURNAL' then Result:=FPrefixe+'_BUDJAL' ;
    if NomChamp='IE_GENERAL' then Result:=FPrefixe+'_BUDGENE' ;
    if NomChamp='IE_SECTION' then Result:=FPrefixe+'_BUDSECT' ;
    if NomChamp='IE_NATUREPIECE' then Result:=FPrefixe+'_NATUREBUD' ;
    END ;
  END ;
END ;

procedure TFFmtDetail.BCopyClick(Sender: TObject);
var Q : TQuery ;
    NatIE,FIE : String ;
    M,M1 : Integer ;
begin
M:=5 ; M1:=10 ;
if FImp='-' then BEGIN Inc(M) ; Inc(M1) ; END ;
// Import dans Impecr / Export dans Ecriture
FIE:='X' ; NatIE:=FNat ;
if FImp='X' then FIE:='-' ;
Q:=OpenSQL('SELECT * FROM FMTIMPDE WHERE ID_NATURE="'+NatIE+'" AND ID_IMPORTATION="'+FIE+'" AND ID_CODE="'+FCode+'"'+
           ' ORDER BY ID_NATURE, ID_IMPORTATION, ID_CODE, ID_NUMLIGNE, ID_NUMCHAMP',TRUE) ;
if Q.Eof then BEGIN HM.Execute(M1,Caption,'') ; Ferme(Q) ; Exit ; END ;
if HM.Execute(M,Caption,'')<>mrYes then BEGIN Ferme(Q) ; Exit ; END ;
Loading:=True ;
ExecuteSQL('DELETE FROM FMTIMPDE WHERE ID_NATURE="'+FNat+'" AND ID_IMPORTATION="'+FImp+'" AND ID_CODE="'+FCode+'"') ;
TFMTIMPDE.Open ;
While Not Q.EOF do
  BEGIN
  TFMTIMPDE.Insert ;
  TFMTIMPDE.FindField('ID_NATURE').AsString:=FNat ;
  TFMTIMPDE.FindField('ID_IMPORTATION').AsString:=FImp ;
  TFMTIMPDE.FindField('ID_CODE').AsString:=Q.FindField('ID_CODE').AsString ;
  TFMTIMPDE.FindField('ID_NUMLIGNE').AsInteger:=Q.FindField('ID_NUMLIGNE').AsInteger ;
  TFMTIMPDE.FindField('ID_NUMCHAMP').AsInteger:=Q.FindField('ID_NUMCHAMP').AsInteger ;
  TFMTIMPDE.FindField('ID_CHAMP').AsString:=TransChamp(Q.FindField('ID_CHAMP').AsString) ;
  TFMTIMPDE.FindField('ID_DEBUT').AsInteger:=Q.FindField('ID_DEBUT').AsInteger ;
  TFMTIMPDE.FindField('ID_LONGUEUR').AsInteger:=Q.FindField('ID_LONGUEUR').AsInteger ;
  TFMTIMPDE.FindField('ID_DECIMAL').AsInteger:=Q.FindField('ID_DECIMAL').AsInteger ;
  TFMTIMPDE.FindField('ID_CORRESP').AsString:=Q.FindField('ID_CORRESP').AsString ;
  TFMTIMPDE.FindField('ID_ALIGNEDROITE').AsString:=Q.FindField('ID_ALIGNEDROITE').AsString ;
  TFMTIMPDE.Post ; Q.Next ;
  END ;
TFMTIMPDE.Close ;
Ferme(Q) ; Charge ; Loading:=False ; FModifie:=True ;
end;

procedure TFFmtDetail.FormClose(Sender: TObject; var Action: TCloseAction);
begin
RegSaveToolbarPos(Self,'FmtImport') ;
end;

procedure TFFmtDetail.BPrintClick(Sender: TObject);
begin
PrintDBGrid (FListe,nil,Caption,'') ;
end;

procedure TFFmtDetail.BAnnulerClick(Sender: TObject);
begin
Charge ; FModifie:=False ;
end;

procedure TFFmtDetail.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var OkG : boolean ;
begin
OkG:=(ActiveControl=FListe) ;
if Not OkG then Exit ;
Case Key of
  VK_DOWN   : if (FListe.Row=FListe.RowCount-1) and not LigneVide(Fliste.Row) then
                 BEGIN
                 FListe.RowCount:=Fliste.RowCount+1 ;
                 FListe.Row:=FListe.Row+1 ;
                 Key:=0 ;
                 END ;
  VK_INSERT : BEGIN BInsertClick(FListe) ; Key:=0 ; END ;
  VK_DELETE : BEGIN BDeleteClick(FListe) ; Key:=0 ; END ;
  END ;
end;

procedure TFFmtDetail.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
if FModifie then
  case HM.Execute(9,Caption,'') of
    mrYes    : BValiderClick(nil);
    mrNo     : ;
    mrCancel : CanClose:=False ;
    END ;
end;

function TFFmtDetail.TransPrefixe(CurNat : String) : String ;
BEGIN
Result:='' ;
if CurNat='FGE' then Result:='G' else
 if CurNat='FTI' then Result:='T' else
  if CurNat='FSE' then Result:='S' else
   if CurNat='FBJ' then Result:='BJ' else
    if CurNat='FBG' then Result:='BG' else
     if CurNat='FBS' then Result:='BS' else
      if CurNat='FEC' then Result:='E' else
       if CurNat='FOD' then Result:='Y' else
        if CurNat='FBE' then Result:='BE' else
         if CurNat='FBA' then Result:='E' else;
          if CurNat='FCC' then Result:='T' ;
if (FImp='X') and ((CurNat='FEC') or (CurNat='FBA') or (CurNat='FBE') or (CurNat='FOD')) then Result:='IE' ;
END ;

procedure TFFmtDetail.OkLibelleClick(Sender: TObject);
Var SQL : String ;
    Q : TQuery ;
    StSource,StDestination : String ;
    i : Integer ;
    F : Boolean ;
    Ind : tLibreChampExp ;
begin
SQL:='Select DH_NOMCHAMP, DH_TYPECHAMP, DH_LIBELLE, DH_CONTROLE, DH_PREFIXE,DH_NOMCHAMP ' ;
SQL:=SQL+' From DECHAMPS Where DH_PREFIXE="'+TransPrefixe(FNat)+'" and (DH_CONTROLE like "%L%" OR DH_CONTROLE like "%F%")' ;
SQL:=SQL+' Order by DH_LIBELLE' ;
Q:=OpenSQL(SQL,True) ;
FChamps.Clear ;
if FImp='X' then FChamps.Items.Add(HM.Mess[2]) else FChamps.Items.Add(HM.Mess[3]) ;
FChamps.Items.Add(HM.Mess[4]) ; FChamps.Values.Add('XX_SENS') ;
While Not Q.EOF do
    BEGIN
    FChamps.Values.Add(Q.Fields[0].AsString) ;
    If OkLibelle.Checked Then
       BEGIN
       FChamps.Items.Add(Q.Fields[2].AsString) ;
       StSource:=Q.Fields[5].AsString ;
       StDestination:=Q.Fields[2].AsString ;
       END Else
       BEGIN
       FChamps.Items.Add(Q.Fields[5].AsString) ;
       StSource:=Q.Fields[2].AsString ;
       StDestination:=Q.Fields[5].AsString ;
       END ;
    For i:=0 To FListe.RowCount-1 Do
      BEGIN
      If Fliste.Cells[1,i]=StSource Then Fliste.Cells[1,i]:=StDestination ;
      END ;
    Q.Next ;
    END ;
Ferme(Q) ;
If FLibre Then For Ind:=_EB_Comptable To _EB_CreditEnCoursEchuNonRegle Do
  BEGIN
  If OkLibelle.Checked Then
    BEGIN
    FChamps.Items.Add(LibLibreExp[Ind,lLib]) ;
    StSource:=LibLibreExp[Ind,lCode] ;
    StDestination:=LibLibreExp[Ind,lLib] ;
    END Else
    BEGIN
    FChamps.Items.Add(LibLibreExp[Ind,lCode]) ;
    StSource:=LibLibreExp[Ind,lLib] ;
    StDestination:=LibLibreExp[Ind,lCode] ;
    END ;
  For i:=0 To FListe.RowCount-1 Do
    BEGIN
    If Fliste.Cells[1,i]=StSource Then Fliste.Cells[1,i]:=StDestination ;
    END ;
  END ;
Fchamps.Refresh ;
FListeRowEnter(Nil,Fliste.Row,F,TRUE);
end;

end.
