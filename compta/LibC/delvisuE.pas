unit DelVisuE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls, SaisUtil, SaisComm,
  hmsgbox,  HQry , Ent1, UTOB,
  {$IFNDEF IMP}
  SaisODA,   //Thl 26/01/2007
  {$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  {$IFNDEF IMP}
   Saisie,
   SaisBor,
  {$ENDIF}
  LettUtil,
  HSysMenu;

type
  TFDelVisuE = class(TForm)
    HP6: TPanel;
    BAnnuler: THBitBtn;
    BAide: THBitBtn;
    BImprimer: THBitBtn;
    H_NB: TLabel;
    GDEL: THGrid;
    HPieces: THMsgBox;
    HCaption: THMsgBox;
    BZoomPiece: THBitBtn;
    HButton: THMsgBox;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure GDELDblClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
  private
    procedure ComComLett ( GDEL : THGrid ; X : T_RAPPAUTO ) ;
  public
    QuelEcr : TTypeEcr ;
    Quoi    : integer ;
    TPiece  : TList ;
  end ;

procedure VisuPiecesGenere ( TPiece : TList ; QuelEcr : TTypeEcr ; Quoi : integer ) ;

implementation

{$R *.DFM}


uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  {$IFDEF EAGLCLIENT}
  UtileAGL,
  {$ELSE}
  PrintDBG,
  {$ENDIF}
{$IFNDEF CMPGIS35}
  CPJUSTISOL_TOF,
{$ENDIF}
  HEnt1 ;



procedure VisuPiecesGenere ( TPiece : TList ; QuelEcr : TTypeEcr ; Quoi : integer ) ;
Var X : TFDelVisuE ;
BEGIN
X:=TFDelVisuE.Create(Application) ;
 Try
  X.TPiece:=TPiece ; X.QuelEcr:=QuelEcr ; X.Quoi:=Quoi ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFDelVisuE.ComComLett ( GDEL : THGrid ; X : T_RAPPAUTO ) ;
BEGIN
GDEL.Cells[0,GDEL.RowCount-1]:=X.General       ; GDEL.Cells[1,GDEL.RowCount-1]:=X.Auxiliaire ;
GDEL.Cells[2,GDEL.RowCount-1]:=IntToStr(X.NbD) ; GDEL.Cells[3,GDEL.RowCount-1]:=FormatFloat(GDEL.ColFormats[3],X.TotalD) ;
GDEL.Cells[4,GDEL.RowCount-1]:=IntToStr(X.NbC) ; GDEL.Cells[5,GDEL.RowCount-1]:=FormatFloat(GDEL.ColFormats[5],X.TotalC) ;
GDEL.Cells[6,GDEL.RowCount-1]:=X.Devise        ;
GDEL.Cells[7,GDEL.RowCount-1]:=X.CodeA         ; GDEL.Cells[8,GDEL.RowCount-1]:=X.CodeZ ;
END ;

procedure TFDelVisuE.FormShow(Sender: TObject);
Var i : integer ;
    O : TOB ;
    X : T_RAPPAUTO ;
begin
Caption:=HPieces.Mess[Quoi] ;
H_Nb.Caption:=HCaption.Mess[Quoi]+' '+IntToStr(TPiece.Count) ;
UpdateCaption(Self) ;
Case Quoi of
        0 : BEGIN
            if QuelEcr=EcrGen then GDEL.ListeParam:='SUPPRECRS' else GDEL.ListeParam:='SUPPRANA' ;
            BZoomPiece.Visible:=False ;
            END ;
      1,2 : GDEL.ListeParam:='VALIDSIMU' ;
3,4,14,19 : GDEL.ListeParam:='CPZOOMMVT' ;
  5,12,13 : GDEL.ListeParam:='CPZOOMANAL' ;
        6 : BEGIN GDEL.ListeParam:='LETTAUTO' ; BZoomPiece.Hint:=HButton.Mess[0] ; HelpContext:=7514100 ; END ;
        7 : GDEL.ListeParam:='CPZOOMMVT' ;
        8 : BEGIN GDEL.ListeParam:='CPZOOMMVT' ; HelpContext:=7244800 ; END ;
        9 : BEGIN GDEL.ListeParam:='SUPPRBUD' ; BZoomPiece.Visible:=False ; END ;
 10,11,15 : GDEL.ListeParam:='CPZOOMMVT' ;
 {b Thl 10/05/2006}
     16,17 : BEGIN
            GDEL.ListeParam:='SUPPRECRS';
            BZoomPiece.Visible:=False ;
            END ;
        18 : GDEL.ListeParam:='CPZOOMTP' ; //YMO 09/01/2006 Ajout cas sur tiers payeurs
 {e Thl}
   END ;
GDEL.RowCount:=2 ;
if Quoi<>6 then BEGIN
   for i:=0 to TPiece.Count-1 do
       BEGIN
       O:=TOB(TPiece[i]) ;
       if O=nil then continue ;
       ComCom1(GDEL,O) ;
       GDEL.RowCount:=GDEL.RowCount+1 ;
       END ;
   END else
   BEGIN
   for i:=0 to TPiece.Count-1 do
       BEGIN
       X:=T_RAPPAUTO(TPiece[i]) ; ComComLett(GDEl,X) ;
       GDEL.RowCount:=GDEL.RowCount+1 ;
       END ;
  BImprimer.Visible := False ;
   END ;
end;

procedure TFDelVisuE.GDELDblClick(Sender: TObject);
begin if BZoomPiece.Visible then BZoomPieceClick(Nil) ; end;

procedure TFDelVisuE.BZoomPieceClick(Sender: TObject);
{$IFNDEF IMP}
Var
    Q     : TQuery ;
    Trouv : boolean ;
    O     : TOB ;
    A     : TActionFiche ;
    M     : RMVT ;
    SQL,StGen,StAux : String ;
{$ENDIF}
begin
{$IFNDEF IMP}
if ((GDEL.Row<1) or (GDEL.Row>=GDEL.RowCount-1)) then Exit ;
A:=taModif ; if Quoi in [3,4,8,14,15] then A:=taConsult; // 14463 
if Quoi=6 then
   BEGIN
   StGen:=GDEL.Cells[0,GDEL.Row] ; StAux:=GDEL.Cells[1,GDEL.Row] ;
{$IFDEF COMPTA}
   if StAux<>'' then JustifSolde(StAux,fbAux) else if StGen<>'' then JustifSolde(StGen,fbGene) ;
{$ENDIF}
   Exit ;
   END ;
O:=TOB(TPiece[GDEL.Row-1]) ; if O=Nil then Exit ;
if QuelEcr=EcrGen then
   BEGIN
   SQL:='SELECT ' + SQLForIdent( fbGene ) + ' FROM ECRITURE WHERE E_JOURNAL="'+O.GetValue('E_JOURNAL')+'"'
       +' AND E_EXERCICE="'+QuelExo(DateToStr(O.GetValue('E_DATECOMPTABLE')))+'"'
       +' AND E_DATECOMPTABLE="'+USDATETIME(O.GetValue('E_DATECOMPTABLE'))+'"'
       +' AND E_QUALIFPIECE="'+O.GetValue('E_QUALIFPIECE')+'"'
       +' AND E_NUMEROPIECE='+IntToStr(O.GetValue('E_NUMEROPIECE')) ;
   Q:=OpenSQL(SQL,True) ;
   Trouv:=Not Q.EOF ; if Trouv then M:=MvtToIdent(Q,fbGene,False) ; Ferme(Q) ;
   if Trouv then
     BEGIN
     if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>'')) then
      {$IFDEF EAGLCLIENT} {26.06.07 FQ20831 YMO Version eAGL}
      LanceSaisieFolioOBM(O,taConsult)
      {$ELSE}
      LanceSaisieFolioOBM(O,taConsult)
      {$ENDIF}
     Else
      LanceSaisie(Nil,A,M) ;
     END ;
   END else if QuelEcr=EcrBud then
   BEGIN
   END else
   BEGIN
   SQL:='SELECT ' + SQLForIdent( fbSect ) + ' FROM ANALYTIQ WHERE Y_JOURNAL="'+O.GetValue('Y_JOURNAL')+'"'
       +' AND Y_EXERCICE="'+QuelExo(DateToStr(O.GetValue('Y_DATECOMPTABLE')))+'"'
       +' AND Y_DATECOMPTABLE="'+USDATETIME(O.GetValue('Y_DATECOMPTABLE'))+'"'
       +' AND Y_QUALIFPIECE="'+O.GetValue('Y_QUALIFPIECE')+'"'
       +' AND Y_NUMEROPIECE='+IntToStr(O.GetValue('Y_NUMEROPIECE'))
       +' AND Y_AXE="'+O.GetValue('Y_AXE')+'"'
       +' AND Y_GENERAL="'+O.GetValue('Y_GENERAL')+'"' ;
   Q:=OpenSQL(SQL,True) ;
   Trouv:=Not Q.EOF ; if Trouv then M:=MvtToIdent(Q,fbSect,False) ; Ferme(Q) ;
   if Trouv then
      {$IFDEF EAGLCLIENT}
      LanceSaisieODA(Nil, taConsult, M) ;  //Thl 26/01/2007
      {$ELSE}
      LanceSaisieODA(Nil,taModif,M) ;
      {$ENDIF}
   END ;
{$ENDIF}
end;

procedure TFDelVisuE.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFDelVisuE.BImprimerClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  i : Integer;
  szSQL : String;
  O : TOB;
{$ENDIF}  
begin
{$IFDEF EAGLCLIENT}
  szSQL := 'WHERE ';
  for i := 0 to TPiece.Count-1 do begin
    O:=TOB(TPiece[i]); if O=Nil then Exit;
    if i > 0 then szSQL := szSQL + ' OR ';
    if QuelEcr=EcrGen then begin
      szSQL := szSQL + '(E_JOURNAL="'+ O.GetValue('E_JOURNAL')+'"'
                     + ' AND E_EXERCICE="'+QuelExo(DateToStr(O.GetValue('E_DATECOMPTABLE')))+'"'
                     + ' AND E_NUMEROPIECE='+IntToStr(O.GetValue('E_NUMEROPIECE')) +')';
      end
    else begin
      szSQL := szSQL + '(Y_JOURNAL="'+ O.GetValue('Y_JOURNAL')+'"'
                     + ' AND Y_EXERCICE="'+QuelExo(DateToStr(O.GetValue('Y_DATECOMPTABLE')))+'"'
                     + ' AND Y_NUMEROPIECE='+IntToStr(O.GetValue('Y_NUMEROPIECE')) +')';
    end;
  end;
  PrintDBGrid(Caption,GDEL.ListeParam,szSQL,'');
{$ELSE}
  PrintDBGrid(GDEL,Nil,Caption,'') ;
{$ENDIF}
end;

procedure TFDelVisuE.FormClose(Sender: TObject; var Action: TCloseAction);
begin
GDEL.VidePile(True) ;
end;

procedure TFDelVisuE.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

end.
