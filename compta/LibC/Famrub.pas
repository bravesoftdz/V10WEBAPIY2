unit FamRub;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, ExtCtrls, StdCtrls, Buttons, hmsgbox, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, Ent1,
  HStatus, HSysMenu,HEnt1, CpteUtil, ADODB ;

Procedure ParametrageFamilleRubrique(Laquelle,LeLib : String ; Var Fam,LibFam : String ; Budget : Boolean ; QueFListe2 : Boolean = FALSE) ;
Procedure RempliComboTypRub(C : THValComboBox) ;
Procedure RempliComboFamRub(C : THValComboBox) ;

Type PColor = Class
        Drapeau : String ;
        END ;

type
  TFFamrub = class(TForm)
    PBouton   : TPanel;
    MsgBox    : THMsgBox;
    Qrf       : THQuery;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    Pappli: TPanel;
    Fliste1: THGrid;
    FListe2: THGrid;
    Nb1: TLabel;
    Tex1: TLabel;
    Panel2: TPanel;
    Nb2: TLabel;
    Tex2: TLabel;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Fliste1MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListe2DrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
    procedure FormCreate(Sender: TObject);
    procedure Fliste1DblClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Laquelle,LeLib : String ;
    LaFam,LibFam   : String ;
    Budget : Boolean ;
    WMinX,WMinY    : Integer ;
    OldBrush       : TBrush ;
    OldPen         : TPen ;
    QueFListe2 : Boolean ;
    Procedure RempliGrid(G : THGrid) ;
    Procedure MajFamilleRubrique ;
    Procedure SelectionneFamille ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure CompteElemSelectionner(Sender : TObject) ;
    procedure CocheDecoche( Sender : TObject ) ;
    Procedure FaitRequeteMaj ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Procedure ParametrageFamilleRubrique(Laquelle,LeLib : String ; Var Fam,LibFam : String ; Budget : Boolean  ; QueFListe2 : Boolean = FALSE) ;
var FFamrub: TFFamrub;
BEGIN
FFamrub:=TFFamrub.Create(Application) ;
 Try
 if (Laquelle<>'') Or QueFListe2 Then FFamrub.WMinX:=260
                 else FFamrub.WMinX:=630 ;
 FFamrub.WMinY:=FFamrub.Height ;
 FFamrub.Laquelle:=Laquelle ;
 FFamrub.LeLib:=LeLib ;
 FFamrub.LaFam:=Fam ;
 FFamrub.LibFam:=LibFam ;
 FFamrub.Budget:=Budget ;
 FFamrub.QueFliste2:=QueFliste2 ;
 FFamrub.ShowModal ;
 Finally
 Fam:=FFamrub.LaFam ;
 LibFam:=FFamrub.LibFam ;
 FFamrub.Free ;
 End ;
END ;

Procedure RempliComboTypRub(C : THValComboBox) ;
BEGIN
FactoriseComboTypeRub(C) ;
END ;

Procedure RempliComboFamRub(C : THValComboBox) ;
Var Q : TQuery ;
    Okok : Boolean ;
BEGIN
Q:=OpenSql('Select CO_CODE,CO_LIBELLE from COMMUN Where CO_TYPE="RBB"',True) ;
C.Values.Clear ;
C.Items.Clear ;
While not Q.Eof do
  BEGIN
  Okok:=True ;
  if Q.Fields[0].AsString='GEN' then C.Values.Add('CBG') else
     if Q.Fields[0].AsString='ANA' then C.Values.Add('CBS') else
        if Q.Fields[0].AsString='G/A' then
           BEGIN
           if EstSerie(S3) then Okok:=False else C.Values.Add('G/S') ;
           END else if Q.Fields[0].AsString='A/G' then
           BEGIN
           if EstSerie(S3) then Okok:=False else C.Values.Add('S/G') ;
           END ;
  if Okok then C.Items.Add(Q.FindField('CO_LIBELLE').AsString) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

procedure TFFamrub.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFFamrub.FormShow(Sender: TObject);
Var Sql : String ;
begin
if Budget then Sql:='Select CO_CODE,CO_LIBELLE From COMMUN Where CO_TYPE="RBB" Order by CO_CODE '
//          else Sql:='Select CC_CODE,CC_LIBELLE From CHOIXCOD Where CC_TYPE="RBF" Order by CC_CODE ' ;
{ Modif CA le 18/04/2001 }
          else Sql:='Select YDS_CODE,YDS_LIBELLE From CHOIXDPSTD Where YDS_TYPE="RBF" Order by YDS_CODE ' ;
{ Fin Modif CA le 18/04/2001 }
Qrf.Close ; Qrf.Sql.Clear ; Qrf.Sql.Add(Sql) ; ChangeSQL(Qrf) ; Qrf.Open ;
RempliGrid(FListe2) ;
if (Laquelle ='') And (Not QueFListe2) then
   BEGIN
   FaitRequeteMaj ; RempliGrid(FListe1) ; FListe1.SetFocus ;
   HelpContext := 7775200 ;
   END else
   BEGIN
   Fliste1.Visible:=False ; Caption:=MsgBox.Mess[2]+' '+Laquelle+' '+LeLib ;
   Width:=260 ; FListe2.Align:=alClient ; Nb1.Visible:=False ; Tex1.Visible:=False ;
   SelectionneFamille ; FListe2.SetFocus ;
   HelpContext := 7775000 ;
   END ;
UpdateCaption(Self) ;
end;

Procedure TFFamrub.SelectionneFamille ;
Var St,St1 : String ;
    i : Byte ;
BEGIN
if LaFam='' then Exit ;
St:=Lafam ;
While St<>'' do
   BEGIN
   St1:=ReadTokenSt(St) ;
   for i:=1 to Fliste2.RowCount-1 do
      if Fliste2.Cells[0,i]=St1 then
         BEGIN PColor(FListe2.Objects[0,i]).Drapeau:='*' ; Break ; END ;
   END ;
CompteElemSelectionner(FListe2) ;
END ;

Procedure TFFamrub.RempliGrid(G : THGrid) ;
Var i : Integer ;
    X : PColor ;
    St : String ;
BEGIN
i:=1 ; G.VidePile(True) ;
While Not Qrf.EOF do
    BEGIN
    X:=PColor.Create ; X.Drapeau:='' ;
    St:=Qrf.Fields[0].AsString ;
    if Budget And (UpperCase(G.Name)='FLISTE2') then
       BEGIN
       if St='ANA' then St:='CBS' else
          if St='A/G' then St:='S/G' else
             if St='GEN' then St:='CBG' else
                 if St='G/A' then St:='G/S' ;
       END ;
    G.Cells[0,i]:=St ;
    G.Objects[0,i]:=X ;
    G.Cells[1,i]:=Qrf.Fields[1].AsString ;
    if UpperCase(G.Name)='FLISTE1' then G.Cells[2,i]:=Qrf.Fields[2].AsString ;
    G.RowCount:=G.RowCount+1 ; Inc(i) ;
    Qrf.Next ;
    END ;
if G.RowCount>2 then G.RowCount:=G.RowCount-1 ;
END ;

procedure TFFamrub.CocheDecoche( Sender : TObject ) ;
Var Grille : THGrid ;
BEGIN
if Sender=Nil then Exit ; Grille:=THGrid(Sender) ;
if PColor(Grille.Objects[0,Grille.Row]).Drapeau='*'
   then PColor(Grille.Objects[0,Grille.Row]).Drapeau:=''
   else PColor(Grille.Objects[0,Grille.Row]).Drapeau:='*' ;
Grille.Invalidate ; CompteElemSelectionner(Grille) ;
END ;

procedure TFFamrub.Fliste1MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if ((ssCtrl in Shift) and (Button=mbLeft)) then CocheDecoche(Sender) ;
end;

procedure TFFamrub.BValiderClick(Sender: TObject);
Var i : Integer ;
begin
if (Laquelle='') And (Not QueFListe2) then
   BEGIN
   if MsgBox.Execute(0,'','')=mrYes then
      BEGIN
      MajFamilleRubrique ; FaitRequeteMaj ; RempliGrid(FListe1) ; CompteElemSelectionner(FListe1) ;
      for i:=1 to FListe2.RowCount-1 do PColor(FListe2.Objects[0,i]).Drapeau:='' ;
      FListe2.Row:=1 ; FListe2.Invalidate ; CompteElemSelectionner(FListe2) ;
      END ;
   END else
   BEGIN
   LaFam:='' ;
   for i:=1 to Fliste2.RowCount-1 do
       if PColor(FListe2.Objects[0,i]).Drapeau='*' then
          BEGIN
          LaFam:=LaFam+Fliste2.Cells[0,i]+';' ;
          LibFam:=LibFam+Fliste2.Cells[1,i]+';' ;
          END ;
   Close ;
   END ;
end;

Procedure TFFamrub.MajFamilleRubrique ;
Var i,j : Integer ;
    St,Sql : String ;
BEGIN
j:=0 ;
for i:=1 to Fliste1.RowCount-1 do
    if PColor(Fliste1.Objects[0,i]).Drapeau='*' then Inc(j) ;
if j>0 then
   BEGIN
   for i:=1 to Fliste2.RowCount-1 do
       if PColor(Fliste2.Objects[0,i]).Drapeau='*' then St:=St+Fliste2.Cells[0,i]+';' ;
   BeginTrans ; InitMove(j,MsgBox.Mess[1]) ;
   for i:=1 to Fliste1.RowCount-1 do
      BEGIN
      if PColor(Fliste1.Objects[0,i]).Drapeau='*' then
         BEGIN
//Simon "
         Sql:='Update RUBRIQUE Set RB_FAMILLES="'+St+'" Where RB_RUBRIQUE="'+FListe1.Cells[0,i]+'"' ;
         Qrf.Close ; Qrf.Sql.Clear ; Qrf.Sql.Add(Sql) ; ChangeSql(Qrf) ; Qrf.ExecSql ;
         MoveCur(False) ;
         END ;
      END ;
   FiniMove ; CommitTrans ;
   END ;
END ;

procedure TFFamrub.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe1.VidePile(True) ; FListe2.VidePile(True) ;
OldBrush.Free ; OldPen.Free ;
end;

procedure TFFamrub.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
END ;

procedure TFFamrub.FListe2DrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
var Text      : array[0..255] of Char;
    F         : TAlignment ;
    Grille : THgrid ;
BEGIN
Grille:=THgrid(Sender) ;
OldBrush.assign(Grille.Canvas.Brush) ;
OldPen.assign(Grille.Canvas.Pen) ;
StrPCopy(Text,Grille.Cells[Col,Row]);
Grille.Canvas.Font.Style:=Grille.Canvas.Font.Style-[fsItalic] ;
if (gdFixed in State) then
  BEGIN
  Grille.Canvas.Brush.Color:=Grille.FixedColor ;
  Grille.Canvas.Font.Color:=Grille.Font.Color ;
  F:=taCenter
  END else
  BEGIN
  F:=taLeftJustify ;
  if (gdSelected in State) then
    BEGIN
    if PColor(Grille.Objects[0,Row]).Drapeau='*' then
       BEGIN
       Grille.Canvas.Brush.Color:=clInactiveCaption ;
       Grille.Canvas.Font.Color:=clHighlightText ;
       Grille.Canvas.Font.Style:=Grille.Canvas.Font.Style+[fsItalic] ;
       END else
       BEGIN
       Grille.Canvas.Brush.Color:=clHighlight ;
       Grille.Canvas.Font.Color:=clHighlightText ;
       END ;
    END else
    BEGIN
    if PColor(Grille.Objects[0,Row]).Drapeau='*' then
       BEGIN
       Grille.Canvas.Brush.Color:=clInactiveCaption ;
       Grille.Canvas.Font.Color:=cl3DLight ;
       END else
       BEGIN
       Grille.Canvas.Brush.Color:=clWindow ;
       Grille.Canvas.Font.Color:=clWindowText ;
       END ;
    END ;
  END ;
   Case F of
        taRightJustify : ExtTextOut(Grille.Canvas.Handle,Rect.Right-Grille.Canvas.TextWidth(Grille.Cells[Col,Row])-3,
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
        taCenter       : ExtTextOut(Grille.Canvas.Handle,Rect.Left+((Rect.Right-Rect.Left-Grille.canvas.TextWidth(Grille.Cells[Col,Row])) div 2),
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
        else             ExtTextOut(Grille.Canvas.Handle,Rect.Left+2,
                         Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
      End ;
if((gdfixed in State) and Grille.Ctl3D) then
   BEGIN
   DrawEdge(Grille.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
   DrawEdge(Grille.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
   END;
Grille.Canvas.Brush.Assign(OldBrush) ; Grille.Canvas.Pen.Assign(OldPen) ;
end;

procedure TFFamrub.FormCreate(Sender: TObject);
begin
OldBrush:=TBrush.Create ; OldPen:=TPen.Create ;
end;

Procedure TFFamrub.CompteElemSelectionner(Sender : TObject) ;
Var i,j : Integer ;
    C : Char ;
BEGIN
j:=0 ; C:=THGrid(Sender).Name[Length(THGrid(Sender).Name)] ;
for i:=1 to THGrid(Sender).RowCount-1 do
    if PColor(THGrid(Sender).Objects[0,i]).Drapeau='*' then Inc(j) ;
Case j of
     0,1: BEGIN
          if C='1' then
             BEGIN
             Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=MsgBox.Mess[3] ;
             END else
             BEGIN
             Nb2.Caption:=IntTostr(j) ; Tex2.Caption:=MsgBox.Mess[3] ;
             END ;
          END ;
     else BEGIN
          if C='1' then
             BEGIN
             Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=MsgBox.Mess[4] ;
             END else
             BEGIN
             Nb2.Caption:=IntTostr(j) ; Tex2.Caption:=MsgBox.Mess[4] ;
             END ;
          END ;
   End ;
END ;

procedure TFFamrub.Fliste1DblClick(Sender: TObject);
begin CocheDecoche(Sender) ; end;

Procedure TFFamrub.FaitRequeteMaj ;
BEGIN
Qrf.Close ; Qrf.Sql.Clear ;
if Budget then Qrf.Sql.Add('Select RB_RUBRIQUE,RB_LIBELLE,RB_FAMILLES From RUBRIQUE Where RB_NATRUB="BUD" Order by RB_RUBRIQUE')
          else Qrf.Sql.Add('Select RB_RUBRIQUE,RB_LIBELLE,RB_FAMILLES From RUBRIQUE Where RB_NATRUB<>"BUD" Order by RB_RUBRIQUE') ; 
ChangeSql(Qrf) ; Qrf.Open ;
END ;

procedure TFFamrub.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
