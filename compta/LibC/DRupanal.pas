unit DRupanal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Hctrls, hmsgbox, Hent1, Grids, HStatus,
  Hqry, Ent1,
{$IFDEF EAGLCLIENT}
  UtileAGL, // PrintDBGrid
  uTob,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  PrintDBG, // PrintDBGrid
{$ENDIF}
  HSysMenu ;

Function DetailPlanRuptureAnal(Nature,Plan,Axe,CodSp : String) : Boolean ;

Type TCodSPlan = Class
     UnTab : Array[1..17] of HTStrings ;
     Constructor Create ;
     Destructor Destroy ; override ;
     End ;

type
  TFDrupanal = class(TForm)
    Pbouton: TPanel;
    FAutoSave: TCheckBox;
    MsgBox: THMsgBox;
    PTop: TPanel;
    TPlan: THLabel;
    Eplan: TEdit;
    TLibplan: THLabel;
    ElibPlan: TEdit;
    Taxe: THLabel;
    Eaxe: TEdit;
    Fliste: THGrid;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    Bdetag: THBitBtn;
    BTag: THBitBtn;
    HMTrad: THSystemMenu;
    Nb1: TLabel;
    Tex1: TLabel;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FlisteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    Nature,Plan,Axe,CodSp : String ;
    StOrdre : String ;
    TabOrd : Array[1..17]of String3 ;
    Cod : Array[0..17]of String ;
    Lib : Array[0..17]of String ;
    TabRequete : array [1..17] of string;
    Generer : Boolean ;
    WMinX,WMinY : Integer ;
    Tablo : TCodSPlan ;
    LgLib : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure Libelle ;
    Procedure QuelOrdre ;
    Procedure CreerRequete( vIndice : integer ; St : String) ;
    Procedure AfficheCode(St,St1 : String ) ;
    Procedure GenereCode(i : Byte) ;
    Procedure RempliTabOrd ;
    Procedure InitQuery ;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure CompteElemSelectionner ;
    Function  ListeVide : Boolean ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure InverseSelection ;
    Procedure RempliTablo ;
    Procedure GenereCodePartiel(j : Integer) ;
  public
    { Déclarations publiques }
  end;


implementation

Uses Rupanal, UtilPgi ;

{$R *.DFM}

Constructor TCodSPlan.Create ;
Var i  :Integer ;
BEGIN
Inherited Create ;
for i:=1 to 17 do UnTab[i]:=Nil ;
END ;

Destructor TCodSPlan.Destroy ;
Var i  :Integer ;
BEGIN
for i:=1 to 17 do if UnTab[i]<>Nil then BEGIN UnTab[i].Free ; UnTab[i]:=Nil ; END ;
Inherited Destroy ;
END ;

Function DetailPlanRuptureAnal(Nature,Plan,Axe,CodSp : String) : Boolean ;
var FDrupanal : TFDrupanal ;
BEGIN
 Result:=False ;
 if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
 FDrupanal:=TFDrupanal.Create(Application) ;
 Try
  FDrupanal.Nature:=Nature ;
  FDrupanal.Plan:=Plan ;
  FDrupanal.Axe:=Axe ;
  FDrupanal.CodSp:=CodSp ;
  FDrupanal.Generer:=False ;
  FDrupanal.ShowModal ;
  if FDrupanal.Generer then Result:=True ;
 Finally
  FDrupanal.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFDrupanal.BFermeClick(Sender: TObject);
begin
  Close ;
end;

////////////////////////////////////////////////////////////////////////////////
Procedure TFDrupanal.Libelle ;
var Q : TQuery ;
begin
  Q := OpenSql('Select CC_LIBELLE From CHOIXCOD Where CC_TYPE="'+Nature+'" And CC_CODE="'+Plan+'"',True) ;
  ElibPlan.Text := Q.FindField('CC_LIBELLE').AsString ;
  Ferme(Q) ;

  Q := OpenSql('Select X_LIBELLE From AXE Where X_AXE="'+Axe+'"',True) ;
  Eaxe.Text:= Q.FindField('X_LIBELLE').AsString ;
  Ferme(Q) ;

  Eplan.Text:=Plan ;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFDrupanal.QuelOrdre ;
var Q : TQuery ;
begin
  Q := OpenSql('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE = "' + Nature +
               '" AND CC_CODE = "' + Plan + '"', True) ;
  StOrdre := Q.FindField('CC_LIBRE').AsString ;
  Ferme(Q) ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 15/04/2004
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFDrupanal.FormShow(Sender: TObject);
var i : Byte ;
begin
  FListe.GetCellCanvas := GetCellCanvas ;
  PopUpMenu := ADDMenuPop(PopUpMenu,'','') ;
  i := 1 ;
  FListe.ColWidths[FListe.ColCount-1] := 0 ;
  Libelle ;
  if CodSp = '' then
  begin
    QuelOrdre ;
    RempliTabOrd ;
    if TabOrd[i] <> '' then
    begin
      InitQuery ;
      GenereCode(i);
    end;
  end
  else
  begin
    RempliTablo ;
    GenereCodePartiel(i) ;
  end;

  if FListe.RowCount > 2 then
    FListe.RowCount := FListe.RowCount - 1 ;

  CompteElemSelectionner ;
  FListe.Invalidate ;
  FListe.SetFocus ;
end;

procedure TFDrupanal.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
  if FListe.Cells[FListe.ColCount-1,ARow] = '*' then
    FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
  else
    FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
end ;

procedure TFDrupanal.FlisteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (ssCtrl in Shift) and (Button=mbLeft) then
    InverseSelection ;
end;

procedure TFDrupanal.BValiderClick(Sender: TObject);
Var i,j : Integer ;
    Trouver : Boolean ;
begin
Trouver:=False ; j:=0 ;
for i:=1 to FListe.RowCount-1 do
   if FListe.Cells[FListe.ColCount-1,i]='*' then BEGIN Trouver:=True ; Inc(j) ; END ;
if Trouver then
   BEGIN
   if CodSp='' then
      BEGIN
      if MsgBox.Execute(0,'','')<>mrYes then Exit ;
         ExecuteSql('Delete From RUPTURE Where RU_NATURERUPT="'+Nature+'" And '+
                    'RU_PLANRUPT="'+Plan+'" And RU_SOCIETE="'+V_PGI.CodeSociete+'"') ;
      END ;
   BeginTrans ; InitMove(j,MsgBox.Mess[1]) ;
   for i:=1 to FListe.RowCount-1 do
       BEGIN
       if FListe.Cells[FListe.ColCount-1,i]='*' then
          BEGIN
          if Not Presencecomplexe('RUPTURE',['RU_NATURERUPT','RU_PLANRUPT','RU_CLASSE'],['=','=','='],[Nature,Plan,Copy(FListe.Cells[0,i],1,17)],['S','S','S']) then
          ExecuteSql('Insert Into RUPTURE (RU_NATURERUPT,RU_PLANRUPT,RU_CLASSE,'+
                     'RU_LIBELLECLASSE,RU_SOCIETE) '+
                     'Values("'+Nature+'","'+Plan+'","'+Copy(FListe.Cells[0,i],1,17)+'",'+
                             '"'+Copy(FListe.Cells[1,i],1,35)+'","'+V_PGI.CodeSociete+'")') ;
          MoveCur(False) ;
          END ;
       END ;
   FiniMove ; CommitTrans ; Generer:=True ;
   END ;
Close ;
end;

Procedure TFDrupanal.InverseSelection ;
BEGIN
if ListeVide then Exit ;
if FListe.Cells[FListe.ColCount-1,FListe.Row]='*' then FListe.Cells[FListe.ColCount-1,FListe.Row]:=''
                                                  else FListe.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
CompteElemSelectionner ; FListe.Invalidate ;
END ;

Function TFDrupanal.ListeVide : Boolean ;
BEGIN Result:=FListe.Cells[0,1]='' ; END ;

Procedure TFDrupanal.CompteElemSelectionner ;
Var i,j : Integer ;
BEGIN
j:=0 ;
if Not ListeVide then
   BEGIN
   for i:=1 to FListe.RowCount-1 do
       if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(j) ;
   END ;
Case j of
     0,1: BEGIN
          Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=MsgBox.Mess[2] ;
          END ;
     else BEGIN
          Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=MsgBox.Mess[3] ;
          END ;
   End ;
END ;

Procedure TFDrupanal.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Visible:=False ; FListe.Visible:=True ; FListe.SetFocus ;
Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ; CompteElemSelectionner ;
end;

procedure TFDrupanal.BdetagClick(Sender: TObject);
begin
  TagDetag(False) ;
end;

procedure TFDrupanal.BTagClick(Sender: TObject);
begin
  TagDetag(True) ;
end;

procedure TFDrupanal.FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((ssShift in Shift) and (Key=VK_DOWN)) or
      ((ssShift in Shift) and (Key=VK_UP)) then
    InverseSelection
  else
  begin
    if (Shift=[]) and (Key=VK_SPACE) then
    begin
      InverseSelection ;
      if (FListe.Row<FListe.RowCount-1) then
        FListe.Row:=FListe.Row+1 ;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 15/04/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFDrupanal.InitQuery ;
var i : Byte ;
begin
  for i := 1 to 17 do
  begin
    if TabOrd[i] <> '' then
    begin
      CreerRequete( i, TabOrd[i]) ;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 15/04/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFDrupanal.GenereCode(i : Byte) ;
var j : Byte ;
    St,St1 : String ;
    lQuery : TQuery;
begin
  try
    try
      lQuery := OpenSQL(TabRequete[i], True);
      while not lQuery.EOF do
      begin
        Cod[i] := lQuery.Fields[0].AsString ;
        Lib[i] := Copy(lQuery.Fields[1].AsString, 1, 17) ;
        St     := '' ;
        St1    := '' ;
        for j := 1 to 17 do
        begin
          St  := St + Cod[j] ;
          St1 := St1 + Lib[j] ;
        end;
        AfficheCode(St,St1) ;
        if (TabOrd[i+1]<>'') and (i+1<=17) then GenereCode(i+1) ;
        lQuery.Next;
      end;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure GenereCode ');
    end;
  finally
    Ferme( lQuery );
  end;
  Cod[i] := '' ;
  Lib[i] := '' ;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFDrupanal.RempliTabOrd ;
var St : String ;
    i : Byte ;
begin
  FillChar(Cod,SizeOf(Cod),#0) ;
  FillChar(Lib,SizeOf(Lib),#0) ;
  FillChar(TabOrd,SizeOf(TabOrd),#0) ;
  FillChar(TabRequete, SizeOf(TabRequete), #0) ;
  i:=1 ;
  while StOrdre <> '' do
  begin
    St:=ReadTokenSt(StOrdre) ;
    TabOrd[i]:=St ;
    Inc(i) ;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 15/04/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFDrupanal.CreerRequete( vIndice : integer; St : String) ;
begin
  TabRequete[vIndice] := 'SELECT PS_CODE,PS_LIBELLE FROM SSSTRUCR WHERE ' +
                         'PS_AXE = "' + Axe + '" AND ' +
                         'PS_SOUSSECTION = "' + St +'" ORDER BY PS_SOUSSECTION';
                         //RR 05/01/2006 FQ17275
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFDrupanal.AfficheCode(St,St1 : String ) ;
begin
  Fliste.Cells[0,FListe.RowCount-1] := St + 'x' ;
  Fliste.Cells[1,FListe.RowCount-1] := St1 ;
  FListe.Cells[FListe.ColCount-1,FListe.RowCount-1] := '*' ;
  FListe.RowCount := FListe.RowCount + 1 ;
end; 

procedure TFDrupanal.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  PrintDBGrid(Caption, FListe.ListeParam, '', '');
{$ELSE}
  PrintDBGrid(FListe,PTop,Caption,'') ;
{$ENDIF}
end;

procedure TFDrupanal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  AvertirTable('ttRuptSect1') ;
  AvertirTable('ttRuptSect2') ;
  AvertirTable('ttRuptSect3') ;
  AvertirTable('ttRuptSect4') ;
  AvertirTable('ttRuptSect5') ;
  FListe.VidePile(False) ;
  Tablo.Destroy ;
end;

procedure TFDrupanal.WMGetMinMaxInfo(var MSG: Tmessage);
begin with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; end;

procedure TFDrupanal.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; Tablo:=TCodSPlan.Create ; end;

procedure TFDrupanal.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

Procedure TFDrupanal.RempliTablo ;
Var i,j : Integer ;
    QLoc,QLib,QSSec : TQuery ;
    St,St1,UnCod : String ;
    Li : HTStrings ;
begin
  i:=1 ;
  QLoc := OpenSql('Select Count(SS_SOUSSECTION) From STRUCRSE Where SS_AXE="'+Axe+'"',True) ;
  LgLib := QLoc.Fields[0].AsInteger ;
  Ferme(QLoc) ;

  Li:=HTSTringList.Create ;
  St:=CodSp ;
  while St<>'' do
  begin
    St1:=ReadTokenSt(St) ;
    if St1<>'' then
    begin
      QLoc := OpenSql('SELECT DISTINCT(PS_SOUSSECTION) FROM SSSTRUCR WHERE ' +
                      'PS_AXE = "' + Axe + '" AND PS_CODE ="' + St1 + '"',TRUE);
      Li.Add(QLoc.Fields[0].AsString+';'+St1) ;
      Ferme(QLoc) ;
    end;
  end;

  QLoc := OpenSql('SELECT SS_SOUSSECTION FROM STRUCRSE WHERE SS_AXE = "'+Axe+'"',True) ;
  while not QLoc.Eof do
  begin
    for j:=0 to Li.Count-1 do
    begin
      if Copy(Li.Strings[j], 1, Pos(';',Li.Strings[j])-1) = QLoc.FindField('SS_SOUSSECTION').AsString then
      begin
        if Tablo.UnTab[i]= nil then
          Tablo.UnTab[i]:=HTStringList.Create ;
        UnCod := Copy(Li.Strings[j],Pos(';',Li.Strings[j])+1,Length(Li.Strings[j])) ;

        QLib := OpenSQL('SELECT PS_LIBELLE FROM SSSTRUCR WHERE ' +
                        'PS_AXE = "' + Axe + '" AND PS_CODE="'+UnCod+'"', True);
        Tablo.UnTab[i].Add(UnCod + ';' + QLib.FindField('PS_LIBELLE').AsString);
        Ferme(QLib);
      end
      else
      begin
        QSSec := OpenSQL('SELECT PS_CODE, PS_LIBELLE FROM SSSTRUCR WHERE ' +
                 'PS_AXE = "' + Axe + '" AND PS_SOUSSECTION = "' +
                 QSSec.FindField('SS_SOUSSECTION').AsString + '"', True);
        while not QSSec.Eof do
        begin
          if Tablo.UnTab[i] = nil then
            Tablo.UnTab[i]:=HTStringList.Create ;
          Tablo.UnTab[i].Add(QSSec.Fields[0].AsString+';'+QSSec.Fields[1].AsString) ;
          QSSec.Next ;
        end ;
        Ferme( QSSec );
      end ;
    end ;
    QLoc.Next ;
    Inc(i) ;
  end; ;
  Ferme(QLoc) ;
end ;

Procedure TFDrupanal.GenereCodePartiel(j : Integer) ;
Var i,k : Integer ;
    St,StC,StL : String ;
BEGIN
if (j<=17) and (Tablo.UnTab[j]<>Nil) then
   BEGIN
   for i:=0 to Tablo.UnTab[j].Count-1 do
      BEGIN
      St:=Tablo.UnTab[j][i] ;
      Cod[j]:=ReadTokenSt(St) ; Lib[j]:=Copy(ReadTokenSt(St),1,35 div LgLib) ;
      StC:='' ; StL:='' ;
      for k:=1 to 17 do BEGIN StC:=StC+Cod[k] ; StL:=StL+Lib[k] ; END ;
      GenereCodePartiel(j+1) ;
      if Cod[LgLib]<>'' then AfficheCode(StC,StL) ;
      END ;
      Cod[j]:='' ; Lib[j]:='' ;
   END ;
END ;

end.
