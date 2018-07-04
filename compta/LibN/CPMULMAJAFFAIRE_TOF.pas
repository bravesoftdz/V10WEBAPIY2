                                   {***********UNITE*************************************************
Auteur  ...... : YMO
Créé le ...... : 20/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULMAJAFFAIRE ()
Mots clefs ... : TOF;CPMULMAJAFFAIRE
*****************************************************************}
Unit CPMULMAJAFFAIRE_TOF ;

Interface

Uses Windows,
     forms,
     sysutils,
     ComCtrls,
     ExtCtrls,  // pout le TPanel
     Classes,
     StdCtrls,
     Controls,
     HTB97,     // pour TToolBarButton97
     Grids,     // Pour le TGridDrawState
     Graphics,  // pour clRed
     AGLInit,      // TheMulQ
{$IFDEF EAGLCLIENT}
    eMul,
    maineagl,
{$ELSE}
    FE_MAIN,
    Mul,
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    SaisBor,       // Pour lanceSaisieFolio
    HDB,
    dbGrids,
{$ENDIF}
     HQry,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
//     HStatus,       // pour InitMove et MoveCur et FiniMove
     UTOF,
     UTOB,
//     uLibWindows,   // pour CDessineTriangle, CVireLigne
     UtilPGI,       // pour Resolution
     ImgList,
     CPDATESAFFAIRES_TOF;

procedure CPLAnceFiche_AFFEnCours(vStRange, vStLequel, vStArgs: string);

Type
  TOF_CPMULMAJAFFAIRE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnAfterFormShow;
    procedure BOuvrirClick(Sender: TObject);
   {$IFDEF EAGLCLIENT}
    procedure FListePostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
//    procedure FListeOnDrawCell(Sender: TObject;ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);
    function  GetColName( vCol : Integer = -1 ) : String ;
    procedure BChercheClick( sender: TObject ) ;
   {$ELSE}
    procedure FListeDrawColumnCell(Sender: TObject;const Rect: TRect; DataCol: Integer; Column: TColumn;State: TGridDrawState);
   {$ENDIF EAGLCLIENT}
    procedure ModifCpteOnClick (Sender : TObject);
    function ModifieEnSerie:Boolean;

  private
    Q   : THQuery ;
    ImCadenas : TImage;
    ImCroix : TImage;
    ImFeuille : TImage;
//    DebChantier : THEdit;
    FinChantier : THEdit;
    S_AffaireEnCours : TCheckBox;
    PTabLibres  : TTabSheet;
{$IFDEF EAGLCLIENT}
    FListe   : THGrid ;
    FOnBChercheClick      : TNotifyEvent ;
{$ELSE}
    FListe   : THDBGrid ;
    FListeDrawColumnCellParent  : TDrawColumnCellEvent ;
{$ENDIF}

  end ;
Implementation

procedure CPLAnceFiche_AFFEnCours(vStRange, vStLequel, vStArgs: string);
begin
  AGLLanceFiche('CP', 'CPMULMAJAFFAIRE', vStRange, vStLequel, vStArgs);
end;


procedure TOF_CPMULMAJAFFAIRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMULMAJAFFAIRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMULMAJAFFAIRE.OnUpdate ;
begin
  Inherited ;  
end ;

procedure TOF_CPMULMAJAFFAIRE.OnLoad ;
begin
  LibellesTableLibre(PTabLibres, 'TS_TABLE', 'S_TABLE', 'S');
  Inherited ;
end ;

procedure TOF_CPMULMAJAFFAIRE.OnArgument (S : String ) ;
begin
  Inherited ;
  Q               := THQuery(GetControl('Q', True)) ;
  ImCadenas       := TImage(GetControl('ImCadenas', True));
  ImCroix         := TImage(GetControl('ImCroix', True));
  ImFeuille       := TImage(GetControl('ImFeuille', True));
//  DebChantier   := THEdit(GetControl('DEBCHANTIER', True));
  FinChantier   := THEdit(GetControl('FINCHANTIER', True));
  S_Affaireencours:= TCheckBox(GetControl('S_AFFAIREENCOURS', True));
  PTabLibres      := TTabSheet(GetControl('PTABLIBRES', True));
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).OnAfterFormShow := Self.OnAfterFormShow;
  FListe          := THGrid(GetControl('FListe', True));
  FOnBChercheClick := TToolBarButton97(GetControl('BCherche')).OnClick  ;
  TToolBarButton97(GetControl ('BCherche')).OnClick := BChercheClick ;
{$ELSE}
  FListe          := THDBGrid(GetControl('FListe', True));
  FListe.OnDrawColumnCell:=FListeDrawColumnCell;
{$ENDIF}
  TButton(GetControl('bOuvrir',True)).OnClick    := bOuvrirClick ;

//  DebChantier.Text:=StDate1900;
  FinChantier.Text:=StDate2099;
  S_AffaireEnCours.Checked:=True;
end ;

procedure TOF_CPMULMAJAFFAIRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPMULMAJAFFAIRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULMAJAFFAIRE.OnCancel () ;
begin
  Inherited ;
end ;

{$IFDEF EAGLCLIENT}
procedure TOF_CPMULMAJAFFAIRE.BChercheClick(sender: TObject);
begin                                      

  if Assigned ( FOnBChercheClick ) then
    FOnBChercheClick( Sender ) ;

   TFMul(Ecran).FListe.PostDrawCell := FListePostDrawCell ;
//  TFMul(Ecran).FListe.OnDrawCell := FListeOnDrawCell;

end;

function TOF_CPMULMAJAFFAIRE.GetColName( vCol : Integer = -1 ) : String;
begin
  if vCol < 0
    then result := FListe.ColNames[ FListe.Col ]
    else if vCol < FListe.ColCount
         then result := FListe.ColNames[ vcol ] ;
end;

procedure TOF_CPMULMAJAFFAIRE.FListePostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
//procedure TOF_CPMULMAJAFFAIRE.FListeOnDrawCell(Sender: TObject;ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);
var Value : String ;
    lstcol  : String ;
    Rc    : TRect ;
    Icone : TBitmap ;
begin

  if ((Q.TQ.EOF) and (Q.TQ.BOF)) then Exit ;

  lStCol   := GetColName( ACol ) ;
  if (lstcol='S_AFFAIREENCOURS') then
  begin

    Value := FListe.Cells[ACol, ARow] ;
    if Value <> '' then
    begin

      If (Q.FindField('S_DEBCHANTIER').AsDateTime=idate1900) and (Q.FindField('S_FINCHANTIER').AsDateTime=idate2099) then
          Icone:=ImCroix.Picture.Bitmap
      else
      if Q.FindField('S_FINCHANTIER').AsDateTime<=StrToDate(FinChantier.Text) then
          Icone:=ImCadenas.Picture.Bitmap
      else
          Icone:=ImFeuille.Picture.Bitmap;

      Rc := FListe.CellRect(ACol, ARow) ;
      FListe.Canvas.FillRect( Rc );
      FListe.Canvas.Draw((Rc.Left+Rc.Right) div 2 -7, (Rc.Top+Rc.Bottom) div 2 -5 , Icone);

    end ;

  end;

end;
{$ELSE}
procedure TOF_CPMULMAJAFFAIRE.FListeDrawColumnCell(Sender: TObject;const Rect: TRect; DataCol: Integer; Column: TColumn;State: TGridDrawState);
var Value : String ;
    Rc    : TRect ;
    Icone : TBitmap ;
begin
  if Assigned(FListeDrawColumnCellParent) then
    FListeDrawColumnCellParent(Sender,Rect,DataCol,Column,State);

  if ((Q.EOF) and (Q.BOF)) then Exit ;

  if (Column.FieldName='S_AFFAIREENCOURS') then
  begin

    Value := FListe.Columns[DataCol].Field.AsString ;
    if Value <> '' then
    begin

      If (Q.FindField('S_DEBCHANTIER').AsDateTime=idate1900) and (Q.FindField('S_FINCHANTIER').AsDateTime=idate2099) then
          Icone:=ImCroix.Picture.Bitmap
      else
      if Q.FindField('S_FINCHANTIER').AsDateTime<=StrToDate(FinChantier.Text) then
          Icone:=ImCadenas.Picture.Bitmap
      else
          Icone:=ImFeuille.Picture.Bitmap;


      Rc := Rect ;
      FListe.Canvas.FillRect( Rc );
      Fliste.Canvas.Draw((Rc.Left+Rc.Right) div 2 -7, (Rc.Top+Rc.Bottom) div 2 -5 , Icone);
    end ;

  end;

end;
{$ENDIF EAGLCLIENT}


procedure TOF_CPMULMAJAFFAIRE.BOuvrirClick(Sender: TObject);
begin
  ModifieEnSerie;
  TFMul(Ecran).BChercheClick(nil) ; // Maj automatique, rafraîchissement grille
end;

procedure TOF_CPMULMAJAFFAIRE.ModifCpteOnClick(Sender: TObject);
{
var lStAxe    : String ;
    lStCpt    : String ;
    lStTable  : String ;
    lStAction : String ;
    lInForm   : Longint ;
}
begin
{
  lInForm   := Longint( Ecran ) ;
  lStAxe    := GetField('S_AXE') ;
  lStCpt    := GetField('S_SECTION') ;
  lStTable  := 'SECTION' ;
  lStAction := 'ACTION=MODIFICATION' ;
  if lStAxe <> ''
    then AglModifCompte( [ lInForm, lStCpt, 'Modif', lStTable, lStAxe, lStAction ] , 0 )
    else AglModifCompte( [ lInForm, lStCpt, 'Modif', lStTable, lStAction ] , 0 ) ;
//  BChercheClick(nil) ;
}
end;

function TOF_CPMULMAJAFFAIRE.ModifieEnSerie:boolean;
var
  {$IFNDEF EAGLCLIENT}
    BM                       : TBookmark;
  {$ENDIF}
    Pref                     : String;
    Ladate                   : String;
    StWhere                  : String;
    SQL                      : String;
    StSQL                    : String;
    FCtrl                    : String;
    i                        : Integer;
    Q                        : THQuery;
BEGIN
Result := True;
FCtrl:='' ;
Pref:='S_' ;
Ladate:=','+Pref+'DATEMODIF="'+UsDateTime(Date)+'"';
StSQL:=ModifZoneSerie;
if StSQL='' then
begin
  Result := False;
  exit ;
end;
StSQL:=StSql+LaDate ;
SQL:='UPDATE SECTION SET ' + StSQL;
Q:=THQuery(Ecran.FindComponent('Q')) ;
if (Q=Nil) then exit;
{$IFNDEF EAGLCLIENT}
    // CA - 23/06/2003 - Pour navigation sur les enregistrements de la query
    TheMulQ:=TFMul(Ecran).Q;
{$ENDIF}
{$IFNDEF EAGLCLIENT}
BM:=Q.GetBookmark;
{$ENDIF}
if FListe.AllSelected then
  BEGIN
  Q.DisableControls ;
  Q.First ;

{$IFDEF EAGLCLIENT}
  While not Q.TQ.Eof do
    BEGIN
    StWhere:=' WHERE S_SECTION="'+ Q.FindField('S_SECTION').AsString+'"';
    ExecuteSQL(SQL+StWhere);
    Q.TQ.Next;
    END ;
{$ELSE}
  While not Q.Eof do
    BEGIN
    StWhere:=' WHERE S_SECTION="'+ Q.FindField('S_SECTION').AsString+'"';
    BeginTrans;
    ExecuteSQL(SQL+StWhere);
    CommitTrans;
    Q.Next ;
    END ;
{$ENDIF}
  Q.EnableControls ;
END
ELSE
BEGIN
  For i:=0 to FListe.NbSelected-1 do
  BEGIN
   FListe.GotoLeBookMark(i) ;
{$IFDEF EAGLCLIENT}
   Q.TQ.Seek(FListe.Row - 1);
   StWhere:=' WHERE S_SECTION="'+ Q.FindField('S_SECTION').AsString+'"';
   ExecuteSQL(SQL+StWhere);
{$ELSE}
   StWhere:=' WHERE S_SECTION="'+ Q.FindField('S_SECTION').AsString+'"';
   ExecuteSQL(SQL+StWhere);
{$ENDIF}
  END ;
END ;
{$IFNDEF EAGLCLIENT}
Q.GotoBookmark(BM); Q.FreeBookmark(BM);
{$ENDIF}
END;

procedure TOF_CPMULMAJAFFAIRE.OnAfterFormShow;
begin
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).FListe.PostDrawCell := FListePostDrawCell ;
//  TFMul(Ecran).FListe.OnDrawCell := FListeOnDrawCell;
{$ENDIF}
end;


Initialization
  registerclasses ( [ TOF_CPMULMAJAFFAIRE ] ) ;
end.

