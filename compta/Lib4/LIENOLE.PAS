unit LienOLE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, OleCtnrs, ExtCtrls, StdCtrls, Db, {$IFNDEF DBXPRESS}dbtables,
  Hqry, Buttons{$ELSE}uDbxDataSet{$ENDIF}, Buttons, Hqry ;

Procedure GereOLE(LeType,LeIdent : String ) ;

type
  TFLienOLE = class(TForm)
    PBoutons: TPanel;
    TLien: THTable;
    FInPlace: TCheckBox;
    FCode: TComboBox;
    BPrev: TBitBtn;
    BNext: TBitBtn;
    BAnnuler: TBitBtn;
    BInsert: TBitBtn;
    BDelete: TBitBtn;
    BValider: TBitBtn;
    BFerme: TBitBtn;
    BAide: TBitBtn;
    BModif: TBitBtn;
    BClose: TBitBtn;
    BLink: TBitBtn;
    MainMenu: TMainMenu;
    Edit1: TMenuItem;
    Cut1: TMenuItem;
    OleContainer1: TOleContainer;
    procedure FInPlaceClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BModifClick(Sender: TObject);
    procedure BCloseClick(Sender: TObject);
    procedure BLinkClick(Sender: TObject);
    procedure OleContainer1Activate(Sender: TObject);
    procedure OleContainer1Deactivate(Sender: TObject);
    procedure OleContainer1ObjectMove(OleContainer: TOleContainer;
      const Bounds: TRect);
    procedure BDeleteClick(Sender: TObject);
  private     { Déclarations privées }
    LeType,LeIdent : String ;
  public      { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure TFLienOLE.FInPlaceClick(Sender: TObject);
begin
OleContainer1.AllowInPlace:=FInPlace.Checked ;
end;

procedure TFLienOLE.FormShow(Sender: TObject);
begin
FCode.Items.Clear ;
TLien.Open ;
TLien.SetRange([LeType,LeIdent],[LeType,LeIdent]) ;
While Not TLien.EOF do
   begin
   FCode.Items.Add(TLien.FindField('LO_CODE').AsString) ;
   TLien.Next ;
   end ;
TLien.Close ;
end;

procedure TFLienOLE.BInsertClick(Sender: TObject);
begin
  if (OleContainer1.State = osEmpty) or
    (MessageDlg('Delete current OLE object?', mtConfirmation, mbOkCancel, 0) = mrOk) then
    if OleContainer1.InsertObjectDialog then
    begin
    Menu:=MainMenu ;
    OleContainer1.DoVerb(ovShow);
    end;
end;

procedure TFLienOLE.BValiderClick(Sender: TObject);
Var s : TMemoryStream;
    k : String ;
begin
TLien.Open ;
if TLien.FindKey([LeType,LeIdent,FCode.text]) then TLien.Edit else TLien.Insert ;
s := TMemoryStream.Create ;
OleContainer1.SaveToStream(s);
TLien.FindField('LO_TYPE').AsString :=LeType ;
TLien.FindField('LO_IDENT').AsString :=LeIdent ;
TLien.FindField('LO_CODE').AsString :=FCode.text ;
TLien.FindField('LO_LIBELLE').AsString :='AAAAAAAA' ;
TBlobField(TLien.FindField('LO_OBJET')).LoadFromStream(s) ;
TLien.Post ;
s.free ;
TLien.Close ;
end;

procedure TFLienOLE.BAnnulerClick(Sender: TObject);
Var s : TMemoryStream ;
    i : integer ;
    k : String ;
begin
TLien.Open ;
if TLien.FindKey([LeType,LeIdent,FCode.text]) then
   begin
   s := TMemoryStream.Create;
   TBlobField(TLien.FindField('LO_OBJET')).SaveToStream(s) ;
   s.position:=0 ;
   OleContainer1.LoadFromStream(s);
   s.Free ;
   end ;
TLien.Close ;
end;

procedure TFLienOLE.BModifClick(Sender: TObject);
begin
Menu:=MainMenu ;
OleContainer1.DoVerb(ovShow);
end;

procedure TFLienOLE.BCloseClick(Sender: TObject);
begin
While OleContainer1.State<>osLoaded do
   OleContainer1.Close ;
end;

procedure TFLienOLE.BLinkClick(Sender: TObject);
begin
OleContainer1.ObjectPropertiesDialog;
end;

Procedure GereOLE(LeType,LeIdent : String ) ;
var FLienOLE: TFLienOLE;

BEGIN
FLienOLE:=TFLienOLE.Create(Application) ;
Try
  FLienOLE.LeType:=LeType ;
  FLienOLE.LeIdent:=LeIdent ;
  FLienOLE.ShowModal ;
  Finally
  FLienOLE.Free ;
  End ;
END ;

procedure TFLienOLE.OleContainer1Activate(Sender: TObject);
begin
Menu:=MainMenu ;
end;

procedure TFLienOLE.OleContainer1Deactivate(Sender: TObject);
begin
Menu:=Nil;
end;

procedure TFLienOLE.OleContainer1ObjectMove(OleContainer: TOleContainer;
  const Bounds: TRect);
begin
   OleContainer1.BoundsRect := Bounds;
end;

procedure TFLienOLE.BDeleteClick(Sender: TObject);
begin
if (OleContainer1.State <> osEmpty) and
    (MessageDlg('Delete current OLE object?', mtConfirmation, mbOkCancel, 0) = mrOk) then
    begin
    OleContainer1.DestroyObject ;
    TLien.Open ;
    if TLien.FindKey([LeType,LeIdent,FCode.text]) then TLien.Delete ;
    TLien.Close ;
    end;

end;

end.
