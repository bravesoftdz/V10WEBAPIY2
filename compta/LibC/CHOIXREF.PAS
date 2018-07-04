unit ChoixRef;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  StdCtrls, Buttons, ExtCtrls, Grids, Hctrls, DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, hmsgbox,
  HSysMenu, HEnt1;

procedure ChoixDesReferences ;

type
  TFChoixRef = class(TForm)
    Liste: THGrid;
    Panel5: TPanel;
    BValider: THBitBtn;
    BHelp: THBitBtn;
    BAbandon: THBitBtn;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    procedure BValiderClick(Sender: TObject);
    procedure ListeKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure ListeDblClick(Sender: TObject);
  private
    procedure ChargeCptesPointables ;
    function  OkRefSaisis : boolean ;
    procedure MajEEXBQ(Tous : boolean) ;
  public
  end;

implementation

uses Choix ;

{$R *.DFM}

procedure ChoixDesReferences ;
var FChoixRef : TFChoixRef ;
BEGIN
FChoixRef:=TFChoixRef.Create(Application) ;
 Try
  FChoixRef.ShowModal ;
 Finally
  FChoixRef.Free ;
 End ;
SourisNormale ;
END ;

procedure TFChoixRef.BValiderClick(Sender: TObject);
var i : Integer ;
begin
if not OkRefSaisis then BEGIN HM.Execute(1,Caption,'') ; ModalResult:=mrNone ; Exit ; END ;
MajEEXBQ(True) ;
for i:=1 to Liste.RowCount-1 do
  if (Liste.Cells[0,i]<>'') and (Liste.Cells[1,i]<>'') then
    ExecuteSQL('UPDATE IMPECR SET IE_REFPOINTAGE="'+Liste.Cells[1,i]+'"'+
               ', IE_DATEPOINTAGE="'+USDateTime(Date)+'" WHERE IE_GENERAL="'+Liste.Cells[0,i]+'"') ;
ModalResult:=mrOk ;
end;

procedure TFChoixRef.ListeKeyPress(Sender: TObject; var Key: Char);
begin
if (Key in['a'..'z']) then Dec(Key,32) ;
end;

procedure TFChoixRef.ChargeCptesPointables ;
var Q1  : TQUERY ;
BEGIN
Liste.RowCount:=2 ;
Q1:=OpenSQL('SELECT DISTINCT IE_GENERAL,IE_REFPOINTAGE,IE_DATEPOINTAGE FROM IMPECR LEFT JOIN GENERAUX ON '
          +'IE_GENERAL=G_GENERAL WHERE G_POINTABLE="X"',True) ;
While not Q1.EOF do
  BEGIN
  Liste.Cells[0,Liste.RowCount-1]:=Q1.Fields[0].AsString ;
  if (Trim(Q1.Fields[1].AsString)<>'') then Liste.Cells[1,Liste.RowCount-1]:=Q1.Fields[1].AsString ;
  Liste.RowCount:=Liste.RowCount+1 ;
  Q1.Next ;
  END ;
if Liste.RowCount>2 then Liste.RowCount:=Liste.RowCount-1 ;
Ferme(Q1) ;
END ;

procedure TFChoixRef.FormShow(Sender: TObject);
begin
ChargeCptesPointables ;
end;

procedure TFChoixRef.MajEEXBQ(Tous : boolean) ;
var Gene,St : String ;
    QEEXBQ : TQuery ;
    i,F,L  : integer ;
begin
F:=1 ; L:=Liste.RowCount ;
if not Tous then BEGIN F:=Liste.Row ; L:=F ; END ;
for i:=F to L do
  BEGIN
  Gene:=Liste.Cells[0,i] ;
  St:=Liste.Cells[1,i] ;
  if (St<>'') and (Gene<>'') then
    BEGIN
    QEEXBQ:=OpenSQL('SELECT * FROM EEXBQ WHERE EE_GENERAL="'+Gene+'" AND EE_REFPOINTAGE="'+St+'"',False) ;
    if QEEXBQ.Eof then BEGIN QEEXBQ.Insert ; InitNew(QEEXBQ) ; END else QEEXBQ.Edit ;
    QEEXBQ.FindField('EE_GENERAL').AsString:=Gene ;
    QEEXBQ.FindField('EE_DATEPOINTAGE').AsDateTime:=Date ;
    QEEXBQ.FindField('EE_REFPOINTAGE').AsString:=St ;
    QEEXBQ.Post ;
    Ferme(QEEXBQ) ;
    END ;
  END ;
END ;
procedure TFChoixRef.ListeDblClick(Sender: TObject);
var St : string ;
begin
MajEEXBQ(False) ;
St:=Choisir(HM.Mess[0],'EEXBQ','EE_REFPOINTAGE','EE_REFPOINTAGE','EE_GENERAL="'+Liste.Cells[0,Liste.Row]+'"','EE_DATEPOINTAGE DESC') ;
if (St<>'') then
  BEGIN
  Liste.Cells[1,Liste.Row]:=St ;
  ExecuteSQL('UPDATE EEXBQ SET EE_DATEPOINTAGE="'+UsDateTime(Date)+'" WHERE EE_GENERAL="'+Liste.Cells[0,Liste.Row]+'" AND EE_REFPOINTAGE="'+St+'"') ;
  END ;
end;

function TFChoixRef.OkRefSaisis : boolean;
var i : integer ;
BEGIN
Result:=False ;
for i:=1 to Liste.RowCount-1 do
  if (Liste.Cells[0,i]<>'') and (Liste.Cells[1,i]='') then Exit ;
Result:=True ;
END ;

end.
