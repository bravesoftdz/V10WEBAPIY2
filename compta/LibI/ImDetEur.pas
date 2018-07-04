unit ImDetEur;

interface

uses
   Classes, Graphics, Controls, Forms,
   Buttons, Grids, Hctrls, ExtCtrls,
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   Hent1, HSysMenu, hmsgbox, Math, ImEnt, StdCtrls;

type
  TFDetailEuro = class(TForm)
    PFen: TPanel;
    G: THGrid;
    POutils: TPanel;
    BAbandon: TBitBtn;
    BAide: TBitBtn;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    fQuery : TQuery;
    procedure InitGrille;
  public
    { Déclarations publiques }
  end;

procedure AfficheDetailEuro (Q : TQuery);

implementation

{$R *.DFM}

procedure AfficheDetailEuro (Q : TQuery);
var  FDetailEuro: TFDetailEuro;
begin
  FDetailEuro := TFDetailEuro.Create(application);
  FDetailEuro.fQuery := Q;
  try
    FDetailEuro.ShowModal;
  finally
     FDetailEuro.Free;
  end;
end;

procedure TFDetailEuro.InitGrille;
var i : integer;
    ColSize : integer;
    NomTablette: string ;
begin
  G.ColAligns[0]:=taCenter ;
  {$IFDEF SERIE1}
  NomTablette:='TTDEVISE' ;
  {$ELSE}
  NomTablette:='TTDEVISETOUTES' ;
  {$ENDIF}
  for i:=1 to G.ColCount - 1 do    G.ColAligns[i] := taRightJustify;
  G.Cells[0,1] := RechDom(NomTablette,V_PGI.DevisePivot,False);
  if Not VHImmo^.TenueEuro then G.Cells[0,2] := HM.Mess[0] else
    G.Cells[0,2] := RechDom(NomTablette,V_PGI.DeviseFongible,False) ;
  ColSize := MaxIntValue ([G.Canvas.TextExtent(G.Cells[0,1]).cx,G.Canvas.TextExtent(G.Cells[0,2]).cx]);
  G.ColWidths[0]:= ColSize+10;
end;

procedure TFDetailEuro.FormShow(Sender: TObject);
begin
  InitGrille;
  G.Cells[1,1]:=StrFMontant(fQuery.FindField('I_MONTANTHT').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[1,2]:=StrFMontant(fQuery.FindField('I_MONTANTHTCONTRE').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[2,1]:=StrFMontant(fQuery.FindField('I_TVARECUPERABLE').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[2,2]:=StrFMontant(fQuery.FindField('I_TVACONTRE').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[3,1]:=StrFMontant(fQuery.FindField('I_TVARECUPEREE').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[3,2]:=StrFMontant(fQuery.FindField('I_TVARECUPCONTRE').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[4,1]:=StrFMontant(fQuery.FindField('I_BASETAXEPRO').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[4,2]:=StrFMontant(fQuery.FindField('I_BASETPCONTRE').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[5,1]:=StrFMontant(fQuery.FindField('I_BASEECO').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[5,2]:=StrFMontant(fQuery.FindField('I_BASEECOCONTRE').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[6,1]:=StrFMontant(fQuery.FindField('I_BASEFISC').AsFloat,15,V_PGI.OkDecV,'',True);
  G.Cells[6,2]:=StrFMontant(fQuery.FindField('I_BASEFISCCONTRE').AsFloat,15,V_PGI.OkDecV,'',True);
end;

procedure TFDetailEuro.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin G.VidePile (True);end;

end.
