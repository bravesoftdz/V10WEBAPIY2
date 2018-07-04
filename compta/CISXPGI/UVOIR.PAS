unit UVoir;

interface

uses
  Windows, SysUtils, Classes, Graphics, Controls, Forms, 
  ComCtrls, ASCIIV, UScript, ExtCtrls;

type
  TVoirDlg = class(TForm)
        ASCIIView1: TASCIIView;
        StatusBar1: TStatusBar;
        procedure ASCIIView1CaretMove(Sender: TObject);
        procedure FormDestroy(Sender: TObject);
        procedure FormClose(Sender: TObject; var Action: TCloseAction);
        procedure ASCIIView1GetLine(Sender: TObject; var ALine: TASCIILine;
          var AColor: TColor; var AGlyph: Integer; ALineN: Integer);
      private
        { Déclarations privées }
        FTypeCar : TTypeCar;

      protected
        procedure SetTypeCar(AValue:TTypeCar);
      public
        { Déclarations publiques }
        property TypeCar : TTypeCar write SetTypeCar;
  end;

var
  VoirDlg: TVoirDlg;

implementation

{$R *.DFM}


procedure TVoirDlg.SetTypeCar(AValue:TTypeCar);
begin
	if FTypeCar <> AValue then
	begin
		FTypeCar := AValue;
		ASCIIView1.Invalidate;
	end;
end;

procedure TVoirDlg.ASCIIView1CaretMove(Sender: TObject);
begin
	with ASCIIView1 do
		StatusBar1.Panels.Items[0].Text := Format('%6d:%4d', [curRow+1, curCol+1]);
end;

procedure TVoirDlg.FormDestroy(Sender: TObject);
begin
	VoirDlg := nil;
end;

procedure TVoirDlg.FormClose(Sender: TObject; var Action: TCloseAction);
begin
	Action := caFree;
end;

procedure TVoirDlg.ASCIIView1GetLine(Sender: TObject;
  var ALine: TASCIILine; var AColor: TColor; var AGlyph: Integer;
  ALineN: Integer);
begin
	if FTypeCar = tcOEM then
		OemToChar(ALine, ALine);
end;

end.
