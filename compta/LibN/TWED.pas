unit TWED;

interface

uses
	Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
	ExtCtrls, UTobView, UTob, Hctrls, Menus, Db, dbTables;

type
   TTobViewForm = class(TForm)
     MainMenu1: TMainMenu;
     TobViewer1: TMenuItem;
     Parametres1: TMenuItem;
     N1: TMenuItem;
     Quitter1: TMenuItem;
    TobViewerED: TTobViewer;
     procedure Quitter1Click(Sender: TObject);
    procedure TobViewerEDDblClickCell(Sender: TTobViewer; ColName: String;
     ACol, ARow: Integer; DataType: TDataType);
    private
    public
    end;

procedure TobViewED(TobED: TOB);

implementation

{$R *.DFM}

procedure TobViewED(TobED: TOB);
begin
	with TTobViewForm.Create( Application ) do
		try
			TobViewerED.TobView( TobGen, Params );
			if Params <> nil then TobViewerED.SetGridView( Params, True );
			ShowModal;
			if Params <> nil then TobViewerEd.GetGridView( Params );
		finally
			Free;
		end;
end;

procedure TTobViewForm.Quitter1Click(Sender: TObject);
begin
Close;
end;

procedure TTobViewForm.TobViewerEDDblClickCell(Sender: TTobViewer; ColName: String; ACol, ARow: Integer; DataType: TDataType);
begin
ShowMessage( ColName+': '+Sender.AsString[ACol, ARow] );
end;

end.

