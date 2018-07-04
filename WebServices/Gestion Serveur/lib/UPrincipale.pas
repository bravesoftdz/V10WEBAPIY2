unit UPrincipale;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Buttons, HTB97, Hctrls, Grids, UgestionINI,UTOB,
  HSysMenu,Udefinitions;

type
  TFprincipale = class(TForm)
    PBAS: TPanel;
    BANNULE: TToolbarButton97;
    GS: THGrid;
    BNEW: TToolbarButton97;
    HmTrad: THSystemMenu;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BANNULEClick(Sender: TObject);
    procedure BNEWClick(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
  private
    fListeInfo : string;
    fgestionIni : TGestionINI;
    fTOBListe : TOB;
    procedure 	AfficheLagrille;
    procedure  ConstitueGrille;
		procedure  PositionneGrid(ID : string);

  public
    { Déclarations publiques }
  end;

var
  Fprincipale: TFprincipale;

implementation
uses UNewCollab,fTest;

{$R *.dfm}


procedure TFprincipale.FormClose(Sender: TObject; var Action: TCloseAction);
begin
	fgestionIni.Free;
end;

procedure TFprincipale.FormShow(Sender: TObject);
begin
	fListeInfo := 'SEL;ID;LIBELLE';
	fgestionIni := TGestionINI.create;
  fTOBListe := fgestionIni.GetTOBList;
  AfficheLagrille;
  HmTrad.ResizeGridColumns(GS);
end;

procedure TFprincipale.BANNULEClick(Sender: TObject);
begin
	close;
end;

procedure TFprincipale.ConstitueGrille;
begin
  GS.ColCount := 3;
  GS.ColWidths [0] := 10;
  GS.cells[1,0] := 'Identification';
  GS.ColWidths [1] := 200;
  GS.cells[2,0] := 'Collaborateur';
  GS.ColWidths [2] := 400;
end;

procedure TFprincipale.BNEWClick(Sender: TObject);
var NewEnreg : TEnreg;
begin
  NewEnreg := nil;
//  Ouvretest;
	AppelAjout (NewEnreg, fgestionIni);
  if NewEnreg <> nil then
  begin
  	AfficheLagrille;
  	PositionneGrid(NewEnreg.ID);
  end;
end;


procedure TFprincipale.AfficheLagrille;
var Indice : Integer;
begin
  GS.VidePile(false);
  ConstitueGrille;
  if fTOBListe.Detail.count > 0 then GS.RowCount := fTOBListe.Detail.count +1
  															else GS.RowCount := 2;
	for Indice := 0 to fTOBListe.Detail.count -1 do
  begin
    fTOBListe.PutGridDetail(GS,false,false,fListeInfo);
  end;
end;

procedure TFprincipale.PositionneGrid(ID: string);
var Indice : Integer;
begin
  GS.BeginUpdate;
  GS.Row := 1;
	for Indice := 0 to fTOBListe.detail.Count -1 do
  begin
		if fTOBListe.detail[Indice].GetString('ID')=ID then
    begin
      GS.Row := Indice +1;
      break;
    end;
  end;
  GS.EndUpdate;
end;

procedure TFprincipale.GSDblClick(Sender: TObject);
begin
	if GS.Row > fTOBListe.Detail.count then Exit;
  if not AppelModif (fTOBListe.detail[GS.row-1].GetString('ID'),fgestionIni) then
  begin
		AfficheLagrille;
    GS.row := 1;
  end;
end;

end.
