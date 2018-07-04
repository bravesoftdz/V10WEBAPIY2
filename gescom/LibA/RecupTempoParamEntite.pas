unit RecupTempoParamEntite;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}

{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  StdCtrls,Dicobtp, hmsgbox, HSysMenu, ExtCtrls, HPanel, Grids, Hctrls,UTOB,LookUp;

type
  TFParamEntite = class(TForm)
    HMtrad: THSystemMenu;
    GSENTITE: THGrid;
    HPanel1: THPanel;
    HPanel2: THPanel;
    FermerEntite: TButton;
    procedure FormShow(Sender: TObject);
    procedure FermerEntiteClick(Sender: TObject);
    procedure GSENTITEElipsisClick(Sender: TObject);
    procedure GSENTITECellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    TobEntite : TOB;
  end;

implementation

{$R *.DFM}

procedure TFParamEntite.FormShow(Sender: TObject);
begin
TobEntite.PutGridDetail (GSENTITE,True,True, '');
end;

procedure TFParamEntite.FermerEntiteClick(Sender: TObject);
begin
TobEntite.GetGridDetail(GSENTITE, TobEntite.detail.count,'','CODE;LIBELLE;Etablissement ');
Self.close;
end;

procedure TFParamEntite.GSENTITEElipsisClick(Sender: TObject);
begin
if (GSENTITE.Col = 2) then
   LookupList(GSENTITE,'Etablissement','ETABLISS','ET_ETABLISSEMENT','ET_LIBELLE','','ET_ETABLISSEMENT',False,-1);
end;


procedure TFParamEntite.GSENTITECellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var //lib : string;
    QQ : TQuery;
begin
if ACol = 2 then
  BEGIN
  if GSENTITE.Cells[ACol,ARow] <> '' then
    BEGIN
    GSENTITE.Cells[ACol,ARow]:=AnsiUppercase (GSENTITE.Cells[ACol,ARow]);
     QQ := OpenSQL('SELECT  ET_ETABLISSEMENT From ETABLISS WHERE ET_ETABLISSEMENT="'+ GSENTITE.Cells[ACol,ARow]+'"',TRUE);
      If QQ.EOF then
       BEGIN
       PGIBoxAF ('Etablissement incorrect', 'Correspondance des entités');
       GSENTITE.Cells[ACol,ARow] :='';
       END;
      Ferme(QQ);
   END;
  end;
end;

end.
