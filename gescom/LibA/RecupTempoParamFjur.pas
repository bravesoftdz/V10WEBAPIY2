unit RecupTempoParamFjur;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,Dicobtp, hmsgbox, HSysMenu, ExtCtrls, HPanel, Grids, Hctrls,UTOB,LookUp;

type
  TFParamFjur = class(TForm)
    HMtrad: THSystemMenu;
    GSFJUR: THGrid;
    HPanel1: THPanel;
    HPanel2: THPanel;
    FermerFJur: TButton;
    procedure FormShow(Sender: TObject);
    procedure FermerFJurClick(Sender: TObject);
    procedure GSFJURElipsisClick(Sender: TObject);
    procedure GSFJURCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    TobFJur : TOB;
  end;

implementation

{$R *.DFM}

procedure TFParamFjur.FormShow(Sender: TObject);
begin
TobFJur.PutGridDetail (GSFJUR,True,True, '');
end;

procedure TFParamFjur.FermerFJurClick(Sender: TObject);
begin
TobfJur.GetGridDetail(GSFJUR, TobFjur.detail.count,'','CODE;LIBELLE;Nouveau code ');
Self.close;
end;

procedure TFParamFjur.GSFJURElipsisClick(Sender: TObject);
begin
if (GSFJur.Col = 2) then
   LookupList(GSFJUR,'Formes juridiques','COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="YFJ"','CO_CODE',False,-1);
end;


procedure TFParamFjur.GSFJURCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var lib : string;
begin
if ACol = 2 then
  BEGIN
  if GSFJUR.Cells[ACol,ARow] <> '' then
    BEGIN
    Lib := (RechDom('YYCODEJURIDIQUE', GSFJUR.Cells[ACol,ARow],False));
    if (lib = '') or (lib = 'Error') then
       BEGIN
       PGIBoxAF ('Forme juridique incorrecte', 'Correspondance des formes juridiques');
       GSFJUR.Cells[ACol,ARow] :='';
       END;
    END;
  END;
end;

end.
