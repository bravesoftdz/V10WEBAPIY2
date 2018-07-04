unit ImpAutoParamMulti;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, HSysMenu, HTB97, StdCtrls, Mask, Hctrls, ExtCtrls, HPanel;

type
  TFImpAutoParamMulti = class(TFVierge)
    HPanel1: THPanel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    SCENARIOMULTI: THCritMaskEdit;
    FORMATMULTI: THValComboBox;
    HLabel3: THLabel;
    BSCENARIO: TToolbarButton97;
    procedure FORMATMULTIChange(Sender: TObject);
    procedure BSCENARIOClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Déclarations privées }
    FFormat : string;
    FScenario : string;
  public
    { Déclarations publiques }
  end;

procedure LanceFicheParametrageImportMulti ( var vFormat : string; var vScenario : string );

implementation

uses LookUp,ScenaCom, HEnt1;

{$R *.DFM}

procedure LanceFicheParametrageImportMulti ( var vFormat : string; var vScenario : string );
var  FImpAutoParamMulti: TFImpAutoParamMulti;
begin
  FImpAutoParamMulti := TFImpAutoParamMulti.Create ( Application ) ;
  try
    FImpAutoParamMulti.ShowModal;
  finally
    vFormat := FImpAutoParamMulti.FFormat;
    vScenario := FImpAutoParamMulti.FScenario;
    FImpAutoParamMulti.Free;
  end;
end;

procedure TFImpAutoParamMulti.FORMATMULTIChange(Sender: TObject);
begin
  inherited;
  SCENARIOMULTI.Plus:='FS_IMPORT="X" AND FS_NATURE="FEC" AND FS_FORMAT="'+FORMATMULTI.Value+'" ';
end;

procedure TFImpAutoParamMulti.BSCENARIOClick(Sender: TObject);
begin
  inherited;
  ParamSupImport('X','FEC',FORMATMULTI.Value,SCENARIOMULTI.Text,taConsult,0,TRUE);
end;

procedure TFImpAutoParamMulti.BValiderClick(Sender: TObject);
begin
  inherited;
  FFormat := FORMATMULTI.Value;
  FScenario := SCENARIOMULTI.Text;
end;

procedure TFImpAutoParamMulti.BFermeClick(Sender: TObject);
begin
  inherited;
  FFormat := '';
  FScenario := '';
end;

procedure TFImpAutoParamMulti.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  inherited;
  CanClose := True;
end;

end.
