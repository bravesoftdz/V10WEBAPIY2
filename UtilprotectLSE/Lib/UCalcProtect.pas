unit UCalcProtect;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, ExtCtrls, Mask, Hctrls, StdCtrls, FileCtrl,HmsgBox;

type
  TFCalcProtec = class(TForm)
    Panel1: TPanel;
    BABANDON: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    IDCLIENT: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    EMPLACESTO: THCritMaskEdit;
    Label3: TLabel;
    DATEFIN: THCritMaskEdit;
    procedure EMPLACESTOElipsisClick(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure OuvreFicheCalc;

implementation

{$R *.dfm}
uses UdroitUtilisation;

procedure OuvreFicheCalc;
var XX : TFCalcProtec;
begin
  XX := TFCalcProtec.create(Application);
  TRY
    XX.ShowModal;
  FINALLY
    XX.Free;
  END;
end;

procedure TFCalcProtec.EMPLACESTOElipsisClick(Sender: TObject);
var Emplace : string;
begin
  SelectDirectory('Entrer le repértoire de destination', 'C:\',Emplace);
  if Emplace <> '' then EMPLACESTO.Text := Emplace;
end;

procedure TFCalcProtec.ToolbarButton971Click(Sender: TObject);
var DateNow : TdateTime;
begin
  DateNow := StrToDate(DateToStr(Now)); 
  IF (IDCLIENT.text = '') {or (StrToDate(DATEFIN.Text)<=DateNow)} or (EMPLACESTO.Text='') then
  begin
    PgiInfo('Veuillez vérifier les informations saisies');
    exit;
  end;
  //
  TRY
    DefiniFichierDroits (EMPLACESTO.Text,IDCLIENT.Text,DATEFIN.Text);
    PgiInfo('Fichier de licence généré.');
  FINALLY
    Close;
  end;
end;

end.
