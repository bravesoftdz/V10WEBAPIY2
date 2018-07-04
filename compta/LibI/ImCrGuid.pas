unit ImCrGuid;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, HPanel, Buttons, ExtCtrls, UTob, Mask;

type
  TFImCrGuid = class(TForm)
    PanelBouton: TPanel;
    Panel1: TPanel;
    BValider: TBitBtn;
    BAnnuler: TBitBtn;
    BAide: TBitBtn;
    HPanel1: THPanel;
    GroupBox4: TGroupBox;
    tJEcheance: THLabel;
    tCTva: THLabel;
    cbEcrEcheance: TCheckBox;
    JEcheance: THCritMaskEdit;
    CTva: THCritMaskEdit;
    GroupBox6: TGroupBox;
    tCBanque: THLabel;
    tJBanque: THLabel;
    cbEcrPaiement: TCheckBox;
    JBanque: THCritMaskEdit;
    CBanque: THCritMaskEdit;
    LJEecheance: THLabel;
    LCTva: THLabel;
    LJBanque: THLabel;
    LCBanque: THLabel;
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées}
    fObAbo : TOB;
  public
    { Déclarations publiques}
  end;

procedure CreationGuideEcheance (OBAbo : TOB);

implementation

{$R *.DFM}

procedure CreationGuideEcheance (OBAbo : TOB);
var  FImCrGuid: TFImCrGuid;
begin
  FImCrGuid:= TFImCrGuid.Create (Application);
  FImCrGuid.fOBAbo := ObAbo;
  try
    FImCrGuid.ShowModal;
  finally
    ObAbo := FImCrGuid.fOBAbo;
    FImCrGuid.Free;
  end;
end;

procedure TFImCrGuid.BValiderClick(Sender: TObject);
begin
//  GenereTOBGuide;
end;
{
procedure TFImCrGuid.GenereTOBGuide;
var OBEcr : TOB;
    dNumLigne : integer;
begin
  if cbEcrEcheance.Checked then
  begin
    OBEcr := TOB.Create ('ECRGUI', ? , - 1);
    OBEcr.PutValue ('EG_TYPE','ABO');
    OBEcr.PutValue ('EG_GUIDE',?);
    OBEcr.PutValue ('EG_NUMLIGNE',0);
    OBEcr.PutValue ('EG_GENERAL',);
    OBEcr.PutValue ('EG_LIBELLE',);
    OBEcr.PutValue ('EG_DEBITDEV',);
    OBEcr.PutValue ('EG_CREDITDEV',);
    OBEcr.PutValue ('EG_MODEPAIE',);
  end;
  if cbEcrPaiement.Checked then
  begin
  end;
end;
}

end.

