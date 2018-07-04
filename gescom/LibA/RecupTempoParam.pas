unit RecupTempoParam;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, ExtCtrls, Dicobtp, Hctrls, HSysMenu;

type
  TFRecupTempoParam = class(TForm)
    GroupBoxArticle: TGroupBox;
    GroupBoxEmploye: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    CheckStat1Prest: TCheckBox;
    CheckStat2Prest: TCheckBox;
    CheckStat1Emp: TCheckBox;
    CheckStat2Emp: TCheckBox;
    Label4: TLabel;
    Label5: TLabel;
    GroupBoxAffaire: TGroupBox;
    Label6: TLabel;
    CheckStat1Aff: TCheckBox;
    CheckStat2Aff: TCheckBox;
    NbStatEmp: TLabel;
    NbStatPrest: TLabel;
    NbStatAff: TLabel;
    FermerParamFicBase: TButton;
    LabelFamilleAff: TLabel;
    RadioFamAffStat: TRadioButton;
    RadioFamAffNon: TRadioButton;
    LabelLiaisonIsis: TLabel;
    LabelIsis: TLabel;
    CheckChangementIsis: TCheckBox;
    Shape1: TShape;
    CheckCodePres: TCheckBox;
    AffCompt: THLabel;
    ComboAffCompt: TComboBox;
    PreCompt: TLabel;
    ComboPreCompt: TComboBox;
    Hmtrad: THSystemMenu;
    ComboStat1Aff: TComboBox;
    ComboStat2Aff: TComboBox;
    CHeckStat3Aff: TCheckBox;
    ComboStat3aff: TComboBox;
    CheckStat4Aff: TCheckBox;
    ComboStat4Aff: TComboBox;
    CheckApport: TCheckBox;
    ComboStat1Emp: TComboBox;
    ComboStat2Emp: TComboBox;
    CheckStat3Emp: TCheckBox;
    ComboStat3Emp: TComboBox;
    CheckStat4Emp: TCheckBox;
    ComboStat4Emp: TComboBox;
    procedure FermerParamFicBaseClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure TFRecupTempoParam.FermerParamFicBaseClick(Sender: TObject);
begin
Self.Close;
end;

end.
