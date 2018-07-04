unit RecupGBDelete;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, ExtCtrls, DicoAf, Hctrls, Spin, Hent1, UTOB;

type
  TFRecupDelete = class(TForm)
    FermerParamFicBase: TButton;
    GroupBox4: TGroupBox;
    LabelDelDim1: TLabel;
    LabelDelDim3: TLabel;
    LabelDelDim4: TLabel;
    LabelDelDim5: TLabel;
    LabelDelDim2: TLabel;
    Label2: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    DelDim1: TCheckBox;
    DelMasqueDim: TCheckBox;
    DelPays: TCheckBox;
    DelArrondis: TCheckBox;
    DelDevises: TCheckBox;
    DelBoutiques: TCheckBox;
    CheckBox5: TCheckBox;
    DelArticles: TCheckBox;
    DelGrilleDim1: TCheckBox;
    DelFamilles: TCheckBox;
    DelCodesPostaux: TCheckBox;
    DelGrilleDim3: TCheckBox;
    DelDim3: TCheckBox;
    DelGrilleDim4: TCheckBox;
    DelDim4: TCheckBox;
    DelGrilleDim5: TCheckBox;
    DelDim5: TCheckBox;
    DelGrilleDim2: TCheckBox;
    DelDim2: TCheckBox;
    ButtonDelete: TButton;
    DelTableA1: TCheckBox;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    DelTableA2: TCheckBox;
    DelTableA3: TCheckBox;
    DelTableA4: TCheckBox;
    DelTableA5: TCheckBox;
    DelTableA6: TCheckBox;
    DelTableA7: TCheckBox;
    DelTableA8: TCheckBox;
    DelTableA9: TCheckBox;
    DelTableA10: TCheckBox;
    DelCollections: TCheckBox;
    DelPoids: TCheckBox;
    DelDemarques: TCheckBox;
    DelCivilites: TCheckBox;
    DelRepresentants: TCheckBox;
    DelClients: TCheckBox;
    DelFournisseurs: TCheckBox;
    DelTableC1: TCheckBox;
    DelTableC2: TCheckBox;
    DelTableC3: TCheckBox;
    DelTableC4: TCheckBox;
    DelTableC5: TCheckBox;
    DelTableC6: TCheckBox;
    DelTableC7: TCheckBox;
    DelTableC8: TCheckBox;
    DelTableC9: TCheckBox;
    DelTableC10: TCheckBox;
    DelCategories: TCheckBox;
    procedure FermerParamFicBaseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ButtonDeleteClick(Sender: TObject);
  private
    ButtonDeleteStatut : boolean;
    Tob_parametrage : TOB;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

Procedure AppelDeleteParam ( var TOBParam : TOB ) ;

implementation

{$R *.DFM}

Procedure AppelDeleteParam ( var TOBParam : TOB ) ;
var X  : TFRecupDelete ;
BEGIN
SourisSablier;
X:=TFRecupDelete.Create(Application) ;
X.Tob_Parametrage:=TOBParam ;
X.Tob_Parametrage.PutEcran(X) ;
try
 X.ShowModal ;
 finally
 X.Free ;
 end ;
SourisNormale ;
END ;

procedure TFRecupDelete.FermerParamFicBaseClick(Sender: TObject);
begin
Tob_Parametrage.GetEcran(Self) ;
Self.Close;
end;

procedure TFRecupDelete.FormShow(Sender: TObject);
begin
//TraduitForm(self);
end;

procedure TFRecupDelete.ButtonDeleteClick(Sender: TObject);
begin
  if ButtonDeleteStatut=TRUE
  then ButtonDelete.Caption:='Tout supprimer'
  else ButtonDelete.Caption:='Ne rien supprimer';
  ButtonDeleteStatut := not ButtonDeleteStatut ;
  DelArrondis.checked := ButtonDeleteStatut ;
  DelBoutiques.checked := ButtonDeleteStatut ;
  DelPays.checked := ButtonDeleteStatut ;
  DelCodesPostaux.checked := ButtonDeleteStatut ;
  DelDevises.checked := ButtonDeleteStatut ;
  DelArrondis.checked := ButtonDeleteStatut ;
  DelFamilles.checked := ButtonDeleteStatut ;
  DelTableA1.checked := ButtonDeleteStatut ;
  DelTableA2.checked := ButtonDeleteStatut ;
  DelTableA3.checked := ButtonDeleteStatut ;
  DelTableA4.checked := ButtonDeleteStatut ;
  DelTableA5.checked := ButtonDeleteStatut ;
  DelTableA6.checked := ButtonDeleteStatut ;
  DelTableA7.checked := ButtonDeleteStatut ;
  DelTableA8.checked := ButtonDeleteStatut ;
  DelTableA9.checked := ButtonDeleteStatut ;
  DelTableA10.checked := ButtonDeleteStatut ;
  DelTableC1.checked := ButtonDeleteStatut ;
  DelTableC2.checked := ButtonDeleteStatut ;
  DelTableC3.checked := ButtonDeleteStatut ;
  DelTableC4.checked := ButtonDeleteStatut ;
  DelTableC5.checked := ButtonDeleteStatut ;
  DelTableC6.checked := ButtonDeleteStatut ;
  DelTableC7.checked := ButtonDeleteStatut ;
  DelTableC8.checked := ButtonDeleteStatut ;
  DelTableC9.checked := ButtonDeleteStatut ;
  DelTableC10.checked := ButtonDeleteStatut ;
  DelMasqueDim.checked := ButtonDeleteStatut ;
  DelGrilleDim1.checked := ButtonDeleteStatut ;
  DelGrilleDim2.checked := ButtonDeleteStatut ;
  DelGrilleDim3.checked := ButtonDeleteStatut ;
  DelGrilleDim4.checked := ButtonDeleteStatut ;
  DelGrilleDim5.checked := ButtonDeleteStatut ;
  DelDim1.checked := ButtonDeleteStatut ;
  DelDim2.checked := ButtonDeleteStatut ;
  DelDim3.checked := ButtonDeleteStatut ;
  DelDim4.checked := ButtonDeleteStatut ;
  DelDim5.checked := ButtonDeleteStatut ;
  DelCollections.checked := ButtonDeleteStatut ;
  DelPoids.checked := ButtonDeleteStatut ;
  DelDemarques.checked := ButtonDeleteStatut ;
  DelCivilites.checked := ButtonDeleteStatut ;
  DelRepresentants.checked := ButtonDeleteStatut ;
  DelCategories.checked := ButtonDeleteStatut ;
  DelArticles.checked := ButtonDeleteStatut ;
  DelClients.checked := ButtonDeleteStatut ;
  DelFournisseurs.checked := ButtonDeleteStatut ;  
end;

end.
