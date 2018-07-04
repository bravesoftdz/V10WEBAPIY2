unit saisieTablecorresp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Hctrls, UTOB, Hent1;

type
  TSaisieTable = class(TForm)
    FermerParamFicBase: TButton;
    GroupBox4: TGroupBox;
    GridReg: THGrid;
    GroupBox5: TGroupBox;
    GridPays: THGrid;
    GroupBox1: TGroupBox;
    GridDevise: THGrid;
    GroupBox2: TGroupBox;
    GridFamille: THGrid;
    procedure FermerParamFicBaseClick(Sender: TObject);
  private
    { Déclarations privées }
  public
   { Déclarations publiques }
    Tob_Reglement : TOB ;
    Tob_Pays      : TOB ;
    Tob_Devise    : TOB ;
    Tob_Famille   : TOB ;
  end;

var
  SaisieTable: TSaisieTable;


Procedure AppelSaisieCorrespondance ( var TOBReg, TOBPays, TOBDevise, TOBFamille: TOB ) ;

implementation

{$R *.DFM}

Procedure AppelSaisieCorrespondance ( var TOBReg, TOBPays , TOBDevise, TOBFamille : TOB ) ;
var X  : TSaisieTable ;
BEGIN
SourisSablier;
X:=TSaisieTable.Create(Application) ;
X.Tob_Reglement:= TOBReg  ;
X.Tob_Pays     := TOBPays ;
X.Tob_Devise   := TOBDevise ;
X.Tob_Famille  := TOBFamille ;

// Affichage
X.GridReg.ColLengths  [0] := -1;
X.GridReg.ColLengths  [1] := -1;
X.GridReg.ColFormats  [2] := 'CB=GCMODEPAIE';

X.GridPays.ColLengths [0] := -1;
X.GridPays.ColLengths [1] := -1;
X.GridPays.ColFormats [2] := 'CB=TTPAYS';

X.GridDevise.ColLengths [0] := -1;
X.GridDevise.ColLengths [1] := -1;
X.GridDevise.ColFormats [2] := 'CB=TTDEVISETOUTES';

X.GridFamille.ColLengths [0] := -1;
X.GridFamille.ColLengths [1] := -1;
X.GridFamille.ColFormats [2] := 'CB=GCFAMILLENIV1';
X.GridFamille.ColFormats [3] := 'CB=GCFAMILLENIV2';
X.GridFamille.ColFormats [4] := 'CB=GCFAMILLENIV3';

X.Tob_Reglement.PutGridDetail (X.GridReg    , False, False, 'CODEGB;LIBELLE;CODEPGI', True);
X.Tob_Pays.PutGridDetail      (X.GridPays   , False, False, 'CODEGB;LIBELLE;CODEPGI', True);
X.Tob_Devise.PutGridDetail    (X.GridDevise , False, False, 'CODEGB;LIBELLE;CODEPGI', True);
X.Tob_Famille.PutGridDetail   (X.GridFamille, False, False, 'CODEGB;LIBELLE;CODEPGI1;CODEPGI2;CODEPGI3', True);

try
 X.ShowModal ;
 finally
 X.Free ;
 end ;
SourisNormale ;
END ;



procedure TSaisieTable.FermerParamFicBaseClick(Sender: TObject);
begin
  if Tob_Reglement.Detail.count > 0 then Tob_Reglement.GetGridDetail (GridReg, Tob_Reglement.Detail.count, '', 'CODEGB;LIBELLE;CODEPGI');
  if Tob_Pays.Detail.count > 0      then Tob_Pays.GetGridDetail (GridPays, Tob_Pays.Detail.count, '', 'CODEGB;LIBELLE;CODEPGI');
  if Tob_Devise.Detail.count > 0   then Tob_Devise.GetGridDetail (GridDevise, Tob_Devise.Detail.count, '', 'CODEGB;LIBELLE;CODEPGI');
  if Tob_Famille.Detail.count > 0   then Tob_Famille.GetGridDetail (GridFamille, Tob_Famille.Detail.count, '', 'CODEGB;LIBELLE;CODEPGI1;CODEPGI2;CODEPGI3');
  Self.Close;
end;

end.
