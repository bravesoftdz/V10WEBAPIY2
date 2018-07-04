{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/01/2007
Modifié le ... :   /  /    
Description .. : Reprise d'un dossier avec fichiers d'origine à l'image des 
Suite ........ : export PGI
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
unit ImportTOB;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, assist, HSysMenu, hmsgbox, StdCtrls, Hctrls, ComCtrls, ExtCtrls,
  HPanel, HTB97, Grids;

type
  TFImportTob = class(TFAssist)
    PCodageSal: TTabSheet;
    TLCodage: TLabel;
    FCodeSal: TRadioGroup;
    PCumuls: TTabSheet;
    BCumuls: TToolbarButton97;
    FCumul: TCheckBox;
    GCumuls: THGrid;
    PLancer: TTabSheet;
    Label3: TLabel;
    BLancer: TToolbarButton97;
    PReport: TTabSheet;
    FReport: TListBox;
    PCreation: TTabSheet;
    PSelect: TTabSheet;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FImportTob: TFImportTob;

implementation

{$R *.dfm}

end.
