unit AffList;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, UiUtil, HTB97,UTOb,PrintDBG, HPanel, Hctrls;

Procedure AfficheListe(LaListe: Tob; LeTitre: String; const LesTitresGrid: array of string);
//Procedure AfficheListe2(LaListe: Tob; LeTitre: String; const LesTitresGrid: array of string);


type
  TFAffList = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bDefaire: TToolbarButton97;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    Fliste: THGrid;
    procedure FormShow(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
  private
    { Déclarations privées }
    Liste: Tob ;
    Titre: string ;
    TitreGrid: boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}
Procedure AfficheListe(LaListe: Tob; LeTitre: String; const LesTitresGrid: array of string );
var FAffList: TFAffList ; i: integer ;
begin
if LaListe=nil then exit ;
FAffList:=TFAffList.Create(Application) ;
With FAffList do begin
  Liste:=LaListe ;
  Titre:=LeTitre ;
  TitreGrid:=True ;
  if High(lesTitresGrid)<>-1 then begin
    for i:=0 to high(lesTitresGrid) do Fliste.Cells[i,0]:=LesTitresGrid[i] ;
    TitreGrid:=False ;
    end ;
  end ;

try
  FAffList.ShowModal ;
finally
  FAffList.Free ;
  Screen.Cursor:=crDefault ;
  end ;
end ;


procedure TFAffList.FormShow(Sender: TObject);
begin
Caption:=Titre ;
Liste.PutGridDetail(Fliste,TitreGrid,True,'',true) ;
end;

procedure TFAffList.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(FListe,nil,Titre,'') ;
end;

end.
