unit EtatTransLB;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, ExtCtrls, Grids, Hctrls, HEnt1 ;

  procedure LanceEtatTransLb ;

type
  TFEtatTransLB = class(TForm)
    GS: THGrid;
    Panel1: TPanel;
    Dock971: TDock97;
    PanelBouton: TToolWindow97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}
procedure LanceEtatTransLb ;
var  X : TFEtatTransLB;
BEGIN
SourisSablier;
X:=TFEtatTransLB.Create(Application) ;
  try
  X.ShowModal ;
  finally
  X.Free ;
  end ;
SourisNormale ;
END ;

procedure TFEtatTransLB.FormShow(Sender: TObject);
   VAR F : TextFile ;
       St, LigneLue : string ;
       okok : boolean ;
       i : integer;
begin
   // Chargement du THGrid
   okok := false ;
   // ouverture de  LBETATTRANS.ini
   St:= ExtractFileDir(Application.exename);
   St:= St+'\LBETATCOM.log'     ;
   AssignFile(F,St) ;
   If FileExists (St) then
   Begin
      Reset(F); // ouverture en lecture
      i:= 1 ;
      While Not Eof(F) do   // Recherche du fichier télétransmis dans LBETATTRANS.INI
      begin
         Readln(F,LigneLue);
         if okok then i:= i+1 else okok := true ;
         GS.Rowcount := i+1;
         GS.Cells[0,i]:= Copy(LigneLue,1,20) ;
         GS.Cells[1,i]:= Copy(LigneLue,20,40) ;
         GS.Cells[2,i]:=Copy(LigneLue,60,25) ;
         GS.Cells[3,i]:=Copy(LigneLue,85,20) ;
         GS.Cells[4,i]:=Copy(LigneLue,105,2) ;
         end;

      CloseFile(F);
   end;
end;

end.
