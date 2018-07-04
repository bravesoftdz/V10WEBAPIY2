unit ULanceFiche;

interface

uses
   Controls,

   Forms, UIUtil, HPanel;

function LanceLaFiche(FFiche_p : TForm) : TModalResult;

implementation

function LanceLaFiche(FFiche_p : TForm) : TModalResult;
var
   CurPanel_l : THPanel;
begin
   result := mrNone;
   CurPanel_l := FindInsidePanel;
   if CurPanel_l = nil then
   begin
      try
         FFiche_p.ShowModal;
      finally
         result := FFiche_p.ModalResult;
         FFiche_p.Free ;
      end;
   end
   else
   begin
      InitInside(FFiche_p, CurPanel_l);
      FFiche_p.Show;
      FFiche_p.SetFocus;
      result := FFiche_p.ModalResult;
   end;
end;
end.
 