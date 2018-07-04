unit YNewEmail;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, HSysMenu, HTB97, StdCtrls, Hctrls, Mask, ExtCtrls, HMsgBox, Hent1;

type
  TFNewEmail = class(TFVierge)
    Panel1: TPanel;
    Email: THCritMaskEdit;
    Nom: THCritMaskEdit;
    HLabel1: THLabel;
    HLabel2: THLabel;
    procedure BValiderClick(Sender: TObject);
    procedure EmailKeyPress(Sender: TObject; var Key: Char);
  end;

function LancerNewEmail (var SEmail : String; var SNom : String) : Boolean;

implementation

{$R *.DFM}

//------------------------------------
//--- Nom   : LancerNewListeDistrib
//--- Objet :
//------------------------------------
function LancerNewEmail (var SEmail : String; var SNom : String) : Boolean;
var FNewEmail: TFNewEmail;
begin
 FNewEmail:= TFNewEmail.Create(Application);
 FNewEmail.Email.Text:=SEmail;
 FNewEmail.Nom.Text:=Snom;
 try
  Result:=(FNewEmail.ShowModal = mrOk);
  SEmail:= FNewEmail.Email.Text;
  SNom:= FNewEmail.Nom.Text;
 finally
  FNewEmail.Free;
 end;
end;

//------------------------------------
//--- Nom   : BValiderClick
//--- Objet :
//------------------------------------
procedure TFNewEmail.BValiderClick(Sender: TObject);
var SEmail     : String;
    LaPosition : Integer;
begin
 inherited;
 SEmail:=Email.Text;
 LaPosition:=pos ('@',SEmail);
 if (LaPosition=0) then
  begin
   PgiInfo ('L''adresse Email n''est pas valide.', TitreHalley);
   Email.SetFocus;
   Exit;
  end;

 SEmail [LaPosition]:='0';
 LaPosition:=pos ('@',SEmail);
 if (LaPosition>0) then
  begin
   PgiInfo ('L''adresse Email n''est pas valide.', TitreHalley);
   Email.SetFocus;
   Exit;
  end;

 ModalResult:=MrOk;
end;

//------------------------------------
//--- Nom   : EmailKeyPress
//--- Objet :
//------------------------------------
procedure TFNewEmail.EmailKeyPress(Sender: TObject; var Key: Char);
begin
 inherited;
 if (Key in [Chr (VK_RETURN),Chr (VK_ESCAPE), chr (VK_PRIOR), Chr (VK_NEXT), chr (VK_END), chr (VK_HOME), chr (VK_UP), chr (VK_DOWN), chr (VK_BACK)]) then Exit;
 if not ((Key in ['0'..'9']) or (Key in ['a'..'z']) or (Key in ['A'..'Z']) or (Key in ['@','_','.','-'])) then key:=#0;
end;

end.
