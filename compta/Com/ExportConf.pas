unit ExportConf;

interface
uses
  Windows, SysUtils, Classes, Forms,
  Hctrls, HTB97,
  UTOF,Vierge;

type
  TOF_ExportConf = Class (TOF)
  procedure OnLoad ; override ;
  procedure OnArgument(stArgument: String); override ;
  procedure OnClose ; override ;
  procedure bClickValide(Sender: TObject);
  procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
  Result    : string;
end;

implementation

procedure TOF_ExportConf.OnArgument(stArgument: String);
var
St        : string;
indice    : integer;
x,x2      : string;
begin
        indice := 1;
        St := ReadTokenSt(stArgument);
        while St <> '' do
        begin
            if indice = 3 then
            begin
                 x2 := St;
                 x := ReadTokenPipe (x2,'\');
                 while x <> '' do
                 begin
                      x := ReadTokenPipe (x2,'\');
                      if x <> '' then
                       St := 'Fichier : ' + x;
                 end;
            end;
             SetControlText ('EDIT'+IntToStr(indice), St);
             St := ReadTokenSt(stArgument);
             inc (indice);
        end;
     TToolbarButton97(GetControl('BValider')).Onclick := bClickValide;
     Ecran.OnKeyDown := OnKeyDownEcran;
end;

procedure TOF_ExportConf.OnLoad ;
begin
     Result := '-';
end;

procedure TOF_ExportConf.OnClose ;
begin
     TFVierge (Ecran).retour := Result;
end;

procedure TOF_ExportConf.bClickValide(Sender: TObject);
begin
     Result := 'X';
end;

procedure TOF_ExportConf.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F10    : begin bClickValide(Sender); Ecran.ModalResult := 1; end;
  end;
end;


Initialization
RegisterClasses([TOF_ExportConf]);

end.
