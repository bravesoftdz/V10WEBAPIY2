unit UTofIntegrInv;

interface

uses UTOF, UTOB;

type
     TOF_IntegrInv = class(TOF)
      TOBSelection : TOB;

      function PutSelectionIntoTOB : boolean;

     published
      procedure OnArgument(stArgument : String); override;
      procedure OnIntegre;

     end;

procedure AGLOnIntegreInv(Parms : Array of Variant; Nb : Integer);


implementation

uses Classes, Controls, Forms, HCtrls, HMsgBox, M3FP,
{$IFNDEF EAGLCLIENT}
     Mul,
{$ELSE}
     eMul,
{$ENDIF}
     IntegreInv, ParamSoc, SysUtils
     ;

const FMess : Array[0..0] of string = ('Veuillez sélectionner au moins une transmission à intégrer');


procedure AGLOnIntegreInv(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                else exit;
if (TOTOF is TOF_IntegrInv) then TOF_IntegrInv(TOTOF).OnIntegre
                            else exit;
end;

function TOF_IntegrInv.PutSelectionIntoTOB : boolean;
var i : integer;
begin
result := true;
with TFMul(Ecran) do
  begin
  if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
    PGIInfo(FMess[0], Caption);
    result := false;
    exit;
    end;

  TOBSelection := TOB.Create('toto', nil, -1);
{$IFDEF EAGLCLIENT}
  Q.TQ.First;
  if FListe.AllSelected then
    TOBSelection.LoadDetailDB('TRANSINVENT', '', '', Q.TQ,false)
{$ELSE}
  Q.DisableControls;
  Q.FindFirst;
  if FListe.AllSelected then
    TOBSelection.LoadDetailDB('TRANSINVENT', '', '', Q,false)
{$ENDIF}
   else
    begin
    for i := 0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.Row-1) ;
      TOB.Create('TRANSINVENT', TOBSelection, -1).SelectDB('',Q.TQ, true);
{$ELSE}
      TOB.Create('TRANSINVENT', TOBSelection, -1).SelectDB('', Q, true);
{$ENDIF}
      end;
    end;
{$IFDEF EAGLCLIENT}
{$ELSE}
  Q.EnableControls;
{$ENDIF}
  end;
end;


procedure TOF_IntegrInv.OnArgument(stArgument : String);
begin
inherited;
SetControlText('XX_WHERE', 'GIT_INTEGRATION<>"X"');
end;

procedure TOF_IntegrInv.OnIntegre;
begin
if not PutSelectionIntoTOB then exit;
IntegreLesInventaires(TOBSelection);
TFMul(Ecran).ChercheClick;
TOBSelection.Free; // Create dans PutSelectionIntoTOB
end;

initialization
RegisterClasses([TOF_IntegrInv]);
RegisterAGLProc('IntegreInv', true, 0, AGLOnIntegreInv);

end.
