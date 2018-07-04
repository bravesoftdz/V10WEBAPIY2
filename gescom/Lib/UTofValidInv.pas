unit UTofValidInv;

interface

uses UTOF, UTOB;

type
     TOF_ValidInv = class(TOF)
      TOBSelection : TOB;

      function PutSelectionIntoTOB : boolean;

     published
      procedure OnArgument(stArgument : String); override;
      procedure OnValidate;

     end;

procedure AGLOnValidateList(Parms : Array of Variant; Nb : Integer);

const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Vous devez renseigner le Tiers Ecarts d''inventaire dans les Paramètes Société'
          {2} ,''
              );

implementation

uses Classes, Controls, Forms, HCtrls, HMsgBox, M3FP, EntGC, HEnt1,
{$IFNDEF EAGLCLIENT}
     Mul,
{$ELSE}
     eMul,
{$ENDIF}
     AssistValidInv, ParamSoc, SysUtils
     ;

const FMess : Array[0..0] of string = ('Veuillez sélectionner au moins une liste à valider');


procedure AGLOnValidateList(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                else exit;
if (TOTOF is TOF_ValidInv) then TOF_ValidInv(TOTOF).OnValidate
                           else exit;
end;

function TOF_ValidInv.PutSelectionIntoTOB : boolean;
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
  if FListe.AllSelected then
    TOBSelection.LoadDetailDB('LISTEINVENT', '', '', Q.TQ,false)
{$ELSE}
  Q.DisableControls;
  if FListe.AllSelected then
    TOBSelection.LoadDetailDB('LISTEINVENT', '', '', Q,false)
{$ENDIF}
   else
    begin
    for i := 0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.Row-1) ;
      TOB.Create('LISTEINVENT', TOBSelection, -1).SelectDB('',Q.TQ, true);
{$ELSE}
      TOB.Create('LISTEINVENT', TOBSelection, -1).SelectDB('', Q, true);
{$ENDIF}
      end;
    end;
{$IFDEF EAGLCLIENT}
{$ELSE}
  Q.EnableControls;
{$ENDIF}
  end;
end;


procedure TOF_ValidInv.OnArgument(stArgument : String);
var Ctl : TControl ;
begin
inherited;
SetControlText('XX_WHERE', 'GIE_DATEINVENTAIRE > "'+USDateTime(GetParamSoc('SO_GCDATECLOTURESTOCK'))+'" AND GIE_VALIDATION<>"X"');

// Libellé Dépôt ou Etablissement
if not VH_GC.GCMultiDepots then
   begin
    Ctl := GetControl('TGIE_DEPOT') ;
    if (Ctl <> Nil) and (Ctl is THLabel) then THLabel(Ctl).Caption := 'Etablissement' ;
   end ;
end;

procedure TOF_ValidInv.OnValidate;
begin
if not PutSelectionIntoTOB then exit;

// Test si le Tiers des Ecarts d'inventaire a été défini dans les paramètres généraux
if GetParamsoc('SO_GCTIERSINV')='' then
    begin
    PGIInfo(TraduireMemoire(TexteMessage[1]),Ecran.Caption);
    TOBSelection.Free;
    exit;
    end;
if EntreeAssistValidInv(TOBSelection) then TFMul(Ecran).ChercheClick;
TOBSelection.Free; // Create dans PutSelectionIntoTOB
end;


initialization
RegisterClasses([TOF_ValidInv]);
RegisterAGLProc('ValidateList', true, 0, AGLOnValidateList);

end.
