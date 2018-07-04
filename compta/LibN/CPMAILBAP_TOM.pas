{ Unité : Source TOM de la TABLE : CPMAILBAP
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001   06/02/06    JP     Création de l'unité
 8.01.001.003   24/01/07    JP     FQ 19561 : Suppression du popup de formatage du texte du memo
--------------------------------------------------------------------------------------}
unit CPMAILBAP_TOM;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFichList,
  {$ELSE}
  FE_Main, FichList, db, dbtables,
  {$ENDIF}
  SysUtils, UTOM, UTob;

type
  TOM_CPMAILBAP = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnDeleteRecord           ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnLoadRecord             ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnArgument   (S : string); override;
    procedure OnClose                  ; override;
    procedure OnCancelRecord           ; override;
  private

  end;

procedure CpLanceFiche_MailsBAP(Range, Lequel, Argument : string);


implementation

uses
  HCtrls, HEnt1, HMsgBox, HTB97, Ent1, HRichOle;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_MailsBAP(Range, Lequel, Argument : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPMAILBAP', Range, Lequel, Argument);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509030;
  {24/01/07 : FQ 19561 : Suppression du popup de formatage du texte du memo}
  {$IFDEF EAGLCLIENT}
  (GetControl('CMA_BLOCNOTE') as THRichEditOle).PlainText := False;
  (GetControl('CMA_BLOCNOTE') as THRichEditOle).DeletePopUp;
  {$ELSE}
  (GetControl('CMA_BLOCNOTE') as THDBRichEditOle).PlainText := True;
  (GetControl('CMA_BLOCNOTE') as THDBRichEditOle).DeletePopUp;
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnCancelRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPMAILBAP.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

initialization
  RegisterClasses([TOM_CPMAILBAP]);

end.
