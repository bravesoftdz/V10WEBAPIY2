{ Unité : Source TOF de la FICHE : CPCRITERECFONB
--------------------------------------------------------------------------------------
    Version   |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
07.01.001.001   05/07/06   JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit CPCRITERECFONB_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  Vierge, SysUtils, HCtrls, UTOF, Forms, ComCtrls;

type
  TOF_CPCRITERECFONB = class (TOF)
  private
    TypeExport : string;
    PeutSortir : Boolean;

    procedure MajSender;
    procedure InitControls;
    function  ValideSaisie : Boolean;
  public
    PageControl : TPageControl;
    FCloseQuery : TCloseQueryEvent;

    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
    procedure BFermeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure CloseQuery(Sender: TObject; var CanClose: Boolean);
  end;

function LanceCritereCfonb(Arg : string) : string;


implementation

uses
  uLibCFONB, HMsgBox, HTB97;

{---------------------------------------------------------------------------------------}
function LanceCritereCfonb(Arg : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AGLLanceFiche('CP', 'CPCRITERECFONB', '', '', Arg);
  if Result = '' then Result := CONST_ANNULER;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCRITERECFONB.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  PageControl := TPageControl(GetControl('PCCRITERE'));
  TypeExport := s;

  FCloseQuery := Ecran.OnCloseQuery;
  Ecran.OnCloseQuery := CloseQuery;
  TToolBarButton97(GetControl('BFERME')).OnMouseDown := BFermeMouseDown;
  PeutSortir := True;
  
       if TypeExport = cfonb_Transfert   then PageControl.ActivePage := TTabSheet(GetControl('TSTRANSFERT'))
  else if TypeExport = cfonb_Prelevement then PageControl.ActivePage := TTabSheet(GetControl('TSPRELEVEMENT'))
  else if TypeExport = cfonb_Virement    then PageControl.ActivePage := TTabSheet(GetControl('TSVIREMENT'));

  {Initialisation des contrôles}
  InitControls;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCRITERECFONB.OnUpdate ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if not ValideSaisie then begin
    PeutSortir := False;
    Exit;
  end
  else
    PeutSortir := True;

  {Récupération de la saisie}
  MajSender;
  TFVierge(Ecran).Retour := 'OK';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCRITERECFONB.MajSender;
{---------------------------------------------------------------------------------------}
begin
  if TypeExport = cfonb_Transfert   then begin
    LaTob.SetString('IMPUTATION', GetControlText('CBIMPUTATIONFRAIS'));
    LaTob.SetString('REMISE'    , GetControlText('CBDEBITREMISE'));
    LaTob.SetString('NATECO'    , GetControlText('CBNATECO'));
    LaTob.SetString('REFERENCE' , GetControlText('EDREFREMISE'));
    LaTob.SetString('DATEREMISE', GetControlText('EDDATEREMISE'));
    if GetCheckBoxState('CKRIBPRINCIPAL') = cbChecked then LaTob.SetString('RIBPRINC', 'X')
                                                      else LaTob.SetString('RIBPRINC', '-');
    if GetCheckBoxState('CKCUMULE') = cbChecked then LaTob.SetString('CUMULE', 'X')
                                                else LaTob.SetString('CUMULE', '-');
  end
  else if TypeExport = cfonb_Prelevement then begin

  end
  else if TypeExport = cfonb_Virement then begin

  end
  {else if LCR then
    if ReftireLib.Checked then LaTob.SetString('REFTIRELIB', 'X')
                          else LaTob.SetString('REFTIRELIB', '-')
   }
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCRITERECFONB.ValideSaisie : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  if TypeExport = cfonb_Transfert   then begin
    if GetControlText('EDREFREMISE') = '' then begin
      HShowMessage('0;' + Ecran.Caption + ';Vous devez indiquer une référence remise.;W;O;O;O;', '', '');
      Result := False;
    end
  end
  else if TypeExport = cfonb_Prelevement then begin

  end
  else if TypeExport = cfonb_Virement then begin

  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCRITERECFONB.CloseQuery(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := PeutSortir;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCRITERECFONB.BFermeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
begin
  PeutSortir := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCRITERECFONB.InitControls;
{---------------------------------------------------------------------------------------}
begin
  if TypeExport = cfonb_Transfert then begin
    THValComboBox(GetControl('CBIMPUTATIONFRAIS')).ItemIndex := 1;
    THValComboBox(GetControl('CBDEBITREMISE'    )).ItemIndex := 1;
  end
  else if TypeExport = cfonb_Prelevement then begin

  end
  else if TypeExport = cfonb_Virement then begin

  end;

end;

initialization
  RegisterClasses([TOF_CPCRITERECFONB]);

end.
