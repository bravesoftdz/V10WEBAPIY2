{ Unité : Source TOF de la FICHE : TRMULTAUXICC
--------------------------------------------------------------------------------------
    Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.05.001.001   17/10/06  JP   Création de l'unité : Mul sur la table ICCTAUXCOMPTE
--------------------------------------------------------------------------------------}
unit TRMULTAUXICC_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     UTOB, MaineAGL, eMUL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main, Mul,
{$ENDIF}
     SysUtils, ComCtrls, uTof;

type
  TOF_TRMULTAUXICC = class (TOF)
    procedure OnArgument(s : string); override;
    procedure OnLoad                ; override;
  private
    procedure BInsertOnClick(Sender : TObject);
    procedure BDeleteOnClick(Sender : TObject);
    procedure FListeDblClick(Sender : TObject);
  end;

procedure TRLance_FicheTauxIcc(Arg : string);

implementation

uses
  HCtrls, HTB97, HEnt1, HMsgBox, ICCTAUXCOMPTE_TOM, Commun, Constantes;

{---------------------------------------------------------------------------------------}
procedure TRLance_FicheTauxIcc(Arg : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('TR', 'TRMULTAUXICC', '', '', Arg);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULTAUXICC.OnArgument(s : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000154;
  if S <> '' then SetControlText('ICD_GENERAL', ReadTokenSt(S));
  TFMul(Ecran).Binsert.OnClick := BInsertOnClick;
  TFMul(Ecran).FListe.OnDblClick := FListeDblClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BDeleteOnClick;
  
  {On limite l'affichage aux comptes courants des dossiers du regroupement Tréso}
  THEdit(GetControl('ICD_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULTAUXICC.OnLoad;
{---------------------------------------------------------------------------------------}
var
  s  : string;
begin
  inherited;
  s := '(ICD_DATEDU <= "' + UsDateTime(StrToDateTime(GetControlText('EDDATE'))) + '" AND ' +
       ' ICD_DATEAU >= "' + UsDateTime(StrToDateTime(GetControlText('EDDATE'))) + '") OR ' +
       '(ICD_DATEDU <= "' + UsDateTime(StrToDateTime(GetControlText('EDDATE_'))) + '" AND ' +
       ' ICD_DATEAU >= "' + UsDateTime(StrToDateTime(GetControlText('EDDATE_'))) + '") OR ' +
       '(ICD_DATEDU >= "' + UsDateTime(StrToDateTime(GetControlText('EDDATE'))) + '" AND ' +
       ' ICD_DATEAU <= "' + UsDateTime(StrToDateTime(GetControlText('EDDATE_'))) + '") ';
  SetControlText('XX_WHERE', s);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULTAUXICC.BDeleteOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if TFMul(Ecran).FListe.nbSelected = 0 then begin
    PGIError('Veuillez sélectionner un élément.', Ecran.Caption);
    Exit;
  end;

  for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
    TFMul(Ecran).FListe.GotoLeBookmark(n);
    ExecuteSQL('DELETE FROM ICCTAUXCOMPTE WHERE ICD_GENERAL = "' + VarToStr(GetField('ICD_GENERAL')) +
               '" AND ICD_DATEDU = "' + UsDateTime(VarToDateTime(GetField('ICD_DATEDU'))) + '"');
  end;
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULTAUXICC.BInsertOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_TRTauxIcc('', '', 'ACTION=CREATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULTAUXICC.FListeDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := VarToStr(GetField('ICD_GENERAL')) + ';' + VarToStr(GetField('ICD_DATEDU')) + ';';
  TRLanceFiche_TRTauxIcc('', s, 'ACTION=MODIFICATION');
  TFMul(Ecran).BCherche.Click;
end;

initialization
  RegisterClasses([TOF_TRMULTAUXICC]);

end.
