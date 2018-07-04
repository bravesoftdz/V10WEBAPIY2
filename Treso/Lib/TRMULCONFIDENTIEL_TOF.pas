{-------------------------------------------------------------------------------------
   Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
08.10.001.001  02/08/07  JP   Création de l'unité : Mul des confidentialités
08.10.001.013  09/10/07  JP   FQ 10529 : gestion de l'alternance User / UserGrp
--------------------------------------------------------------------------------------}
unit TRMULCONFIDENTIEL_TOF;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, UTOF;

type
  TOF_TRMULCONFIDENTIEL = class (TOF)
    procedure OnArgument(S : string); override;
  private
    procedure SlctAllClick   (Sender : TObject);
    procedure BInsertClick   (Sender : TObject);
    procedure BDeleteClick   (Sender : TObject);
    procedure ListDblClick   (Sender : TObject);
  end ;

procedure TRLanceFiche_MulConfidentialite(Range, Lequel, Arguments : string);


implementation

uses
  HMsgBox, PROSPECTCONF_TOM, HTB97, AglInit, HQry, ULibConfidentialite;


{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulConfidentialite(Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('TR', 'TRMULCONFIDENTIEL', Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONFIDENTIEL.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000157;
  TToolbarButton97(GetControl('BINSERT'   )).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BDELETE'   )).OnClick := BDeleteClick;
  TToolbarButton97(GetControl('BOUVRIR'   )).OnClick := ListDblClick;
  TFMul(Ecran).FListe.OnDblClick := ListDblClick;
  TFMul(Ecran).SetDBListe(GetLaDBListe);
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;
  THValComboBox(GetControl('RTC_PRODUITPGI')).Plus := 'AND CO_CODE <> "CPT"';

  {09/10/07 : FQ 10529 : gestion du libellé et de la tablettre en fonction du paramsoc}
  GereTabletteUser(GetControl('RTC_INTERVENANT'));
  GereTabletteUser(GetControl('RTC_INTERVENANT_'));
  GereCaptionUser(THLabel(GetControl('TRTC_INTERVENANT')));
  {09/10/07 : FQ 10529 : On cache ces zones car pour le moment, elles n'ont pas d'intérêt}
  SetControlVisible('RTC_TYPECONF' , False);
  SetControlVisible('TRTC_TYPECONF', False);
  SetControlVisible('RTC_PRODUITPGI' , False);
  SetControlVisible('TRTC_PRODUITPGI', False);
  GetControl('RTC_INTERVENANT'  ).Top := GetControl('RTC_TYPECONF').Top;
  GetControl('RTC_INTERVENANT_' ).Top := GetControl('RTC_TYPECONF').Top;
  GetControl('TRTC_INTERVENANT' ).Top := GetControl('TRTC_TYPECONF').Top;
  GetControl('TRTC_INTERVENANT_').Top := GetControl('TRTC_TYPECONF').Top;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONFIDENTIEL.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  w : string;
begin
  if TFMul(Ecran).FListe.AllSelected then begin
    if PgiAsk(TraduireMemoire('Êtes-vous sûr de vouloir supprimer toutes les confidentialités sélectionnées'), Ecran.Caption) = mrYes then begin
      w := RecupWhereCritere(TFMul(Ecran).Pages);
      ExecuteSQL('DELETRE FROM PROSPECTCONF ' + w);

      TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
    end;
  end
  else begin
    if TFMul(Ecran).FListe.nbSelected = 0 then begin
      HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '');
      Exit;
    end;

    if PgiAsk(TraduireMemoire('Êtes-vous sûr de vouloir supprimer le(s) confidentialité(s) sélectionnée(s)'), Ecran.Caption) = mrYes then begin
      {On boucle sur la sélection}
      for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
        TFMul(Ecran).FListe.GotoLeBookmark(n);
        BeginTrans;
        try
          {Suppression du contrat courant ...}
          ExecuteSQL('DELETE FROM PROSPECTCONF WHERE RTC_PRODUITPGI = "'  + GetField('RTC_PRODUITPGI') + '" AND ' +
                                                    'RTC_INTERVENANT = "' + GetField('RTC_INTERVENANT') + '" AND ' +
                                                    'RTC_TYPECONF = "'    + GetField('RTC_TYPECONF') + '"');
          CommitTrans;
        except
          on E : Exception do begin
            RollBack;
            PGIError(E.Message);
          end;
        end;
      end;
      TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONFIDENTIEL.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Confidentialite('', '', ActionToString(taCreatEnSerie));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONFIDENTIEL.ListDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  lClef : string;
begin
  lClef := GetField('RTC_INTERVENANT') + ';' + GetField('RTC_TYPECONF') + ';' + GetField('RTC_PRODUITPGI') + ';';
  TRLanceFiche_Confidentialite('', lClef, ActionToString(taModif));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONFIDENTIEL.SlctAllClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(Fiche.bSelectAll);
end;

initialization
  RegisterClasses([TOF_TRMULCONFIDENTIEL]);

end.
