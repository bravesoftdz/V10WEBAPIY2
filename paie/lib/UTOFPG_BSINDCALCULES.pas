unit UTOFPG_BSINDCALCULES;

interface
uses
Classes,
SysUtils,
Controls,
Utof,
Htb97,
Vierge,
{$IFDEF EAGLCLIENT}
eMul,MainEagl,
{$ELSE}
Mul,Fe_main,
{$ENDIF}
HmsgBox,
HCtrls;

type
  TOF_MUL_BSINDCALCULES = class(TOF)
    procedure OnArgument(S: string); override;
  private
   procedure BtTraitementClick ( sender : TObject);
  End;

  TOF_BSINDICATION = class(TOF)
    procedure OnUpdate; override;
    procedure OnCancel; override;
  End;

implementation


{ TOF_MUL_BSINDCALCULES }

procedure TOF_MUL_BSINDCALCULES.OnArgument(S: string);
Var Btn : TToolBarButton97;
begin
  inherited;
Btn := TToolBarButton97(GetControl('BSUPPRIMER'));
if Assigned(Btn) then Btn.onclick := BtTraitementClick;

Btn := TToolBarButton97(GetControl('BTINDICATION'));
if Assigned(Btn) then Btn.onclick := BtTraitementClick;
end;


procedure TOF_MUL_BSINDCALCULES.BtTraitementClick(sender: TObject);
Var
   i, Increm, Tot                                       : integer;
   Indic, BsPresent, st, Indication,MajIndication       : string ;
   DateDebut, DateFin                                   : TDateTime;
   ToutSelect                                           : Boolean;
begin
 if TToolBarButton97(Sender).Name = 'BSUPPRIMER' Then
   St := 'supprimer'
 else
 if TToolBarButton97(Sender).Name = 'BTINDICATION' then
      St := 'modifier';

 { Gestion de la sélection des indicateurs }
  if (TFMul(Ecran).FListe.nbSelected > 0) or (TFMul(Ecran).FListe.AllSelected) then
    if PgiAsk('Etês-vous sûr de vouloir '+St+' les indicateurs sélectionnés?', Ecran.caption) = mrYes then
    begin
      St := ''; ToutSelect := False;
      If TFMul(Ecran).FListe.AllSelected then
        Begin
        ToutSelect := True;
        Tot := TFMul(Ecran).Q.RecordCount;
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Fetchlestous;
        {$ENDIF}
        TFMul(Ecran).Q.First;
        End
      else Tot := TFMul(Ecran).FListe.NbSelected;

      if TToolBarButton97(Sender).Name = 'BTINDICATION' then
        Begin
        MajIndication := AglLanceFiche('PAY','BSINDICATION','','','');
        if MajIndication = '' then exit;
        End;


      { Balayage des lignes sélectionnées }
      for i := 0 to Tot - 1 do
      begin
        {$IFDEF EAGLCLIENT}
        if not ToutSelect then  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
        {$ENDIF}
        if TFMul(Ecran).FListe.NbSelected > 0 then TFMul(Ecran).FListe.GotoLeBookmark(i);

        Indic      := TFmul(Ecran).Q.FindField('PBC_INDICATEURBS').asstring;
        DateDebut  := TFmul(Ecran).Q.FindField('PBC_DATEDEBUT').AsDateTime;
        DateFin    := TFmul(Ecran).Q.FindField('PBC_DATEFIN').AsDateTime;
        BsPresent  := TFmul(Ecran).Q.FindField('PBC_BSPRESENTATION').asstring;
        Increm     := TFmul(Ecran).Q.FindField('PBC_INCREM').AsInteger;
        Indication := TFmul(Ecran).Q.FindField('PBC_PGINDICATION').AsString;

        if ((Indication = 'VER') AND (TToolBarButton97(Sender).Name = 'BSUPPRIMER'))
        OR ((TToolBarButton97(Sender).Name = 'BTINDICATION') AND ((Indication = 'PER' ) OR (Indication = MajIndication)) ) then
          Begin
          if ToutSelect then TFMul(Ecran).Q.Next;
          continue;
          End;

        if TToolBarButton97(Sender).Name = 'BSUPPRIMER' Then
              St := 'DELETE FROM BILANSOCIAL '+
                    'WHERE PBC_INDICATEURBS = "'+Indic+'" '+
                    'AND PBC_DATEDEBUT ="'+USDateTime(DateDebut)+'" '+
                    'AND PBC_DATEFIN ="'+USDateTime(DateFin)+'" '+
                    'AND PBC_BSPRESENTATION ="'+BsPresent+'" '+
                    'AND PBC_INCREM ='+IntToStr(Increm)
        else
        if TToolBarButton97(Sender).Name = 'BTINDICATION' then
              St := 'UPDATE BILANSOCIAL SET PBC_PGINDICATION="'+MajIndication+'" '+
                    'WHERE PBC_INDICATEURBS = "'+Indic+'" '+
                    'AND PBC_DATEDEBUT ="'+USDateTime(DateDebut)+'" '+
                    'AND PBC_DATEFIN ="'+USDateTime(DateFin)+'" '+
                    'AND PBC_BSPRESENTATION ="'+BsPresent+'" '+
                    'AND PBC_INCREM ='+IntToStr(Increm) ;
        ExecuteSql(St);
        if ToutSelect then TFMul(Ecran).Q.Next;
      end;
    TFMul(Ecran).FListe.ClearSelected;
    TFMul(Ecran).BCherche.Click;
    End;
end;
        (*if TToolBarButton97(Sender).Name = 'BTINDICATION' then
          Begin
          if Indication = 'CAL' then Indication := 'NON'
          else if Indication = 'NON' then Indication := 'CAL'
          else Begin
             if TFMul(Ecran).FListe.AllSelected then TFMul(Ecran).Q.Next;
             continue;
             End;
          End;   *)

{ TOF_BSINDICATION }

procedure TOF_BSINDICATION.OnCancel;
begin
  inherited;
TFVierge(Ecran).Retour := '';
end;

procedure TOF_BSINDICATION.OnUpdate;
begin
  inherited;
TFVierge(Ecran).Retour := GetControlText('PGINDICATION');
end;

initialization
  registerclasses([TOF_MUL_BSINDCALCULES,TOF_BSINDICATION]);
end.
