{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 5/05/2000
Modifié le ... : 01/08/2000
Description .. : TOF Pour l'écran de lancement d'état d'inventaire
Suite ........ : permanent
Mots clefs ... : INVENTAIRE;PERMANENT;STOCK
*****************************************************************}
unit UTOFInvPerm;

interface

uses stdctrls, Classes, UTOF, HCtrls;

type
    TOF_InvPerm = class(TOF)
    private
      FSelectCombo,
      FCBJour,
      FCBMois,
      FCBAnnee : THValComboBox;
      FValorizChk, FDtlChk : TCheckBox;

      procedure RefreshYearz;
      procedure RefreshMonths;
      procedure RefreshDayz;

    published
      procedure OnNew; override;
      procedure OnLoad; override;
      procedure OnComboChange(Sender : TObject);
      procedure OnComboPopUp(Sender : TObject);
      procedure OnValorizClick(Sender : TObject);
      procedure OnDtlClick(Sender : TObject);
      procedure OnYearChange(Sender : TObject);
      procedure OnMonthChange(Sender : TObject);
    end;

const
    NbRupts = 5;

implementation

uses Controls, SysUtils, HEnt1, Graphics,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     UTOB, HTB97;

procedure TOF_InvPerm.RefreshYearz;
var Q : TQuery;
    Y : String;
begin
FCBAnnee.Clear;
Y := FormatDateTime('yyyy', Date);
FCBAnnee.Items.AddObject(Y, Pointer(StrToInt(Y)));

Q := OpenSQL('SELECT DISTINCT GQ_DATECLOTURE FROM DISPO WHERE GQ_CLOTURE="X" '+
             'ORDER BY GQ_DATECLOTURE DESC', true);
while not Q.EOF do
  begin
  Y := FormatDateTime('yyyy', Q.FindField('GQ_DATECLOTURE').AsDateTime);
  if FCBAnnee.Items.IndexOf(Y) = -1
   then FCBAnnee.Items.AddObject(Y, Pointer(StrToInt(Y)));

  Q.Next;
  end;
Ferme(Q);

FCBAnnee.ItemIndex := 0;
end;

procedure TOF_InvPerm.RefreshMonths;
var Q : TQuery;
    D, Inf, Sup : TDate;
begin
FCBMois.Clear;
if (FCBAnnee.Text = FormatDateTime('yyyy', Date))
 then FCBMois.Items.AddObject(FormatDateTime('mmmm', Date), Pointer(StrToInt(FormatDateTime('m', Date))));

Inf := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]), 1, 1);
Sup := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]), 12, 31);
Q := OpenSQL('SELECT DISTINCT GQ_DATECLOTURE FROM DISPO WHERE GQ_CLOTURE="X" '+
             'AND GQ_DATECLOTURE>="'+USDateTime(Inf)+'" AND GQ_DATECLOTURE<="'+USDateTime(Sup)+'" '+
             'ORDER BY GQ_DATECLOTURE DESC', true);
while not Q.EOF do
  begin
  D := Q.FindField('GQ_DATECLOTURE').AsDateTime;
  if (FCBMois.Items.IndexOf(FormatDateTime('mmmm', D)) = -1)
   then FCBMois.Items.AddObject(FormatDateTime('mmmm', D), Pointer(StrToInt(FormatDateTime('m', D))));

  Q.Next;
  end;
Ferme(Q);

FCBMois.ItemIndex := 0;
end;

procedure TOF_InvPerm.RefreshDayz;
var Q : TQuery;
    D, Inf, Sup : TDate;
begin
FCBJour.Clear;
if (FCBAnnee.Text = FormatDateTime('yyyy', Date)) and (FCBMois.Text = FormatDateTime('mmmm', Date))
 then FCBJour.Items.AddObject(FormatDateTime('dd', Date), Pointer(StrToInt(FormatDateTime('d', Date))));

Inf := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]),
                  Integer(FCBMois.Items.Objects[FCBMois.ItemIndex]), 1);
Sup := FinDeMois(Inf);
Q := OpenSQL('SELECT DISTINCT GQ_DATECLOTURE FROM DISPO WHERE GQ_CLOTURE="X" '+
             'AND GQ_DATECLOTURE>="'+USDateTime(Inf)+'" AND GQ_DATECLOTURE<="'+USDateTime(Sup)+'" '+
             'ORDER BY GQ_DATECLOTURE DESC', true);
while not Q.EOF do
  begin
  D := Q.FindField('GQ_DATECLOTURE').AsDateTime;
  if (FCBJour.Items.IndexOf(FormatDateTime('dd', D)) = -1)
   then FCBJour.Items.AddObject(FormatDateTime('dd', D), Pointer(StrToInt(FormatDateTime('d', D))));

  Q.Next;
  end;
Ferme(Q);

FCBJour.ItemIndex := 0;
end;

procedure TOF_InvPerm.OnYearChange(Sender : TObject);
begin
RefreshMonths;
RefreshDayz;
end;

procedure TOF_InvPerm.OnMonthChange(Sender : TObject);
begin
RefreshDayz;
end;


procedure TOF_InvPerm.OnNew;
var i : integer;
begin
inherited;
FSelectCombo := THValComboBox(Ecran.FindComponent('KELPRIX'));
FDtlChk := TCheckBox(Ecran.FindComponent('CHKDTL'));
FValorizChk := TCheckBox(Ecran.FindComponent('CHKVALOR'));
FCBJour := THValComboBox(Ecran.FindComponent('CBJOUR'));
FCBMois := THValComboBox(Ecran.FindComponent('CBMOIS'));
FCBAnnee := THValComboBox(Ecran.FindComponent('CBANNEE'));

FSelectCombo.ItemIndex := 0;
FDtlChk.OnClick := OnDtlClick;
FValorizChk.OnClick := OnValorizClick;
OnDtlClick(self);
OnValorizClick(self);
FCBAnnee.OnChange := OnYearChange;
FCBMois.OnChange := OnMonthChange;

RefreshYearz;
RefreshMonths;
RefreshDayz;

for i := 1 to NbRupts do
  begin
  THValComboBox(Ecran.FindComponent('RUPT'+inttostr(i))).OnChange := OnComboChange;
  THValComboBox(Ecran.FindComponent('RUPT'+inttostr(i))).OnDropDown := OnComboPopup;
  end;

for i := 1 to 3 do
  THLabel(Ecran.FindComponent('TGA_FAMILLENIV'+inttostr(i))).Caption := RechDom('GCLIBFAMILLE', 'LF'+inttostr(i), false);

SetControlVisible ('BFiltre', False);
SetControlVisible ('FFiltres', False);
end;

procedure TOF_InvPerm.OnLoad;
var StXXWhere : String;
    DateInv : TDate;
    Q : TQuery;
begin
inherited;
DateInv := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]),
                      Integer(FCBMois.Items.Objects[FCBMois.ItemIndex]),
                      Integer(FCBJour.Items.Objects[FCBJour.ItemIndex]));
if DateInv = Date
  then StXXWhere := 'GQ_CLOTURE<>"X"'
  else StXXWhere := 'GQ_CLOTURE="X" AND GQ_DATECLOTURE="'+USDateTime(DateInv)+'"';
THCritMaskEdit(Ecran.FindComponent('XX_WHERE')).Text := StXXWhere;
THCritMaskEdit(Ecran.FindComponent('DATE')).Text := DateToStr(DateInv);

if FValorizChk.Checked then
  begin
  THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE10')).Text := RechDom('GCINFPRIXIP', FSelectCombo.Value, true);  // Champ
  Q := OpenSQL('SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE="GPI" AND CO_CODE="'+FSelectCombo.Value+'"', true);
  if not Q.EOF then THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE11')).Text := Q.FindField('CO_LIBRE').AsString; // Lib.
  Ferme(Q);
  end else
  begin
  THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE10')).Text := '';
  THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE11')).Text := '';
  end;

  TToolbarButton97(Ecran.FindComponent('bExport')).Visible := True;

//if not TCheckBox(Ecran.FindComponent('CHKDTL')).Checked then TCheckBox(Ecran.FindComponent('CHKQZERO')).Checked := false;
//if not TCheckBox(Ecran.FindComponent('CHKVALOR')).Checked then TCheckBox(Ecran.FindComponent('CHKSTOCKNEG')).Checked := false;
end;

procedure TOF_InvPerm.OnComboChange(Sender : TObject);
var RuptNb : String;
    FCombo : THValComboBox;
    NextComboName : String;
begin
FCombo := (Sender as THValComboBox);
RuptNb := Copy(FCombo.Name,5,1);

if FCombo.Value = '' then
  begin
  THCritMaskEdit(Ecran.FindComponent('XX_RUPTURE'+RuptNb)).Text := '';
  THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE'+RuptNb)).Text := '';
  end else
  begin
  THCritMaskEdit(Ecran.FindComponent('XX_RUPTURE'+RuptNb)).Text := RechDom('GCGROUPINVPERM', FCombo.Value, true);
  if (FCombo.Value = 'LF1') or (FCombo.Value = 'LF2') or (FCombo.Value = 'LF3')
    then THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE'+RuptNb)).Text := RechDom('GCLIBFAMILLE', FCombo.Value, false)
    else THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE'+RuptNb)).Text := RechDom('GCGROUPINVPERM', FCombo.Value, false);
  end;

NextComboName := FCombo.Name; NextComboName[5] := Chr(Ord(NextComboName[5])+1);
if (Ecran.FindComponent(NextComboName) <> nil) and (FCombo.Value <> '') then
  begin
  THValComboBox(Ecran.FindComponent(NextComboName)).Enabled := true;
  THValComboBox(Ecran.FindComponent(NextComboName)).Color := clWindow;
  end else
  while Ecran.FindComponent(NextComboName) <> nil do
    begin
    with THValComboBox(Ecran.FindComponent(NextComboName)) do
      begin Enabled := False; Color := clBtnFace; Value := ''; end;
    NextComboName[5] := Chr(Ord(NextComboName[5])+1);
    end;
end;

procedure TOF_InvPerm.OnComboPopUp(Sender : TObject);
var St_Plus, St_Value, St_Text, St : String;
    FCombo : THValComboBox;
    Indice, i_ind : Integer;
begin
FCombo := (Sender as THValComboBox);
Indice := StrToInt(Copy(FCombo.Name,5,1));
St_Plus := '';
St_Value := FCombo.Value;
St_Text := FCombo.Text;
for i_ind := 1 to NbRupts do
  begin
  if i_ind = Indice then Continue;
  St := String(THValComboBox(Ecran.FindComponent('RUPT'+IntToStr(i_ind))).Value);
  if St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"'+St+'"';
  end;
FCombo.Plus := St_Plus;
FCombo.Value := St_Value;
FCombo.Text := St_Text;
end;

procedure TOF_InvPerm.OnValorizClick(Sender : TObject);
begin
THValComboBox(Ecran.FindComponent('KELPRIX')).Enabled := FValorizChk.Checked;
with TCheckBox(Ecran.FindComponent('CHKSTOCKNEG')) do
  begin
  if not FValorizChk.Checked then Checked := false;
  Enabled := FValorizChk.Checked;
  end;
end;

procedure TOF_InvPerm.OnDtlClick(Sender : TObject);
begin
with TCheckBox(Ecran.FindComponent('CHKQZERO')) do
  begin
  if not FDtlChk.Checked then Checked := false;
  Enabled := FDtlChk.Checked;
  end;
end;

initialization
RegisterClasses([TOF_InvPerm]);

end.
