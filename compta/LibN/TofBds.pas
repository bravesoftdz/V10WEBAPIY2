unit TofBds;

interface
uses Classes, StdCtrls, UTof, HCtrls, QRS1, Ent1, SysUtils, dbtables, TofMeth, Spin;

// TOF.OnArgument sur show
// TOF.Load après le valideclick

type
  TOF_BDS = class(TOF_Meth)
  private
    Cpt1, Cpt2: THEdit;
    Exo1, Exo2, ExoBds, Etab, Devise : THValComboBox ;
    Date1, Date2, Date3, Date4 : THEdit;
    DateD, DateF: TDatetime;
    DevPivot: THEdit;
    procedure CompteOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure ExoBdsOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnUpdate ; override ;
    procedure OnNew ; override ;
  end ;

  TOF_BALANCE = class(TOF_Meth)
  private
    Cpt1, Cpt2: THEdit;
    Exo1, Exo2, ExoBds, Etab, Devise : THValComboBox ;
    Date1, Date2, Date3, Date4 : THEdit;
    DateD, DateF: TDatetime;
    DevPivot: THEdit;
    Comparatif, Blocnote: TCheckbox;
    Rupture: THEdit;
    Ruptures: TSpinEdit;
    Histo: THLabel;
    TypeBal: THEdit;
    procedure CompteOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure ExoBdsOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure ComparatifOnClick(Sender: TObject) ;
    procedure InitCritereAvance;
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnUpdate ; override ;
    procedure OnNew ; override ;
  end ;


implementation

uses HEnt1, ZBalance, HMsgBox;

procedure TOF_BDS.CompteOnExit(Sender: TObject) ;
BEGIN
DoCompteOnExit(THEdit(Sender), Cpt1, Cpt2);
END;

procedure TOF_BDS.ExoOnChange(Sender: TObject) ;
BEGIN
if Sender = Exo1 then begin
  DoExoToDateOnChange(Exo1, Date1, Date2);
  DateD := StrToDate(Date1.Text);
  DateF := StrToDate(Date2.Text);
end;
END;

procedure TOF_BDS.ExoBdsOnChange(Sender: TObject);
var s: string;
BEGIN
s:=ExoBds.Value;
Exo2.ItemIndex:=ExoBds.ItemIndex;
ReadTokenST(s);
Date3.Text:=ReadTokenST(s);
Date4.Text:=ReadTokenST(s);
END;

procedure TOF_BDS.DateOnExit(Sender: TObject) ;
BEGIN
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
END;

procedure TOF_BDS.OnArgument(Arguments : string) ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
  Caption := TraduireMemoire('Balance générale comparative');
  UpdateCaption(Ecran);
  Exo1:=THValComboBox(GetControl('EXO1')) ;
  Exo2:=THValComboBox(GetControl('EXO2')) ;
  ExoBds:=THValComboBox(GetControl('EXOBDS')) ;
  Cpt1:=THEdit(GetControl('G_GENERAL')) ;
  Cpt2:=THEdit(GetControl('G_GENERAL_')) ;
  Date1:=THEdit(GetControl('DATE1')) ;
  Date2:=THEdit(GetControl('DATE2')) ;
  Date3:=THEdit(GetControl('DATE3')) ;
  Date4:=THEdit(GetControl('DATE4')) ;
  Etab:=THValComboBox(GetControl('ETABLISSEMENT')) ;
  Devise:=THValComboBox(GetControl('DEVISE')) ;
  DevPivot:=THEdit(GetControl('DEVPIVOT'));
end;
  if Exo1<>nil then begin
    Exo1.OnChange:=ExoOnChange ;
    Exo1.Value:=VH^.Entree.Code ;
  end ;
  if ExoBds<>nil then begin
    ExoBds.OnChange:=ExoBdsOnChange ;
    FillComboBds(ExoBds) ;
    if Exo2<>nil then FillComboValues(ExoBds, Exo2);
    ExoBdsOnChange(ExoBds);
  end ;
  if Cpt1<>nil then Cpt1.OnExit:=CompteOnExit ;
  if Cpt2<>nil then Cpt2.OnExit:=CompteOnExit ;
  if Date1<>nil then Date1.OnExit:=DateOnExit ;
  if Date2<>nil then Date2.OnExit:=DateOnExit ;
  if (Etab<>nil) and (Etab.ItemIndex<>-1) then Etab.ItemIndex := 0 ;
  if Devise<>nil then Devise.ItemIndex := 0;
  if (Devise<>nil) and (DevPivot<>nil) then
    DevPivot.Text:=Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
  TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];
END ;

procedure TOF_BDS.OnUpdate ;
BEGIN
inherited ;
if Trim(Cpt1.Text) = '' then Cpt1.Text := '00000000';
if Trim(Cpt2.Text) = '' then Cpt2.Text := 'ZZZZZZZZ';
END;

procedure TOF_BDS.OnNew;
begin
if ExoBds.ItemIndex<0 then begin
  HShowMessage('0;Il n''y a pas de balance comparative;Aucun enregistrement à imprimer;E;O;O;O;','','') ;
  SetControlEnabled(TFQRS1(Ecran).BValider.Name, false) ;
  Abort;
end;
end;


{TOF_BALANCE}

procedure TOF_BALANCE.OnArgument(Arguments : string) ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
  PageOrder(Pages);
  Caption := TraduireMemoire('Balance générale');
  UpdateCaption(Ecran);
  Exo1:=THValComboBox(GetControl('EXO1')) ;
  Exo2:=THValComboBox(GetControl('EXO2')) ;
  ExoBds:=THValComboBox(GetControl('EXOBDS')) ;
  Cpt1:=THEdit(GetControl('G_GENERAL')) ;
  Cpt2:=THEdit(GetControl('G_GENERAL_')) ;
  Date1:=THEdit(GetControl('DATE1')) ;
  Date2:=THEdit(GetControl('DATE2')) ;
  Date3:=THEdit(GetControl('DATE3')) ;
  Date4:=THEdit(GetControl('DATE4')) ;
  Etab:=THValComboBox(GetControl('ETABLISSEMENT')) ;
  Devise:=THValComboBox(GetControl('DEVISE')) ;
  DevPivot:=THEdit(GetControl('DEVPIVOT'));
  Comparatif:=TCheckbox(GetControl('Comparatif'));
  Blocnote:=TCheckbox(GetControl('Blocnote'));
  Ruptures:=TSpinEdit(GetControl('RUPTURES'));
  Rupture:=THEdit(GetControl('Rupture'));
  TypeBal:=THEdit(GetControl('TypeBal'));
  Histo:=THLabel(GetControl('Historique'));
  Pages.ActivePage:=Pages.Pages[0];
end;
  if Exo1<>nil then begin
    Exo1.OnChange:=ExoOnChange ;
    Exo1.Value:=VH^.Entree.Code ;
  end ;
  if ExoBds<>nil then begin
    ExoBds.OnChange:=ExoBdsOnChange ;
    FillComboBds(ExoBds) ;
    if Exo2<>nil then FillComboValues(ExoBds, Exo2);
    ExoBdsOnChange(ExoBds);
  end ;
  if Cpt1<>nil then Cpt1.OnExit:=CompteOnExit ;
  if Cpt2<>nil then Cpt2.OnExit:=CompteOnExit ;
  if Date1<>nil then Date1.OnExit:=DateOnExit ;
  if Date2<>nil then Date2.OnExit:=DateOnExit ;
  if Comparatif<>nil then Comparatif.OnClick:=ComparatifOnClick;
  if Etab<>nil then Etab.ItemIndex := 0 ;
  if Devise<>nil then Devise.ItemIndex := 0;
  if (Devise<>nil) and (DevPivot<>nil) then
    DevPivot.Text:=Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
  InitCritereAvance;
END ;

procedure TOF_BALANCE.OnUpdate ;
BEGIN
  inherited ;
  if Trim(Cpt1.Text) = '' then Cpt1.Text := '00000000';
  if Trim(Cpt2.Text) = '' then Cpt2.Text := 'ZZZZZZZZ';
  Rupture.Text:=IntToStr(Ruptures.Value);
END;

procedure TOF_BALANCE.OnNew;
begin
if ExoBds.Visible and (ExoBds.ItemIndex<0) and (Comparatif.Checked) then begin
  HShowMessage('0;Il n''y a pas de balance comparative;Aucun enregistrement à imprimer;E;O;O;O;','','') ;
  SetControlEnabled(TFQRS1(Ecran).BValider.Name, false) ;
  Abort;
end;
end;

procedure TOF_BALANCE.InitCritereAvance;
begin
  if Ruptures <> nil then begin
    Ruptures.Value := 0;
    Rupture.Text := InttoStr(Ruptures.Value);
  end;
  if Blocnote <> nil then
    Blocnote.Checked := false;
  if Comparatif <> nil then begin
    Comparatif.Checked := true;
    ComparatifOnClick(nil);
  end;
end;

procedure TOF_BALANCE.CompteOnExit(Sender: TObject) ;
BEGIN
DoCompteOnExit(THEdit(Sender), Cpt1, Cpt2);
END;

procedure TOF_BALANCE.ExoOnChange(Sender: TObject) ;
BEGIN
if Sender = Exo1 then begin
  DoExoToDateOnChange(Exo1, Date1, Date2);
  DateD := StrToDate(Date1.Text);
  DateF := StrToDate(Date2.Text);
end;
END;

procedure TOF_BALANCE.ExoBdsOnChange(Sender: TObject);
var s: string;
BEGIN
s:=ExoBds.Value;
Exo2.ItemIndex:=ExoBds.ItemIndex;
ReadTokenST(s);
Date3.Text:=ReadTokenST(s);
Date4.Text:=ReadTokenST(s);
END;

procedure TOF_BALANCE.DateOnExit(Sender: TObject) ;
BEGIN
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
END;


procedure TOF_BALANCE.ComparatifOnClick(Sender: TObject);
begin
  if Comparatif.Checked then begin
    Histo.Enabled := true;
    ExoBds.Enabled := true;
    TypeBal.Text := ':BDS§BAL';
  end else begin
    Histo.Enabled := false;
    ExoBds.Enabled := false;
    TypeBal.Text := '';
  end;
end;

initialization
RegisterClasses([TOF_BDS]) ;
RegisterClasses([TOF_BALANCE]) ;

end.
