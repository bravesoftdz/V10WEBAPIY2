{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 07/02/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit TofTabAv;

interface

uses Classes, StdCtrls, UTof, HCtrls,Ent1, SysUtils, TofMeth, HTB97,
{$IFDEF EAGLCLIENT}
     eQRS1,
     MaineAGL,  // AGLLanceFiche
{$ELSE}
     QRS1,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,  // AGLLanceFiche
//     HQuickrp,
{$ENDIF}
     CritEdt,  // TCritEdtChaine
     AGLInit,  // TheData
     Filtre,   // Filtre
     Extctrls; // TTimer

// TOF.OnArgument sur show
// TOF.Load après le valideclick

type
  TOF_TABAV = class(TOF_Meth)
  private
    Jal1, Jal2: THEdit;
    Exo, Devise : THValComboBox ;
    Date1, Date2 : THEdit;
    DateD, DateF: TDatetime;
    Periode, NbPeriode, DevIdx: THEdit; // FQ 20486 
    DevPivot: THEdit;
    BHelp: TToolbarButton97;
    xxWhere : THEdit;
    procedure BHelpClick(Sender: TObject);
    procedure JalOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnUpdate ; override ;
    procedure OnNew ; override ;
    procedure OnLoad ; override ;
    procedure FTimerTimer(Sender: TObject); // GCO
  end ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  HEnt1;

procedure TOF_TABAV.JalOnExit(Sender: TObject) ;
BEGIN
DoJalOnExit(THEdit(Sender), Jal1, Jal2);
END;

procedure TOF_TABAV.ExoOnChange(Sender: TObject) ;
var ExoDate: TExoDate;
    MM,AA,NbMois: word;
BEGIN
DoExoToDateOnChange(Exo, Date1, Date2);
DateD := StrToDate(Date1.Text);
DateF := StrToDate(Date2.Text);
ExoDate.Deb := StrToDate(Date1.Text);
ExoDate.Fin := StrToDate(Date2.Text);
NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;
if NbMois > 12 then
  Date2.Text := DateToStr(FinDeMois(PlusMois(ExoDate.Deb, 11)));
END;

procedure TOF_TABAV.DateOnExit(Sender: TObject) ;
BEGIN
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
END;

procedure TOF_TABAV.OnArgument(Arguments : string) ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
NatureEtat := 'JTA';
//RRO le 14032003
ChoixEtat:=TRUE ;
Exo:=THValComboBox(GetControl('E_EXERCICE')) ;
Jal1:=THEdit(GetControl('E_JOURNAL')) ;
Jal2:=THEdit(GetControl('E_JOURNAL_')) ;
Date1:=THEdit(GetControl('E_DATECOMPTABLE')) ;
Date2:=THEdit(GetControl('E_DATECOMPTABLE_')) ;
{ FQ 20486 BVE 18.06.07 }
Periode:=THEdit(GetControl('PERIODE')) ;
NbPeriode:=THEdit(GetControl('NbPERIODE')) ;
DevIdx:=THEdit(GetControl('DEVIDX')) ;
{ END FQ 20486 }
Devise:=THValComboBox(GetControl('E_DEVISE')) ;
DevPivot:=THEdit(GetControl('DEVPIVOT'));
BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
xxWhere:=THEdit(GetControl('XX_WHERE'));
end;
if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick;
if Exo<>nil then begin
  Exo.OnChange:=ExoOnChange ;
  Exo.Value:=VH^.Entree.Code ;
end ;
if Jal1<>nil then Jal1.OnExit:=JalOnExit ;
if Jal2<>nil then Jal2.OnExit:=JalOnExit ;
if Date1<>nil then Date1.OnExit:=DateOnExit ;
if Date2<>nil then Date2.OnExit:=DateOnExit ;
if Devise<>nil then Devise.ItemIndex := 0;
if (Devise<>nil) and (DevPivot<>nil) then
  DevPivot.Text:=Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];

// GCO - 01/07/2002
FTimer.OnTimer := FTimerTimer;
// FIN GCO

END ;

procedure TOF_TABAV.OnNew;
BEGIN
inherited;
  with TFQRS1(Ecran) do begin
    HelpContext:=7454000;
    if CodeEtat = 'JTL' then begin
      Caption := TraduireMemoire('Tableau d''avancement Mouvements');
    end
    else if CodeEtat = 'JTM' then begin
      Caption := TraduireMemoire('Tableau d''avancement Montants');
    end;
    UpdateCaption(Ecran);
  end;
END;

procedure TOF_TABAV.OnUpdate ;
var
  ExoDate : TExoDate;
  MM,AA,NbMois: Word;
BEGIN
inherited ;
if Trim(Jal1.Text) = '' then Jal1.Text := '000';
if Trim(Jal2.Text) = '' then Jal2.Text := 'ZZZ';
ExoDate.Deb := StrToDate(Date1.Text);
ExoDate.Fin := StrToDate(Date2.Text);
NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;
if Periode <> nil then Periode.Text := IntToStr(GetPeriode(ExoDate.Deb));
if NbPeriode <> nil then NbPeriode.Text := IntToStr(NbMois);
if DevIdx <> nil then DevIdx.Text := IntToStr(Devise.ItemIndex);
  { FQ 20969 BVE 18.07.07 }
  if xxWhere <> nil then
  begin
     xxWhere.Text := '(' + xxWhere.Text + ') AND ' +
                     '( E_DATECOMPTABLE >= "' + UsDateTime(StrToDate(Date1.Text)) + '" AND '+
                     '  E_DATECOMPTABLE <= "' + UsDateTime(StrToDate(Date2.Text)) + '" )';
  end;
  { END FQ 20969 }
END;

procedure TOF_TABAV.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;

procedure TOF_TABAV.OnLoad;
begin
  inherited;
  if xxWhere <> nil then begin
     xxWhere.Text := 'E_QUALIFPIECE="N" and E_ECRANOUVEAU="N"';
     if ComboEtab.ItemIndex > 0 then
        xxWhere.Text := xxWhere.Text + ' and E_ETABLISSEMENT="' + ComboEtab.Value + '"';
     if Devise.ItemIndex > 0 then
        xxWhere.Text := xxWhere.Text + ' and E_DEVISE="' + Devise.Value + '"';
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2002
Modifié le ... : 06/09/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_TABAV.FTimerTimer(Sender: TObject);
begin
  if FCritEdtChaine <> nil then
  begin
    with FCritEdtChaine do
    begin
      if CritEdtChaine.UtiliseCritStd then
      begin
        Exo.Value := CritEdtChaine.Exercice.Code ;
        Date1.Text := DateToStr(CritEdtChaine.Exercice.Deb);
        Date2.Text := DateToStr(CritEdtChaine.Exercice.Fin);
      end;
    end;
  end;
  inherited;
end;

initialization
RegisterClasses([TOF_TABAV]) ;

end.
