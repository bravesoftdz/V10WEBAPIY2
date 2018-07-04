unit TofQrS1;

interface

uses Classes, StdCtrls, UTof, HCtrls, QRS1, Ent1, SysUtils, dbtables;

// TOF.OnArgument sur show
// TOF.Load après le valideclick

type
  TOF_QRS1 = class(TOF)
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnLoad ; override ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure PeriodeOnChange(Sender: TObject) ;
  end ;

implementation

uses HEnt1;

//V_PGI.DevisePivot
procedure TOF_QRS1.ExoOnChange(Sender: TObject) ;
var Fiche : TFQRS1 ; Exo : THValComboBox ; Date1, Date2 : THCritMaskEdit;
    P1, P2 : THValComboBox ;
    ExoDate: TExoDate;
    MM,AA,NbMois: word;
BEGIN
Fiche:=TFQRS1(Ecran) ;
Exo:=THValComboBox(Fiche.FindComponent('E_EXERCICE')) ;
Date1:=THCritMaskEdit(Fiche.FindComponent('E_DATECOMPTABLE')) ;
Date2:=THCritMaskEdit(Fiche.FindComponent('E_DATECOMPTABLE_')) ;
P1:=THValComboBox(Fiche.FindComponent('PERIODE1')) ;
P2:=THValComboBox(Fiche.FindComponent('PERIODE2')) ;
if (Exo<>nil) and (Date1<>nil) and (Date2<>nil) and (P1 <> nil) and (P2 <> nil) then
begin
  ExoToDates(Exo.Value, Date1, Date2);
  ListePeriode(Exo.Value, P1.Items, P1.Values, true);
  ListePeriode(Exo.Value, P2.Items, P2.Values, false);
  P1.ItemIndex := P1.Values.IndexOf(Date1.Text);
  P2.ItemIndex := P2.Values.IndexOf(Date2.Text);
  ExoDate.Deb := StrToDate(Date1.Text);
  ExoDate.Fin := StrToDate(Date2.Text);
  NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;
  if NbMois > 12 then
  begin
    P1.ItemIndex := P1.Values.IndexOf(DateToStr(PlusMois(ExoDate.Fin, -12) +1));
    P1.Enabled := false;
  end else P1.Enabled := true;
end;
END;

procedure TOF_QRS1.PeriodeOnChange(Sender: TObject) ;
var Fiche : TFQRS1 ; Date1, Date2 : THCritMaskEdit;
    P1, P2 : THValComboBox ;
BEGIN
Fiche:=TFQRS1(Ecran) ;
P1:=THValComboBox(Fiche.FindComponent('PERIODE1')) ;
P2:=THValComboBox(Fiche.FindComponent('PERIODE2')) ;
Date1:=THCritMaskEdit(Fiche.FindComponent('E_DATECOMPTABLE')) ;
Date2:=THCritMaskEdit(Fiche.FindComponent('E_DATECOMPTABLE_')) ;
if (Sender = P1) and (Date1 <> nil) then begin
  Date1.Text := THValComboBox(Sender).Value;
  if StrToDate(Date1.Text) > StrToDate(Date2.Text) then begin
   P2.ItemIndex := P1.ItemIndex;
   Date2.Text := Date1.Text;
  end;
end else if (Sender = P2) and (Date2 <> nil) then begin
  Date2.Text := THValComboBox(Sender).Value;
  if StrToDate(Date1.Text) > StrToDate(Date2.Text) then begin
    P1.ItemIndex := P2.ItemIndex;
    Date1.Text := Date2.Text;
  end;
end;
END;


procedure TOF_QRS1.OnArgument(Arguments : string) ;
var Fiche : TFQRS1 ; Exo : THValComboBox ;
begin
inherited ;
Fiche:=TFQRS1(Ecran) ;
Exo:=THValComboBox(Fiche.FindComponent('E_EXERCICE')) ;
if Exo<>nil then
  begin
  Exo.OnChange:=ExoOnChange ;
  Exo.Value:=VH^.Entree.Code ;
  end ;
Exo:=THValComboBox(Fiche.FindComponent('PERIODE1')) ;
if Exo<>nil then Exo.OnChange:=PeriodeOnChange ;
Exo:=THValComboBox(Fiche.FindComponent('PERIODE2')) ;
if Exo<>nil then Exo.OnChange:=PeriodeOnChange ;
Exo:=THValComboBox(Fiche.FindComponent('E_ETABLISSEMENT')) ;
if Exo<>nil then Exo.ItemIndex := 0 ;
Exo:=THValComboBox(Fiche.FindComponent('E_DEVISE')) ;
if Exo<>nil then Exo.ItemIndex := Exo.Values.IndexOf(V_PGI.DevisePivot);
end ;

procedure TOF_QRS1.OnLoad ;
var
  aEdit : TEdit ;
  Fiche : TFQRS1 ;
  Exo : THValComboBox ;
  ExoDate : TExoDate;
  MM,AA,NbMois: Word;
  Date1, Date2 : THCritMaskEdit;
BEGIN
inherited ;
Fiche:=TFQRS1(Ecran) ;
Date1:=THCritMaskEdit(Fiche.FindComponent('E_DATECOMPTABLE')) ;
Date2:=THCritMaskEdit(Fiche.FindComponent('E_DATECOMPTABLE_')) ;
//QuelDateDeExo(Exo.Value, ExoDate);
ExoDate.Deb := StrToDate(Date1.Text);
ExoDate.Fin := StrToDate(Date2.Text);
NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;

aEdit := TEdit(Fiche.FindComponent('PERIODE'));
if aEdit <> nil then aEdit.Text := IntToStr(GetPeriode(ExoDate.Deb));
aEdit := TEdit(Fiche.FindComponent('NBPERIODE'));
if aEdit <> nil then aEdit.Text := IntToStr(NbMois);
END;



initialization
RegisterClasses([TOF_QRS1]) ;

end.
