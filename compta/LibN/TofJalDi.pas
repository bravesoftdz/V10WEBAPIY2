unit TofJalDi;

interface

uses Classes, StdCtrls, UTof, HCtrls, QRS1, Ent1, SysUtils, dbtables, TofMeth;

// TOF.OnArgument sur show
// TOF.Load après le valideclick

type
  TOF_JALDI = class(TOF_Meth)
  public
    Jal1, Jal2: THEdit;
    Exo, Etab, Devise : THValComboBox ;
    Date1, Date2 : THEdit;
    Num1, Num2: THEdit;
    DateD, DateF: TDatetime;
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnLoad ; override ;
    procedure JalOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure NumOnExit(Sender: TObject) ;
//    procedure JalElipsisClick(Sender: TObject);
  end ;

implementation

uses HEnt1;


procedure TOF_JALDI.ExoOnChange(Sender: TObject) ;
BEGIN
DoExoToDateOnChange(Exo, Date1, Date2);
DateD := StrToDate(Date1.Text);
DateF := StrToDate(Date2.Text);
END;

procedure TOF_JALDI.JalOnExit(Sender: TObject) ;
BEGIN
DoJalOnExit(THEdit(Sender), Jal1, Jal2);
END;

procedure TOF_JALDI.DateOnExit(Sender: TObject) ;
BEGIN
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
END;

procedure TOF_JALDI.NumOnExit(Sender: TObject) ;
BEGIN
DoNumOnExit(THEdit(Sender), Num1, Num2);
END;

procedure TOF_JALDI.OnArgument(Arguments : string) ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
Exo:=THValComboBox(FindComponent('E_EXERCICE')) ;
Jal1:=THEdit(FindComponent('E_JOURNAL')) ;
Jal2:=THEdit(FindComponent('E_JOURNAL_')) ;
Date1:=THEdit(FindComponent('E_DATECOMPTABLE')) ;
Date2:=THEdit(FindComponent('E_DATECOMPTABLE_')) ;
Etab:=THValComboBox(FindComponent('E_ETABLISSEMENT')) ;
Devise:=THValComboBox(FindComponent('E_DEVISE')) ;
Num1:=THEdit(FindComponent('E_NUMEROPIECE')) ;
Num2:=THEdit(FindComponent('E_NUMEROPIECE_')) ;
if Exo<>nil then begin Exo.OnChange:=ExoOnChange; Exo.Value:=VH^.Entree.Code; end ;
if Jal1<>nil then Jal1.OnExit:=JalOnExit;
if Jal2<>nil then Jal2.OnExit:=JalOnExit;
//if Jal1<>nil then begin Jal1.OnExit:=JalOnExit; Jal1.OnElipsisClick := JalElipsisClick; end;
//if Jal2<>nil then begin Jal2.OnExit:=JalOnExit; Jal1.OnElipsisClick := JalElipsisClick; end;
if Date1<>nil then Date1.OnExit:=DateOnExit ;
if Date2<>nil then Date2.OnExit:=DateOnExit ;
if Num1 <> nil then Num1.OnExit := NumOnExit;
if Num2 <> nil then Num2.OnExit := NumOnExit;
if Etab<>nil then Etab.ItemIndex := 0 ;
if Devise<>nil then Devise.ItemIndex := Devise.Values.IndexOf(V_PGI.DevisePivot);
end;
END;

procedure TOF_JALDI.OnLoad ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
if Trim(Jal1.Text) = '' then Jal1.Text := '000';
if Trim(Jal2.Text) = '' then Jal2.Text := 'ZZZ';
end;
END;


initialization
RegisterClasses([TOF_JALDI]) ;

end.
