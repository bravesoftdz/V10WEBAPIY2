{***********UNITE*************************************************
Auteur  ...... :  3GG
Créé le ...... : 14/02/2006
Modifié le ... : 15/02/2006
Description .. : Source TOF de la FICHE : PGANALYSESAL ()
Mots clefs ... : TOF;PGANALYSESAL
*****************************************************************}
Unit PGANALYSESAL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     stat ,
     UTOF ; 

Type
  TOF_PGANALYSESAL = Class (TOF)
  private
    PSA_SALARIE : THedit;
    PSA_SALARIE_ : THedit;

    //  Complément à zéro des matricules saisis sur - 10 caractères
    procedure Psa_SalarieExit(Sender: TObject);
    procedure Psa_Salarie_Exit(Sender: TObject);
  public
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arguments : String ) ; override ;
  end ;

Implementation
  uses entpaie,
       P5def,
       Pgoutils,
       Pgoutils2;
procedure TOF_PGANALYSESAL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESAL.OnArgument (Arguments : String ) ;
var
  Arg, Ch1, Ch2, Ch3: string;
  Num: Integer;
  Pref: string;

  begin
  Inherited ;
  Arg := Arguments;
  Arg := Trim(Arg);
  Psa_Salarie := ThEdit(Getcontrol('PSA_SALARIE'));
  Psa_Salarie_ := ThEdit(Getcontrol('PSA_SALARIE_'));
  Pref := 'PSA';
      //  Complément à zéro des matricules saisis sur - 10 caractères
  if PSA_SALARIE <> nil then Psa_Salarie.OnExit := Psa_SalarieExit;
  if PSA_SALARIE_ <> nil then Psa_Salarie_.OnExit := Psa_Salarie_Exit;

  for Num := 1 to 4 do
  begin
    if Num > 4 then Break;
    Ch1 := Pref + '_TRAVAILN' + IntToStr(Num);
    Ch2 := 'T' + Pref + '_TRAVAILN' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_TRAVAILN' + IntToStr(Num);
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;

  for Num := 1 to 4 do
  begin
    if Num > 4 then Break;
    Ch1 := Pref + '_LIBREPCMB' + IntToStr(Num);
    Ch2 := 'T' + Pref + '_LIBREPCMB' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_LIBREPCMB' + IntToStr(Num);
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;
end ;

//  Complément à zéro des matricules saisis sur - 10 caractères
procedure TOF_PGANALYSESAL.Psa_SalarieExit(Sender: TObject);
var
  Edit : ThEdit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;
procedure TOF_PGANALYSESAL.Psa_Salarie_Exit(Sender: TObject);
var
  Edit : ThEdit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;


Initialization
  registerclasses ( [ TOF_PGANALYSESAL ] ) ; 
end.
