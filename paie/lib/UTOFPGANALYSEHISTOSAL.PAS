{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPGANALYSEHISTOSAL ()
Mots clefs ... : TOF;UTOFPGANALYSEHISTOSAL
*****************************************************************
PT1   21/06/2007 FC V_72 FQ 14330 Accès à tort aux salariés confidentiels
}
Unit UTOFPGANALYSEHISTOSAL ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
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
     UTOF,EntPaie, PGOutils2 ; 

Type
  TOF_UTOFPGANALYSEHISTOSAL = Class (TOF)
    public
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      Procedure PGRechSal (FF : TForm) ;
      procedure ExitEdit(Sender: TObject);
    private
      procedure ActiveWhere(Sender: TObject);
      procedure OnClickSalarieSortie(Sender: TObject);
  end ;

Implementation

procedure TOF_UTOFPGANALYSEHISTOSAL.ActiveWhere(Sender: TObject);
var
  StWhere,Info : String;
  DateArret,DateApplic : TDateTime;
  StDateArret : String;
  i : integer;
  St : String;
begin
  SetControlText('XX_WHERE','');
  
  StWhere := ' AND PHD_PGTYPEINFOLS="SAL"';

  // Critère matricule salarié
  if (GetControlText('PSA_SALARIE') <> '') then
    stWhere := stWhere + ' AND PSA_SALARIE = "' + GetControlText('PSA_SALARIE') + '"';

  // Critère nom salarié
  if (GetControlText('PSA_LIBELLE') <> '') then
    stWhere := stWhere + ' AND PSA_LIBELLE LIKE "' + GetControlText('PSA_LIBELLE') + '%"';

  // Critère établissement
  if (GetControlText('PSA_ETABLISSEMENT') <> '') and (GetControlText('PSA_ETABLISSEMENT') <> '<<Tous>>') then
    stWhere := stWhere + ' AND PSA_ETABLISSEMENT = "' + GetControlText('PSA_ETABLISSEMENT') + '"';

  // Information modifiée
  if (GetControlText('PHD_PGINFOSMODIF') <> '') and (GetControlText('PHD_PGINFOSMODIF') <> '<<Tous>>') then
  begin
    Info := GetControlText('PHD_PGINFOSMODIF');
    stWhere := stWhere + ' AND (';
    while Info <> '' do
    begin
      stWhere := stWhere + ' PHD_PGINFOSMODIF = "' + READTOKENST(Info) + '"';
      if (Info <> '') then
        stWhere := stWhere + ' OR ';
    end;
    stWhere := stWhere + ')';
  end;

  // Gérer la case à cocher exclure
  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET'))) then
  Begin
    DateArret := StrtoDate(GetControlText('DATEARRET'));
    StDateArret := ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
    StDateArret := StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
    stWhere := stWhere + StDateArret;
  End;

  // Zone de travail 1 à 4
  for i := 1 to 4 do
  begin
    if (GetControlText('PSA_TRAVAILN' + IntToStr(i)) <> '') and (GetControlText('PSA_TRAVAILN' + IntToStr(i)) <> '<<Tous>>') then
      stWhere := stWhere + ' AND PSA_TRAVAILN' + IntToStr(i) + ' = "' + GetControlText('PSA_TRAVAILN' + IntToStr(i)) + '"';
  end;

  // Code statistique
  if (GetControlText('PSA_CODESTAT') <> '') and (GetControlText('PSA_CODESTAT') <> '<<Tous>>') then
    stWhere := stWhere + ' AND PSA_CODESTAT = "' + GetControlText('PSA_CODESTAT') + '"';

  // Dates application
  DateApplic := StrtoDate(GetControlText('PHD_DATEAPPLIC'));
  stWhere := stWhere + ' AND PHD_DATEAPPLIC>="' + UsDateTime(DateApplic) + '"';
  DateApplic := StrtoDate(GetControlText('PHD_DATEAPPLIC_'));
  stWhere := stWhere + ' AND PHD_DATEAPPLIC<="' + UsDateTime(DateApplic) + '"';

  //DEB PT1
  St := SQLConf('SALARIES');
  if St <> '' then
    stWhere := stWhere + ' AND ' + St;
  //FIN PT1

  SetControlText('XX_WHERE',stWhere);
  SetControlText('XX_ORDERBY','PHD_SALARIE,PHD_PGINFOSMODIF,PHD_DATEAPPLIC');
end;

procedure TOF_UTOFPGANALYSEHISTOSAL.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_UTOFPGANALYSEHISTOSAL.OnLoad ;
begin
  Inherited ;
  ActiveWhere(nil);
end ;

procedure TOF_UTOFPGANALYSEHISTOSAL.OnArgument (S : String ) ;
var
  Check : TCheckBox;
begin
  Inherited ;
  Check := TCheckBox(GetControl('CKSORTIE'));
  if (Check <> nil) then
    Check.OnClick:=OnClickSalarieSortie;

  PGRechSal (Ecran);
end ;

procedure TOF_UTOFPGANALYSEHISTOSAL.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_UTOFPGANALYSEHISTOSAL.PGRechSal(FF: TForm);
var i : Integer;
LeControl : TControl;
begin
  for i := 0 to FF.ComponentCount - 1 do
  begin
    LeControl := TControl(FF.Components[i]);
    if LeControl is THLabel then continue;
    if (LeControl is THEdit) AND ((pos('SALARIE', LeControl.Name) > 0) OR (pos('SALARIE_', LeControl.Name) > 0)) then
      begin
      THEdit(LeControl).OnExit := ExitEdit;
      end
      else Continue;
  end;
end;

Initialization
  registerclasses ( [ TOF_UTOFPGANALYSEHISTOSAL ] ) ;
end.
