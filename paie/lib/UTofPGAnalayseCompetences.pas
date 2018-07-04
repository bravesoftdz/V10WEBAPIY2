{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGANALYSECOMPETENCES ()
Mots clefs ... : TOF;PGANALYSECOMPETENCES
*****************************************************************}
Unit UTofPGAnalayseCompetences ;

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
     EntPaie,
     P5Def,
     HMsgBox,
     PGOutilsFormation,
     PGOutils2,
     ParamDat,
     UTOF ; 

Type
  TOF_PGANALYSECOMPETENCES = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure ExitEdit(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
  end ;

Implementation

procedure TOF_PGANALYSECOMPETENCES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSECOMPETENCES.OnLoad ;
var Where : String;
begin
  Inherited ;
  If GetCheckBoxState('CKSORTIE') = CbChecked then Where :='(PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATEARRET')))+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) '
  else Where := '';
  If TCheckBox(GetControl('CDIPLOME')) <> Nil then
  begin
       If GetCheckBoxState('CDIPLOME') = CbChecked then
       begin
            If Where <> '' then Where := Where + ' AND PFO_DATEFIN IN (SELECT MAX(PFO_DATEFIN) FROM FORMATIONS WHERE PFO_SALARIE=PSA_SALARIE AND PFO_NATUREFORM="004")'
            else Where := 'PFO_DATEFIN IN (SELECT MAX(PFO_DATEFIN) FROM FORMATIONS WHERE PFO_SALARIE=PSA_SALARIE AND PFO_NATUREFORM="004")';
       end;
  end;
  SetControlText('XX_WHERE',Where);
end ;

procedure TOF_PGANALYSECOMPETENCES.OnArgument (S : String ) ;
var Num : Integer;
    Edit : THEdit;
    Check : TCheckBox;
begin
  Inherited ;
  //Gestion onglet salariés
  For Num  :=  1 to VH_Paie.PGNbreStatOrg do
  begin
     if Num >4 then Break;
     VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
  end;
  VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
  For Num := 1 to VH_Paie.PgNbCombo do
  begin
     if Num >4 then Break;
     VisibiliteChampLibreSal(IntToStr(Num),GetControl ('PSA_LIBREPCMB'+IntToStr(Num)),GetControl ('TPSA_LIBREPCMB'+IntToStr(Num)));
  end;
  //Gestion onglet compétences
  For Num  :=  1 to 5 do
  begin
     VisibiliteChampCompetence (IntToStr(Num),GetControl ('PCO_TABLELIBRERH'+IntToStr(Num)),GetControl ('TPCO_TABLELIBRERH'+IntToStr(Num)));
  end;
  For Num  :=  1 to 3 do
  begin
       VisibiliteChampCompetenceRess (IntToStr(Num),GetControl ('PCH_TABLELIBRECR'+IntToStr(Num)),GetControl ('TPCH_TABLELIBRECR'+IntToStr(Num)));
  end;
  Edit := THEdit(GetControl('PSA_SALARIE'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
  Edit := THEdit(GetControl('DATEARRET'));
  if Edit <> nil then Edit.OnElipsisClick := DateElipsisclick;
  Check := TCheckBox(GetControl('CKSORTIE'));
  If Check <> Nil then Check.OnClick:=OnClickSalarieSortie;
end ;



procedure TOF_PGANALYSECOMPETENCES.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGANALYSECOMPETENCES.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_PGANALYSECOMPETENCES.DateElipsisclick(Sender: TObject);
var
  key : char;
begin
  key := '*';
  ParamDate (Ecran, Sender, Key);
end;

Initialization
  registerclasses ( [ TOF_PGANALYSECOMPETENCES ] ) ;
end.

