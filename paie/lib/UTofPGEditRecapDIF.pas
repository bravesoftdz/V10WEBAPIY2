{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITRECAPDIF ()
Mots clefs ... : TOF;PGEDITRECAPDIF
*****************************************************************}
{
PT1 26/01/2007  V_80 FC Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PT2  09/10/2007  FL  V_7  Prise en compte des salariés sortis + Plus de sélection établissement
}
Unit UTofPGEditRecapDIF ;

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
     paramsoc,
     PGEditOutils2,
     PGOutils2,
     EntPaie,
     LookUp,
     UTOF ,
     PGOutils; 

Type
  TOF_PGEDITRECAPDIF = Class (TOF)
    procedure OnLoad               ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
      {$IFDEF EMANAGER}
    TypeUtilisat : String;
    {$ENDIF}
    procedure ClickSortie(Sender : TObject);
    procedure ExitEdit(Sender: TObject);
    {$IFDEF EMANAGER}
    procedure SalarieElipsisClick(Sender : Tobject);
    procedure RespElipsisClick(Sender : TObject);
    {$ENDIF}
  end ;

Implementation

procedure TOF_PGEDITRECAPDIF.OnLoad  ;
{$IFDEF EMANAGER}
var StWhere : String;
    MultiNiveau : Boolean;
    DateArret : TDateTime; //PT2
{$ENDIF}
begin
     Inherited ;
     {$IFDEF EMANAGER}
     MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
     If typeUtilisat = 'R' then
     begin
          If MultiNiveau then StWhere := '(PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
          else StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     end
      else StWhere := 'PSE_RESPONSFOR IN '+
     '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     //PT2 - Fin
     if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
     Begin
          DateArret:=StrtoDate(GetControlText('DATEARRET'));
          StWhere := StWhere + ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
     End;
     //PT2 - Fin
     SetControlText('XX_WHERE',StWhere);
     {$ENDIF}
end;

procedure TOF_PGEDITRECAPDIF.OnArgument (S : String ) ;
var Min,Max : String;
    Check : TCheckBox;
    Defaut: THEdit;
begin
  Inherited ;
  SetControlText('DATEARRET',DateToStr(V_PGI.DateEntree));
  SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));
  {$IFNDEF EMANAGER} //PT2
  SetControlText('PSA_ETABLISSEMENT', Min);
  SetControlText('PSA_ETABLISSEMENT_', Max);
  {$ENDIF}
//  SetControltext('PSA_DATEENTREE',DateToStr(IDate1900));
//  SetControltext('PSA_DATEENTREE_',DateToStr(V_PGI.DateEntree));
  Check := TCheckBox(GetControl('CKSORTIE'));
  If Check <> Nil then
  begin
       Check.OnClick := ClickSortie;
       Check.Checked := True;
  end;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  {$IFDEF EMANAGER}
  SetControlVisible('CMULTINIVEAU',True);
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnElipsisClick:= SalarieElipsisClick;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut<>nil then Defaut.OnElipsisClick:= SalarieElipsisClick;
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  If Defaut<>nil then Defaut.OnElipsisClick := SalarieElipsisClick;
  SetControlVisible('PSE_RESPONSFOR',True);
  SetControlVisible('TPSE_RESPONSFOR',True);
  SetControlVisible('LIBRESP',True);
  SetControlCaption('LIBRESP','');
  Defaut:=ThEdit(getcontrol('PSE_RESPONSFOR'));
  If Defaut<>nil then Defaut.OnElipsisClick := RespElipsisClick;
  {$ELSE}
  RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
  SetControlText('PSA_SALARIE',Min);
  SetControlText('PSA_SALARIE_',Max);
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  {$ENDIF}
end ;

procedure TOF_PGEDITRECAPDIF.ClickSortie(Sender : TObject);
begin
  If GetCheckBoxState('CKSORTIE') = CbChecked then
  SetControlText('XX_WHERE','(PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATEARRET')))+'")')
  else SetControlText('XX_WHERE','');
end;

procedure TOF_PGEDITRECAPDIF.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

{$IFDEF EMANAGER}
procedure TOF_PGEDITRECAPDIF.SalarieElipsisClick(Sender : Tobject);
var StWhere,StOrder : String;
    DateArret : TDateTime;
begin
     //PT2 - Début
     If (GetControlText('CKSORTIE') = 'X') and (IsValidDate(GetControlText('DATEARRET'))) Then
     Begin
          DateArret:=StrtoDate(GetControlText('DATEARRET'));
          StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(DateArret) + '") AND ';
     End;
     If (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) Then
     Begin
       If typeUtilisat = 'R' then StWhere := StWhere + 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'
       Else StWhere := StWhere + 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")';
     End
     Else
     Begin
     //PT2 - Fin
       If typeUtilisat = 'R' then StWhere := StWhere + 'PSE_CODESERVICE IN '+ //PT1
     '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
     ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
     ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))'
       else StWhere := StWhere + 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+ //PT1
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     End;
     If GetControlText('RESPONSFOR') <> '' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'"';
     //If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'; //PT1
     StOrder := 'PSA_SALARIE';
     StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
     LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;


procedure TOF_PGEDITRECAPDIF.RespElipsisClick(Sender : TObject);
var StOrder,StWhere : String;
begin
         If TypeUtilisat = 'R' then StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
        else StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
        StOrder := 'PSI_LIBELLE';
        LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
end;

{$ENDIF}

Initialization
  registerclasses ( [ TOF_PGEDITRECAPDIF ] ) ;
end.

