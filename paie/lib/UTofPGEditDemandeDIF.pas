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

}
Unit UTofPGEditDemandeDIF ;

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
     LookUp,
     paramsoc,
     PGEditOutils2,
     PGOutils2,
     PGOutils,
     P5DEF,
     EntPaie,
     UTOF ; 

Type
  TOF_PGEDITDEMANDEDIF = Class (TOF)
    procedure OnLoad  ;   override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFDEF EMANAGER}
    TypeUtilisat : String;
    {$ENDIF}
    procedure ExitEdit(Sender: TObject);
    {$IFDEF EMANAGER}
    procedure SalarieElipsisClick(Sender : Tobject);
    procedure RespElipsisClick(Sender : TObject);
    {$ENDIF}
  end ;

Implementation

procedure TOF_PGEDITDEMANDEDIF.OnLoad ;
{$IFDEF EMANAGER}
var StWhere : String;
    MultiNiveau : Boolean;
{$ENDIF}
begin
     Inherited ;
     {$IFDEF EMANAGER}
     MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
     If typeUtilisat = 'R' then
     begin
          If MultiNiveau then StWhere := '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
          else StWhere := 'PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     end
     else StWhere := '(PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))))';
     SetControlText('XX_WHERE',StWhere);
     {$ENDIF}
end;

procedure TOF_PGEDITDEMANDEDIF.OnArgument (S : String ) ;
var Defaut: THEdit;
    Num : Integer;
    {$IFNDEF EMANAGER}
    Min,Max : String;
    {$ENDIF}
begin
  Inherited ;
  SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
       if Num >4 then Break;
       VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
  end;
  VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
  {$IFDEF EMANAGER}
  SetControlVisible('CMULTINIVEAU',True);
  Defaut:=ThEdit(getcontrol('PFI_SALARIE'));
  If Defaut<>nil then Defaut.OnElipsisClick:= SalarieElipsisClick;
  Defaut:=ThEdit(getcontrol('PFI_SALARIE_'));
  If Defaut<>nil then Defaut.OnElipsisClick:= SalarieElipsisClick;
  Defaut:=ThEdit(getcontrol('PFI_RESPONSFOR'));
  If Defaut<>nil then Defaut.OnElipsisClick := RespElipsisClick;
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  {$ELSE}
  RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
  SetControlText('PFI_SALARIE',Min);
  SetControlText('PFI_SALARIE_',Max);
  {$ENDIF}
end ;

procedure TOF_PGEDITDEMANDEDIF.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

{$IFDEF EMANAGER}
procedure TOF_PGEDITDEMANDEDIF.SalarieElipsisClick(Sender : Tobject);
var StWhere,StOrder : String;
begin
     If typeUtilisat = 'R' then StWhere := 'PSE_CODESERVICE IN '+
     '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
     ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
     ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))'
     else StWhere := 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     If GetControlText('RESPONSFOR') <> '' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'"';
     If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     StOrder := 'PSA_SALARIE';
     StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
     LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;

procedure TOF_PGEDITDEMANDEDIF.RespElipsisClick(Sender : TObject);
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
  registerclasses ( [ TOF_PGEDITDEMANDEDIF ] ) ;
end.

