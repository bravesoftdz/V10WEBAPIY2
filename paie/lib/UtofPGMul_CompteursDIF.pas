{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULCOMPTEURSDIF ()
Mots clefs ... : TOF;PGMULCOMPTEURSDIF
*****************************************************************}
{
PT1 26/01/2007  V_80 FC Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PT2  09/10/2007  FL  V_7  Prise en compte des salariés sortis
}
Unit UtofPGMul_CompteursDIF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     FE_Main,
     HDB,
{$else}
     eMul,
     uTob,
     MainEAGL,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HQry,
     HMsgBox,
     EntPaie,
     P5DEF,
     PGOutils2,
     LookUp,
     UTOF,
     PGOutils ;

Type
  TOF_PGMULCOMPTEURSDIF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    {$IFDEF EMANAGER}
    TypeUtilisat : String;
    {$ENDIF}
    procedure GrilleDblClick (Sender : TObject) ;
    procedure OnClickSalarieSortie(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    {$IFDEF EMANAGER}
    procedure SalarieElipsisClick(Sender : Tobject);
    procedure RespElipsisClick(Sender : TObject);
    {$ENDIF}
  end ;

Implementation


procedure TOF_PGMULCOMPTEURSDIF.GrilleDblClick (Sender : TObject) ;
var Salarie : String;
    Q_Mul : THQuery ;
begin
     {$IFDEF EAGLCLIENT}
     TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;  //PT1
     {$ENDIF}
     Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
     Salarie := Q_Mul.FindField('PSA_SALARIE').AsString;
     AGLLanceFiche('PAY','COMPTEURSDIF','','',Salarie + ';'+GetControlText('DATEARRET'));
end;

procedure TOF_PGMULCOMPTEURSDIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCOMPTEURSDIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCOMPTEURSDIF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCOMPTEURSDIF.OnLoad ;
var
DateArret : TdateTime;
StDateArret : string;
StWhere,WhereResp : String;
{$IFDEF EMANAGER}
MultiNiveau : Boolean;
{$ENDIF}
begin
  Inherited ;
  StWhere := '';
  If GetControlText('RESPONSFOR') <> '' then WhereResp := 'PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'" AND '
  else WhereResp := '';
  {$IFDEF EMANAGER}
  MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
//  StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")';
  If TypeUtilisat = 'R' then
  begin
       If MultiNiveau then StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+'(PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
         ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))))'
         else StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")';
  end
  else StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+'PSE_RESPONSFOR IN '+
            '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))))';
  {$ENDIF}
  if  TCheckBox(GetControl('CKSORTIE'))<>nil then
  Begin
  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
     Begin
     DateArret:=StrtoDate(GetControlText('DATEARRET'));
     StDateArret:='(PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
     StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
     If StWhere <> '' then StWhere := StWhere + ' AND '+ StDateArret
     else StWhere := StDateArret;
     End;
  end;
  SetControlText('XX_WHERE',StWhere);
end ;

procedure TOF_PGMULCOMPTEURSDIF.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        Defaut: THEdit;
        Check : TCheckBox;
        Num : Integer;
begin
  Inherited ;

        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        SetControlText('DATEARRET',DateToStr(V_PGI.DateEntree));
        Check:=TCheckBox(GetControl('CKSORTIE'));
        if Check=nil then
        Begin
        //Deb PT6
        SetControlVisible('DATEARRET',False);
        SetControlVisible('TDATEARRET',False);
        //FIN PT6
        End
        else
        begin
             Check.OnClick:=OnClickSalarieSortie;
             Check.Checked := True;
        end;

        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
         if Num >4 then Break;
          VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
        Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
        If Defaut<>nil then Defaut.OnExit:=ExitEdit;
        {$IFDEF EMANAGER}
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
       else TypeUtilisat := 'S';
        If Defaut<>nil then Defaut.OnElipsisClick := SalarieElipsisClick;
        SetControlVisible('RESPONSFOR',True);
        SetControlVisible('TRESPONSFOR',True);
        SetControlVisible('LIBRESP',True);
        SetControlCaption('LIBRESP','');
        Defaut:=ThEdit(getcontrol('RESPONSFOR'));
        If Defaut<>nil then Defaut.OnElipsisClick := RespElipsisClick;
        {$ENDIF}
end ;

procedure TOF_PGMULCOMPTEURSDIF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCOMPTEURSDIF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCOMPTEURSDIF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCOMPTEURSDIF.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGMULCOMPTEURSDIF.OnClickSalarieSortie(Sender: TObject);
begin
//DEB PT7
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
//FIN PT7
end;

{$IFDEF EMANAGER}
procedure TOF_PGMULCOMPTEURSDIF.SalarieElipsisClick(Sender : Tobject);
var StWhere,StOrder : String;
    DateArret : TDateTime; //PT2
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
     //If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     StOrder := 'PSA_SALARIE';
     StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
     LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;

procedure TOF_PGMULCOMPTEURSDIF.RespElipsisClick(Sender : TObject);
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
  registerclasses ( [ TOF_PGMULCOMPTEURSDIF ] ) ;
end.

