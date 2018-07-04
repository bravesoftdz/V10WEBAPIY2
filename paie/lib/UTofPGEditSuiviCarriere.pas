{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 04/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITSUIVICARRIERE ()
Mots clefs ... : TOF;PGEDITSUIVICARRIERE
*****************************************************************}
Unit UTofPGEditSuiviCarriere ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,
{$else}
     eMul,uTob,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,ParamDat ,
     PGOutils,PGOutils2,PGEditOutils,PGEditOutils2,ParamSoc,P5DEF;

Type
  TOF_PGEDITSUIVICARRIERE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure ExitEdit(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure AccesDatePaie(Sender : TObject);
    Procedure AccesDateFormation(Sender : TObject);
    Procedure OnClickSalarieSortie(Sender : TObject);
  end ;

Implementation

procedure TOF_PGEDITSUIVICARRIERE.OnUpdate;
begin
  Inherited ;
        If GetCheckBoxState('CKSORTIE') = CbChecked then
        begin
                SetControlText('XX_WHERE','PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'"');
        end
        else SetControlText('XX_WHERE','');
end;

procedure TOF_PGEDITSUIVICARRIERE.OnArgument (S : String ) ;
var Edit : THEdit;
    Check : TCheckBox;
    Min,Max : String;
    Num : Integer;
begin
  Inherited ;
        SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
        Edit := THEdit(GetControl('DEBPAIE'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := THEdit(GetControl('FINPAIE'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := THEdit(GetControl('DEBFORM'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := THEdit(GetControl('FINFORM'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := ThEdit(getcontrol('PSA_SALARIE'));
        if Edit<>nil then Edit.OnExit:=ExitEdit;
        Edit := ThEdit(getcontrol('PSA_SALARIE_'));
        if Edit<>nil then Edit.OnExit:=ExitEdit;
        Check := TCheckBox(GetControl('CBPAIE'));
        If Check <> Nil then Check.OnClick := AccesDatePaie;
        Check := TCheckBox(GetControl('CBFORMATION'));
        If Check <> Nil then Check.OnClick := AccesDateFormation;
        RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
        SetControlText('PSA_SALARIE',Min);
        SetControlText('PSA_SALARIE_',Max);
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat(GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT'));
        Check := TCheckBox(GetControl('CKSORTIE'));
        If Check <> Nil then Check.OnClick := OnClickSalarieSortie;
end ;

procedure TOF_PGEDITSUIVICARRIERE.DateElipsisclick(Sender: TObject);
var
  key : char;
begin
  key := '*';
  ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITSUIVICARRIERE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
        edit:=THEdit(Sender);
        if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDITSUIVICARRIERE.AccesDatePaie(Sender : TObject);
begin
        If GetCheckBoxState('CBPAIE') = CbChecked then
        begin
                SetControlEnabled('DEBPAIE',True);
                SetControlEnabled('FINPAIE',True);
        end
        else
        begin
                SetControlEnabled('DEBPAIE',False);
                SetControlEnabled('FINPAIE',False);
        end;
end;

Procedure TOF_PGEDITSUIVICARRIERE.AccesDateFormation(Sender : TObject);
begin
        If GetCheckBoxState('CBFORMATION') = CbChecked then
        begin
                SetControlEnabled('DEBFORM',True);
                SetControlEnabled('FINFORM',True);
        end
        else
        begin
                SetControlEnabled('DEBFORM',False);
                SetControlEnabled('FINFORM',False);
        end;
end;

Procedure TOF_PGEDITSUIVICARRIERE.OnClickSalarieSortie(Sender : TObject);
begin
        SetControlenabled('DATESORTIE',(GetControltext('CKSORTIE')='X'));
        SetControlenabled('TDATESORTIE',(GetControltext('CKSORTIE')='X'));
end;

Initialization
  registerclasses ( [ TOF_PGEDITSUIVICARRIERE ] ) ;
end.

