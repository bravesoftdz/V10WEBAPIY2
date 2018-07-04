{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 11/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EDITFRAISGENERAUX ()
Mots clefs ... : TOF;EDITFRAISGENERAUX
*****************************************************************
PT1 09/12/2004 JL V_60 Suppression de la requete dans éxé
PT2 12/01/2005 JL V_60 Cumul 09 net imposable par défaut
PT3 08/02/2005 JL V_60 FQ 11938 Valeur de PHC_CUMUL forcé à 09 si non renseigné
PT4 11/05/2007 JL V_72 FQ 14011 cumul 01 au lieu de 09
}
Unit UTofPGEditFraisGeneraux ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,QRS1,
{$else}
     eMul,uTob,eQRS1,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,Paramdat,ParamSoc,HQRY,PGEditOutils,EntPaie,PGOutils,PGOutils2 ;

Type
  TOF_EDITFRAISGENERAUX = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure DateElipsisClick(Sender : TOBject);
    procedure ExitEdit(Sender : Tobject);
    procedure ExitCumul (Sender : TObject);
  end ;

Implementation

procedure TOF_EDITFRAISGENERAUX.OnUpdate ;
var SQL,Where : String;
    Pages : TPageControl;
begin
  Inherited ;
        {Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        If Where <> '' then Where := Where+' ';
        SQL := 'select  top 10 sum(phc_montant) PHC_MONTANT,PHC_SALARIE,PSA_SEXE,PSA_LIBELLEEMPLOI,'+
        '(Select SUM (PHB_MTREM) FROM HISTOBULLETIN LEFT JOIN REMUNERATION ON PHB_RUBRIQUE=PRM_RUBRIQUE '+
        'WHERE PHB_SALARIE=PHC_SALARIE  AND PHB_NATURERUB="AAA" AND PHB_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PHC_DATEFIN_')))+'" '+
        'AND PHB_DATEFIN>="'+UsDateTime(StrToDate(GetControlText('PHC_DATEFIN')))+'" AND PRM_AVANTAGENAT<>"") AVTNATURE, '+
        '(Select SUM (PHB_MTREM) FROM HISTOBULLETIn LEFT JOIN REMUNERATION ON PHB_RUBRIQUE=PRM_RUBRIQUE '+
        'WHERE PHB_SALARIE=PHC_SALARIE  AND PHB_NATURERUB="AAA" AND PHB_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PHC_DATEFIN_')))+'" '+
        'AND PHB_DATEFIN>="'+UsDateTime(StrToDate(GetControlText('PHC_DATEFIN')))+'" AND PRM_FRAISPROF="X") FRAISPROF, '+
        'PSA_ADRESSE1,PSA_CODEPOSTAL,PSA_VILLE '+
        'from histocumsal '+
        'left join SALARIES ON PSA_SALARIE=PHC_SALARIE '+Where+
        'group by PHC_SALARIE,PSA_SEXE,PSA_LIBELLEEMPLOI,PSA_ADRESSE1,PSA_CODEPOSTAL,PSA_VILLE ORDER BY PHC_MONTANt DESC';
        TFQRS1(Ecran).WhereSQL:=SQL;}
end ;

procedure TOF_EDITFRAISGENERAUX.OnArgument (S : String ) ;
var Edit,ECumul : THEdit;
begin
  Inherited ;
     SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
     Edit := ThEdit(getcontrol('PHC_SALARIE'));
     If Edit <> nil then Edit.OnExit:=ExitEdit;
     Edit := ThEdit(getcontrol('PHC_SALARIE_'));
     If Edit <> nil then Edit.OnExit:=ExitEdit;
     Edit := ThEdit(getcontrol('PHC_DATEFIN'));
     If Edit <> nil then Edit.OnElipsisClick := DateElipsisClick;
     Edit := ThEdit(getcontrol('PHC_DATEFIN_'));
     If Edit <> nil then Edit.OnElipsisClick := DateElipsisClick;
     SetControlText('PHC_CUMULPAIE','01');           // PT2 PT4
     SetControlCaption('LIBCUMUL',RechDom('PGCUMULPAIE','09',False)); // PT2
     ECumul := THEdit(GetControl('PHC_CUMULPAIE'));
     If ECumul <> Nil then ECumul.OnExit := ExitCumul;
end ;

procedure TOF_EDITFRAISGENERAUX.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_EDITFRAISGENERAUX.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

//DEBUT PT3
procedure TOF_EDITFRAISGENERAUX.ExitCumul (Sender : TObject);
begin
     If GetControlText('PHC_CUMULPAIE') = '' then
     begin
          SetControlText('PHC_CUMULPAIE','01'); //PT4
     end;
end;
//FIN PT3

Initialization
  registerclasses ( [ TOF_EDITFRAISGENERAUX ] ) ;
end.

