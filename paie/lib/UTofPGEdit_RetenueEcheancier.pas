{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 23/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EDITECHEANCIER ()
Mots clefs ... : TOF;PGEDITRETENUESAL
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
PT1 : 21/12/2006 FC V_70 FQ 13539 Zero à gauche matricule
PT2 : 30/01/2007 V_80 FC Mise en place filtrage des habilitations/poupulations
PT3 : 02/07/2007 FC V_72 FQ 14333 Confidentialité
}
Unit UTofPGEdit_RetenueEcheancier ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,Mul,
{$ELSE}
     eQRS1,eMul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,HQry,ParamSoc,
     PgOutils2,  //AffectDefautCode
     EntPaie,    //VH_paie
     P5Def   // MonHabilitation
     ;

Type
  TOF_PGEDITECHEANCIERRET = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure GestionPeriode (Sender : TObject);
    procedure DateElipsisClick (Sender : TObject);
    procedure ExitSalarie (Sender : TObject);
  end ;

Implementation

procedure TOF_PGEDITECHEANCIERRET.OnUpdate ;
var StWhere,StJoin,StOrder,StSelect : String;
    Pages : TPageControl;
    StConf : String; //PT3
begin
  Inherited ;
        StJoin := '';
        Pages := TPageControl(GetControl('Pages'));
        StWhere := RecupWhereCritere(Pages);
        StConf := SQLConf('SALARIES');    //PT3
        If (GetControlText('PSA_ETABLISSEMENT') <> '') or (GetCheckBoxState('CBALPHA') = CbChecked) or (Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'')) then //PT2
        begin
                StJoin :='LEFT JOIN SALARIES ON PHR_SALARIE=PSA_SALARIE ';
               // If GetControlText('ETAB') <> '' then StEtab := 'PSA_ETABLISSEMENT="'+GetControlText('ETAB')+'"';
        end;
        if GetCheckBoxState('CBALPHA') = CbChecked then StOrder := ' ORDER BY PSA_LIBELLE,PRE_SALARIE,PRE_ORDRE,PHR_DATEDEBUT'
        Else StOrder := ' ORDER BY PRE_SALARIE,PRE_ORDRE,PHR_DATEDEBUT';
        StSelect := 'SELECT PRE_SALARIE,PRE_ORDRE,PHR_SALARIE,PHR_MONTANT,PHR_ARRIERE,PHR_CUMULVERSE,PHR_MONTANTMENS,PHR_DATEPAIEMENT,PHR_ORDRE,PHR_DATEDEBUT,PHR_DATEFIN,PHR_REPRISEARR,'+
        'PRE_LIBELLE,PRE_RETENUESAL,PRE_NBMOIS,PRE_MONTANTMENS,PRE_MONTANTTOT,PRE_DATEDEBUT,PRE_DATEFIN,'+
        'ANN_NOMPER,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE,ANN_APPAYS '+
        'FROM RETENUESALAIRE LEFT JOIN HISTORETENUE ON PHR_SALARIE=PRE_SALARIE AND PHR_ORDRE=PRE_ORDRE '+
        'LEFT JOIN ANNUAIRE ON PRE_BENEFRSGU=ANN_GUIDPER ';
        If GetControlText('RETENUESALARIE') <> '' then
        begin
                If StWhere <> '' then
                  StWhere := StWhere +
                            'AND PHR_ORDRE='+GetcontrolText('RETENUESALARIE')+''  // DB2
                else
                  StWhere := 'WHERE PHR_ORDRE='+GetcontrolText('RETENUESALARIE')+''; // DB2
        end;
        If GetControlText('ETATRETENUE') = 'ACT' then
        begin
                If StWhere <> '' then StWhere := StWhere + 'AND PRE_ACTIF="X"'
                else StWhere := 'WHERE PRE_ACTIF="X"';
        end;
        If GetControlText('ETATRETENUE') = 'NON' then
        begin
                If StWhere <> '' then StWhere := StWhere + 'AND PRE_ACTIF="-"'
                else StWhere := 'WHERE PRE_ACTIF="-"';
        end;

        //DEB PT3
        If StConf <> '' then
        begin
          If StWhere <> '' then
            StWhere := StWhere + ' AND PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+')'
          else
            StWhere := 'WHERE PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+')';
        end;
        //FIN PT3

        TFQRS1(Ecran).WhereSQL := StSelect + StJoin + StWhere + StOrder;
end ;

procedure TOF_PGEDITECHEANCIERRET.OnArgument (S : String ) ;
var Check : TCheckBox;
    Edit : THEdit;
begin
  Inherited ;
        SetControlText('DATEHISTO',DateToStr(Date));
        SetControlEnabled('DATEHISTO',False);
        Check := TCheckBox(GetControl('CBHISTO'));
    //    If Check <> Nil then Check.OnClick := VisualiserHistorique;
        SetControlCaption('LIBSALARIE','');
        Edit := THEdit(GetControl('DATEHISTO'));
        If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit (GetControl('PHR_SALARIE'));
        If Edit <> Nil then Edit.OnExit := ExitSalarie;
        SetControlText('ETATRETENUE','ACT');
        SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));
        SetControlText('PSA_ETABLISSEMENT','');
end ;

procedure TOF_PGEDITECHEANCIERRET.GestionPeriode (Sender : TObject);
begin
        If GetCheckBoxState('CPERIODE')=CbChecked then SetControlEnabled('DATEDEBUT',True)
        Else SetControlEnabled('DATEDEBUT_',False);
end;

procedure TOF_PGEDITECHEANCIERRET.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITECHEANCIERRET.ExitSalarie (Sender : TObject);
var St : String;
  edit : thedit;

begin
        St := GetControlText('PHR_SALARIE');
        If St ='' then SetControlenabled('RETENUESALARIE',False)
        Else SetControlenabled('RETENUESALARIE',True);
        SetControlProperty('RETENUESALARIE','Plus',St);
        SetControlText('RETENUESALARIE','');
  // DEB PT1
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
 // FIN PT1
end;

Initialization
  registerclasses ( [ TOF_PGEDITECHEANCIERRET ] ) ;
end.

