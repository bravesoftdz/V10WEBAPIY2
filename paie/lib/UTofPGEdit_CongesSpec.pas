{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDIT_CONGESSPEC ()
Mots clefs ... : TOF;PGEDIT_CONGESSPEC
*****************************************************************}
{
PT1 26/01/2007  V_80 FC Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PT2 03/12/2007 V_72  RM FQ 14020 : Compléter le matricule salarié par des zeros automatiquement
}
Unit UTofPGEdit_CongesSpec ;

Interface

Uses StdCtrls,Controls,Classes, UTOB,
{$IFDEF EAGLCLIENT}
     eQRS1,UtilEAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtREtat,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,LookUp,HQry,HTB97,ParamDat,PGOutils,PgOutils2 ;

Type
  TOF_PGEDIT_CONGESSPEC = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure DateElipsisclick(Sender: TObject);
    procedure SalarieElipsisClick(Sender : TObject);
    Procedure ExitEdit(Sender: TObject);  //PT2
  end ;

Implementation

procedure TOF_PGEDIT_CONGESSPEC.OnUpdate ;
var StSQL,StWhere : String;
    Pages : TPageControl;
begin
  Inherited ;
        If GetControlText('ETAB') <> '' then
        begin
                Pages := TPageControl(GetControl('Pages'));
                StWhere := RecupWhereCritere(Pages);
                StSQL := 'SELECT CONGESSPEC.* FROM CONGESSPEC LEFT JOIN SALARIES ON PSA_SALARIE=PCS_SALARIE';
                If StWhere <> '' Then StWhere := ' '+StWhere+' AND PSA_ETABLISSEMENT="'+GetControlText('ETAB')+'"'
                Else StWhere := ' WHERE PSA_ETABLISSEMENT="'+GetControlText('ETAB')+'"';
                TFQRS1(Ecran).WhereSQL := StSQL+StWhere;
        end;
end ;

procedure TOF_PGEDIT_CONGESSPEC.OnArgument (S : String ) ;
var Edit : THEdit;
    aa,mm,jj,Mois : word;
    DateDebut,DateFin : TDateTime;
    Q :TQuery;
    Max : Integer;
begin
  Inherited ;
        Edit := THEdit(GetControl('PCS_DATEDEBUT'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('PCS_DATEDEBUT_'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        DecodeDate(V_PGI.DateEntree,aa,mm,jj);
        Mois:=1;
        If (mm<7) and (mm>3) Then Mois:=4;
        If (mm<10) and (mm>6) Then Mois:=7;
        If mm>9 Then Mois:=10;
        DateDebut:=EncodeDate(aa,Mois,1);
        DateFin:=PlusMois(DateDebut,2);
        DateFin:=FinDeMois(DateFin);
        SetControlText('PCS_DATEDEBUT',DateToStr(dateDebut));
        SetControlText('PCS_DATEDEBUT_',DateToStr(DateFin));
        SetControlCaption('LIBSAL','');
        Edit := THEdit(GetControl('PCS_SALARIE'));
        If Edit <> Nil then Edit.onElipsisClick := SalarieElipsisClick;
        If Edit <> Nil then Edit.OnExit := ExitEdit;    //PT2
        SetControlText('TDATEDELIV',dateToStr(V_PGI.DateEntree));
        Q := OpenSQL('SELECT MAX (PCS_NUMCERTIFICAT) MAXCERTIF FROM CONGESSPEC WHERE PCS_NUMCERTIFICAT<>""',True);
        If Not Q.Eof then
        begin
                If IsNumeric(Q.FindField('MAXCERTIF').AsString) then Max := StrToInt(Q.FindField('MAXCERTIF').AsString)
                else Max := 0;
        end
        Else Max := 0;
        Ferme(Q);
        Max := Max + 1;
        SetControlText('NUMCERTIF',ColleZeroDevant(Max,8));
end ;

procedure TOF_PGEDIT_CONGESSPEC.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDIT_CONGESSPEC.SalarieElipsisClick(Sender : TObject);
var  St : String;
begin
        St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
        'WHERE PSE_INTERMITTENT="X"';
        St := RecupClauseHabilitationLookupList(St);  //PT1
        LookupList(THEdit(Sender),'Liste des intermittents','SALARIES','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM','','', True,-1,St);
end;

//PT2 AJout de la procedure ===========================>
Procedure TOF_PGEDIT_CONGESSPEC.ExitEdit(Sender: TObject);
var edit : thedit;
Begin
edit:=THEdit(Sender);
If edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
End;

Initialization
  registerclasses ( [ TOF_PGEDIT_CONGESSPEC ] ) ;
end.

