{***********UNITE*************************************************
Auteur  ...... :  JL
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULHISTORETENUE ()
Mots clefs ... : TOF;PGMULHISTORETENUE
*****************************************************************}
{
PT1 : 21/12/2006 FC V_70 FQ 13539 Zero à gauche matricule
PT2 : 02/07/2007 FC V_72 FQ 14333 Confidentialité
PT3 : 02/07/2008 FC V_850 FQ 15604 "Incorrect syntaxe near the keyword 'AND' en consultation
}
Unit UTofPG_MulHistoRetenue;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,Mul,
     {$ELSE}
     MainEAGL,eMul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,paramdat,UTob,HTB97,
     PgOutils2,  //AffectDefautCode
     EntPaie    //VH_paie
     ;

Type
  TOF_PGMULHISTORETENUE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Salarie,Ordre : String;
    procedure DateElipsisClick (Sender : TObject);
    procedure ExitSalarie (Sender : TObject);
    procedure DebutHisto (Sender : TObject);
    procedure FinHisto (Sender : TObject);
    procedure ClickMainLevee(Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULHISTORETENUE.OnLoad ;
var StWhere : String;
  StConf : String;  //PT2
begin
  Inherited ;
        StWhere := '';
        StConf := SQLConf('SALARIES');  //PT2
        If GetControlText('RETENUESAL') <> '' then
          StWhere := 'PHR_ORDRE='+GetControlText('RETENUESAL')+'';      // DB2
        If GetCheckBoxState('CBDEBUTHISTO') = CbChecked then
        begin
                If StWhere <> '' then StWhere := StWhere + ' AND ';
                StWhere := StWhere + 'PHR_DATEDEBUT>="'+UsdateTime(StrToDate(GetControlText('DATEDEBUT')))+'"';
        end;
        If GetCheckBoxState('CBFINHISTO') = CbChecked then
        begin
                If StWhere <> '' then StWhere := StWhere + ' AND ';
                StWhere := StWhere + 'PHR_DATEFIN<="'+UsdateTime(StrToDate(GetControlText('DATEFIN')))+'"';
        end;
        //DEB PT2
        If StConf <> '' then
        begin
          If StWhere <> '' then
            StWhere := StWhere + ' AND PHR_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+')'
          else
            StWhere := 'PHR_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+')';
        end;
        //FIN PT2
        SetControlText('XX_WHERE',StWhere);
        TFMul(Ecran).FListe.OnDblClick := GrilleDblClick;
end ;

procedure TOF_PGMULHISTORETENUE.OnArgument (S : String ) ;
var Edit:THEdit;
    Check : TCheckBox;
    Q : TQuery;
    Bt : TToolBarButton97;
begin
Inherited ;
        Salarie := ReadTokenPipe (S,';');
        Ordre := ReadTokenPipe (S,';');
        SetControlCaption('LIBSAL','');
        If (Salarie <> '') and (ordre <> '') then
        begin
                SetControlText('PHR_SALARIE',Salarie);
                SetControlEnabled('PHR_SALARIE',False);
                SetControlProperty('RETENUESAL','Plus',Salarie);
                SetControlText('RETENUESAL',Ordre);
                Ecran.Caption := 'Retenue salarié '+Salarie + ' ' + rechDom('PGSALARIE',Salarie,False);
                UpdateCaption(Ecran);
        end
        Else SetControlEnabled('RETENUESAL',False);
        Edit := THEdit(GetControl('DATEDEBUT'));
        If Edit <> Nil Then Edit.OnClick := DateElipsisClick;
        Edit := THEdit(GetControl('DATEFIN'));
        If Edit <> Nil Then Edit.OnClick := DateElipsisClick;
        Edit := THEdit (GetControl('PHR_SALARIE'));
        If Edit <> Nil then Edit.OnExit := ExitSalarie;
        SetControlText('DATEDEBUT',DateToStr(Date));
        SetControlText('DATEFIN',DateToStr(Date));
        Check := TCheckBox(Getcontrol('CBDEBUTHISTO'));
        If Check <> Nil then Check.OnClick := DebutHisto;
        Check := TCheckBox(Getcontrol('CBFINHISTO'));
        If Check <> Nil then Check.OnClick := FinHisto;
        SetControlProperty('PAVANCE','TabVisible',False);
        SetControlProperty('PCOMPLEMENT','TabVisible',False);
        SetControlVisible('BINSERT',False);
        SetControlVisible('BOUVRIR',False);
        Q := OpenSQL('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_DATEDEBUT DESC',True);
        if not Q.eof then
        begin
                SetControlText('DATEDEBUT',DateToStr(Q.FindField('PEX_DATEDEBUT').AsDateTime));
                SetControlChecked('CBDEBUTHISTO',True);
                SetControlText('DATEFIN',DateToStr(Date));
                SetControlChecked('CFINHISTO',True);
        end;
        Ferme(Q);
        Bt := TToolBarButton97(GetControl('BMAINLEVEE'));
        If Bt <> Nil then Bt.OnClick := ClickMainLevee;
end ;

procedure TOF_PGMULHISTORETENUE.DateElipsisClick(Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULHISTORETENUE.ExitSalarie (Sender : TObject);
var St : String;
  edit : thedit;

begin
        St := GetControlText('PHR_SALARIE');
        If St ='' then SetControlenabled('RETENUESAL',False)
        Else SetControlenabled('RETENUESAL',True);
        SetControlProperty('RETENUESAL','Plus',St);
        SetControlText('RETENUESAL','');
  //DEB PT1
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
 // FIN PT1     
end;

procedure TOF_PGMULHISTORETENUE.DebutHisto (Sender : TObject);
begin
        If (TCheckBox(Sender).Checked) = True then SetControlEnabled('DATEDEBUT',True)
        Else SetControlEnabled('DATEDEBUT',False);
end;

procedure TOF_PGMULHISTORETENUE.FinHisto (Sender : TObject);
begin
        If (TCheckBox(Sender).Checked) = True then SetControlEnabled('DATEFIN',True)
        Else SetControlEnabled('DATEFIN',False);
end;

procedure TOF_PGMULHISTORETENUE.ClickMainLevee(Sender : TObject);
begin
  AGLLanceFiche('PAY','RETENUEMAINLEVEE','','',Salarie+';'+Ordre+';ACTION=CREATION');
end;

procedure TOF_PGMULHISTORETENUE.GrilleDblClick(Sender : TObject);
var DD,DF : TDateTime;
begin
  //DEB PT3 Je ne vois pas comment ça a pu fonctionner sans qu'on récupère les valeurs de la ligne
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
  {$ENDIF}
  Salarie := TFMul(Ecran).Q.FindField('PHR_SALARIE').AsString;
  Ordre := TFMul(Ecran).Q.FindField('PHR_ORDRE').AsString;
  //FIN PT3
  DD := TFMul(Ecran).Q.FindField('PHR_DATEDEBUT').AsDateTime;
  DF := TFMul(Ecran).Q.FindField('PHR_DATEFIN').AsDateTime;
  if (Salarie <> '') and (Ordre <> '') then   //PT3
    AGLLanceFiche('PAY','RETENUEMAINLEVEE','',Salarie+';'+Ordre+';'+DateToStr(DD)+';'+DateToStr(DF),'');
end;


Initialization
  registerclasses ( [ TOF_PGMULHISTORETENUE ] ) ;
end.
