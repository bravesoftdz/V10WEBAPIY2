{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGRAZFORMATION ()
Mots clefs ... : TOF;PGRAZFORMATION
*****************************************************************}
Unit UTofPGRazFormation ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,HTB97,UTOB ;

Type
  TOF_PGRAZFORMATION = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure EffacerDonneesMillesime(Sender : TObject);
    procedure DateElipsisclick(Sender : TObject);
    procedure ChangeMillesime (Sender : TObject);
  end ;

Implementation

procedure TOF_PGRAZFORMATION.OnArgument (S : String ) ;
var BEfface : TToolBarButton97;
    Q : TQuery;
    Edit : THEdit;
    Combo : THValComboBox;
begin
  Inherited ;
        BEfface := TToolBarButton97(GetControl('BEFFACE'));
        If BEfface <> Nil Then BEfface.OnClick := EffacerDonneesMillesime;
        Q := OpenSQL('SELECT PFE_MILLESIME,PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION ORDER BY PFE_DATEDEBUT DESC',True);
        If not Q.Eof then
        begin
                SetControlText('PFE_MILLESIME',Q.FindField('PFE_MILLESIME').AsString);
                SetControlText('PFE_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                SetControlText('PFE_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
        Edit := THEdit(GetControl('PFE_DATEDEBUT'));
        if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := THEdit(GetControl('PFE_DATEFIN'));
        if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Combo := THValComboBox(GetControl('PFE_MILLESIME'));
        If Combo <> Nil Then Combo.OnChange := ChangeMillesime;
end ;

procedure TOF_PGRAZFORMATION.EffacerDonneesMillesime (Sender : TObject);
var Q : TQuery;
    TobRealise : Tob;
    i : Integer;
begin
        If PGIAsk('Etes vous sûr de vouloir effacer toutes les données de cet exercice ?',Ecran.Caption) = MrYes then
        begin
                //Supp réalisé
                If GetControlText('REALISE') = 'X' then
                begin
                        Q := OpenSQL('SELECT * FROM SESSIONSTAGE WHERE PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('PFE_DATEDEBUT')))+'" '+
                        'AND PSS_DATEDEBUT<="'+UsDateTime(StrToDate(GetControlText('PFE_DATEFIN')))+'"',true);
                        TobRealise := Tob.create('SESSIONSTAGE',Nil,-1);
                        TobRealise.LoadDetailDB('SESSIONSTAGE','','',Q,False);
                        Ferme(Q);
                        For i := 0 to TobRealise.detail.Count-1 do
                        begin
                                //Supp des formations
                                ExecuteSQL('DELETE FROM FORMATIONS WHERE PFO_CODESTAGE="'+TobRealise.Detail[i].GetValue('PSS_CODESTAGE')+'" '+
                                'AND PFO_ORDRE='+IntToStr(TobRealise.Detail[i].GetValue('PSS_ORDRE'))+' AND PFO_MILLESIME="'+TobRealise.Detail[i].GetValue('PSS_MILLESIME')+'"');                        //Supp des frais//DB2
                                ExecuteSQL('DELETE FROM FRAISSALFORM WHERE PFS_CODESTAGE="'+TobRealise.Detail[i].GetValue('PSS_CODESTAGE')+'" '+
                                'AND PFS_ORDRE='+IntToStr(TobRealise.Detail[i].GetValue('PSS_ORDRE'))+' AND PFS_MILLESIME="'+TobRealise.Detail[i].GetValue('PSS_MILLESIME')+'"');   //DB2
                                //Supp des animateurs
                                ExecuteSQL('DELETE FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+TobRealise.Detail[i].GetValue('PSS_CODESTAGE')+'" '+
                                'AND PAN_ORDRE='+IntToStr(TobRealise.Detail[i].GetValue('PSS_ORDRE'))+' AND PAN_MILLESIME="'+TobRealise.Detail[i].GetValue('PSS_MILLESIME')+'"'); //DB2
                                //Supp des investissements
                                ExecuteSQL('DELETE FROM INVESTSESSION WHERE PIS_CODESTAGE="'+TobRealise.Detail[i].GetValue('PSS_CODESTAGE')+'" '+
                                'AND PIS_ORDRE='+IntToStr(TobRealise.Detail[i].GetValue('PSS_ORDRE'))+' AND PIS_MILLESIME="'+TobRealise.Detail[i].GetValue('PSS_MILLESIME')+'"');  //DB2
                                //Supp des sessions
                                TobRealise.Detail[i].DeleteDB(False);
                        end;
                        TobRealise.free;
                end;
                //Supp des inscriptions au budget
                If GetControlText('PREVISIONNEL') = 'X' then
                begin
                        ExecuteSQL('DELETE FROM STAGE WHERE PST_MILLESIME="'+GetControlText('PFE_MILLESIME')+'"');
                        //Supp des stage
                        ExecuteSQL('DELETE FROM INSCFORMATION WHERE PFI_MILLESIME="'+GetControlText('PFE_MILLESIME')+'"');
                        //Cursus
                        ExecuteSQL('DELETE FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+GetControlText('PFE_MILLESIME')+'"');
                end;
        end;
end;

procedure TOF_PGRAZFORMATION.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGRAZFORMATION.ChangeMillesime (Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+GetControlText('PFE_MILLESIME')+'"',True);
        If not Q.Eof then
        begin
                SetControlText('PFE_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                SetControlText('PFE_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;


Initialization
  registerclasses ( [ TOF_PGRAZFORMATION ] ) ;
end.

