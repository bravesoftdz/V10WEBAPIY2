{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCURSUSSESSION ()
Mots clefs ... : TOF;PGCURSUSSESSION
*****************************************************************
PT1   : 09/12/2003 V_50 JL Corection : zone rangcursus n'était pas MAJ après création
---- JL 20/03/2006 modification clé annuaire ----
}
Unit UTofPGCursusSession ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
     MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTob,HSysMenu,HTB97,ParamDat ;

Type
  TOF_PGCURSUSSESSION = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    LeCursus,Action : String;
    LeRang : Integer;
    procedure RemplirGrilleSession;
    procedure ChangeCursus(Sender : TOBject);
    procedure Validation(Sender : TObject);
    procedure OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleElipsisClick (Sender : TObject);
    procedure InscrptionStagiaire(Sender : TObject);
    procedure SupprimerCursus (Sender : TObject);
  end ;

Implementation

procedure TOF_PGCURSUSSESSION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGCURSUSSESSION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGCURSUSSESSION.OnArgument (S : String ) ;
var Combo : THValComboBox;
    BValid,BStagiaire,BDelete : TToolBarButton97;
    GSessions : THGrid;
begin
  Inherited ;
        Action := ReadTokenPipe(S,';');
        If Action = 'MODIFICATION' then
        begin
                LeCursus := ReadTokenPipe (S,';');
                If LeCursus <> '' then
                begin
                        SetControlText('CURSUS',LeCursus);
                        LeRang := StrToInt(ReadTokenPipe(S,';'));
                        SetControlEnabled('CURSUS',False);
                        SetControlText('RANGCURSUS',IntToStr(LeRang));
                        RemplirGrilleSession;
                        SetControlVisible('BSTAGIAIRE',True);
                end
                Else LeRang := 0;
        end
        else LeRang := 0;
        If LeRang = 0 then SetControlVisible('BSTAGIAIRE',False);
        Combo := THValComboBox(GetControl('CURSUS'));
        If Combo <> Nil then Combo.OnChange := ChangeCursus;
        BValid := TToolBarButton97(GetControl('BVAL'));
        If BValid <> Nil then BValid.OnClick := Validation;
        GSessions := THGrid(GetControl('GSESSIONS'));
        If GSessions <> Nil then
        begin
                GSessions.OnCellEnter := OnCellEnter;
                GSessions.OnCellExit := OnCellExit;
                GSessions.OnElipsisClick := GrilleElipsisClick;
                GSessions.ColTypes[2]:='D';
                GSessions.ColTypes[3]:='D';
                GSessions.ColFormats[2] := ShortDateFormat;
                GSessions.ColFormats[3] := ShortDateFormat;
                GSessions.ColEditables[0]:=False;
                GSessions.ColEditables[1]:=False;
                GSessions.ColEditables[4]:=False;
                GSessions.ColEditables[5]:=False;
        end;
        BStagiaire := TToolBarButton97(GetControl('BSTAGIAIRE'));
        If BStagiaire <> Nil then BStagiaire.OnClick := InscrptionStagiaire;
        BDelete := TToolBarButton97(GetControl('BDELETE'));
        If BDelete <> Nil then BDelete.OnClick := SupprimerCursus;
end ;

procedure TOF_PGCURSUSSESSION.RemplirGrilleSession;
var GSessions : THGrid;
    HMTrad : THSystemMenu;
    Q : TQuery;
    TobFormations : Tob;
    i,Session : Integer;
    Stage : String;
begin
        GSessions := THGrid(GetControl('GSESSIONS'));
        If GSessions = Nil then Exit;
//        If LeRang > 0 then
        //SetControlEnabled('GSESSIONS',False);
        Q := OpenSQL('SELECT PCU_LIBELLE,PCU_DATEDEBUT,PCU_DATEFIN FROM CURSUS WHERE PCU_CURSUS="'+LeCursus+'"'+
        ' AND PCU_RANGCURSUS='+IntToStr(LeRang)+'',true);            //DB2
        if Not Q.eof then
        begin
                SetControlText('DATEDEB',DateToStr(Q.FindField('PCU_DATEDEBUT').AsDateTime));
                SetControlText('DATEFIN',DateToStr(Q.FindField('PCU_DATEFIN').AsDateTime));
                SetControlText('LIBELLE',Q.FindField('PCU_LIBELLE').AsString);
        end;
        Ferme(Q);
        Q := OpenSQL('SELECT PCC_CODESTAGE,PCC_ORDRE FROM CURSUSSTAGE WHERE PCC_RANGCURSUS='+IntToStr(LeRang)+' '+  //DB2
        'AND PCC_CURSUS="'+LeCursus+'"',True);
        TobFormations := Tob.Create('Les formations',Nil,-1);
        TobFormations.LoadDetailDB('Les Formations','','',Q,False);
        Ferme(Q);
        For i := 1 to GSessions.RowCount - 1 do
        begin
                GSessions.Rows[i].Clear;
        end;
        GSessions.RowCount := 2;
        For i := 0 to TobFormations.Detail.Count -1 do
        begin
                Stage := TobFormations.Detail[i].GetValue('PCC_CODESTAGE');
                Session := TobFormations.Detail[i].GetValue('PCC_ORDRE');
                If i > 0 then GSessions.RowCount := GSessions.RowCount + 1;
                GSessions.CellValues[0,i+1] := Stage;
                GSessions.CellValues[1,i+1] := RechDom('PGSTAGEFORM',Stage,False);
                If LeRang > 0 then
                begin
                        Q := OpenSQL('SELECT PSS_DATEDEBUT,PSS_DATEFIN FROM SESSIONSTAGE'+
                        ' WHERE PSS_CODESTAGE="'+Stage+'" AND PSS_ORDRE='+IntToStr(Session)+'',True);  //DB2
                        If Not Q.eof then
                        begin
                                GSessions.CellValues[2,i+1] := DateToStr(Q.FindField('PSS_DATEDEBUT').AsDateTime);
                                GSessions.CellValues[3,i+1] := DateToStr(Q.FindField('PSS_DATEFIN').AsDateTime);
                        end;
                        Ferme(Q);
                end;
                GSessions.cellValues[4,i+1] := IntToStr(Session);
                GSessions.CellValues[5,i+1] := '0000';
        end;
        TobFormations.free;
        HMTrad:=THSystemMenu(GetControl('HMTrad'));
        HMTrad.ResizeGridColumns (GSessions);
end;

procedure TOF_PGCURSUSSESSION.ChangeCursus(Sender : TOBject);
begin
        LeCursus := GetControlText('CURSUS');
        RemplirGrilleSession;
end;

procedure TOF_PGCURSUSSESSION.Validation(Sender : TObject);
var TobSessions,TS,TobCursusStage,TCS,TobCursus,TC,TobInvest,TI : Tob;
    Q : TQuery;
    IMax,i : Integer;
    GSessions : THGrid;
    Stage,St,TestDate : String;
begin
        GSessions := THGrid(GeTControl('GSESSIONS'));
        If GSessions = Nil then Exit;
        For i := 1 to GSessions.RowCount -1 do
        begin
                TestDate := GSessions.CellValues[2,i];
                If Not IsValidDate (TestDate) then
                begin
                        PGIBox(TestDate+' n''est pas une date correcte, vous devez la modifier',Ecran.Caption);
                        GSessions.Row := i;
                        GSessions.Col := 2;
                        exit;
                end;
                TestDate := GSessions.CellValues[3,i];
                If Not IsValidDate (TestDate) then
                begin
                        PGIBox(TestDate+' n''est pas une date correcte, vous devez la modifier',Ecran.Caption);
                        GSessions.Row := i;
                        GSessions.Col := 2;
                        exit;
                end;
        end;
        Q := OpenSQL('SELECT MAX(PCU_RANGCURSUS) FROM CURSUS WHERE PCU_CURSUS="'+LeCursus+'"',True);
        if Not Q.EOF then
                begin
                        IMax := Q.Fields[0].AsInteger;
                        if IMax <> 0 then
                        IMax := IMax + 1
                        else
                        IMax := 1;
                end
                else IMax := 1 ;
        Ferme(Q);
        LeRang := IMax;
        SetControlText('RANGCURSUS',IntToStr(LeRang));  //PT1
        TobCursusStage := Tob.Create('Les cursus et stage',Nil,-1);
        TobSessions := Tob.Create('Les sessions',Nil,-1);
        For i := 1 to GSessions.RowCount -1 do
        begin
                Stage := GSessions.CellValues[0,i];
                If Stage = '' then continue;
                Q := OpenSQL('SELECT MAX(PSS_ORDRE) FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Stage+'"',True);
                if Not Q.EOF then
                begin
                        IMax := Q.Fields[0].AsInteger;
                        if IMax <> 0 then
                        IMax := IMax + 1
                        else
                        IMax := 1;
                end
                else IMax := 1 ;
                Ferme(Q) ;
                TCS := Tob.Create('CURSUSSTAGE',TobCursusStage,-1);
                TCS.PutValue('PCC_RANGCURSUS',Lerang);
                TCS.PutValue('PCC_CURSUS',LeCursus);
                TCS.PutValue('PCC_CODESTAGE',Stage);
                TCS.PutValue('PCC_ORDRE',IMax);
                TCS.InsertDB(Nil,False);
                TS := Tob.Create('SESSIONSTAGE',TobSessions,-1);
                TS.PutValue('PSS_CODESTAGE',Stage);
                TS.PutValue('PSS_ORDRE',IMax);
                TS.PutValue('PSS_MILLESIME','0000');
                TS.PutValue('PSS_DATEDEBUT',StrToDate(GSessions.CellValues[2,i]));
                If (StrToDate(GetControlText('DATEDEB')) = IDate1900) or (StrToDate(GetControlText('DATEDEB')) > StrToDate(GSessions.CellValues[2,i])) then
                SetControlText('DATEDEB',GSessions.CellValues[2,i]);
                TS.PutValue('PSS_DATEFIN',StrToDate(GSessions.CellValues[3,i]));
                If (StrToDate(GetControlText('DATEFIN')) = IDate1900) or (StrToDate(GetControlText('DATEFIN')) < StrToDate(GSessions.CellValues[3,i])) then
                SetControlText('DATEFIN',GSessions.CellValues[3,i]);
                st  :=  'SELECT PST_COUTBUDGETE,PST_COUTFONCT,PST_COUTORGAN,PST_COUTUNITAIRE,PST_COUTMOB,PST_DUREESTAGE,PST_JOURSTAGE,PST_INCLUSDECL,PST_ACTIONFORM,'+
                'PST_JOURSTAGE,PST_CENTREFORMGU,PST_LIEUFORM,PST_NATUREFORM,PST_TYPCONVFORM,PST_ORGCOLLECTPGU,PST_ORGCOLLECTSGU'+
                ',PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8'+
                ' FROM STAGE WHERE PST_CODESTAGE="'+Stage+'" AND PST_MILLESIME="0000"';
                Q := OpenSQL(st,TRUE) ;
                if Not Q.EOF then
                begin
                        TS.PutValue ('PSS_CENTREFORMGU',Q.FindField('PST_CENTREFORMGU').AsString);
                        TS.PutValue('PSS_LIEUFORM',Q.FindField('PST_LIEUFORM').AsString);
                        TS.PutValue('PSS_ORGCOLLECTPGU',Q.FindField('PST_ORGCOLLECTPGU').AsString);
                        TS.PutValue('PSS_ORGCOLLECTSGU',Q.FindField('PST_ORGCOLLECTSGU').AsString);
                        TS.PutValue('PSS_NATUREFORM',Q.FindField('PST_NATUREFORM').AsString);
                        TS.PutValue('PSS_TYPCONVFORM',Q.FindField('PST_TYPCONVFORM').AsString);
                        TS.PutValue('PSS_COUTPEDAG',Q.FindField('PST_COUTBUDGETE').AsFloat);
                        TS.PutValue('PSS_COUTFONCT',Q.FindField('PST_COUTFONCT').AsFloat);
                        TS.PutValue('PSS_COUTORGAN',Q.FindField('PST_COUTORGAN').AsFloat);
                        TS.PutValue('PSS_COUTUNITAIRE',Q.FindField('PST_COUTUNITAIRE').AsFloat);
                        TS.PutValue('PSS_COUTMOB',Q.FindField('PST_COUTMOB').AsFloat);
                        TS.PutValue('PSS_FORMATION1',Q.FindField('PST_FORMATION1').AsString);
                        TS.PutValue('PSS_FORMATION2',Q.FindField('PST_FORMATION2').AsString);
                        TS.PutValue('PSS_FORMATION3',Q.FindField('PST_FORMATION3').AsString);
                        TS.PutValue('PSS_FORMATION4',Q.FindField('PST_FORMATION4').AsString);
                        TS.PutValue('PSS_FORMATION5',Q.FindField('PST_FORMATION5').AsString);
                        TS.PutValue('PSS_FORMATION6',Q.FindField('PST_FORMATION6').AsString);
                        TS.PutValue('PSS_FORMATION7',Q.FindField('PST_FORMATION7').AsString);
                        TS.PutValue('PSS_FORMATION8',Q.FindField('PST_FORMATION8').AsString);
                        TS.PutValue('PSS_DUREESTAGE',Q.FindField('PST_DUREESTAGE').Asfloat);
                        TS.PutValue('PSS_JOURSTAGE',Q.FindField('PST_JOURSTAGE').Asfloat);
                        TS.PutValue('PSS_INCLUSDECL',Q.FindField('PST_INCLUSDECL').AsString);
                        TS.PutValue('PSS_ACTIONFORM',Q.FindField('PST_ACTIONFORM').AsString);
                end;
                Ferme (Q);
                TS.PutValue ('PSS_DEBUTDJ','MAT');
                TS.PutValue ('PSS_FINDJ','PAM');
                TS.PutValue('PSS_EFFECTUE','-');
                TS.PutValue('PSS_CLOTUREINSC','-');
                TS.PutValue('PSS_VALIDORG','-');
                TS.PutValue('PSS_HEUREDEBUT',0);
                TS.PutValue('PSS_HEUREFIN',0);
                TS.PutValue('PSS_AVECCLIENT','-');
                TS.InsertDB(Nil,False);
                Q := OpenSQL('SELECT * FROM INVESTSESSION '+
                'WHERE PIS_CODESTAGE="'+Stage+'" AND PIS_ORDRE=-1 AND PIS_MILLESIME="0000"',True);
                TobInvest := Tob.Create('Les investissements',Nil,-1);
                TobInvest.LoadDetailDB('INVESTSESSION','','',Q,False);
                Ferme(Q);
                TI := TobInvest.FindFirst(['PIS_CODESTAGE'],[Stage],False);
                While TI <> Nil do
                begin
                        TI.ChangeParent(TobInvest,-1);
                        TI.PutValue('PIS_ORDRE',IMax);
                        TI.InsertOrUpdateDB(False);
                        TI := TobInvest.FindNext(['PIS_CODESTAGE'],[Stage],False);
                end;
               TobInvest.Free;
        end;
        TobCursus := Tob.Create('les cursus',Nil,-1);
        TC := Tob.Create('CURSUS',TobCursus,-1);
        TC.PutValue('PCU_CURSUS',LeCursus);
        TC.PutValue('PCU_RANGCURSUS',LeRang);
        TC.PutValue('PCU_LIBELLE',GetControlText('LIBELLE'));
        TC.PutValue('PCU_DATEDEBUT',StrToDate(GetControlText('DATEDEB')));
        TC.PutValue('PCU_DATEFIN',StrToDate(GetControlText('DATEFIN')));
        TC.InsertDB(Nil,False);
        TobCursus.Free;
        TobSessions.Free;
        TobCursusStage.Free;
        If ACtion = 'CREATION' then
        begin
                Action := 'MODIFICATION';
                SetControlVisible('BSTAGIAIRE',True);
        end;
        RemplirGrilleSession;
end;

procedure TOF_PGCURSUSSESSION.OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
        If (ACol = 2) or (ACol = 3) then THGrid(Sender).ElipsisButton := True
        Else THGrid(Sender).ElipsisButton := False;
        If (ACol = 0) or (ACol = 1) then THGrid(Sender).Col := 3;
end;

procedure TOF_PGCURSUSSESSION.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var Q : TQuery;
    DureeStage : Double;
    DateDeb,DateFin : TDateTime;
    GSessions : THGrid;
    Stage,TestDate : String;
begin
        GSessions := THGrid(GetControl('GSESSIONS'));
        If ACol = 2 then
        begin
                TestDate := GSessions.CellValues[2,Arow];
                If GSessions.CellValues[0,Arow] = '' then Exit;
                If not IsValidDate(TestDate) then
                begin
                        If not (IsValidDate(TestDate)) and (TestDate <> '  /  /    ') and (TestDate<>'') then  //PT1
                        begin
                                PGIBox(TestDate+' n''est pas une date correcte',Ecran.Caption);
                                GSessions.Col := ACol;
                                GSessions.Row := ARow;
                                Exit;
                        end;
                        Exit;
                end;
                Stage := GSessions.CellValues[0,ARow];
                DureeStage := 0;
                Q := OpenSQL('SELECT PST_JOURSTAGE FROM STAGE WHERE PST_CODESTAGE="'+Stage+'" AND PST_MILLESIME="0000"',True);
                If Not Q.Eof then
                begin
                        DureeStage := Q.FindField('PST_JOURSTAGE').AsFloat;
                end;
                Ferme(Q);
                DateDeb := StrToDate(TestDate);
                If DureeStage>0 Then DateFin  :=  PLUSDATE (DateDeb,Trunc(DureeStage-1),'J')
                Else DateFin := DateDeb;
                GSessions.CellValues[3,ARow] := DateToStr(DateFin);
        end;
        If ACol = 3 then
        begin
        TestDate := GSessions.CellValues[2,Arow];
        If not (IsValidDate(TestDate)) and (TestDate <> '  /  /    ') and (TestDate<>'') then  //PT1
                begin
                        PGIBox(TestDate+' n''est pas une date correcte',Ecran.Caption);
                        GSessions.Col := ACol;
                        GSessions.Row := ARow;
                        Exit;
                end;
        end;
end;

procedure TOF_PGCURSUSSESSION.GrilleElipsisClick (Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGCURSUSSESSION.InscrptionStagiaire(Sender : TObject);
var St : String;
begin
        St := getControlText('CURSUS')+';'+GetControlText('RANGCURSUS');
        AGLLanceFiche('PAY','CURSUSSALARIE','','',St);
end;

procedure TOF_PGCURSUSSESSION.SupprimerCursus (Sender : TObject);
var Rep : Integer;
    GSessions : THGrid;
    i : Integer;
    NumOrdre,Stage,Millesime : String;
begin
        Rep := PGIAskCancel('Attention, la suppression du cursus va entrainer celle de toute les sessions et stagiaires attachés à ce cursus#13#10 Voulez-Vous continuer',Ecran.Caption);
        If Rep <> MrYes then Exit;
        GSessions := THGrid(GetControl('GSESSIONS'));
        If GSessions = Nil then Exit;
        For i := 1 to GSessions.RowCount - 1 do
        begin
                Stage := GSessions.Cellvalues[0,i];
                NumOrdre := GSessions.Cellvalues[4,i];
                Millesime := GSessions.Cellvalues[5,i];
                ExecuteSQL('DELETE FROM SESSIONANIMAT WHERE '+
                'PAN_CODESTAGE="'+Stage+'" AND PAN_ORDRE='+NumOrdre+' AND PAN_MILLESIME="'+Millesime+'"');
                ExecuteSQL('DELETE FROM FORMATIONS WHERE '+
                'PFO_CODESTAGE="'+Stage+'" AND PFO_ORDRE='+NumOrdre+' AND PFO_MILLESIME="'+Millesime+'"');
                ExecuteSQL('DELETE FROM SESSIONSTAGE WHERE '+
                'PSS_CODESTAGE="'+Stage+'" AND PSS_ORDRE='+NumOrdre+' AND PSS_MILLESIME="'+Millesime+'"');
        end;
        ExecuteSQL('DELETE FROM CURSUS WHERE PCU_CURSUS="'+LeCursus+'" AND PCU_RANGCURSUS='+IntToStr(LeRang)+'');  //DB2
        ExecuteSQL('DELETE FROM CURSUSSTAGE WHERE PCC_CURSUS="'+LeCursus+'" AND PCC_RANGCURSUS='+IntToStr(LeRang)+'');  //DB2
        Ecran.Close;
end;

Initialization
  registerclasses ( [ TOF_PGCURSUSSESSION ] ) ;
end.

