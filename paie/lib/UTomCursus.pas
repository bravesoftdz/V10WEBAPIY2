{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CURSUS (CURSUS)
Mots clefs ... : TOM;CURSUS
*****************************************************************
PT1 | 24/11/2003 | V_50  | JL | Correction "conversion type variant ncorrect" en CWAS
--- | 20/03/2006 |       | JL | modification clé annuaire ----
PT2 | 01/10/2007 | V_80  | FL | Suppression de la clause sur PCC_SALARIE qui n'existe pas
PT3 | 02/04/2008 | V_80  | FL | Partage formation
PT4 | 02/04/2008 | V_80  | FL | Prise en compte des groupes de travail
PT5 | 03/04/2008 | V_80  | FL | Restriction sur les prédéfinis
PT6 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue
}
Unit UTomCursus ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     eFiche,eFichList,   
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOM,UTob,PgOutils2,PgOutilsFormation ;

Type
  TOM_CURSUS = Class (TOM)
    procedure OnUpdateRecord             ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField (F: TField)  ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnDeleteRecord             ; override ;
    private
    StArgument : String;
    end ;

Implementation

Uses GalOutil,PGOutils;

procedure TOM_CURSUS.OnUpdateRecord ;
var TobSessions,TS,TobCursusStage,TCS,TobInvest,TI,TobStage : Tob;
    Q : TQuery;
    IMax,i,LeRang : Integer;
    Stage : String;
begin
  Inherited ;
        If PGBundleInscFormation then //PT3
        begin
          If GetField('PCU_PREDEFINI') = 'STD' then SetField('PCU_NODOSSIER','000000')
          else SetField('PCU_NODOSSIER',V_PGI.NoDossier);
        end;
        If StArgument ='SESSIONCURSUS' then
        begin
                If (DS.State in [dsInsert]) then
                begin
                        Q := OpenSQL('SELECT MAX(PCU_RANGCURSUS) FROM CURSUS WHERE PCU_CURSUS="'+GetField('PCU_CURSUS')+'"',True);
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
                        SetField('PCU_RANGCURSUS',IntToStr(LeRang));  //PT1
                        TobCursusStage := Tob.Create('Lescursusetstage',Nil,-1);
                        TobSessions := Tob.Create('Les sessions',Nil,-1);
                        Q := OpenSQL('SELECT STAGE.* FROM STAGE LEFT JOIN CURSUSSTAGE ON PST_CODESTAGE=PCC_CODESTAGE '+
                        'WHERE PST_MILLESIME="0000" AND PCC_CURSUS="'+GetField('PCU_CURSUS')+'" AND PCC_RANGCURSUS=0 AND PCC_SALARIE=""',True);
                        TobStage := Tob.Create('LesStages',Nil,-1);
                        TobStage.LoadDetailDB('LesSatges','','',Q,False);
                        Ferme(Q);
                        For i := 0 to TobStage.detail.Count - 1 do
                        begin
                                Stage := TobStage.Detail[i].GetValue('PST_CODESTAGE');
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
                                TCS.PutValue('PCC_CURSUS',GetField('PCU_CURSUS'));
                                TCS.PutValue('PCC_CODESTAGE',Stage);
                                TCS.PutValue('PCC_ORDRE',IMax);
                                TCS.PutValue('PCC_SALARIE','');
                                TCS.InsertDB(Nil,False);
                                TS := Tob.Create('SESSIONSTAGE',TobSessions,-1);
                                TS.PutValue('PSS_CODESTAGE',Stage);
                                TS.PutValue('PSS_ORDRE',IMax);
                                TS.PutValue('PSS_MILLESIME','0000');
                                TS.PutValue('PSS_DATEDEBUT',IDate1900);
                                TS.PutValue('PSS_DATEFIN',IDate1900);
                                TS.PutValue ('PSS_CENTREFORMGU',TobStage.Detail[i].GetValue('PST_CENTREFORMGU'));
                                TS.PutValue('PSS_LIEUFORM',TobStage.Detail[i].GetValue('PST_LIEUFORM'));
                                TS.PutValue('PSS_ORGCOLLECTPGU',TobStage.Detail[i].GetValue('PST_ORGCOLLECTPGU'));
                                TS.PutValue('PSS_ORGCOLLECTSGU',TobStage.Detail[i].GetValue('PST_ORGCOLLECTSGU'));
                                TS.PutValue('PSS_NATUREFORM',TobStage.Detail[i].GetValue('PST_NATUREFORM'));
                                TS.PutValue('PSS_TYPCONVFORM',TobStage.Detail[i].GetValue('PST_TYPCONVFORM'));
                                TS.PutValue('PSS_COUTPEDAG',TobStage.Detail[i].GetValue('PST_COUTBUDGETE'));
                                TS.PutValue('PSS_COUTFONCT',TobStage.Detail[i].GetValue('PST_COUTFONCT'));
                                TS.PutValue('PSS_COUTORGAN',TobStage.Detail[i].GetValue('PST_COUTORGAN'));
                                TS.PutValue('PSS_COUTUNITAIRE',TobStage.Detail[i].GetValue('PST_COUTUNITAIRE'));
                                TS.PutValue('PSS_COUTMOB',TobStage.Detail[i].GetValue('PST_COUTMOB'));
                                TS.PutValue('PSS_FORMATION1',TobStage.Detail[i].GetValue('PST_FORMATION1'));
                                TS.PutValue('PSS_FORMATION2',TobStage.Detail[i].GetValue('PST_FORMATION2'));
                                TS.PutValue('PSS_FORMATION3',TobStage.Detail[i].GetValue('PST_FORMATION3'));
                                TS.PutValue('PSS_FORMATION4',TobStage.Detail[i].GetValue('PST_FORMATION4'));
                                TS.PutValue('PSS_FORMATION5',TobStage.Detail[i].GetValue('PST_FORMATION5'));
                                TS.PutValue('PSS_FORMATION6',TobStage.Detail[i].GetValue('PST_FORMATION6'));
                                TS.PutValue('PSS_FORMATION7',TobStage.Detail[i].GetValue('PST_FORMATION7'));
                                TS.PutValue('PSS_FORMATION8',TobStage.Detail[i].GetValue('PST_FORMATION8'));
                                TS.PutValue('PSS_DUREESTAGE',TobStage.Detail[i].GetValue('PST_DUREESTAGE'));
                                TS.PutValue('PSS_JOURSTAGE',TobStage.Detail[i].GetValue('PST_JOURSTAGE'));
                                TS.PutValue('PSS_INCLUSDECL',TobStage.Detail[i].GetValue('PST_INCLUSDECL'));
                                TS.PutValue('PSS_ACTIONFORM',TobStage.Detail[i].GetValue('PST_ACTIONFORM'));
                                TS.PutValue('PSS_CURSUS',GetField('PCU_CURSUS'));
                                TS.PutValue('PSS_RANGCURSUS',LeRang);
                                TS.PutValue('PSS_DEBUTDJ','MAT');
                                TS.PutValue('PSS_FINDJ','PAM');
                                TS.PutValue('PSS_EFFECTUE','-');
                                TS.PutValue('PSS_CLOTUREINSC','-');
                                TS.PutValue('PSS_VALIDORG','-');
                                TS.PutValue('PSS_HEUREDEBUT',0);
                                TS.PutValue('PSS_HEUREFIN',0);
                                TS.PutValue('PSS_AVECCLIENT','-');
                                TS.InsertDB(Nil,False);
                                Q := OpenSQL('SELECT * FROM INVESTSESSION '+
                                'WHERE PIS_CODESTAGE="'+Stage+'" AND PIS_ORDRE=-1 AND PIS_MILLESIME="0000"',True);  //DB2
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
                        TobSessions.Free;
                        TobCursusStage.Free;
                        TobStage.Free;
                end;
        end;
end;

procedure TOM_CURSUS.OnNewRecord ;
var IMax : Integer;
    Q : TQuery;
begin
  Inherited ;
        SetField('PCU_RANGCURSUS',0);  //PT1
        SetField('PCU_DATEDEBUT',IDate1900);
        SetField('PCU_DATEFIN',IDate1900);
        If StArgument <> 'SESSIONCURSUS' then
        begin
                Q := OpenSQL('SELECT MAX(PCU_CURSUS) MAXCUR FROM CURSUS',True);
                If Not Q.Eof then
                begin
                        If Q.FindField('MAXCUR').AsString <> '' then IMax := StrToInt(Q.FindField('MAXCUR').AsString)
                        else IMax := 0;
                end
                else IMax := 0;
                Ferme(Q);
                SetField('PCU_CURSUS',IntToStr(IMax+1));
        end;
        SetField('PCU_INCLUSCAT','X');
        SetField('PST_NODOSSIER',V_PGI.NoDossier);
        If PGDroitMultiForm then //PT3
        begin
          If PGDroitMultiForm then SetField('PCU_PREDEFINI','STD')
          else SetField('PCU_PREDEFINI','DOS');
        end
        Else SetField('PCU_PREDEFINI','DOS');
end ;

procedure TOM_CURSUS.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CURSUS.OnChangeField (F: TField);
var Cursus : String;
begin
        Inherited ;
        If F.FieldName = 'PCU_CURSUS' then
        begin
                If GetField('PCU_CURSUS') <> '' then
                begin
                        Cursus := Trim(GetField('PCU_CURSUS'));
                        Cursus := ColleZeroDevant(StrToInt(trim(Cursus)),3);
                        If Cursus <> GetField('PCU_CURSUS') then SetField('PCU_CURSUS',Cursus);
                end;
        end;
end;

procedure TOM_CURSUS.OnArgument ( S: String ) ;
begin
  Inherited ;
    StArgument := ReadTokenPipe(S,';');
    SetControlProperty('PCU_CURSUS','EditMask','999');
    If StArgument = 'SESSIONCURSUS' then
    begin
            SetControlVisible('PCU_DATEDEBUT',True);
            SetControlVisible('PCU_DATEFIN',True);
            SetControlVisible('TPCU_DATEDEBUT',True);
            SetControlVisible('TPCU_DATEFIN',True);
    end;
    
    //PT3 - Début
	If PGBundleCatalogue then //PT6
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('PCU_NODOSSIER',False);
			SetControlText   ('PCU_NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PCU_PREDEFINI',False);
			SetControlText   ('PCU_PREDEFINI','');
		end
    	Else If V_PGI.ModePCL='1' Then 
    	Begin
    		SetControlProperty ('PCU_NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT4
		End;
	end
	else
	begin
		SetControlVisible('PCU_PREDEFINI', False);
		SetControlVisible('TPCU_PREDEFINI',False);
		SetControlVisible('PCU_NODOSSIER',     False);
		SetControlVisible('TPCU_NODOSSIER',    False);
	end;
	//PT3 - Fin

    GereAccesPredefini (THValComboBox(GetControl('PCU_PREDEFINI'))); //PT5
end ;

procedure TOM_CURSUS.OnDeleteRecord ;
var TobCursusStage : Tob ;
    Q : TQuery ;
    i : Integer ;
begin
  Inherited ;
        Q := OpenSQL('SELECT * FROM CURSUSSTAGE WHERE PCC_CURSUS="' + GetField('PCU_CURSUS') + '" AND PCC_RANGCURSUS='+IntToStr(GetField('PCU_RANGCURSUS')),True);  //DB2 //PT2
        TobCursusStage := Tob.Create('CURSUSSTAGE',Nil,-1);
        TobCursusStage.LoadDetailDB('CURSUSSTAGE','','',Q,False);
        Ferme(Q);
        For i := 0 To TobCursusStage.Detail.Count -1 do
        begin
                TobCursusStage.Detail[i].DeleteDB(False);
        end;
        TobCursusStage.Free;
end ;


Initialization
  registerclasses ( [ TOM_CURSUS ] ) ;
end.

