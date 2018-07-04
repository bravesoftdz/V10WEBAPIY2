{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCURSUSSESSION ()
Mots clefs ... : TOF;PGCURSUSSESSION
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
}
Unit UTofPGCursusSalarie ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTob,HSysMenu,HTB97,ParamDat ;

Type
  TOF_PGCURSUSSALARIE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    LeCursus : String;
    LeRang : Integer;
    GSalaries : THGrid;
    procedure RemplirGrilleSal;
    procedure AjouterFormation (Sender : TObject);
    procedure SupprimerFormation (Sender : TObject);
    procedure EffacerZoneFormation (Sender : TObject);
    procedure ClickGrille(Sender : TObject);
    end ;

Implementation

procedure TOF_PGCURSUSSALARIE.OnArgument (S : String ) ;
var BAjout,BSupp,BEfface : TToolBarButton97;
begin
  Inherited ;
        LeCursus := ReadTokenPipe (S,';');
        GSalaries := THGrid(GetControl('GRILLESAL'));
        If LeCursus <> '' then
                begin
                LeRang := StrToInt(ReadTokenPipe(S,';'));
                RemplirGrilleSal;
                end
        Else LeRang := 0;
        If GSalaries <> Nil then GSalaries.OnClick := ClickGrille;
        BAjout := TToolBarButton97(GetControl('BAJOUTSTAGE'));
        If BAjout <> Nil Then BAjout.OnClick := AjouterFormation;
        BSupp := TToolBarButton97(GetControl('BSUPPSTAGE'));
        If BSupp <> Nil Then BSupp.OnClick := SupprimerFormation;
        BEfface := TToolBarButton97(GetControl('BEFFACE'));
        If BEfface <> Nil Then BEfface.OnClick := EffacerZoneFormation;
        SetControlVisible('BVALIDER',False);
end ;

procedure TOF_PGCURSUSSALARIE.RemplirGrilleSal;
var HMTrad : THSystemMenu;
    Q : TQuery;
    TobSalaries : Tob;
    i : Integer;
    Salarie,Nom,Prenom,Etab,LibEmploi : String;
begin
        If GSalaries = Nil then Exit;
        If LeRang > 0 then
         Q := OpenSQL('SELECT PFO_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI FROM FORMATIONS '+
         'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
         'LEFT JOIN CURSUSSTAGE ON PFO_CODESTAGE=PCC_CODESTAGE AND PFO_ORDRE=PCC_ORDRE '+
         'WHERE PCC_RANGCURSUS='+IntToStr(LeRang)+' AND PCC_CURSUS="'+LeCursus+'" '+  //DB2
         'GROUP BY PFO_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI',True);
        TobSalaries := Tob.Create('Les salaries',Nil,-1);
        TobSalaries.LoadDetailDB('Les salaries','','',Q,False);
        Ferme(Q);
        For i := 1 to GSalaries.RowCount - 1 do
        begin
                GSalaries.Rows[i].Clear;
        end;
        GSalaries.RowCount := 2;
        For i := 0 to TobSalaries.Detail.Count -1 do
        begin
                Salarie := TobSalaries.Detail[i].GetValue('PFO_SALARIE');
                Nom := TobSalaries.Detail[i].GetValue('PSA_LIBELLE');
                Prenom := TobSalaries.Detail[i].GetValue('PSA_PRENOM');
                Etab := TobSalaries.Detail[i].GetValue('PSA_ETABLISSEMENT');
                LibEmploi := TobSalaries.Detail[i].GetValue('PSA_LIBELLEEMPLOI');
                If i > 0 then GSalaries.RowCount := GSalaries.RowCount + 1;
                GSalaries.CellValues[0,i+1] := Salarie;
                GSalaries.CellValues[1,i+1] := Nom;
                GSalaries.CellValues[2,i+1] := Prenom;
                GSalaries.CellValues[3,i+1] := RechDom('TTETABLISSEMENT',Etab,False);
                GSalaries.CellValues[4,i+1] := RechDom('PGLIBEMPLOI',LibEmploi,False);
        end;
        TobSalaries.free;
        HMTrad:=THSystemMenu(GetControl('HMTrad'));
        HMTrad.ResizeGridColumns (GSalaries);
end;

procedure TOF_PGCURSUSSALARIE.AjouterFormation (Sender : TObject);
var TobFormations,TobSessions,TF: Tob;
    Q : TQuery;
    j : Integer;
    Salarie : String;
    Nom,Prenom,Etab,LibEmploi : String;
    TN1,TN2,TN3,TN4,LC1,LC2,LC3,LC4,CS : String;
begin
        If GSalaries = Nil then Exit;
        Salarie := GetControlText('SALARIE');
        If Salarie='' then exit;
        Q := OpenSQL('SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_LIBELLE,PSA_PRENOM,'+
        'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4 '+
        'FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
        If Not Q.eof then
        begin
                Nom := Q.FindField('PSA_LIBELLE').AsString;
                Prenom := Q.FindField('PSA_PRENOM').AsString;
                Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
                LibEmploi := Q.FindField('PSA_LIBELLEEMPLOI').AsString;
                TN1 := Q.FindField('PSA_TRAVAILN1').AsString;
                TN2 := Q.FindField('PSA_TRAVAILN2').AsString;
                TN3 := Q.FindField('PSA_TRAVAILN3').AsString;
                TN4 := Q.FindField('PSA_TRAVAILN4').AsString;
                LC1 := Q.FindField('PSA_LIBREPCMB1').AsString;
                LC2 := Q.FindField('PSA_LIBREPCMB2').AsString;
                LC3 := Q.FindField('PSA_LIBREPCMB3').AsString;
                LC4 := Q.FindField('PSA_LIBREPCMB4').AsString;
                CS := Q.FindField('PSA_CODESTAT').AsString;
        end;
        Ferme(Q);
        Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_DATEDEBUT,PSS_DATEFIN,'+
        'PSS_FORMATION1,PSS_FORMATION2,PSS_FORMATION3,PSS_FORMATION4,PSS_FORMATION5,'+
        'PSS_FORMATION6,PSS_FORMATION7,PSS_FORMATION8,PSS_DEBUTDJ,PSS_FINDJ,PSS_ORGCOLLECTSGU FROM SESSIONSTAGE '+
        'LEFT JOIN CURSUSSTAGE ON PSS_ORDRE=PCC_ORDRE AND PSS_CODESTAGE=PCC_CODESTAGE '+
        'WHERE PCC_RANGCURSUS='+IntToStr(LeRang)+' AND PCC_CURSUS="'+LeCursus+'"',True);  //DB2
        TobSessions := Tob.Create('Les sessions',Nil,-1);
        TobSessions.LoadDetailDB('Les sessions','','',Q,false);
        Ferme(Q);
        TobFormations := Tob.Create('Les formations',Nil,-1);
        For j := 0 to TobSessions.Detail.Count - 1 do
        begin
                If Salarie = '' then continue;
                TF := Tob.Create('FORMATIONS',TobFormations,-1);
                TF.PutValue('PFO_SALARIE',Salarie);
                TF.PutValue('PFO_CODESTAGE',TobSessions.Detail[j].GetValue('PSS_CODESTAGE'));
                TF.PutValue('PFO_ORDRE',TobSessions.Detail[j].GetValue('PSS_ORDRE'));
                TF.PutValue('PFO_MILLESIME',TobSessions.Detail[j].GetValue('PSS_MILLESIME'));
                TF.PutValue('PFO_DATEDEBUT',TobSessions.Detail[j].GetValue('PSS_DATEDEBUT'));
                TF.PutValue('PFO_DATEFIN',TobSessions.Detail[j].GetValue('PSS_DATEFIN'));
                TF.PutValue('PFO_FORMATION1',TobSessions.Detail[j].GetValue('PSS_FORMATION1'));
                TF.PutValue('PFO_FORMATION2',TobSessions.Detail[j].GetValue('PSS_FORMATION2'));
                TF.PutValue('PFO_FORMATION3',TobSessions.Detail[j].GetValue('PSS_FORMATION3'));
                TF.PutValue('PFO_FORMATION4',TobSessions.Detail[j].GetValue('PSS_FORMATION4'));
                TF.PutValue('PFO_FORMATION5',TobSessions.Detail[j].GetValue('PSS_FORMATION5'));
                TF.PutValue('PFO_FORMATION6',TobSessions.Detail[j].GetValue('PSS_FORMATION6'));
                TF.PutValue('PFO_FORMATION7',TobSessions.Detail[j].GetValue('PSS_FORMATION7'));
                TF.PutValue('PFO_FORMATION8',TobSessions.Detail[j].GetValue('PSS_FORMATION8'));
                TF.PutValue('PFO_DEBUTDJ',TobSessions.Detail[j].GetValue('PSS_DEBUTDJ'));
                TF.PutValue('PFO_FINDJ',TobSessions.Detail[j].GetValue('PSS_FINDJ'));
                TF.PutValue('PFO_ORGCOLLECTGU',TobSessions.Detail[j].GetValue('PSS_ORGCOLLECTSGU'));
                TF.PutValue('PFO_NOMSALARIE',Nom);
                TF.PutValue('PFO_PRENOM',Prenom);
                TF.PutValue('PFO_ETABLISSEMENT',Etab);
                TF.PutValue('PFO_LIBEMPLOIFOR',LibEmploi);
                TF.PutValue('PFO_TRAVAILN1',TN1);
                TF.PutValue('PFO_TRAVAILN2',TN2);
                TF.PutValue('PFO_TRAVAILN3',TN3);
                TF.PutValue('PFO_TRAVAILN4',TN4);
                TF.PutValue('PFO_LIBREPCMB1',LC1);
                TF.PutValue('PFO_LIBREPCMB2',LC2);
                TF.PutValue('PFO_LIBREPCMB3',LC3);
                TF.PutValue('PFO_LIBREPCMB4',LC4);
                TF.PutValue('PFO_CODESTAT',CS);
                TF.InsertOrUpdateDB(False);
        end;
        TobSessions.Free;
        TobFormations.Free;
        SetControlText('SALARIE','');
        SetControlCaption('LIBSALARIE','');
        RemplirGrilleSal;
end;

procedure TOF_PGCURSUSSALARIE.SupprimerFormation (Sender : TObject);
var TobSessions : Tob;
    Q : TQuery;
    i : Integer;
    Stage,Ordre,Millesime,Salarie : String;
begin
        Salarie := GetControlText('SALARIE');
        If Salarie ='' then Exit;
        Q:= OpenSQL('SELECT PSS_CODESTAGE,PSS_MILLESIME,PSS_ORDRE FROM SESSIONSTAGE '+
        'LEFT JOIN CURSUSSTAGE ON PSS_CODESTAGE=PCC_CODESTAGE AND PSS_ORDRE=PCC_ORDRE '+
        'WHERE PCC_RANGCURSUS='+IntToStr(LeRang)+' AND PCC_CURSUS="'+LeCursus+'"',True);//DB2
        TobSessions := Tob.Create('Les sessions',Nil,-1);
        TobSessions.LoadDetailDB('Les sessions','','',Q,false);
        Ferme(Q);
        For i := 0 to TobSessions.Detail.Count -1 do
        begin
                Stage := TobSessions.detail[i].GetValue('PSS_CODESTAGE');
                Ordre := IntToStr(TobSessions.detail[i].GetValue('PSS_ORDRE'));
                Millesime := TobSessions.detail[i].GetValue('PSS_MILLESIME');
                ExecuteSQL('DELETE FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" '+
                'AND PFO_CODESTAGE="'+Stage+'" AND PFO_ORDRE='+Ordre+' AND PFO_MILLESIME="'+Millesime+'"');  //DB2
        end;
        TobSessions.Free;
        SetControlText('SALARIE','');
        SetControlCaption('LIBSALARIE','');
        RemplirGrilleSal;
end;

procedure TOF_PGCURSUSSALARIE.EffacerZoneFormation (Sender : TObject);
begin
        SetControlText('SALARIE','');
        SetControlCaption('LIBSALARIE','');
end;

procedure TOF_PGCURSUSSALARIE.ClickGrille(Sender : TObject);
var Salarie : String;
    Ou : Integer;
begin
        Ou := GSalaries.Row;
        Salarie := GSalaries.CellValues[0,Ou];
        SetControlText('SALARIE',Salarie);
        SetControlCaption('LIBSALARIE',RechDom('PGSALARIE',Salarie,False));
end;

Initialization
  registerclasses ( [ TOF_PGCURSUSSALARIE ] ) ;
end.

