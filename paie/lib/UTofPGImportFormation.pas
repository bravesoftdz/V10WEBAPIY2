{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 29/05/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGIMPORTFORMATION ()
Mots clefs ... : TOF;PGIMPORTFORMATION
*****************************************************************}
Unit UTofPGImportFormation ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms,
     uTob,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     Utobxls,
     HMsgBox,
     PGOutils2,
     UTobDebug,
     ed_tools,
     UTOF ;

Type
  TOF_PGIMPORTFORMATION = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    procedure ImportForm;
    procedure ChangeChemin(Sender : TObject);
  end ;

Implementation

procedure TOF_PGIMPORTFORMATION.ChangeChemin(Sender : TObject);
var Chemin : String;
begin
     Chemin := GetControlText('REPERTOIRE');
     SetControlChecked('FICHIERORG',FileExists(Chemin+'\ANNUAIRE.txt'));
     SetControlChecked('FICHIERSTAGE',FileExists(Chemin+'\STAGE.txt'));
     SetControlChecked('FICHIERSESSION',FileExists(Chemin+'\SESSIONSTAGE.txt'));
     SetControlChecked('FICHIERSTAGIAIRES',FileExists(Chemin+'\FORMATIONS.txt'));
     SetControlChecked('FICHIERANIM',FileExists(Chemin+'\SESSIONANIMAT.txt'));
     SetControlChecked('CONVERSIONSAL',FileExists(Chemin+'\CCMX.xls'));
end;

procedure TOF_PGIMPORTFORMATION.ImportForm ;
var TobFSession,TobFStage,TobFForm,TobFAnim,TobFAnn : tob;
    TobConvertStage,TobConvertSession,TobConvertAnn : Tob;
    TobSalaries : Tob;
    TC,T : Tob;
    i,f : Integer;
    Ancien,Nouveau,Stage,Salarie,NomOrg,FileN : String;
    IAncien,INouveau,IMax,ISal,Centre : Integer;
    Q : TQuery;
    TobStage,TobSessionStage,TobFormations,TobAnnuaire,TobAnimateurs : Tob;
    IAnn,INouveauAnn,NewCCMx : Integer;
    Guid,AncienGuid,NouveauGuid,NewCCMxGUid : String;
    TypePer : string;
    TobSessions,TS : Tob;
    Montant : Double;
    QS,QF : TQuery;
    TobFormation: Tob;
     LeStage, LaSession: string;
  CoutSalarie, TotalHeures, HeuresSal, CoutAnim, CoutGlobal, CoutParStagiaire: Double;
  matriculeSal,Chemin : String;
begin

     Chemin := GetControlText('REPERTOIRE');

     TobSalaries  := Tob.Create('LesSalaries',Nil,-1);
     FileN := Chemin+'\CCMX.xls';

     ImportTOBFromXLS (TobSalaries, FileN);

     For i := 0 to TobSalaries.Detail.Count - 1 do
     begin
          If IsNumeric(TobSalaries.Detail[i].GetValue('MATRICULE')) and (TobSalaries.Detail[i].GetValue('MATRICULE')<>'') then
          begin
            TobSalaries.Detail[i].Putvalue('MATRICULE',ColleZeroDevant(StrToInt(TobSalaries.Detail[i].GetValue('MATRICULE')),10));
          end;
          If IsNumeric(TobSalaries.Detail[i].GetValue('CCMX')) and (TobSalaries.Detail[i].GetValue('CCMX')<>'') then
          begin
            TobSalaries.Detail[i].Putvalue('CCMX',ColleZeroDevant(StrToInt(TobSalaries.Detail[i].GetValue('CCMX')),6));
          end;
     end;
     TobFAnn := Tob.Create('les organismes',Nil,-1);
     TOBLoadFromFile(Chemin+'\ANNUAIRE.txt', nil, TobFAnn);



     TobConvertStage := Tob.Create('Nouveau stage',Nil,-1);
     TobConvertSession := Tob.Create('Nouveau numero',Nil,-1);
     TobConvertAnn := Tob.Create('Nouveau annuaire',Nil,-1);

     InitMoveProgressForm(nil, 'Récupération de l''annuaire', 'Veuillez patienter SVP ...', TobFAnn.Detail.Count - 1, FALSE, TRUE);
     TobAnnuaire := Tob.Create('ANNUAIRE',Nil,-1);
     try
    BeginTrans;
     For i := 0 to TobFAnn.Detail.Count - 1 do
     begin
          T := Tob.Create('ANNUAIRE',TobAnnuaire,-1);
          Q := OpenSQL('SELECT MAX(ANN_CODEPER) FROM ANNUAIRE',TRUE) ;
          if not Q.Eof then
          begin
               If Q.Fields[0].AsString <> '' Then Imax := Q.Fields[0].AsInteger
               Else Imax := 0;
               Imax := Imax + 1;
          end
          else IMax := 1;
          Ferme(Q) ;
          Ancien := TobFAnn.Detail[i].getValue('ANN_NOMPER');
          TC := Tob.Create('Fille convert',TobConvertAnn,-1);
          TC.AddChampSupValeur('ANCIEN',Ancien);
          TC.AddChampSupValeur('NOUVEAU',IMax);
          T.PutValue('ANN_GUIDPER', AglGetGuid());
          TC.AddChampSupValeur('GUID',T.GetValue('ANN_GUIDPER'));
          T.PutValue('ANN_CODEPER', IMax);
          T.PutValue('ANN_FAMPER', 'SOC');
          If (Ancien = 'FAFIEC') or (ancien='FONGECIF') then TypePer := 'FCO'
          else TypePer := 'CFO';
          T.PutValue('ANN_TYPEPER', TypePer);
          T.PutValue('ANN_NOM1',TobFAnn.Detail[i].GetValue('ANN_NOMPER'));
          T.PutValue('ANN_NOMPER',TobFAnn.Detail[i].GetValue('ANN_NOMPER'));
          T.PutValue('ANN_APNOM',TobFAnn.Detail[i].GetValue('ANN_APNOM'));
          T.PutValue('ANN_APRUE1',TobFAnn.Detail[i].GetValue('ANN_APRUE1'));
          T.PutValue('ANN_APRUE2',TobFAnn.Detail[i].GetValue('ANN_APRUE2'));
          T.PutValue('ANN_APCPVILLE',TobFAnn.Detail[i].GetValue('ANN_APCPVILLE'));
          T.PutValue('ANN_EMAIL',TobFAnn.Detail[i].GetValue('ANN_EMAIL'));
          T.PutValue('ANN_FAX',TobFAnn.Detail[i].GetValue('ANN_FAX'));
          T.PutValue('ANN_TEL1',TobFAnn.Detail[i].GetValue('ANN_TEL1'));
          T.InsertDB(Nil);
          If Ancien = 'CCMX' then
          begin
               Q := OpenSQL('SELECT MAX(ANN_CODEPER) FROM ANNUAIRE',TRUE) ;
               if not Q.Eof then
               begin
                    If Q.Fields[0].AsString <> '' Then Imax := Q.Fields[0].AsInteger
                    Else Imax := 0;
                    Imax := Imax + 1;
               end
               else IMax := 1;
               Ferme(Q) ;
               NewCCMx := IMax;
               TypePer := 'FCO';
               T.PutValue('ANN_TYPEPER', TypePer);
               T.PutValue('ANN_CODEPER', IMax);
               T.PutValue('ANN_GUIDPER', AglGetGuid());
               NewCCMxGUid := T.GetValue('ANN_GUIDPER');
               T.InsertDB(Nil);
          end;
          MoveCurProgressForm(TobFAnn.Detail[i].getValue('ANN_NOMPER'));
     end;
     FiniMoveProgressForm;
     TobFAnn.Free;
     TobAnnuaire.Free;
     TobFStage := Tob.Create('Les stages',Nil,-1);
     TOBLoadFromFile(Chemin+'\STAGE.txt', nil, TobFStage);
     TobStage := Tob.Create('STAGE',Nil,-1);
     InitMoveProgressForm(nil, 'Récupération des stages', 'Veuillez patienter SVP ...', TobFStage.Detail.Count - 1, FALSE, TRUE);
     For i := 0 to TobFStage.Detail.Count - 1 do
     begin
          T := Tob.Create('STAGE',TobStage,-1);
          Q := OpenSQL('SELECT MAX(PST_CODESTAGE) FROM STAGE where pst_CODEStage <>"2222"',TRUE) ;
          if not Q.Eof then
          begin
               If Q.Fields[0].AsString <> '' Then Imax := Q.Fields[0].AsInteger
               Else Imax := 0;
               Imax := Imax + 1;
          end
          else IMax := 1;
          Ferme(Q) ;
          Ancien := TobFStage.Detail[i].getValue('PST_CODESTAGE');
          Nouveau := ColleZeroDevant(IMax,6);
          T.PutValue('PST_CODESTAGE',Nouveau);
          TC := Tob.Create('Fille convert',TobConvertStage,-1);
          TC.AddChampSupValeur('ANCIEN',Ancien);
          TC.AddChampSupValeur('NOUVEAU',Nouveau);
          TC.AddChampSupValeur('INCLUS',TobFStage.Detail[i].GetValue('PST_INCLUSDECL'));
          T.PutValue('PST_MILLESIME','0000');
          T.PutValue('PST_ACTIF','X');
          T.PutValue('PST_LIBELLE',TobFStage.Detail[i].GetValue('PST_LIBELLE'));
          T.PutValue('PST_NATUREFORM',TobFStage.Detail[i].GetValue('PST_NATUREFORM'));
          T.PutValue('PST_TYPCONVFORM',TobFStage.Detail[i].GetValue('PST_TYP_CONV_FORM'));
          T.PutValue('PST_ACTIONFORM',TobFStage.Detail[i].GetValue('PST_ACTIONFORM'));
          T.PutValue('PST_DUREESTAGE',TobFStage.Detail[i].GetValue('PST_DUREESTAGE'));
          T.PutValue('PST_JOURSTAGE',TobFStage.Detail[i].GetValue('PST_JOURSTAGE'));
          T.PutValue('PST_INCLUSDECL',TobFStage.Detail[i].GetValue('PST_INCLUSDECL'));
          T.PutValue('PST_FORMATION1','12');
          T.PutValue('PST_TYPCONVFORM','ANN');
          NomOrg := TobFStage.Detail[i].GetValue('CENTREFORM');
          TC := TobConvertAnn.FindFirst(['ANCIEN'],[NomOrg],False);
          If TC <> Nil then
          begin
               T.PutValue('PST_CENTREFORM',TC.GetValue('NOUVEAU'));
               T.PutValue('PST_CENTREFORMGU',TC.GetValue('GUID'));
          end;
          NomOrg := TobFStage.Detail[i].GetValue('ORGCOLLECT');
          TC := TobConvertAnn.FindFirst(['ANCIEN'],[NomOrg],False);
          If TC <> Nil then
          begin
               If NomOrg ='FONGECIF' then
               begin
                T.PutValue('PST_ORGCOLLECTP',TC.GetValue('NOUVEAU'));
                T.PutValue('PST_ORGCOLLECTS',TC.GetValue('NOUVEAU'));
               T.PutValue('PST_ORGCOLLECTPGU',TC.GetValue('GUID'));
                T.PutValue('PST_ORGCOLLECTSGU',TC.GetValue('GUID'));
               end
               else
               begin
                If NomOrg='FAFIEC' then
                begin
                     T.PutValue('PST_ORGCOLLECTP',TC.GetValue('NOUVEAU'));
                     T.PutValue('PST_ORGCOLLECTPGU',TC.GetValue('GUID'));
                end
                else
                begin
                     T.PutValue('PST_ORGCOLLECTP',NewCCMx);
                     T.PutValue('PST_ORGCOLLECTPGU',NewCCMxGUid);
                end;
                T.PutValue('PST_ORGCOLLECTS',NewCCMx);
                T.PutValue('PST_ORGCOLLECTSGU',NewCCMxGUid);
               end;
          end
          else
          begin
                T.PutValue('PST_ORGCOLLECTP',NewCCMx);
                T.PutValue('PST_ORGCOLLECTS',NewCCMx);
               T.PutValue('PST_ORGCOLLECTPGU',NewCCMxGUid);
                T.PutValue('PST_ORGCOLLECTSGU',NewCCMxGUid);
          end;

          T.InsertDB(Nil);
          MoveCurProgressForm(TobFStage.Detail[i].GetValue('PST_LIBELLE'));
     end;
     FiniMoveProgressForm;
     TobStage.Free;
     TobFStage.Free;

     TobFSession := Tob.Create('Les sessions',Nil,-1);
     TOBLoadFromFile(Chemin+'\SESSIONSTAGE.txt', nil, TobFSession);
     TobSessionStage := Tob.Create('SESSIONSTAGE',Nil,-1);
     InitMoveProgressForm(nil, 'Récupération des sessions', 'Veuillez patienter SVP ...', TobFSession.Detail.Count - 1, FALSE, TRUE);
     For i := 0 to TobFSession.Detail.Count - 1 do
     begin
          Ancien := TobFSession.Detail[i].getValue('PSS_CODESTAGE');
          TC := TobConvertStage.FindFirst(['ANCIEN'],[Ancien],False);
          Nouveau := TC.GetValue('NOUVEAU');
          Q := OpenSQL('SELECT MAX(PSS_ORDRE) FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Nouveau+'"',TRUE) ;
          if not Q.Eof then
          begin
               If Q.Fields[0].AsString <> '' Then Imax := Q.Fields[0].AsInteger
               Else Imax := 0;
               Imax := Imax + 1;
          end
          else IMax := 1;
          Ferme(Q) ;
          T := Tob.Create('SESSIONSTAGE',TobSessionstage,-1);
//        T.PutValue
          T.PutValue('PSS_MILLESIME','0000');
          T.PutValue('PSS_CODESTAGE',Nouveau);
          T.PutValue('PSS_ORDRE',IMax);
          T.PutValue('PSS_LIBELLE',TobFSession.Detail[i].GetValue('PSS_LIBELLE'));
          T.PutValue('PSS_NATUREFORM',TobFSession.Detail[i].GetValue('PSS_NATUREFORM'));
          T.PutValue('PSS_JOURSTAGE',TobFSession.Detail[i].GetValue('PSS_JOURSTAGE'));
          T.PutValue('PSS_DUREESTAGE',TobFSession.Detail[i].GetValue('PSS_DUREESTAGE'));
          T.PutValue('PSS_DATEDEBUT',TobFSession.Detail[i].GetValue('PSS_DATEDEBUT'));
          T.PutValue('PSS_DATEFIN',TobFSession.Detail[i].GetValue('PSS_DATEFIN'));
          T.PutValue('PSS_EFFECTUE',TobFSession.Detail[i].GetValue('PSS_EFFECTUE'));
          T.PutValue('PSS_ACTIONFORM',TobFSession.Detail[i].GetValue('PSS_ACTIONFORM'));
          T.PutValue('PSS_INCLUSDECL',TC.GetValue('INCLUS'));
          T.PutValue('PSS_FORMATION1','12');
          T.PutValue('PSS_TYPCONVFORM','ANN');

          NomOrg := TobFSession.Detail[i].GetValue('CENTREFORM');
          TC := TobConvertAnn.FindFirst(['ANCIEN'],[NomOrg],False);
          If TC <> Nil then
          begin
               Centre := TC.GetValue('NOUVEAU');
               T.PutValue('PSS_CENTREFORM',TC.GetValue('NOUVEAU'));
               T.PutValue('PSS_CENTREFORMGU',TC.GetValue('GUID'));
          end;

          IAncien := TobFSession.Detail[i].GetValue('PSS_ORDRE');
          TC := Tob.Create('Fille convert',TobConvertSession,-1);
          TC.AddChampSupValeur('ANCIEN',IAncien);
          TC.AddChampSupValeur('NOUVEAU',IMax);
          TC.AddChampSupValeur('ANCSTAGE',Ancien);
          TC.AddChampSupValeur('NEWSTAGE',Nouveau);
          TC.AddChampSupValeur('CENTREFORM',Centre);

          NomOrg := TobFSession.Detail[i].GetValue('ORGCOLLECT');
          TC := TobConvertAnn.FindFirst(['ANCIEN'],[NomOrg],False);
          If TC <> Nil then
          begin
               If NomOrg ='FONGECIF' then
               begin
                T.PutValue('PSS_ORGCOLLECTP',TC.GetValue('NOUVEAU'));
                T.PutValue('PSS_ORGCOLLECTS',TC.GetValue('NOUVEAU'));
                T.PutValue('PSS_ORGCOLLECTPGU',TC.GetValue('GUID'));
                T.PutValue('PSS_ORGCOLLECTSGU',TC.GetValue('GUID'));
               end
               else
               begin
                If NomOrg='FAFIEC' then
                begin
                     T.PutValue('PSS_ORGCOLLECTP',TC.GetValue('NOUVEAU'));
                     T.PutValue('PSS_ORGCOLLECTPGU',TC.GetValue('GUID'));
                end
                else
                begin
                     T.PutValue('PSS_ORGCOLLECTP',NewCCMx);
                     T.PutValue('PSS_ORGCOLLECTPGU',NewCCMxGUid);
                end;
                T.PutValue('PSS_ORGCOLLECTS',NewCCMx);
                T.PutValue('PSS_ORGCOLLECTSGU',NewCCMxGUid);
               end;
          end
          else
          begin
                T.PutValue('PSS_ORGCOLLECTP',NewCCMx);
                T.PutValue('PSS_ORGCOLLECTS',NewCCMx);
                T.PutValue('PSS_ORGCOLLECTPGU',NewCCMxGUid);
                T.PutValue('PSS_ORGCOLLECTSGU',NewCCMxGUid);
          end;

          T.InsertDB(Nil);
          MoveCurProgressForm(Ancien);
     end;
     FiniMoveProgressForm;
     TobSessionStage.Free;
     TobFSession.Free;

     TobFormations := Tob.Create('FORMATIONS',Nil,-1);
     TobFForm := Tob.Create('Les formations',Nil,-1);
     TOBLoadFromFile(Chemin+'\FORMATIONS.txt', nil, TobFForm);
     InitMoveProgressForm(nil, 'Récupération des formations', 'Veuillez patienter SVP ...', TobFForm.Detail.Count - 1, FALSE, TRUE);
     For i := 0 to TobFForm.Detail.Count - 1 do
     begin
          Salarie := TobFForm.Detail[i].GetValue('PFO_SALARIE');
          TC := TobSalaries.FindFirst(['CCMX'],[Salarie],False);
          T := Tob.Create('FORMATIONS',TobFormations,-1);
          T.PutValue('PFO_MILLESIME','0000');
          matriculeSal := TC.GetValue('MATRICULE');
          T.PutValue('PFO_SALARIE',matriculeSal);
          T.PutValue('PFO_NOMSALARIE',TC.GetValue('NOM'));
          T.PutValue('PFO_PRENOM',TC.GetValue('PRENOM'));
          IAncien := TobFForm.Detail[i].GetValue('PFO_ORDRE');
          Ancien := TobFForm.Detail[i].GetValue('PFO_CODESTAGE');
          TC := TobConvertSession.FindFirst(['ANCIEN','ANCSTAGE'],[IAncien,Ancien],False);
          INouveau := TC.GetValue('NOUVEAU');
          T.PutValue('PFO_CENTREFORM',TC.GetValue('CENTREFORM'));
               T.PutValue('PFO_CENTREFORMGU',TC.GetValue('GUID'));
          T.PutValue('PFO_ORDRE',INouveau);
          TC := TobConvertStage.FindFirst(['ANCIEN'],[Ancien],False);
          Nouveau := TC.GetValue('NOUVEAU');
          T.PutValue('PFO_CODESTAGE',Nouveau);
          T.PutValue('PFO_DATEDEBUT',TobFForm.Detail[i].GetValue('PFO_DATEDEBUT'));
          T.PutValue('PFO_DATEFIN',TobFForm.Detail[i].GetValue('PFO_DATEFIN'));
          T.PutValue('PFO_EFFECTUE',TobFForm.Detail[i].GetValue('PFO_EFFECTUE'));
          T.PutValue('PFO_NATUREFORM',TobFForm.Detail[i].GetValue('PFO_NATUREFORM'));
          T.PutValue('PFO_TYPEPLANPREV',TobFForm.Detail[i].GetValue('PFO_TYPEPLANPREV'));
          T.PutValue('PFO_TYPOFORMATION',TobFForm.Detail[i].GetValue('PFO_TYPOFORMATION'));
          T.PutValue('PFO_NBREHEURE',TobFForm.Detail[i].GetValue('PFO_NBREHEURE'));
          T.PutValue('PFO_HTPSTRAV',TobFForm.Detail[i].GetValue('PFO_HTPSTRAV'));
          T.PutValue('PFO_HTPSNONTRAV',TobFForm.Detail[i].GetValue('PFO_HTPSNONTRAV'));
          T.PutValue('PFO_COUTREELSAL',TobFForm.Detail[i].GetValue('PFO_COUTREELSAL'));
          T.PutValue('PFO_FRAISREEL',TobFForm.Detail[i].GetValue('PFO_FRAISREEL'));
          T.PutValue('PFO_AUTRECOUT',TobFForm.Detail[i].GetValue('PFO_AUTRECOUT'));
          T.PutValue('PFO_FRAISPLAF',TobFForm.Detail[i].GetValue('PFO_FRAISPLAF'));
          T.PutValue('PFO_ALLOCFORM',TobFForm.Detail[i].GetValue('PFO_ALLOCFORM'));
          T.PutValue('PFO_FORMATION1','12');

          Q := OpenSQL('SELECT PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4'+
          ',PSA_CODESTAT,PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4'+
          ',PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI FROM SALARIES WHERE PSA_SALARIE="'+matriculeSal+'"',True);
          If Not Q.Eof then
          begin                                               
          T.PutValue('PFO_TRAVAILN1',Q.FindField('PSA_TRAVAILN1').AsString);
          T.PutValue('PFO_TRAVAILN2',Q.FindField('PSA_TRAVAILN2').AsString);
          T.PutValue('PFO_TRAVAILN3',Q.FindField('PSA_TRAVAILN3').AsString);
          T.PutValue('PFO_TRAVAILN4',Q.FindField('PSA_TRAVAILN4').AsString);
          T.PutValue('PFO_CODESTAT',Q.FindField('PSA_CODESTAT').AsString);
          T.PutValue('PFO_LIBREPCMB1',Q.FindField('PSA_LIBREPCMB1').AsString);
          T.PutValue('PFO_LIBREPCMB2',Q.FindField('PSA_LIBREPCMB2').AsString);
          T.PutValue('PFO_LIBREPCMB3',Q.FindField('PSA_LIBREPCMB3').AsString);
          T.PutValue('PFO_LIBREPCMB4',Q.FindField('PSA_LIBREPCMB4').AsString);
          T.PutValue('PFO_ETABLISSEMENT',Q.FindField('PSA_ETABLISSEMENT').AsString);
          T.PutValue('PFO_LIBEMPLOIFOR',Q.FindField('PSA_LIBELLEEMPLOI').AsString);
          end;
          Ferme(Q);



          NomOrg := TobFForm.Detail[i].GetValue('ORGCOLLECT');
          TC := TobConvertAnn.FindFirst(['ANCIEN'],[NomOrg],False);
          If TC <> Nil then
          begin
               if NomOrg = 'FONGECIF' then
               begin
                 T.PutValue('PFO_ORGCOLLECT',TC.GetValue('NOUVEAU'));
               T.PutValue('PFO_ORGCOLLECTGU',TC.GetValue('GUID'));
               end
               else
               begin
                 T.PutValue('PFO_ORGCOLLECT',NewCCMx);
                 T.PutValue('PFO_ORGCOLLECTGU',NewCCMxGUid);
               end;
          end
          else
          begin
                T.PutValue('PFO_ORGCOLLECT',NewCCMx);
                T.PutValue('PFO_ORGCOLLECTGU',NewCCMxGUid);
          end;

          T.InsertOrUpdateDB(False);
          QS:=OpenSQL('SELECT * FROM SESSIONSTAGE WHERE '+
            'PSS_CODESTAGE="'+Nouveau+'" AND PSS_MILLESIME="0000"'+
            ' AND PSS_ORDRE='+IntToStr(INouveau)+'',True);
           TobSessions := Tob.create('SESSIONSTAGE',Nil,-1);
           TobSessions.LoadDetailDB('SESSIONSTAGE','','',QS,False);
           Ferme(QS);
           QF := OpenSQL('SELECT SUM(PFO_AUTRECOUT) AS MONTANT FROM FORMATIONS WHERE '+
            'PFO_CODESTAGE="'+Nouveau+'" AND PFO_MILLESIME="0000"'+
            ' AND PFO_ORDRE='+IntToStr(INouveau)+'',True);
          Montant := 0;
          if not QF.eof then Montant := QF.FindField('MONTANT').AsFloat;
          Ferme(QF);
          TS := TobSessions.FindFirst([''],[''],False);
          If TS <> Nil then
          begin
                TS.PutValue('PSS_COUTPEDAG',Montant);
                TS.UpdateDB(False);
          end;
          TobSessions.Free;
          MoveCurProgressForm(matriculeSal);
     end;
     FiniMoveProgressForm;
     TobFForm.Free;
     TobFAnim := Tob.Create('Les aniamteurs',Nil,-1);
     TOBLoadFromFile(Chemin+'\SESSIONANIMAT.txt', nil, TobFAnim);
     TobAnimateurs := Tob.Create('SESSIONANIMAT',Nil,-1);
     InitMoveProgressForm(nil, 'Récupération des animateurs', 'Veuillez patienter SVP ...', TobFAnim.Detail.Count - 1, FALSE, TRUE);
     For i := 0 to TobFAnim.Detail.Count - 1 do
     begin
          Salarie := TobFAnim.Detail[i].GetValue('PAN_SALARIE');
          TC := TobSalaries.FindFirst(['CCMX'],[Salarie],False);
          T := Tob.Create('SESSIONANIMAT',TobAnimateurs,-1);
          T.PutValue('PAN_MILLESIME','0000');
          matriculeSal := TC.GetValue('MATRICULE');
          T.PutValue('PAN_SALARIE',matriculeSal);
          IAncien := TobFAnim.Detail[i].GetValue('PAN_ORDRE');
          Ancien := TobFAnim.Detail[i].GetValue('PAN_CODESTAGE');
          TC := TobConvertSession.FindFirst(['ANCIEN','ANCSTAGE'],[IAncien,Ancien],False);
          INouveau := TC.GetValue('NOUVEAU');
          T.PutValue('PAN_ORDRE',INouveau);
          TC := TobConvertStage.FindFirst(['ANCIEN'],[Ancien],False);
          Nouveau := TC.GetValue('NOUVEAU');
          T.PutValue('PAN_CODESTAGE',Nouveau);
          T.PutValue('PAN_DATEDEBUT',TobFAnim.Detail[i].GetValue('PAN_DATEDEBUT'));
          T.PutValue('PAN_DATEFIN',TobFAnim.Detail[i].GetValue('PAN_DATEFIN'));
          T.PutValue('PAN_NBREHEURE',TobFAnim.Detail[i].GetValue('PAN_NBREHEURE'));
          T.PutValue('PAN_SALAIREANIM',TobFAnim.Detail[i].GetValue('PAN_SALAIREANIM'));
          T.InsertOrUpdateDB(False);
          QS:=OpenSQL('SELECT * FROM SESSIONSTAGE WHERE '+
            'PSS_CODESTAGE="'+Nouveau+'" AND PSS_MILLESIME="0000"'+
            ' AND PSS_ORDRE='+IntToStr(INouveau)+'',True);
           TobSessions := Tob.create('SESSIONSTAGE',Nil,-1);
           TobSessions.LoadDetailDB('SESSIONSTAGE','','',QS,False);
           Ferme(QS);
           QF := OpenSQL('SELECT SUM(PAN_SALAIREANIM) AS MONTANT FROM SESSIONANIMAT WHERE '+
            'PAN_CODESTAGE="'+Nouveau+'" AND PAN_MILLESIME="0000"'+
            ' AND PAN_ORDRE='+IntToStr(INouveau)+'',True);
          Montant := 0;
          if not QF.eof then Montant := QF.FindField('MONTANT').AsFloat;
          Ferme(QF);
          TS := TobSessions.FindFirst([''],[''],False);
          If TS <> Nil then
          begin
                TS.PutValue('PSS_COUTSALAIR',Montant);
                TS.UpdateDB(False);
          end;
          TobSessions.Free;
          MoveCurProgressForm(matriculeSal);
     end;
     FiniMoveProgressForm;
     TobConvertAnn.Free;
     TobAnimateurs.Free;
     TobFormations.Free;
     TobSalaries.Free;
     TobConvertStage.Free;
     TobConvertSession.Free;



     TobFAnim.Free;



  TotalHeures := 0;
  CoutAnim := 0;
  CoutGlobal := 0;
  CoutParStagiaire := 0;
  Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_COUTSALAIR,PSS_COUTPEDAG,PSS_COUTUNITAIRE FROM SESSIONSTAGE ', True); //DB2
  TobSessions := Tob.create('SESSIONSTAGE',Nil,-1);
  TobSessions.LoadDetailDB('SESSIONSTAGE','','',QS,False);
  Ferme(Q);
  InitMoveProgressForm(nil, 'calcul des coûts des sessions', 'Veuillez patienter SVP ...', TobSessions.Detail.Count - 1, FALSE, TRUE);
  For i := 0 to TobSessions.Detail.Count - 1 do
  begin
    CoutAnim := TobSessions.Detail[i].getValue('PSS_COUTSALAIR');
    CoutGlobal := TobSessions.Detail[i].getValue('PSS_COUTPEDAG');
    CoutParStagiaire := TobSessions.Detail[i].getValue('PSS_COUTUNITAIRE');
    LeStage := TobSessions.Detail[i].getValue('PSS_CODESTAGE');
    LaSession := IntToStr(TobSessions.Detail[i].getValue('PSS_ORDRE'));
    Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) AS NBHEURE FROM FORMATIONS WHERE PFO_CODESTAGE="' + LeStage + '" AND PFO_ORDRE=' + LaSession + ' AND PFO_MILLESIME="0000"', True);
    if not Q.eof then TotalHeures := Q.FindField('NBHEURE').AsFloat;
    Ferme(Q);
    Q := OpenSQL('SELECT * FROM FORMATIONS WHERE PFO_CODESTAGE="' + LeStage + '" AND PFO_ORDRE=' + LaSession + ' AND PFO_MILLESIME="0000" ', True); //DB2
    TobFormation := Tob.Create('FORMATIONS', nil, -1);
    TobFormation.LoadDetailDB('FORMATIONS', '', '', Q, False);
    Ferme(Q);
    for f := 0 to TobFormation.Detail.Count - 1 do
    begin
      CoutSalarie := 0;
      T := TobFormation.Detail[f];
      if TotalHeures > 0 then
      begin
        HeuresSal := T.GetValue('PFO_NBREHEURE');
        CoutSalarie := CoutAnim;
        CoutSalarie := (CoutSalarie * HeuresSal) / TotalHeures;
      end;
      CoutSalarie := CoutSalarie + T.GetValue('PFO_AUTRECOUT');
      T.PutValue('PFO_AUTRECOUT', CoutSalarie);
      T.UpdateDB(False);
    end;
    TobFormation.Free;
    MoveCurProgressForm(LeStage);
   end;
   FiniMoveProgressForm;
   CommitTrans;
   PGIBox('L''import est terminé', 'Import formation');
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la mise à jour', 'Import formation');
  end;
end;




procedure TOF_PGIMPORTFORMATION.OnUpdate ;
begin
  Inherited ;
  ImportForm;
end ;

procedure TOF_PGIMPORTFORMATION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGIMPORTFORMATION.OnArgument (S : String ) ;
var edit : THEdit;
begin
  Inherited ;
  Edit := THEdit(GetControl('REPERTOIRE'));
  If Edit <> Nil then Edit.OnChange := ChangeChemin;
end ;

procedure TOF_PGIMPORTFORMATION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGIMPORTFORMATION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGIMPORTFORMATION.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGIMPORTFORMATION ] ) ; 
end.


