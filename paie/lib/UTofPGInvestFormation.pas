{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 15/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : INVESTFORMATION ()
Mots clefs ... : TOF;PGINVESTFORMATION
*****************************************************************
PT1 19/11/2003  V_50 JL Ajout calcul colonne validation organisme collecteur
PT2 19/11/2003  V_50 JL Suppression saisie des taux de charge
---- JL 20/03/2006 modification clé annuaire ----
PT3 04/04/2006 V_65 FQ 13008 Conversion de type variant ioncorrect
PT4 09/08/2007 V_80 Supp frais fixe
}


Unit UTofPGInvestFormation ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HTB97,UTOB,PGOutilsFormation,EntPaie,HSysMenu,Grids,HStatus,ed_tools ;

Type
  TOF_PGINVESTFORMATION = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    GInvest : THGrid;
    Millesime : String;
    procedure AfficherGrille;
    procedure MajInvestissement(Sender : TObject);
    procedure CalculerMasseSalariale (Sender : TObject);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure MajMontants(Sender : TObject);
    procedure MajGrille;
    procedure ChgtValeurTablette(Sender : Tobject);
  end ;

Implementation


procedure TOF_PGINVESTFORMATION.ChgtValeurTablette(Sender : Tobject);
var Q : TQuery;
    TobSession,TobInvest,T,TobMaj : Tob;
    i,s,Rep : Integer;
    DD,DF : TDateTime;
    Finance,CodeStage : String;
    Ordre : Integer;
begin
     Q := OpenSQL('SELECT * FROM INVESTFORMATION WHERE PIF_MILLESIME="'+Millesime+'"',True);
     TobInvest := Tob.Create('INVESTFORMATION',Nil,-1);
     TobInvest.LoadDetailDB('INVESTFORMATION', '', '', Q, False);
     Ferme(Q);
     InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TobInvest.Detail.Count, FALSE, TRUE);
     InitMove(TobInvest.Detail.Count, '');
     For i := 0 to TobInvest.Detail.Count - 1 do
     begin
          TobMaj := Tob.Create('INVESTFORMATION',Nil,-1);
          T := Tob.Create('INVESTFORMATION',TobMaj,-1);
          Finance := TobInvest.Detail[i].GetValue('PIF_INVESTFORM');
          If Finance = 'PDF' then T.PutValue('PIF_INVESTFORM','PLF')
          else if Finance = 'ALT'then T.PutValue('PIF_INVESTFORM','PRO')
          else continue;
          T.PutValue('PIF_MILLESIME',TobInvest.Detail[i].GetValue('PIF_MILLESIME'));
          T.PutValue('PIF_ORGCOLLECT',TobInvest.Detail[i].GetValue('PIF_ORGCOLLECT'));
          T.PutValue('PIF_POURCENTAGE',TobInvest.Detail[i].GetValue('PIF_POURCENTAGE'));
          T.PutValue('PIF_MTVERSEMENT',TobInvest.Detail[i].GetValue('PIF_MTVERSEMENT'));
          T.PutValue('PIF_MTREALISE',TobInvest.Detail[i].GetValue('PIF_MTREALISE'));
          T.PutValue('PIF_MTREALISEOPCA',TobInvest.Detail[i].GetValue('PIF_MTREALISEOPCA'));
          T.PutValue('PIF_ORGCOLLECTGU',TobInvest.Detail[i].GetValue('PIF_ORGCOLLECTGU'));
          T.InsertDb(Nil);
          MoveCurProgressForm('Financcement : '+Finance);
     end;
     TobMaj.Free;
     FiniMoveProgressForm;
     For i := 0 to TobInvest.Detail.Count - 1 do
     begin
          Finance := TobInvest.Detail[i].GetValue('PIF_INVESTFORM');
          If (Finance = 'PDF') or (Finance = 'ALT') then TobInvest.Detail[i].DeleteDB(False);
     end;
     TobInvest.Free;
     DD := IDate1900;
     DF := IDate1900;
     Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime+'"',True);
     If Not Q.Eof then
     begin
          DD := Q.FindField('PFE_DATEDEBUT').AsDateTime;
          DF := Q.FindField('PFE_DATEFIN').AsDateTime;
     end;
     Ferme(Q);
     Q := OpenSQL('SELECT PIS_CODESTAGE,PIS_INVESTFORM,PIS_ORDRE FROM INVESTSESSION JOIN SESSIONSTAGE '+
     'ON PIS_MILLESIME=PSS_MILLESIME AND PSS_CODESTAGE=PIS_CODESTAGE AND PIS_ORDRE=PSS_ORDRE '+
     'WHERE PSS_DATEDEBUT>="'+UsDateTime(DD)+'"'+
     ' AND PSS_DATEDEBUT<="'+UsDateTime(DF)+'"',True);
     TobSession := Tob.Create('les sessions',Nil,-1);
     TobSession.LoadDetailDB('les sessions', '', '', Q, False);
     Ferme(Q);
     InitMoveProgressForm(nil, 'Modification des sessions', 'Veuillez patienter SVP ...', TobSession.Detail.Count, FALSE, TRUE);
     InitMove(TobSession.Detail.Count, '');
     For s := 0 to TobSession.Detail.Count - 1 do
     begin
          Finance := TobSession.Detail[s].GetValue('PIS_INVESTFORM');
          Ordre := TobSession.Detail[s].GetValue('PIS_ORDRE');
          CodeStage := TobSession.Detail[s].GetValue('PIS_CODESTAGE');
          Q := OpenSQL('SELECT * FROM INVESTSESSION WHERE '+
          'PIS_CODESTAGE="'+CodeStage+'" AND PIS_ORDRE='+IntToStr(Ordre),True);
          TobInvest := Tob.Create('INVESTSESSION',Nil,-1);
          TobInvest.LoadDetailDB('INVESTSESSION', '', '', Q, False);
          Ferme(Q);
          For i := 0 to TobInvest.Detail.Count - 1 do
          begin
               TobMaj := Tob.Create('INVESTSESSION',Nil,-1);
               T := Tob.Create('INVESTSESSION',TobMaj,-1);
               Finance := TobInvest.Detail[i].GetValue('PIS_INVESTFORM');
               If Finance = 'PDF' then T.PutValue('PIS_INVESTFORM','PLF')
               else If Finance = 'ALT' then T.PutValue('PIS_INVESTFORM','PRO')
               else continue;
               T.PutValue('PIS_CODESTAGE',TobInvest.Detail[i].GetValue('PIS_CODESTAGE'));
               T.PutValue('PIS_ORDRE',TobInvest.Detail[i].GetValue('PIS_ORDRE'));
               T.PutValue('PIS_MILLESIME',TobInvest.Detail[i].GetValue('PIS_MILLESIME'));
               T.PutValue('PIS_INVESTFORM',Finance);
               T.PutValue('PIS_COUTPEDAG',TobInvest.Detail[i].GetValue('PIS_COUTPEDAG'));
               T.PutValue('PIS_TAUXPEDAG',TobInvest.Detail[i].GetValue('PIS_TAUXPEDAG'));
               T.PutValue('PIS_COUTSALAIRE',TobInvest.Detail[i].GetValue('PIS_COUTSALAIRE'));
               T.PutValue('PIS_TAUXSALAIRE',TobInvest.Detail[i].GetValue('PIS_TAUXSALAIRE'));
               T.PutValue('PIS_FRAISFORM',TobInvest.Detail[i].GetValue('PIS_FRAISFORM'));
               T.PutValue('PIS_TAUXSALFRAIS',TobInvest.Detail[i].GetValue('PIS_TAUXSALFRAIS'));
               T.PutValue('PIS_MONTANT',TobInvest.Detail[i].GetValue('PIS_MONTANT'));
               T.InsertOrUpdateDb(False);

               end;
          TOBMAJ.free;
          For i := 0 to TobInvest.Detail.Count - 1 do
          begin
               Finance := TobInvest.Detail[i].GetValue('PIS_INVESTFORM');
               If (Finance = 'PDF') or (Finance='ALT') then TobInvest.Detail[i].DeleteDB(False);
          end;
          TobInvest.Free;
          MoveCurProgressForm('Stage : '+CodeStage);
     end;
     TobSession.Free;
     FiniMoveProgressForm;
     //LesSatges
     Rep := PGIAsk ('Voulez vous mettre à jour le catalogue ?','Mise à jour des types d''investissement');
     If Rep = MrYes then
     begin
       Q := OpenSQL('SELECT PIS_CODESTAGE,PIS_INVESTFORM,PIS_ORDRE FROM INVESTSESSION JOIN STAGE '+
      'ON PST_CODESTAGE=PIS_CODESTAGE AND PST_MILLESIME=PIS_MILLESIME '+
      'WHERE PIS_ORDRE=-1 AND (PST_MILLESIME="0000" OR PST_MILLESIME="'+Millesime+'")',True);
      TobSession := Tob.Create('les sessions',Nil,-1);
      TobSession.LoadDetailDB('les sessions', '', '', Q, False);
      Ferme(Q);
      InitMoveProgressForm(nil, 'Modification du catalogue', 'Veuillez patienter SVP ...', TobSession.Detail.Count, FALSE, TRUE);
      InitMove(TobSession.Detail.Count, '');
      For s := 0 to TobSession.Detail.Count - 1 do
      begin
          Finance := TobSession.Detail[s].GetValue('PIS_INVESTFORM');
          CodeStage := TobSession.Detail[s].GetValue('PIS_CODESTAGE');
          Q := OpenSQL('SELECT * FROM INVESTSESSION WHERE '+
          'PIS_CODESTAGE="'+CodeStage+'" AND PIS_ORDRE=-1',True);
          TobInvest := Tob.Create('INVESTSESSION',Nil,-1);
          TobInvest.LoadDetailDB('INVESTSESSION', '', '', Q, False);
          Ferme(Q);
          For i := 0 to TobInvest.Detail.Count - 1 do
          begin
               TobMaj := Tob.Create('INVESTSESSION',Nil,-1);
               T := Tob.Create('INVESTSESSION',TobMaj,-1);
               Finance := TobInvest.Detail[i].GetValue('PIS_INVESTFORM');
               If Finance = 'PDF' then T.PutValue('PIS_INVESTFORM','PLF')
               else If Finance = 'ALT' then T.PutValue('PIS_INVESTFORM','PRO')
               else continue;
               T.PutValue('PIS_CODESTAGE',TobInvest.Detail[i].GetValue('PIS_CODESTAGE'));
               T.PutValue('PIS_ORDRE',TobInvest.Detail[i].GetValue('PIS_ORDRE'));
               T.PutValue('PIS_MILLESIME',TobInvest.Detail[i].GetValue('PIS_MILLESIME'));
               T.PutValue('PIS_COUTPEDAG',TobInvest.Detail[i].GetValue('PIS_COUTPEDAG'));
               T.PutValue('PIS_TAUXPEDAG',TobInvest.Detail[i].GetValue('PIS_TAUXPEDAG'));
               T.PutValue('PIS_COUTSALAIRE',TobInvest.Detail[i].GetValue('PIS_COUTSALAIRE'));
               T.PutValue('PIS_TAUXSALAIRE',TobInvest.Detail[i].GetValue('PIS_TAUXSALAIRE'));
               T.PutValue('PIS_FRAISFORM',TobInvest.Detail[i].GetValue('PIS_FRAISFORM'));
               T.PutValue('PIS_TAUXSALFRAIS',TobInvest.Detail[i].GetValue('PIS_TAUXSALFRAIS'));
               T.PutValue('PIS_MONTANT',TobInvest.Detail[i].GetValue('PIS_MONTANT'));
               T.InsertOrUpdateDb(False);
               end;
          TOBMAJ.free;
          For i := 0 to TobInvest.Detail.Count - 1 do
          begin
               Finance := TobInvest.Detail[i].GetValue('PIS_INVESTFORM');
               If (Finance = 'PDF') or (Finance='ALT') then TobInvest.Detail[i].DeleteDB(False);
          end;
          TobInvest.Free;
          MoveCurProgressForm('Stage : '+CodeStage);

      end;
      TobSession.Free;
      FiniMoveProgressForm;
     end;
     GInvest.ColFormats[0] := 'CB=PGTYPEPLANPREV';
     AfficherGrille;
end;

procedure TOF_PGINVESTFORMATION.OnArgument (S : String ) ;
var BValider,BCalcMasse,BMajMontant,Bt : TToolBarButton97;
    Num : Integer;
    HMTrad : THSystemMenu;
    Tablette : String;
begin
  Inherited ;
        Millesime := ReadTokenPipe(S,';');
        GInvest := THGrid(GetControl('GINVEST'));
        BValider := TToolBarButton97(GetControl('BVAL'));
        If BValider <> Nil then BValider.OnClick := MajInvestissement;
        If ExisteSQL('SELECT PIF_INVESTFORM FROM INVESTFORMATION WHERE PIF_MILLESIME="'+Millesime+'"'+
        ' AND (PIF_INVESTFORM="PDF" OR PIF_INVESTFORM="ALT")') then
        begin
             SetControlVIsible('BREAFFECTATION',True);
             Tablette := 'PGFINANCEFORMATION';
        end
        Else
        begin
             SetControlVIsible('BREAFFECTATION',True);
             Tablette := 'PGTYPEPLANPREV';
//             Tablette := 'PGFINANCEFORMATION';
        end;
        //GInvest.ColFormats[0] := 'CB=PGFINANCEFORMATION';
        GInvest.ColFormats[0] := 'CB='+Tablette;
        GInvest.ColFormats[1] := 'CB=PGORGCOLLECTEUR';
        GInvest.ColFormats[2] := '# ##0.00';
        GInvest.ColFormats[3] := '# ##0.00';
        GInvest.ColFormats[4] := '# ##0.00';
        GInvest.ColFormats[5] := '# ##0.00';
        AfficherGrille;
        GInvest.OnCellExit := OnCellExit;
        BCalcMasse := TToolBarButton97(GetControl('BCALCMASSE'));
        If BCalcMasse <> Nil then BCalcMasse.OnClick := CalculerMasseSalariale;
        Ecran.Caption := 'Saisie des investissements pour le millésime '+Millesime;
        UpdateCaption(Ecran);
        BMajMontant := TToolBarButton97(GetControl('BMAJ'));
        If BMajMontant <> Nil then BMajMontant.OnClick := MajMontants;
        HMTrad:=THSystemMenu(GetControl('HMTrad'));
        HMTrad.ResizeGridColumns(GInvest);
        Bt := TToolBarButton97(GetControl('BREAFFECTATION'));
        If Bt <> Nil then Bt.OnClick := ChgtValeurTablette;
end ;

procedure TOF_PGINVESTFORMATION.AfficherGrille;
var i : Integer;
    Q : TQuery;
    MasseSal,MasseSalCDD : Double;
    TobInvest : Tob;
begin
        For i := 1 to GInvest.ColCount - 1 do
        begin
                GInvest.CellValues[0,i] := '';
                GInvest.CellValues[1,i] := '';
        end;
        Q := OpenSQL('SELECT PIF_INVESTFORM,PIF_MILLESIME,PIF_ORGCOLLECTGU,PIF_POURCENTAGE,PIF_MTVERSEMENT,PIF_MTREALISE,PIF_MTREALISEOPCA FROM INVESTFORMATION '+
        'WHERE PIF_MILLESIME="'+Millesime+'"',True);
        TobInvest := Tob.Create('Les investissements',Nil,-1);
        TobInvest.LoadDetailDB('INVESTFORMATION','','',Q,False);
        Ferme(Q);
        For i := 0 to TobInvest.Detail.Count - 1 do
        begin
                GInvest.CellValues[0,i+1] :=  TobInvest.Detail[i].GetValue('PIF_INVESTFORM');
                GInvest.CellValues[1,i+1] :=  TobInvest.Detail[i].GetValue('PIF_ORGCOLLECTGU');
                GInvest.CellValues[2,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIF_POURCENTAGE'));
                GInvest.CellValues[3,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIF_MTVERSEMENT'));
                GInvest.CellValues[4,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIF_MTREALISE'));
                GInvest.CellValues[5,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIF_MTREALISEOPCA'));
        end;
        TobInvest.Free;
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN,PFE_MASSESAL,PFE_MASSESALCDD'+
        ',PFE_COUTFIXEFOR7,PFE_COUTFIXEFOR8,PFE_TAUXCHARGEC,PFE_TAUXCHARGENC,'+ //PT4
        'PFE_OPCACTFIXE1,PFE_OPCACTFIXE2,PFE_OPCACTFIXE3,PFE_OPCACTFIXE4,PFE_OPCACTFIXE5,PFE_OPCACTFIXE6,PFE_OPCACTFIXE7,PFE_OPCACTFIXE8 '+
        'FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime+'"',true);
        If Not Q.Eof then
        begin
                MasseSal := Q.FindField('PFE_MASSESAL').AsFloat;
                MasseSalCDD := Q.FindField('PFE_MASSESALCDD').AsFloat;
                SetControlText('DATEDEBPAIE',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                SetControlText('DATEFINPAIE',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
                For i := 1 to 8 do
                begin
                        SetControlText('ORGCOLLECT'+IntToStr(i),IntToStr(Q.FindField('PFE_OPCACTFIXE'+IntToStr(i)).AsInteger));
                end;
        end
        Else
        begin
                MasseSal := 0;
                MasseSalCDD := 0;
                SetControlText('DATEDEBPAIE',DateToStr(IDate1900));
                SetControlText('DATEFINPAIE',DateToStr(IDate1900));
                For i := 1 to 8 do
                begin
                        SetControlText('ORGCOLLECT1','');
                end;
        end;
        Ferme(Q);
        SetControlText('MASSESALARIALE',FloatToStr(MasseSal));
        SetControlText('MASSESALARIALECDD',FloatToStr(MasseSalCDD));
end;

procedure TOF_PGINVESTFORMATION.MajInvestissement(Sender : TObject);
var Q:Tquery;
    TobInvest,TI,TobExerFormation : tob;
    i,j : Integer;
    MessError : String;
begin
        MessError := '';
        If Not IsNumeric(GetControlText('MASSESALARIALE')) then MessError := Messerror + '- #13#10 La masse salariale doit être numérique';
        If Not IsNumeric(GetControlText('MASSESALARIALECDD')) then MessError := Messerror + '- #13#10 La masse salariale CDD doit être numérique';
        If MessError <> '' then
        begin
                PGIBox('Validation impossible car '+MessError,Ecran.Caption);
                Exit;
        end;
  //      If PGIAskCancel('Voulez vous mettre à jour',Ecran.Caption) <> MrYes Then Exit;
        ExecuteSQL('DELETE FROM INVESTFORMATION WHERE PIF_MILLESIME="'+Millesime+'"');
        Q := OpenSQL('SELECT * FROM INVESTFORMATION '+
        'WHERE PIF_MILLESIME="'+Millesime+'"',True);
        TobInvest := Tob.Create('Les investissments',Nil,-1);
        TobInvest.LoadDetailDB('INVESTFORMATION','','',Q,False);
        Ferme(Q);
        For i := 1 to GInvest.RowCount - 1 do
        begin
                If (GInvest.CellValues[0,i] <> '') AND (GInvest.CellValues[0,i] <> 'AUC') then
                begin
                        TI := Tob.Create('INVESTFORMATION',TobInvest,-1);
                        TI.PutValue('PIF_INVESTFORM',GInvest.CellValues[0,i]);
                        TI.PutValue('PIF_MILLESIME',Millesime);
                        TI.PutValue('PIF_ORGCOLLECTGU',GInvest.CellValues[1,i]);
                        If (GInvest.CellValues[2,i] <> '') and (IsNumeric(GInvest.CellValues[2,i])) then TI.PutValue('PIF_POURCENTAGE',StrToFloat(GInvest.CellValues[2,i]))
                        else TI.PutValue('PIF_POURCENTAGE',0); //PT3
                        If (GInvest.CellValues[2,i] <> '') and (IsNumeric(GInvest.CellValues[3,i])) then TI.PutValue('PIF_MTVERSEMENT',StrToFloat(GInvest.CellValues[3,i]))
                        else TI.PutValue('PIF_MTVERSEMENT',0); //PT3
                        If (GInvest.CellValues[2,i] <> '') and (IsNumeric(GInvest.CellValues[4,i])) then TI.PutValue('PIF_MTREALISE',StrToFloat(GInvest.CellValues[4,i]))
                        else TI.PutValue('PIF_MTREALISE',0);        //PT3
                        If (GInvest.CellValues[2,i] <> '') and (IsNumeric(GInvest.CellValues[5,i])) then TI.PutValue('PIF_MTREALISEOPCA',StrToFloat(GInvest.CellValues[5,i]))
                        else TI.PutValue('PIF_MTREALISEOPCA',0);         //PT3
                        TI.InsertDB(Nil,False);
                end;
        end;
        TobInvest.Free;
        Q := OpenSQL('SELECT * FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime+'"',True);
        TobExerFormation := Tob.Create ('EXERFORMATION',Nil,-1);
        TobExerFormation.LoadDetailDB('EXERFORMATION','','',Q,False);
        Ferme(Q);
        For i := 0 To TobExerFormation.Detail.Count - 1 do
        begin
                For j := 1 to 8 do
                begin
                        if GetControlText('ORGCOLLECT'+IntToStr(j)) <> '' then TobExerFormation.Detail[i].PutValue('PFE_OPCACTFIXE'+IntToStr(j),StrToInt(GetControlText('ORGCOLLECT'+IntToStr(j))))
                        else TobExerFormation.Detail[i].PutValue('PFE_OPCACTFIXE'+IntToStr(j),-1)
                end;
                TobExerFormation.Detail[i].PutValue('PFE_MASSESAL',StrToFloat(GetControlText('MASSESALARIALE')));
                TobExerFormation.Detail[i].PutValue('PFE_MASSESALCDD',StrToFloat(GetControlText('MASSESALARIALECDD')));
                TobExerFormation.Detail[i].UpdateDB(False);
        end;
        TobExerFormation.Free;
end;

procedure TOF_PGINVESTFORMATION.CalculerMasseSalariale (Sender : TObject);
var Q : TQuery;
    TobPaie : Tob;
    i : Integer;
    Brut,BrutCDD : Double;
begin
        Brut := 0;
        BrutCDD := 0;
        Q := OpenSQL('SELECT PPU_CBRUT,PCI_TYPECONTRAT,PPU_DATEDEBUT FROM PAIEENCOURS LEFT JOIN CONTRATTRAVAIL '+
        'ON PPU_SALARIE=PCI_SALARIE AND PPU_DATEDEBUT>=PCI_DEBUTCONTRAT AND PPU_DATEDEBUT<=PCI_FINCONTRAT '+
        'WHERE PPU_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBPAIE')))+'" '+
        'AND PPU_DATEFIN<="'+UsDateTime(StrToDATE(GetControlText('DATEFINPAIE')))+'" ',True);
        TobPaie := Tob.Create('Les paie',Nil,-1);
        TobPaie.LoadDetailDB('Les paies','','',Q,False);
        Ferme(Q);
        For i := 0 to TobPaie.Detail.Count-1 do
        begin
                If TobPaie.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CCD' then
                BrutCDD := BrutCDD + TobPaie.Detail[i].GetValue('PPU_CBRUT')
                else Brut := Brut + TobPaie.Detail[i].GetValue('PPU_CBRUT');
        end;
        TobPaie.Free;
        SetControlText('MASSESALARIALE',FloatToStr(Brut));
        SetControlText('MASSESALARIALECDD',FloatToStr(BrutCDD));
end;

procedure TOF_PGINVESTFORMATION.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
        If ACol = 0 then
        begin
                If (GInvest.CellValues[0,ARow] <> '') AND (GInvest.CellValues[0,ARow] <> 'AUC') then
                begin
                        If GInvest.CellValues[2,ARow] = '' then GInvest.CellValues[2,ARow] := '0';
                        If GInvest.CellValues[3,ARow] = '' then GInvest.CellValues[3,ARow] := '0';
                        If GInvest.CellValues[4,ARow] = '' then GInvest.CellValues[4,ARow] := '0';
                        If GInvest.CellValues[5,ARow] = '' then GInvest.CellValues[5,ARow] := '0';
                end
                else
                begin
                        GInvest.CellValues[1,ARow] := '-1';
                        GInvest.CellValues[2,ARow] := '';
                        GInvest.CellValues[3,ARow] := '';
                        GInvest.CellValues[4,ARow] := '';
                        GInvest.CellValues[5,ARow] := '';
                end;
        end;
end;

procedure TOF_PGINVESTFORMATION.MajMontants(Sender : TObject);
begin
        GInvest.Options := GInvest.Options - [GoEditing];
        MajGrille;
        GInvest.Options := GInvest.Options + [GoEditing];
end;
procedure TOF_PGINVESTFORMATION.MajGrille;
var i : Integer;
    Invest,OrgCollect : String;
    Pourcentage,MasseSal : Double;
    Q : TQuery;
    DateDebut,DateFin : TDateTime;
begin

        DateDebut := IDate1900;
        DateFin := IDate1900;
        MasseSal := StrToFloat(GetControlText('MASSESALARIALE'));
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime+'"',True);
        If Not Q.Eof then
        begin
                DateDebut := Q.FindField('PFE_DATEDEBUT').AsDateTime;
                DateFin := Q.FindField('PFE_DATEFIN').AsDateTime;
        end;
        Ferme(Q);
        For i := 1 to GInvest.RowCount - 1 do
        begin
                Invest := GInvest.CellValues[0,i];
                If (Invest <> '') AND (Invest <> 'AUC') then
                begin
                        OrgCollect := GInvest.CellValues[1,i];
                        If (GInvest.CellValues[2,i]<>'') and (IsNumeric(GInvest.CellValues[2,i])) then Pourcentage := StrToFloat(GInvest.CellValues[2,i])
                        else Pourcentage := 0;
                        Pourcentage := Pourcentage / 100;
                        Q := OpenSQL('SELECT SUM (PIS_MONTANT) TOTAL FROM INVESTSESSION '+
                        'LEFT JOIN SESSIONSTAGE ON PIS_CODESTAGE=PSS_CODESTAGE AND PIS_ORDRE=PSS_ORDRE AND PIS_MILLESIME=PSS_MILLESIME '+
                        'WHERE PSS_EFFECTUE="X" AND PSS_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+'" '+
                        'AND PIS_INVESTFORM="'+Invest+'" AND (PSS_ORGCOLLECTPGU="'+OrgCollect+'" OR PSS_ORGCOLLECTSGU="'+OrgCollect+'")',True);  //DB2
                        If not Q.eof then GInvest.CellValues[4,i] := FloatToStr (Arrondi (Q.FindField('TOTAL').AsFloat,2));
                        GInvest.CellValues[3,i] := FloatToStr(Arrondi(MasseSal * Pourcentage,2));
                        Ferme(Q);
                        //DEBUT PT1
                        Q := OpenSQL('SELECT SUM (PIS_MONTANT) TOTAL FROM INVESTSESSION '+
                        'LEFT JOIN SESSIONSTAGE ON PIS_CODESTAGE=PSS_CODESTAGE AND PIS_ORDRE=PSS_ORDRE AND PIS_MILLESIME=PSS_MILLESIME '+
                        'WHERE PSS_EFFECTUE="X" AND PSS_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+'" '+
                        'AND PIS_INVESTFORM="'+Invest+'" AND PSS_VALIDORG="X" AND (PSS_ORGCOLLECTPGU="'+OrgCollect+'" OR PSS_ORGCOLLECTSgu="'+OrgCollect+'")',True);  //DB2
                        If not Q.eof then GInvest.CellValues[5,i] := FloatToStr (Arrondi (Q.FindField('TOTAL').AsFloat,2));
                        Ferme(Q);
                        //FIN PT1
                end;
        end;
end;

//procedure TOF_PGINVESTFORMATION.Imprimer(Sendeer : TObject);
//begin
  //      LanceEtatTOB('E','PFO','P',TobEtat,True,False,False,Pages,'','',False);
//end;

Initialization
  registerclasses ( [ TOF_PGINVESTFORMATION ] ) ;
end.


