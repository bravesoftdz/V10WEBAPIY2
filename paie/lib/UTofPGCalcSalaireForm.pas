{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCALCSALAIREFORM ()
Mots clefs ... : TOF;PGCALCSALAIREFORM
*****************************************************************
PT1  19/12/2003 V_50 JL Ajout maj salaire animateurs
PT2  16/11/2004 V_60 JL Correction maj cout du previsionnel
PT3  07/12/2004 V_60 JL Correction maj des frais
PT4  09/03/2005 V_60 JL Correction erreur indice hors des limites en maj
PT5  14/11/2005 V_60 JL Correction calcul cout pédagogique prévisionnel
PT6  30/06/2006 V_70 JL FQ 13337 1 repas au lieu de 2 si pas d'hébergement
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
PT7  16/04/2007 V_720 FL FQ 13568 Prise en compte du paramétrage prévisionnel pour la maj des coûts
PT8  18/04/2007 V_720 FL FQ 14071 La mise à jour des coûts doit prendre en compte l'exercice en cours
PT9  10/05/2007 V_720 FL FQ 13567 Prise en compte des populations pour le calcul des coûts prévisionnels
PT10 08/02/2008 V_803 FL Correction du traitement multidossier pour le calcul des frais
PT11 12/02/2008 V_803 FL Correction du calcul du taux horaire dans le cas du multidossier
}
Unit UTofPGCalcSalaireForm ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,
{$ELSE}
     EMul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTob,EntPaie,UTOFPGMul_SessionStage,ParamDat, UtilPGI,
     HTB97,PGOutilsFormation,HStatus,ed_tools,UTOFPGMulStage,UTobDebug,PgPopulOutils ;  //PT9

Type
  TOF_PGCALCSALAIREFORM = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    TypeCalcul,MillesimeMul : String;
    TauxChargeBudget : Double;
    procedure DateElipsisclick(Sender : TObject);
    procedure ExitMethodeValo(Sender : TObject);
    procedure ExitMethodeCalc(Sender : TObject);
    procedure LanceMajSalairesForms(Sender : TObject);
    procedure LanceMajSalairePrev(Sender : TObject);
    procedure MajTableFormations(LeStage , LeMillesime : String; LaSession : Integer);
    procedure MAJTableInscFormation(LeStage , LeMillesime : String);
    procedure AccesPeriodePaie(Sender : TObject);
    Function MajFrais(TypeFrais,LeSalarie,LeStage , LeMillesime : String; LaSession : Integer) : Double;
    Function MajSalaire(LeSalarie,LeStage , LeMillesime , LibEmploi , MillesimeEC : String; LaSession : Integer;NoDossier : String) : Double;
  end ;

Implementation

procedure TOF_PGCALCSALAIREFORM.OnArgument (S : String ) ;
var THDate : THEdit;
    THCalc,THValo : THValComboBox;
    BCalc : TToolBarButton97;
    DateDeb,DateFin : TDateTime;
    Check : TCheckBox;
begin
  Inherited ;
          TypeCalcul := Trim(ReadTokenSt(S));
        DateDeb := IDate1900;
        DateFin := IDate1900;
        BCalc := TToolBarButton97(GetControl('BLANCECALC'));
        If TypeCalcul = 'PREVISIONNEL' then
        begin
                MillesimeMul := Trim(ReadTokenSt(S));
                SetActiveTabSheet('PPREVISIONNEL');
                SetControlProperty('PCRITERES','TabVisible',False);
                SetControlProperty('PCRITERES','Visible',False);
                If BCalc <> Nil Then BCalc.OnClick := LanceMajSalairePrev;
                SetControlVisible('CBPREVPEDAG',True); //PT2
        end
        else
        begin
                RendMillesimeRealise(DateDeb,DateFin);
                SetActiveTabSheet('PCRITERES');
                SetControlProperty('PPREVISIONNEL','TabVisible',False);
                SetControlProperty('PPREVISIONNEL','Visible',False);
                If BCalc <> Nil Then BCalc.OnClick := LanceMajSalairesForms;
                SetControltext('DATEDEB',DateToStr(DateDeb));
                SetControltext('DATEFIN',DateToStr(DateFin));
                SetControlText('METHODEVALO',VH_Paie.PGForValoSalaire);
                THDate := THEdit(GetControl('DATEDEBUT'));
                If THDate <> Nil then THDate.OnElipsisClick := DateElipsisclick;
                THDate := THEdit(GetControl('DATEFIN'));
                If THDate <> Nil then THDate.OnElipsisClick := DateElipsisclick;
                THCalc := THValComboBox(GetControl('METHODECALC'));
                if THCalc <> Nil then THCalc.OnChange := ExitMethodeCalc;
                THValo := THValComboBox(GetControl('METHODEVALO'));
                if THValo <> Nil then THValo.OnChange := ExitMethodeValo;
                Check := TCheckBox(GetControl('TCHOIXPERIODE'));
                If Check <> Nil then Check.Onclick := AccesPeriodePaie;
                ExitMethodeValo(THValo);
        end ;
end;

procedure TOF_PGCALCSALAIREFORM.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGCALCSALAIREFORM.ExitMethodeValo(Sender : TObject);
begin
        SetControlText('METHODECALC',VH_Paie.PGForMethodeCalc);
        if getControlText('METHODEVALO') = 'VCC' then SetControlEnabled('METHODECALC',False)
        Else SetControlEnabled('METHODECALC',True);
end;

procedure TOF_PGCALCSALAIREFORM.ExitMethodeCalc(Sender : TObject);
begin
        If GetControlText('METHODECALC') <> 'THS' then SetControlEnabled('TCHOIXPERIODE',True)
        else SetControlEnabled('TCHOIXPERIODE',False);
end;

procedure TOF_PGCALCSALAIREFORM.LanceMajSalairesForms(Sender : TObject);
var Stage,Millesime : String;
    Session,i : Integer;
begin
        If (Grille = nil) then Exit;
        if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
        begin
                MessageAlerte('Aucun élément sélectionné');
                exit;
        end;
        if ((Grille.nbSelected) > 0) AND (not Grille.AllSelected ) then
        begin
                InitMoveProgressForm (NIL,'Calcul des coûts par formation', 'Veuillez patienter SVP ...',Grille.nbSelected,FALSE,TRUE);
                InitMove(Grille.nbSelected,'');
                for i := 0 to Grille.NbSelected-1 do
                begin
                        Grille.GotoLeBookmark(i);
                        {$IFDEF EAGLCLIENT}
                        TFM.Q.TQ.Seek(Grille.Row-1) ; //PT2
                        {$ENDIF}
                        Stage := Q_mul.FindField('PSS_CODESTAGE').AsString;
                        Session := Q_mul.FindField('PSS_ORDRE').AsInteger;
                        Millesime := q_mul.FindField('PSS_MILLESIME').AsString;
                        MajTableFormations(Stage,Millesime,Session);
                        If GetCheckBoxState('CALCCTPEDAG') = CbChecked then MAJCoutsFormation(Millesime,Stage,IntToStr(Session));
                        If GetCheckBoxState('CALCINVEST') = CbChecked then CalcCtInvestSession('TOUS',Stage,Millesime,Session);
                        MoveCurProgressForm(RechDom('PGSTAGEFORM',Stage,False));
                end;
                FiniMoveProgressForm;
                Grille.ClearSelected;
        end;
        If (Grille.AllSelected=TRUE) then
        begin
             {$IFDEF EAGLCLIENT}
             if (TFM.bSelectAll.Down) then  //PT2
             TFM.Fetchlestous;
             {$ENDIF}
                Q_mul.First;
                InitMoveProgressForm (NIL,'Calcul des coûts par formation', 'Veuillez patienter SVP ...',Q_mul.RecordCount,FALSE,TRUE);
                InitMove(Q_mul.RecordCount,'');
                while Not Q_mul.EOF do
                begin
                        Stage := Q_mul.FindField('PSS_CODESTAGE').AsString;
                        Session := Q_mul.FindField('PSS_ORDRE').AsInteger;
                        Millesime := Q_mul.FindField('PSS_MILLESIME').AsString;
                        MajTableFormations(Stage,Millesime,Session);
                        If GetCheckBoxState('CALCCTPEDAG') = CbChecked then MAJCoutsFormation(Millesime,Stage,IntToStr(Session));
                        If GetCheckBoxState('CALCINVEST') = CbChecked then CalcCtInvestSession('TOUS',Stage,Millesime,Session);
                        MoveCurProgressForm(RechDom('PGSTAGEFORM',Stage,False));
                        Q_mul.Next;
                 end;
                 FiniMoveProgressForm;
                Grille.AllSelected := False;
        end;
end;

procedure TOF_PGCALCSALAIREFORM.LanceMajSalairePrev(Sender : TObject);
var Stage,Millesime : String;
    i : Integer;
    Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="'+MillesimeMul+'"',True);
        If Not Q.Eof then TauxChargeBudget := Q.FindField('PFE_TAUXBUDGET').AsFloat
        else TauxChargeBudget := 1;
        Ferme(Q);
 //DEBUT PT2
        If (Liste = nil) then Exit;
        if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
        begin
                MessageAlerte('Aucun élément sélectionné');
                exit;
        end;
        if ((Liste.nbSelected) > 0) AND (not Liste.AllSelected ) then
        begin
                InitMoveProgressForm (NIL,'Calcul des coûts par formation', 'Veuillez patienter SVP ...',Liste.nbSelected,FALSE,TRUE);
                InitMove(Liste.nbSelected,'');
                for i := 0 to Liste.NbSelected-1 do
                begin
                        Liste.GotoLeBookmark(i);
                        {$IFDEF EAGLCLIENT}
                        TFMStage.Q.TQ.Seek(Liste.Row-1) ;
                        {$ENDIF}
                        Stage := Q_MulStage.FindField('PST_CODESTAGE').AsString;
                        Millesime := Q_MulStage.FindField('PST_MILLESIME').AsString;
                        MAJTableInscFormation(Stage,Millesime);
                        MoveCurProgressForm(RechDom('PGSTAGEFORM',Stage,False));
                end;
                FiniMoveProgressForm;
                Liste.ClearSelected;
        end;
        If (Liste.AllSelected=TRUE) then
        begin
             {$IFDEF EAGLCLIENT}
             if (TFMStage.bSelectAll.Down) then
             TFMStage.Fetchlestous;
             {$ENDIF}
                Q_MulStage.First;
                InitMoveProgressForm (NIL,'Calcul des coûts par formation', 'Veuillez patienter SVP ...',Q_MulStage.RecordCount,FALSE,TRUE);
                InitMove(Q_MulStage.RecordCount,'');
                while Not Q_MulStage.EOF do
                begin
                        Stage := Q_MulStage.FindField('PST_CODESTAGE').AsString;
                        Millesime := Q_MulStage.FindField('PST_MILLESIME').AsString;
                        MAJTableInscFormation(Stage,Millesime);
                        MoveCurProgressForm(RechDom('PGSTAGEFORM',Stage,False));
                        Q_MulStage.Next;
                 end;
                 FiniMoveProgressForm;
                Liste.AllSelected := False;
 //FIN PT2
        end;
end;

procedure TOF_PGCALCSALAIREFORM.MAJTableInscFormation(LeStage , LeMillesime : String);
var TobInsc : Tob;
    Q : TQuery;
    i : Integer;
    Salarie,LibelleEmploi,Etab,Lieu,LibelleInsc : String;
    Duree,Salaire,NbHeure : Double;
    FraisH,FraisR,FraisT,NbJour,NbInsc,TotalFrais : Double;
    CoutUnit,CoutFonc,CoutFoncSta ,NbMax,CoutAnim,CoutAnimSta,NbInscStage,NbDivD,NbDivE,NbCout : Double;
    NbheureSta,NbJourSta : Double;
    DateExercice : TDateTime; //PT8
    Req : String; //PT9
    CalcFrais : Boolean; //PT9
begin
        Q := OpenSQL('SELECT * FROM INSCFORMATION WHERE PFI_CODESTAGE="'+LeStage+'" AND PFI_MILLESIME="'+LeMillesime+'" '+
        'AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL")',True);
        TobInsc := Tob.Create('INSCFORMATION',Nil,-1);
        TobInsc.LoadDetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE ' +
                'WHERE PST_CODESTAGE="' + LeStage + '" AND PST_MILLESIME="' + LeMillesime + '"', True);
        if not Q.Eof then
        begin
                CoutFonc := Q.FindField('PST_COUTBUDGETE').AsFloat;
                CoutUnit := Q.FindField('PST_COUTUNITAIRE').AsFloat;
                CoutAnim := Q.FindField('PST_COUTSALAIR').AsFloat;
                NbMax := Q.FindField('PST_NBSTAMAX').AsInteger;
                NbheureSta := Q.FindField('PST_DUREESTAGE').AsInteger;
                NbJourSta := Q.FindField('PST_JOURSTAGE').AsInteger;
        end
        else
        begin
                CoutUnit := 0;
                CoutAnim := 0;
                Coutfonc := 0;
                NbMax := 0;
                NbheureSta := 0;
                NbJourSta := 0;
        end;
        Ferme(Q);

        //PT9 - Début
        CalcFrais := True;

        { Si la gestion des frais s'effectue par population et que celle propre à la formation
          n'est pas valide, il faut interdire la saisie des inscriptions. }
        If VH_Paie.PGForGestFraisByPop And Not CanUsePopulation(TYP_POPUL_FORM_PREV) Then
        Begin
              PGIBox('Attention, les populations de formation ne sont pas valides. #13#10'+
                     'La mise à jour des frais ne sera pas effectuée.',Ecran.Caption);
              CalcFrais := False;
        End;
        //PT9 - Fin

        for i := 0 to TobInsc.Detail.count - 1 do
        begin
                Salarie := TobInsc.Detail[i].GetValue('PFI_SALARIE');
                LibelleEmploi := TobInsc.Detail[i].GetValue('PFI_LIBEMPLOIFOR');
                Duree := TobInsc.Detail[i].GetValue('PFI_DUREESTAGE');
                Etab := TobInsc.Detail[i].GetValue('PFI_ETABLISSEMENT');
                Lieu := TobInsc.Detail[i].GetValue('PFI_LIEUFORM');
 // DEBUT PT2
                NbInsc  := TobInsc.Detail[i].GetValue('PFI_NBINSC');

                If GetCheckBoxState('CBDUREE') = CbChecked then
                begin
                        if TobInsc.Detail[i].GetValue('PFI_DUREESTAGE') <> (NbheureSta * TobInsc.Detail[i].GetValue('PFI_NBINSC')) then TobInsc.Detail[i].PutValue('PFI_DUREESTAGE', NbheureSta * TobInsc.Detail[i].GetValue('PFI_NBINSC'));
                        if TobInsc.Detail[i].GetValue('PFI_JOURSTAGE') <> (NbJourSta * TobInsc.Detail[i].GetValue('PFI_NBINSC')) then TobInsc.Detail[i].PutValue('PFI_JOURSTAGE', NbJourSta * TobInsc.Detail[i].GetValue('PFI_NBINSC'));
                        if Salarie = '' then
                        begin
                              if TobInsc.Detail[i].GetValue('PFI_NBINSC') > 1 then LibelleInsc := IntToStr(TobInsc.Detail[i].GetValue('PFI_NBINSC')) + ' salariés'
                              else LibelleInsc := IntToStr(TobInsc.Detail[i].GetValue('PFI_NBINSC')) + ' salarié';
                              If LibelleInsc <> TobInsc.Detail[i].GetValue('PFI_LIBELLE') then
                              begin
                                TobInsc.Detail[i].PutValue('PFI_LIBELLE',LibelleInsc);
                              end;
                        end;
                end;
                If GetCheckBoxState('CBPREVPEDAG') = CbChecked then
                begin
                        Q := OpenSQL('SELECT SUM(PFI_NBINSC) NBINSC FROM INSCFORMATION ' +
                          'WHERE PFI_CODESTAGE="' + LeStage + '" AND PFI_MILLESIME="' + LeMillesime + '" ' +
                          'AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL")', True);
                        if not Q.Eof then NbInscStage := Q.FindField('NBINSC').AsInteger
                        else NbInscStage := 1;
                        Ferme(Q);
                        if NbMax > 0 then
                        begin
                          NbDivD := NbInscStage / NbMax;
                          NbDivE := arrondi(NbDivD, 0);
                          if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
                          else NbCout := NbDivE;
                        end
                        else NbCout := 1;
                        CoutFoncSta := (NbCout * CoutFonc) / NbInscStage; //PT5
                        CoutAnimSta := (NbCout * CoutAnim) / NbInscStage;
                        TobInsc.Detail[i].PutValue('PFI_AUTRECOUT',NbInsc * (CoutFoncSta + CoutAnimSta + Coutunit));
                end;
 //FIN PT2
                if GetCheckBoxState('CBPREVSALAIRE') = CbChecked then
                begin
                        If Salarie <> '' Then
                        begin
                                //PT7 - Début
                                // Lorsque le salarié est connu, on calcule selon la méthode paramétrée dans le ParamSoc
                                If (VH_Paie.PGForValoSalairePrev = 'VCR') Then
                                //PT8 - Début
                                Begin
                                     DateExercice := 0;
                                     Q := OpenSQL('SELECT PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_MILLESIME="'+LeMillesime+'"',True);
                                     If Not Q.Eof Then DateExercice := Q.FindField('PFE_DATEDEBUT').AsDateTime;
                                     Ferme(Q);
                                     Salaire := ForTauxHoraireReel(Salarie, iDate1900, iDate1900, '', ValPrevisionnel, DateExercice);
                                End
                                //PT8 - Fin
                                Else
                                     Salaire := ForTauxHoraireCategoriel(LibelleEmploi,LeMillesime);
                                //PT7 - Fin
                                TobInsc.Detail[i].PutValue('PFI_COUTREELSAL',Salaire*Duree*TauxChargeBudget);
                        end
                        Else
                        begin
                                Salaire := ForTauxHoraireCategoriel(LibelleEmploi,MillesimeMul);
                                NbHeure := TobInsc.Detail[i].GetValue('PFI_DUREESTAGE');
                                TobInsc.Detail[i].PutValue('PFI_COUTREELSAL',Salaire*NbHeure*TauxChargeBudget);
                        end;
                end;
                If (GetCheckBoxState('CBPREVFORFAIT') = CbChecked) And (CalcFrais) then //PT9
                begin
                      Req := 'SELECT PFF_FRAISHEBERG,PFF_FRAISREPAS,PFF_FRAISTRANSP FROM FORFAITFORM '+
                             'WHERE PFF_MILLESIME="'+MillesimeMul+'" AND PFF_LIEUFORM="'+Lieu+'"';

                      //PT9 - Début
                      If (VH_Paie.PGForGestFraisByPop) Then
                      Begin
                           If (Salarie <> '') Then
                                Req := Req + ' AND PFF_POPULATION IN (SELECT PNA_POPULATION FROM SALARIEPOPUL WHERE PNA_SALARIE="'+Salarie+'" AND PNA_TYPEPOP="'+TYP_POPUL_FORM_PREV+'")'
                           Else
                                // Recherche de la population correspondant aux données actuelles
                                Req := Req + ' AND PFF_POPULATION="'+RecherchePopulation(TobInsc.Detail[i])+'"';
                      End
                      Else
                           Req := Req + ' AND PFF_ETABLISSEMENT="'+Etab+'"';
                      //PT9 - Fin

                      Q := OpenSQL(Req,True);
                      FraisH := 0; FraisR := 0; FraisT := 0;
                      if not Q.eof then
                      begin
                        FraisH := Q.FindField('PFF_FRAISHEBERG').AsFloat;
                        FraisR := Q.FindField('PFF_FRAISREPAS').AsFloat;
                        FraisT := Q.FindField('PFF_FRAISTRANSP').AsFloat;
                      end;
                      Ferme(Q);
                      NbJour := TobInsc.Detail[i].GetValue('PFI_JOURSTAGE');
                      If Salarie = '' then NbJour := NbJour / NbInsc;
                      NbInsc := TobInsc.Detail[i].GetValue('PFI_NBINSC');
                      FraisH := FraisH*(NbJour-1);
                      If FraisH<0 then FraisH := 0;
                      If FraisH > 0 then FraisR := FraisR*((NbJour*2)-1) //PT6
                      else FraisR := FraisR*NbJour;
                      If FraisR<0 then FraisR := 0;
                      TotalFrais := FraisH+FraisR+FraisT;
                      If Salarie = '' then TobInsc.Detail[i].PutValue('PFI_FRAISFORFAIT',TotalFrais*NbInsc)  //PT2
                      else TobInsc.Detail[i].PutValue('PFI_FRAISFORFAIT',TotalFrais);
                end;
                TobInsc.Detail[i].UpdateDB(False);
        end;
end;

procedure TOF_PGCALCSALAIREFORM.MajTableFormations(LeStage , LeMillesime : String; LaSession : Integer);
var TobFormations,TobAnimat : Tob;
    Q : TQuery;
    i : Integer;
    TotalSal,FraisPlaf,FraisReel,TauxCadre,TauxNonCadre,SalaireAnim,FraisANim : Double;
    MillesimeEC,Salarie,LibelleEmploi,dadscat,AvecClient,Animateur : String;
    Nodossier : String;
begin
        MillesimeEC := '0000';
        TauxCadre := 1;
        TauxNonCadre := 1;
        Q := OpenSQL('SELECT PFE_MILLESIME,PFE_TAUXCHARGENC,PFE_TAUXCHARGEC FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-" ORDER BY PFE_MILLESIME DESC',true);
        If Not Q.eof then
        begin
                MillesimeEC := Q.FindField('PFE_MILLESIME').AsString;
                TauxCadre := Q.FindField('PFE_TAUXCHARGEC').AsFloat;
                TauxNonCadre := Q.FindField('PFE_TAUXCHARGENC').AsFloat;
        end;
        Ferme(Q);
        if (GetCheckBoxState('CALCSALAIRE') = CbChecked) or (GetCheckBoxState('CALCFFRAIS') = CbChecked) then
        begin
                //DEBUT PT1
                if GetCheckBoxState('CALCCTPEDAG') = CbChecked then
                begin
                        Q := OpenSQL('SELECT PSS_AVECCLIENT FROM SESSIONSTAGE WHERE '+
                        'PSS_ORDRE=' + IntToStr(LaSession) + ' AND '+  //DB2
                        'PSS_CODESTAGE="' + LeStage + '" AND '+
                        'PSS_MILLESIME="' + LeMillesime + '"',True);
                        If Not Q.Eof then AvecClient := Q.FindField('PSS_AVECCLIENT').AsString
                        else AvecClient := '-';
                        Ferme(Q);
                        Q := OpenSQL('SELECT * FROM SESSIONANIMAT WHERE '+
                        'PAN_ORDRE=' + IntToStr(LaSession) + ' AND '+  //DB2
                        'PAN_CODESTAGE="' + LeStage + '" AND '+
                        'PAN_MILLESIME="' + LeMillesime + '" AND '+
                        'PAN_SALARIE<>""',True);
                        TobAnimat := Tob.Create('SESSIONANIMAT',Nil,-1);
                        TobAnimat.LoadDetailDB('SESSIONANIMAT','','',Q,False);
                        Ferme(Q);
                        For i := 0 To TobAnimat.Detail.Count - 1 do
                        begin
                                If AvecClient = '-' then
                                begin
                                        //TotalSal := 0;
                                        Animateur := TobAnimat.Detail[i].GetValue('PAN_SALARIE');
                                        If PGBundleHierarchie then
                                        begin
                                          Q := OpenSQL('SELECT DOS_NOMBASE FROM INTERIMAIRES LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER WHERE PSI_INTERIMAIRE="'+TobAnimat.Detail[i].GetValue('PAN_SALARIE')+'"',True); //PT10
                                          Nodossier := Q.FindField('DOS_NOMBASE').AsString;
                                          Ferme(Q);
                                          Q := OpenSQL('SELECT PSA_LIBELLEEMPLOI,PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+TobAnimat.Detail[i].GetValue('PAN_SALARIE')+'"',True,-1,'',false,NoDossier); //PT10
                                        end
                                        else  Q := OpenSQL('SELECT PSA_LIBELLEEMPLOI,PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+Animateur+'"',True);
                                        If not Q.Eof then
                                        begin
                                                LibelleEmploi := Q.FindField('PSA_LIBELLEEMPLOI').AsString;
                                                dadscat := Q.FindField('PSA_DADSCAT').AsString;
                                        end
                                        else begin Libelleemploi := ''; dadscat := ''; end;
                                        Ferme(Q);
                                        TotalSal := MajSalaire(Animateur,LeStage, LeMillesime,LibelleEmploi,MillesimeEC,LaSession,NoDossier);
                                        TotalSal := TobAnimat.Detail[i].GetValue('PAN_NBREHEURE') * TotalSal;
                                        TotalSal := Arrondi(TotalSal,2);
                                        If (TauxCadre <> 1) or (TauxNonCadre <> 1) then
                                        begin
                                                If (dadscat = '01') or (dadscat = '02') then TotalSal := Arrondi(TotalSal * TauxCadre,2)
                                                else TotalSal := Arrondi(TotalSal * TauxNonCadre,2);
                                        end;
                                        TobAnimat.Detail[i].PutValue('PAN_SALAIREANIM',TotalSal);
                                        TobAnimat.Detail[i].UpdateDB(False);
                                end
                                else TobAnimat.Detail[i].PutValue('PAN_SALAIREANIM',0);
                                TobAnimat.Detail[i].UpdateDB(False);
                        end;
                        TobAnimat.Free;
                        SalaireAnim := 0;
                        FraisAnim := 0;
                        Q := OpenSQL('SELECT SUM (PAN_SALAIREANIM) SALAIREANIM FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+LeStage+'" AND '+
                        'PAN_ORDRE='+IntToStr(LaSession)+' AND PAN_MILLESIME="'+LeMillesime+'"',True);  //DB2
                        If Not Q.Eof then SalaireAnim := Q.FindField('SALAIREANIM').AsFloat;
                        Ferme(Q);
                        Q := OpenSQL('SELECT SUM (PFS_MONTANT) FRAISANIM FROM FRAISSALFORM LEFT JOIN SESSIONANIMAT ON PFS_SALARIE=PAN_SALARIE AND PFS_CODESTAGE=PAN_CODESTAGE AND PAN_ORDRE=PFS_ORDRE '+
                        'WHERE PAN_CODESTAGE="'+LeStage+'" AND PAN_FRASPEDAG="X" AND '+
                        'PAN_ORDRE='+IntToStr(LaSession)+' AND PAN_MILLESIME="'+LeMillesime+'"',True);  //DB2
                        If Not Q.Eof then FraisANim := Q.FindField('FRAISANIM').AsFloat;
                        Ferme(Q);
                        ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_COUTSALAIR='+StrfPoint(SalaireAnim + FraisANim)+' '+
                        'WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE='+IntToStr(LaSession)+' AND PSS_MILLESIME="'+LeMillesime+'"');  //DB2
                end;
                //FIN PT1
                Q := OpenSQL('SELECT * FROM FORMATIONS WHERE '+
                'PFO_ORDRE=' + IntToStr(LaSession) + ' AND '+  //DB2
                'PFO_CODESTAGE="' + LeStage + '" AND '+
                'PFO_MILLESIME="' + LeMillesime + '" AND '+
                'PFO_EFFECTUE="X"',True);
                TobFormations := Tob.Create('FORMATIONS',Nil,-1);
                TobFormations.LoadDetailDB('FORMATIONS','','',Q,False);
                Ferme(Q);
                For i := 0 To TobFormations.Detail.Count - 1 do
                begin
                        Salarie := TobFormations.Detail[i].GetValue('PFO_SALARIE');  //PT3  et PT4
                        LibelleEmploi := TobFormations.detail[i].GetValue('PFO_LIBEMPLOIFOR');
                        //TotalSal := 0;
                        If GetCheckBoxState('CALCSALAIRE') = CbChecked then
                        begin
                                //PT13 - Début
                                If PGBundleHierarchie then
                                begin
                                  Q := OpenSQL('SELECT DOS_NOMBASE FROM INTERIMAIRES LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER WHERE PSI_INTERIMAIRE="'+Salarie+'"',True);
                                  Nodossier := Q.FindField('DOS_NOMBASE').AsString;
                                  Ferme(Q);
                                end;
                                //PT13 - Fin
                                TotalSal := MajSalaire(Salarie,LeStage, LeMillesime,LibelleEmploi,MillesimeEC,LaSession,NoDossier);
                                TotalSal := TobFormations.Detail[i].GetValue('PFO_NBREHEURE') * TotalSal;
                                TotalSal := Arrondi(TotalSal,2);
                                If (TauxCadre <> 1) or (TauxNonCadre <> 1) then
                                begin
                                        If PGBundleHierarchie then
                                        begin
                                          //PT13
                                          //Q := OpenSQL('SELECT DOS_NOMBASE FROM INTERIMAIRES LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER WHERE PSI_INTERIMAIRE="'+TobFormations.Detail[i].GetValue('PFO_SALARIE')+'"',True);
                                          //Nodossier := Q.FindField('DOS_NOMBASE').AsString;
                                          //Ferme(Q);
                                          Q := OpenSQL('SELECT PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+TobFormations.Detail[i].GetValue('PFO_SALARIE')+'"',True,-1,'',false,NoDossier);
                                        end
                                        else Q := OpenSQL('SELECT PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+TobFormations.Detail[i].GetValue('PFO_SALARIE')+'"',True);
                                        If (Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02') then TotalSal := Arrondi(TotalSal * TauxCadre,2)
                                        else TotalSal := Arrondi(TotalSal * TauxNonCadre,2);
                                        Ferme(Q);
                                end;
                                TobFormations.Detail[i].PutValue('PFO_COUTREELSAL',TotalSal);
                        end;
                        If GetCheckBoxState('CALCFFRAIS') = CbChecked then
                        begin
                                FraisReel := MajFrais('REEL',Salarie,LeStage,LeMillesime,LaSession);
                                TobFormations.Detail[i].PutValue('PFO_FRAISREEL',FraisReel);
                                FraisPlaf := MajFrais('PLAF',Salarie,LeStage,LeMillesime,LaSession);
                                TobFormations.Detail[i].PutValue('PFO_FRAISPLAF',FraisPlaf);
                        end;
                        TobFormations.Detail[i].UpdateDB(False);
                end;
                TobFormations.Free;
        end;
end;

procedure TOF_PGCALCSALAIREFORM.AccesPeriodePaie(Sender : TObject);
begin
        If TCheckBox(Sender).Checked = True then
        begin
                SetControlEnabled('DATEDEB',True);
                SetControlEnabled('DATEFIN',True);
        end
        Else
        begin
                SetControlEnabled('DATEDEB',False);
                SetControlEnabled('DATEFIN',False);
        end;
end;

Function TOF_PGCALCSALAIREFORM.MajFrais(TypeFrais,LeSalarie,LeStage , LeMillesime : String; LaSession : Integer) : Double;
var Q : TQuery;
begin
        Result := 0;
        If TypeFrais = 'REEL' then
        begin
                Q := OpenSQL('SELECT SUM(PFS_MONTANT) TOTAL FROM FRAISSALFORM WHERE PFS_CODESTAGE="'+LeStage+'"' +
                ' AND PFS_MILLESIME="'+LeMillesime+'" AND PFS_ORDRE='+IntToStr(LaSession)+' AND PFS_SALARIE="'+LeSalarie+'"',True);  //DB2
                If Not Q.Eof then Result := Q.FindField('TOTAL').AsFloat;
                Ferme(Q);
        end
        // Plafonnés
        Else
        begin
                Q := OpenSQL('SELECT SUM(PFS_MONTANTPLAF) TOTAL FROM FRAISSALFORM WHERE PFS_CODESTAGE="'+LeStage+'"' +
                ' AND PFS_MILLESIME="'+LeMillesime+'" AND PFS_ORDRE='+IntToStr(LaSession)+' AND PFS_SALARIE="'+LeSalarie+'"',True);  //DB2
                If Not Q.Eof then Result := Q.FindField('TOTAL').AsFloat;
                Ferme(Q);
        end;
end;

Function TOF_PGCALCSALAIREFORM.MajSalaire(LeSalarie,LeStage , LeMillesime,LibEmploi,MillesimeEC : String; LaSession : Integer;NoDossier : String) : Double;
var TobPaie : Tob;
    Q : TQuery;
    NbPaie,j : Integer;
    TotalSal : Double;
    MValo,MCalc : String;
    dateDeb, DateFin : TDateTime;
    LeDossier : String;
//  StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare: string;
//  stLesBases, laSoc, SchemaName : string;
//  bSaveEnableDeShare, bSaveExisteDeShare : boolean;
begin
        MValo := GetControlText('METHODEVALO');
        MCalc := GetControlText('METHODECALC');
        dateDeb := StrToDate(GetControlText('DATEDEB'));
        DateFin := StrToDate(GetControlText('DATEFIN'));
        TotalSal := 0;
{        DB := TDataBase.Create(Appli);
        LaSoc := 'BUNDLE2';
        ChargeDBParams (laSoc, StDriver, StServer, StPath, Stbase, StUser, StPassWord,
                        StODBC, StOptions, StGroup, StShare);
        AssignDBParamsStrings ('BUNDLE2', StDriver, StServer, StPath, StBase, StUser, StPassWord,
                               StODBC, StOptions, DB);
        DB.Connected := TRUE;}
        if MValo = 'VCR' then
        begin
                If MCalc = 'THS' then TotalSal := ForTauxHoraireReel(LeSalarie,StrToDate(GetControlText('DATEDEB')),StrToDate(GetControlText('DATEFIN')),MCalc,False,0,NoDossier) //PT13
                Else
                begin
                        If GetControlText('TCHOIXPERIODE') = '-' then TotalSal := ForTauxHoraireReel(LeSalarie,IDate1900,IDate1900,MCalc,False,0,NoDossier) //PT13
                        Else
                        begin
                                NbPaie := 0;
                                LeDossier := '';
                                If PGBundleHierarchie then
                                begin
                                  Q := OpenSQL('SELECT DOS_NOMBASE FROM INTERIMAIRES LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER WHERE PSI_INTERIMAIRE="'+LeSalarie+'"',True);
                                  LeDossier := Q.FindField('DOS_NOMBASE').AsString;
                                  Ferme(Q);
                                  Q := OpenSQL('SELECT PPU_DATEDEBUT,PPU_DATEFIN FROM '+GetBase(LeDossier,'PAIEENCOURS') +
                                  ' WHERE PPU_SALARIE="'+LeSalarie+'" AND PPU_DATEDEBUT>="'+UsDateTime(DateDeb)+'" '+
                                  'AND PPU_DATEDEBUT<="'+UsDateTime(DateFin)+'"',True);//,-1,'',false,LeDossier);
                                end
                                else
                                begin
                                  Q := OpenSQL('SELECT PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS '+
                                  'WHERE PPU_SALARIE="'+LeSalarie+'" AND PPU_DATEDEBUT>="'+UsDateTime(DateDeb)+'" '+
                                  'AND PPU_DATEDEBUT<="'+UsDateTime(DateFin)+'"',True);
                                end;
                                TobPaie := Tob.Create('Les paies',Nil,-1);
                                TobPaie.LoadDetailDB('Les Paie','','',Q,False);
                                Ferme(Q);
                                For j := 0 To TobPaie.Detail.Count - 1 do
                                begin
                                        NbPaie := NbPaie +1;
                                        TotalSal := TotalSal + ForTauxHoraireReel(LeSalarie,TobPaie.Detail[j].GetValue('PPU_DATEDEBUT'),TobPaie.Detail[j].GetValue('PPU_DATEFIN'),MCalc,ValReel,IDate1900, LeDossier);
                                end;
                                If NbPaie > 0 then TotalSal := TotalSal / NbPaie;
                                TobPaie.Free;
                        end;
                end;
        end;
        if MValo = 'VCC' then
        begin
                TotalSal := ForTauxHoraireCategoriel(LibEmploi,MillesimeEC);
        end;
        Result := TotalSal
end;

Initialization
  registerclasses ( [ TOF_PGCALCSALAIREFORM ] ) ;
end.



