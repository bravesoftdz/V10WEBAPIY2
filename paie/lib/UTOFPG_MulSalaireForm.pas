{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 14/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULSALAIREFORM ()
Mots clefs ... : TOF;PGMULSALAIREFORM
*****************************************************************
PT1 08/04/2003 JL V_42 Correction date pour calcul
PT2 20/07/2005 JL V_60 FQ 12096 calcul avec ou sans charge patronales
PT3 20/07/2005 JL V_60 FQ 12095 vérif gestion budget
PT4 16/03/2006 JL V_650 FQ 12981 Gestion accès charges patronales
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
PT5 16/04/2007 FL V_720 FQ 13568 Proposer la méthode de calcul réelle par défaut si elle est
                                 basée sur une valorisation catégorielle, prévisionnelle sinon.
PT6 08/01/2008 FL V_802 FQ 13408 En taux horaire salarié, proposer aussi le taux de charge patronal
}
Unit UTOFPG_MulSalaireForm;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,
{$ELSE}
       eMul,
{$ENDIF}
     forms,sysutils,UTOB,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HTB97,PGOutilsFormation,ed_tools,EntPaie,HStatus,ParamDat,UtilPGI ;

Type
  TOF_PGMULSALAIREFORM = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure CalculerClick(Sender : TObject);
    procedure DateElipsisclick(Sender : TObject);
    procedure ChangementCalcul(Sender : TObject);
    procedure CalculerClickMultiDossier(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULSALAIREFORM.OnArgument (S : String ) ;
var Q : TQuery;
    Millesime,Combo : THValComboBox;
    BCalculer : TToolbarButton97;
    THDate : THEdit;
    DateDeb,DateFin : TDateTime;
begin
  Inherited ;
        Millesime := THValComboBox(GetControl('PSF_MILLESIME'));
        Datedeb := IDate1900;
        dateFin := IDate1900;
        Q := OpenSQL('SELECT PFE_MILLESIME,PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" ORDER BY PFE_MILLESIME DESC',true);
        if not Q.eof then
        begin
                Millesime.Value := Q.FindField('PFE_MILLESIME').AsString;
                DateDeb := Q.FindField('PFE_DATEDEBUT').AsDateTime;
                DateFin := Q.FindField('PFE_DATEFIN').AsDateTime;
        end;
        Ferme(Q);
        BCalculer := TToolBarButton97(GetControl('BCALCULER'));
        If BCalculer <> Nil then
        begin
          If PGBundleInscFormation then BCalculer.OnClick := CalculerClickMultiDossier
          else BCalculer.OnClick := CalculerClick;
        end;
        THDate := THEdit(GetControl('DATEDEB'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        THDate := THEdit(GetControl('DATEFIN'));
        SetControlText('DATEDEB',DateToStr(DateDeb));
        SetControlText('DATEFIN',DateToStr(DateFin));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        Combo := THValComboBox(GetControl('METHODECALC'));
        If Combo <> Nil then Combo.OnChange := ChangementCalcul;
        //PT5 - Début
        If (VH_Paie.PGForValoSalaire = 'VCC') Then
          SetControlText('METHODECALC',VH_Paie.PGForMethodeCalc)
        Else
          SetControlText('METHODECALC',VH_Paie.PGForMethodeCalcPrev);
        //PT5 - Fin
end ;

procedure TOF_PGMULSALAIREFORM.CalculerClick(Sender : TObject);
var Q : TQuery;
    TMaj,TobForfait,TobSalaire,TS,TobLibEmploi,T:Tob;
    Millesime:THValcomboBox;
    a,Duree,i:Integer;
    LibelleEmploi,Where:String;
    NbInsc,Salaire,HeuresTrav,HeuresEtab,TotSalaire:Double;
    TobInsc,TobStage:Tob;
    DateDeb,DateFin:TDateTime;
    TauxCadre, TauxNonCadre : Double; //PT6
begin
        Millesime := THValComboBox(GetControl('PSF_MILLESIME'));
        if ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE ((PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-") OR (PFE_ACTIF="X" AND PFE_CLOTURE="-")) AND PFE_MILLESIME="'+Millesime.Value+'"') then
        begin
                if PGIAsk('Attention, le calcul des salaires va effacer toutes les données éxistantes pour ce millésime.#13#10Voulez-vous continuer ?','Calcul des salaires moyens pour le millésime '+Millesime.Value)=mrYes then
                begin
                        ExecuteSQL('DELETE FROM SALAIREFORM WHERE PSF_MILLESIME="'+Millesime.Value+'"');
                        Q := OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PLE"',True);
                        TobLibEmploi := Tob.Create('Les libellés emploi',Nil,-1);
                        TobLibEmploi.LoadDetailDB('CHOIXCOD','','',Q,False);
                        Ferme(Q);
                        Q := OpenSQL('SELECT PSF_MILLESIME FROM SALAIREFORM WHERE PSF_MILLESIME="'+Millesime.Value+'"',true);
                                TobForfait := Tob.Create('SALAIREFORM',Nil,-1);
                                TobForfait.LoadDetailDB('SALAIREFORM','','',Q,False);
                                Ferme(Q);
                        If GetControlText('METHODECALC') <> 'THS' then
                        begin
                                DateDeb := StrToDate(GetControlText('DATEDEB'));      //DEBUT PT1
                                DateFin := StrToDate(GetControlText('DATEFIN'));
                                Q := OpenSQL('SELECT PPU_CHEURESTRAV,PPU_CBRUT,PPU_CCOUTPATRON,ETB_HORAIREETABL,PPU_LIBELLEEMPLOI FROM PAIEENCOURS LEFT JOIN ETABCOMPL'+
                                        ' ON PPU_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE PPU_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PPU_DATEFIN<="'+UsDateTime(DateFin)+'"',True);
                                TobSalaire := Tob.Create('Les paies',Nil,-1);
                                TobSalaire.LoadDetailDB('PAIEENCOURS','','',Q,False);
                                Ferme(Q);
                                InitMoveProgressForm (NIL,'Calcul des salaires moyens par libellé emploi', 'Veuillez patienter SVP ...',TobLibEmploi.Detail.Count,FALSE,TRUE);
                                InitMove(tobLibEmploi.Detail.Count,'');
                                For i := 0 to tobLibEmploi.Detail.Count-1 do
                                begin
                                        a := 0;
                                        LibelleEmploi := TobLibEmploi.Detail[i].getValue('CC_CODE');
                                        MoveCurProgressForm(RechDom('PGLIBEMPLOI',LibelleEmploi,False));
                                        TotSalaire := 0;
                                        TS := TobSalaire.FindFirst(['PPU_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        While TS <> Nil do
                                        begin
                                                a := a+1;
                                                HeuresTrav := TS.GetValue('PPU_CHEURESTRAV');
                                                If GetCheckBoxState('CHARGES') = CbChecked then Salaire := TS.GetValue('PPU_CBRUT')+TS.GetValue('PPU_CCOUTPATRON')//PT2
                                                else Salaire := TS.GetValue('PPU_CBRUT');
                                                HeuresEtab := TS.GetValue('ETB_HORAIREETABL');
                                                If heurestrav <> 0 then Salaire := (Salaire * HeuresEtab)/HeuresTrav;
                                                If HeuresEtab <> 0 then TotSalaire := TotSalaire + (Salaire/HeuresEtab);
                                                TS := TobSalaire.FindNext(['PPU_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        end;
                                        if a <> 0 then TotSalaire := TotSalaire/a;
                                        TMaj := Tob.Create('SALAIREFORM',Nil,-1);
                                        TMaj.PutValue('PSF_MILLESIME',Millesime.Value);
                                        TMaj.PutValue('PSF_LIBEMPLOIFOR',LibelleEmploi);
                                        TMaj.PutValue('PSF_MONTANT',Arrondi(TotSalaire,2));
                                        TMaj.InsertOrUpdateDB(False);
                                        TMaj.Free;
                                end;
                                FiniMoveProgressForm;
                                TobSalaire.Free;
                        end
                        else
                        begin
                                Q := OpenSQL('SELECT PSA_LIBELLEEMPLOI,PSA_TAUXHORAIRE,PSA_DADSCAT FROM SALARIES WHERE PSA_DATESORTIE="'+UsDateTime(IDATE1900)+'"',True); //PT6
                                TobSalaire := Tob.Create('Les paies',Nil,-1);
                                TobSalaire.LoadDetailDB('PAIEENCOURS','','',Q,False);
                                Ferme(Q);
                                
                                //PT6 - Début
                                TauxCadre    := 1;
                                TauxNonCadre := 1;
                                Q := OpenSQL('SELECT PFE_TAUXCHARGEC,PFE_TAUXCHARGENC FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime.Value+'"',True);
                                If Not Q.EOF Then 
                                Begin
                                	TauxCadre    := Q.FindField('PFE_TAUXCHARGEC').AsFloat;
                                	TauxNonCadre := Q.FindField('PFE_TAUXCHARGENC').AsFloat;
                                End;
                                //PT6 - Fin
                                
                                InitMoveProgressForm (NIL,'Calcul des salaires moyens par libellé emploi', 'Veuillez patienter SVP ...',TobLibEmploi.Detail.Count,FALSE,TRUE);
                                InitMove(tobLibEmploi.Detail.Count,'');
                                For i := 0 to tobLibEmploi.Detail.Count-1 do
                                begin
                                        a := 0;
                                        LibelleEmploi := TobLibEmploi.Detail[i].getValue('CC_CODE');
                                        MoveCurProgressForm(RechDom('PGLIBEMPLOI',LibelleEmploi,False));
                                        TotSalaire := 0;
                                        TS := TobSalaire.FindFirst(['PSA_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        While TS <> Nil do
                                        begin
                                                // Si le taux horaire est à 0, c'est qu'il n'a pas été renseigné
                                                // Il n'est donc pas nécessaire de le prendre en compte
                                                If TS.GetValue('PSA_TAUXHORAIRE') > 0 Then //PT6
                                                Begin
                                                    a := a+1;
                                                    TotSalaire := TotSalaire + TS.GetValue('PSA_TAUXHORAIRE');
                                                    //PT6 - Début
                                                    If GetCheckBoxState('CHARGES') = CbChecked then
                                                    Begin
                                                		If (TS.GetValue('PSA_DADSCAT') = '01') Or (TS.GetValue('PSA_DADSCAT') = '02') Then 
                                                			TotSalaire := TotSalaire * TauxCadre
                                                		Else
                                                        	TotSalaire := TotSalaire * TauxNonCadre;
                                                    End;
                                                    //PT6 - Fin
                                                End;
                                                TS := TobSalaire.FindNext(['PSA_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        end;
                                        if a <> 0 then TotSalaire := TotSalaire/a;
                                        TMaj := Tob.Create('SALAIREFORM',Nil,-1);
                                        TMaj.PutValue('PSF_MILLESIME',Millesime.Value);
                                        TMaj.PutValue('PSF_LIBEMPLOIFOR',LibelleEmploi);
                                        TMaj.PutValue('PSF_MONTANT',Arrondi(TotSalaire,2));
                                        TMaj.InsertOrUpdateDB(False);
                                        TMaj.Free;
                                end;
                                FiniMoveProgressForm;
                                TobSalaire.Free;
                        end;
                        TobLibEmploi.Free;
                        TobForfait.Free;
                        if ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" AND PFE_MILLESIME="'+Millesime.Value+'"') then
                        begin
                            if VH_Paie.PGForPrevisionnel = True then // PT3
                            begin
                                if PGIAsk('Voulez-vous mettre à jour le budget ?','Calcul des salaires moyens pour le millésime '+Millesime.Value)=mrYes then
                                begin
                                        If VH_Paie.PGForValoSalaire = 'VCC' then Where := 'WHERE PFI_MILLESIME="'+Millesime.Value+'"'
                                        Else Where := 'WHERE PFI_SALARIE="" AND PFI_MILLESIME="'+Millesime.Value+'"';
                                        Q := OpenSQL('SELECT * FROM INSCFORMATION '+Where,True);
                                        TobInsc := Tob.Create('INSCFORMATION',Nil,-1);
                                        TobInsc.LoadDetailDB('INSCFORMATION','','',Q,False);
                                        Ferme(Q);
                                        Q := OpenSQL('SELECT PST_CODESTAGE,PST_DUREESTAGE FROM STAGE WHERE PST_MILLESIME="'+Millesime.Value+'"',True);
                                        TobStage := Tob.Create('Les stages',Nil,-1);
                                        TobStage.LoadDetailDB('STAGE','','',Q,False);
                                        Ferme(Q);
                                        T := TobInsc.FindFirst([''],[''],False);
                                        While T <> Nil do
                                        begin
                                                Duree := 0;
                                                Salaire := RecupSalaireFormationNonNom(T.GetValue('PFI_MILLESIME'),T.GetValue('PFI_LIBEMPLOIFOR'));
                                                NbInsc := T.GetValue('PFI_NBINSC');
                                                TS := TobStage.FindFirst(['PST_CODESTAGE'],[T.GetValue('PFI_CODESTAGE')],False);
                                                If TS <> Nil Then Duree := TS.GetValue('PST_DUREESTAGE');
                                                T.PutValue('PFI_COUTREELSAL',arrondi(Salaire*NbInsc*Duree,2));
                                                T.UpdateDB;
                                                T := TobInsc.FindNext([''],[''],False);
                                        end;
                                        TobInsc.Free;
                                        TobStage.Free;
                                end;
                            end;
                        end;
                end;
                TFMul(Ecran).BChercheClick(Nil);
        end
        Else PGIBox('Calcul impossible car ni l''exercice ni le budget ne sont ouverts pour l''année '+Millesime.Value,Ecran.Caption);
end;

procedure TOF_PGMULSALAIREFORM.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULSALAIREFORM.ChangementCalcul(Sender : TObject);
begin
        If THValComboBox(Sender).Value = 'THS' then
        begin
                SetControlVisible('DATEDEB',False);
                SetControlVisible('TDATEDEB',False);
                SetControlVisible('DATEFIN',False);
                SetControlVisible('TDATEFIN',False);
                SetControlVisible('CHARGES',True);  //PT4 //PT6
        end
        Else
        begin
                SetControlVisible('DATEDEB',True);
                SetControlVisible('TDATEDEB',True);
                SetControlVisible('DATEFIN',True);
                SetControlVisible('TDATEFIN',True);
                SetControlVisible('CHARGES',True); //PT4
        end;
end;

procedure TOF_PGMULSALAIREFORM.CalculerClickMultiDossier(Sender : TObject);
var Q : TQuery;
    TMaj,TobForfait,TobSalaire,TS,TobLibEmploi,T:Tob;
    Millesime:THValcomboBox;
    a,Duree,i:Integer;
    LibelleEmploi,Where:String;
    NbInsc,Salaire,HeuresTrav,HeuresEtab,TotSalaire:Double;
    TobInsc,TobStage:Tob;
    DateDeb,DateFin:TDateTime;
    NoDossier : String;
    TobNoDossier : Tob;
    NbDos : Integer;
begin
        Millesime := THValComboBox(GetControl('PSF_MILLESIME'));
        if ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE ((PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-") OR (PFE_ACTIF="X" AND PFE_CLOTURE="-")) AND PFE_MILLESIME="'+Millesime.Value+'"') then
        begin
                if PGIAsk('Attention, le calcul des salaires va effacer toutes les données éxistantes pour ce millésime.#13#10Voulez-vous continuer ?','Calcul des salaires moyens pour le millésime '+Millesime.Value)=mrYes then
                begin
                        DateDeb := StrToDate(GetControlText('DATEDEB'));
                        DateFin := StrToDate(GetControlText('DATEFIN'));

                        ExecuteSQL('DELETE FROM SALAIREFORM WHERE PSF_MILLESIME="'+Millesime.Value+'"');
                        Q := OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PLE"',True);
                        TobLibEmploi := Tob.Create('Les libellés emploi',Nil,-1);
                        TobLibEmploi.LoadDetailDB('CHOIXCOD','','',Q,False);
                        Ferme(Q);
                        Q := OpenSQL('SELECT PSF_MILLESIME FROM SALAIREFORM WHERE PSF_MILLESIME="'+Millesime.Value+'"',true);
                                TobForfait := Tob.Create('SALAIREFORM',Nil,-1);
                                TobForfait.LoadDetailDB('SALAIREFORM','','',Q,False);
                                Ferme(Q);
                        If GetControlText('METHODECALC') <> 'THS' then
                        begin
                                //DateDeb := StrToDate(GetControlText('DATEDEB'));      //DEBUT PT1 
                                //DateFin := StrToDate(GetControlText('DATEFIN'));
                                InitMoveProgressForm (NIL,'Calcul des salaires moyens par libellé emploi', 'Veuillez patienter SVP ...',TobLibEmploi.Detail.Count,FALSE,TRUE);
                                InitMove(tobLibEmploi.Detail.Count,'');
                                For i := 0 to tobLibEmploi.Detail.Count-1 do
                                begin
                                  Q := OpenSQL('SELECT DISTINCT(DOS_NOMBASE) FROM INTERIMAIRES LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER',True);
                                  TobNoDossier := Tob.Create('Les dossiers',Nil,-1);
                                  TobNoDossier.LoadDetailDB('les dossiers','','',Q,False);
                                  Ferme(Q);
                                  a := 0;
                                  TotSalaire := 0;
                                  LibelleEmploi := TobLibEmploi.Detail[i].getValue('CC_CODE');
                                  MoveCurProgressForm(RechDom('PGLIBEMPLOI',LibelleEmploi,False));
                                  For NbDos := 0 to TobNoDossier.Detail.Count-1  do
                                  begin
                                        NoDossier := TobNoDossier.Detail[NbDos].GetValue('DOS_NOMBASE');
                                        Q := OpenSQL('SELECT PPU_CHEURESTRAV,PPU_CBRUT,PPU_CCOUTPATRON,ETB_HORAIREETABL,PPU_LIBELLEEMPLOI FROM '+GetBase(NoDossier,'PAIEENCOURS') +' LEFT JOIN '+GetBase(NoDossier,'ETABCOMPL') +''+
                                        ' ON PPU_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE PPU_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PPU_DATEFIN<="'+UsDateTime(DateFin)+'"',True);
                                        TobSalaire := Tob.Create('Les paies',Nil,-1);
                                        TobSalaire.LoadDetailDB('les paies','','',Q,False);
                                        Ferme(Q);
                                        TS := TobSalaire.FindFirst(['PPU_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        While TS <> Nil do
                                        begin
                                                a := a+1;
                                                HeuresTrav := TS.GetValue('PPU_CHEURESTRAV');
                                                If GetCheckBoxState('CHARGES') = CbChecked then Salaire := TS.GetValue('PPU_CBRUT')+TS.GetValue('PPU_CCOUTPATRON')//PT2
                                                else Salaire := TS.GetValue('PPU_CBRUT');
                                                HeuresEtab := TS.GetValue('ETB_HORAIREETABL');
                                                If heurestrav <> 0 then Salaire := (Salaire * HeuresEtab)/HeuresTrav;
                                                If HeuresEtab <> 0 then TotSalaire := TotSalaire + (Salaire/HeuresEtab);
                                                TS := TobSalaire.FindNext(['PPU_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        end;
                                        if a <> 0 then TotSalaire := TotSalaire/a;
                                  end;
                                  TobNoDossier.Free;
                                  TMaj := Tob.Create('SALAIREFORM',Nil,-1);
                                  TMaj.PutValue('PSF_MILLESIME',Millesime.Value);
                                  TMaj.PutValue('PSF_LIBEMPLOIFOR',LibelleEmploi);
                                  TMaj.PutValue('PSF_MONTANT',Arrondi(TotSalaire,2));
                                  TMaj.InsertOrUpdateDB(False);
                                  TMaj.Free;
                                end;
                                FiniMoveProgressForm;
                                If TobSalaire <> Nil Then FreeAndNil(TobSalaire);
                        end
                        else
                        begin
                                InitMoveProgressForm (NIL,'Calcul des salaires moyens par libellé emploi', 'Veuillez patienter SVP ...',TobLibEmploi.Detail.Count,FALSE,TRUE);
                                InitMove(tobLibEmploi.Detail.Count,'');
                                For i := 0 to tobLibEmploi.Detail.Count-1 do
                                begin
                                  Q := OpenSQL('SELECT DISTINCT(PSI_NODOSSIER) FROM INETRIMAIRES',True);
                                  TobNoDossier := Tob.Create('Les dossiers',Nil,-1);
                                  TobNoDossier.LoadDetailDB('les dossiers','','',Q,False);
                                  Ferme(Q);
                                  a := 0;
                                  TotSalaire := 0;
                                  LibelleEmploi := TobLibEmploi.Detail[i].getValue('CC_CODE');
                                  MoveCurProgressForm(RechDom('PGLIBEMPLOI',LibelleEmploi,False));
                                  For NbDos := 0 to TobNoDossier.Detail.Count-1  do
                                  begin
                                        NoDossier := TobNoDossier.Detail[NbDos].GetValue('PSI_NODOSSIER');
                                        Q := OpenSQL('SELECT PPU_CHEURESTRAV,PPU_CBRUT,PPU_CCOUTPATRON,ETB_HORAIREETABL,PPU_LIBELLEEMPLOI FROM '+GetBase(NoDossier,'PAIEENCOURS') +' LEFT JOIN '+GetBase(NoDossier,'ETABCOMPL') +
                                        ' ON PPU_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE PPU_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PPU_DATEFIN<="'+UsDateTime(DateFin)+'"',True,-1,'',false,NoDossier);
                                        TobSalaire := Tob.Create('Les paies',Nil,-1);
                                        TobSalaire.LoadDetailDB('les paies','','',Q,False);
                                        Ferme(Q);
                                        TS := TobSalaire.FindFirst(['PSA_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        While TS <> Nil do
                                        begin
                                                a := a+1;
                                                TotSalaire := TotSalaire + TS.GetValue('PSA_TAUXHORAIRE');
                                                TS := TobSalaire.FindNext(['PSA_LIBELLEEMPLOI'],[LibelleEmploi],False);
                                        end;
                                        if a <> 0 then TotSalaire := TotSalaire/a;
                                  end;
                                  TobNoDossier.Free;
                                  TMaj := Tob.Create('SALAIREFORM',Nil,-1);
                                  TMaj.PutValue('PSF_MILLESIME',Millesime.Value);
                                  TMaj.PutValue('PSF_LIBEMPLOIFOR',LibelleEmploi);
                                  TMaj.PutValue('PSF_MONTANT',Arrondi(TotSalaire,2));
                                  TMaj.InsertOrUpdateDB(False);
                                  TMaj.Free;
                                end;
                                FiniMoveProgressForm;
                                TobSalaire.Free;
                        end;
                        TobLibEmploi.Free;
                        TobForfait.Free;
                        if ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" AND PFE_MILLESIME="'+Millesime.Value+'"') then
                        begin
                            if VH_Paie.PGForPrevisionnel = True then // PT3
                            begin
                                if PGIAsk('Voulez-vous mettre à jour le budget ?','Calcul des salaires moyens pour le millésime '+Millesime.Value)=mrYes then
                                begin
                                        If VH_Paie.PGForValoSalaire = 'VCC' then Where := 'WHERE PFI_MILLESIME="'+Millesime.Value+'"'
                                        Else Where := 'WHERE PFI_SALARIE="" AND PFI_MILLESIME="'+Millesime.Value+'"';
                                        Q := OpenSQL('SELECT * FROM INSCFORMATION '+Where,True);
                                        TobInsc := Tob.Create('INSCFORMATION',Nil,-1);
                                        TobInsc.LoadDetailDB('INSCFORMATION','','',Q,False);
                                        Ferme(Q);
                                        Q := OpenSQL('SELECT PST_CODESTAGE,PST_DUREESTAGE FROM STAGE WHERE PST_MILLESIME="'+Millesime.Value+'"',True);
                                        TobStage := Tob.Create('Les stages',Nil,-1);
                                        TobStage.LoadDetailDB('STAGE','','',Q,False);
                                        Ferme(Q);
                                        T := TobInsc.FindFirst([''],[''],False);
                                        While T <> Nil do
                                        begin
                                                Duree := 0;
                                                Salaire := RecupSalaireFormationNonNom(T.GetValue('PFI_MILLESIME'),T.GetValue('PFI_LIBEMPLOIFOR'));
                                                NbInsc := T.GetValue('PFI_NBINSC');
                                                TS := TobStage.FindFirst(['PST_CODESTAGE'],[T.GetValue('PFI_CODESTAGE')],False);
                                                If TS <> Nil Then Duree := TS.GetValue('PST_DUREESTAGE');
                                                T.PutValue('PFI_COUTREELSAL',arrondi(Salaire*NbInsc*Duree,2));
                                                T.UpdateDB;
                                                T := TobInsc.FindNext([''],[''],False);
                                        end;
                                        TobInsc.Free;
                                        TobStage.Free;
                                end;
                            end;
                        end;
                end;
                TFMul(Ecran).BChercheClick(Nil);
        end
        Else PGIBox('Calcul impossible car ni l''exercice ni le budget ne sont ouverts pour l''année '+Millesime.Value,Ecran.Caption);
end;


Initialization
  registerclasses ( [ TOF_PGMULSALAIREFORM ] ) ;
end.
