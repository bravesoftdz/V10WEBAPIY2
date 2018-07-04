{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : TOM Gestion des stages
Mots clefs ... : PAIE
*****************************************************************
PT1  | 15/05/2003 | V_42  | JL | Gestion investissement
PT2  | 26/11/2003 | V_50  | JL | Ajout type investissement = plan de formation par défaut.
PT3  | 06/10/2005 | V_60  | JL | FQ 12087 Griser descriptif
PT4  | 30/03/2006 | V_65  | JL | FQ 13006 Modif duplication
PT5  | 30/08/2006 | V_70  | JL | FQ 13468 Combo au lieu de THEdit pour centre de formation
PT6  | 30/08/2006 | V_70  | JL | 13460 Gestion chgt nb animateurs et changements heure (modif du coût)
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT7  | 03/05/2007 | V_720 | FL | FQ 14051 Gestion du bouton "Libellés Emplois"
PT8  | 12/06/2007 | V_720 | FL | FQ 13334 Permettre la modification du salaire animateur et correction du calcul du coût
PT9  | 14/06/2007 | V_720 | JL | Gestion partage formation
PT10 | 14/11/2007 | V_80  | FL | CWAS : Chargement de la grille d'investissement sur affichage de l'onglet coût
     |            |       |    |        afin d'éviter le message "Impossible de focaliser une fenêtre inactive".
PT11 | 24/10/2007 | V_8   | FL | Emanager / Report / Ajout gestion favoris
PT12 | 02/04/2008 | V_8   | FL | Restriction sur les prédéfinis
PT13 | 02/04/2008 | V_8   | FL | Uniformisation des types de plans avec le réalisé
PT14 | 02/04/2008 | V_8   | FL | Adaptation aux groupes de travail
PT15 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue
PT16 | 17/04/2008 | V_804 | FL | Ne pas autoriser la modification des emplois associés si consultation
}
unit UTOMSTAGE;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Spin,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fe_Main,Fiche,
{$ELSE}
       MaineAgl,eFiche,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOM,UTOB,HTB97,PgOutils,PgOutils2,EntPaie,PgoutilsFormation ;

Type
     TOM_STAGE = Class(TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnAfterUpdateRecord; override;
       procedure OnNewRecord; override;
       procedure OnChangeField( F: TField); override;
       procedure OnDeleteRecord; override;
       procedure OnUpdateRecord ; override;
       private
       TypeSaisie,MillesimeConsult,StageDupliquer{$IFDEF EMANAGER},MillesimeEC{$ENDIF} : String;  //PT11
       GInvest : THGrid;
       SalAnimUser : Boolean; //PT8
       SalaireAnimateur,TauxBudget : Double; //PT8
       GrilleChargee : Boolean;  //PT10
       procedure BObjectifsClick(Sender : TObject);
       procedure BDupliquerClick(Sender : TObject);
       procedure BFormationClick(Sender : TObject);
       Procedure BEmploisClick (Sender : TOBject); //PT7
       procedure RecupInfosStage;
       procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
       procedure MajInvest;
       procedure GriserDescriptif(Acces : Boolean); //PT3
       {procedure MajPrev;}
       Function  CalcSalaireAnimateur : Double; //PT8
       {$IFDEF EMANAGER}
       procedure AjouterAuxFavoris(Sender : Tobject);  //PT11
       procedure RetirerDesFavoris(Sender : Tobject);  //PT11
       {$ENDIF}
       procedure onEnterOngletCouts (Sender : TObject); //PT10
     END ;

implementation

Uses GalOutil, TntWideStrings;

{ TOM_STAGE }

procedure TOM_STAGE.OnUpdateRecord;
Var
    Rep : Integer;
begin
  inherited;
        If GetField('PST_PREDEFINI') = 'STD' then SetField('PST_NODOSSIER','000000');//PT9
        If (DS.State in [dsInsert]) Then
        begin
                If GetField('PST_CODESTAGE') = '--CURSUS--' then
                begin
                        PGIBox('Ce code est résevé, veuillez en saisir un autre',Ecran.Caption);
                        LastError := 1;
                        SetFocusControl('PST_CODESTAGE');
                        Exit;
                end;
                TFFiche(Ecran).Retour := GetField('PST_CODESTAGE');
        end;

        //PT8 - Début
        If SalAnimUser Then
        Begin
             Rep := PGIAsk('Attention, la valeur saisie pour les salaires des animateurs est différente de celle calculée. Voulez-vous garder la valeur saisie?');
             If Rep = MrNo then
             Begin
                  SetField('PST_COUTSALAIR', CalcSalaireAnimateur);
                  SalAnimUser := False;
             End;
        End
        Else
             SetField('PST_COUTSALAIR', CalcSalaireAnimateur);
        //PT8 - Fin
end;

procedure TOM_STAGE.OnDeleteRecord;
begin
  inherited;
        If ExisteSQL('SELECT PSS_CODESTAGE FROM SESSIONSTAGE WHERE PSS_MILLESIME="'+GetField('PST_MILLESIME')+'" AND PSS_CODESTAGE="'+GetField('PST_CODESTAGE')+'"') Then
        begin
                LastError := 1;
                PGIBox('Vous ne pouvez pas supprimer cette formation car des sessions de formations lui sont affectées',TFFiche(Ecran).Caption);
                Exit;
        end;
        If GetField('PST_MILLESIME') <> '0000' Then
        begin
                If ExisteSQL('SELECT PFI_ETABLISSEMENT FROM INSCFORMATION WHERE PFI_MILLESIME="'+GetField('PST_MILLESIME')+'" AND PFI_CODESTAGE="'+GetField('PST_CODESTAGE')+'"') Then
                begin
                        LastError := 1;
                        PGIBox('Vous ne pouvez pas supprimer cette formation car des inscriptions au budget lui sont affectées',TFFiche(Ecran).Caption);
                        Exit;
                end;
        end;
end;

procedure TOM_STAGE.OnArgument(stArgument : String);
var BObjectifs,BDupliquer,BFormation,BEmplois{$IFDEF EMANAGER},Bt{$ENDIF} : TToolBarButton97;
    Num : Integer;
    Stage : String;
begin
 inherited;
        {$IFDEF EMANAGER}  //PT11
        MillesimeEC := RendMillesimeEManager;
        {$ENDIF}
        AvertirTable('PGORGCOLLECTEUR');
        AvertirTable('PGCENTREFORMATION');
        AvertirTable('PGLIEUFORMATION');
        TypeSaisie := ReadTokenPipe(StArgument,';');
        Stage := ReadTokenPipe(StArgument,';');
        MillesimeConsult := ReadTokenPipe(StArgument,';');
        BObjectifs:=TToolBarButton97(GetControl('BOBJECTIFS'));
        SetcontrolVisible('BOBJECTIFS',False);                 // SUPPRESSION OBJECTIFS);
        If BObjectifs <> Nil Then BObjectifs.OnClick := BObjectifsClick;
        BDupliquer := TToolBarButton97(GetControl('BDUPLIQUER'));
        BEmplois := TToolBarButton97(GetControl('BEMPLOIS')); //PT7
        If BEmplois <> Nil Then BEmplois.OnClick := BEmploisClick;

        If (TypeSaisie = 'SAISIECAT') and (BDupliquer <> Nil) Then
        begin
                BDupliquer.Visible := True;
                BDupliquer.OnClick := BDupliquerClick;
                SetControlVisible('BFORMATION',False);
        end;
        If TypeSaisie = 'CONSULTCAT' then
        begin
                SetControlVisible('BDUPLIQUER',False);
                SetControlVisible('BFORMATION',False);
                SetControlVisible('BOBJECTIFS',False);
                //PT11 - Début
                {$IFDEF EMANAGER}
                Bt := TToolBarButton97(GetControl('BFAVORIS'));
                If Bt <> Nil then bt .OnClick := AjouterAuxFavoris;
                Bt := TToolBarButton97(GetControl('BNONFAVORIS'));
                If Bt <> Nil then bt .OnClick := RetirerDesFavoris;
                {$ENDIF}
                //PT11 - Fin
        end;
        If TypeSaisie = 'SAISIEBUD' Then
        begin
                If Not (ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+MillesimeConsult+'" AND PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-"')) Then
                TFFiche(Ecran).FTypeAction := TaConsult
                Else TFFiche(Ecran).FTypeAction := TaModif;
                BFormation := TToolBarButton97(GetControl('BFORMATION'));
                If BFormation<>Nil Then BFormation.onclick := BFormationClick;
                BFormation.Visible := False;
                SetControlEnabled('PDESCRIPTION',False);
                GriserDescriptif(False);
                //SetControlEnabled('PCOMP',False);
                SetControlProperty('BFORMATION','Hint','Inscriptions au budget');
        end;
        If TypeSaisie = 'CONSULTCAT' Then
        begin
                TFFiche(Ecran).FTypeAction := taConsult;
                SetControlVisible('BFORMATION',False);
        end;
        If (TypeSaisie = 'CONSULTCAT') or (TypeSaisie = 'SAISIECAT') Then
        begin
                SetControlVIsible('PST_COUTSALAIR',False);
                SetControlVIsible('TPST_COUTSALAIR',False);
        end;
        If TypeSaisie <> 'SAISIECAT' then SetControlVisible('BINSERT',False);
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        StageDupliquer := '';
        if (VH_PAIE.PGTypeCodeStage='ALP') then
        SetControlProperty('PST_CODESTAGE','EditMask','');
        if (VH_PAIE.PGTypeCodeStage='NUM') then
        SetControlProperty('PST_CODESTAGE','EditMask','999999');
        TFFiche(Ecran).MonoFiche := True;
        //DEBUT PT1
        GInvest := THGrid(GetControl('GINVESTSESSION'));
        //PT10 - Début
        TTabSheet(GetControl('TBSHTCOUT',true)).OnShow := onEnterOngletCouts;
        GrilleChargee := False;
        //If GInvest <> nil then
        //begin
        //        GInvest.ColFormats[0] := 'CB=PGFINANCEFORMATION';
        //        GInvest.ColFormats[1] := 'CB=PGOUINON';
        //        GInvest.ColFormats[2] := '# ##0.00';
        //        GInvest.ColFormats[3] := 'CB=PGOUINON';
        //        GInvest.ColFormats[4] := '# ##0.00';
        //        GInvest.ColFormats[5] := 'CB=PGOUINON';
        //        GInvest.ColFormats[6] := '# ##0.00';
        //        GInvest.ColFormats[7] := '# ##0';
        //        GInvest.OnCellExit := OnCellExit;
        //end;
        //PT10 - Fin
        //FIN PT1
        SalaireAnimateur := -1; //PT8
        SalAnimUser := False; //PT8
        //DEBUT PT9
        If PGBundleCatalogue then //PT15
        begin
          If (not PGDroitMultiForm) Or ((StArgument='ACTION=MODIFICATION') or (StArgument='ACTION=CONSULTATION')) then
          begin
            SetControlEnabled('PST_PREDEFINI',False);
            SetControlEnabled('PST_NODOSSIER',False);
          end
       	  Else If V_PGI.ModePCL='1' Then 
       		SetControlProperty('PST_NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT14
        end
        else
        begin
          SetControlVisible('PST_PREDEFINI',False);
          SetControlVisible('PST_NODOSSIER',False);
          SetControlVisible('TPST_PREDEFINI',False);
          SetControlVisible('TPST_NODOSSIER',False);
        end;
        //FIN PT9

        GereAccesPredefini (THValComboBox(GetControl('PST_PREDEFINI'))); //PT12

        If Not JailedroitTag(47160) Then SetControlVisible('BEMPLOIS', False); //PT16
end;

procedure TOM_STAGE.OnChangeField( F : TField);
var TotalCout : double;
    CodeStage : String;
    {SalaireAnim,NbHeure,NbAnim : Double;
    QAnim : TQuery;
    Millesime : String;}
begin
  inherited;
        If F.FieldName='PST_CODESTAGE' Then
        begin
                if (GetField('PST_CODESTAGE')) <> '' then
                begin
                        CodeStage := Trim(GetField('PST_CODESTAGE'));
                        if (VH_PAIE.PGTypeCodeStage='NUM') and (isnumeric(CodeStage)) and (CodeStage  <>  '') then
                        begin
                                CodeStage  :=  ColleZeroDevant(StrToInt(trim(CodeStage)),6);
                                if CodeStage  <>  (GetField('PST_CODESTAGE')) then SetField('PST_CODESTAGE',CodeStage);
                        end
                        else if (VH_PAIE.PGTypeCodeStage='ALP') AND (CodeStage <> GetField('PST_CODESTAGE')) then
                        SetField('PST_CODESTAGE',CodeStage);
                end;
        end;
        If F.FieldName='PST_NATUREFORM' Then
        begin
                If GetField('PST_NATUREFORM')='001' Then
                begin
                        SetControlVisible('TPST_CENTREFORMGU',False);
                        SetField('PST_CENTREFORMGU','');
                        SetControlVisible('PFO_NBANIM',True);
                        SetControlVIsible('TPFO_NBANIM',True);
                        SetControlEnabled('PFO_COUTSALAIR',False);
                end
                Else
                begin
                        SetControlVisible('PFO_NBANIM',False);
                        SetControlVIsible('TPFO_NBANIM',False);
                        SetControlEnabled('PFO_COUTSALAIR',True);
                        SetControlVisible('TPST_CENTREFORMGU',True);
                        SetControlVisible('PST_CENTREFORMGU',True);
                end;
        end;
        If TypeSaisie = 'SAISIECAT' Then
        begin
                If F.FieldName='PST_COUTMOB' Then
                begin
                        TotalCout := GetField('PST_COUTMOB')+GetField('PST_COUTORGAN')+GetField('PST_COUTFONCT');
                        If TotalCout <> 0 Then
                        begin
                                SetField('PST_COUTBUDGETE',TotalCout);
                                SetControlEnabled('PST_COUTBUDGETE',False);
                        end
                        else SetControlEnabled('PST_COUTBUDGETE',True);
                end;
                If F.FieldName='PST_COUTORGAN' Then
                begin
                        TotalCout := GetField('PST_COUTMOB')+GetField('PST_COUTORGAN')+GetField('PST_COUTFONCT');
                        If TotalCout <> 0 Then
                        begin
                                SetField('PST_COUTBUDGETE',TotalCout);
                                SetControlEnabled('PST_COUTBUDGETE',False);
                        end
                        Else SetControlEnabled('PST_COUTBUDGETE',True);
                end;
                If F.FieldName='PST_COUTFONCT' Then
                begin
                        TotalCout := GetField('PST_COUTMOB')+GetField('PST_COUTORGAN')+GetField('PST_COUTFONCT');
                        If TotalCout <> 0 Then
                        begin
                                SetField('PST_COUTBUDGETE',TotalCout);
                                SetControlEnabled('PST_COUTBUDGETE',False);
                        end
                        Else SetControlEnabled('PST_COUTBUDGETE',True);
                end;
        end;
        //DEBUT PT6
        //PT8 - Début
        {If (F.FieldName='PST_DUREESTAGE') or (F.FieldName='PST_NBANIM') Then
        begin
             Millesime := GetField('PST_MILLESIME');
             If Millesime <> '0000' then
             begin
                  NbHeure := GetField('PST_DUREESTAGE');
                  NbAnim := GetField('PST_NBANIM');
                  QAnim := OpenSQl('SELECT PFE_SALAIREANIM FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
                  salaireanim := 0;
                  if not QAnim.eof then SalaireAnim := QAnim.FindField('PFE_SALAIREANIM').AsFloat;
                  Ferme(QAnim);
                  SetField('PST_COUTSALAIR', SalaireAnim * NbHeure * NbAnim);
             end;
        end; }

        If (F.FieldName='PST_COUTSALAIR') Then
             If GetField('PST_COUTSALAIR') <> CalcSalaireAnimateur Then
                  SalAnimUser := True;
        //PT8 - Fin
        //FIN PT6
end;


procedure TOM_STAGE.OnLoadRecord;
var ExisteSession,ExisteBudget : Boolean;
//    i : Integer;
//    Q : TQuery;
//    TobInvest : Tob;
begin
  inherited;
        If Not AfterInserting then
        begin
                ExisteSession := ExisteSQL('SELECT PSS_CODESTAGE FROM SESSIONSTAGE WHERE PSS_MILLESIME="'+GetField('PST_MILLESIME')+'" AND PSS_CODESTAGE="'+GetField('PST_CODESTAGE')+'"');
                If GetField('PST_MILLESIME') <> '0000' Then ExisteBudget := ExisteSQL('SELECT PFI_ETABLISSEMENT FROM INSCFORMATION WHERE PFI_MILLESIME="'+GetField('PST_MILLESIME')+'" AND PFI_CODESTAGE="'+GetField('PST_CODESTAGE')+'"')
                Else ExisteBudget := ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="'+getField('PST_CODESTAGE')+'" AND PST_MILLESIME <> "0000"');
                If  (ExisteBudget) or (ExisteSession)Then
                begin
//                        SetControlEnabled('PCOMP',False);
                        SetControlEnabled('PDESCRIPTION',False);
                        GriserDescriptif(False);
                end
                else
                begin
                        SetControlEnabled('PGENERAL',True);
                        SetControlEnabled('PCOMP',True);
                        SetControlEnabled('PDESCRIPTION',True);
                        GriserDescriptif(True);
                end;
        end
        else
        begin
                SetControlEnabled('BOBJECTIFS',False);
                SetControlEnabled('BFORMATION',False);
                SetControlEnabled('BDUPLIQUER',False);
        end;
        If TypeSaisie = 'SAISIECAT' Then
        begin
                TFFiche(Ecran).Caption := 'Catalogue : formation '+GetField('PST_CODESTAGE')+' '+GetField('PST_LIBELLE');
                UpdateCaption(Ecran);
        end;
        //DEBUT PT1
        //PT10 - Début
        //PT11 - Début
        {$IFDEF EMANAGER}
        If TypeSaisie = 'CONSULTCAT' then
        begin
          If existeSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_LIBELLE="'+V_PGI.userSalarie+'" AND PCC_CODESTAGE="'+GetField('PST_CODESTAGE')+'"') then
          Begin
               SetControlEnabled('BFAVORIS',False);
               SetControlEnabled('BNONFAVORIS',True);
          End
          Else
          Begin
               SetControlEnabled('BFAVORIS',True);
               SetControlEnabled('BNONFAVORIS',False);
          End;
        end;
        {$ENDIF}
        //PT11 - Fin
        //Q := OpenSQL('SELECT PIS_INVESTFORM,PIS_COUTPEDAG,PIS_COUTSALAIRE,PIS_FRAISFORM,PIS_TAUXSALFRAIS,PIS_TAUXPEDAG,PIS_TAUXSALAIRE,PIS_MONTANT FROM INVESTSESSION '+
        //'WHERE PIS_CODESTAGE="'+GetField('PST_CODESTAGE')+'" AND PIS_ORDRE=-1 AND PIS_MILLESIME="'+GetField('PST_MILLESIME')+'"',True); //DB2
        //TobInvest := Tob.Create('Les investissements',Nil,-1);
        //TobInvest.LoadDetailDB('INVESTSESSION','','',Q,False);
        //Ferme(Q);
        //
        //For i := 0 to TobInvest.Detail.Count - 1 do
        //begin
        //        GInvest.CellValues[0,i+1] :=  TobInvest.Detail[i].GetValue('PIS_INVESTFORM');
        //        If TobInvest.Detail[i].GetValue('PIS_COUTPEDAG') = 'X' then GInvest.CellValues[1,i+1] := 'OUI'
        //        else GInvest.CellValues[1,i+1] := 'NON';
        //        GInvest.CellValues[2,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_TAUXPEDAG'));
        //        If TobInvest.Detail[i].GetValue('PIS_COUTSALAIRE') = 'X' then GInvest.CellValues[3,i+1] := 'OUI'
        //        else GInvest.CellValues[3,i+1] := 'NON';
        //        GInvest.CellValues[4,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_TAUXSALAIRE'));
        //        If TobInvest.Detail[i].GetValue('PIS_FRAISFORM') = 'X' then GInvest.CellValues[5,i+1] := 'OUI'
        //        else GInvest.CellValues[5,i+1] := 'NON';
        //        GInvest.CellValues[6,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_TAUXSALFRAIS'));
        //        GInvest.CellValues[7,i+1] :=  '0';
        //end;
        //TobInvest.Free;
        //PT10 - Fin
        //FIN PT1
end;

procedure TOM_STAGE.OnAfterUpdateRecord;
begin
inherited;
        SetControlEnabled('BOBJECTIFS',True);
        SetControlEnabled('BFORMATION',True);
        SetControlEnabled('BDUPLIQUER',True);
        MajInvest;
        AvertirTable('PGSTAGEFORMCOMPLET');
        AvertirTable('PGSTAGEFORMINITIALE');
end;

procedure TOM_STAGE.OnNewRecord;
var Q : TQuery;
    IMax : Integer;
begin
  inherited;
        SetField('PST_NODOSSIER',V_PGI.NoDossier);
        If PGDroitMultiForm then SetField('PST_PREDEFINI','STD')
        else SetField('PST_PREDEFINI','DOS'); //PT9
        If TypeSaisie='INITIALE' then SetField('PST_NATUREFORM','004')
        else SetField('PST_NATUREFORM','001');
        SetField('PST_ACTIF','X');
        SetField('PST_MILLESIME','0000');
        SetField('PST_TYPCONVFORM','ANN');
        SetField('PST_ORGCOLLECTP',-1);
        SetField('PST_ORGCOLLECTS',-1);
        SetField('PST_ACCEPTBUD','-');
        SetField('PST_NBANIM',1);
        SetField('PST_COMPTERENDU','-');
        SetField('PST_ATTESTPRESENC','-');
        SetField('PST_QUESTAPPREC','-');
        SetField('PST_QUESTEVALUAT','-');
        SetField('PST_SUPPCOURS','-');
        SetField('PST_VIDEOPROJ','-');
        SetField('PST_RETROPROJ','-');
        SetField('PST_JEUXROLE','-');
        SetField('PST_ETUDECAS','-');
        SetField('PST_INCLUSDECL','X');
        If StageDupliquer <> '' Then
        begin
                RecupInfosStage;
                StageDupliquer := '';
        end;
        If VH_Paie.PGStageFormAuto=True Then
        begin
                Q := OpenSQL('SELECT MAX(PST_CODESTAGE) FROM STAGE',TRUE) ;
                if NOT Q.Eof then
                begin
                        If Q.Fields[0].AsString <> '' Then Imax := Q.Fields[0].AsInteger
                        Else Imax := 0;
                        SetField ('PST_CODESTAGE',IntToStr (Imax + 1));
                end
                else SetField ('PST_CODESTAGE',IntToStr(1));
                Ferme(Q) ;
        end;
        //DEBUT PT2
        //GInvest.CellValues[0,1] := 'PDF'; //PT13
        GInvest.CellValues[0,1] := 'PLF'; 
        GInvest.CellValues[1,1] := 'OUI';
        GInvest.CellValues[2,1] := '1';
        GInvest.CellValues[3,1] := 'OUI';
        GInvest.CellValues[4,1] := '1';
        GInvest.CellValues[5,1] := 'OUI';
        GInvest.CellValues[6,1] := '1';
        //FIN PT2
end;

procedure TOM_STAGE.BObjectifsclick(Sender:TObject);
begin
        AglLanceFiche('PAY','MUL_OBJECTIFSTAGE', '', '', TypeSaisie+';'+GetField('PST_CODESTAGE'));
end;

procedure TOM_STAGE.BFormationClick(Sender:TObject);
begin
        AglLanceFiche('PAY','SAISIEINSCBUDGET','','','SAISIESTAGE;;'+GetField('PST_CODESTAGE')+';'+GetField('PST_MILLESIME'));
end;

procedure TOM_STAGE.BDupliquerClick(Sender:TObject);
begin
        StageDupliquer := GetField('PST_CODESTAGE');
        TFFiche(Ecran).BinsertClick(TFFiche                              (Ecran).BInsert);
end;

procedure TOM_STAGE.RecupInfosStage;
var Q : TQuery;
begin
        Q := OpenSQL('SELECT * FROM STAGE WHERE PST_CODESTAGE="'+StageDupliquer+'" AND PST_MILLESIME="0000"',True);
        if not Q.eof then
        begin
                SetField('PST_LIBELLE',Q.FindField('PST_LIBELLE').AsString);
                SetField('PST_LIBELLE1',Q.FindField('PST_LIBELLE1').AsString);
                SetField('PST_DIPLOMANT',Q.FindField('PST_DIPLOMANT').AsString);
                SetField('PST_CENTREFORMGU',Q.FindField('PST_CENTREFORMGU').AsString);
                SetField('PST_COUTSALAIR',Q.FindField('PST_COUTSALAIR').AsFloat);
                SetField('PST_COUTMOB',Q.FindField('PST_COUTMOB').AsFloat);
                SetField('PST_COUTFONCT',Q.FindField('PST_COUTFONCT').AsFloat);
                SetField('PST_COUTORGAN',Q.FindField('PST_COUTORGAN').AsFloat);
                SetField('PST_COUTUNITAIRE',Q.FindField('PST_COUTUNITAIRE').AsFloat);
                SetField('PST_COUTBUDGETE',Q.FindField('PST_COUTBUDGETE').AsFloat);
                SetField('PST_FORMATION1',Q.FindField('PST_FORMATION1').AsString);
                SetField('PST_FORMATION2',Q.FindField('PST_FORMATION2').AsString);
                SetField('PST_FORMATION3',Q.FindField('PST_FORMATION3').AsString);
                SetField('PST_FORMATION4',Q.FindField('PST_FORMATION4').AsString);
                SetField('PST_FORMATION5',Q.FindField('PST_FORMATION5').AsString);
                SetField('PST_FORMATION6',Q.FindField('PST_FORMATION6').AsString);
                SetField('PST_FORMATION7',Q.FindField('PST_FORMATION7').AsString);
                SetField('PST_FORMATION8',Q.FindField('PST_FORMATION8').AsString);
                SetField('PST_DUREESTAGE',Q.FindField('PST_DUREESTAGE').AsFloat);
                SetField('PST_JOURSTAGE',Q.FindField('PST_JOURSTAGE').AsFloat);
                SetField('PST_NBSTAMIN',Q.FindField('PST_NBSTAMIN').AsInteger);
                SetField('PST_NBSTAMAX',Q.FindField('PST_NBSTAMAX').AsInteger);
                SetField('PST_NATUREFORM',Q.FindField('PST_NATUREFORM').AsString);
                SetField('PST_TYPCONVFORM',Q.FindField('PST_TYPCONVFORM').AsString);
                SetField('PST_ORGCOLLECTPGU',Q.FindField('PST_ORGCOLLECTPGU').AsString);
                SetField('PST_ORGCOLLECTSGU',Q.FindField('PST_ORGCOLLECTSGU').AsString);
                SetField('PST_LIEUFORM',Q.FindField('PST_LIEUFORM').AsString);
                SetField('PST_ACTIF',Q.FindField('PST_ACTIF').AsString);
                SetField('PST_ACCEPTBUD',Q.FindField('PST_ACCEPTBUD').AsString);
                SetField('PST_NBANIM',Q.FindField('PST_NBANIM').AsInteger);
                SetField('PST_FRAISANIM',Q.FindField('PST_FRAISANIM').AsString); //PT4
                SetField('PST_COMPTERENDU',Q.FindField('PST_COMPTERENDU').AsString);
                SetField('PST_ATTESTPRESENC',Q.FindField('PST_ATTESTPRESENC').AsString);
                SetField('PST_QUESTAPPREC',Q.FindField('PST_QUESTAPPREC').AsString);
                SetField('PST_QUESTEVALUAT',Q.FindField('PST_QUESTEVALUAT').AsString);
                SetField('PST_AUTREEVALUAT',Q.FindField('PST_AUTREEVALUAT').AsString);
                SetField('PST_SUPPCOURS',Q.FindField('PST_SUPPCOURS').AsString);
                SetField('PST_VIDEOPROJ',Q.FindField('PST_VIDEOPROJ').AsString);
                SetField('PST_RETROPROJ',Q.FindField('PST_RETROPROJ').AsString);
                SetField('PST_JEUXROLE',Q.FindField('PST_JEUXROLE').AsString);
                SetField('PST_ETUDECAS',Q.FindField('PST_ETUDECAS').AsString);
                SetField('PST_AUTREMOYEN',Q.FindField('PST_AUTREMOYEN').AsString);
        end;
        Ferme(Q);
end;

procedure TOM_STAGE.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
        ForceUpdate;
        If ACol = 0 then
        begin
                If (GInvest.CellValues[0,ARow] <> '') AND (GInvest.CellValues[0,ARow] <> 'AUC') then
                begin
                        If GInvest.CellValues[1,ARow] = '' then GInvest.CellValues[1,ARow] := 'OUI';
                        If GInvest.CellValues[2,ARow] = '' then GInvest.CellValues[2,ARow] := '1';
                        If GInvest.CellValues[3,ARow] = '' then GInvest.CellValues[3,ARow] := 'OUI';
                        If GInvest.CellValues[4,ARow] = '' then GInvest.CellValues[4,ARow] := '1';
                        If GInvest.CellValues[5,ARow] = '' then GInvest.CellValues[5,ARow] := 'OUI';
                        If GInvest.CellValues[6,ARow] = '' then GInvest.CellValues[6,ARow] := '1';
                end
                else
                begin
                        GInvest.CellValues[1,ARow] := '';
                        GInvest.CellValues[2,ARow] := '0';
                        GInvest.CellValues[3,ARow] := '';
                        GInvest.CellValues[4,ARow] := '0';
                        GInvest.CellValues[5,ARow] := '';
                        GInvest.CellValues[6,ARow] := '0';
                end;
        end;
end;

procedure TOM_STAGE.MajInvest;
var Q:Tquery;
    TobInvest,TI : tob;
    i : Integer;
begin
        ExecuteSQL('DELETE FROM INVESTSESSION WHERE PIS_CODESTAGE="'+GetField('PST_CODESTAGE')+'" AND PIS_ORDRE=-1 AND PIS_MILLESIME="'+GetField('PST_MILLESIME')+'"'); //DB2
        Q := OpenSQL('SELECT * FROM INVESTSESSION '+
        'WHERE PIS_CODESTAGE="'+GetField('PST_CODESTAGE')+'" AND PIS_ORDRE=-1 AND PIS_MILLESIME="'+GetField('PST_MILLESIME')+'"',True); //DB2
        TobInvest := Tob.Create('INVESTSESSION',Nil,-1);
        TobInvest.LoadDetailDB('INVESTSESSION','','',Q,False);
        Ferme(Q);
        For i := 1 to 2 do
        begin
                If (GInvest.CellValues[0,i] <> '') AND (GInvest.CellValues[0,i] <> 'AUC') then
                begin
                        TI := Tob.Create('INVESTSESSION',TobInvest,-1);
                        TI.PutValue('PIS_CODESTAGE',GetField('PST_CODESTAGE'));
                        TI.PutValue('PIS_ORDRE',-1);
                        TI.PutValue('PIS_MILLESIME',GetField('PST_MILLESIME'));
                        TI.PutValue('PIS_INVESTFORM',GInvest.CellValues[0,i]);
                        If GInvest.CellValues[1,i] = 'OUI' then TI.PutValue('PIS_COUTPEDAG','X')
                        else TI.PutValue('PIS_COUTPEDAG','-');
                        If IsNumeric(GInvest.CellValues[2,i]) then TI.PutValue('PIS_TAUXPEDAG',StrToFloat(GInvest.CellValues[2,i]));
                        If GInvest.CellValues[3,i] = 'OUI' then TI.PutValue('PIS_COUTSALAIRE','X')
                        else TI.PutValue('PIS_COUTSALAIRE','-');
                        If IsNumeric(GInvest.CellValues[4,i]) then TI.PutValue('PIS_TAUXSALAIRE',StrToFloat(GInvest.CellValues[4,i]));
                        If GInvest.CellValues[5,i] = 'OUI' then TI.PutValue('PIS_FRAISFORM','X')
                        else TI.PutValue('PIS_FRAISFORM','-');
                        If IsNumeric(GInvest.CellValues[6,i]) then TI.PutValue('PIS_TAUXSALFRAIS',StrToFloat(GInvest.CellValues[6,i]));
                        If IsNumeric(GInvest.CellValues[7,i]) then TI.PutValue('PIS_MONTANT',StrToFloat(GInvest.CellValues[7,i]));
                        TI.InsertDB(Nil,False);
                end;
        end;
        TobInvest.Free;
end;

procedure TOM_STAGE.GriserDescriptif(Acces : Boolean); //PT3
begin
     SetControlEnabled('PST_COMPTERENDU',Acces);
     SetControlEnabled('PST_QUESTAPPREC',Acces);
     SetControlEnabled('PST_ATTESTPRESENC',Acces);
     SetControlEnabled('PST_QUESTEVALUAT',Acces);
     SetControlEnabled('PST_AUTREEVALUAT',Acces);
     SetControlEnabled('PST_SUPPCOURS',Acces);
     SetControlEnabled('PST_RETROPROJ',Acces);
     SetControlEnabled('PST_ETUDECAS',Acces);
     SetControlEnabled('PST_VIDEOPROJ',Acces);
     SetControlEnabled('PST_JEUXROLE',Acces);
     SetControlEnabled('PST_AUTREMOYEN',Acces);
     SetControlEnabled('TPST_AUTREEVALUAT',Acces);
     SetControlEnabled('TPST_AUTREMOYEN',Acces);
end;

{procedure TOM_STAGE.MajPrev;
var Q : TQuery;
    TStage : Tob;
    Rep,i : Integer;
    Millesime : String;
begin
     Millesime := RendMillesimePrevisionnel;
     Q := OpenSQL('SELECT * FROM STAGE WHERE PST_CODESTAGE="'+GetField('PST_CODESTAGE')+'" AND PST_MILLESIME="'+Millesime+'"',True);
     TStage := Tob.Create('STAGE',Nil,-1);
     TStage.LoadDetailDB('STAGE','','',Q,False);
     Ferme(Q);
     If TStage.Detail.Count > 0 then
     begin
          Rep := PGIAsk('Voulez vous mettre à jour le stage prévisionnel '+Millesime);
          If Rep = MrYes then
          begin
               For i := 0 to TStage.Detail.Count - 1 do
               begin
                     TStage.Detail[i].PutValue('PST_LIBELLE',GetField('PST_LIBELLE'));
                     TStage.Detail[i].PutValue('PST_LIBELLE1',GetField('PST_LIBELLE1'));
                     TStage.Detail[i].PutValue('PST_DIPLOMANT',GetField('PST_DIPLOMANT'));
                     TStage.Detail[i].PutValue('PST_CENTREFORMGU',GetField('PST_CENTREFORMGU'));
                     TStage.Detail[i].PutValue('PST_COUTSALAIR',GetField('PST_COUTSALAIR'));
                     TStage.Detail[i].PutValue('PST_COUTMOB',GetField('PST_COUTMOB'));
                     TStage.Detail[i].PutValue('PST_COUTFONCT',GetField('PST_COUTFONCT'));
                     TStage.Detail[i].PutValue('PST_COUTORGAN',GetField('PST_COUTORGAN'));
                     TStage.Detail[i].PutValue('PST_COUTUNITAIRE',GetField('PST_COUTUNITAIRE'));
                     TStage.Detail[i].PutValue('PST_COUTEVAL',GetField('PST_COUTEVAL'));
                     TStage.Detail[i].PutValue('PST_COUTBUDGETE',GetField('PST_COUTBUDGETE'));
                     TStage.Detail[i].PutValue('PST_FORMATION1',GetField('PST_FORMATION1'));
                     TStage.Detail[i].PutValue('PST_FORMATION2',GetField('PST_FORMATION2'));
                     TStage.Detail[i].PutValue('PST_FORMATION3',GetField('PST_FORMATION3'));
                     TStage.Detail[i].PutValue('PST_FORMATION4',GetField('PST_FORMATION4'));
                     TStage.Detail[i].PutValue('PST_FORMATION5',GetField('PST_FORMATION5'));
                     TStage.Detail[i].PutValue('PST_FORMATION6',GetField('PST_FORMATION6'));
                     TStage.Detail[i].PutValue('PST_FORMATION7',GetField('PST_FORMATION7'));
                     TStage.Detail[i].PutValue('PST_FORMATION8',GetField('PST_FORMATION8'));
                     TStage.Detail[i].PutValue('PST_INCLUSDECL',GetField('PST_INCLUSDECL'));
                     TStage.Detail[i].PutValue('PST_DUREESTAGE',GetField('PST_DUREESTAGE'));
                     TStage.Detail[i].PutValue('PST_JOURSTAGE',GetField('PST_JOURSTAGE'));
                     TStage.Detail[i].PutValue('PST_NBSTAMIN',GetField('PST_NBSTAMIN'));
                     TStage.Detail[i].PutValue('PST_NBSTAMAX',GetField('PST_NBSTAMAX'));
                     TStage.Detail[i].PutValue('PST_NATUREFORM',GetField('PST_NATUREFORM'));
                     TStage.Detail[i].PutValue('PST_TYPCONVFORM',GetField('PST_TYPCONVFORM'));
                     TStage.Detail[i].PutValue('PST_ACTIONFORM',GetField('PST_ACTIONFORM'));
                     TStage.Detail[i].PutValue('PST_ORGCOLLECTPGU',GetField('PST_ORGCOLLECTPGU'));
                     TStage.Detail[i].PutValue('PST_ORGCOLLECTSGU',GetField('PST_ORGCOLLECTSGU'));
                     TStage.Detail[i].PutValue('PST_LIEUFORM',GetField('PST_LIEUFORM'));
                     TStage.Detail[i].PutValue('PST_ACTIF',GetField('PST_ACTIF'));
                     TStage.Detail[i].PutValue('PST_ACCEPTBUD',GetField('PST_ACCEPTBUD'));
                     TStage.Detail[i].PutValue('PST_NBANIM',GetField('PST_NBANIM'));
                     TStage.Detail[i].PutValue('PST_FRAISANIM',GetField('PST_FRAISANIM'));
                     TStage.Detail[i].PutValue('PST_COMPTERENDU',GetField('PST_COMPTERENDU'));
                     TStage.Detail[i].PutValue('PST_ATTESTPRESENC',GetField('PST_ATTESTPRESENC'));
                     TStage.Detail[i].PutValue('PST_QUESTAPPREC',GetField('PST_QUESTAPPREC'));
                     TStage.Detail[i].PutValue('PST_QUESTEVALUAT',GetField('PST_QUESTEVALUAT'));
                     TStage.Detail[i].PutValue('PST_AUTREEVALUAT',GetField('PST_AUTREEVALUAT'));
                     TStage.Detail[i].PutValue('PST_SUPPCOURS',GetField('PST_SUPPCOURS'));
                     TStage.Detail[i].PutValue('PST_VIDEOPROJ',GetField('PST_VIDEOPROJ'));
                     TStage.Detail[i].PutValue('PST_RETROPROJ',GetField('PST_RETROPROJ'));
                     TStage.Detail[i].PutValue('PST_JEUXROLE',GetField('PST_JEUXROLE'));
                     TStage.Detail[i].PutValue('PST_ETUDECAS',GetField('PST_ETUDECAS'));
                     TStage.Detail[i].PutValue('PST_AUTREMOYEN',GetField('PST_AUTREMOYEN'));
                     TStage.Detail[i].PutValue('PST_REGLEMENTAIRE',GetField('PST_REGLEMENTAIRE'));
                     TStage.Detail[i].PutValue('PST_NIVEAUFORMINIT',GetField('PST_NIVEAUFORMINIT'));
                     TStage.Detail[i].PutValue('PST_TYPEFORMINIT',GetField('PST_TYPEFORMINIT'));
                     TStage.Detail[i].PutValue('PST_DOMFORMINIT',GetField('PST_DOMFORMINIT'));
                     TStage.Detail[i].UpdateDB;
               end;
          end;
     end;
     TStage.Free;
end;}

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/05/2007
Modifié le ... :   /  /
Description .. : Affichage de l'écran de sélection des emplois valides pour le
Suite ........ : stage  / PT7
Mots clefs ... :
*****************************************************************}
procedure TOM_STAGE.BEmploisClick(Sender: TOBject);
Var StAction : String;
begin
    If TFFiche(Ecran).FTypeAction = taConsult Then StAction := 'CONSULTATION' Else StAction := 'MODIFICATION'; //PT16
        AglLanceFiche('PAY','MULSTAGEEMPLOI', '', '', GetField('PST_CODESTAGE')+';-1;'+StAction);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 12/06/2007 / PT8
Modifié le ... :   /  /
Description .. : Calcule le montant des salaires animateurs
Mots clefs ... :
*****************************************************************}
function TOM_STAGE.CalcSalaireAnimateur: Double;
var NbHeure,NbAnim : Double;
    QAnim : TQuery;
    Millesime : String;
begin
     Result := 0;

     Millesime := GetField('PST_MILLESIME');
     If Millesime <> '0000' then
     begin
          NbHeure := GetField('PST_DUREESTAGE');
          NbAnim := GetField('PST_NBANIM');

          // On n'effectue la requête qu'une seule fois en testant la valeur d'initialisation de la variable à -1
          If SalaireAnimateur = -1 Then
          Begin
               QAnim := OpenSQl('SELECT PFE_SALAIREANIM,PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
               SalaireAnimateur := 0;
               TauxBudget := 0;
               if not QAnim.eof then
               Begin
                    SalaireAnimateur := QAnim.FindField('PFE_SALAIREANIM').AsFloat;
                    TauxBudget := QAnim.FindField('PFE_TAUXBUDGET').AsFloat;
               End;

               Ferme(QAnim);
          End;

          Result := (SalaireAnimateur * NbHeure * NbAnim * TauxBudget);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/11/2007 / PT10
Modifié le ... :   /  /
Description .. : Mise en forme de la grille d'investissement
Suite ........ : et chargement des données associées au stage
Mots clefs ... :
*****************************************************************}
procedure TOM_STAGE.onEnterOngletCouts(Sender: TObject);
var i : Integer;
    Q : TQuery;
    TobInvest : Tob;
begin
    If (Not GrilleChargee) And (GInvest <> Nil) Then
    Begin
        //GInvest.ColFormats[0] := 'CB=PGFINANCEFORMATION'; //PT13
        GInvest.ColFormats[0] := 'CB=PGTYPEPLANPREV'; 
        GInvest.ColFormats[1] := 'CB=PGOUINON';
        GInvest.ColFormats[2] := '# ##0.00';
        GInvest.ColFormats[3] := 'CB=PGOUINON';
        GInvest.ColFormats[4] := '# ##0.00';
        GInvest.ColFormats[5] := 'CB=PGOUINON';
        GInvest.ColFormats[6] := '# ##0.00';
        GInvest.ColFormats[7] := '# ##0';
        GInvest.OnCellExit := OnCellExit;

        Q := OpenSQL('SELECT PIS_INVESTFORM,PIS_COUTPEDAG,PIS_COUTSALAIRE,PIS_FRAISFORM,PIS_TAUXSALFRAIS,PIS_TAUXPEDAG,PIS_TAUXSALAIRE,PIS_MONTANT FROM INVESTSESSION '+
        'WHERE PIS_CODESTAGE="'+GetField('PST_CODESTAGE')+'" AND PIS_ORDRE=-1 AND PIS_MILLESIME="'+GetField('PST_MILLESIME')+'"',True); //DB2
        TobInvest := Tob.Create('Les investissements',Nil,-1);
        TobInvest.LoadDetailDB('INVESTSESSION','','',Q,False);
        Ferme(Q);

        For i := 0 to TobInvest.Detail.Count - 1 do
        begin
            GInvest.CellValues[0,i+1] :=  TobInvest.Detail[i].GetValue('PIS_INVESTFORM');
            If TobInvest.Detail[i].GetValue('PIS_COUTPEDAG') = 'X' then GInvest.CellValues[1,i+1] := 'OUI'
            else GInvest.CellValues[1,i+1] := 'NON';
            GInvest.CellValues[2,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_TAUXPEDAG'));
            If TobInvest.Detail[i].GetValue('PIS_COUTSALAIRE') = 'X' then GInvest.CellValues[3,i+1] := 'OUI'
            else GInvest.CellValues[3,i+1] := 'NON';
            GInvest.CellValues[4,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_TAUXSALAIRE'));
            If TobInvest.Detail[i].GetValue('PIS_FRAISFORM') = 'X' then GInvest.CellValues[5,i+1] := 'OUI'
            else GInvest.CellValues[5,i+1] := 'NON';
            GInvest.CellValues[6,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_TAUXSALFRAIS'));
            GInvest.CellValues[7,i+1] :=  '0';
        end;
        TobInvest.Free;

        GrilleChargee := True;
    End;

{$IFDEF EMANAGER}
{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 24/10/2007 / PT11
Modifié le ... :   /  /
Description .. : Ajout d'un stage aux favoris
Mots clefs ... :
*****************************************************************}
procedure TOM_STAGE.AjouterAuxFavoris(Sender : Tobject);
var Q : TQuery;
    TobCursus : Tob;
    IMax : Integer;
begin
     Q := OpenSQL('SELECT MAX(PCC_ORDRE) FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_CODESTAGE="'+GetField('PST_CODESTAGE')+'"',True);
     if Not Q.EOF then
     begin
          IMax := Q.Fields[0].AsInteger;
          if IMax <> 0 then IMax := IMax + 1
          else IMax := 1;
     end
     else IMax := 1 ;
     Ferme(Q) ;
     TobCursus := Tob.Create('CURSUSSTAGE',Nil,-1);
     TobCursus.PutValue('PCC_CURSUS','---');
     TobCursus.PutValue('PCC_LIBELLE',V_PGI.UserSalarie);
     TobCursus.PutValue('PCC_ORDRE',-1);
     TobCursus.PutValue('PCC_RANGCURSUS',Imax);
     TobCursus.PutValue('PCC_CODESTAGE',GetField('PST_CODESTAGE'));
     TobCursus.PutValue('PCC_MILLESIME',MillesimeEC);
     TobCursus.InsertDB(Nil);
     SetControlENabled('BFAVORIS',False);
     SetControlEnabled('BNONFAVORIS',True);
     TobCursus.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 24/10/2007 / PT11
Modifié le ... :   /  /
Description .. : Suppression d'un stage des favoris
Mots clefs ... :
*****************************************************************}
procedure TOM_STAGE.RetirerDesFavoris(Sender : Tobject);
begin
    
     ExecuteSQL('DELETE FROM CURSUSSTAGE WHERE PCC_LIBELLE="'+V_PGI.UserSalarie+'" AND '+
     'PCC_CURSUS="---" AND PCC_MILLESIME="'+MillesimeEC+'" AND PCC_CODESTAGE="'+getField('PST_CODESTAGE')+'"');
     SetControlEnabled('BFAVORIS',True);
     SetControlEnabled('BNONFAVORIS',False);
end;
{$ENDIF}

end;

Initialization
    registerclasses([TOM_STAGE]) ;
end.

