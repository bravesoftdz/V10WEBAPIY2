{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULSTAGE ()
Mots clefs ... : TOF;PGMULSTAGE
*****************************************************************
PT1 | 24/04/2003 | V_42  | JL | Développement pour CWAS
PT2 | 16/11/2004 | V_60  | JL | Modifs pour eFormation + accès maj cout prévisionnel
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT3 | 14/05/2007 | V_720 | FL | FQ 13567 Vérification de la validité de la population avant saisie au budget
PT4 | 15/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT5 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT6 | 15/04/2008 | V_804 | FL | Prise en compte du millésime sélectionné lors d'une inscription, même si le stage n'est pas encore ouvert
PT7 | 17/04/2008 | V_804 | FL | Ajout de PGBundleCatalogue
PT8 | 17/04/2008 | V_804 | FL | Ne pas autoriser la modification d'éléments STD si pas droits multi
}
Unit UTOFPGMulStage;

Interface                                          

Uses StdCtrls,Controls,Classes,                     
{$IFNDEF EAGLCLIENT}
     DBGrids, HDB,Mul,Fiche,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ELSE}
     UtileAGL,MaineAgl,emul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,Hqry,HTB97,UTob,PgoutilsFormation,EntPaie,UTomInscrFormation,PGPopulOutils ;

Type
TOF_PGMULSTAGEINIT = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
    Q_Mul : THQuery ;
    Bt : TToolbarButton97;
    procedure GrilleDblClick(Sender : TObject);
    procedure NouveauStage(Sender : TObject);
  end ;

  TOF_PGMULSTAGEEFOR = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    MillesimeEC,TypeSaisie,typeUtilisat : String;
    THVMillesime : THValComboBox;
    Previsionnel : Boolean;
    Q_MulEfor : THQuery;
    {$IFNDEF EAGLCLIENT}
    ListeEfor : THDBGrid;
    {$ELSE}
    ListeEfor : THGrid;
    {$ENDIF}
    procedure GrilleDblClick(Sender : TObject);
  end ;

  TOF_PGMULSTAGE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TypeUtilisat,MillesimeEC,TypeSaisie : String;
    THVMillesime : THValComboBox;
    Previsionnel : Boolean;
    procedure GrilleDblClick(Sender : TObject);
    procedure NouveauStage(Sender : TObject);
    procedure ChangeMillesime(Sender : TObject);
  end ;
var
    Q_MulStage : THQuery;
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    TFMStage : TFMul;
    {$ENDIF}

Implementation
Uses GalOutil;

procedure TOF_PGMULSTAGE.NouveauStage;
{$IFNDEF EMANAGER}
Var
Bt  :  TToolbarButton97;
{$ENDIF}
begin
        {$IFNDEF EMANAGER}
        AglLanceFiche('PAY','STAGE','','','SAISIECAT;ACTION=CREATION');
        Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
        {$ELSE}
        AglLanceFiche('PAY','MUL_STAGEEFOR','','','NEWINSC');
        {$ENDIF}
end;

procedure TOF_PGMULSTAGE.OnLoad ;
var ChoixStage :  THValComboBox;
    SQL,StWhere : String;
begin
  Inherited ;
        {$IFDEF EMANAGER}
        THVMillesime.Value := MillesimeEC;
        SetControlText('CHOIXSTAGE','SB');
        {$ENDIF}
        StWhere := '';
        MillesimeEC  := THVMillesime.Value;
        ChoixStage := THValComboBox(GetControl('CHOIXSTAGE'));
        If TypeSaisie = 'SAISIEBUD' Then
        begin
                TFMul(Ecran).Caption := 'Formations budgetisées en '+MillesimeEC;
                UpdateCaption(TFMul(Ecran)) ;
                SetControltext('PST_MILLESIME',MillesimeEC);
        end;
        If TypeSaisie = 'CALCSALAIRE' Then
        begin
                TFMul(Ecran).Caption := 'Maj des coûts pour les formations budgetisées en '+MillesimeEC;
                UpdateCaption(TFMul(Ecran)) ;
                SetControltext('PST_MILLESIME',MillesimeEC);
        end;
        If (TypeSaisie = 'RECUPSTAGEDIF') or (TypeSaisie = 'SESSION') OR (TypeSaisie = 'INSCBUDGET') Then
        begin
                If ChoixStage.Value = 'SB' Then SetControlText('PST_MILLESIME',MillesimeEC)
                Else If ChoixStage.Value = 'SNB' Then
                begin
                        SetControlText('PST_MILLESIME','0000');
                        StWhere := 'PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'")';
                end
                Else
                begin
                        SetControlText('PST_MILLESIME','');
                        SQL := '((PST_MILLESIME="'+MillesimeEC+'") OR (PST_MILLESIME="0000" AND PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'")))';
                        StWhere := SQL;
                end;
        end;
        // Debut PT1
        If TypeSaisie = 'CWASINSCBUDGET' Then
        begin
                If ChoixStage.Value = 'SB' Then
                begin
                        SetControlText('PST_MILLESIME',MillesimeEC);
                        If TypeUtilisat = 'R' then StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE '+
                        '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'+
                        ' AND PFI_MILLESIME="'+MillesimeEC+'")'
                        else StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
                        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))) AND PFI_MILLESIME="'+MillesimeEC+'")';
                end
                Else If ChoixStage.Value = 'SNB' Then
                begin
                        SetControlText('PST_MILLESIME','0000');
                        If TypeUtilisat = 'R' then StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE NOT IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE '+
                        '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                        ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'+
                        ' AND PFI_MILLESIME="'+MillesimeEC+'")'
                        else StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE NOT IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))) AND PFI_MILLESIME="'+MillesimeEC+'")';
                end
                Else
                begin
                        SetControlText('PST_MILLESIME','');
                        SQL:= 'PST_ACTIF="X" AND ((PST_MILLESIME="'+MillesimeEC+'") OR (PST_MILLESIME="0000" AND PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'")))';
                        StWhere := SQL;
                end;
        end;
        If (TypeSaisie = 'SAISIECAT') or (TypeSaisie = 'CONSULCAT') Then
        begin
                If ChoixStage.Value = 'A' Then StWhere := 'PST_ACTIF="X"'
                Else If ChoixStage.value = 'NA' Then StWhere := 'PST_ACTIF="-"';
        end;
        If GetControlText('DECLARATION') = 'OUI' then
        begin
                If StWhere <> '' then StWhere := StWHere + ' AND PST_INCLUSDECL="X"'
                else StWhere := 'PST_INCLUSDECL="X"';
        end;
        If GetControlText('DECLARATION') = 'NON' then
        begin
                If StWhere <> '' then StWhere := StWHere + ' AND PST_INCLUSDECL="-"'
                else StWhere := 'PST_INCLUSDECL="-"';
        end;
        SetControlText('XX_WHERE',StWhere);

        {$IFNDEF EMANAGER}
        SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PST_PREDEFINI'),GetControlText('NODOSSIER'),'PST')); //PT4
        {$ENDIF}
end ;

procedure TOF_PGMULSTAGE.OnArgument (S : String ) ;
Var ChoixStage : THValComboBox;
    Q : TQuery;
    Num : Integer;
    Libelle,LibMillesime : THLabel;
    BOuv : TToolBarButton97;
begin
Inherited ;
        TypeSaisie := Trim(ReadTokenSt(S));
        millesimeEC := '';
        Previsionnel := False;
        {$IFDEF EMANAGER}
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        {$ENDIF}
        If (TypeSaisie = 'SAISIEBUD') or (TypeSaisie = 'INSCBUDGET') or (TypeSaisie = 'CWASINSCBUDGET') Then
        begin
                MillesimeEC := RendMillesimePrevisionnel;
                Previsionnel := True;
        end
        Else
        begin
                Q := OpenSQL('SELECT PFE_MILLESIME,PFE_CLOTUREPREV FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-" ORDER BY PFE_MILLESIME DESC',True);
                if not Q.eof then
                begin
                        MillesimeEC := Q.FindField('PFE_MILLESIME').AsString;
                        If Q.FindField('PFE_CLOTUREPREV').Asstring = 'X' Then Previsionnel := True
                        Else Previsionnel := False;
                end;
                Ferme(Q);
        end;
        Q_MulStage := THQuery(Ecran.FindComponent('Q'));
        THVMillesime := THValComboBox(GetControl('THVMILLESIME'));
        if THVMillesime <> Nil Then THVMillesime.OnChange := ChangeMillesime;
        ChoixStage := THValComboBox(GetControl('CHOIXSTAGE'));
        Libelle := THLabel(GetControl('TCHOIXSTAGE'));
        LibMillesime := THLabel(GetControl('LTHVMILLESIME'));
        If (TypeSaisie = 'INSCBUDGET') or (TypeSaisie = 'CWASINSCBUDGET') Then
        begin                     
                Libelle.Visible := True;
                Libelle.Caption := 'Liste';
                THVMillesime.visible := True;
                THVMillesime.Value := MillesimeEC;
                LibMillesime.Visible := True;
                ChoixStage.Items.Add('<<TOUS>>');
                ChoixStage.Items.Add('Formations budgetisées en '+MillesimeEC);
                ChoixStage.Items.Add('Catalogue formations');
                ChoixStage.Values.Add('TOUS');
                ChoixStage.Values.Add('SB');
                ChoixStage.Values.Add('SNB');
                ChoixStage.Value := 'SB';
                ChoixStage.Visible := True;
                SetControlVisible('BINSERT',False);
        end;
        If TypeSaisie = 'RECUPSTAGEDIF' then
        begin
                Libelle.Visible := True;
                Libelle.Caption := 'Liste';
                THVMillesime.visible := True;
                THVMillesime.Value := MillesimeEC;
                LibMillesime.Visible := True;
                ChoixStage.Items.Add('<<TOUS>>');
                ChoixStage.Items.Add('Formations budgetisées en '+MillesimeEC);
                ChoixStage.Items.Add('Catalogue formations');
                ChoixStage.Values.Add('TOUS');
                ChoixStage.Values.Add('SB');
                ChoixStage.Values.Add('SNB');
                ChoixStage.Value := 'SB';
                ChoixStage.Visible := True;
                SetControlVisible('BINSERT',False);
                TFMul(Ecran).Caption := 'Choix de la formation';
        end;
        If (TypeSaisie = 'SAISIEBUD') or (TypeSaisie = 'CALCSALAIRE')Then
        begin
                ChoixStage.Visible := False;
                SetControlVisible('BINSERT',False);
                SetControltext('PST_MILLESIME',MillesimeEC);
                THVMillesime.Visible := True;
                LibMillesime.Visible := True;
                THVMillesime.Value := MillesimeEC;
                SetControlText('PST_MILLESIME','2002');
                If TypeSaisie = 'CALCSALAIRE' then
                begin
                        TFMul(Ecran).Caption := 'Maj des coûts pour les formations budgetisées en '+MillesimeEC;
                        SetControlVisible('BSELECTALL',True);
                end
                else TFMul(Ecran).Caption := 'Formations budgetisées en '+MillesimeEC;
                Libelle.Visible := False;
        end;
        If TypeSaisie = 'BUDGETR' Then
        begin
                ChoixStage.Visible := False;
                THVMillesime.visible := False;
                LibMillesime.Visible := False;
                SetControlVisible('BINSERT',False);
                SetControlText('PST_MILLESIME','0000');
                TFMul(Ecran).Caption := 'Liste des formations disponibles';
                Libelle.Visible := False;
                SetControlVisible('BINSERT',False);
        end;
        If (TypeSaisie = 'SAISIECAT') or (TypeSaisie = 'CONSULTCAT') Then
        begin
                ChoixStage.Visible := True;
                Libelle.Visible := True;
                Libelle.Caption := 'Catalogue';
                THVMillesime.Visible := False;
                LibMillesime.Visible := False;
                SetControltext('PST_MILLESIME','0000');
                If TypeSaisie = 'CONSULTCAT' then SetControlVisible('BINSERT',False)
                Else SetControlVisible('BINSERT',True);
                TFMul(Ecran).Caption := 'Liste des formations';
                ChoixStage.Items.Add('<<TOUS>>');
                ChoixStage.Items.Add('Formations actives');
                ChoixStage.Items.Add('Formations non actives');
                ChoixStage.Values.Add('TOUS');
                ChoixStage.Values.Add('A');
                ChoixStage.Values.Add('NA');
                ChoixStage.Value := 'A';
        end;
        If TypeSaisie = 'SESSION' Then
        begin
                Libelle.Visible := True;
                Libelle.Caption := 'Budget';
                THVMillesime.visible := True;
                THVMillesime.Value := MillesimeEC;
                LibMillesime.Visible := True;
                ChoixStage.Items.Add('<<TOUS>>');
                ChoixStage.Items.Add('Formations budgetisées en '+MillesimeEC);
                ChoixStage.Items.Add('Catalogue formations');
                ChoixStage.Values.Add('TOUS');
                ChoixStage.Values.Add('SB');
                ChoixStage.Values.Add('SNB');
                ChoixStage.Value := '';
                ChoixStage.Visible := True;
                SetControlVisible('BINSERT',False);
        end;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste := THGrid(GetControl('FLISTE'));
        TFMStage := TFMUL(Ecran);
        {$ENDIF}
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        UpdateCaption(TFMul(Ecran)) ;
        if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
        BOuv  :=  TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv  <>  NIL then  BOuv.OnClick  :=  NouveauStage;
        SetControlText('DECLARATION','TOUS');
        {$IFDEF EMANAGER}
        SetControlEnabled('THVMILLESIME',False);
        SetControlVisible('PAvance',False);
        SetControlProperty('PAvance','TabVisible',False);
        SetControlVisible('BINSERT',True);
        SetControlEnabled('CHOIXSTAGE',False);
        {$ENDIF}

        {$IFNDEF EMANAGER}
         //DEBUT PT4
        If PGBundleCatalogue then //PT7
        begin
          If not PGDroitMultiForm then
          begin
            SetControlEnabled('NODOSSIER',False);
            SetControlText   ('NODOSSIER',V_PGI.NoDossier);
            SetControlEnabled('PST_PREDEFINI',False);
            SetControlText   ('PST_PREDEFINI',''); //PT4
            SetControlProperty ('PST_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT5
          end
	      Else If V_PGI.ModePCL='1' Then
	      Begin
	      	SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT5          
	      	SetControlProperty ('PST_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT5
	      End;
        end
        else
        begin
          SetControlVisible('PST_PREDEFINI', False);
          SetControlVisible('TPST_PREDEFINI',False);
          SetControlVisible('NODOSSIER',     False);
          SetControlVisible('TNODOSSIER',    False);
        end;
        //FIN PT4
        {$ENDIF}
end ;

procedure TOF_PGMULSTAGE.GrilleDblClick(Sender : TObject);
var Stage,Millesime,Pred,StAction : String;
    MillesimeEnCours : Boolean;
    Q : TQuery;
    {$IFNDEF EMANAGER}
    Bt  :  TToolbarButton97;
    {$ENDIF}
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;  //PT1
        {$ENDIF}
        If Q_MulStage.Eof then
        begin
                PGIBox('Vous devez sélectionner une formation',Ecran.Caption);
                Exit;
        end;
        If TypeSaisie = 'RECUPSTAGEDIF' then
        begin
             TFMul(Ecran).Retour := Q_MulStage.FindField ('PST_CODESTAGE').AsString;
             TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
             Exit;
        end;
        Stage := Q_MulStage.FindField ('PST_CODESTAGE').AsString;
        //Millesime := Q_MulStage.FindField ('PST_MILLESIME').AsString; //PT6
        Millesime := THVMillesime.Value;
        MillesimeEnCours := ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+GetControlText('THVMILLESIME')+'" '+
        'AND PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-"');
        If Stage <> '' Then
        begin
                If TypeSaisie = 'SAISIEBUD' Then
                begin
                        if MillesimeEnCours=True then AglLanceFiche('PAY','STAGE', '',Stage+';'+Millesime ,'SAISIEBUD;'+ Stage+';'+THVMillesime.Value+';ACTION=MODIFICATION' )
                        Else AglLanceFiche('PAY','STAGE', '',Stage+';'+Millesime ,'SAISIEBUD;'+ Stage+';'+THVMillesime.Value+';ACTION=CONSULTATION' );
                end;
                If TypeSaisie = 'INSCBUDGET' Then
                begin
                        //PT3 - Début
                        { Si la gestion des frais s'effectue par population et que celle propre à la formation
                          n'est pas valide, il faut interdire la saisie des inscriptions. }
                        If VH_Paie.PGForGestFraisByPop And Not CanUsePopulation(TYP_POPUL_FORM_PREV) Then
                        Begin
                              PGIBox('Attention, les populations de formation ne sont pas valides. Le calcul des frais n''est donc pas possible. #13#10'+
                                     'Veuillez effectuer les opérations nécessaires avant de poursuivre.',Ecran.Caption);
                              Exit;
                        End;
                        //PT3 - Fin
                        if MillesimeEnCours=True then AglLanceFiche('PAY','SAISIEINSCBUDGET','','','SAISIESTAGE;;'+Stage+';'+Millesime)
                        Else AglLanceFiche('PAY','SAISIEINSCBUDGET','','','SAISIESTAGE;;'+Stage+';'+Millesime+';ACTION=CONSULTATION');
                end;
                If TypeSaisie = 'CWASINSCBUDGET' Then
                begin
                        Stage := Q_MulStage.FindField ('PST_CODESTAGE').AsString;
                        Millesime := Q_MulStage.FindField ('PST_MILLESIME').AsString;
                        if MillesimeEnCours=True then AglLanceFiche('PAY','SAISIEINSCBUDGET','','','CWASINSCBUDGET;;'+Stage+';'+Millesime)
                        Else AglLanceFiche('PAY','SAISIEINSCBUDGET','','','CWASINSCBUDGET;;'+Stage+';'+Millesime+';ACTION=CONSULTATION');
                end;
        If TypeSaisie = 'BUDGETR' Then AglLanceFiche('PAY','STAGE', '',Stage+';'+Millesime , 'BUDGETR;'+ Stage+';'+THVMillesime.Value+';ACTION=CONSULTATION' );
        If TypeSaisie = 'SAISIECAT' Then
        Begin
            //PT8
            Q := OpenSQL('SELECT PST_PREDEFINI FROM STAGE WHERE PST_CODESTAGE="'+Stage+'" AND PST_MILLESIME="'+Q_MulStage.FindField ('PST_MILLESIME').AsString+'"', True);
            If Not Q.EOF Then Pred := Q.FindField('PST_PREDEFINI').AsString;
            Ferme(Q);
            If (Pred='STD') And (Not PGDroitMultiForm) Then StAction := 'CONSULTATION'
            Else StAction := 'MODIFICATION';
            AglLanceFiche('PAY','STAGE', '',Stage+';'+Millesime ,'SAISIECAT;'+ Stage+';'+THVMillesime.Value+';ACTION='+StAction);
        End;
        If TypeSaisie = 'CONSULTCAT' Then AglLanceFiche('PAY','STAGE', '',Stage+';'+Millesime ,'CONSULTCAT;'+ Stage+';'+THVMillesime.Value+';ACTION=CONSULTATION' );
        If TypeSaisie = 'SESSION' Then AglLanceFiche('PAY','STAGE', '', Stage+';'+Millesime , 'SESSION;'+Stage+';'+THVMillesime.Value+';ACTION=CONSULTATION' );
        If TypeSaisie = 'CALCSALAIRE' Then
        begin
                if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
                begin
                        MessageAlerte('Aucun élément sélectionné');
                        Exit;
                end;
                AGLLanceFiche('PAY','CALCSALAIREFORM','','','PREVISIONNEL;'+GetControltext('THVMILLESIME'));
        end;
        {$IFNDEF EMANAGER}
        Bt :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
        {$ENDIF}
        end;
end;

procedure TOF_PGMULSTAGE.ChangeMillesime(Sender : TObject);
begin
end;
{TOF_PGMULSTAGEINIT}
procedure TOF_PGMULSTAGEINIT.NouveauStage;
begin
        AglLanceFiche('PAY','DIPLOME','','','INITIALE;ACTION=CREATION');
        if Bt = NIL then Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
end;

procedure TOF_PGMULSTAGEINIT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSTAGEINIT.OnArgument (S : String ) ;
Var BOuv : TToolBarButton97;
begin
Inherited ;
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste := THGrid(GetControl('FLISTE'));
        {$ENDIF}
        if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
        BOuv  :=  TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv  <>  NIL then  BOuv.OnClick  :=  NouveauStage;
end ;

procedure TOF_PGMULSTAGEINIT.GrilleDblClick(Sender : TObject);
var Stage,Millesime : String;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Stage := Q_Mul.FindField ('PST_CODESTAGE').AsString;
        Millesime := Q_Mul.FindField ('PST_MILLESIME').AsString;
        AglLanceFiche('PAY','DIPLOME', '',Stage+';'+Millesime ,'INITIALE;'+ Stage+';;ACTION=MODIFICATION' );
        if Bt = NIL then Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
end;

{TOF_PGMULSTAGEEFOR}
procedure TOF_PGMULSTAGEEFOR.OnLoad ;
var ChoixStage :  THValComboBox;
    SQL,StWhere : String;
begin
  Inherited ;
        {$IFDEF EMANAGER}
        THVMillesime.Value := MillesimeEC;
        SetControlText('CHOIXSTAGE','SNB');
        {$ENDIF}
        StWhere := '';
        MillesimeEC  := THVMillesime.Value;
        ChoixStage := THValComboBox(GetControl('CHOIXSTAGE'));
        If ChoixStage.Value = 'SB' Then
                begin
                        SetControlText('PST_MILLESIME',MillesimeEC);
                        If TypeUtilisat = 'R' then StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE '+
                        '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'+
                        ' AND PFI_MILLESIME="'+MillesimeEC+'")'
                        else StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))) AND PFI_MILLESIME="'+MillesimeEC+'")';
                end
                Else If ChoixStage.Value = 'SNB' Then
                begin
                        SetControlText('PST_MILLESIME','0000');
                        If TypeUtilisat = 'R' then StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE NOT IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE '+
                        '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                        ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'+
                        ' AND PFI_MILLESIME="'+MillesimeEC+'")'
                        else StWhere := 'PST_ACTIF="X" AND PST_CODESTAGE NOT IN (SELECT DISTINCT (PFI_CODESTAGE) FROM INSCFORMATION WHERE PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))) AND PFI_MILLESIME="'+MillesimeEC+'")';
                end
                Else
                begin
                        SetControlText('PST_MILLESIME','');
                        SQL:= 'PST_ACTIF="X" AND ((PST_MILLESIME="'+MillesimeEC+'") OR (PST_MILLESIME="0000" AND PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'")))';
                        StWhere := SQL;
                end;
        SetControlText('XX_WHERE',StWhere);
end ;

procedure TOF_PGMULSTAGEEFOR.OnArgument (S : String ) ;
Var ChoixStage : THValComboBox;
    Num : Integer;
    Libelle,LibMillesime : THLabel;
begin
Inherited ;
        {$IFNDEF EAGLCLIENT}
        ListeEfor := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        ListeEfor := THGrid(GetControl('FLISTE'));
        {$ENDIF}
        TypeSaisie := Trim(ReadTokenSt(S));
        millesimeEC := '';
        Previsionnel := False;
        {$IFDEF EMANAGER}
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        {$ENDIF}
        MillesimeEC := RendMillesimePrevisionnel;
        Previsionnel := True;
        THVMillesime := THValComboBox(GetControl('THVMILLESIME'));
        ChoixStage := THValComboBox(GetControl('CHOIXSTAGE'));
        Libelle := THLabel(GetControl('TCHOIXSTAGE'));
        Libelle.Caption := 'Liste';
        LibMillesime := THLabel(GetControl('LTHVMILLESIME'));
        Libelle.Visible := True;
        THVMillesime.visible := True;
        THVMillesime.Value := MillesimeEC;
        LibMillesime.Visible := True;
        ChoixStage.Items.Add('<<TOUS>>');
        ChoixStage.Items.Add('formations budgetisées en '+MillesimeEC);
        ChoixStage.Items.Add('Catalogue formations');
        ChoixStage.Values.Add('TOUS');
        ChoixStage.Values.Add('SB');
        ChoixStage.Values.Add('SNB');
        ChoixStage.Value := 'SNB';
        ChoixStage.Visible := True;
        SetControlVisible('BINSERT',False);
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        UpdateCaption(TFMul(Ecran)) ;
        if ListeEfor  <>  NIL then ListeEfor.OnDblClick  :=  GrilleDblClick;
        SetControlText('DECLARATION','TOUS');
        {$IFDEF EMANAGER}
        SetControlEnabled('THVMILLESIME',False);
        SetControlVisible('PAvance',False);
        SetControlProperty('PAvance','TabVisible',False);
        SetControlEnabled('CHOIXSTAGE',False);
        {$ENDIF}
end ;

procedure TOF_PGMULSTAGEEFOR.GrilleDblClick(Sender : TObject);
var Stage,Millesime : String;
    MillesimeEnCours : Boolean;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(ListeEFor.Row-1) ;  //PT1
        {$ENDIF}
        Q_MulEfor := THQuery(Ecran.FindComponent('Q'));
        If Q_MulEfor.Eof then
        begin
                PGIBox('Vous devez sélectionner une formation',Ecran.Caption);
                Exit;
        end;
        Stage := Q_MulEfor.FindField ('PST_CODESTAGE').AsString;
        Millesime := Q_MulEfor.FindField ('PST_MILLESIME').AsString;
        MillesimeEnCours := ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+GetControlText('THVMILLESIME')+'" '+
        'AND PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-"');
        If Stage <> '' Then
        begin
                if MillesimeEnCours=True then AglLanceFiche('PAY','SAISIEINSCBUDGET','','','CWASINSCBUDGET;;'+Stage+';'+Millesime)
                Else AglLanceFiche('PAY','SAISIEINSCBUDGET','','','CWASINSCBUDGET;;'+Stage+';'+Millesime+';ACTION=CONSULTATION');
        end;
       // if Bt = NIL then Bt  :=  TToolbarButton97 (GetControl('BCherche'));
       // if Bt  <>  NIL then Bt.click;
end;

Initialization
  registerclasses ( [ TOF_PGMULSTAGE,TOF_PGMULSTAGEINIT,TOF_PGMULSTAGEEFOR ] ) ;
end.
