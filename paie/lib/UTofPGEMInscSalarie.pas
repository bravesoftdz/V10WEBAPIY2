{***********UNITE*************************************************
Auteur  ......  :
Créé le ...... : 18/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SAISIEINSCBUDGET ()
Mots clefs ... : TOF;SAISIEINSCBUDGET
*****************************************************************
PT1 24/04/2003  V_42 JL  Développement pour compatibilité CWAS
PT2 28/09/2007  V_7  FL  Adaptation cursus + accès assistant
PT3 22/10/2007  V_7  FL  Utilisation de EM_MULCATALOGUE pour l'inscription en masse + correction filtre libellé emploi
}
Unit UTofPGEMInscSalarie;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,dbtables,Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOB,UTOF,uTableFiltre,SaisieList,
     hTB97,PGOutilsFormation,EntPaie,LookUp;


Type
  TOF_PGEMSAISIESAL = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    TypeSaisie,MillesimeEC,LeSalarie,SQL,TypeUtilisat,Utilisateur : String; //PT2
    TF :  TTableFiltre;
    procedure StageElipsisClick(Sender : TObject);
    //procedure ParamTV(Sender : TObject);
    procedure ChangeFiltreEmploi (Sender : TObject);
    procedure InscriptionMultiples (Sender : TObject);
  end ;


Implementation


procedure TOF_PGEMSAISIESAL.OnLoad ;
begin
  Inherited ;
       TF.ChangeColFormat('PFI_LIEUFORM','CB=PGLIEUFORMATION');
       TFSaisieList(Ecran).BCherche.Click;
       If Not ExisteSQL('SELECT PFI_SALARIE FROM INSCFORMATION '+SQL) Then
       begin
           // Bt := TToolBarButton97(GetControl('BInsert'));
           // Bt.Click;
           //TF.State := DsInsert;
           TF.Insert;
       end;
       
end ;

procedure TOF_PGEMSAISIESAL.OnArgument (S : String ) ;
var Q : TQuery;
    Millesime : String;
    Edit : THEdit;
    Combo : THValComboBox;
    Stage,TypeInsc,LibEmploi,SQL : String; //PT2
    Bt : TToolbarButton97;
begin
  Inherited ;
       TypeSaisie := Trim(ReadTokenPipe(S, ';'));
       TypeInsc := Trim(ReadTokenPipe(S, ';'));
       Stage := Trim(ReadTokenPipe(S, ';'));
       Millesime := Trim(ReadTokenPipe(S, ';'));
       LeSalarie := Trim(ReadTokenPipe(S, ';'));
       LibEmploi := Trim(ReadTokenPipe(S, ';'));
       {$IFDEF EMANAGER}
       ReadTokenSt(S);
       Utilisateur := ReadTokenSt(S);
       {$ENDIF}
       SetControlText('PSA_SALARIE',LeSalarie);

       //PT2 - Début
       If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+Utilisateur+'"') then TypeUtilisat := 'R'
       else TypeUtilisat := 'S';
       //PT2 - Fin

       Q := OpenSQL('SELECT PSA_LIBELLEEMPLOI,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_DATENAISSANCE,PSA_DATEENTREE FROM SALARIES WHERE PSA_SALARIE="'+LeSalarie+'"',True);
       if Not Q.Eof then
       begin
            SetControlText('LESALARIE',Q.FindField('PSA_LIBELLE').AsString+ ' '+Q.FindField('PSA_PRENOM').AsString);
            SetControlText('LIBEMPLOI',Q.FindField('PSA_LIBELLEEMPLOI').AsString);
            SetControlText('ETABLISSEMENT',Q.FindField('PSA_ETABLISSEMENT').AsString);
            SetControlText('NAISSANCE',Q.FindField('PSA_DATENAISSANCE').AsString);
            SetControlText('ENTREE',Q.FindField('PSA_DATEENTREE').AsString);
       end;
       Ferme(Q);
       MillesimeEC := RendMillesimeEManager;
       If ExisteSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+MillesimeEC+'" AND PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'"') then SetControlChecked('FAVORIS',True);
       If (TypeSaisie = 'SAISIESTAGE') or (TypeSaisie = 'CWASINSCBUDGET')  Then
       begin
                SetControlVisible('BTree',False);
                SetControlVisible('PANTREEVIEW',False);
       end;
       TF  :=  TFSaisieList(Ecran).LeFiltre;
       Edit := THEdit(GetControl('PFI_CODESTAGE'));
       If Edit <> Nil Then Edit.OnelipsisClick := StageElipsisClick;
       //PT2 - Début
       SQL := 'WHERE PFI_SALARIE="'+LeSalarie+'" AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CODESTAGE<>"--CURSUS--"'; //PT2
       SQL := SQL + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,False);
       //PT2 - Fin
       {$IFDEF EMANAGER}
       If LibEmploi <> '' then
       begin
            SetControlText('FILTREEMPLOI',LibEmploi);
            //PT2 - Début
            SQL := 'WHERE PFI_LIBEMPLOIFOR="'+LibEmploi+'" AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CODESTAGE<>"--CURSUS--"'; //PT2
            SQL := SQL + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,False)
            //PT2 - Fin
       end;
       {$ENDIF}
       TF.WhereTable := SQL;

       {$IFNDEF EMANAGER}SetControlVisible('BPARAMLISTE',True);{$ENDIF}
       SetActiveTabSheet('PSALARIE');
       SetControlProperty('PFI_NBINSC','MinValue',1);
       SetControlProperty('PFI_NBINSC','MaxValue',10000);
       SetControlEnabled('PFI_JOURSTAGE',False);
       TFSaisieList(Ecran).ParamTreeView := False;
       If LeSalarie = '' then
       Begin
            TFSaisieList(Ecran).caption := 'Inscriptions non nominatives au prévisionnel '+MillesimeEC;
            SetControlVisible('PanPied', False); //PT2
       End
       else
            TFSaisieList(Ecran).caption := 'Inscriptions du salarié '+GetControlText('LESALARIE')+' au prévisionnel '+MillesimeEC;
       UpdateCaption(Ecran);

       //PT2 - Début
       SetControlProperty('PFI_LIBEMPLOIFOR','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
                        'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND '+
                        AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,True)+')');
       SetControlProperty('FILTREEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
                        'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND '+
                        AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,True)+')');
       //PT2 - Fin
       Combo := THValComboBox(GetControl('FILTREEMPLOI'));
       If Combo <> Nil then Combo.OnChange := ChangeFiltreEmploi;
       Bt := TToolbarButton97(GetControl('BMULTI'));
       If Bt <> Nil then Bt.onClick := InscriptionMultiples;

end ;                                                     

procedure TOF_PGEMSAISIESAL.StageElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
       { StWhere  :=  'PST_ACTIF="X" AND'+
        ' ((PST_MILLESIME="'+MillesimeEC+'") OR (PST_MILLESIME="0000" AND '+
        'PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'")))';
        If GetCheckBoxState('FAVORIS') = CbChecked then StWhere := StWhere + ' AND PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_LIBELLE="'+V_PGI.UserSalarie+'")';
        StOrder  :=  'PST_MILLESIME,PST_LIBELLE';
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,StOrder, True,-1);}
        StWhere := 'PST_ACTIF="X" AND CO_TYPE="PFN" AND '+  //PT2
        ' PST_MILLESIME="0000" AND PST_NATUREFORM<>"004"';
        If GetCheckBoxState('FAVORIS') = CbChecked then StWhere := StWhere + ' AND PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")';
        StWhere := StWhere + ' AND PST_FORMATION3="'+CYCLE_EN_COURS_EM+'"'; 
        StOrder := 'PST_LIBELLE,PST_MILLESIME';
        {$IFNDEF EAGLCLIENT}
        LookupList(THDBEdit(Sender),'Liste des stages','STAGE LEFT JOIN COMMUN ON PST_NATUREFORM=CO_CODE','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,CO_LIBELLE',StWhere,StOrder, True,-1);
        {$ELSE}
        LookupList(THEdit(Sender),'Liste des stages','STAGE LEFT JOIN COMMUN ON PST_NATUREFORM=CO_CODE','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,CO_LIBELLE',StWhere,StOrder, True,-1);
        {$ENDIF}
end;

{procedure TOF_PGEMSAISIESAL.ParamTV(Sender : TObject);
begin
        TFSaisieList(Ecran).BChercheClick(Sender);
end;}

procedure TOF_PGEMSAISIESAL.ChangeFiltreEmploi (Sender : TObject);
var LibEmploi : String; //pt3
begin
        if TF.State = DsEdit then
        begin
                PGIBox('Changement impossible car vous êtes en modification', Ecran.Caption);
                Exit;
        end;
        if TF.State = DsInsert then
        begin
                PGIBox('Changement impossible car vous êtes en création', Ecran.Caption);
                Exit;
        end;
       //PT2 - Début
       LibEmploi := GetControlText('FILTREEMPLOI'); //pt3
       If LibEmploi <> '' then
            SQL := 'WHERE PFI_LIBEMPLOIFOR="'+LibEmploi+'" AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CODESTAGE<>"--CURSUS--"'
       Else
            SQL := 'WHERE PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CODESTAGE<>"--CURSUS--"'; //pt3
            //If GetControltext('FILTREEMPLOI') <> '' then SQL := 'WHERE PFI_SALARIE="'+LeSalarie+'" AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_LIBEMPLOIFOR="'+GetControlText('FILTREEMPLOI')+'" AND PFI_CODESTAGE<>"--CURSUS--"' //PT2 //pt3
            //else SQL := 'WHERE PFI_SALARIE="'+LeSalarie+'" AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CODESTAGE<>"--CURSUS--"'; //PT2
       SQL :=SQL + ' AND '+AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,False);
       //PT2 - Fin
        TF.WhereTable:= SQL;
        TF.RefreshLignes;
end;

procedure TOF_PGEMSAISIESAL.InscriptionMultiples (Sender : TObject);
var Ret : String;
    t,TobStage,TobInvest,TI : Tob;
    Salaire : Double;
    Q,QAnim : TQuery;
    For1,For2,For3,For4,For5,For6,For7,For8 : String;
    Nature,Lieu : String;
    JourStage,Duree,FraisH,FraisR,FraisT,TotalFrais : Double;
    TauxChargeBudget,NbJours : Double;
    SalaireAnim,NbAnim : Double;
    LibreCmb1,LibreCmb2,LibreCmb3,LibreCmb4,TravailN1,TravailN2,TravailN3,TravailN4 : String;
    CodeStat,Etab,Emploi,DadsCat,Stage : String;
begin
    Ret := AGLLanceFiche('PAY','EM_MULCATALOGUE','','','SAISIEMASSE'); //PT3
    Q := OpenSQL('SELECT PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="' + MillesimeEC + '"', True);
    if not Q.Eof then TauxChargeBudget := Q.FindField('PFE_TAUXBUDGET').AsFloat
    else TauxChargeBudget := 1;
    Ferme(Q);
    Q := OpenSQL('SELECT PSA_DADSCAT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,' +
     'PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI ' +
    'FROM SALARIES WHERE PSA_SALARIE="'+LeSalarie+'"',True);
    If Not Q.Eof then
    begin
         Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
         Emploi := Q.FindField('PSA_LIBELLEEMPLOI').AsString;
         TravailN1 := Q.FindField('PSA_TRAVAILN1').AsString;
         TravailN2 := Q.FindField('PSA_TRAVAILN2').AsString;
         TravailN3 := Q.FindField('PSA_TRAVAILN3').AsString;
         TravailN4 := Q.FindField('PSA_TRAVAILN4').AsString;
         CodeStat := Q.FindField('PSA_CODESTAT').AsString;
         LibreCmb1 := Q.FindField('PSA_LIBREPCMB1').AsString;
         LibreCmb2 := Q.FindField('PSA_LIBREPCMB2').AsString;
         LibreCmb3 := Q.FindField('PSA_LIBREPCMB3').AsString;
         LibreCmb4 := Q.FindField('PSA_LIBREPCMB4').AsString;
         DadsCat := Q.FindField('PSA_DADSCAT').AsString;
    end;
    Ferme(Q);
    If Ret = '' then exit;
    TF.DisableTom;
    TF.StartUpdate;
    Stage := ReadTokenPipe(Ret,';');
    While Stage <> '' do
    begin
         Q := OpenSQL('SELECT PST_JOURSTAGE,PST_LIEUFORM,PST_DUREESTAGE,PST_NBANIM,' +
         'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_NATUREFORM' +
         ' FROM STAGE WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="0000"', True);
         if not Q.eof then
         begin
              JourStage := Q.FindField('PST_JOURSTAGE').AsFloat;
              Duree := Q.FindField('PST_DUREESTAGE').AsFloat;
              Lieu := Q.FindField('PST_LIEUFORM').AsString;
              For1 := Q.FindField('PST_FORMATION1').AsString;
              For2 := Q.FindField('PST_FORMATION2').AsString;
              For3 := Q.FindField('PST_FORMATION3').AsString;
              For4 := Q.FindField('PST_FORMATION4').AsString;
              For5 := Q.FindField('PST_FORMATION5').AsString;
              For6 := Q.FindField('PST_FORMATION6').AsString;
              For7 := Q.FindField('PST_FORMATION7').AsString;
              For8 := Q.FindField('PST_FORMATION8').AsString;
              Nature := Q.FindField('PST_NATUREFORM').AsString;
              NbAnim := Q.FindField('PST_NBANIM').AsFloat;
         end;
         Ferme(Q);
         if not ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + MillesimeEC + '"') then
         begin
              QAnim := OpenSQl('SELECT PFE_SALAIREANIM FROM EXERFORMATION WHERE PFE_MILLESIME="' + MillesimeEC + '"', True);
              salaireanim := 0;
              if not QAnim.eof then SalaireAnim := QAnim.FindField('PFE_SALAIREANIM').AsFloat;
              Ferme(QAnim);
              Q := OpenSQL('SELECT * FROM STAGE WHERE ' +
              'PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="0000"', True);
              TobStage := Tob.Create('STAGE', nil, -1);
              TobStage.loadDetailDB('STAGE', '', '', Q, False);
              Ferme(Q);
              T := TobStage.FindFirst(['PST_CODESTAGE'], [Stage], False);
              if T <> nil then
              begin
                   T.ChangeParent(TobStage, -1);
                   T.PutValue('PST_MILLESIME', MillesimeEC);
                   T.PutValue('PST_COUTSALAIR', SalaireAnim * Duree * NbAnim);
                   T.InsertOrUpdateDB;
              end;
              TobStage.Free;
              // Duplication investissement pour la formation
              Q := OpenSQL('SELECT * FROM INVESTSESSION ' +
              'WHERE PIS_CODESTAGE="' + Stage + '" AND PIS_ORDRE=-1 AND PIS_MILLESIME="0000"', True); //DB2
              TobInvest := Tob.Create('INVESTSESSION', nil, -1);
              TobInvest.LoadDetailDB('INVESTSESSION', '', '', Q, False);
              Ferme(Q);
              TI := TobInvest.FindFirst(['PIS_CODESTAGE'], [Stage], False);
              while TI <> nil do
              begin
                   TI.ChangeParent(TobInvest, -1);
                   TI.PutValue('PIS_MILLESIME', GetField('PFI_MILLESIME'));
                   TI.InsertOrUpdateDB(False);
                   TI := TobInvest.FindNext(['PIS_CODESTAGE'], [Stage], False);
              end;
         TobInvest.Free;
         end;
         TF.Insert;
         TF.PutValue('PFI_SALARIE',LeSalarie);
         TF.PutValue('PFI_LIBELLE',Rechdom('PGSALARIE',LeSalarie,False));
         TF.PutValue('PFI_CODESTAGE',Stage);
         TF.PutValue('PFI_RESPONSFOR',Utilisateur);
         TF.PutValue('PFI_MILLESIME',MillesimeEC);
         TF.PutValue('PFI_NBINSC',1);
         TF.PutValue('PFI_ETABLISSEMENT',Etab);
         TF.PutValue('PFI_LIBEMPLOIFOR',Emploi);
         TF.PutValue('PFI_TRAVAILN1',TravailN1);
         TF.PutValue('PFI_TRAVAILN2',TravailN2);
         TF.PutValue('PFI_TRAVAILN3',TravailN3);
         TF.PutValue('PFI_TRAVAILN4',TravailN4);
         TF.PutValue('PFI_CODESTAT',CodeStat);
         TF.PutValue('PFI_LIBREPCMB1', LibreCmb1);
         TF.PutValue('PFI_LIBREPCMB2', LibreCmb2);
         TF.PutValue('PFI_LIBREPCMB3', LibreCmb3);
         TF.PutValue('PFI_LIBREPCMB4', LibreCmb4);
         TF.PutValue('PFI_DADSCAT', DadsCat);
         Q := OpenSQL('SELECT MAX (PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_ETABLISSEMENT="' + TF.GetValue('PFI_ETABLISSEMENT') + '" AND PFI_MILLESIME="' + MillesimeEC + '"', True);
         if not Q.eof then TF.PutValue('PFI_RANG', Q.FindField('RANG').AsInteger + 1)
         else TF.PutValue('PFI_RANG', 1);
         Ferme(Q);
         TF.PutValue('PFI_ETATINSCFOR','ATT');
         TF.PutValue('PFI_TYPEPLANPREV','PLF');
         TF.PutValue('PFI_JOURSTAGE', JourStage);
         TF.PutValue('PFI_DUREESTAGE', Duree);
         TF.PutValue('PFI_HTPSTRAV', Duree);
         TF.PutValue('PFI_HTPSNONTRAV', 0);
         TF.PutValue('PFI_LIEUFORM', Lieu);
         TF.PutValue('PFI_FORMATION1', For1);
         TF.PutValue('PFI_FORMATION2', For2);
         TF.PutValue('PFI_FORMATION3', For3);
         TF.PutValue('PFI_FORMATION4', For4);
         TF.PutValue('PFI_FORMATION5', For5);
         TF.PutValue('PFI_FORMATION6', For6);
         TF.PutValue('PFI_FORMATION7', For7);
         TF.PutValue('PFI_FORMATION8', For8);
         TF.PutValue('PFI_NATUREFORM', Nature);
         TF.PutValue('PFI_NIVPRIORITE', '01');
         TF.PutValue('PFI_MOTIFINSCFOR', '2');
         if VH_Paie.PGForValoSalaire = 'VCR' then Salaire := ForTauxHoraireReel(TF.Getvalue('PFI_SALARIE'))
         else Salaire := ForTauxHoraireCategoriel(TF.GetValue('PFI_LIBEMPLOIFOR'), MillesimeEC);
         TF.PutValue('PFI_COUTREELSAL', Salaire * Duree * TauxChargeBudget);
         Q := OpenSQL('SELECT PFF_FRAISHEBERG,PFF_FRAISREPAS,PFF_FRAISTRANSP FROM FORFAITFORM ' +
          'WHERE PFF_MILLESIME="' + MillesimeEC + '" AND PFF_LIEUFORM="' + Lieu + '"' + //PT3
          ' AND PFF_ETABLISSEMENT="' + TF.Getvalue('PFI_ETABLISSEMENT') + '"', True);
          FraisH := 0;
          FraisR := 0;
          FraisT := 0;
          if not Q.eof then
          begin
            FraisH := Q.FindField('PFF_FRAISHEBERG').AsFloat;
            FraisR := Q.FindField('PFF_FRAISREPAS').AsFloat;
            FraisT := Q.FindField('PFF_FRAISTRANSP').AsFloat;
          end;
          Ferme(Q);
          NbJours := JourStage;
          FraisH := FraisH * (NbJours - 1);
          if FraisH < 0 then FraisH := 0;
          FraisR := FraisR * ((NbJours * 2) - 1);
          if FraisR < 0 then FraisR := 0;
          TotalFrais := FraisH + FraisR + FraisT;
          TF.PutValue('PFI_FRAISFORFAIT', TotalFrais);
          TF.Post;
          {$IFNDEF EMANAGER}
          MajCoutPrev(Stage,MillesimeEC,TF);
          {$ENDIF}
          Stage := ReadTokenPipe(Ret,';');
    end;
    TF.EndUpdate;
    TF.EnableTom;
end;

Initialization
  registerclasses ( [ TOF_PGEMSAISIESAL ] ) ;
end.

