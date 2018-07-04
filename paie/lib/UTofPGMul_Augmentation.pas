{***********UNITE*************************************************
 Auteur  ...... : JL
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMUL_AUGMENTATION ()
Mots clefs ... : TOF;saisie des augmentations de salaires
*****************************************************************}
Unit UTofPGMul_Augmentation ;

Interface

Uses StdCtrls,Controls,Classes, 
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,HDB,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HQry,PgOutils2,P5Def,HTB97,EntPaie,uTob,HSysMenu,Ed_Tools ,P5Util,LookUp,ParamSoc;

Type
  TOF_PGMUL_AUGMENTATION = Class (TOF)
    procedure OnClose                   ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Argument : String;
    NbMEquivalenceAnn : Integer;
      Q_MulAug : THQuery;
    {$IFNDEF EAGLCLIENT}
        ListeAug : THDBGrid ;
        {$ELSE}
        ListeAug : THGrid ;
        {$ENDIF}
    Consultation : Boolean;
    RequeteSalAugm : String;
    procedure GrilleDblClick (Sender : TObject);
    procedure ExitEdit(Sender: TObject);
    procedure CreerLigneSalarie;
    procedure ChargerTobAug(AnneeSaisie : String);
    Function RendSalairesTob(TypeSalaire : String;TS : Tob) : Double;
    Function SalarieExclus (TS,TC : Tob;Fixe,Variable : Double) : String;
    Function Comparaison(Operateur : String;Val1,Val2 : Variant) : Boolean;
    procedure SuppAugmentations(Sender : Tobject);
    procedure AfficheLegende(Sender : TObject);
    Function RendRequeteResp(ValeurResp,ChampResp : String) : String;
    procedure SalarieElipsisClick(Sender : TObject);
    procedure RespAbsElipsisClick(Sender : TObject);
//    Procedure Valider(Sender : Tobject);
//    Procedure Refuser(Sender : Tobject);
  end ;
  var
 //   Q_MulAug : THQuery;
    {$IFNDEF EAGLCLIENT}
//        ListeAug : THDBGrid ;
        {$ELSE}
//        ListeAug : THGrid ;
        {$ENDIF}
        TobAug : Tob;

Implementation

procedure TOF_PGMUL_AUGMENTATION.OnClose ;
begin
  Inherited ;
 // ListeAug := Nil;
end;

procedure TOF_PGMUL_AUGMENTATION.OnLoad ;
var StWhere,Annee,WhereResp : String;
    CRespVar : THMultiValComboBox;
begin
  Inherited ;
        StWhere := '';
        {$IFDEF EMANAGER}
        If Argument = 'VALIDRESP' then
         begin
                StWhere := '(PBG_RESPONSVAR="'+V_PGI.UserSalarie+'" OR PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICES ON PSE_CODESERVICE=PGS_CODESERVICE LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))';
         end
         else
         begin
                If Consultation then StWhere := '(PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSVAR<>"'+V_PGI.UserSalarie+'" AND (PSE_RESPONSABS="'+V_PGI.UserSalarie+'" OR PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSABS="'+V_PGI.UserSalarie+'")))))'
                else StWhere := '(PBG_RESPONSVAR="'+V_PGI.UserSalarie+'" OR PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_ASSISTNDF="'+V_PGI.UserSalarie+'"))';

         end;

        {$ENDIF}
        If GetControlText('ANNEE') <> '' then
        begin
             Annee := RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False);
             If StWhere = '' then StWhere := '(PBG_ANNEE="'+Annee+'")'
             else StWhere := StWhere + 'AND (PBG_ANNEE="'+Annee+'")';
        end;
        CRespVar := THMultiValComboBox(GetControl('RESPONSVAR'));
        If (GetControlText('RESPONSVAR') <> '') and (CRespVar.tous = False) then
        begin
             WhereResp := RendRequeteResp(GetControlText('RESPONSVAR'),'PSE_RESPONSVAR');
             If StWhere = '' then StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+')'
             else  StWhere := StWhere + 'AND PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+')';
        end;
        If (GetControlText('RESPONSABS') <> '') then
        begin
             If StWhere = '' then StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSABS="'+GetControlText('RESPONSABS')+'")'
             else  StWhere := StWhere + 'AND PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSABS="'+GetControlText('RESPONSABS')+'")';
        end;
        If StWhere <> '' then stWhere := StWhere + ' AND PSA_SORTIEDEFINIT<>"X"'
        else stWhere := 'PSA_SORTIEDEFINIT<>"X"';
        SetControlText('XX_WHERE',StWHere);
end ;

procedure TOF_PGMUL_AUGMENTATION.OnArgument (S : String ) ;
var Num : Integer;
    Defaut : THEdit;
    BSupp,BLegende : TToolBarButton97;
    MultiC : THMultiValComboBox;
begin
  Inherited ;
        NbMEquivalenceAnn := GetParamSocSecur('SO_PGAUGMNBMOISANN',False);
        If NbMEquivalenceAnn <= 0 then NbMEquivalenceAnn := 12;    
        Argument := ReadTokenPipe(S,';');
        Consultation := False;
        If Argument = 'PROPOSITION' then
        begin
             Consultation := True;
             TFMul(Ecran).Caption := 'Propositions d''augmentations';
             Argument := 'SAISIE';
        end;
        MultiC := THMultiValComboBox(GetControl('PBG_VALIDERPAR'));
        If MultiC <> nil then MultiC.plus := ' AND PSI_INTERIMAIRE IN '+
        '(SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE WHERE PHO_NIVEAUH<=2)';
        Q_MulAug := THQuery(Ecran.FindComponent('Q'));
        SetControlCaption('LIBSALARIE','');
        SetControlCaption('LIBABSENCE','');
         {$IFNDEF EAGLCLIENT}
        ListeAug := THDBGrid(GetControl('FListe'));
        {$ELSE}
        ListeAug := THGrid(GetControl('FListe'));
        {$ENDIF}
        SetControlText('ANNEE',GetParamSoc('SO_PGAUGANNEE'));

        Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
        If Defaut<>nil then
        begin
             Defaut.OnExit:=ExitEdit;
             defaut.OnElipsisClick := SalarieElipsisClick;
        end;
        Defaut:=ThEdit(getcontrol('RESPONSABS'));
        If Defaut<>nil then
        begin
             Defaut.OnExit:=ExitEdit;
             defaut.OnElipsisClick := RespAbsElipsisClick;
        end;
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
         end;
         If ListeAug <> Nil Then ListeAug.OnDblClick := GrilleDblClick ;
         VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
         If Argument = 'IMPORTSALARIE' then
         begin
              Ecran.Caption := 'Import des salariés pour augmentation';
              SetControlProperty('BOuvrir','Hint','Importer les salariés');
         end;
         If Argument = 'VALIDDRH' then
         begin
              Ecran.Caption := 'Validation des augmentations';
         end;
         If Argument = 'VALIDRESP' then
         begin
                TFMul(Ecran).SetDBListe('PGAUGMRESPVALID');
             //   if Q_MulAug <> nil then Q_MulAug.Liste := 'PGAUGMRESPVALID';
                Ecran.Caption := 'Validation des augmentations';
                {$IFDEF EMANAGER}
                SetControlProperty('RESPONSABLE','Plus',' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))');
                SetControlVisible('PBG_VALIDERPAR',False);
                SetControlVisible('TPBG_VALIDERPAR',False);
                {$ENDIF}
         end;
         If Argument = 'INTEGRATION' then
         begin
              Ecran.Caption := 'Intégration des augmentations';
              SetControlProperty('BOuvrir','Hint','Intégrer les augmentations');
         end;
         If Argument = 'SAISIE' then
         begin
              {$IFDEF EMANAGER}
//              SetControlenabled('RESPONSVAR',False);
//              SetControlVisible('RESPONSVAR',False);
              SetControlVisible('PBG_VALIDERPAR',False);
//              SetControlVisible('TRESPONSVAR',False);
              SetControlVisible('TPBG_VALIDERPAR',False);
{              If Consultation then
              begin
//                   SetControlText('RESPONSABS',V_PGI.UserSalarie);
//                   SetControlVisible('TRESPONSABS',False);
//                   SetControlVisible('RESPONSABS',False);
              end
              else SetControlText('RESPONSVAR',V_PGI.UserSalarie);}
              {$ENDIF}
         end;
         If Argument = 'SUPPRESSION' then
         begin
              SetControlVisible('BOuvrir',False);
              Ecran.Caption := 'Suppression des augmentations';
              SetControlProperty('BOuvrir','Hint','Intégrer les augmentations');
              SetControlVisible('BDelete',True);
              BSupp := TToolBarButton97(GetControl('BDelete'));
              If BSupp <> Nil then BSupp.OnClick := SuppAugmentations;
         end
         else SetControlVisible('BDelete',false);
         UpdateCaption(TFMul(Ecran));
         BLegende := TToolBarButton97(GetControl('BLEGENDE'));
         If BLegende <> Nil then BLegende.OnClick := AfficheLegende;
         {$IFDEF EMANAGER}
         SetControlEnabled('ANNEE',False);
         MultiC := THMultiValComboBox(GetControl('RESPONSVAR'));
         MultiC.plus :=  ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))';
         if ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE="'+V_PGI.UserSalarie+'" AND PSE_ADMINISTABS="X"') then
         begin
              SetControlEnabled('BParamListe',true);
              SetControlVisible('BParamListe',true);
              SetControlEnabled('BFiltre',True);
         end
         else
         begin
              SetControlEnabled('BParamListe',False);
              SetControlEnabled('BFiltre',False);
         end;
         {$ENDIF}

end ;

procedure TOF_PGMUL_AUGMENTATION.GrilleDblClick (Sender : TObject);
var Retour,Annee : String;
    Salarie : String;
    i : Integer;
begin
        If TFMul(Ecran).bSelectAll.Down then ListeAug.AllSelected := True;
        If Argument = 'INTEGRATION' then
        begin
                if (ListeAug = nil) then Exit;
                if (ListeAug.NbSelected = 0) and (not ListeAug.AllSelected) then
                begin
                        MessageAlerte('Aucun élément sélectionné');
                        exit;
                end;
                AGLLanceFiche('PAY','AUGM_INTEGRA','','','');
        end;
        If (Argument = 'SAISIE') or (Argument = 'VALIDRESP') or (Argument = 'VALIDDRH') then
        begin
                If Consultation then Argument := 'PROPOSITION';
                if (ListeAug = nil) then Exit;
                If Q_MulAug.RecordCount < 1 then
                begin
                     PGIBox('Aucun salarié ne correspond aux critères',ecran.Caption);
                     Exit;
                end;
                {$IFNDEF EMANAGER}
                if (ListeAug.NbSelected = 0) and (not ListeAug.AllSelected) then
                begin
                        MessageAlerte('Vous devez sélectionner le(s) salarié(s) à saisir');
                        exit;
                end;
                {$ENDIF}
                If GetControlText('ANNEE') = '' then
                begin
                     PGIBox('Vous devez renseigner l''année',ecran.Caption);
                     Exit;
                end;
                RequeteSalAugm := '';
                Annee := RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False);
                If (ListeAug.AllSelected) or (ListeAug.NbSelected = 0) then
                begin
                     {$IFDEF EAGLCLIENT}
                     if (TFMul(Ecran).bSelectAll.Down) or (ListeAug.NbSelected = 0) then
                     TFMul(Ecran).Fetchlestous;
                     {$ENDIF}
                     TFmul(Ecran).Q.First;
                     while not TFmul(Ecran).Q.EOF do
                     begin
                       Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
                       RequeteSalAugm := RequeteSalAugm + ' "' + Salarie + '",';
                       TFmul(Ecran).Q.Next;
                     end;
                end
                else
                begin

                     for i := 0 to ListeAug.NbSelected-1 do
                     begin
                          ListeAug.GotoLeBookmark(i);
                          {$IFDEF EAGLCLIENT}
                          TFmul(Ecran).Q.TQ.Seek(ListeAug.Row-1) ;
                          {$ENDIF}
                          Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
                          RequeteSalAugm := RequeteSalAugm + ' "' + Salarie + '",';
                     end;
                end;

                if RequeteSalAugm <> '' then RequeteSalAugm := ' PSA_SALARIE IN (' + Copy(RequeteSalAugm, 1, Length(RequeteSalAugm) - 1) + ')';
                ChargerTobAug(Annee);
                If TobAug.Detail.Count <=0 then
                begin
                     PGIBox('Aucun salarié ne correspond aux critères d''augmentation');
                     Retour := '';
                end
                else Retour := AGLLanceFiche('PAY','AUGMENTATIONSAL','','',Argument+';'+Annee);
                TobAug.Free;
                SetControlProperty('BSelectAll','Down',False);
                ListeAug.AllSelected := False;
                ListeAug.ClearSelected;
                If Retour = 'MODIF' then TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
                If Argument = 'PROPOSITION' then Argument := 'SAISIE';
        end;
        If Argument = 'IMPORTSALARIE' then CreerLigneSalarie;
end;

procedure TOF_PGMUL_AUGMENTATION.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

Procedure TOF_PGMUL_AUGMENTATION.CreerLigneSalarie;
var TobSal,TobAug,TA,TCritere : Tob;
    i : Integer;
    Q : TQuery;
    Fixe,Variable : Double;
    StWhere : String;
begin
   StWhere := '';
   If GetControlText('ANNEE') = '' then
   begin
       PGIBox('Vous devez choisir une année',Ecran.Caption);
   end;
   StWHere := 'WHERE PSA_ETABLISSEMENT="'+GetControlText('PSA_ETABLISSEMENT')+'"';  
   If StWhere <> '' then StWhere := StWhere + ' AND (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+UsDateTime(IDate1900)+'")'
   Else StWhere := 'WHERE PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+UsDateTime(IDate1900)+'"';
   Q := OpenSQL('SELECT * FROM SALARIES '+StWhere,True);
   TobSal := Tob.Create('LesSalaris',Nil,-1);
   TobSal.LoadDetailDB('LesSalaries','','',Q,False);
   Ferme(Q);
   Q := OpenSQL('SELECT * FROM BUDGETPAIE WHERE PBG_ANNEE="'+RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False)+'"',True);
   TobAug :=  Tob.Create('BUDGETPAIE',Nil,-1);
   TobAug.LoadDetailDB('BUDGETPAIE','','',Q,False);
   ferme(Q);
   Q := OpenSQL('SELECT * FROM AUGMPARAM WHERE PAP_ACTIF="X"',True);
   TCritere :=  Tob.Create('Lesexclusions',Nil,-1);
   TCritere.LoadDetailDB('Lesexclusions','','',Q,False);
   ferme(Q);
   {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
         {$ENDIF}
   InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', TobSal.Detail.Count,
                    False,True);
   For i := 0 to TobSal.Detail.Count - 1 do
   begin
        TA := TobAug.FindFirst(['PBG_SALARIE'],[TobSal.Detail[i].getValue('PSA_SALARIE')],False);
        If TA = Nil then
        begin
             TA := Tob.Create('BUDGETPAIE',TobAug,-1);
             TA.PutValue('PBG_SALARIE',TobSal.Detail[i].getValue('PSA_SALARIE'));
             TA.PutValue('PBG_TYPEBUDG','AUG');
             TA.PutValue('PBG_NUMORDRE',1);
  //                 TA.PutValue('PBG_DATEBUDG'
             TA.PutValue('PBG_ANNEE',RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False));
             Fixe := RendSalairesTob('FIXE',TobSal.Detail[i]);
             Variable := RendSalairesTob('VARIABLE',TobSal.Detail[i]);
             TA.PutValue('PBG_FIXEAV',Fixe);
             TA.PutValue('PBG_VARIABLEAV',Variable);
             TA.PutValue('PBG_FIXEAP',Fixe);
             TA.PutValue('PBG_VARIABLEAP',Variable);
             TA.PutValue('PBG_ETATINTAUGM',SalarieExclus (TobSal.Detail[i],TCritere,Fixe,Variable));
             {PBG_ETABLISSEMENT
             PBG_LIBELLE
             PBG_TRAVAILN1
             PBG_TRAVAILN2
             PBG_TRAVAILN3
             PBG_TRAVAILN4
             PBG_CODESTAT
             PBG_LIBREPCMB1
             PBG_LIBREPCMB2
             PBG_LIBREPCMB3
             PBG_LIBREPCMB4
             PBG_CONDEMPLOI
             PBG_DADSPROF
             PBG_LIBELLEEMPLOI
             PBG_MOTIFCTINT
             PBG_DEBUTEMPLOI
             PBG_FINEMPLOI
             PBG_SALAIREMOIS1
             PBG_SALAIREMOIS2
             PBG_SALAIREMOIS3
             PBG_SALAIREMOIS4
             PBG_SALAIREMOIS5
             PBG_SALAIRANN1
             PBG_SALAIRANN2
             PBG_SALAIRANN3
             PBG_SALAIRANN4
             PBG_SALAIRANN5
             PBG_HORAIREMOIS
             PBG_UNITEPRISEFF
             PBG_BUDGETAPE
             PBG_TYPECONTRAT
             PBG_DADSEXOBASE
             PBG_MOTIFAUGM

             PBG_OFFREEMPLOI
             PBG_TYPOFFREEMPLOI
             PBG_SUPPORTOFFRE
             PBG_FIXEAP
             PBG_VARIABLEAP
             PBG_PCTFIXE
             PBG_PCTVARIABLE
             PBG_RESPONSVAR
             PBG_VALIDERPAR
             PBG_VALIDELE
             PBG_ACCEPTERPAFR
             PBG_ACCEPTELE
             PBG_OFFREPOURVUE
             PBG_CONFIDENTIEL
             PBG_COMMENTAIREABR
             PBG_COMMENTAIRE  }
             TA.InsertDB(Nil);
        end;
        MoveCurProgressForm ('Salarié : '+
                       TobSal.Detail[i].GetValue('PSA_LIBELLE'));
   end;
   TobAug.Free;
   TobSal.Free;
   TCritere.Free;
   FiniMoveProgressForm;
end;

procedure TOF_PGMUL_AUGMENTATION.ChargerTobAug(AnneeSaisie : String);
var TobSal,TSA,TobMul,TA : Tob;
    i : Integer;
    HoraireMens,Effectif : Double;
    Q : TQuery;
    Salarie,Etat : String;
    PremMois, PremAnnee,Age : WORD;
    Anciennete,AncAnnee,AncMois : Integer;
    StAge,StAnciennete : String;
    DateNaiss,DateANC : TDateTime;
    LibStat,LibOrg1,LibOrg2,LibOrg3,LibOrg4 : String;
    AncBrut,NewBrut,PctBrut,coeff : Double;
    PctAugmDec : Integer;
begin
        PctAugmDec := GetParamSoc('SO_PGAUGMPCTDEC');
        If VH_Paie.PGNbreStatOrg > 0 then LibOrg1 := VH_Paie.PGLibelleOrgStat1 + ' : '
        else LibOrg1 := '';
        If VH_Paie.PGNbreStatOrg > 1 then LibOrg2 := VH_Paie.PGLibelleOrgStat2 + ' : '
        else LibOrg2 := '';
        If VH_Paie.PGNbreStatOrg > 2 then LibOrg3 := VH_Paie.PGLibelleOrgStat3 + ' : '
        else LibOrg3 := '';
        If VH_Paie.PGNbreStatOrg > 3 then LibOrg4 := VH_Paie.PGLibelleOrgStat4 + ' : '
        else LibOrg4 := '';
        LibStat := VH_Paie.PGLibCodeStat;
        TobAug := Tob.Create('augmentation',Nil,-1);
        Q := OpenSQL('SELECT PSA_LIBELLE,PSA_SALARIE,PSA_PRENOM,PSA_LIBELLEEMPLOI,PSA_DATEANCIENNETE,PSA_REGULANCIEN,PSA_DATEENTREE,PSA_CONVENTION'+
                  ',PSA_DATENAISSANCE,PSA_UNITEPRISEFF,PSA_HORAIREMOIS,PSA_ETABLISSEMENT,PSA_CODESTAT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4'+
                  ',PBG_FIXEAV,PBG_FIXEAP,PBG_PCTFIXE,PBG_VARIABLEAV,PBG_VARIABLEAP,PBG_PCTVARIABLE,PBG_ETATINTAUGM '+
                  ',PBG_MOTIFAUGM,PBG_COMMENTAIREABR,PBG_LIBELLEEMPLOI'+
                  ' FROM SALARIES LEFT JOIN BUDGETPAIE ON PSA_SALARIE=PBG_SALARIE '+
                  'WHERE '+RequeteSalAugm +' AND PBG_TYPEBUDG="AUG" AND PBG_ANNEE="'+AnneeSaisie+'"',True);
        TobSal := Tob.Create('LeSalarie',Nil,-1);
        TobSal.LoadDetailDB('LeSalarie','','',Q,False);
        Ferme(Q);
        {$IFDEF EMANAGER}
        If (ListeAug.AllSelected) or (ListeAug.NbSelected = 0) then
        {$ELSE}
        If ListeAug.AllSelected then
        {$ENDIF}
        begin
             {$IFDEF EAGLCLIENT}
             if (TFMul(Ecran).bSelectAll.Down) or (ListeAug.NbSelected = 0) then
             TFMul(Ecran).Fetchlestous;
             {$ENDIF}
             TobMul := Tob.Create('DonnesMul',Nil,-1);
             Q_MulAug.First;
             {$IFNDEF EAGLCLIENT}
             {$ELSE}
             {$ENDIF}
             InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', Q_MulAug.RecordCount,
                    False,True);
             While not Q_MulAug.eof do
             begin
                  TSA := TobSal.FindFirst(['PSA_SALARIE'],[Q_MulAug.FindField('PSA_SALARIE').AsString],False);
                  Etat := Q_MulAug.FindField('PBG_ETATINTAUGM').AsString;
                  If (Etat <> '011')  or (Argument = 'VALIDDRH') then
                  begin
                  TA := Tob.Create('FilleTobAug',TobAug,-1);
                  TA.AddChampSupValeur('INDICATEUR','',False);
                  TA.AddChampSupValeur('PSA_SALARIE',TSA.GetValue('PSA_SALARIE'),False);
                  TA.AddChampSupValeur('PSA_LIBELLE',TSA.GetValue('PSA_LIBELLE') + ' '+ TSA.GetValue('PSA_PRENOM'),False);
                  TA.AddChampSupValeur('PBG_LIBELLEEMPLOI',RechDom('PGLIBEMPLOI',TSA.GetValue('PSA_LIBELLEEMPLOI'),False),False);
                  TA.AddChampSupValeur('PSA_DATEENTREE',TSA.GetValue('PSA_DATEENTREE'),False);
                  TA.AddChampSupValeur('PBG_FIXEAV',TSA.GetValue('PBG_FIXEAV'),False);
                  TA.AddChampSupValeur('PBG_FIXEAP',TSA.GetValue('PBG_FIXEAP'),False);
                  TA.AddChampSupValeur('PBG_PCTFIXE',TSA.GetValue('PBG_PCTFIXE'),False);
                  TA.AddChampSupValeur('PBG_VARIABLEAV',TSA.GetValue('PBG_VARIABLEAV'),False);
                  TA.AddChampSupValeur('PBG_VARIABLEAP',TSA.GetValue('PBG_VARIABLEAP'),False);
                  TA.AddChampSupValeur('PBG_PCTVARIABLE',TSA.GetValue('PBG_PCTVARIABLE'),False);
                  TA.AddChampSupValeur('PBG_COMMENTAIREABR',TSA.GetValue('PBG_COMMENTAIREABR'),False);
                  TA.AddChampSupValeur('PBG_ETATINTAUGM',TSA.GetValue('PBG_ETATINTAUGM'),False);
                  AncBrut := TSA.GetValue('PBG_FIXEAV') + TSA.GetValue('PBG_VARIABLEAV');
                  NewBrut := TSA.GetValue('PBG_FIXEAP') + TSA.GetValue('PBG_VARIABLEAP');
                  If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                  else PctBrut := 0;
                  Effectif := TSA.GetValue('PSA_UNITEPRISEFF');
                  If Effectif = 0 then Effectif := 1;
                  Coeff := 1 / Effectif;
                  TA.AddChampSupValeur('TOTALAV',AncBrut,False);
                  TA.AddChampSupValeur('PCTTOTAL',PctBrut,False);
                  TA.AddChampSupValeur('TOTALAP',NewBrut,False);
                  TA.AddChampSupValeur('TPSPLEIN',NewBrut*coeff,False);
                  TA.AddChampSupValeur('SALANNAV',TSA.GetValue('PBG_FIXEAV')*NbMEquivalenceAnn + TSA.GetValue('PBG_VARIABLEAV')*12,False);
                  TA.AddChampSupValeur('SALANNAP',TSA.GetValue('PBG_FIXEAP')*NbMEquivalenceAnn + TSA.GetValue('PBG_VARIABLEAP')*12,False);
                  TA.AddChampSupValeur('SALANNTPSP',(TSA.GetValue('PBG_FIXEAP')*NbMEquivalenceAnn + TSA.GetValue('PBG_VARIABLEAP')*12)*Coeff,False);
                  TA.AddChampSupValeur('PBG_MOTIFAUGM',TSA.GetValue('PBG_MOTIFAUGM'),False);
                  TA.AddChampSupValeur('ETAT',TSA.GetValue('PBG_ETATINTAUGM'),False);
                  TA.AddChampSupValeur('MOTIF',TSA.GetValue('PBG_MOTIFAUGM'),False);
                  //Infoscompl
                  DateNaiss := TSA.Getvalue('PSA_DATENAISSANCE');
                  DateANC  := TSA.Getvalue('PSA_DATEANCIENNETE');
                  Anciennete := AncienneteMois(DateANC, Date);
                  AncAnnee := Anciennete div 12;
                  AncMois := Anciennete - AncAnnee * 12;
                  Age := 0;
                  PremMois := 0;
                  PremAnnee := 0;
                  If DateNaiss > IDate1900 then AglNombreDeMoisComplet(DateNaiss, Date, PremMois, PremAnnee, Age)
                  else Age := 0;
                  Age := Age div 12;
                  If Age > 0 then StAge := IntToStr(Age) + ' ans'
                  else StAge := 'Non renseigné';
                  If AncAnnee > 0 then
                  begin
                       If AncMois > 0 then StAnciennete := IntToStr(AncAnnee) + ' an(s)' + ' et '+ IntToStr(AncMois)+ ' mois'
                       else StAnciennete := IntToStr(AncAnnee) + ' an(s)';
                  end
                  else StAnciennete := IntToStr(AncMois) + ' mois';
                  TA.AddChampSupValeur('AGE',StAge,False);
                  TA.AddChampSupValeur('ANCIENETE',StAnciennete,False);
                  TA.AddChampSupValeur('PSA_ETABLISSEMENT',RechDom('TTETABLISSEMENT',TSA.GetValue('PSA_ETABLISSEMENT'),False),False);
                  HoraireMens := TSA.GetValue('PSA_HORAIREMOIS');
                  //TA.AddChampSupValeur('EFFECTIF','Pris dans l''effectif pour '+FloatToStr(Effectif),False);
                  TA.AddChampSupValeur('EFFECTIF',Effectif,False);
                  TA.AddChampSupValeur('HORAIRE','Horaire mensuel : '+FloatToStr(HoraireMens),False);
                  If (TSA.GetValue('PSA_TRAVAILN1') <> '') and (VH_Paie.PGNbreStatOrg > 0) then TA.AddChampSupValeur('PSA_TRAVAILN1',LibOrg1 + RechDom('PGTRAVAILN1',TSA.GetValue('PSA_TRAVAILN1'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN1',LibOrg1,False);
                  If (TSA.GetValue('PSA_TRAVAILN2') <> '') and (VH_Paie.PGNbreStatOrg > 1) then TA.AddChampSupValeur('PSA_TRAVAILN2',LibOrg2 + RechDom('PGTRAVAILN2',TSA.GetValue('PSA_TRAVAILN2'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN2',LibOrg2,False);
                  If (TSA.GetValue('PSA_TRAVAILN3') <> '') and (VH_Paie.PGNbreStatOrg > 2) then TA.AddChampSupValeur('PSA_TRAVAILN3',LibOrg3 + RechDom('PGTRAVAILN3',TSA.GetValue('PSA_TRAVAILN3'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN3',LibOrg3,False);
                  If (TSA.GetValue('PSA_TRAVAILN4') <> '') and (VH_Paie.PGNbreStatOrg > 3) then TA.AddChampSupValeur('PSA_TRAVAILN4',LibOrg4 + RechDom('PGTRAVAILN4',TSA.GetValue('PSA_TRAVAILN4'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN4',LibOrg4,False);
                  If TSA.GetValue('PSA_CODESTAT') <> '' then TA.AddChampSupValeur('PSA_CODESTAT',LibStat + RechDom('PGCODESTAT',TSA.GetValue('PSA_CODESTAT'),False),False)
                  else TA.AddChampSupValeur('PSA_CODESTAT',LibStat,False);
                  TA.AddChampSupValeur('CONVENTION',TSA.GetValue('PSA_CONVENTION'),False);
                   If Effectif <> 1 then TA.AddChampSupValeur('TPSPART','#ICO#84',False)
                  else TA.AddChampSupValeur('TPSPART','#ICO#81',False);
                  end;
                  MoveCurProgressForm ('Salarié : '+Q_MulAug.FindField('PSA_LIBELLE').AsString);
                  Q_MulAug.Next;
             end;
        end
        else
        begin
             InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', ListeAug.NbSelected-1,
                    False,True);
             for i := 0 to ListeAug.NbSelected-1 do
             begin
                  ListeAug.GotoLeBookmark(i);
                  {$IFDEF EAGLCLIENT}
                  TFmul(Ecran).Q.TQ.Seek(ListeAug.Row-1) ;
                  {$ENDIF}
                   Salarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
                   TSA := TobSal.FindFirst(['PSA_SALARIE'],[Salarie],False);
                  Etat := TSA.GetValue('PBG_ETATINTAUGM');
                  if (Etat <> '011') or (Argument = 'VALIDDRH') then
                  begin
                  TA := Tob.Create('FilleTobAug',TobAug,-1);
                  TA.AddChampSupValeur('INDICATEUR','',False);
                  TA.AddChampSupValeur('PSA_SALARIE',TSA.GetValue('PSA_SALARIE'),False);
                  TA.AddChampSupValeur('PSA_LIBELLE',TSA.GetValue('PSA_LIBELLE') + ' '+ TSA.GetValue('PSA_PRENOM'),False);
                  TA.AddChampSupValeur('PBG_LIBELLEEMPLOI',RechDom('PGLIBEMPLOI',TSA.GetValue('PSA_LIBELLEEMPLOI'),False),False);
                  TA.AddChampSupValeur('PSA_DATEENTREE',TSA.GetValue('PSA_DATEENTREE'),False);
                  TA.AddChampSupValeur('PBG_FIXEAV',TSA.GetValue('PBG_FIXEAV'),False);
                  TA.AddChampSupValeur('PBG_FIXEAP',TSA.GetValue('PBG_FIXEAP'),False);
                  TA.AddChampSupValeur('PBG_PCTFIXE',TSA.GetValue('PBG_PCTFIXE'),False);
                  TA.AddChampSupValeur('PBG_VARIABLEAV',TSA.GetValue('PBG_VARIABLEAV'),False);
                  TA.AddChampSupValeur('PBG_VARIABLEAP',TSA.GetValue('PBG_VARIABLEAP'),False);
                  TA.AddChampSupValeur('PBG_PCTVARIABLE',TSA.GetValue('PBG_PCTVARIABLE'),False);
                  TA.AddChampSupValeur('PBG_COMMENTAIREABR',TSA.GetValue('PBG_COMMENTAIREABR'),False);
                  TA.AddChampSupValeur('PBG_ETATINTAUGM',TSA.GetValue('PBG_ETATINTAUGM'),False);
                  AncBrut := TSA.GetValue('PBG_FIXEAV') + TSA.GetValue('PBG_VARIABLEAV');
                  NewBrut := TSA.GetValue('PBG_FIXEAP') + TSA.GetValue('PBG_VARIABLEAP');
                  If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                  else PctBrut := 0;
                  Effectif := TSA.GetValue('PSA_UNITEPRISEFF');
                  If Effectif = 0 then Effectif := 1;
                  Coeff := 1 / Effectif;
                  TA.AddChampSupValeur('TOTALAV',AncBrut,False);
                  TA.AddChampSupValeur('PCTTOTAL',PctBrut,False);
                  TA.AddChampSupValeur('TOTALAP',NewBrut,False);
                  TA.AddChampSupValeur('TPSPLEIN',NewBrut*coeff,False);
                  TA.AddChampSupValeur('SALANNAV',TSA.GetValue('PBG_FIXEAV')*NbMEquivalenceAnn + TSA.GetValue('PBG_VARIABLEAV')*12,False);
                  TA.AddChampSupValeur('SALANNAP',TSA.GetValue('PBG_FIXEAP')*NbMEquivalenceAnn + TSA.GetValue('PBG_VARIABLEAP')*12,False);
                  TA.AddChampSupValeur('SALANNTPSP',(TSA.GetValue('PBG_FIXEAP')*NbMEquivalenceAnn + TSA.GetValue('PBG_VARIABLEAP')*12)*Coeff,False);
                  TA.AddChampSupValeur('PBG_MOTIFAUGM',TSA.GetValue('PBG_MOTIFAUGM'),False);
                  TA.AddChampSupValeur('ETAT',TSA.GetValue('PBG_ETATINTAUGM'),False);
                  TA.AddChampSupValeur('MOTIF',TSA.GetValue('PBG_MOTIFAUGM'),False);
                  //Infoscompl
                  DateNaiss := TSA.Getvalue('PSA_DATENAISSANCE');
                  DateANC  := TSA.Getvalue('PSA_DATEANCIENNETE');
                  Anciennete := AncienneteMois(DateANC, Date);
                  AncAnnee := Anciennete div 12;
                  AncMois := Anciennete - AncAnnee * 12;
                  Age := 0;
                  PremMois := 0;
                  PremAnnee := 0;
                  If DateNaiss > IDate1900 then AglNombreDeMoisComplet(DateNaiss, Date, PremMois, PremAnnee, Age)
                  else Age := 0;
                  Age := Age div 12;
                  If Age > 0 then StAge := IntToStr(Age) + ' ans'
                  else StAge := 'Non renseigné';
                  If AncAnnee > 0 then
                  begin
                       If AncMois > 0 then StAnciennete := IntToStr(AncAnnee) + ' an(s)' + ' et '+ IntToStr(AncMois)+ 'mois'
                       else StAnciennete := IntToStr(AncAnnee) + ' an(s)';
                  end
                  else StAnciennete := IntToStr(AncMois) + ' mois';
                  TA.AddChampSupValeur('AGE',StAge,False);
                  TA.AddChampSupValeur('ANCIENETE',StAnciennete,False);
                  TA.AddChampSupValeur('PSA_ETABLISSEMENT',RechDom('TTETABLISSEMENT',TSA.GetValue('PSA_ETABLISSEMENT'),False),False);
                  HoraireMens := TSA.GetValue('PSA_HORAIREMOIS');
                  TA.AddChampSupValeur('EFFECTIF',FloatToStr(Effectif),False);
//                  TA.AddChampSupValeur('EFFECTIF','Pris dans l''effectif pour '+FloatToStr(Effectif),False);
                  TA.AddChampSupValeur('HORAIRE','Horaire mensuel : '+FloatToStr(HoraireMens),False);
                  If (TSA.GetValue('PSA_TRAVAILN1') <> '') and (VH_Paie.PGNbreStatOrg > 0) then TA.AddChampSupValeur('PSA_TRAVAILN1',LibOrg1 + RechDom('PGTRAVAILN1',TSA.GetValue('PSA_TRAVAILN1'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN1',LibOrg1,False);
                  If (TSA.GetValue('PSA_TRAVAILN2') <> '') and (VH_Paie.PGNbreStatOrg > 1) then TA.AddChampSupValeur('PSA_TRAVAILN2',LibOrg2 + RechDom('PGTRAVAILN2',TSA.GetValue('PSA_TRAVAILN2'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN2',LibOrg2,False);
                  If (TSA.GetValue('PSA_TRAVAILN3') <> '') and (VH_Paie.PGNbreStatOrg > 2) then TA.AddChampSupValeur('PSA_TRAVAILN3',LibOrg3 + RechDom('PGTRAVAILN3',TSA.GetValue('PSA_TRAVAILN3'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN3',LibOrg3,False);
                  If (TSA.GetValue('PSA_TRAVAILN4') <> '') and (VH_Paie.PGNbreStatOrg > 3) then TA.AddChampSupValeur('PSA_TRAVAILN4',LibOrg4 + RechDom('PGTRAVAILN4',TSA.GetValue('PSA_TRAVAILN4'),False),False)
                  else TA.AddChampSupValeur('PSA_TRAVAILN4',LibOrg4,False);
                  If TSA.GetValue('PSA_CODESTAT') <> '' then TA.AddChampSupValeur('PSA_CODESTAT',LibStat + RechDom('PGCODESTAT',TSA.GetValue('PSA_CODESTAT'),False),False)
                  else TA.AddChampSupValeur('PSA_CODESTAT',LibStat,False);
                  TA.AddChampSupValeur('CONVENTION',TSA.GetValue('PSA_CONVENTION'),False);
                  If Effectif <> 1 then TA.AddChampSupValeur('TPSPART','#ICO#84',False)
                  else TA.AddChampSupValeur('TPSPART','#ICO#81',False);
                  end;
                  MoveCurProgressForm ('Salarié : '+TSA.GetValue('PSA_LIBELLE'));

             end;
             ListeAug.ClearSelected;
             TobMul.Free;
        end;
        TobSal.Free;
        FiniMoveProgressForm;

end;

Function TOF_PGMUL_AUGMENTATION.RendSalairesTob(TypeSalaire : String;TS : Tob) : Double;
var LesSalaires,TempSalaire,Champ : String;
    Montant : Double;
    i : Integer;
begin
     LesSalaires := '';
     If TypeSalaire = 'FIXE' then LesSalaires := GetParamSoc('SO_PGAUGFIXE');
     If TypeSalaire = 'VARIABLE' then LesSalaires := GetParamSoc('SO_PGAUGVARIABLE');
     If LesSalaires = '' then
     begin
          Result := 0;
          Exit;
     end;
     Montant := 0;
     While LesSalaires <> '' do
     begin
          Champ := '';
          TempSalaire := ReadTokenPipe(LesSalaires,';');
          For i := 1 to 5 do
          begin
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PSA_SALAIRANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PSA_SALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then Montant := Montant + TS.GetValue(champ);
     end;
     Result := Montant;
end;

Function TOF_PGMUL_AUGMENTATION.SalarieExclus (TS,TC : Tob;Fixe,Variable : Double) : String;
var Anciennete : Word;
    i,a : Integer;
    LeCritere,LaDonnee,Valeur,Operateur,NouvelEtat : String;
    DE : TDateTime;
    Q : TQuery;
    TobDetail : Tob;
    VerifCritere : Boolean;
begin
     Result := '001';
     For i := 0 to TC.detail.Count - 1 do
     begin
          LeCritere := TC.Detail[i].GetValue('PAP_CRITEREAUGM');
          NouvelEtat := TC.Detail[i].GetValue('PAP_ETATINTAUGM');
          Q := OpenSQL('SELECT * FROM AUGMEXCLUS WHERE PAE_CRITEREAUGM="'+LeCritere+'"',True);
          TobDetail := Tob.Create('LeDetail',Nil,-1);
          TobDetail.LoadDetailDB('LeDetail','','',Q,False);
          Ferme(Q);
          VerifCritere := True;
          For a := 0 to TobDetail.detail.Count - 1 do
          begin
               Operateur := Tobdetail.Detail[a].GetValue('PAE_OPERATTEST');
               Valeur := Tobdetail.Detail[a].GetValue('PAE_VALEUREX');
               LaDonnee := TobDetail.Detail[a].GetValue('PAE_CRITEREAUGEX');
               //Durée d'absence
               If LaDonnee = 'ABS' then
               begin
               end
               //Ancienneté
               else If LaDonnee = 'ANC' then
               begin
                    DE := TS.Getvalue('PSA_DATEENTREE');
                    Anciennete := AncienneteMois(DE, Date);
                    If not Comparaison(Operateur,Anciennete,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Dernière augmentation
               else If LaDonnee = 'DAG' then
               begin
               end
               //Motif de sortie
               else If LaDonnee = 'MOS' then
               begin
               end
               //Niveau de rémuinération fixe
               else If LaDonnee = 'NIF' then
               begin
                    If not Comparaison(Operateur,Fixe,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Niveau de rémunération total
               else If LaDonnee = 'NIT' then
               begin
                    If not Comparaison(Operateur,Fixe + Variable,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Niveau de rémunération variable
               else If LaDonnee = 'NIV' then
               begin
                    If not Comparaison(Operateur,Variable,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Temps de travail
               else If LaDonnee = 'TPS' then
               begin
               end
               //Type de contrat
               else If LaDonnee = 'TYC' then
               begin
               end;
          end;
          TobDetail.Free;
          If VerifCritere then
          begin
               Result := NouvelEtat;
               If NouvelEtat = '011' then Exit;
          end;
     end;
end;

Function TOF_PGMUL_AUGMENTATION.Comparaison(Operateur : String;Val1,Val2 : Variant) : Boolean;
begin
     Result := False;
     If Operateur = '>' then
     begin
          If Val1 > Val2 then Result := True;
     end
     else If Operateur = '>=' then
     begin
          If Val1 >= Val2 then Result := True;
     end
     else If Operateur = '<' then
     begin
          If Val1 < Val2 then Result := True;
     end
     else If Operateur = '<=' then
     begin
          If Val1 <= Val2 then Result := True;
     end
     else If Operateur = '=' then
     begin
          If Val1 = Val2 then Result := True;
     end
     else If Operateur = '<>' then
     begin
          If Val1 <> Val2 then Result := True;
     end;
end;

procedure TOF_PGMUL_AUGMENTATION.SuppAugmentations(Sender : Tobject);
var Salarie,TypeB,Annee : String;
    NumOrdre,i : Integer;
    Rep : Word;
begin
        if (ListeAug.NbSelected = 0) and (not ListeAug.AllSelected) then
        begin
             MessageAlerte('Vous devez sélectionner le(s) salarié(s) à supprimer');
             exit;
        end;
        Rep := PGIAskCancel('Voulez-vous supprimer les augmentations sélectionnées ?',Ecran.Caption);
        If Rep <> MrYes then Exit;
        If ListeAug.AllSelected then
        begin
             {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
             {$ENDIF}
//             TobMul := Tob.Create('DonnesMul',Nil,-1);
             Q_MulAug.First;
             InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', Q_MulAug.RecordCount,
                    False,True);

             While Not Q_MulAug.Eof do
             begin
                  Salarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
                  TypeB := Q_MulAug.FindField('PBG_TYPEBUDG').AsString;
                  NumOrdre := Q_MulAug.FindField('PBG_NUMORDRE').AsInteger;
                  Annee := Q_MulAug.FindField('PBG_ANNEE').AsString;
                  ExecuteSQL('DELETE FROM BUDGETPAIE WHERE '+
                  'PBG_SALARIE="'+Salarie+'" AND PBG_ANNEE="'+Annee+'" AND PBG_TYPEBUDG="'+TypeB+'" '+
                  'AND PBG_NUMORDRE='+IntToStr(NumOrdre));
                  MoveCurProgressForm ('Salarié : '+Salarie);
                  MoveCurProgressForm ('Salarié : '+Salarie);
                  Q_MulAug.Next;
             end;
        end
        else
        begin
             InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', ListeAug.NbSelected,
                    False,True);
             for i := 0 to ListeAug.NbSelected -1 do
             begin
                  ListeAug.GotoLeBookmark(i);
                  {$IFDEF EAGLCLIENT}
                  TFmul(Ecran).Q.TQ.Seek(ListeAug.Row-1) ;
                  {$ENDIF}
                  Salarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
                  TypeB := Q_MulAug.FindField('PBG_TYPEBUDG').AsString;
                  NumOrdre := Q_MulAug.FindField('PBG_NUMORDRE').AsInteger;
                  Annee := Q_MulAug.FindField('PBG_ANNEE').AsString;
                  ExecuteSQL('DELETE FROM BUDGETPAIE WHERE '+
                  'PBG_SALARIE="'+Salarie+'" AND PBG_ANNEE="'+Annee+'" AND PBG_TYPEBUDG="'+TypeB+'" '+
                  'AND PBG_NUMORDRE='+IntToStr(NumOrdre));
                  MoveCurProgressForm ('Salarié : '+Salarie);
             end;
        end;
        FiniMoveProgressForm;
        ListeAug.AllSelected := False;
        ListeAug.ClearSelected;
        TFMul(Ecran).bSelectAll.Down := False;
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMUL_AUGMENTATION.AfficheLegende(Sender : TObject);
begin
     AGLLanceFiche('PAY','AUGMLEGENDE','','','');
end;

Function TOF_PGMUL_AUGMENTATION.RendRequeteResp(ValeurResp,ChampResp : String) : String;
var StResp,Requete : String;
begin
     Result := '';
     requete := '';
     While ValeurResp <> '' do
     begin
          StResp := ReadTokenPipe(ValeurResp,';');
          If Requete <> '' then requete := Requete + ' OR ' + ChampResp+'="'+StResp+'"'
          else Requete := ChampResp+'="'+StResp+'"';
     end;
     Result := Requete;
end;

procedure TOF_PGMUL_AUGMENTATION.SalarieElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StOrder := 'PSA_LIBELLE';
        StWhere := '';
        {$IFDEF EMANAGER}
        If Argument = 'SAISIE' then StWhere := 'PSE_RESPONSVAR="'+V_PGI.UserSalarie+'"';
        If Argument = 'VALIDRESP' then
        begin
             StWhere := '(PSA_SALARIE="'+V_PGI.UserSalarie+'" OR PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICES ON PSE_CODESERVICE=PGS_CODESERVICE LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))';
        end;
        {$ENDIF}
        If GetControlText('RESPONSABS') <> '' then
        begin
             If StWhere <> '' then StWhere := StWhere + ' AND PSE_RESPONSABS="'+GetControlText('RESPONSABS')+'"'
             else StWhere := 'PSE_RESPONSABS="'+GetControlText('RESPONSABS')+'"';
        end;
        LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;

procedure TOF_PGMUL_AUGMENTATION.RespAbsElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StOrder := 'PSI_LIBELLE';
        StWhere := 'PSI_INTERIMAIRE IN (SELECT DISTINCT PSE_RESPONSABS FROM DEPORTSAL)';
        {$IFDEF EMANAGER}
//        If Argument = 'SAISIE' then StWhere := 'PSE_RESPONSVAR="'+V_PGI.UserSalarie+'" AND PSA_SALARIE IN (SELECT DISTINCT PSE_RESPONSABS FROM DEPORTSAL)';
//        If Argument = 'VALIDRESP' then
//        begin
          If consultation then StWhere := '(PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSABS FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSABS="'+V_PGI.UserSalarie+'")))'
          else StWhere := '(PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSABS FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))';
//        end;
        {$ENDIF}
        LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES LEFT JOIN DEPORTSAL ON PSI_INTERIMAIRE=PSE_SALARIE','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
end;

Initialization
  registerclasses ( [ TOF_PGMUL_AUGMENTATION ] ) ;
end.

