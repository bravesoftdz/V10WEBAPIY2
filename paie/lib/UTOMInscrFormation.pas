{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 12/07/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : INSCFORMATION (INSCFORMATION)
Mots clefs ... : TOM;INSCFORMATION
*****************************************************************
PT1  | 31/10/2003 | V_50  | JL | Modif pour accès en validation dans la fiche INSCFORMATION
PT2  | 12/10/2003 | V_50  | JL | Modif pour validation par responsables
PT3  | 12/10/2003 | V_50  | JL | Calcul en fonction du lieu, et nb heures du stagiaires au lieu du stage
PT4  | 17/12/2003 | V_50  | JL | Suppression de l'investissement correspondant aux stage si plus d'inscriptions
PT5  | 01/09/2004 | V_50  | JL | Ajout taux de charge pour le budget
PT6  | 11/04/2005 | V_60  | JL | FQ 12174 correction blocage saisie si pas de validation
PT7  | 31/03/2006 | V_65  | JL | FQ 13005 Plus de validation par défaut en cas de demande de DIF
PT8  | 12/04/2006 | V_65  | JL | FQ 13034 Conditionnement motif décision
PT9  | 12/04/2006 | V_65  | JL | FQ 13033 Pour le DIF millésime par défaut = plan de formation au lieu de prévisionnel
PT10 | 30/06/2006 | V_70  | JL | FQ 13337 1 repas au lieu de 2 si pas d'hébergement
---- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT11 | 26/01/2007 | V_75  | JL | Supp Tree View
PT12 | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
     |            |       |    | pour les critères code salarié uniquement
PT13 | 13/04/2007 | V_720 | FL | FQ 13568 Utilisation des méthodes prévisionnelles de calcul du coût salarié
PT14 | 17/04/2007 | V_720 | FL | FQ 14071 Calcul du T.H. salarié : prise en compte de l'historique salarié s'il existe
PT15 | 10/05/2007 | V_720 | FL | FQ 13567 Prise en compte des populations pour le calcul des frais prévisionnnels
PT16 | 14/11/2007 | V_80  | FL | Déplacement de la création du stage au prévisionnel dans PGOutilsFormation
PT17 | 28/09/2007 | V_7   | FL | Emanager / Report / Adaptation cursus + accès assistant
PT18 | 22/10/2007 | V_7   | FL | Report / Adaptation CWAS
PT19 | 04/12/2007 | V_8   | FL | Emanager / Report / Gestion des sous-niveaux
PT20 | 08/01/2008 | V_802 | FL | Correction de bugs divers
PT21 | 05/02/2008 | V_803 | FL | La civilité n'est pas indiquée dans la table Intérimaires
PT22 | 12/02/2008 | V_803 | FL | Multidossier : rechercher le nom du salarié dans la tablette PGSALARIEINT
PT23 | 10/03/2008 | V_803 | FL | Appel de la fonction PrepareMailFormation pour le corps du mail à envoyer
PT24 | 04/04/2008 | V_803 | FL | Correctif sur StageElipsis en cas de partage formation 
PT25 | 15/04/2008 | V_804 | FL | Récupération du mail d'abord depuis la table UTILISAT, puis ensuite depuis DEPORTSAL
PT26 | 05/05/2008 | V_804 | FL | Mise à jour du libellé des inscrits aux cursus + affectation du dossier de connexion lors d'insc non nominatives
PT27 | 02/06/2008 | V_804 | FL | Envoi d'un mail lors d'une décision DIF
PT28 | 05/06/2008 | V_804 | FL | FQ 15458 Report V7 Maj des coûts de tous les salariés, même ceux qui ne sont pas visibles (confidentiels)
PT29 | 11/06/2008 | V_804 | FL | Possibilité de suppression d'une inscription faisant partie d'un cursus
PT30 | 06/09/2008 | V_850 | NA | FQ 15694 Controle date de début de l'exercice >= date sortie si elle est renseignée
}
unit UTOMInscrFormation;

interface

uses StdCtrls, Controls, Classes,UtilWord,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  Fiche, FE_Main,HDB,
  {$ELSE}
  eFiche, eFichList,MainEAGL,
  {$ENDIF}
  sysutils,  HCtrls, HEnt1, HMsgBox, UTOM, UTob, LookUp, uTableFiltre,P5DEF,
  SaisieList, HSysMenu, HTB97, PgoutilsFormation, ParamDat, EntPaie,  HStatus,
  ed_tools,ParamSoc,MailOl,PGOutils2,PGOutils;

type
  TOM_INSCFORMATION = class(TOM)
    procedure Onclose; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    private
    {$IFDEF EMANAGER}
    TypeUtilisat : String;
    {$ENDIF}
    SalarieDIF : String;
    InitSalarie, ArgStage, ArgMillesime, LeCursus,ArgSalarie : string; //PT18
    Utilisateur, MillesimeEC, TypeSaisie, TypeInsc: string;
    TF: TTableFiltre;
    MAilSaisie,MailDecision : Boolean;
    THVMillesime: THValComboBox;
    AfficherLibelle,NewCursus: Boolean;
    InitNbSal: Integer;
    TauxChargeBudget: Double;
    CoutAtt, CoutValid: Double;
    TobNature : Tob;
    DefautHTpsT : Boolean;
    TypeDePlan : String;
    EtatInit : String;
    InscriptionsStd : Boolean;
    WDatedebexerc : Tdatetime; // pt30
    procedure BAcceptClick(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure AfficheInfosStage(Sender: TObject);
    procedure SalarieElipsisClick(Sender: TObject);
    procedure BSelectionClick(Sender: TObject);
    procedure InscriptionCursus;
    procedure SuppressionCursus;
    procedure CreerCursusPrev;
    procedure CreerStagePrev(Stage: string);
    function CalCoutBudgetFor(Stage, Millesime, NbInsc: string): Double;
    procedure MajAfficheCoutInsc(Individuel, Global: Boolean);
    procedure MajEtatFormCursus(cursus,Etab,millesime,etat,motifinsc,motifdec : String;rang : Integer);
    procedure BZoomCLick(Sender : TObject);
//    Function RendAllocationFormation(Salarie : String) : Double;
    procedure AffectationHeures;
    procedure EnvoiMailresponsable(Action : String);
    procedure RecupInfosStage;
    procedure VoirCompteursDIF(Sender : TObject);
    procedure StageElipsisClick(Sender : TObject);
    Procedure MajCoutPrev(Stage,Millesime : String);
  end;


implementation

procedure TOM_INSCFORMATION.OnClose;
begin
inherited;
If TobNature <> Nil then freeandnil(TobNature);
end;

procedure TOM_INSCFORMATION.OnDeleteRecord;
var
  Q: TQuery;
  NbInsc : Integer;
  {$IFNDEF EMANAGER}
  Rep : Word;
  {$ENDIF}
begin
  inherited;
  If Ecran Is TFFiche then TFFiche(Ecran).Retour := 'MODIF';
  {$IFDEF EMANAGER}
  if (TypeSaisie = 'CWASINSCBUDGET') or (TypeSaisie = 'CURSUS') then
  begin
        If GetField('PFI_ETATINSCFOR') = 'VAL' then
        begin
                PGIBox('Vous ne pouvez pas supprimer cette isncription car elle a été validée',Ecran.Caption);
                LastError := 1;
                Exit;
        end;
  end;
  {$ENDIF}
  if GetField('PFI_CURSUS') <> '' then
  begin
    If GetField('PFI_CODESTAGE') = '--CURSUS--' then SuppressionCursus
    else
    begin
      //PT29
      If PGIAsk(TraduireMemoire('Attention : L''inscription fait partie du cursus '+RechDom('PGCURSUS', GetField('PFI_CURSUS'), False)+'.#13#10Êtes-vous sûr de vouloir la supprimer?'),Ecran.Caption) = mrNo Then
      Begin
      	LastError := 1; 
      	Exit;
      End;
      //PGIBox('Vous ne pouvez pas supprimer individuellement cette inscription car elle fait partie du cursus : ' + RechDom('PGCURSUS', GetField('PFI_CURSUS'), False),
      //Ecran.Caption);
    end;
  end;
  NbInsc := 0;
  Q := OpenSQL('SELECT COUNT(PFI_ETABLISSEMENT) AS NBETAB FROM INSCFORMATION WHERE PFI_MILLESIME="' + GetField('PFI_MILLESIME') + '"' +
    ' AND PFI_CODESTAGE="' + GetField('PFI_CODESTAGE') + '"', True);
  if not Q.Eof then NbInsc := Q.FindField('NBETAB').AsInteger;
  Ferme(Q);
  if NbInsc = 1 then
  begin
//      Stage := GetField('PFI_CODESTAGE');
//      ExecuteSQL('DELETE FROM STAGE WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + GetField('PFI_MILLESIME') + '"');
//      ExecuteSQL('DELETE FROM INVESTSESSION WHERE PIS_ORDRE=-1 AND PIS_CODESTAGE="' + Stage + '" AND PIS_MILLESIME="' + GetField('PFI_MILLESIME') + '"'); //PT4 //DB2
//    end
//    else
//    begin
  {$IFNDEF EMANAGER}
      Rep := PGIAsk('Attention, il ne reste plus d''inscriptions pour la formation ' + RechDom('PGSTAGEFORM', GetField('PFI_CODESTAGE'), False) +
        '#13#10 Voulez-vous enlever formation du budget ' + GetField('PFI_MILLESIME'), Ecran.Caption);
      if Rep = MrYes then
      begin
        ExecuteSQL('DELETE FROM STAGE WHERE PST_CODESTAGE="' + GetField('PFI_CODESTAGE') + '" AND PST_MILLESIME="' + GetField('PFI_MILLESIME') + '"');
        ExecuteSQL('DELETE FROM INVESTSESSION WHERE PIS_ORDRE=-1 AND PIS_CODESTAGE="' + GetField('PFI_CODESTAGE') + '" AND PIS_MILLESIME="' + GetField('PFI_MILLESIME') + '"');
          //PT4 DB2
      end;
     {$ENDIF}
//    end;
  end;
end;

procedure TOM_INSCFORMATION.OnNewRecord;
var
  Stage: string;
  Q: TQuery;
begin
  inherited;
  SetField('PFI_PREDEFINI','DOS');
  SetField('PFI_NODOSSIER','000000');
  If TypeSaisie = 'SAISIESALARIE' then SetField('PFI_SALARIE',ArgStage);
  {$IFDEF EMANAGER}
  SetField('PFI_MOTIFINSCFOR','2'); //PT17
  if TypeSaisie = 'GESTIONDIF' then SetField('PFI_MOTIFINSCFOR','5');
  SetField('PFI_NIVPRIORITE','01'); //PT17
  {$ENDIF}
  if TypeSaisie = 'GESTIONDIF' then
  begin
       SetField('PFI_TYPEPLANPREV','DIF');
       SetControlEnabled('PFI_MILLESIME',False);
  end
  else If VH_Paie.PGForPrevisionnel then SetField('PFI_TYPEPLANPREV','PLF') //PT6
  else SetField('PFI_TYPEPLANPREV','DIF');
  SetField('PFI_DATEDIF',V_PGI.DateEntree);
  if TypeSaisie = 'CURSUS' then SetField('PFI_CURSUS', LeCursus);
  SetField('PFI_TYPOFORMATION','001');
  SetField('PFI_DATEACCEPT', IDate1900);
  SetField('PFI_REFUSELE', IDate1900);
  SetField('PFI_REPORTELE', IDate1900);
  SetField('PFI_ETATINSCFOR', 'ATT');
  SetField('PFI_NBINSC', 1);
  SetField('PFI_REALISE', '-');
  If TypeSaisie <> 'SAISIESALARIE' then SetField('PFI_CODESTAGE', ArgStage);
  SetField('PFI_LIBELLE', '1 Salarié');
  SetField('PFI_MILLESIME', MillesimeEC);
  {$IFDEF EMANAGER}
  If typeUtilisat = 'R' then SetField('PFI_RESPONSFOR', utilisateur);
  {$ENDIF}
  if TypeSaisie <> 'CURSUS' then
  begin
    if ((TypeSaisie = 'CWASINSCBUDGET') and (ArgStage<>'')) or (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESTAGE') then //PT18
    //if (TypeSaisie = 'CWASINSCBUDGET') or (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESTAGE') then
    begin
      Stage := ArgStage;
      Q := OpenSQL('SELECT PST_JOURSTAGE,PST_LIEUFORM,PST_DUREESTAGE,' +
        'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_NATUREFORM' +
        ' FROM STAGE WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="0000"', True);
      if not Q.eof then
      begin
        SetField('PFI_JOURSTAGE', Q.FindField('PST_JOURSTAGE').AsFloat);
        SetField('PFI_DUREESTAGE', Q.FindField('PST_DUREESTAGE').AsFloat);
        if (DefautHTpsT) or (GetField('PFI_TYPEPLANPREV') = 'PLF') then
        begin
             SetField('PFI_HTPSTRAV', Q.FindField('PST_DUREESTAGE').AsFloat);
             SetField('PFI_HTPSNONTRAV', 0);
        end
        else
        begin
             SetField('PFI_HTPSTRAV', 0);
             SetField('PFI_HTPSNONTRAV', Q.FindField('PST_DUREESTAGE').AsFloat);
        end;
        SetField('PFI_LIEUFORM', Q.FindField('PST_LIEUFORM').AsString);
        SetField('PFI_FORMATION1', Q.FindField('PST_FORMATION1').AsString);
        SetField('PFI_FORMATION2', Q.FindField('PST_FORMATION2').AsString);
        SetField('PFI_FORMATION3', Q.FindField('PST_FORMATION3').AsString);
        SetField('PFI_FORMATION4', Q.FindField('PST_FORMATION4').AsString);
        SetField('PFI_FORMATION5', Q.FindField('PST_FORMATION5').AsString);
        SetField('PFI_FORMATION6', Q.FindField('PST_FORMATION6').AsString);
        SetField('PFI_FORMATION7', Q.FindField('PST_FORMATION7').AsString);
        SetField('PFI_FORMATION8', Q.FindField('PST_FORMATION8').AsString);
        SetField('PFI_NATUREFORM', Q.FindField('PST_NATUREFORM').AsString);
      end;
      Ferme(Q);
    end;
  end
  else SetField('PFI_CODESTAGE', '--CURSUS--');
  //PT18 - Début
  if (TypeSaisie = 'CWASINSCBUDGET') and (ArgStage = '') then
  begin
       SetField('PFI_SALARIE',ArgSalarie);
       If (ArgSalarie = '') and (THValComboBox(GetControl('FILTREEMPLOI')) <> Nil) then
       begin
            If GetControlText('FILTREEMPLOI')<>'' then SetField('PFI_LIBEMPLOIFOR',getControlText('FILTREEMPLOI'));
       end;
  end;
  //PT18 - FIn
end;

procedure TOM_INSCFORMATION.OnUpdateRecord;
var
  Q: TQuery;
  Salaire: Double;
  MessError: string;
  NbInsc : Integer;
  NbJour, NbHeure, FraisH, FraisR, FraisT, TotalFrais : Double;
  DateExercice : TDateTime; //PT14
  Req : String; //PT15
  TobEcran : TOB; //PT15
begin
  inherited;
  If Ecran Is TFFiche then TFFiche(Ecran).Retour := 'MODIF';
  If GetField('PFI_SALARIE') <> '' then
  begin
       Q := OpenSQL('SELECT PSE_RESPONSFOR FROM DEPORTSAL WHERE PSE_SALARIE="' + GetField('PFI_SALARIE') + '"', True);
       if not Q.eof then SetField('PFI_RESPONSFOR', Q.Findfield('PSE_RESPONSFOR').AsString);
       Ferme(Q);
  end;
  If getField('PFI_CODESTAGE') = '' then
  begin
       PGIBox('Vous devez renseigner la formation',Ecran.Caption);
       LastError := 1;
       Exit;
  end;
  If GetField('PFI_TYPEPLANPREV') <> 'PLF' then
  begin
       If GetField('PFI_DATEDIF') <= IDate1900 then SetField('PFI_DATEDIF',V_PGI.DateEntree);
       If getField('PFI_SALARIE') = '' then
       begin
            PGIBox('Vous devez renseigner le salarié',Ecran.Caption);
            LastError := 1;
            Exit;
       end;
  end;
  if (VH_Paie.PGForValidPrev = False) and (GetField('PFI_TYPEPLANPREV')<>'DIF') then SetField('PFI_ETATINSCFOR', 'VAL'); //PT7
  if (DS.State in [dsInsert]) or (IsFieldModified('PFI_NBINSC')) OR (IsFieldModified('PFI_HTPSTRAV')) then
  begin
    if DS.State in [dsInsert] then
    begin
      If GetField('PFI_SALARIE') <> '' then
      begin
        if ExisteSQL('SELECT PFI_SALARIE FROM INSCFORMATION WHERE PFI_SALARIE="' + GetField('PFI_SALARIE') + '" AND ' +
        'PFI_CODESTAGE="' + GetField('PFI_CODESTAGE') + '" AND PFI_MILLESIME="' + MillesimeEC + '"') then
        begin
          if PGIAskCancel('Attention il existe déja une inscription pour ce salarié et cette formation, voulez-vous continuer ?', Ecran.Caption) <> MrYes then
          begin
            LastError := 1;
            Exit;
          end;
        end;
      end;
    end;
    MessError := '';
    if GetField('PFI_LIBEMPLOIFOR') = '' then
    begin
      SetFocusControl('PFI_LIBEMPLOIFOR');
      MessError := MessError+'- Le libellé emploi#13#10';
    end;
    {$IFDEF EMANAGER}
    If GetField('PFI_SALARIE') = '' then
    begin
            If TypeUtilisat = 'R' then
            begin
                    If Not ExisteSQL('SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
                        'WHERE PSA_LIBELLEEMPLOI="'+getField('PFI_LIBEMPLOIFOR')+'" AND '+ //PT17
                        '(PSE_RESPONSFOR="'+Utilisateur+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                        ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))') then
                    begin
                                PGIBox('Vous ne pouvez pas saisir d''inscription pour ce libellé emploi',Ecran.Caption);
                                LastError := 1;
                                Exit;
                    end;
            end
            else
            begin
                    PGIBox('Vous ne pouvez pas saisir d''inscriptions non nominatives',Ecran.Caption);
                    LastError := 1;
                    Exit;
            end;
    end;
    if GetField('PFI_MOTIFINSCFOR') = '' then
    begin
        SetFocusControl('PFI_LIBEMPLOIFOR');
        MessError := MessError+'- Le motif d''inscription#13#10';
    end;
    {$ENDIF}
    if GetField('PFI_ETABLISSEMENT') = '' then
    begin
      SetFocusControl('PFI_ETABLISSEMENT');
      MessError := MessError + '- L''établissement';
    end;
    if MessError <> '' then
    begin
      LastError := 1;
      PGIBox('Vous devez renseigner :#13#10 ' + MessError,Ecran.Caption);
      Exit;
      //LastErrorMsg := TraduireMemoire('Vous devez renseigner :#13#10 '+MessError,TFSaisieList(Ecran).Caption);
    end
    else
    begin
      if typeSaisie <> 'CURSUS' then
      begin
        NbInsc := GetField('PFI_NBINSC');
        if NbInsc <= 0 then NbInsc := 1;
        if GetField('PFI_SALARIE') <> '' then
        begin
          if VH_Paie.PGForValoSalairePrev = 'VCR' Then
          //PT14 - Début
          Begin
               DateExercice := 0;
               Q := OpenSQL('SELECT PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_MILLESIME="'+GetField('PFI_MILLESIME')+'"',True);
               If Not Q.Eof Then DateExercice := Q.FindField('PFE_DATEDEBUT').AsDateTime;
               Ferme(Q);
               Salaire := ForTauxHoraireReel(GetField('PFI_SALARIE'),0,0,'',ValPrevisionnel,DateExercice); //PT13 //PT14
          End
          //PT14 - Fin
          else
               Salaire := ForTauxHoraireCategoriel(GetField('PFI_LIBEMPLOIFOR'), GetField('PFI_MILLESIME'));
          SetField('PFI_COUTREELSAL', Salaire * GetField('PFI_HTPSTRAV') * TauxChargeBudget); //PT5
        end
        else
        begin
          Salaire := ForTauxHoraireCategoriel(GetField('PFI_LIBEMPLOIFOR'), MillesimeEC);
          If argStage = '' then NbHeure := GetField('PFI_DUREESTAGE') //PT18
          else NbHeure := StrtoFloat(Getcontroltext('DUREESTAGE')); 
          SetField('PFI_COUTREELSAL', Salaire * NbHeure * TauxChargeBudget); //PT5
        end;
        Req := 'SELECT PFF_FRAISHEBERG,PFF_FRAISREPAS,PFF_FRAISTRANSP FROM FORFAITFORM ' +
               'WHERE PFF_MILLESIME="' + MillesimeEC + '" AND PFF_LIEUFORM="' + Getfield('PFI_LIEUFORM') + '"'; //PT3
        //PT15 - Début
        If (VH_Paie.PGForGestFraisByPop) Then
        Begin
               // Un salarié a été saisi : on récupère sa population directement
               If GetField('PFI_SALARIE') <> '' Then
                    Req := Req + ' AND PFF_POPULATION IN (SELECT PNA_POPULATION FROM SALARIEPOPUL WHERE PNA_SALARIE="'+GetField('PFI_SALARIE')+'")'
               // Un nombre de salarié a été saisi (et donc pas de salarié précis) : on essaye de déterminer la population
               // en fonction des données présentes
               Else
               Begin
                    TobEcran := Tob.Create('INSCFORMATION', Nil, -1);
                    TobEcran.GetEcran(Ecran);
                    Req := Req + ' AND PFF_POPULATION="'+RecherchePopulation(TobEcran)+'"';
                    FreeAndNil (TobEcran);
               End;
        End
        Else
        //PT15 - Fin
               Req := Req + ' AND PFF_ETABLISSEMENT="' + GetField('PFI_ETABLISSEMENT') + '"';

        Q := OpenSQL(Req,True);
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
        NbJour := GetField('PFI_JOURSTAGE'); //PT3
        If GetField('PFI_SALARIE') = '' then NbJour := NbJour / NbInsc;
        FraisH := FraisH * (NbJour - 1);
        if FraisH < 0 then FraisH := 0;
        If FraisH > 0 then FraisR := FraisR*((NbJour*2)-1) //PT10
        else FraisR := FraisR*NbJour;
        if FraisR < 0 then FraisR := 0;
        TotalFrais := FraisH + FraisR + FraisT;
        If GetField('PFI_SALARIE') = '' then SetField('PFI_FRAISFORFAIT', TotalFrais * NbInsc)
        else SetField('PFI_FRAISFORFAIT', TotalFrais);
        SetField('PFI_AUTRECOUT', CalCoutBudgetFor(GetField('PFI_CODESTAGE'), GetField('PFI_MILLESIME'), IntToStr(NbInsc)));
      end;
    end;
  end;
  if DS.State in [dsInsert] then
  begin
    NewCursus := True;
    if GetField('PFI_MILLESIME') = '0000' then SetField('PFI_MILLESIME', MillesimeEC);
    Q := OpenSQL('SELECT MAX (PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_ETABLISSEMENT="' + Getfield('PFI_ETABLISSEMENT') + '" AND PFI_MILLESIME="' + MillesimeEC + '"', True);
    if not Q.eof then SetField('PFI_RANG', Q.FindField('RANG').AsInteger + 1)
    else SetField('PFI_RANG', 1);
    Ferme(Q);
  end
  else NewCursus := False;

  if GetField('PFI_CURSUS') <> '' then
  begin
   MajEtatFormCursus(GetField('PFI_CURSUS'),GetField('PFI_ETABLISSEMENT'),GetField('PFI_MILLESIME'),GetField('PFI_ETATINSCFOR'),GetField('PFI_MOTIFINSCFOR'),GetField('PFI_MOTIFDECISION'),GetField('PFI_RANG'));
  end;
  MailDecision := False;
  MailSaisie := False;
  {$IFDEF EMANAGER}
  If GetField('PFI_TYPEPLANPREV')=  'DIF' then
  begin
       If DS.State in [DsInsert] then MailSaisie := True
       else if EtatInit <> GetField('PFI_ETATINSCFOR') then MailDecision := True;
  end;
  {$ELSE}
  If GetField('PFI_TYPEPLANPREV')=  'DIF' then
  begin
    If DS.State in [DsInsert] then
    begin
         If GetParamSoc('SO_PGDIFMAILSAISIE') = true then EnvoiMailresponsable('SAISIE');
    end
    else
    begin
         if (GetParamSoc('SO_PGDIFMAILVALID') = true ) and (EtatInit <> GetField('PFI_ETATINSCFOR')) then EnvoiMailresponsable('DECISION');
    end;
  end;
//  If GetParamSoc('SO_PGDIFMAILVALID') = true then
//  begin
//       If (EtatInit <> GetField('PFI_ETATINSCFOR') ) and (GetField('PFI_ETATINSCFOR') <> 'ATT') then EnvoiMailresponsable('DECISION');
//  end;
  {$ENDIF}
end;

procedure TOM_INSCFORMATION.OnLoadRecord;
var
  i: Integer;
  Libelle, Nature,Stage: string;
  TobNat : Tob;
  Salarie : String;
begin
  inherited;
  SetControlEnabled('PFI_NATUREFORM',False);
  If SalarieDIF <> '' then SetField('PFI_SALARIE',SalarieDIF);
  EtatInit := GetField('PFI_ETATINSCFOR');
  TypeDePlan := GetField('PFI_TYPEPLANPREV');
  if not (Ecran is TFSaisieList) then
  begin
       If TypeDePlan = 'DIF' then
       begin
            TFFiche(Ecran).Caption := 'Saisie Droit Individuel à la Formation';
            UpdateCaption(TFFiche(Ecran));
       end;
  end;
  If TypeSaisie = 'GESTIONDIF' then
  begin
       if (DS.State in [dsInsert]) then SetControlEnabled('BZOOM',True)
       else SetControlEnabled('BZOOM',False);
  end;
  If GetField('TYPEPLANPREV') <> 'PLF' then
  begin
       SetControlEnabled('PFI_LIBEMPLOIFOR',False);
       SetControlenabled('PFI_ETABLISSEMENT',False);
  end;
  Salarie := GetField('PFI_SALARIE');
  If (GetField('PFI_ETATINSCFOR') = 'VAL') and (VH_Paie.PGForValidPrev = False) then
  begin
        SetControlEnabled('PFI_NBINSC',False);
        SetControlEnabled('PFI_SALARIE',False);
        SetControlEnabled('PFI_ETABLISSEMENT',False);
        SetControlEnabled('PFI_LIBEMPLOIFOR',False);
  end
  else
  begin
        SetControlEnabled('PFI_NBINSC',True);
        SetControlEnabled('PFI_SALARIE',True);
        SetControlEnabled('PFI_ETABLISSEMENT',True);
        SetControlEnabled('PFI_LIBEMPLOIFOR',True);
  end;
  SetControlEnabled('PFI_DADSCAT', False);
  if not AfterInserting then
  begin
    SetControlEnabled('PFI_ETABLISSEMENT', False);
    SetControlEnabled('PFI_LIBEMPLOIFOR', False);
  end
  else
  begin
    SetControlEnabled('PFI_LIBEMPLOIFOR', True);
    SetControlEnabled('PFI_ETABLISSEMENT', True);
  end;
  InitSalarie := GetField('PFI_SALARIE');
  InitNbSal := GetField('PFI_NBINSC');
  if not (Ecran is TFSaisieList) then
  begin
    if GetField('PFI_CURSUS') <> '' then
    begin
      TFFiche(Ecran).Caption := 'Inscription de ' + GetField('PFI_LIBELLE') + ' au cursus ' + RechDom('PGCURSUS', GetField('PFI_CURSUS'), False);
      SetControlVisible('PFI_CODESTAGE', False);
      SetControlVisible('TPFI_CODESTAGE', False);
      SetControlVisible('PFI_CURSUS', True);
      SetControlVisible('TPFI_CURSUS', True);
    end
    else
    begin
      TFFiche(Ecran).Caption := 'Inscription de ' + GetField('PFI_LIBELLE') + ' à la formation ' + RechDom('PGSTAGEFORM', GetField('PFI_CODESTAGE'), False);
      SetControlVisible('PFI_CODESTAGE', True);
      SetControlVisible('TPFI_CODESTAGE', True);
      SetControlVisible('PFI_CURSUS', False);
      SetControlVisible('TPFI_CURSUS', False);
    end;
    UpdateCaption(Ecran);
  end;
  if (Ecran is TFSaisieList) then MajAfficheCoutInsc(True, False);
  if TypeSaisie = 'VALIDRESPONSFOR' then
  begin
   Stage := TF.GetValue('PFI_CODESTAGE');
  //  Q := OpenSQL('SELECT PST_NATUREFORM FROM STAGE WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + MillesimeEC + '"', True);
  //  TobNat := Tob.Create('Naturestage',Nil,-1);
  //  TobNat.LoadDetailDB('naturestage','','',Q,False);
  //  Ferme(Q);
   // If TobNat.Detail.Count = 1 then Nature := TobNat.Detail[0].GetString('PST_NATUREFORM')
   // else Nature := '';
    TobNat := TobNature.FindFirst(['CODESTAGE'],[Stage],False);
    If TobNat <> Nil then Nature := TobNat.GetValue('NATURE')
    else Nature := '';
    TFSaisieList(Ecran).RichMessage1.Clear;
    TFSaisieList(Ecran).RichMessage1.Lines.Append('Infos formations');
    If Nature <> '' then TFSaisieList(Ecran).RichMessage1.Lines.Append('        Nature : ' + RechDom('PGNATUREFORM', Nature, False));
    for i := 1 to VH_Paie.NBFormationLibre do
    begin
      if i = 1 then Libelle := VH_Paie.FormationLibre1;
      if i = 2 then Libelle := VH_Paie.FormationLibre2;
      if i = 3 then Libelle := VH_Paie.FormationLibre3;
      if i = 4 then Libelle := VH_Paie.FormationLibre4;
      if i = 5 then Libelle := VH_Paie.FormationLibre5;
      if i = 6 then Libelle := VH_Paie.FormationLibre6;
      if i = 7 then Libelle := VH_Paie.FormationLibre7;
      if i = 8 then Libelle := VH_Paie.FormationLibre8;
      TFSaisieList(Ecran).RichMessage1.Lines.Append('        ' + Libelle + ' : ' + RechDom('PGFORMATION' + IntToStr(i), GetField('PFI_FORMATION' + IntToStr(i)), False));
    end;
  end;
  //AfficheInfosStage;
  If TypeSaisie = 'CURSUS' then
  begin
       SetControlEnabled('PFI_TYPEPLANPREV',False);
  end;
end;

procedure TOM_INSCFORMATION.OnAfterUpdateRecord;
var
  Q: TQuery;
  TobStage, T: Tob;
  Stage: string;
begin
  inherited;
  If MAilSaisie then EnvoiMailresponsable('SAISIE');
  If MailDecision then EnvoiMailresponsable('DECISION');
  if (TypeSaisie = 'SAISIESALARIE') or (TypeSaisie = 'GESTIONDIF') or (TypeSaisie = 'SAISIEVALID') or (TypeSaisie = 'CWASINSCBUDGET') or (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESTAGE') then
  begin
    Stage := GetField('PFI_CODESTAGE');
    CreerStagePrev(Stage);
  end;
  if (TypeSaisie = 'GESTIONDIF') or (TypeSaisie = 'SAISIEVALID') or (TypeSaisie = 'VALIDRESPONSFOR') or (TypeSaisie = 'VALIDRESPCURSUS') then
  begin
    if GetField('PFI_DATEACCEPT') <> IDate1900 then
    begin
      If (Ecran is TFSaisieList) then Stage := TF.GetValue('PFI_CODESTAGE')
      else Stage := GetField('PFI_CODESTAGE');
      Q := OpenSQL('SELECT * FROM STAGE WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + MillesimeEC + '"', True); //PT2
      TobStage := Tob.Create('STAGE', nil, -1);
      TobStage.LoadDetailDB('STAGE', '', '', Q, false);
      Ferme(Q);
      T := TobStage.FindFirst(['PST_CODESTAGE', 'PST_MILLESIME'], [Stage, MillesimeEC], False); //PT2
      if T <> nil then
      begin
        if T.Getvalue('PST_ACCEPTBUD') <> 'X' then
        begin
          T.PutValue('PST_ACCEPTBUD', 'X');
          T.UpdateDB(False);
        end;
      end;
      TobStage.Free;
    end;
  end;
  if TypeSaisie = 'CURSUS' then
  begin
    If NewCursus then
    begin
            CreerCursusPrev;
            InscriptionCursus;
    end;
  end;
  if (Ecran is TFSaisieList) then
     begin
     MajAfficheCoutInsc(True, True);
     {$IFNDEF EMANAGER} //PT17
     MajCoutPrev(GetField('PFI_CODESTAGE'),GetField('PFI_MILLESIME'));
     {$ENDIF}
     end;
end;

procedure TOM_INSCFORMATION.OnChangeField(F: TField);
var
  Q: TQuery;
  DureeStage, JourStage: Double;
  Etat : String;
  Acces : Boolean;
  wdatesortie : tdatetime; // pt30
  {$IFDEF EMANAGER}
  SQL : String;
  {$ENDIF}
begin
  inherited;
  {$IFDEF EMANAGER}
  If F.FieldName = 'PFI_NATUREFORM' then
  begin
       If GetField('PFI_NATUREFORM') = '002' then
       begin
            SetControlEnabled('PFI_ETATINSCFOR',False);
            SetControlEnabled('PFI_MOTIFETATINSC',False);
            SetControlEnabled('PFI_DATEACCEPT',False);
            If (DS.State in [DSInsert]) and (GetField('PFI_ETATINSCFOR') <> 'ATT') then SetField('PFI_ETATINSCFOR','ATT');
       end
       else
       begin
            SetControlEnabled('PFI_ETATINSCFOR',True);
            SetControlEnabled('PFI_MOTIFETATINSC',True);
            SetControlEnabled('PFI_DATEACCEPT',True);
       end;
  end;
  {$ENDIF}
  If F.FieldName = 'PFI_ETATINSCFOR' then
  begin
       Etat := GetField('PFI_ETATINSCFOR');
       Acces := false;
       If Etat = 'REF' then SetControlProperty('PFI_MOTIFETATINSC','DataType','PGMOTIFETATINSC3')
       else If Etat = 'REP' then SetControlProperty('PFI_MOTIFETATINSC','DataType','PGMOTIFETATINSC2')
       else If Etat = 'VAL' then SetControlProperty('PFI_MOTIFETATINSC','DataType','PGMOTIFETATINSC1')
       else
       begin
//            SetControlProperty('PFI_MOTIFETATINSC','DataType','');
            SetField('PFI_MOTIFETATINSC','');     //PT8
            Acces := True;
       end;
       SetControlEnabled('PFI_SALARIE',Acces);
       SetControlEnabled('PFI_DATEDIF',Acces);
       SetControlEnabled('PFI_MOTIFINSCFOR',Acces);
       SetControlEnabled('PFI_DUREESTAGE',Acces);
       SetControlEnabled('PFI_JOURSTAGE',Acces);
       SetControlEnabled('PFI_MOTIFETATINSC',not Acces);  //PT8
       SetControlEnabled('PFI_HTPSTRAV',Acces);
       SetControlEnabled('PFI_HTPSNONTRAV',Acces);
//       SetControlEnabled('PFI_DATEACCEPT',Acces);
       If (IsFieldModified('PFI_ETATINSCFOR')) and (GetField('PFI_ETATINSCFOR') <> 'ATT') then
       begin
            SetField('PFI_DATEACCEPT',V_PGI.DateEntree);
       end;
  end;
  If F.FieldName = 'PFI_TYPEPLANPREV' then
  begin
       If TypeDePlan <> GetField('PFI_TYPEPLANPREV') then
       begin
            AffectationHeures;
            TypeDePlan := GetField('PFI_TYPEPLANPREV');
       end;
  end;
  //PT17 - Début
  //PT26 On fait la même chose pour l'appli que pour eManager
  //  {$IFDEF EMANAGER}
  if (TypeSaisie='CURSUS') And (F.FieldName = 'PFI_NBINSC') And (GetField('PFI_SALARIE') = '') and (InitNbSal <> GetField('PFI_NBINSC')) then
  begin
      if GetField('PFI_NBINSC') > 1 then SetField('PFI_LIBELLE', IntToStr(GetField('PFI_NBINSC')) + ' salariés')
      else SetField('PFI_LIBELLE', IntToStr(GetField('PFI_NBINSC')) + ' salarié');
      InitNbSal := GetField('PFI_NBINSC');
  end;
  //  {$ENDIF}
  //PT17 - Fin
  if( (TypeSaisie = 'GESTIONDIF') or (TypeSaisie='CWASINSCBUDGET') or (TypeSaisie='SAISIESTAGE')) and (F.FieldName = 'PFI_NBINSC') and (IsFieldModified('PFI_NBINSC')) then
  begin
    Q := OpenSQL('SELECT PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE WHERE PST_CODESTAGE="' + Getfield('PFI_CODESTAGE') + '" AND PST_MILLESIME="' + GetField('PFI_MILLESIME') + '"',
      True);
    if not Q.Eof then
    begin
      DureeStage := Q.FindField('PST_DUREESTAGE').AsFloat;
      JourStage := Q.FindField('PST_JOURSTAGE').AsFloat;
//      Ferme(Q);
    end
    else
    begin
      Ferme(Q);
      Q := OpenSQL('SELECT PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE WHERE PST_CODESTAGE="' + Getfield('PFI_CODESTAGE') + '" AND PST_MILLESIME="0000"', True);
      if not Q.Eof then
      begin
        DureeStage := Q.FindField('PST_DUREESTAGE').AsFloat;
        JourStage := Q.FindField('PST_JOURSTAGE').AsFloat;
      end
      else
      begin
        DureeStage := 0;
        JourStage := 0;
      end;
//      Ferme(Q);
    end;
    Ferme(Q);
    if GetField('PFI_DUREESTAGE') <> (DureeStage * GetField('PFI_NBINSC')) then SetField('PFI_DUREESTAGE', DureeStage * GetField('PFI_NBINSC'));
    if GetField('PFI_JOURSTAGE') <> (JourStage * GetField('PFI_NBINSC')) then SetField('PFI_JOURSTAGE', JourStage * GetField('PFI_NBINSC'));
    if (GetField('PFI_SALARIE') = '') and (InitNbSal <> GetField('PFI_NBINSC')) then
    begin
      if GetField('PFI_NBINSC') > 1 then SetField('PFI_LIBELLE', IntToStr(GetField('PFI_NBINSC')) + ' salariés')
      else SetField('PFI_LIBELLE', IntToStr(GetField('PFI_NBINSC')) + ' salarié');
      InitNbSal := GetField('PFI_NBINSC');
    end;
  end;
  if ((TypeSaisie <> 'VALIDRESPONSFOR') or (TypeSaisie = 'VALIDRESPCURSUS')) then
  begin
    if (TypeSaisie='SAISIESALARIE') or (TypeSaisie = 'GESTIONDIF') or (TypeSaisie = 'CWASINSCBUDGET') or (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESTAGE') or (TypeSaisie = 'CURSUS') then
    begin
      if (F.FieldName = 'PFI_SALARIE') and (GetField('PFI_SALARIE') <> '') then
      begin
        SetControlEnabled('PFI_DADSCAT', False);
        //                                SetControlEnabled('PFI_NBINSC',False);

        // deb pt30 : si date de début exercice n'est pas renseignée: recherche de cette date
        if wdatedebexerc = 0 then begin
         Q := OpenSQL('SELECT PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" ORDER BY PFE_MILLESIME DESC', True);
         if not Q.Eof then wdatedebexerc := Q.FindField('PFE_DATEDEBUT').AsDateTime
         else wdatedebexerc := IDate1900;
        end;
        Ferme(Q);
        // fin pt30

        if (InitSalarie <> GetField('PFI_SALARIE')) Or ((TypeSaisie = 'SAISIESALARIE') AND (DS.State in [dsInsert]) )or ((ArgSalarie<>'') AND (DS.State in [dsInsert]) )then //PT18
        begin
          {$IFDEF EMANAGER}
          If TypeUtilisat = 'R' then SQL := 'SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE="' + GetField('PFI_SALARIE') + '" AND '+
          '(PSE_RESPONSFOR="'+Utilisateur+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))'
          else SQL := 'SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR IN '+
          '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+Utilisateur+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+Utilisateur+'"))) AND PSE_SALARIE="' + GetField('PFI_SALARIE') + '"';  //PT17
          if not ExisteSQL(SQL) then
          begin
            PGIBox('Vous n''êtes pas responsable de ce salarié, veuillez saisir un autre matricule', Ecran.Caption);
            SetFocusControl('PFI_SALARIE');
          end
          else
          begin
            {$ENDIF}
            //                                        SetField('PFI_NBINSC',1);
            SetField('PFI_LIBELLE', RechDom('PGSALARIE', GetField('PFI_SALARIE'), FALSE));
            If PGBundleHierarchie then
            begin
              SetField('PFI_LIBELLE', RechDom('PGSALARIEINT', GetField('PFI_SALARIE'), FALSE)); //PT22
            // pt30 Q := OpenSQL('SELECT PSI_NODOSSIER,PSI_TRAVAILN1,PSI_TRAVAILN2,PSI_TRAVAILN3,PSI_TRAVAILN4,PSI_CODESTAT,' +
                  Q := OpenSQL('SELECT PSI_NODOSSIER,PSI_TRAVAILN1,PSI_TRAVAILN2,PSI_TRAVAILN3,PSI_TRAVAILN4,PSI_CODESTAT,PSI_DATESORTIE,' +  // pt30
                  'PSI_LIBREPCMB1,PSI_LIBREPCMB2,PSI_LIBREPCMB3,PSI_LIBREPCMB4,PSI_ETABLISSEMENT,PSI_LIBELLEEMPLOI ' +
                  'FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="' + GetField('PFI_SALARIE') + '"', True);
              if not Q.eof then
              begin
                  SetField('PFI_ETABLISSEMENT', Q.FindField('PSI_ETABLISSEMENT').AsString);
                  SetField('PFI_LIBEMPLOIFOR', Q.FindField('PSI_LIBELLEEMPLOI').AsString);
                  SetField('PFI_CODESTAT', Q.FindField('PSI_CODESTAT').AsString);
                  SetField('PFI_TRAVAILN1', Q.FindField('PSI_TRAVAILN1').AsString);
                  SetField('PFI_TRAVAILN2', Q.FindField('PSI_TRAVAILN2').AsString);
                  SetField('PFI_TRAVAILN3', Q.FindField('PSI_TRAVAILN3').AsString);
                  SetField('PFI_TRAVAILN4', Q.FindField('PSI_TRAVAILN4').AsString);
                  SetField('PFI_LIBREPCMB1', Q.FindField('PSI_LIBREPCMB1').AsString);
                  SetField('PFI_LIBREPCMB2', Q.FindField('PSI_LIBREPCMB2').AsString);
                  SetField('PFI_LIBREPCMB3', Q.FindField('PSI_LIBREPCMB3').AsString);
                  SetField('PFI_LIBREPCMB4', Q.FindField('PSI_LIBREPCMB4').AsString);
//                  SetField('PFI_DADSCAT', Q.FindField('PSI_DADSCAT').AsString);
                  SetField('PFI_NODOSSIER', Q.FindField('PSI_NODOSSIER').AsString);
                  SetField('PFI_PREDEFINI', 'DOS');
                  Wdatesortie := Q.findfield('PSI_DATESORTIE').Asdatetime; // pt30
              end;
              Ferme(Q);
            end
            else
            begin
				Ferme(Q);
             //pt30   Q := OpenSQL('SELECT PSA_DADSCAT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,' +
                  Q := OpenSQL('SELECT PSA_DADSCAT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_DATESORTIE,' +   // pt30
                  'PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI ' +
                  'FROM SALARIES WHERE PSA_SALARIE="' + GetField('PFI_SALARIE') + '"', True);
                if not Q.eof then
                begin
                  SetField('PFI_ETABLISSEMENT', Q.FindField('PSA_ETABLISSEMENT').AsString);
                  SetField('PFI_LIBEMPLOIFOR', Q.FindField('PSA_LIBELLEEMPLOI').AsString);
                  SetField('PFI_CODESTAT', Q.FindField('PSA_CODESTAT').AsString);
                  SetField('PFI_TRAVAILN1', Q.FindField('PSA_TRAVAILN1').AsString);
                  SetField('PFI_TRAVAILN2', Q.FindField('PSA_TRAVAILN2').AsString);
                  SetField('PFI_TRAVAILN3', Q.FindField('PSA_TRAVAILN3').AsString);
                  SetField('PFI_TRAVAILN4', Q.FindField('PSA_TRAVAILN4').AsString);
                  SetField('PFI_LIBREPCMB1', Q.FindField('PSA_LIBREPCMB1').AsString);
                  SetField('PFI_LIBREPCMB2', Q.FindField('PSA_LIBREPCMB2').AsString);
                  SetField('PFI_LIBREPCMB3', Q.FindField('PSA_LIBREPCMB3').AsString);
                  SetField('PFI_LIBREPCMB4', Q.FindField('PSA_LIBREPCMB4').AsString);
                  SetField('PFI_DADSCAT', Q.FindField('PSA_DADSCAT').AsString);
                  SetField('PFI_NODOSSIER', '000000');
                  SetField('PFI_PREDEFINI', 'DOS');
                  Wdatesortie := Q.findfield('PSA_DATESORTIE').asdatetime; // pt30
                end;
                Ferme(Q);
            end;
            // deb pt30 controle date debut exercice >= date de sortie
            If ((wdatesortie > 10) and  (wdatesortie <= wdatedebexerc))  then
               Pgibox('La date de début de l''exercice est supérieure à la date de sortie',Ecran.Caption);
            // fin pt30
            
            InitSalarie := GetField('PFI_SALARIE');
            {$IFDEF EMANAGER}
          end;
          {$ENDIF}
        end;
        SetControlEnabled('PFI_ETABLISSEMENT', False);
        SetControlEnabled('PFI_LIBEMPLOIFOR', False);
      end;
      if (F.FieldName = 'PFI_SALARIE') and (GetField('PFI_SALARIE') = '') then
      begin
        SetControlEnabled('PFI_ETABLISSEMENT', True);
        SetControlEnabled('PFI_LIBEMPLOIFOR', True);
        //                                SetControlEnabled('PFI_NBINSC',true);
        SetControlEnabled('PFI_DADSCAT', True);
        if InitNbSal <> GetField('PFI_NBINSC') then
        begin
          if GetField('PFI_NBINSC') > 1 then SetField('PFI_LIBELLE', IntToStr(GetField('PFI_NBINSC')) + ' salariés')
          else SetField('PFI_LIBELLE', IntToStr(GetField('PFI_NBINSC')) + ' salarié');
          InitNbSal := GetField('PFI_NBINSC');
        end;
        If PGBundleInscFormation Then SetField('PFI_NODOSSIER', V_PGI.NoDossier); //PT26
      end;
    end;
  end;
  if (TypeSaisie = 'GESTIONDIF') or (TypeSaisie = 'SAISIEVALID') or (TypeSaisie = 'VALIDRESPONSFOR') or (TypeSaisie = 'VALIDATION') or (TypeSaisie = 'VALIDRESPCURSUS') then // PT1
  begin
    if F.FieldName = 'PFI_ETATINSCFOR' then
    begin
      if GetField('PFI_ETATINSCFOR') = 'ATT' then
      begin
        SetControlEnabled('PFI_MOTIFETATINSC', False);
      end;
      if GetField('PFI_ETATINSCFOR') = 'REF' then
      begin
        SetControlEnabled('PFI_MOTIFETATINSC', True);
        SetControlProperty('PFI_MOTIFETATINSC', 'Plus', ' AND CC_LIBRE="REF"');
      end;
      if GetField('PFI_ETATINSCFOR') = 'VAL' then
      begin
        SetControlEnabled('PFI_MOTIFETATINSC', True);
        SetControlProperty('PFI_MOTIFETATINSC', 'Plus', ' AND CC_LIBRE="VAL"');
      end;
      if GetField('PFI_ETATINSCFOR') = 'REP' then
      begin
        SetControlEnabled('PFI_MOTIFETATINSC', True);
        SetControlProperty('PFI_MOTIFETATINSC', 'Plus', ' AND CC_LIBRE="REP"');
      end;
    end;
  end;
    if (TypeSaisie = 'GESTIONDIF') and (F.FieldName = 'PFI_CODESTAGE') and (IsFieldModified('PFI_CODESTAGE')) then
  begin
    RecupInfosStage;
    AffectationHeures;
  end;
  //PT18 - Début
  If ((TypeSaisie = 'CWASINSCBUDGET') and (ARgStage='') and (DS.State in [dsInsert]) ) then
  begin
    if F.FieldName = 'PFI_CODESTAGE' then
    begin
      If GetField('PFI_CODESTAGE') <> '' then
      begin
         RecupInfosStage;
      end;
    end;
  end;
  //PT18 - Fin
end;


procedure TOM_INSCFORMATION.OnArgument(S: string);
var
  Q: TQuery;
  BSelect, BAccept,BZoom,BCompteurs : TToolBarButton97;
//  THDate, THSal: THEdit;
  Num : Integer;
  {$IFNDEF EAGLCLIENT}
  Edit   : THDBEdit;
  Edit2  : THEdit;
  {$ELSE}
  Edit : THEdit;
  {$ENDIF}
  DD,DF : TdateTime;
begin
  inherited;
  wdatedebexerc := 0; // pt30 initialise la date de début d'exercice

  InscriptionsStd := False;
  If Ecran Is TFFiche then TFFiche(Ecran).Retour := 'RIEN';
  DefautHTpsT := GetParamSoc('SO_PGDIFTPSTRAV');
  Utilisateur := V_PGI.UserSalarie;

  {$IFDEF EMANAGER}
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+Utilisateur+'"') then TypeUtilisat := 'R'  //PT17
  else TypeUtilisat := 'S';
  {$ENDIF}                   

  TypeSaisie := ReadTokenSt(S);
  SalarieDIF := '';
  if TypeSaisie = 'CURSUS' then
  begin
    LeCursus := Trim(ReadTokenPipe(S, ';'));
    TypeInsc := '';
    ArgStage := '';
    ArgMillesime := Trim(ReadTokenPipe(S, ';'));
    SetControlEnabled('PFI_TYPEPLANPREV',False);
    If ExisteSQL('SELECT PCU_CURSUS FROM CURSUS WHERE PCU_PREDEFINI="STD" AND PCU_CURSUS="'+LeCursus+'"') then InscriptionsSTD := True; 

  end
  else
  begin
       If TypeSaisie = 'GESTIONDIF' then
       begin
            TypeInsc := ReadTokenSt(S);
            SalarieDIF := Trim(ReadTokenPipe(S, ';'));
       end
       else
       begin
            TypeInsc := ReadTokenSt(S);
            ArgStage := Trim(ReadTokenPipe(S, ';'));
            ArgMillesime := Trim(ReadTokenPipe(S, ';'));
            ArgSalarie := Trim(ReadTokenPipe(S, ';'));  //PT18
       end;
  end;

  If TypeSaisie = 'SAISIEDIF' then setControlEnabled('PFI_TYPEPLANPREV',False);
  If not (Ecran is TFSaisieList) then
  begin
         {$IFNDEF EAGLCLIENT}
         Edit := THDBEdit(GetControl('PFI_CODESTAGE'));
         {$ELSE}
         Edit := THEdit(GetControl('PFI_CODESTAGE'));
         {$ENDIF}
         If Edit <> Nil then Edit.OnElipsisClick := StageElipsisClick;
         BCompteurs := TToolBarButton97(GetControl('BCOMPTEURS'));
         If BCompteurs <> Nil Then BCompteurs.OnClick := VoirCompteursDIF;
         For Num  :=  1 to VH_Paie.PGNbreStatOrg do
         begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFI_TRAVAILN'+IntToStr(Num)),GetControl ('TPFI_TRAVAILN'+IntToStr(Num)));
                VisibiliteChampLibreSal(IntToStr(Num),GetControl ('PFI_LIBREPCMB'+IntToStr(Num)),GetControl ('TPFI_LIBREPCMB'+IntToStr(Num)));
                SetControlEnabled('PFI_TRAVAILN'+IntToStr(Num),False);
                SetControlEnabled('PFI_LIBREPCMB'+IntToStr(Num),False);
         end;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PFI_FORMATION'+IntToStr(Num)),GetControl ('TPFI_FORMATION'+IntToStr(Num)));
                SetControlEnabled('PFI_FORMATION'+IntToStr(Num),False);
        end;
        VisibiliteStat(GetControl ('PFI_CODESTAT'),GetControl ('TPFI_CODESTAT'));
        SetControlEnabled('PFI_CODESTAT',False);
  end;

  If TypeSaisie = 'GESTIONDIF' then
  begin
       BZoom := TToolBarButton97(GetControl('BZOOM'));
       If BZoom <> Nil then BZoom.OnClick := BZoomClick;
       SetControlEnabled('PFI_TYPEPLANPREV',False);
  end
  else
    SetControlVisible ('BZOOM', False);

  if (ArgMillesime = '') or (ArgMillesime = '0000') then
  begin
    MillesimeEC := '';
    {$IFNDEF EMANAGER} //PT17
    If (TypeSaisie = 'GESTIONDIF') then MillesimeEC := RendMillesimeRealise(DD,DF)
    else MillesimeEC := RendMillesimePrevisionnel;
    {$ELSE}
    millesimeEC := RendMillesimeEManager;
    {$ENDIF}
  end
  else MillesimeEC := ArgMillesime;

  //DEBUT PT5
  Q := OpenSQL('SELECT PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="' + MillesimeEC + '"', True);
  if not Q.Eof then TauxChargeBudget := Q.FindField('PFE_TAUXBUDGET').AsFloat
  else TauxChargeBudget := 1;
  Ferme(Q);
  //FIN PT5

  if (Ecran <> nil) and (Ecran is TFSaisieList) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    {$IFDEF EMANAGER}
    SetControlVisible('BParamListe', False);
    {$ENDIF}
    if TypeSaisie <> 'CURSUS' then TF.OnSetNavigate := AfficheInfosStage;
    AfficherLibelle := True;
    if (TypeSaisie <> 'CURSUS') and (TypeSaisie <> 'VALIDRESPCURSUS') then
    begin
      If (ArgStage <> '') or (ArgSalarie <> '' ) then  //PT1
      begin
      SetControlVisible('PanPied', True);
      SetControlVisible('PCPied', True);
      end;
    end
    else
    begin
      SetControlVisible('PanPied', False);
      SetControlVisible('PCPied', False);
    end;
    if (TypeSaisie = 'VALIDRESPONSFOR') or (TypeSaisie = 'VALIDRESPCURSUS') then
    begin
      SetControlVisible('BACCEPTATION', True);
      TFSaisieList(Ecran).Caption := 'Validation des inscriptions au budget de l''année ' + MillesimeEC;
      TF.LaGrid.MultiSelect := true;
    end;
    if TypeSaisie = 'SAISIEINSC' then // Saisie des inscriptions
    begin
      SetControlProperty('PCPIED', 'ActivePage', 'PFORMATION');
      SetControlVisible('BACCEPTATION', False);
      TFSaisieList(Ecran).Caption := 'Saisie des inscriptions au budget de l''année ' + MillesimeEC;
    end;
    if TypeSaisie = 'CURSUS' then // Saisie des cursus
    begin
      TFSaisieList(Ecran).Caption := 'Saisie des inscriptions aux cursus de l''année ' + MillesimeEC;
      {$IFDEF EMANAGER}
      Utilisateur := ReadTokenSt(S);
      If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+Utilisateur+'"') then TypeUtilisat := 'R'
      else TypeUtilisat := 'S';
      {$ENDIF}
    end;
    if TypeSaisie = 'SAISIESTAGE' then // Saisie des inscriptions
    begin
      SetControlVisible('BACCEPTATION', False);
      TFSaisieList(Ecran).Caption := 'Saisie des inscriptions au budget de l''année ' + MillesimeEC;
    end;
    if TypeSaisie = 'CWASINSCBUDGET' then // Saisie des inscriptions par responsable
    begin
      SetControlProperty('PFI_SALARIE', 'DataType', '');
      SetControlVisible('BACCEPTATION', False);
      TFSaisieList(Ecran).Caption := 'Saisie des inscriptions au budget de l''année ' + MillesimeEC;
      {$IFDEF EMANAGER}
      ReadTokenSt(S);
      ReadTokenSt(S);
      Utilisateur := ReadTokenSt(S);
      If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+Utilisateur+'"') then TypeUtilisat := 'R'
      else TypeUtilisat := 'S';
      {$ENDIF}
    end;
    if TypeSaisie = 'VALIDINSC' then // Validation des inscriptions
    begin
      SetControlVisible('BACCEPTATION', True);
      TFSaisieList(Ecran).Caption := 'Validation des inscriptions au budget de l''année ' + MillesimeEC;
      TF.LaGrid.MultiSelect := true;
    end;
    if TypeSaisie = 'VALIDWEBINSC' then // Validation des inscriptions par responsable
    begin
      SetControlVisible('BACCEPTATION', True);
      TFSaisieList(Ecran).Caption := 'Validation des inscriptions au budget de l''année ' + MillesimeEC;
      TF.LaGrid.MultiSelect := true;
    end;
    if TypeSaisie = 'CONSULTATION' then TFSaisieList(Ecran).FTypeAction := TaConsult;
    if TypeSaisie = 'CONSULTATION' then TFSaisieList(Ecran).Caption := 'Consultation des inscriptions au budget prévisionnel de l''année ' + MillesimeEC;
    if TypeSaisie = 'SAISIEVALID' then TFSaisieList(Ecran).Caption := 'Validation des inscriptions au budget prévisionnel de l''année ' + MillesimeEC;
    if (TypeSaisie = 'CWASINSCBUDGET') or (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESTAGE') then
      TFSaisieList(Ecran).Caption := 'Saisie des inscriptions au budget prévisionnel de l''année ' + MillesimeEC;
    UpdateCaption(TFSaisieList(Ecran));
    if (TypeSaisie = 'SAISIEVALID') or (TypeSaisie = 'VALIDRESPONSFOR') or (TypeSaisie = 'VALIDRESPCURSUS') then
    begin
      TF.LaGrid.MultiSelect := true;
    end
    else SetControlVisible('BACCEPTATION', False);
    SetControlVisible('Page1', False);
    SetControlProperty('Page1', 'TabVisible', False);
    SetControlVisible('Page2', False);
    SetControlProperty('Page2', 'TabVisible', False);
    SetControlVisible('Page3', False);
    SetControlProperty('Page3', 'TabVisible', False);
    if (TypeSaisie = 'VALIDRESPONSFOR') or (TypeSaisie = 'VALIDRESPCURSUS') then
    begin
      BAccept := TToolBarButton97(GetControl('BACCEPTATION'));
      if BAccept <> nil then BAccept.OnClick := BAcceptClick;
    end;
{    if TypeSaisie = 'SAISIESTAGE' then // Saisie des inscriptions
    begin
      THSal := THEdit(GetControl('PFI_SALARIE'));
      if THSal <> nil then THSal.OnElipsisClick := SalarieElipsisClick;
    end;
    if TypeSaisie = 'CWASINSCBUDGET' then // Saisie des inscriptions par responsable
    begin
      THSal := THEdit(GetControl('PFI_SALARIE'));
      if THSal <> nil then THSal.OnElipsisClick := SalarieElipsisClick;
    end;}
     if TypeSaisie = 'VALIDINSC' then // Validation des inscriptions
    begin
      BAccept := TToolBarButton97(GetControl('BACCEPTATION'));
      if BAccept <> nil then BAccept.OnClick := BAcceptClick;
    end;
    if TypeSaisie = 'VALIDWEBINSC' then // Validation des inscriptions par responsable
    begin
      BAccept := TToolBarButton97(GetControl('BACCEPTATION'));
      if BAccept <> nil then BAccept.OnClick := BAcceptClick;
    end;
    if (TypeSaisie = 'SAISIEVALID') or (TypeSaisie = 'VALIDRESPONSFOR') or (TypeSaisie = 'VALIDRESPCURSUS') then
    begin
      BAccept := TToolBarButton97(GetControl('BACCEPTATION'));
      if BAccept <> nil then BAccept.OnClick := BAcceptClick;
      BSelect := TToolBarButton97(GetControl('BSELECTALL'));
      if BSelect <> nil then BSelect.OnClick := BSelectionClick;
    end
    else SetControlVisible('BACCEPTATION', False);
    THVMillesime := THValComboBox(GetControl('THVMILLESIME'));

    Edit := {$IFDEF EAGLCLIENT}THEdit{$ELSE}THDBEdit{$ENDIF}(GetControl('PFI_DATEACCEPT'));
    if Edit <> nil then Edit.OnElipsisClick := DateElipsisClick;
    Edit := {$IFDEF EAGLCLIENT}THEdit{$ELSE}THDBEdit{$ENDIF}(GetControl('PFI_REPORTELE'));
    if Edit <> nil then Edit.OnElipsisClick := DateElipsisClick;
    Edit := {$IFDEF EAGLCLIENT}THEdit{$ELSE}THDBEdit{$ENDIF}(GetControl('PFI_REFUSELE'));
    if Edit <> nil then Edit.OnElipsisClick := DateElipsisClick;
  end
  else
  begin // DEBUT PT1
    SetControlEnabled('PFI_RESPONSFOR', False);
    SetcontrolEnabled('PFI_LIBELLE', False);
    SetControlEnabled('PFI_FRAISFORFAIT', False);
    SetControlEnabled('PFI_COUTREELSAL', False);
    SetControlEnabled('PFI_AUTRECOUT', False);
    SetControlEnabled('PFI_CONFIDENTIEL', False);
    if TFFiche(Ecran).TYpeAction = TaConsult then SetControlVisible('BDefaire', False);
  end; // FIN PT1

  {$IFDEF EMANAGER}
  SetControlEnabled('PFI_REALISE',False);
  {$ENDIF}

  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PFI_SALARIE'));
  {$ELSE}
  Edit  := Nil;
  Edit2 := Nil;
  If (GetControl('PFI_SALARIE') Is THEdit) Then
    Edit2 := THEdit(GetControl('PFI_SALARIE'))
  Else
    Edit  := THDBEdit(GetControl('PFI_SALARIE'));
  {$ENDIF}
  if Edit <> nil Then Edit.OnElipsisClick := SalarieElipsisClick
  {$IFNDEF EAGLCLIENT}
  Else If Edit2 <> Nil Then Edit2.OnElipsisClick := SalarieElipsisClick
  {$ENDIF};

  //PT17 - Début
  {$IFDEF EMANAGER}
  If (ArgStage<>'') And (ExisteSQL('SELECT 1 FROM CURSUSSTAGE WHERE PCC_CURSUS<>"---" AND PCC_CURSUS<>"-C-" AND PCC_CODESTAGE="'+ArgStage+'"')) Then
     PGIInfo('Ce stage fait partie d''un cursus.',Ecran.Caption);
  {$ENDIF}
  //PT17 - Fin
end;

procedure TOM_INSCFORMATION.BAcceptClick(Sender: TObject);
var
  i: integer;
begin
  If TF.Reccount = 0 then exit;
  if ((TF.LaGrid.nbSelected) > 0) and (not TF.LaGrid.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TF.LaGrid.nbSelected, FALSE, TRUE);
    InitMove(TF.LaGrid.NbSelected, '');
    TF.DisableTom;
    TF.StartUpdate;
    for i := 0 to TF.LaGrid.NbSelected - 1 do
    begin
      TF.LaGrid.GotoLeBOOKMARK(i);
      TF.SelectRecord(TF.LaGrid.Row);
      MoveCur(False);
      If TF.GetValue('PFI_CURSUS') <> '' then MajEtatFormCursus(TF.GetValue('PFI_CURSUS'),TF.GetValue('PFI_ETABLISSEMENT'),TF.GetValue('PFI_MILLESIME'),TF.GetValue('PFI_ETATINSCFOR'),TF.GetValue('PFI_MOTIFINSCFOR'),TF.GetValue('PFI_MOTIFDECISION'),TF.GetValue('PFI_RANG'));
      TF.PutValue('PFI_DATEACCEPT', Date);
      TF.PutValue('PFI_ETATINSCFOR', 'VAL');
      TF.Post;
      MoveCurProgressForm(TF.GetValue('PFI_LIBELLE'));
    end;
    TF.EndUpdate;
    TF.EnableTom;
    FiniMove;
    FiniMoveProgressForm;
  end;
  if TF.LaGrid.AllSelected then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TF.LaGrid.RowCount, FALSE, TRUE);
    InitMove(TF.LaGrid.NbSelected, '');
    TF.DisableTom;
    TF.StartUpdate;
    for i := 1 to TF.LaGrid.RowCount do
    begin
      TF.SelectRecord(i);
      MoveCur(False);
      If TF.GetValue('PFI_CURSUS') <> '' then MajEtatFormCursus(TF.GetValue('PFI_CURSUS'),TF.GetValue('PFI_ETABLISSEMENT'),TF.GetValue('PFI_MILLESIME'),TF.GetValue('PFI_ETATINSCFOR'),TF.GetValue('PFI_MOTIFINSCFOR'),TF.GetValue('PFI_MOTIFDECISION'),TF.GetValue('PFI_RANG'));
      TF.PutValue('PFI_DATEACCEPT', Date);
      TF.PutValue('PFI_ETATINSCFOR', 'VAL');
      TF.Post;
      MoveCurProgressForm(TF.GetValue('PFI_LIBELLE'));
    end;
    TF.EndUpdate;
    TF.EnableTom;
    FiniMove;
    FiniMoveProgressForm;
  end;
  TF.LaGrid.ClearSelected;
  TF.LaGrid.AllSelected := False;
  SetControlProperty('BSELECTALL', 'Down', False);
end;

procedure TOM_INSCFORMATION.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOM_INSCFORMATION.AfficheInfosStage(Sender: TObject);
var
  Q: TQuery;
  Num: Integer;
  HMTrad: THSystemMenu;
begin
  Q := OpenSQL('SELECT PST_PREDEFINI,PST_LIBELLE,PST_LIBELLE1,PST_LIEUFORM,PST_NATUREFORM,PST_TYPCONVFORM,PST_COUTBUDGETE,PST_FORMATION1,' +
    'PST_NBSTAMAX,PST_COUTUNITAIRE,PST_NBSTAMIN,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,' +
    'PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE WHERE PST_CODESTAGE="' + ArgStage + '" AND PST_MILLESIME="' + ArgMillesime + '"', True);
  if not Q.eof then
  begin
    Setcontroltext('STAGE', Q.FindField('PST_LIBELLE').AsString + ' ' + Q.FindField('PST_LIBELLE1').AsString);
    SetControlText('LIEUFORM', Q.FindField('PST_LIEUFORM').AsString);
    SetControlText('NATUREFORM', Q.FindField('PST_NATUREFORM').AsString);
    Setcontroltext('DUREESTAGE', FloatToStr(Q.FindField('PST_DUREESTAGE').AsFloat));
    Setcontroltext('JOURSTAGE', FloatToStr(Q.FindField('PST_JOURSTAGE').AsFloat));
    Setcontroltext('NBMIN', IntToStr(Q.FindField('PST_NBSTAMIN').AsInteger));
    Setcontroltext('NBMAX', IntToStr(Q.FindField('PST_NBSTAMAX').AsInteger));
    Setcontroltext('COUTGLOBAL', FloatToStr(Q.FindField('PST_COUTBUDGETE').AsFloat));
    Setcontroltext('COUTFORFAITAIRE', FloatToStr(Q.FindField('PST_COUTUNITAIRE').AsFloat));
    Setcontroltext('FORMATION1', Q.FindField('PST_FORMATION1').AsString);
    Setcontroltext('FORMATION2', Q.FindField('PST_FORMATION2').AsString);
    Setcontroltext('FORMATION3', Q.FindField('PST_FORMATION3').AsString);
    Setcontroltext('FORMATION4', Q.FindField('PST_FORMATION4').AsString);
    Setcontroltext('FORMATION5', Q.FindField('PST_FORMATION5').AsString);
    Setcontroltext('FORMATION6', Q.FindField('PST_FORMATION6').AsString);
    Setcontroltext('FORMATION7', Q.FindField('PST_FORMATION7').AsString);
    Setcontroltext('FORMATION8', Q.FindField('PST_FORMATION8').AsString);
    If Q.FindField('PST_PREDEFINI').AsString = 'STD' then InscriptionsStd := True;
  end;
  Ferme(Q);
  SetControlVisible('PFORMATIONLIBERE', True);
  SetControlProperty('PFORMATIONLIBRE', 'TabVisible', True);

  for Num := 1 to VH_Paie.NBFormationLibre do
  begin
    if Num > 8 then Break;
    VisibiliteChampFormation(IntToStr(Num), GetControl('FORMATION' + IntToStr(Num)), GetControl('TFORMATION' + IntToStr(Num)));
    SetControlEnabled('FORMATION' + IntToStr(Num), False);
  end;
  HMTrad := THSystemMenu(GetControl('HMTRAD'));
  HMTrad.ResizeGridColumns(TF.LaGrid);
  if (Ecran is TFSaisieList) then MajAfficheCoutInsc(False, True);
end;

procedure TOM_INSCFORMATION.SalarieElipsisClick(Sender: TObject);
var
  StFrom, StWhere: string;
  Q: TQuery;
  DD: TDateTime;
begin
  Q := OpenSQL('SELECT PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" ORDER BY PFE_MILLESIME DESC', True);
  if not Q.Eof then DD := Q.FindField('PFE_DATEDEBUT').AsDateTime
  else DD := IDate1900;
  // pt30 mémorise la date de début d'exercice
  wdatedebexerc := DD;  // pt30

  Ferme(Q);
  If PGBundleHierarchie then
  begin
    StWhere := '(PSI_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSI_DATESORTIE>="' + UsDateTime(DD) + '")';
//    StFrom := 'INTERIMAIRES';
//    If (not InscriptionsStd) or (Not PGDroitMultiForm) then StWhere := StWhere+' AND PSI_NODOSSIER="'+V_PGI.NoDossier+'"';
  end
  else
  begin
     StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE>="' + UsDateTime(DD) + '")';
     StFrom := 'SALARIES';
  end;
  {$IFDEF EMANAGER}
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  //PT19 - Début
  StWhere := StWhere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
//  If TypeUtilisat = 'R' then
//  begin
//       If THEdit(GetControl('CHIERARCHIE')) <> Nil then
//       begin
//            If GetControlText('CHIERARCHIE') = 'DIRECT' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+Utilisateur+'"'
//            else If GetControlText('CHIERARCHIE') = 'INDIRECT' then StWhere := StWhere + ' AND PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
//       ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'"))'
//            else StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+Utilisateur+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
//       ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))';
//       end
//       else StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+Utilisateur+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
//       ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))';
//  end
//  else StWhere := StWhere + ' AND PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
//     'WHERE (PGS_SECRETAIREFOR="'+Utilisateur+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+Utilisateur+'")))';
  //PT19 - Fin
  {$ENDIF}
  StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT12
  If PGBundleHierarchie then
//  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSI_INTERIMAIRE', 'PSI_LIBELLE,PSI_PRENOM,PSI_NODOSSIER', StWhere, 'PSI_INTERIMAIRE', TRUE, -1)
    ElipsisSalarieMultidos(Sender, StWhere)
  else LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
end;

procedure TOM_INSCFORMATION.BSelectionClick(Sender: TObject);
var
  BSelect: TToolBarButton97;
begin
  BSelect := TToolBarButton97(GetControl('BSELECTALL'));
  if BSelect.Down then TF.LaGrid.AllSelected := True
  else TF.LaGrid.AllSelected := False;
end;

procedure TOM_INSCFORMATION.InscriptionCursus;
var
  Q: TQuery;
  TobLesFormations, T: Tob;
  i, NbInsc: Integer;
  CoutSal, TotalFrais,NbJourCur, CoutSalCur, TotalFraisCur, NbHCur,CtPedagCursus,CtPedag: Double;
  FraisH, FraisR, FraisT, Salaire, NbHeure, NbJour: Double;
  Req : String; //PT15
  TobEcran : TOB; //PT15
begin
  TF.SelectRecord(TF.LaGrid.Row);
  Q := OpenSQL('SELECT PCC_CODESTAGE,PST_JOURSTAGE,PST_LIEUFORM,PST_DUREESTAGE,' +
    'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_NATUREFORM ' + //PT18
    'FROM CURSUSSTAGE LEFT JOIN STAGE ON PCC_CODESTAGE=PST_CODESTAGE AND PCC_MILLESIME=PST_MILLESIME WHERE PCC_CURSUS="' + LeCursus +
      '" AND PCC_RANGCURSUS=0 AND PCC_MILLESIME="'+MillesimeEC+'"', True); //DB2
  TobLesFormations := Tob.Create('LesFormations', nil, -1);
  TobLesFormations.LoadDetailDB('LesFormations', '', '', Q, False);
  Ferme(Q);
  CoutSalCur := 0;
  TotalFraisCur := 0;
  NbHCur := 0;
  NbJourCur := 0;
  CtPedagCursus := 0;
  for i := 0 to TobLesFormations.Detail.Count - 1 do
  begin
    T := Tob.Create('INSCFORMATION', nil, -1);
    T.PutValue('PFI_SALARIE',TF.GetValue('PFI_SALARIE'));
    T.PutValue('PFI_ETABLISSEMENT',TF.GetValue('PFI_ETABLISSEMENT'));
    T.PutValue('PFI_TYPEPLANPREV',TF.GetValue('PFI_TYPEPLANPREV'));
    T.PutValue('PFI_RANG',TF.GetValue('PFI_RANG'));
    T.PutValue('PFI_MILLESIME',TF.GetValue('PFI_MILLESIME'));
    T.PutValue('PFI_LIBELLE',TF.GetValue('PFI_LIBELLE'));
    T.PutValue('PFI_CONFIDENTIEL',TF.GetValue('PFI_CONFIDENTIEL'));
    T.PutValue('PFI_DATEACCEPT',TF.GetValue('PFI_DATEACCEPT'));
    T.PutValue('PFI_ACCEPTPAR',TF.GetValue('PFI_ACCEPTPAR'));
    T.PutValue('PFI_ETATINSCFOR',TF.GetValue('PFI_ETATINSCFOR'));
    T.PutValue('PFI_MOTIFINSCFOR',TF.GetValue('PFI_MOTIFINSCFOR'));
    T.PutValue('PFI_MOTIFETATINSC',TF.GetValue('PFI_MOTIFETATINSC'));
    T.PutValue('PFI_NIVPRIORITE',TF.GetValue('PFI_NIVPRIORITE'));
    T.PutValue('PFI_REFUSELE',TF.GetValue('PFI_REFUSELE'));
    T.PutValue('PFI_REFUSPAR',TF.GetValue('PFI_REFUSPAR'));
    T.PutValue('PFI_REFUSFORM',TF.GetValue('PFI_REFUSFORM'));
    T.PutValue('PFI_REPORTELE',TF.GetValue('PFI_REPORTELE'));
    T.PutValue('PFI_REPORTERPAR',TF.GetValue('PFI_REPORTERPAR'));
    T.PutValue('PFI_REPORTFORM',TF.GetValue('PFI_REPORTFORM'));
    T.PutValue('PFI_BUDGETE',TF.GetValue('PFI_BUDGETE'));
    T.PutValue('PFI_CURSUS',TF.GetValue('PFI_CURSUS'));
    T.PutValue('PFI_TRAITEMENT',TF.GetValue('PFI_TRAITEMENT'));
    T.PutValue('PFI_RESPONSFOR',TF.GetValue('PFI_RESPONSFOR'));
    T.PutValue('PFI_NBINSC',TF.GetValue('PFI_NBINSC'));
    T.PutValue('PFI_LIBEMPLOIFOR',TF.GetValue('PFI_LIBEMPLOIFOR'));
    T.PutValue('PFI_METIERFOR',TF.GetValue('PFI_METIERFOR'));
    T.PutValue('PFI_TRAVAILN1',TF.GetValue('PFI_TRAVAILN1'));
    T.PutValue('PFI_TRAVAILN2',TF.GetValue('PFI_TRAVAILN2'));
    T.PutValue('PFI_TRAVAILN3',TF.GetValue('PFI_TRAVAILN3'));
    T.PutValue('PFI_TRAVAILN4',TF.GetValue('PFI_TRAVAILN4'));
    T.PutValue('PFI_CODESTAT',TF.GetValue('PFI_CODESTAT'));
    T.PutValue('PFI_LIBREPCMB1',TF.GetValue('PFI_LIBREPCMB1'));
    T.PutValue('PFI_LIBREPCMB2',TF.GetValue('PFI_LIBREPCMB2'));
    T.PutValue('PFI_LIBREPCMB3',TF.GetValue('PFI_LIBREPCMB3'));
    T.PutValue('PFI_LIBREPCMB4',TF.GetValue('PFI_LIBREPCMB4'));
    T.PutValue('PFI_DADSCAT',TF.GetValue('PFI_DADSCAT'));
    T.PutValue('PFI_REALISE',TF.GetValue('PFI_REALISE'));
    
  	T.PutValue('PFI_PREDEFINI', TF.GetValue('PFI_PREDEFINI'));
  	T.PutValue('PFI_NODOSSIER', TF.GetValue('PFI_NODOSSIER'));
    
    T.PutValue('PFI_CODESTAGE', TobLesFormations.Detail[i].GetValue('PCC_CODESTAGE'));
    T.PutValue('PFI_JOURSTAGE', TobLesFormations.Detail[i].GetValue('PST_JOURSTAGE'));
    T.PutValue('PFI_DUREESTAGE', TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE'));
    if (DefautHTpsT) or (GetField('PFI_TYPEPLANPREV') = 'PLF') then
    begin
        T.PutValue('PFI_HTPSTRAV', TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE'));
        T.PutValue('PFI_HTPSNONTRAV', 0);
    end
    else
    begin
        T.PutValue('PFI_HTPSTRAV', 0);
        T.PutValue('PFI_HTPSNONTRAV', TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE'));
    end;
    T.PutValue('PFI_LIEUFORM', TobLesFormations.Detail[i].GetValue('PST_LIEUFORM'));
    T.PutValue('PFI_FORMATION1', TobLesFormations.Detail[i].GetValue('PST_FORMATION1'));
    T.PutValue('PFI_FORMATION2', TobLesFormations.Detail[i].GetValue('PST_FORMATION2'));
    T.PutValue('PFI_FORMATION3', TobLesFormations.Detail[i].GetValue('PST_FORMATION3'));
    T.PutValue('PFI_FORMATION4', TobLesFormations.Detail[i].GetValue('PST_FORMATION4'));
    T.PutValue('PFI_FORMATION5', TobLesFormations.Detail[i].GetValue('PST_FORMATION5'));
    T.PutValue('PFI_FORMATION6', TobLesFormations.Detail[i].GetValue('PST_FORMATION6'));
    T.PutValue('PFI_FORMATION7', TobLesFormations.Detail[i].GetValue('PST_FORMATION7'));
    T.PutValue('PFI_FORMATION8', TobLesFormations.Detail[i].GetValue('PST_FORMATION8'));
    T.PutValue('PFI_NATUREFORM', TobLesFormations.Detail[i].GetValue('PST_NATUREFORM'));
  
    if T.GetValue('PFI_SALARIE') <> '' then
    begin
      if VH_Paie.PGForValoSalairePrev = 'VCR' then
          Salaire := ForTauxHoraireReel(T.GetValue('PFI_SALARIE'),0,0,'',ValPrevisionnel) //PT13
      else
          Salaire := ForTauxHoraireCategoriel(T.Getvalue('PFI_LIBEMPLOIFOR'), T.Getvalue('PFI_MILLESIME'));
      CoutSal := Salaire * TobLesFormations.Detail[i].GetValue('PST_JOURSTAGE') * T.Getvalue('PFI_NBINSC');
      Q := OpenSQL('SELECT PSE_RESPONSFOR FROM DEPORTSAL WHERE PSE_SALARIE="' + T.Getvalue('PFI_SALARIE') + '"', True);
      if not Q.eof then
      begin
        T.PutValue('PFI_RESPONSFOR', Q.Findfield('PSE_RESPONSFOR').AsString);
        SetField('PFI_RESPONSFOR', Q.Findfield('PSE_RESPONSFOR').AsString);
      end;
      Ferme(Q);
      T.PutValue('PFI_DUREESTAGE', TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE') * T.Getvalue('PFI_NBINSC')); //PT17
      if (DefautHTpsT) or (GetField('PFI_TYPEPLANPREV') = 'PLF') then
      begin
         T.PutValue('PFI_HTPSTRAV', TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE') * T.Getvalue('PFI_NBINSC')); //PT17
         T.PutValue('PFI_HTPSNONTRAV', 0);
      end
      else
      begin
         T.PutValue('PFI_HTPSTRAV', 0);
         T.PutValue('PFI_HTPSNONTRAV', TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE') * T.Getvalue('PFI_NBINSC'));  //PT17
      end;
    end
    else
    begin
      T.PutValue('PFI_DUREESTAGE', TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE') * T.Getvalue('PFI_NBINSC'));
      Salaire := ForTauxHoraireCategoriel(TF.GetValue('PFI_LIBEMPLOIFOR'), MillesimeEC);
      NbInsc := T.Getvalue('PFI_NBINSC');
      NbHeure := TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE');
      CoutSal := Salaire * NbHeure * NbInsc;
    end;
    T.PutValue('PFI_COUTREELSAL', CoutSal * TauxChargeBudget); //PT5
    CoutSalCur := CoutSalCur + CoutSal;
    //PT15 - Début
    Req := 'SELECT PFF_FRAISHEBERG,PFF_FRAISREPAS,PFF_FRAISTRANSP FROM FORFAITFORM ' +
           'WHERE PFF_MILLESIME="' + T.GetValue('PFI_MILLESIME') + '" AND PFF_LIEUFORM="' + T.GetValue('PFI_LIEUFORM') + '"'; //PT3
    If (VH_Paie.PGForGestFraisByPop) Then
    Begin
           // Un salarié a été saisi : on récupère sa population directement
           If GetField('PFI_SALARIE') <> '' then
                Req := Req + ' AND PFF_POPULATION IN (SELECT PNA_POPULATION FROM SALARIEPOPUL WHERE PNA_SALARIE="'+GetField('PFI_SALARIE')+'")'
           // Un nombre de salarié a été saisi (et donc pas de salarié précis) : on essaye de déterminer la population
           // en fonction des données présentes
           Else
           Begin
                TobEcran := Tob.Create('INSCFORMATION', Nil, -1);
                TobEcran.GetEcran(Ecran);
                Req := Req + ' AND PFF_POPULATION="'+RecherchePopulation(TobEcran)+'"';
                FreeAndNil (TobEcran);
           End;
    End
    Else
          Req := Req + ' AND PFF_ETABLISSEMENT="' + T.GetValue('PFI_ETABLISSEMENT') + '"';
    //PT15 - Fin
    Q := OpenSQL(Req, True);
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
    NbJour := T.GetValue('PFI_JOURSTAGE'); //PT3
    NbInsc := T.GetValue('PFI_NBINSC');
    FraisH := FraisH * (NbJour - 1);
    if FraisH < 0 then FraisH := 0;
    If FraisH > 0 then FraisR := FraisR*((NbJour*2)-1) //PT10
    else FraisR := FraisR*NbJour;
    if FraisR < 0 then FraisR := 0;
    TotalFrais := FraisH + FraisR + FraisT;
    TotalFrais := TotalFrais * NbInsc;
    T.PutValue('PFI_FRAISFORFAIT', TotalFrais);
    T.PutValue('PFI_JOURSTAGE', NbJour * NbInsc);
    TotalFraisCur := TotalFraisCur + TotalFrais;
    NbHCur := NbHcur + (TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE') * NbInsc);
    NbJourCur := NbJourCur + (NbJour * NbInsc);
    CtPedag := CalCoutBudgetFor(TobLesFormations.Detail[i].GetValue('PCC_CODESTAGE'), GetField('PFI_MILLESIME'), IntToStr(NbInsc));
    CtPedagCursus := CtPedagCursus + CtPedag;
    T.PutValue('PFI_AUTRECOUT', CtPedag);
    T.InsertOrUpdateDB;
    if T <> nil then T.Free;
  end;
  TF.StartUpdate;
  TF.DisableTOM;
  TF.PutValue('PFI_COUTREELSAL', CoutSalCur * TauxChargeBudget);
  TF.PutValue('PFI_FRAISFORFAIT', TotalFraisCur);
  TF.PutValue('PFI_DUREESTAGE', NbHCur);
  if (DefautHTpsT) or (GetField('PFI_TYPEPLANPREV') = 'PLF') then
  begin
    TF.PutValue('PFI_HTPSTRAV', NbHCur);
    TF.PutValue('PFI_HTPSNONTRAV', 0);
  end
  else
  begin
    TF.PutValue('PFI_HTPSTRAV', 0);
    TF.PutValue('PFI_HTPSNONTRAV', NbHCur);
  end;
  TF.PutValue('PFI_JOURSTAGE', NbJourCur);
  TF.PutValue('PFI_AUTRECOUT', CtPedagCursus);
  
  TF.Post;
  TF.EnableTOM;
  TF.EndUpdate;
  if TobLesFormations <> nil then TobLesFormations.Free;
end;

procedure TOM_INSCFORMATION.SuppressionCursus;
begin
  ExecuteSQL('DELETE FROM INSCFORMATION WHERE PFI_CURSUS="' + GetField('PFI_CURSUS') + '" AND PFI_CODESTAGE<>"--CURSUS--" ' +
    'AND PFI_ETABLISSEMENT="' + GetField('PFI_ETABLISSEMENT') + '" AND PFI_RANG="' + IntToStr(GetField('PFI_RANG')) + '" AND PFI_MILLESIME="' + GetField('PFI_MILLESIME') + '"');
end;

procedure TOM_INSCFORMATION.CreerCursusPrev;
var
  Q: TQuery;
  TobCursusStage: Tob;
  i: Integer;
  Stage: string;
begin
  if not ExisteSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="' + LeCursus + '" AND PCC_MILLESIME="' + MillesimeEC + '"') then
  begin
    Q := OpenSQL('SELECT * FROM CURSUSSTAGE WHERE PCC_CURSUS="' + LeCursus + '" AND PCC_MILLESIME="0000"', True);
    TobCursusStage := Tob.Create('CURSUSSTAGE', nil, -1);
    TobCursusStage.LoadDetailDB('CURSUSSTAGE', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TobCursusStage.Detail.Count - 1 do
    begin
      Stage := TobCursusStage.Detail[i].GetValue('PCC_CODESTAGE');
      TobCursusStage.Detail[i].PutValue('PCC_MILLESIME', MillesimeEC);
      TobCursusStage.InsertOrUpdateDB();
      if not ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + MillesimeEC + '"') then
      begin
        CreerStagePrev(Stage);
      end;
    end;
    TobCursusStage.Free;
  end;
end;

procedure TOM_INSCFORMATION.CreerStagePrev(Stage: string);
var
  //Q, QAnim: TQuery;       //PT17
  //TobStage, T, TobInvest, TI: Tob;
  //NbHeure: Integer;
  //SalaireAnim, NbAnim: Double;
  Bt: TToolBarButton97;
begin
    CreerStageAuPrevisionnel (Stage, MillesimeEC); //PT16

    {$IFNDEF EMANAGER} //PT17 : si rafraîchissement de la liste, on perd les infos de l'inscription
    If Ecran Is TFSaisieList then
    begin
        TF.RefreshEntete;
        SetControltext('XX_WHERE', 'PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + MillesimeEC + '"');
        Bt := TToolBarButton97(GetControl('BCHERCHE'));
        if Bt <> nil then Bt.Click;
    end;                 
    {$ENDIF}
end;

function TOM_INSCFORMATION.CalCoutBudgetFor(Stage, Millesime, NbInsc: string): Double;
var
  NbAnim,NbInscStage, NbMax: Integer;
  NbHeure ,SalaireAnim,CoutFonc, CoutUnit, CoutAnim, NbDivD, NbDivE, NbCout: Double;
  Q,QAnim : TQuery;
begin
  Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX FROM STAGE ' +
    'WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + Millesime + '"', True);
  if not Q.Eof then
  begin
    CoutFonc := Q.FindField('PST_COUTBUDGETE').AsFloat;
    CoutUnit := Q.FindField('PST_COUTUNITAIRE').AsFloat;
    CoutAnim := Q.FindField('PST_COUTSALAIR').AsFloat;
    NbMax := Q.FindField('PST_NBSTAMAX').AsInteger;
    Ferme(Q);
  end
  else
  begin
    Ferme(Q);
    Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,PST_DUREESTAGE,PST_NBANIM FROM STAGE ' +
      'WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="0000"', True);
    if not Q.Eof then
    begin
      CoutFonc := Q.FindField('PST_COUTBUDGETE').AsFloat;
      CoutUnit := Q.FindField('PST_COUTUNITAIRE').AsFloat;
      NbMax := Q.FindField('PST_NBSTAMAX').AsInteger;
      NbHeure := Q.FindField('PST_DUREESTAGE').AsFloat;
      NbAnim := Q.FindField('PST_NBANIM').AsInteger;
      Ferme(Q);
      QAnim := OpenSQl('SELECT PFE_SALAIREANIM FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
      salaireanim := 0;
      if not QAnim.eof then SalaireAnim := QAnim.FindField('PFE_SALAIREANIM').AsFloat;
      Ferme(QAnim);
      CoutAnim := SalaireAnim * NbHeure * NbAnim;
    end
    else
    begin
      CoutUnit := 0;
      CoutAnim := 0;
      Coutfonc := 0;
      NbMax := 0;
      Ferme(Q);
    end;
  end;
  Q := OpenSQL('SELECT SUM(PFI_NBINSC) NBINSC FROM INSCFORMATION ' +
    'WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '" ' +
    'AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL") ' +
    'AND (PFI_ETABLISSEMENT<>"' + GetField('PFI_ETABLISSEMENT') + '" OR PFI_RANG<>' + IntToStr(GetField('PFI_RANG')) + ')', True);
  if not Q.Eof then NbInscStage := Q.FindField('NBINSC').AsInteger
  else NbInscStage := 0;
  Ferme(Q);
  NbInscStage := NbInscStage + GetField('PFI_NBINSC');
  if NbMax > 0 then
  begin
    NbDivD := NbInscStage / NbMax;
    NbDivE := arrondi(NbDivD, 0);
    if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
    else NbCout := NbDivE
  end
  else NbCout := 1;
  CoutFonc := (NbCout * CoutFonc) / NbInscStage;
  CoutAnim := (NbCout * CoutAnim) / NbInscStage;
  Result := (StrToInt(NbInsc)) * (coutFonc + CoutAnim + Coutunit);
end;

procedure TOM_INSCFORMATION.MajAfficheCoutInsc(Individuel, Global: Boolean);
var
  CtInd: Double;
  i: Integer;
  Q : TQuery;
  T : Tob;
  Nature : String;
begin

  if Global then
  begin
     If TobNature <> Nil then freeandnil(TobNature);
    TobNature := Tob.Create('Naturedesform',Nil,-1);
    CoutATT := 0;
    coutValid := 0;
    if TF.TobFiltre.Detail.Count > 0 then
    begin
      for i := 0 to TF.TobFiltre.Detail.Count - 1 do
      begin
        if TF.TobFiltre.Detail[i].GeTvalue('PFI_ETATINSCFOR') = 'ATT' then
          CoutAtt := CoutAtt + TF.TobFiltre.Detail[i].GeTvalue('PFI_AUTRECOUT') + TF.TobFiltre.Detail[i].GeTvalue('PFI_COUTREELSAL') +
            TF.TobFiltre.Detail[i].GeTvalue('PFI_FRAISFORFAIT');
        if TF.TobFiltre.Detail[i].GeTvalue('PFI_ETATINSCFOR') = 'VAL' then
          CoutValid := CoutValid + TF.TobFiltre.Detail[i].GeTvalue('PFI_AUTRECOUT') + TF.TobFiltre.Detail[i].GeTvalue('PFI_COUTREELSAL') +
            TF.TobFiltre.Detail[i].GeTvalue('PFI_FRAISFORFAIT');
            T := TobNature.FindFirst(['CODESTAGE'],[TF.TobFiltre.Detail[i].GeTvalue('PFI_CODESTAGE')],False);
            If T = Nil then
            begin
                T:= Tob.Create('FilleNat',TobNature,-1);
                Q := OpenSQL('SELECT PST_NATUREFORM FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'" AND PST_CODESTAGE="'+TF.TobFiltre.Detail[i].GeTvalue('PFI_CODESTAGE')+'"',True);
                If Not Q.Eof then Nature := Q.FindField('PST_NATUREFORM').AsString
                else Nature := '';
                Ferme(Q);
                T.AddChampSupValeur('CODESTAGE',TF.TobFiltre.Detail[i].GeTvalue('PFI_CODESTAGE'));
                T.AddChampSupValeur('NATURE',Nature);
            end;
      end;
    end;
  end;
  if Individuel then
  begin
    CtInd := GetField('PFI_AUTRECOUT') + GetField('PFI_COUTREELSAL') + GetField('PFI_FRAISFORFAIT');
  end
  else CtInd := 0;
  if TypeSaisie = 'VALIDRESPONSFOR' then
  begin
    TFSaisieList(Ecran).RichMessage2.Clear;
    TFSaisieList(Ecran).RichMessage2.Lines.Append('Coûts prévisionnels pour le responsable:');
    TFSaisieList(Ecran).RichMessage2.Lines.Append('        Inscription en cours : ' + FloatToStr(CtInd));
    TFSaisieList(Ecran).RichMessage2.Lines.Append('        Inscription(s) validée(s) : ' + FloatToStr(CoutValid));
    TFSaisieList(Ecran).RichMessage2.Lines.Append('        Inscription(s) en attente : ' + FloatToStr(CoutAtt));
  end;
  If TypeSaisie = 'VALIDRESPCURSUS' then
  begin
        TFSaisieList(Ecran).RichMessage1.Clear;
        TFSaisieList(Ecran).RichMessage1.Lines.Append('Coûts prévisionnels pour le responsable:');
        TFSaisieList(Ecran).RichMessage1.Lines.Append('        Inscription en cours : ' + FloatToStr(CtInd));
        TFSaisieList(Ecran).RichMessage1.Lines.Append('        Inscription(s) validée(s) : ' + FloatToStr(CoutValid));
        TFSaisieList(Ecran).RichMessage1.Lines.Append('        Inscription(s) en attente : ' + FloatToStr(CoutAtt));
  end;
end;

procedure TOM_INSCFORMATION.MajEtatFormCursus(cursus,Etab,millesime,etat,motifinsc,motifdec : String;rang : Integer);
var Q : TQuery;
    TobInsc : Tob;
    i : Integer;
begin
Q := OpenSQL('SELECT * FROM INSCFORMATION WHERE PFI_CURSUS="' + cursus + '" AND ' +
        'PFI_RANG=' + IntToStr(Rang) + ' AND PFI_ETABLISSEMENT="' + Etab + '" ' +
        'AND PFI_MILLESIME="' + millesime + '" AND PFI_CODESTAGE<>"--CURSUS--"', True);
      TobInsc := Tob.Create('INSCFORMATION', nil, -1);
      TobInsc.LoadDetailDB('INSCFORMATION', '', '', Q, False);
      Ferme(Q);
    for i := 0 to TobInsc.Detail.Count - 1 do
      begin
        TobInsc.Detail[i].PutValue('PFI_ETATINSCFOR', etat);
        TobInsc.Detail[i].PutValue('PFI_MOTIFINSCFOR', motifinsc);
        TobInsc.Detail[i].PutValue('PFI_MOTIFDECISION', motifdec);
        TobInsc.Detail[i].UpdateDB(False);
      end;
      TobInsc.Free;
end;


procedure TOM_INSCFORMATION.BZoomCLick(Sender : TObject);
var Stage : String;
begin
     Stage := AGLLanceFiche('PAY','MUL_STAGE','','','RECUPSTAGEDIF');
     SetField('PFI_CODESTAGE',Stage);
     SetControlCaption('LIBSTAGE',RechDom('PGSTAGEFORM',Stage,False));
end;

{Function TOM_INSCFORMATION.RendAllocationFormation(Salarie : String) : Double;
var DateCalcul : TDateTime;
    Q : TQuery;
    Net,Heures,Alloc : Double;
begin
     DateCalcul := PlusDate(V_PGI.DateEntree,-1,'A');
     Q := OpenSQL('SELECT SUM(PPU_CNETAPAYER) NET,SUM(PPU_CHEURESTRAV) HEURES FROM PAIEENCOURS WHERE PPU_DATEFIN>="'+UsDateTime(DateCalcul)+'" '+
     'AND PPU_SALARIE="'+Salarie+'"',True);
     If Not Q.Eof then
     begin
          Net := Q.FindField('NET').AsFloat;
          Heures := Q.FindField('HEURES').AsFloat;
     end
     else
     begin
          Net := 0;
          Heures := 0;
     end;
     Ferme(Q);
     If Heures <> 0 then Alloc := Net / Heures
     else Alloc := 0;
     Alloc := Alloc / 2;
     Result := Alloc;
end;         }

procedure TOM_INSCFORMATION.AffectationHeures;
begin
     if (DefautHTpsT) or (GetField('PFI_TYPEPLANPREV') = 'PLF') then
     begin
          SetField('PFI_HTPSTRAV', GetField('PFI_DUREESTAGE'));
          SetField('PFI_HTPSNONTRAV', 0);
     end
     else
     begin
          SetField('PFI_HTPSTRAV', 0);
          SetField('PFI_HTPSNONTRAV', GetField('PFI_DUREESTAGE'));
     end;
end;

procedure TOM_INSCFORMATION.EnvoiMailresponsable(Action : String);
var Texte: HTStrings;
    Titre, Destinataire: string;
    Responsable : string;
    Q : TQuery;
    TobEcran : TOB;
    DestCopie : String;
begin
	 //PT23
     Responsable := GetField('PFI_RESPONSFOR');
     Destinataire := '';
     Q := OpenSQL('SELECT US_EMAIL,PSE_EMAILPROF FROM DEPORTSAL LEFT JOIN UTILISAT ON PSE_SALARIE=US_AUXILIAIRE WHERE PSE_SALARIE="'+Responsable+'"',True); //PT25
     If Not Q.Eof then 
     Begin
     	If Q.FindField('US_EMAIL').AsString <> '' Then //PT25
     		Destinataire := Q.FindField('US_EMAIL').AsString
     	Else
     		Destinataire := Q.FindField('PSE_EMAILPROF').AsString;
     End;
     Ferme(Q);
     
     If Destinataire = '' then Exit;
     
     If action = 'SAISIE' then
     begin
          //PT23 - Appel de PrepareMailFormation à la place
          TobEcran := TOB.Create('INSCFORMATION', Nil, -1);
          TobEcran.GetEcran(Ecran);
          PrepareMailFormation(GetField('PFI_SALARIE'),GetField('PFI_RESPONSFOR'),'SAISIEDIF',TobEcran,Titre,Texte,True);
          FreeAndNil(TobEcran);
  		  //PT23 - Fin
  		  If (VH_PAIE.PGForMailAdr <> '') Then DestCopie := VH_PAIE.PGForMailAdr;
  		  
          SendMail(Titre, Destinataire, DestCopie, Texte, '', True,1,'','',False,False);
          Texte.free;
     end;
     If Action = 'DECISION' then
     begin
     	  //PT27 - Envoi du mail
     	  //Exit; Pourquoi y avait-il un exit ici???
          TobEcran := TOB.Create('INSCFORMATION', Nil, -1);
          TobEcran.GetEcran(Ecran);
          PrepareMailFormation(GetField('PFI_SALARIE'),GetField('PFI_RESPONSFOR'),'DECISIONDIF',TobEcran,Titre,Texte,True);
          FreeAndNil(TobEcran);
  		  If (VH_PAIE.PGForMailAdr <> '') Then DestCopie := VH_PAIE.PGForMailAdr;
  		  
          SendMail(Titre, Destinataire, DestCopie, Texte, '', True,1,'','',False,False);
          Texte.free;
     end;
end;

procedure TOM_INSCFORMATION.RecupInfosStage;
var Q : TQuery;
    DureeStage,JourStage : Double;
    WhereMillesime : String;
begin
     If Not ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="'+GetField('PFI_CODESTAGE')+'" AND PST_MILLESIME="'+GetField('PFI_MILLESIME')+'"') Then WhereMillesime := '0000'
     else WhereMillesime := GetField('PFI_MILLESIME');
     Q := OpenSQL('SELECT PST_PREDEFINI,PST_NODOSSIER,PST_DUREESTAGE,PST_JOURSTAGE,PST_LIEUFORM,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,'+
     'PST_FORMATION7,PST_FORMATION8,PST_NATUREFORM FROM STAGE WHERE PST_MILLESIME="'+WhereMillesime+'" AND PST_CODESTAGE="'+GetField('PFI_CODESTAGE')+'"',True);
     If Not Q.Eof then
     begin
          SetField('PFI_LIEUFORM',Q.FindField('PST_LIEUFORM').AsString);
          SetField('PFI_FORMATION1',Q.FindField('PST_FORMATION1').AsString);
          SetField('PFI_FORMATION2',Q.FindField('PST_FORMATION2').AsString);
          SetField('PFI_FORMATION3',Q.FindField('PST_FORMATION3').AsString);
          SetField('PFI_FORMATION4',Q.FindField('PST_FORMATION4').AsString);
          SetField('PFI_FORMATION5',Q.FindField('PST_FORMATION5').AsString);
          SetField('PFI_FORMATION6',Q.FindField('PST_FORMATION6').AsString);
          SetField('PFI_FORMATION7',Q.FindField('PST_FORMATION7').AsString);
          SetField('PFI_FORMATION8',Q.FindField('PST_FORMATION8').AsString);
          SetField('PFI_NATUREFORM',Q.FindField('PST_NATUREFORM').AsString);
          If Q.FindField('PST_PREDEFINI').AsString = 'STD' then InscriptionsStd := True;
          DureeStage := Q.FindField('PST_DUREESTAGE').AsFloat;
          JourStage := Q.FindField('PST_JOURSTAGE').AsFloat;
     end
     else
     begin
          Ferme(Q);
          Q := OpenSQL('SELECT PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE WHERE PST_CODESTAGE="' + Getfield('PFI_CODESTAGE') + '" AND PST_MILLESIME="0000"', True);
          if not Q.Eof then
          begin
               DureeStage := Q.FindField('PST_DUREESTAGE').AsFloat;
               JourStage := Q.FindField('PST_JOURSTAGE').AsFloat;
          end
          else
          begin
               DureeStage := 0;
               JourStage := 0;
          end;
     end;
     Ferme(Q);
     if GetField('PFI_DUREESTAGE') <> (DureeStage * GetField('PFI_NBINSC')) then SetField('PFI_DUREESTAGE', DureeStage * GetField('PFI_NBINSC'));
     if GetField('PFI_JOURSTAGE') <> (JourStage * GetField('PFI_NBINSC')) then SetField('PFI_JOURSTAGE', JourStage * GetField('PFI_NBINSC'));
     //PT17 - Début
     {$IFDEF EMANAGER}
     If ExisteSQL('SELECT 1 FROM CURSUSSTAGE WHERE PCC_CURSUS<>"---" AND PCC_CURSUS<>"-C-" AND PCC_CODESTAGE="'+GetField('PFI_CODESTAGE')+'"') Then
     PGIInfo('Ce stage fait partie d''un cursus.',Ecran.Caption);
     {$ENDIF}
     //PT17 - Fin
end;

procedure TOM_INSCFORMATION.VoirCompteursDIF(Sender : TObject);
begin
     If GetField('PFI_SALARIE') = '' then
     begin
          PGIBox('Vous devez renseigner un salarié',Ecran.Caption);
          Exit;
     end;
     AGLLanceFiche('PAY','COMPTEURSDIF','','',GetField('PFI_SALARIE'));
end;

procedure TOM_INSCFORMATION.StageElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StWhere := 'PST_ACTIF="X" AND CO_TYPE="PFN" AND '+  //PT2
                  ' PST_MILLESIME="0000" ';
        StOrder := 'PST_LIBELLE,PST_MILLESIME';
        
        //PT24
		If PGBundleInscFormation then
		Begin
			If not PGDroitMultiForm then
				StWhere := StWhere + DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)
	       	Else If V_PGI.ModePCL='1' Then 
				StWhere := StWhere + DossiersAInterroger('','','PST',True,True);
		End;
        
        {$IFNDEF EAGLCLIENT}
        LookupList(THDBEdit(Sender),'Liste des stages','STAGE LEFT JOIN COMMUN ON PST_NATUREFORM=CO_CODE','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,CO_LIBELLE',StWhere,StOrder, True,-1);
        {$ELSE}
        LookupList(THEdit(Sender),'Liste des stages','STAGE LEFT JOIN COMMUN ON PST_NATUREFORM=CO_CODE','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,CO_LIBELLE',StWhere,StOrder, True,-1);
        {$ENDIF}
end;

Procedure TOM_INSCFORMATION.MajCoutPrev(Stage,Millesime : String);
var Q : TQuery;
    i : Integer;
    CoutFonc,CoutUnit,CoutAnim,NbMax, NbDivD,NbDivE : Double;
    NbInsc,NbCout,CoutFoncSta,CoutAnimSta : Double;
    NbInscStage : Integer;
    TobInsc : TOB;
begin
     CoutFonc := 0;
     CoutUnit := 0;
     CoutAnim := 0;
     NbMax := 0;
     //NbDivD := 0;
     //NbDivE := 0;
     Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE ' +
     'WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + Millesime + '"', True);
     If Not Q.eof then
     begin
          CoutFonc := Q.FindField('PST_COUTBUDGETE').AsFloat;
          CoutUnit := Q.FindField('PST_COUTUNITAIRE').AsFloat;
          CoutAnim := Q.FindField('PST_COUTSALAIR').AsFloat;
          NbMax := Q.FindField('PST_NBSTAMAX').AsInteger;
     end;
     Ferme(Q);
     Q := OpenSQL('SELECT SUM(PFI_NBINSC) NBINSC FROM INSCFORMATION ' +
                          'WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '" ' +
                          'AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL")', True);
     if not Q.Eof then NbInscStage := Q.FindField('NBINSC').AsInteger
     else NbInscStage := 1;
     Ferme(Q);
     
     //PT28 - Début
     // Plus de traitement à partir du TF pour que tous les salariés soient traités, même ceux qui ne sont pas visibles (confidentiels)
     (*TF.DisableTom;
     TF.StartUpdate;
     for i:=1 to TF.LaGrid.RowCount - 1 do
     begin
          TF.SelectRecord(i);
          MoveCur(False);
          NbInsc  := TF.GetValue('PFI_NBINSC');
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
          If (TF.GetValue('PFI_ETATINSCFOR')='ATT') or (TF.GetValue('PFI_ETATINSCFOR')='VAL') then TF.PutValue('PFI_AUTRECOUT',NbInsc * (CoutFoncSta + CoutAnimSta + Coutunit))
          else TF.PutValue('PFI_AUTRECOUT',0);
          TF.Post;
     end;
     TF.EndUpdate;
     TF.EnableTom;*)
     
     TobInsc := TOB.Create('~LesInscrits', Nil, -1);
     TobInsc.LoadDetailDBFromSQL('INSCFORMATION','SELECT * FROM INSCFORMATION WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '"');
     
     for i:=0 to TobInsc.Detail.Count - 1 do
     begin
          NbInsc  := TobInsc.Detail[i].GetValue('PFI_NBINSC');
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
          If (TobInsc.Detail[i].GetValue('PFI_ETATINSCFOR')='ATT') or (TobInsc.Detail[i].GetValue('PFI_ETATINSCFOR')='VAL') then 
          	TobInsc.Detail[i].PutValue('PFI_AUTRECOUT',NbInsc * (CoutFoncSta + CoutAnimSta + Coutunit))
          else 
          	TobInsc.Detail[i].PutValue('PFI_AUTRECOUT',0);
     end;
     
     TobInsc.UpdateDB;
     FreeAndNil(TobInsc);
     
     TF.RefreshLignes;
     //PT28 - Fin
end;

initialization
  registerclasses([TOM_INSCFORMATION]);
end.

