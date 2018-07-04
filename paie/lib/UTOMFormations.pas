{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 08/07/2002
Modifié le ...  :    /  /
Description .. : Source TOM de la TABLE : FORMATIONS (FORMATIONS)
Mots clefs ... : TOM;FORMATIONS
*****************************************************************
PT1  | 04/11/2003 | JL | V_50  | Modif erreur sur calcul salaire (* millésime)
PT2  | 12/11/2003 | JL | V_50  | Gestion affichage récupération prévisionnel en CWAS
PT3  | 19/11/2003 | JL | V_50  | Correction viloation d'accès fiche de présence (suppression contrôle sur taux de charge)
PT4  | 24/11/2003 | JL | V_50  | Correction LookUpList pour CWAS
PT5  | 27/01/2005 | JL | V_60  | Recherche prévisonnel sur code groupe au lieu du stage
PT6  | 28/02/2005 | JL | V_60  | FQ 12036 modif titre pour validation
PT7  | 12/07/2005 | PH | V_60  | Bug saisie des diplômes des salariés Erreur sur la TOB ?????
PT8  | 16/03/2006 | JL | V_650 | FQ 12977 Changemetn calcul du salaire pour permettre saisie salaires a 0.
PT9  | 26/04/2006 | JL | V_650 | FQ 12985 Bouton nouveau supp si inscriptions clôturées
PT10 | 06/06/2006 |    | V_65  | FQ 13028 gestion rapprochement DIF sans millésime
PT11 | 30/08/2006 | JL | V_70  | FQ 13462 Calcul auto du coût pédagogique
PT12 | 13/09/2006 | JL | V_70  | Calcul du salaire après une inscription qui n'était plus effectué
PT13 | 13/09/2006 | JL | V_70  | Ajout taux de charge dans validation en masse présence
PT14 | 09/10/2006 | JL | V_70  | Supp des frais associés au stagiaire
PT15 | 25/01/2007 | JL | V_75  | Supp du treeview dans fiche saisie liste
PT16 | 26/01/2007 | FC | V_80  | Mise en place du filtage habilitation
     |            |    |       | Prise en compte de la clause habilitation dans les lookuplist
PT17 | 12/04/2007 | FL | V_720 | FQ 14047 Suppression de l'onglet "Récupération budget" qui n'est plus utilisé.
     |            |    |       |    Gestion du bouton de récupération du prévisionnel.
PT18 | 16/04/2007 | FL | V_720 | FQ 13568 Correction de la recherche du millésime pour la formation
PT19 | 17/04/2007 | FL | V_720 | FQ 14071 Calcul du T.H. salarié : prise en compte de l'historique salarié s'il existe
PT20 | 14/06/2007 | JL | V_720 | Modif pour partage formation
PT21 | 02/04/2008 | FL | V803  | Suite des modifications pour le partage
PT22 | 28/05/2008 | FL | V_804 | FQ 14052 Déplacement de certaines fonctions dans PGOutilsFormation pour le DIF
PT23 | 28/05/2008 | FL | V_804 | Ajout de l'IDSESSIONFOR
PT24 | 05/06/2008 | FL | V_804 | FQ 15458 Report V7 Mise à jour des coûts même pour les salariés non visibles (confidentiels)
}
Unit UTOMFormations;


Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Fiche,Fe_Main,FichList,
{$ELSE}
      MaineAgl,eFiche,eFichList,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,LookUp,uTableFiltre,
     SaisieList,EntPaie,HSysMenu,Graphics,PgOutilsFormation,HTB97,ParamDat,AGLInit,HStatus,ed_tools,ParamSoc,PGOutils,PGCommun;

Type TOM_FORMATIONS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnCancelRecord             ; override ;  //PT17
    private
    InitOrdre,InitStage,InitSalarie : String;
    MillesimeEC,TypeSaisie : String;
    TF :  TTableFiltre;
    LeStage,LaSession,LeMillesime,LeCursus : String;
    TauxChargeC,TauxChargeNC : Double;
    LeRangCursus : Integer;
    DefautHTpsT : Boolean;
    InitTypePlan : String;
    InitDatePaie : TDateTime;
    SalarieInscription : Boolean;
    procedure ExitSession(Sender : TObject);
    procedure StageElipsisClick(Sender : TObject);
    procedure AfficheGrilleFrais;
    procedure BPresenceClick(Sender : TObject);
    procedure BAcceptClick(Sender : TObject);
    procedure AfficheInfos(Sender : TObject);
    procedure AfficheSession;
    procedure VerfifClotureInsc(Stage,Session,Millesime : String);
    procedure DateElipsisclick(Sender :  TObject);
//    Procedure AfficheGrilleInsc(Stage :  String); //PT17
    Procedure RecupInfosSalarie(TS : String);
    procedure BSelectionClick (Sender : TObject);
    procedure SalarieElipsisClick (Sender : TObject);
    procedure CodeSalarieElipsisClick (Sender : TObject);
    //procedure MajCompetences( Matricule : String;Present : Boolean); //PT22
    procedure InscriptionCursus;
    procedure MajCompetencesInitiales;
    procedure VerifReaPrevisionnel(Stage,Salarie,Effectue : String);
//    procedure InsererDIFBulletin; //PT22
//    procedure SupprimerDIF(Supp : Boolean; var SuppAbsence,SuppRubrique : Boolean); //PT22
    procedure VoirMouvementsGenere(Sender : TObject);
    procedure VoirCompteursDIF(Sender : TObject);
    procedure MettreAjourLeSalaire;
    end ;

Implementation

Uses UtilPGI;

procedure TOM_FORMATIONS.OnAfterDeleteRecord;
begin
Inherited ;
	//PT24
	MAJCoutsFormation(LeMillesime, LeStage, LaSession, Nil);
	TF.RefreshLignes;
end;

procedure TOM_FORMATIONS.OnDeleteRecord;
var Rep : Integer;
    SuppAbsence,SuppRubrique : Boolean;
    TobInscription : TOB; //PT22
begin
     Inherited ;
        ExecuteSQL('DELETE FROM FRAISSALFORM WHERE PFS_SALARIE="'+GetField('PFO_SALARIE')+'" AND '+
                'PFS_CODESTAGE="'+GetField('PFO_CODESTAGE')+'" AND PFS_ORDRE='+IntToStr(GetField('PFO_ORDRE'))); //PT14
        If TypeSaisie = 'INITIALE' then
        begin
                Rep := PGIAsk('La suppression de cette formation va entraîner la suppression des compétences liées pour ce salarié,#13#10 voulez-vous continuer ?',Ecran.Caption);
                If Rep = MrYes then
                begin
                        ExecuteSQL('DELETE FROM RHCOMPETRESSOURCE WHERE PCH_CODESTAGE="'+GetField('PFO_CODESTAGE')+'" AND PCH_SESSIONSTAGE='+IntToStr(GetField('PFO_ORDRE')));
                end
                else
                begin
                        LastError := 1;
                        Exit;
                end;
        end;
        //SupprimerDIF(True,SuppAbsence,SuppRubrique); //PT22
        TobInscription := TOB.Create('FORMATIONS', Nil, -1);
        TobInscription.GetEcran(Ecran);
        SupprimerDIF(TobInscription,InitDatePaie,True,SuppAbsence,SuppRubrique);
        FreeAndNil(TobInscription);
end;

procedure TOM_FORMATIONS.OnAfterUpdateRecord;
var TobSession : Tob;
    Q : TQuery;
    Stage,Session,Millesime : String;
    i : Integer;
    Present : Boolean;
    TobInscription : TOB; //PT22
begin
     Inherited ;
       If TypeSaisie = 'INITIALE' then
       begin
           MajCompetencesInitiales;
           exit;
       end
       Else If (TypeSaisie='SAISIEFP') Then
       Begin
           //PT22
           TobInscription := TOB.Create('FORMATIONS', Nil, -1);
           TobInscription.GetEcran(Ecran);
           InsererDIFBulletin(TobInscription,InitDatePaie);
           SetField('PFO_DATEPAIE', TobInscription.GetValue('PFO_DATEPAIE'));
           FreeAndNil(TobInscription);
       End;
        InitTypePlan := GetField('PFO_TYPEPLANPREV');
        If TypeSaisie = 'CURSUS' then
        begin
                InscriptionCursus;
                Exit;
        end;
//        If (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'CWASSAISIEFORMATION') or (TypeSaisie = 'SAISIESESSION') Then AfficheGrilleInsc(LeStage);//PT15 //PT17
        If TypeSaisie = 'SAISIEFP' Then
        begin
                Stage  := LeStage;//PT15
                Session := LaSession;//PT15
                Millesime := LeMillesime;//PT15
                If GetField('PFO_EFFECTUE') ='X' Then
                begin
                        Present := True;
                        Q := OpenSQL('SELECT * FROM SESSIONSTAGE WHERE'+
                        ' PSS_CODESTAGE="'+Stage+'" AND PSS_ORDRE='+Session+' AND PSS_MILLESIME="'+Millesime+'" AND PSS_EFFECTUE<>"X"',true); //DB2
                        TobSession := Tob.Create('SESSIONSTAGE',Nil,-1);
                        TobSession.LoadDetailDB('SESSIONSTAGE','','',Q,False);
                        Ferme(Q);
                        For i :=0 to TobSession.Detail.Count - 1 do
                        begin
                                TobSession.Detail[i].PutValue('PSS_EFFECTUE','X');
                                TobSession.Detail[i].PutValue('PSS_CLOTUREINSC','X');
                                TobSession.Detail[i].UpdateDB;
                        end;
                        TobSession.Free;
                end
                else Present := False;
                //MAJCoutsFormation(Millesime,Stage,Session);
                
                //PT24 - Plus de passage du TF pour que tous les salariés soient traités, même ceux qui ne sont pas visibles (confidentiels)
                //TF.StartUpdate;
                //MAJCoutsFormation(GetField('PFO_MILLESIME'), GetField('PFO_CODESTAGE'), IntToStr(GetField('PFO_ORDRE')),TF); //PT11
                MAJCoutsFormation(GetField('PFO_MILLESIME'), GetField('PFO_CODESTAGE'), IntToStr(GetField('PFO_ORDRE')),Nil);
                //TF.EndUpdate;
                TF.RefreshLignes;
                
                MajCompetences(GetField('PFO_SALARIE'),LeStage,LaSession,LeMillesime,Present); //PT22
        end;
       If isFieldModified( 'PFO_COUTREELSAL' ) then CalcCtInvestSession('SALAIRE',GetField('PFO_CODESTAGE'),GetField('PFO_MILLESIME'),GetField('PFO_ORDRE'));
       InitDatePaie := GetField('PFO_DATEPAIE');
       If GetField('PFO_PGTYPESESSION') = 'SOS' then MajSousSessionSal(getField('PFO_SALARIE'),getField('PFO_CODESTAGE'),IntToStr(getField('PFO_NOSESSIONMULTI')));

       SetControlEnabled('BRECUPERATION', True); //PT17
end;

procedure TOM_FORMATIONS.AfficheInfos(Sender : TObject);
begin
    //    SetControlVisible('BINSERT',true);     //PT9 Mis en commentaire
        SetControlVisible('BDELETE',True);
        If TypeSaisie<>'VALIDRESPONSFOR' Then AfficheSession;
end;

procedure TOM_FORMATIONS.AfficheSession;
var Stage,Session,Millesime : String;
    Q : TQuery;
    HeureDeb,HeureFin : TDateTime;
    Num : Integer;
    HMTrad : THSystemMenu;
begin
        If (TypeSaisie='SAISIEINSC') or (TypeSaisie='CWASSAISIEFORMATION')or (typeSaisie='SAISIEFP') or (TypeSaisie='SAISIESESSION') Then
        begin
                Stage := LeStage;
                Session := LaSession;
                Millesime := LeMillesime;
        end
        Else
        begin
                 //DEBUT PT15
                Stage := LeStage;
                Session := LaSession;
                Millesime := LeMillesime;
                //FIN PT15
        end;
        Q := OpenSQL('SELECT PSS_HEUREDEBUT,PSS_HEUREFIN,PSS_DUREESTAGE,PSS_LIBELLE,PST_LIBELLE,PST_LIBELLE1,PSS_DATEDEBUT,PSS_DATEFIN,PSS_DUREESTAGE,PSS_JOURSTAGE'+
           ',PST_NATUREFORM,PST_LIEUFORM,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
           ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8 FROM SESSIONSTAGE'+
           ' LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME'+
           ' WHERE PSS_CODESTAGE="'+Stage+'" AND PSS_MILLESIME="'+Millesime+'" AND PSS_ORDRE="'+Session+'"',True);  //DB2
        if not Q.eof then
        begin
                HeureDeb := Q.FindField('PSS_HEUREDEBUT').AsDatetime;
                HeureFin := Q.FindField('PSS_HEUREFIN').AsDateTime;
                SetControlText('HEUREDEBUT',FormatDateTime('hh:mm',HeureDeb));
                SetControlText('HEUREFIN',FormatDateTime('hh:mm',HeureFin));
                SetControlText('STAGE',Q.FindField('PST_LIBELLE').AsString);
                SetControlText('SESSION',Q.FindField('PSS_LIBELLE').AsString);
                SetControlText('DATEDEBUT',DateToStr(Q.FindField('PSS_DATEDEBUT').AsDateTime));
                SetControlText('DATEFIN',DateToStr(Q.FindField('PSS_DATEFIN').AsDateTime));
                SetControlProperty('NATUREFORM','Value',Q.FindField('PST_NATUREFORM').AsString);
                SetControlProperty('LIEUFORMATION','Value',Q.FindField('PST_LIEUFORM').AsString);
                if typeSaisie<>'SAISIEFP' then
                begin
                        Setcontroltext('FORMATION1',Q.FindField('PST_FORMATION1').AsString);
                        Setcontroltext('FORMATION2',Q.FindField('PST_FORMATION2').AsString);
                        Setcontroltext('FORMATION3',Q.FindField('PST_FORMATION3').AsString);
                        Setcontroltext('FORMATION4',Q.FindField('PST_FORMATION4').AsString);
                        Setcontroltext('FORMATION5',Q.FindField('PST_FORMATION5').AsString);
                        Setcontroltext('FORMATION6',Q.FindField('PST_FORMATION6').AsString);
                        Setcontroltext('FORMATION7',Q.FindField('PST_FORMATION7').AsString);
                        Setcontroltext('FORMATION8',Q.FindField('PST_FORMATION8').AsString);
                        For Num  :=  1 to VH_Paie.NBFormationLibre do
                        begin
                                if Num >8 then Break;
                                VisibiliteChampFormation (IntToStr(Num),GetControl ('FORMATION'+IntToStr(Num)),GetControl ('TFORMATION'+IntToStr(Num)));
                                SetControlEnabled('FORMATION'+IntToStr(Num),False);
                        end;
                end;
        end;
        Ferme(Q);
        HMTrad:=THSystemMenu(GetControl('HMTrad'));
        HMTrad.ResizeGridColumns(TF.LaGrid);
end;

procedure TOM_FORMATIONS.OnNewRecord ;
var Q : TQuery;
begin
  Inherited ;
        If (TypeSaisie <> 'SAISIEFP') AND (TypeSaisie <> 'INITIALE') then // PT7
        begin
             Q := OpenSQL('SELECT PSS_DUREESTAGE,PSS_IDSESSIONFOR FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+LeStage+'"'+ //PT15 //PT23
             ' AND PSS_ORDRE='+LaSession+''+  //DB2
             ' AND PSS_MILLESIME="'+Lemillesime+'"',true);
             if not Q.eof then
             begin
                  SetField('PFO_NBREHEURE',Q.FindField('PSS_DUREESTAGE').AsFloat);
                  SetField('PFO_IDSESSIONFOR',Q.FindField('PSS_IDSESSIONFOR').AsString); //PT23
                  if (DefautHTpsT) or (GetField('PFO_TYPEPLANPREV') = 'PLF') then
                  begin
                       SetField('PFO_HTPSTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                       SetField('PFO_HTPSNONTRAV',0);
                  end
                  else
                  begin
                       SetField('PFO_HTPSTRAV',0);
                       SetField('PFO_HTPSNONTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                  end;
             end;
             Ferme(Q);
        end;
        SetField('PFO_TYPEPLANPREV','PLF');
        If TypeSaisie = 'CURSUS' then
        begin
                SetField('PFO_RANGCURSUS',LeRangCursus);
                SetField('PFO_CURSUS',LeCursus);
        end;
        SetControlEnabled('PFO_SALARIE',True);
        SetField('PFO_EFFECTUE','-');
        SetField('PFO_CIF','-');
        SetField('PFO_ALTERNANCE','-');
        SetField('PFO_TYPOFORMATION','001');
        TF.PutValue('PFO_PLANFORM','-');
        TF.PutValue('PFO_EAO','-');
        TF.PutValue('PFO_COINVEST','-');
        TF.PutValue('PFO_ORDRE',-1);
        TF.PutValue('PFO_ORGCOLLECT',-1);
        TF.PutValue('PFO_DATECOINVES',IDate1900);
        TF.PutValue('PFO_DATEACCEPT',IDate1900);
        TF.PutValue('PFO_REFUSELE',IDate1900);
        TF.PutValue('PFO_REPORTELE',IDate1900);
        SetField('PFO_ETATINSCFOR','ATT');
        If (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESESSION') or (TypeSaisie = 'SAISIEFP') Then
        begin
                if VH_Paie.PGForValidSession = False then SetField('PFO_ETATINSCFOR','VAL');
        end;
        If TypeSaisie = 'INITIALE' then
        begin
                SetField('PFO_DATEFIN',Date);
                TF.PutValue('PFO_NATUREFORM','004');
                TF.PutValue('PFO_MILLESIME','0000');
        end;
end ;

procedure TOM_FORMATIONS.OnUpdateRecord ;
var //TauxHoraireMontant : Double;
    Q       : TQuery;
    IMin    : Integer;
    Req,Pfx : String;
begin
	Inherited ;
	If DS.State in [DSInsert] then
	begin
		InitTypeplan := GetField('PFO_TYPEPLANPREV');
		InitDatePaie := GetField('PFO_DATEPAIE');
	end;
	//        SetField('PFO_HTPSNONTRAV',GetField('PFO_NBREHEURE') - GetField('PFO_HTPSTRAV'));
	If TypeSaisie = 'INITIALE' then
	begin
		If GetField('PFO_CODESTAGE') = '' then
		begin
			PGIBox('Vous devez renseigner le diplôme',Ecran.Caption);
			LastError := 1;
			SetFocusControl('PFO_CODESTAGE');
			Exit;
		end;
		If DS.State in [DSInsert] then
		begin
			Q := OpenSQL('SELECT MIN(PFO_ORDRE) MINI FROM FORMATIONS WHERE PFO_SALARIE="'+TF.TOBFiltre.GetValue('PSA_SALARIE')+'"',True);
			If Not Q.Eof then IMin := Q.FindField('MINI').AsInteger
			else IMin := 0;
			Ferme(Q);

			If Imin < 0 then IMin := IMin - 1
			else IMin := -1;
			SetField('PFO_ORDRE',IMin);

			Q := OpenSQL('SELECT * FROM STAGE WHERE PST_MILLESIME="0000" AND PST_CODESTAGE="'+GetField('PFO_CODESTAGE')+'"',True);
			If Not Q.Eof then
			begin
				SetField('PFO_FORMATION1',Q.FindField('PST_FORMATION1').AsString);
				SetField('PFO_FORMATION2',Q.FindField('PST_FORMATION2').AsString);
				SetField('PFO_FORMATION3',Q.FindField('PST_FORMATION3').AsString);
				SetField('PFO_FORMATION4',Q.FindField('PST_FORMATION4').AsString);
				SetField('PFO_FORMATION5',Q.FindField('PST_FORMATION5').AsString);
				SetField('PFO_FORMATION6',Q.FindField('PST_FORMATION6').AsString);
				SetField('PFO_FORMATION7',Q.FindField('PST_FORMATION7').AsString);
				SetField('PFO_FORMATION8',Q.FindField('PST_FORMATION8').AsString);
				SetField('PFO_LIEUFORM',Q.FindField('PST_LIEUFORM').AsString);
			end;
			Ferme(Q);

			//PT21 - Début
			Req := 'SELECT * FROM SALARIES WHERE PSA_SALARIE="'+TF.TOBFiltre.GetValue('PSA_SALARIE')+'"';
			Req := AdaptMultiDosForm (Req, False);
			//PT21 - Fin

			Q := OpenSQL(Req,True);
			If Not Q.Eof then
			begin
				SetField('PFO_TRAVAILN1',		Q.FindField(Pfx+'_TRAVAILN1').AsString);
				SetField('PFO_TRAVAILN2',		Q.FindField(Pfx+'_TRAVAILN2').AsString);
				SetField('PFO_TRAVAILN3',		Q.FindField(Pfx+'_TRAVAILN3').AsString);
				SetField('PFO_TRAVAILN4',		Q.FindField(Pfx+'_TRAVAILN4').AsString);
				SetField('PFO_CODESTAT',		Q.FindField(Pfx+'_CODESTAT').AsString);
				SetField('PFO_ETABLISSEMENT',	Q.FindField(Pfx+'_ETABLISSEMENT').AsString);
				SetField('PFO_LIBEMPLOIFOR',	Q.FindField(Pfx+'_LIBELLEEMPLOI').AsString);
				SetField('PFO_NOMSALARIE',		Q.FindField(Pfx+'_LIBELLE').AsString);
				SetField('PFO_PRENOM',			Q.FindField(Pfx+'_PRENOM').AsString);
			end;
			Ferme(Q);
			//PT23 - Correction du prédéfini
			//DEBUT PT20
			(*If PGBundleInscFormation then
			begin
			SetField('PFO_NODOSSIER',GetNoDossierSalarie (TF.TOBFiltre.GetValue('PSA_SALARIE')))
			Q := OpenSQL('SELECT PSI_NODOSSIER FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+TF.TOBFiltre.GetValue('PSA_SALARIE')+'"',True);
			If Not Q.eof then
			begin
			SetField('PFO_NODOSSIER',Q.FindField('PSI_NODOSSIER').AsString);
			SetField('PFO_PREDEFINI','DOS');
			end;
			Ferme(Q);
			end
			else
			begin
			SetField('PFO_NODOSSIER',V_PGI.NoDossier);
			end;
			SetField('PFO_PREDEFINI','DOS');*)
			//FINPT20
			If PGBundleInscFormation Then
			Begin
				TF.PutValue('PFO_NODOSSIER', GetNoDossierSalarie(TF.TOBFiltre.GetValue('PSA_SALARIE')));
				TF.PutValue('PFO_PREDEFINI', 'DOS');
			End
			Else
			Begin
				TF.PutValue('PFO_NODOSSIER', V_PGI.NoDossier);
				TF.PutValue('PFO_PREDEFINI', 'STD');
			End;
		end;
		Exit;
	end;

	If GetField('PFO_SALARIE')='' Then
	begin
		LastError := 1;
		PgiBox('Vous devez renseigner le code salarié',TFSaisieList(Ecran).Caption);
	end;
	If TypeSaisie = 'CURSUS' then Exit;
	If (TypeSaisie <> 'SAISIEFP') and (TypeSaisie <> 'CWASVALIDINSC') and (TypeSaisie <> 'VALIDINSC')  Then
	begin
		If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_SALARIE="'+Getfield('PFO_SALARIE')+'" '+
		'AND (PFO_CODESTAGE<>"'+GetField('PFO_CODESTAGE')+'" OR PFO_ORDRE<>'+IntToStr(GetField('PFO_ORDRE'))+' OR PFO_MILLESIME<>"'+Getfield('PFO_MILLESIME')+'") '+ //DB2
		'AND ((PFO_DATEDEBUT>="'+UsDateTime(GetField('PFO_DATEDEBUT'))+'" AND PFO_DATEDEBUT<="'+UsDateTime(GetField('PFO_DATEFIN'))+'") '+
		'OR (PFO_DATEFIN>="'+UsDateTime(GetField('PFO_DATEDEBUT'))+'" AND PFO_DATEFIN<="'+UsDateTime(GetField('PFO_DATEFIN'))+'") OR '+
		'(PFO_DATEDEBUT<="'+UsDateTime(GetField('PFO_DATEDEBUT'))+'" AND PFO_DATEFIN>="'+UsDateTime(GetField('PFO_DATEFIN'))+'"))') then
		begin
			If PGIAsk('Attention, ce salarié est déjà inscrit à une session dont les dates coïncident.#13#10 Voulez-vous continuer?',Ecran.Caption)<>MrYes then
			begin
				LastError := 1;
				Exit;
			end;
		end;
		If DS.State in [DsInsert] then
		begin
			If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_SALARIE="'+Getfield('PFO_SALARIE')+'" '+
			'AND PFO_CODESTAGE="'+GetField('PFO_CODESTAGE')+'" AND PFO_ORDRE='+IntToStr(GetField('PFO_ORDRE'))+' AND PFO_MILLESIME="'+Getfield('PFO_MILLESIME')+'"') then  //DB2
			begin
				PGIBox('Saisie impossible car ce salarié est déjà inscrit.',Ecran.Caption);
				LastError := 1;
				Exit;
			end;
		end;
	end;
	If (TypeSaisie='SAISIEFP') Then
	begin
		If GetField('PFO_NBREHEURE') <> (GetField('PFO_HTPSTRAV') + GetField('PFO_HTPSNONTRAV')) then
		begin
			PGIBox('La somme des heures tarvaillées  et non travaillées doit être égale aux heures totales',Ecran.Caption);
			LastError := 1;
			Exit;
		end;
		{TauxHoraire := 0;       //PT8 = Mis en commentaire, fais sur onchangefield
		if VH_Paie.PGForValoSalaire='VCR' then TauxHoraire := ForTauxHoraireReel(GetField('PFO_SALARIE'));
		if VH_Paie.PGForValoSalaire='VCC' then TauxHoraire := ForTauxHoraireCategoriel(GetField('PFO_LIBEMPLOIFOR'),MillesimeEC);
		Montant := TauxHoraire*GetField('PFO_HTPSTRAV');
		//PT3 suppression contrôle taux cadre <> 1
		Q := OpenSQL('SELECT PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True);
		If (Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02') then Montant := Arrondi(Montant * TauxChargeC,2)
		else Montant := Arrondi(Montant * TauxChargeNC,2);
		Ferme(Q);
		SetField('PFO_COUTREELSAL',Montant);}
		VerifReaPrevisionnel(GetField('PFO_CODESTAGE'),GetField('PFO_SALARIE'),GetField('PFO_EFFECTUE'));
	end;
end ;

procedure TOM_FORMATIONS.OnLoadRecord ;
var CodeStage,Millesime : String;
begin
  Inherited ;
        InitDatePaie := GetField('PFO_DATEPAIE');
        If TypeSaisie = 'INITIALE' then Exit;
        InitSalarie := GetField('PFO_SALARIE');
        If not AfterInserting Then SetControlEnabled('PFO_SALARIE',False)
        Else SetControlEnabled('PFO_SALARIE',True);
        If not (Ecran is TFSaisieList) Then
        begin
                InitOrdre := GetField('PFO_ORDRE');
                CodeStage := GetField('PFO_CODESTAGE');
                Millesime := GetField('PFO_MILLESIME');
                initStage := GetField('PFO_CODESTAGE');
                SetControlProperty('SESSION','Plus','PSS_CODESTAGE="'+GetField('PFO_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PFO_MILLESIME')+'"');
                SetControlText('SESSION',IntToStr(GetField('PFO_ORDRE')));
        end
        Else
        begin
                If TypeSaisie='FRAIS' Then AfficheGrilleFrais;
        end;
        InitTypePlan := GetField('PFO_TYPEPLANPREV');
        If TypeSaisie = 'SAISIEFP' then
        begin
             SetControlProperty('BCOMPTEURS','Hint','Voir les compteurs de ' +GetField('PFO_NOMSALARIE') + ' '+GetField('PFO_PRENOM'));
        end;
        If (GetField('PFO_SALARIE') <> '') and (GetField('PFO_EFFECTUE') = '-') then SalarieInscription := True  //PT12
        else SalarieInscription := False;

        SetControlEnabled('BRECUPERATION', True); //PT17
end ;


{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 13/04/2007 / PT17
Modifié le ... :   /  /    
Description .. : Surcharge de l'évènement OnCancelRecord lancé sur 
Suite ........ : l'annulation d'une modification
Mots clefs ... : 
*****************************************************************}
Procedure TOM_FORMATIONS.OnCancelRecord ;
Begin
     Inherited;
     SetControlEnabled('BRECUPERATION', True);
End;

procedure TOM_FORMATIONS.OnChangeField ( F: TField ) ;
var Q : TQuery;
    Millesime : String;
    TSession : THValComboBox;
    i : Integer;
begin
  Inherited ;
        If TypeSaisie = 'INITIALE' then Exit;
        If (Ecran is TFSaisieList) Then
        begin
                If (TypeSaisie = 'CWASSAISIEFORMATION') or (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'VALIDINSC') or (TypeSaisie = 'SAISIEFP') or (TypeSaisie = 'SAISIESESSION') or (TypeSaisie = 'CWASVALIDINSC') Then
                begin
                        If (F.FieldName = 'PFO_SALARIE') and (TFSaisielist(Ecran).FTypeAction=TaCreat) then
                        begin
                                If (GetField('PFO_SALARIE')<>'') and( InitSalarie = '') Then RecupInfosSalarie(TypeSaisie);
                        end;
                end;
        end;
        If not (Ecran is TFSaisieList) Then
        begin
                If F.FieldName = 'PFO_COINVEST' Then
                begin
                        If GetField('PFO_COINVEST') = 'X' Then
                        begin
                                SetControlEnabled('PFO_DATECOINVES',True);
                                SetControlEnabled('PFO_INDEMCOINVES',True);
                        end
                        Else
                        begin
                                SetControlEnabled('PFO_DATECOINVES',False);
                                SetControlEnabled('PFO_INDEMCOINVES',False);
                        end;
                end;
                If F.FieldName = 'PFO_CODESTAGE' then
                begin
                        TSession := THValComboBox(GetControl('SESSION'));
                        TSession.Plus := 'PSS_CODESTAGE="'+GetField('PFO_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PFO_MILLESIME')+'"';
                        If GetField('PFO_CODESTAGE') <> '' Then SetControlEnabled('SESSION',True)
                        Else SetControlenabled('SESSION',False);
                        If GetField('PFO_CODESTAGE') <> InitStage Then
                        begin
                                If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="'+GetField('PFO_CODESTAGE')+'" AND PST_MILLESIME="'+MillesimeEC+'"') Then
                                begin
                                        SetField('PFO_MILLESIME',MillesimeEC);
                                        Millesime := MillesimeEC;
                                end
                                Else
                                begin
                                        SetField('PFO_MILLESIME','0000');
                                        Millesime := '0000';
                                end;
                                Q := OpenSQL('SELECT PSS_DUREESTAGE,PST_TYPCONVFORM FROM STAGE '+
                                'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME '+
                                'WHERE PST_CODESTAGE="'+GetField('PFO_CODESTAGE')+'"'+
                                ' AND PST_MILLESIME="'+Millesime+'"',True);
                                If not Q.eof then
                                begin
                                        SetField('PFO_NBREHEURE',Q.FindField('PSS_DUREESTAGE').AsFloat);
                                        if (DefautHTpsT) or (GetField('PFO_TYPEPLANPREV') = 'PLF') then
                                        begin
                                             SetField('PFO_HTPSTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                                             SetField('PFO_HTPSNONTRAV',0);
                                        end
                                        else
                                        begin
                                             SetField('PFO_HTPSTRAV',0);
                                             SetField('PFO_HTPSNONTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                                        end;
                                        SetField('PFO_TYPEFORM',Q.FindField('PST_TYPCONVFORM').AsString);
                                end;
                                Ferme(Q);
                               // ForceUpdate;
                        end;
                end;
                If F.FieldName = 'PFO_MILLESIME' Then
                begin
                        TSession := THValComboBox(GetControl('SESSION'));
                        TSession.Plus := 'PSS_CODESTAGE="'+GetField('PFO_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PFO_MILLESIME')+'"';
                end;
        end;
        If (TypeSaisie = 'SAISIEFP')  Then
        begin
                If F.FieldName = 'PFO_EFFECTUE' Then
                begin
                        If (GetField('PFO_EFFECTUE') = 'X') and (GetField('PFO_NBREHEURE')=0) Then
                        begin
                                Q := OpenSQL('SELECT PSS_DUREESTAGE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PFO_CODESTAGE')+'"'+
                                ' AND PSS_ORDRE='+IntToStr(GetField('PFO_ORDRE'))+''+  //DB2
                                ' AND PSS_MILLESIME="'+GetField('PFO_MILLESIME')+'"',true);
                                if not Q.eof then
                                begin
                                     SetField('PFO_NBREHEURE',Q.FindField('PSS_DUREESTAGE').AsFloat);
                                     if (DefautHTpsT) or (GetField('PFO_TYPEPLANPREV') = 'PLF') then
                                     begin
                                          SetField('PFO_HTPSTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                                          SetField('PFO_HTPSNONTRAV',0);
                                     end
                                     else
                                     begin
                                          SetField('PFO_HTPSTRAV',0);
                                          SetField('PFO_HTPSNONTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                                     end;
                                end;
                                Ferme(Q);
                        end;
                        If GetField('PFO_EFFECTUE') = '-' Then
                        begin
                             SetField('PFO_NBREHEURE',0);
                             SetField('PFO_HTPSTRAV',0);
                             SetField('PFO_HTPSNONTRAV',0);
                        end;
                        if (SalarieInscription) and (IsFieldModified('PFO_EFFECTUE')) then  MettreAjourLeSalaire;
                end;
        end;
        If TypeSaisie = 'FRAIS' Then
        begin
                For i := 1 to 6 do
                begin
                        If F.FieldName = ('PFO_FRAISLIBRE'+IntToStr(i)) Then
                        SetField('PFO_FRAISPLAF'+IntToStr(i),VerifPlafondFormation(IntToStr(i),MillesimeEC,GetField('PFO_FRAISLIBRE'+IntToStr(i))));
                end;
        end;
        If (TypeSaisie = 'VALIDINSC') or (TypeSaisie = 'VALIDWEBINSC') or (TypeSaisie = 'CWASVALIDINSC') Then
        begin
                If F.FieldName = 'PFO_DATEACCEPT' Then
                begin
                        If GetField('PFO_DATEACCEPT')<>IDate1900 Then
                        begin
                                SetControlEnabled('PFO_REFUSELE',False);
                                SetControlenabled('PFO_REPORTELE',False);
                                //SetField('PFO_REFUSELE',IDate1900);
                                //SetField('PFO_MOTIFREFUS','');
                                //SetField('REFUSEPAR','');
                                SetControlEnabled('PFO_REFUSEPAR',False);
                                SetControlEnabled('PFO_REFUSFORM',False);
                                SetControlEnabled('PFO_REPORTFORM',False);
                                SetControlEnabled('PFO_REPORTEPAR',False);
                        end
                        Else
                        begin
                                SetControlEnabled('PFO_REPORTELE',True);
                                SetControlEnabled('PFO_REFUSELE',True);
                                SetControlEnabled('PFO_REFUSEPAR',True);
                                SetControlEnabled('PFO_REFUSFORM',True);
                                SetControlEnabled('PFO_REPORTFORM',True);
                                SetControlEnabled('PFO_REPORTEPAR',True);
                        end;
                end;
                If F.FieldName = 'PFO_REFUSELE' Then
                begin
                        If GetField('PFO_REFUSELE')<>IDate1900 then
                        begin
                                //SetField('PFO_DATEACCEPT',IDate1900);
                                SetControlenabled('PFO_DATEACCEPT',False);
                                SetControlEnabled('PFO_REPORTFORM',False);
                                SetControlEnabled('PFO_REPORTEPAR',False);
                                SetControlenabled('PFO_DATEACCEPT',False);
                                SetControlEnabled('PFO_REFUSELE',True);
                                SetControlEnabled('PFO_REFUSEPAR',True);
                        end
                        Else
                        begin
                                SetControlEnabled('PFO_DATEACCEPT',True);
                                SetControlEnabled('PFO_REFUSEPAR',False);
                                SetControlenabled('PFO_REPORTELE',True);
                                SetControlEnabled('PFO_REPORTFORM',True);
                                SetControlEnabled('PFO_REPORTEPAR',True);
                                SetControlenabled('PFO_DATEACCEPT',True);
                        end;
                end;
                If F.FieldName = 'PFO_REPORTELE' Then
                begin
                        If GetField('PFO_REPORTELE')<>IDate1900 then
                        begin
                                SetControlEnabled('PFO_REPORTFORM',True);
                                SetControlEnabled('PFO_REPORTEPAR',True);
                        end
                        Else
                        begin
                                SetControlEnabled('PFO_REPORTFORM',False);
                                SetControlEnabled('PFO_REPORTEPAR',False);
                        end;
                end;
        end;
        If F.FieldName = 'PFO_ETATINSCFOR' then
        begin
                If GetField('PFO_ETATINSCFOR') = 'ATT' then
                begin
                        SetControlEnabled('PFO_MOTIFETATINSC',False);
                end;
                If GetField('PFO_ETATINSCFOR') = 'REF' then
                begin
                        SetControlEnabled('PFO_MOTIFETATINSC',True);
                        SetControlProperty('PFO_MOTIFETATINSC','Plus',' AND CC_LIBRE="REF"');
                end;
                If GetField('PFO_ETATINSCFOR') = 'VAL' then
                begin
                        SetControlEnabled('PFO_MOTIFETATINSC',True);
                        SetControlProperty('PFO_MOTIFETATINSC','Plus',' AND CC_LIBRE="VAL"');
                end;
        end;
        If F.FieldName = 'PFO_NBREHEURE' then
        begin
             if IsFieldModified('PFO_NBREHEURE') then
             begin
                  if (DefautHTpsT) or (GetField('PFO_TYPEPLANPREV') = 'PLF') then
                  begin
                       SetField('PFO_HTPSTRAV',GetField('PFO_NBREHEURE'));
                       SetField('PFO_HTPSNONTRAV',0);
                  end
                  else
                  begin
                       SetField('PFO_HTPSTRAV',0);
                       SetField('PFO_HTPSNONTRAV',GetField('PFO_NBREHEURE'));
                  end;
                  MettreAjourLeSalaire; //PT8
             end;
        end;

        { Désactivation du bouton de récupération des stagiaires prévisionnels si
          - création d'une nouvelle ligne (TYPEPLANPREV sera alors modifié)
          - modification d'une ligne (ETATINSCFOR est le seul élément modifiable)
        }
        If IsFieldModified('PFO_TYPEPLANPREV') Or IsFieldModified('PFO_ETATINSCFOR') Then   //PT17
          SetControlEnabled('BRECUPERATION', False);
end ;

procedure TOM_FORMATIONS.OnArgument ( S : String ) ;
var TSession : THValComboBox;
    {$IFNDEF EAGLCLIENT}
    THStage : THDBEdit;
    {$ELSE}
    THStage : THEdit;
    {$ENDIF}
    Q : TQuery;
    Edit : THEdit;
    BPresence,BAccept,BSelect,BCompteurs,BMvts : TToolBarButton97;
    Pgeneral,PComplement,PSalarie : TTabSheet;
    THDate : THEdit;
    FinSess : TDateTime; //PT18
begin
  Inherited ;
        DefautHTpsT := GetParamSoc('SO_PGDIFTPSTRAV');
        TypeSaisie := ReadTokenSt(S);
        If TypeSaisie = 'RECUPPREVREA' then TypeSaisie := 'SAISIEFP';
        If TypeSaisie = 'RECUPPREVPLA' then TypeSaisie := 'SAISIESTAGEFORMAT';
        If TypeSaisie = 'INITIALE' then
        begin
                TF  :=  TFSaisieList(Ecran).LeFiltre;
                Exit;
        end;
        If TypeSaisie ='CURSUS' then
        begin
                LeCursus := Trim(ReadTokenPipe(S,';'));
                LeRangCursus := StrToInt(Trim(ReadTokenPipe(S,';')));
                LeMillesime := Trim(ReadTokenPipe(S,';'));
                Q := OpenSQL('SELECT PCC_CODESTAGE,PCC_ORDRE FROM CURSUSSTAGE '+
                'WHERE PCC_CURSUS="'+LeCursus+'" AND PCC_RANGCURSUS='+IntToStr(LeRangCursus)+' AND PCC_SALARIE',true); //DB2
                If Not Q.eof then
                begin
                        Q.First;
                        LeStage := Q.FindField('PCC_CODESTAGE').AsString;
                        LaSession := IntToStr(Q.FindField('PCC_ORDRE').AsInteger);
                end
                else
                begin
                        LeStage := '';
                        LaSession := '';
                end;
                Ferme(Q);
        end
        else
        begin
                LeStage := Trim(ReadTokenPipe(S,';'));
                LaSession := Trim(ReadTokenPipe(S,';'));
                LeMillesime := Trim(ReadTokenPipe(S,';'));
        end;
        SetControlEnabled('PFO_SALARIE',False);

        //PT18 - Début
        MillesimeEC := '';
        TauxChargeC := 1;
        TauxChargeNC := 1;

        If ((LeStage <> '') And (LaSession <> '') And (LeMillesime <> '')) Then
        Begin
          // Récupération de la date de début et de fin de la session
          Q := OpenSQL('SELECT PSS_PREDEFINI,PSS_DATEFIN FROM SESSIONSTAGE WHERE '
                     + 'PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE="'+LaSession+'" AND PSS_MILLESIME="'+LeMillesime+'"',True);
          If Not Q.EOF Then
            begin
               FinSess := Q.FindField('PSS_DATEFIN').AsDateTime;
            end
          Else
               FinSess := iDate2099;
          Ferme(Q);

          { Sélection du millésime et des taux à appliquer par rapport à la session en cours.
            On ne compare que la date de fin de la session afin de trouver une correspondance même
            dans le cas où celle-ci se déroulerait sur 2 exercices.
          }
          Q := OpenSQL('SELECT PFE_MILLESIME,PFE_TAUXCHARGENC,PFE_TAUXCHARGEC FROM EXERFORMATION'
                     +' WHERE PFE_DATEDEBUT <= "'+UsDateTime(FinSess)+'" AND PFE_DATEFIN >= "'+UsDateTime(FinSess)+'"',True);

          If not Q.EOF Then
          Begin
              MillesimeEC := Q.FindField('PFE_MILLESIME').AsString;
              TauxChargeC := Q.FindField('PFE_TAUXCHARGEC').AsFloat;
              TauxChargeNC := Q.FindField('PFE_TAUXCHARGENC').AsFloat;
          End;
          Ferme(Q);
          End
        Else
        Begin
          // Sélection du millésime et des taux à appliquer les plus récents par rapport à l'état
          Q := OpenSQL('SELECT PFE_MILLESIME,PFE_TAUXCHARGENC,PFE_TAUXCHARGEC FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-" ORDER BY PFE_DATEDEBUT DESC',True);
          If not Q.eof then
          begin
                Q.First;
                MillesimeEC := Q.FindField('PFE_MILLESIME').AsString;
                TauxChargeC := Q.FindField('PFE_TAUXCHARGEC').AsFloat;    //PT1
                TauxChargeNC := Q.FindField('PFE_TAUXCHARGENC').AsFloat;
          end;
          Ferme(Q);
        End;
        //PT18 - Fin

        If (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'CWASSAISIEFORMATION') or (TypeSaisie = 'SAISIEFP') or (TypeSaisie = 'SAISIESESSION') Then
        begin
                If (LeStage<>'') and (LaSession<>'') and (LeMillesime<>'') Then VerfifClotureInsc(LeStage,LaSession,LeMillesime);
        end;
        If not (Ecran is TFSaisieList) Then
        begin
                TSession := THValComboBox(GetControl('SESSION'));
                If TSession<>Nil Then TSession.onExit := ExitSession;
                {$IFNDEF EAGLCLIENT}
                THStage := THDBEdit(GetControl('PFO_CODESTAGE'));
                {$ELSE}
                THStage := THEdit(GetControl('PFO_CODESTAGE'));
                {$ENDIF}
                If THStage<>Nil Then THStage.OnElipsisClick := StageElipsisClick;
        end
        Else
        begin
                TF  :=  TFSaisieList(Ecran).LeFiltre;
                if (Ecran<>nil) and (Ecran is TFSaisieList ) then
                TF.OnSetNavigate  :=  AfficheInfos;
                SetControlVisible('PanPied',True);
                SetControlVisible('PCPied',True);
                SetControlVisible('Page1',False);
                SetControlProperty('Page1','TabVisible',False);
                SetControlVisible('Page2',False);
                SetControlProperty('Page2','TabVisible',False);
                SetControlVisible('Page3',False);
                SetControlProperty('Page3','TabVisible',False);
                If TypeSaisie = 'FRAIS' Then
                begin
                        AfficheGrilleFrais;
                        SetControlVisible('BINSERT',False);
                        TFSaisieList(Ecran).Caption := 'Saisie des frais des salariés pour les sessions de formation de l''année '+MillesimeEC;
                end
                else If TypeSaisie = 'INSCRIPTION' Then
                begin
                       // SetControlVisible('PINSCSAL',True);
                       // SetControlProperty('PINSCSAL','TabVisible',true);
                        TFSaisieList(Ecran).Caption := 'Inscription des salariés aux sessions de formation de l''année '+MillesimeEC;
                end
                else If TypeSaisie = 'SAISIEWEB' Then
                begin
                        TFSaisieList(Ecran).Caption := 'Inscription des salariés aux sessions de formation de l''année '+MillesimeEC;
                    //    SetControlVisible('PINSCSAL',True);
                  //      SetControlProperty('PINSCSAL','TabVisible',true);
                end
                else If TypeSaisie = 'CWASSAISIEFORMATION' Then
                begin
                        TFSaisieList(Ecran).Caption := 'Inscription des salariés aux sessions de formation de l''année '+MillesimeEC;
                    //    SetControlVisible('PINSCSAL',True);
                   //     SetControlProperty('PINSCSAL','TabVisible',true);
//                        AfficheGrilleInsc(LeStage);  // PT2 //PT17
                end
                else If (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESESSION') Then
                begin
                    //    SetControlVisible('PINSCSAL',True);
                   //     SetControlProperty('PINSCSAL','TabVisible',true);
//                        AfficheGrilleInsc(LeStage); //PT17
                        TFSaisieList(Ecran).Caption := 'Inscription des salariés aux sessions de formation de l''année '+MillesimeEC;
                end
                else If (TypeSaisie = 'VALIDINSC') or (TypeSaisie = 'CWASVALIDINSC') Then
                begin
                        TFSaisieList(Ecran).Caption := 'Validation des inscriptions aux sessions de formation de l''année '+MillesimeEC;// PT6
                        BAccept := TToolBarButton97(GetControl('BACCEPTATION'));
                        If BAccept<>Nil Then BAccept.OnClick := BAcceptClick;
                        TF.LaGrid.MultiSelect := true;
                        BSelect := TToolBarButton97(GetControl('BSELECTALL'));
                        If BSelect <> Nil then BSelect.OnClick := BSelectionClick;
                end
                else If TypeSaisie = 'SAISIEFP' Then
                begin
                        TFSaisieList(Ecran).Caption := 'Saisie présence des salariés aux sessions formation de l''année '+MillesimeEC;
                        BPresence := TToolBarButton97(GetControl('BPRESENCE'));
                        If BPresence<>Nil then BPresence.OnClick := BPresenceClick;
                        TF.LaGrid.MultiSelect := true;
                        BSelect := TToolBarButton97(GetControl('BSELECTALL'));
                        If BSelect <> Nil then BSelect.OnClick := BSelectionClick;
                end;
                UpdateCaption(TFSaisieList(Ecran)) ;
        end;
        If TypeSaisie = 'SAISIECOMPL' Then
                begin
                        SetActiveTabSheet('PCOMPLFORM');
                        PGeneral := TTabSheet(getControl('PGeneral'));
                        PComplement := TTabSheet(GetControl('PCOMPLFORM'));
                        PSalarie := TTabSheet(GetControl('PSALARIE'));
                        PComplement.TabVisible := True;
                        PGeneral.tabVisible := False;
                        PSalarie.TabVisible := False;
                        Pgeneral.Visible := False;
                        PSalarie.Visible := False;
                end;
        If (Ecran is TFSaisieList) Then
        begin
                THDate := THEdit(GetControl('PFO_DATEDEBUT'));
                If thdATE<>NIL Then THDate.OnElipsisClick := DateElipsisClick;
                THDate := THEdit(GetControl('PFO_DATEFIN'));
                If thdATE<>NIL Then THDate.OnElipsisClick := DateElipsisClick;
                THDate := THEdit(GetControl('PFO_DATEACCEPT'));
                If thdATE<>NIL Then THDate.OnElipsisClick := DateElipsisClick;
                THDate := THEdit(GetControl('PFO_REPORTELE'));
                If thdATE<>NIL Then THDate.OnElipsisClick := DateElipsisClick;
                THDate := THEdit(GetControl('PFO_REFUSELE'));
                If thdATE<>NIL Then THDate.OnElipsisClick := DateElipsisClick;
                Edit := THEdit(GetControl('PFO_NOMSALARIE'));
                If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
                Edit := THEdit(GetControl('PFO_SALARIE'));
                If Edit <> Nil then Edit.OnElipsisClick := CodeSalarieElipsisClick;
        end;
        BCompteurs := TToolBarButton97(GetControl('BCOMPTEURS'));
        If BCompteurs <> Nil Then BCompteurs.OnClick := VoirCompteursDIF;
        BMvts := TToolBarButton97(GetControl('BMOUVEMENTS'));
        If BMvts <> Nil Then BMvts.OnClick := VoirMouvementsGenere;
        {$IFDEF EMANAGER}
        TFFiche(Ecran).FTYpeAction := TaConsult;
        {$ENDIF}
end ;

procedure TOM_FORMATIONS.ExitSession(Sender : TObject);
var CodeStage,Ordre,Millesime : String;
    Q : TQuery;
begin
        iF THValComboBox(Sender).Name<>'SESSION' Then Exit;
        If THValComboBox(Sender).Value = '' Then Exit;
        If THValComboBox(Sender).Value<>InitOrdre Then
        begin
                CodeStage := GetField('PFO_CODESTAGE');
                Ordre := THValComboBox(Sender).Value;
                Millesime := GetField('PFO_MILLESIME');
                Q := OpenSQL('SELECT PSS_LIBELLE,PSS_DATEDEBUT,PSS_DEBUTDJ,PSS_DATEFIN,PSS_FINDJ,PSS_CENTREFORMGU'+
                ' FROM SESSIONSTAGE WHERE PSS_MILLESIME="'+Millesime+'"'+
                ' AND PSS_CODESTAGE="'+CodeStage+'" AND PSS_ORDRE='+Ordre,True); //DB2
                if not Q.eof then
                begin
                        SetField('PFO_DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsDateTime);
                        SetField('PFO_DEBUTDJ',Q.FindField('PSS_DEBUTDJ').AsString);
                        SetField('PFO_DATEFIN',Q.FindField('PSS_DATEFIN').AsDateTime);
//                        SetField('PFO_DATEPAIE',V_PGI.DateEntree);
                        SetField('PFO_FINDJ',Q.FindField('PSS_FINDJ').AsString);
                        SetField('PFO_CENTREFORMGU',Q.FindField('PSS_CENTREFORMGU').AsString);
                end;
                Ferme(Q);
                InitOrdre := Ordre;
                SetField('PFO_ORDRE',StrToInt(Ordre));
                ForceUpdate;
        end;
end;

procedure TOM_FORMATIONS.StageElipsisClick(Sender : TObject);
var  StWhere,StOrder : String;
begin
        StWhere := 'PST_ACTIF="X" AND'+
        ' (PST_MILLESIME="'+MillesimeEC+'") OR (PST_MILLESIME="0000" AND '+
        'PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'"))';
        StOrder := 'PST_MILLESIME,PST_LIBELLE';
        {$IFNDEF EAGLCLIENT}
        LookupList(THDBEdit(Sender),'Liste des stages','STAGE LEFT JOIN LIEUFORMATION ON PST_LIEUFORM=PLF_LIEUFORM','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,PLF_LIEUFORM,PST_MILLESIME',StWhere,StOrder, True,-1);
        {$ELSE}
        LookupList(THEdit(Sender),'Liste des stages','STAGE LEFT JOIN LIEUFORMATION ON PST_LIEUFORM=PLF_LIEUFORM','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,PLF_LIEUFORM,PST_MILLESIME',StWhere,StOrder, True,-1);
        {$ENDIF}
end;

procedure TOM_FORMATIONS.AfficheGrilleFrais;
var NbFrais,i,a,b,c : Integer;
begin
        NbFrais := VH_Paie.PGFNbFraisLibre;
        For i := nbFrais+1 to 6  do
        begin
                a := 3+((i-1)*3);
                b := 4+((i-1)*3);
                c := 5+((i-1)*3);
                TF.LaGrid.DeleteCol(a);
                TF.LaGrid.DeleteCol(b);
                TF.LaGrid.DeleteCol(c);
        end;
        For i := 1 to nbFrais Do
        begin
                If i=1 Then
                begin
                        TF.LaGrid.ColAligns[3] := TaRightJustify;
                        TF.LaGrid.CellValues[4,0] := 'Qté';
                        TF.LaGrid.ColAligns[4] := TaRightJustify;
                        TF.LaGrid.CellValues[5,0] := 'Montant';
                        TF.LaGrid.ColAligns[5] := TaRightJustify;
                        TF.LaGrid.ColColors[5] := clBlue;
                end;
                If i=2 Then
                begin
                        TF.LaGrid.ColAligns[6] := TaRightJustify;
                        TF.LaGrid.CellValues[7,0] := 'Qté';
                        TF.LaGrid.ColAligns[7] := TaRightJustify;
                        TF.LaGrid.CellValues[8,0] := 'Montant';
                        TF.LaGrid.ColAligns[8] := TaRightJustify;
                        TF.LaGrid.ColColors[8] := clBlue	;
                end;
                If i=3 Then
                begin
                        TF.LaGrid.ColAligns[9] := TaRightJustify;
                        TF.LaGrid.CellValues[10,0] := 'Qté';
                        TF.LaGrid.ColAligns[10] := TaRightJustify;
                        TF.LaGrid.CellValues[11,0] := 'Montant';
                        TF.LaGrid.ColAligns[11] := TaRightJustify;
                        TF.LaGrid.ColColors[11] := clBlue;
                end;
                If i=4 Then
                begin
                        TF.LaGrid.ColAligns[12] := TaRightJustify;
                        TF.LaGrid.CellValues[13,0] := 'Qté';
                        TF.LaGrid.ColAligns[13] := TaRightJustify;
                        TF.LaGrid.CellValues[14,0] := 'Montant';
                        TF.LaGrid.ColAligns[14] := TaRightJustify;
                        TF.LaGrid.ColColors[14] := clBlue;
                end;
                If i=5 Then
                begin
                        TF.LaGrid.ColAligns[15] := TaRightJustify;
                        TF.LaGrid.CellValues[16,0] := 'Qté';
                        TF.LaGrid.ColAligns[16] := TaRightJustify;
                        TF.LaGrid.CellValues[17,0] := 'Montant';
                        TF.LaGrid.ColAligns[17] := TaRightJustify;
                        TF.LaGrid.ColColors[17] := clBlue;
                end;
                If i=6 Then
                begin
                        TF.LaGrid.ColAligns[18] := TaRightJustify;
                        TF.LaGrid.CellValues[19,0] := 'Qté';
                        TF.LaGrid.ColAligns[19] := TaRightJustify;
                        TF.LaGrid.CellValues[20,0] := 'Montant';
                        TF.LaGrid.ColAligns[20] := TaRightJustify;
                        TF.LaGrid.ColColors[20] := clBlue;
                end;
        end;
end;

procedure TOM_FORMATIONS.BPresenceClick(Sender : TObject);
Var i : Integer;
    DureeStage : Double;
    Q : TQuery;
    Stage,Millesime,Session : String;
    TauxHoraire,Montant : Double;
begin
        //DEBUT PT15
        Stage := LeStage;
        Millesime := LeMillesime;
        Session := LaSession;
        //FIN PT15
        Q := OpenSQL('SELECT PSS_DUREESTAGE'+
        ' FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Stage+'"'+
        ' AND PSS_ORDRE='+Session+' AND PSS_MILLESIME="'+Millesime+'"',True);  //DB2
        DureeStage := 0;
        if not Q.eof then DureeStage := Q.FindField('PSS_DUREESTAGE').AsFloat;
        Ferme(Q);
        TF.DisableTOM;
        if ((TF.LaGrid.nbSelected)>0) AND (not TF.LaGrid.AllSelected ) then
        begin
                InitMoveProgressForm (NIL,'Début du traitement', 'Veuillez patienter SVP ...',TF.LaGrid.nbSelected,FALSE,TRUE);
                InitMove(TF.LaGrid.NbSelected,'');
                TF.StartUpdate;
                for i:=0 to TF.LaGrid.NbSelected-1 do
                begin
                        TF.LaGrid.GotoLeBOOKMARK(i);
                        TF.SelectRecord(TF.LaGrid.Row);
                        MoveCur(False);
                        VerifReaPrevisionnel(TF.GetValue('PFO_CODESTAGE'),TF.GetValue('PFO_SALARIE'),TF.GetValue('PFO_EFFECTUE'));
                        TF.PutValue('PFO_EFFECTUE','X');
                        TF.PutValue('PFO_NBREHEURE',DureeStage);
                        if (DefautHTpsT) or (TF.GetValue('PFO_TYPEPLANPREV') = 'PLF') then
                        begin
                             TF.PutValue('PFO_HTPSTRAV',DureeStage);
                             TF.PutValue('PFO_HTPSNONTRAV',0);
                        end
                        else
                        begin
                             TF.PutValue('PFO_HTPSTRAV',0);
                             TF.PutValue('PFO_HTPSNONTRAV',DureeStage);
                        end;
                        TauxHoraire := 0;
                        if VH_Paie.PGForValoSalaire = 'VCR' then TauxHoraire := ForTauxHoraireReel(TF.GetValue('PFO_SALARIE'));
                        if VH_Paie.PGForValoSalaire = 'VCC' then TauxHoraire := ForTauxHoraireCategoriel(TF.GetValue('PFO_LIBEMPLOIFOR'),MillesimeEC);
                        Montant := TauxHoraire*dureeStage;
                        
                        //PT21
                        If PGBundleHierarchie Then
                        	Q := OpenSQL('SELECT PSA_DADSCAT FROM '+GetBase(GetBaseSalarie(GetField('PFO_SALARIE')),'SALARIES')+' WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True)
                        Else
                        	Q := OpenSQL('SELECT PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True);  //PT13
                        If (Not Q.EOF) And ((Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02')) then 
                        	Montant := Arrondi(Montant * TauxChargeC,2)
                        else 
                        	Montant := Arrondi(Montant * TauxChargeNC,2);
                        Ferme(Q);
                        TF.PutValue('PFO_COUTREELSAL',Montant);
                        TF.Post;
                        MajCompetences(TF.GetValue('PFO_SALARIE'),LeStage,LaSession,LeMillesime,True); //PT22
                        MoveCurProgressForm(TF.GetValue('PFO_SALARIE'));
                end;
                TF.EndUpdate;
                FiniMove;
                FiniMoveProgressForm;
        end;
        if TF.LaGrid.AllSelected then
        begin
                InitMoveProgressForm (NIL,'Début du traitement', 'Veuillez patienter SVP ...',TF.LaGrid.RowCount,FALSE,TRUE);
                InitMove(TF.LaGrid.NbSelected,'');
                TF.StartUpdate;
                for i:=1 to TF.LaGrid.RowCount do
                begin
                        TF.SelectRecord(i);
                        MoveCur(False);
                        TF.PutValue('PFO_EFFECTUE','X');
                        TF.PutValue('PFO_NBREHEURE',DureeStage);
                         if (DefautHTpsT) or (TF.GetValue('PFO_TYPEPLANPREV') = 'PLF') then
                        begin
                             TF.PutValue('PFO_HTPSTRAV',DureeStage);
                             TF.PutValue('PFO_HTPSNONTRAV',0);
                        end
                        else
                        begin
                             TF.PutValue('PFO_HTPSTRAV',0);
                             TF.PutValue('PFO_HTPSNONTRAV',DureeStage);
                        end;
                        VerifReaPrevisionnel(TF.GetValue('PFO_CODESTAGE'),TF.GetValue('PFO_SALARIE'),TF.GetValue('PFO_EFFECTUE'));
                        TauxHoraire := 0;
                        if VH_Paie.PGForValoSalaire = 'VCR' then TauxHoraire := ForTauxHoraireReel(TF.GetValue('PFO_SALARIE'));
                        if VH_Paie.PGForValoSalaire = 'VCC' then TauxHoraire := ForTauxHoraireCategoriel(TF.GetValue('PFO_LIBEMPLOIFOR'),MillesimeEC);
                        Montant := TauxHoraire*dureeStage;
                        
                        //PT21
                        If PGBundleHierarchie Then
                        	Q := OpenSQL('SELECT PSA_DADSCAT FROM '+GetBase(GetBaseSalarie(GetField('PFO_SALARIE')),'SALARIES')+' WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True)
                        Else
                        	Q := OpenSQL('SELECT PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True);  //PT13

                        If (Not Q.EOF) And ((Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02')) then 
							Montant := Arrondi(Montant * TauxChargeC,2)
                        else 
                        	Montant := Arrondi(Montant * TauxChargeNC,2);
                        Ferme(Q);
                        TF.PutValue('PFO_COUTREELSAL',Montant);
                        TF.Post;
                        MajCompetences(TF.GetValue('PFO_SALARIE'),LeStage,LaSession,LeMillesime,True); //PT22
                        MoveCurProgressForm(TF.GetValue('PFO_SALARIE'));
                end;
                TF.EndUpdate;
                FiniMove;
                FiniMoveProgressForm;
        end;
        TF.EnableTOM;
        TF.LaGrid.ClearSelected;
        MAJCoutsFormation(Millesime,Stage,Session);
        SetControlProperty('BSELECTALL','Down',False);
end;

procedure TOM_FORMATIONS.BAcceptClick(Sender : TObject);
var i : integer;
begin
        if ((TF.LaGrid.nbSelected)>0) AND (not TF.LaGrid.AllSelected ) then
        begin
                InitMoveProgressForm (NIL,'Début du traitement', 'Veuillez patienter SVP ...',TF.LaGrid.nbSelected,FALSE,TRUE);
                InitMove(TF.LaGrid.NbSelected,'');
                TF.StartUpdate;
                for i:=0 to TF.LaGrid.NbSelected-1 do
                begin
                        TF.LaGrid.GotoLeBOOKMARK(i);
                        TF.SelectRecord(TF.LaGrid.Row);
                        MoveCur(False);
                        TF.PutValue('PFO_DATEACCEPT',Date);
                        TF.PutValue('PFO_ETATINSCFOR','VAL');
                        TF.Post;
                        MoveCurProgressForm(TF.GetValue('PFO_SALARIE'));
                end;
                TF.EndUpdate;
                FiniMove;
                FiniMoveProgressForm;
        end;
        if TF.LaGrid.AllSelected then
        begin
                InitMoveProgressForm (NIL,'Début du traitement', 'Veuillez patienter SVP ...',TF.LaGrid.RowCount,FALSE,TRUE);
                InitMove(TF.LaGrid.NbSelected,'');
                TF.StartUpdate;
                for i:=1 to TF.LaGrid.RowCount do
                begin
                        TF.SelectRecord(i);
                        MoveCur(False);
                        TF.PutValue('PFO_DATEACCEPT',Date);
                        TF.PutValue('PFO_ETATINSCFOR','VAL');
                        TF.Post;
                        MoveCurProgressForm(TF.GetValue('PFO_SALARIE'));
                end;
                TF.EndUpdate;
                FiniMove;
                FiniMoveProgressForm;
        end;
        TF.LaGrid.ClearSelected;
        SetControlProperty('BSELECTALL','Down',False);
end;

procedure TOM_FORMATIONS.VerfifClotureInsc(Stage,Session,Millesime : String);
var Q : TQuery;
begin
        if (Ecran is TFSaisieList) and ((TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'CWASSAISIEFORMATION') or (TypeSaisie = 'SAISIESESSION')) Then
        begin
                If TF=Nil Then TF  :=  TFSaisieList(Ecran).LeFiltre;
                Q := OpenSQL('SELECT PSS_CLOTUREINSC FROM SESSIONSTAGE WHERE '+
                'PSS_CODESTAGE="'+Stage+'" AND PSS_ORDRE='+Session+' AND PSS_MILLESIME="'+Millesime+'"',True);  //DB2
                If Q.eof then
                begin
                        Ferme(Q);
                        Exit;
                end;
                If Q.FindField('PSS_CLOTUREINSC').AsString = 'X' Then
                begin
                        SetControlVisible('INFOCLOTURE',true);
                        TFSaisieList(Ecran).FTypeAction := taConsult;
                        TF.SaisieEnGrid := False;
                        SetControlVisible('BDELETE',False);
                        SetControlVisible('BINSERT',False);
                        SetControlEnabled('BINSERT',False);
                        //PT21
                        SetControlVisible('BMULTISAL',     False);
                        SetControlVisible('BRECUPERATION', False);
                end
                Else
                begin
                        SetControlVisible('INFOCLOTURE',False);
                        TF.SaisieEnGrid := True;
                        SetcontrolEnabled('BINSERT',True);
                        SetControlVisible('BDELETE',True);
                end;
                Ferme(Q);
        end;
end;

procedure TOM_FORMATIONS.DateElipsisclick(Sender : TObject);
var key : char;
begin
        key  :=  '*';
        ParamDate (Ecran, Sender, Key);
end;

{Procedure TOM_FORMATIONS.AfficheGrilleInsc(Stage : String); //PT17
var TobSal : Tob;
    QExer,QS : TQuery;
    GSalaries : THGrid;
    i : Integer;
    DateDebut,DateFin : TDateTime;
    HMTrad :  THSystemMenu;
    StWhere : String;
begin
        StWhere := '';
        {$IFDEF EMANAGER}
        {StWhere := ' AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        {$ENDIF}
        {GSalaries := THGrid(GetControl('GSALARIES'));
        For i := 1 to GSalaries.RowCount-1 do
        begin
                GSalaries.DeleteRow(i);
        end;
        GSalaries.RowCount := 2;
        GSalaries.ColCount := 8;
        GSalaries.SortEnabled := True;
        //GSalaries.CellValues[0,0] := 'Type de plan';
        //GSalaries.CellValues[1,0] := 'Matricule';
        //GSalaries.CellValues[2,0] := 'Nom';
        //GSalaries.CellValues[3,0] := 'Etablissement';
        //GSalaries.CellValues[4,0] := 'Libellé Emploi';
        //GSalaries.CellValues[5,0] := 'Sur tps trav.';
        //GSalaries.CellValues[6,0] := 'Hors tps trav.';
        //GSalaries.CellValues[7,0] := 'Typologie';
        GSalaries.ColWidths[7] := -1;
        GSalaries.ColFormats[0] := 'CB=PGTYPEPLANPREV';
        GSalaries.ColFormats[3] := 'CB=TTETABLISSEMENT';
        GSalaries.ColFormats[4] := 'CB=PGLIBEMPLOI';
        GSalaries.ColFormats[5] := '# ##0.00';
        GSalaries.ColFormats[6] := '# ##0.00';
        GSalaries.FixedRows := 1;
        Qexer := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+MillesimeEC+'"',True);
        DateDebut := Date;
        DateFin := Date;
        If not QExer.eof then
        begin
                DateDebut := QExer.FindField('PFE_DATEDEBUT').AsDateTime;
                DateFin := QExer.FindField('PFE_DATEFIN').AsDateTime;
        end;
        Ferme(QExer);
        QS := OpenSQL('SELECT PFI_HTPSTRAV,PFI_HTPSNONTRAV,PFI_TYPOFORMATION,PFI_TYPEPLANPREV,PFI_SALARIE,PFI_NBINSC,PFI_ETABLISSEMENT,PFI_LIBEMPLOIFOR,PFI_MILLESIME,PFI_LIBELLE '+
        'FROM INSCFORMATION'+
        ' WHERE PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE="'+Stage+'" AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_SALARIE<>""'+
        ' AND PFI_SALARIE NOT IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_CODESTAGE="'+Stage+'" AND'+
        ' PFO_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'")'+StWhere,True);
        TobSal := Tob.Create('Les Salariés',Nil,-1);
        TobSal.LoadDetailDB('INSCFORMATION','','',QS,False);
        Ferme(QS);
        TobSal.PutGridDetail(GSalaries,False,True,'PFI_TYPEPLANPREV;PFI_SALARIE;PFI_LIBELLE;PFI_ETABLISSEMENT;PFI_LIBEMPLOIFOR;PFI_HTPSTRAV;PFI_HTPSNONTRAV;PFI_TYPOFORMATION',False);
        TobSal.Free;
        HMTrad:=THSystemMenu(GetControl('HMTrad'));
        HMTrad.ResizeGridColumns (GSalaries);
end;  }

Procedure TOM_FORMATIONS.RecupInfosSalarie(TS : String);
var Q : TQuery;
    Session,Stage,Millesime : String;
//    TauxHoraire,Montant : Double;
    DatePaie,DateSession : TDateTime;
begin
        {$IFDEF EMANAGER}
        If Not ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND PSE_SALARIE="'+GetField('PFO_SALARIE')+'"') then
        begin
                PGIBox('Vous n''êtes pas responsable de ce salarié, veuillez saisir un autre matricule',Ecran.Caption);
                SetFocusControl('PFO_SALARIE');
                Exit;
        end;
        {$ENDIF}
        Q := OpenSQL('SELECT MAX (PPU_DATEFIN) MAXDATE FROM PAIEENCOURS WHERE PPU_SALARIE="'+GetField('PFO_SALARIE')+'"',True);
        If Not Q.Eof then DatePaie := Q.FindField('MAXDATE').AsDateTime
        Else DatePaie := V_PGI.DateEntree;
        DatePaie := PlusMois(DatePaie,1);
        DatePaie := FinDeMois(DatePaie);
        Ferme(Q);
        If DatePaie < V_PGI.dateEntree then DatePaie := FinDeMois(V_PGI.DateEntree);
        SetField('PFO_DATEPAIE',DatePaie);
        //DEBUT PT15
        Q := OpenSQL('SELECT PSS_DATEDEBUT FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE='+LaSession+' AND PSS_MILLESIME="'+LeMillesime+'"',True);
        If not Q.Eof then DateSession := Q.FindField('PSS_DATEDEBUT').AsDateTime
        else dateSession := IDate1900;
        Ferme(Q);
        //FIN PT15
        If (PGBundleHierarchie) and (PGDroitMultiForm) then
        begin
          Q := OpenSQL('SELECT PSI_NODOSSIER,PSI_DATESORTIE,PSI_TRAVAILN1,PSI_TRAVAILN2,PSI_TRAVAILN3,PSI_TRAVAILN4,PSI_CODESTAT,PSI_LIBELLE,PSI_ETABLISSEMENT,PSI_LIBELLEEMPLOI,PSI_PRENOM,PSE_RESPONSFOR'+
           ',PSI_LIBREPCMB1,PSI_LIBREPCMB2,PSI_LIBREPCMB3,PSI_LIBREPCMB4 FROM INTERIMAIRES LEFT JOIN DEPORTSAL ON PSI_INTERIMAIRE=PSE_SALARIE WHERE PSI_INTERIMAIRE="'+GetField('PFO_SALARIE')+'"',True);
          if not Q.eof then
          begin
                  If (Q.FindField('PSI_DATESORTIE').AsDateTime>10) and  (Q.FindField('PSI_DATESORTIE').AsDateTime<=DateSession) then
                  begin
                          PGIBox('Saisie impossible pour ce salarié, #13#10 sa date de sortie est supérieure à la date de début de la session',Ecran.Caption);
                          SetFocusControl('PFO_SALARIE');
                          SetField('PFO_SALARIE','');
                          Exit;
                  end;
                  SetField('PFO_TRAVAILN1',		Q.FindField('PSI_TRAVAILN1').AsString);
                  SetField('PFO_TRAVAILN2',		Q.FindField('PSI_TRAVAILN2').AsString);
                  SetField('PFO_TRAVAILN3',		Q.FindField('PSI_TRAVAILN3').AsString);
                  SetField('PFO_TRAVAILN4',		Q.FindField('PSI_TRAVAILN4').AsString);
                  SetField('PFO_CODESTAT',		Q.FindField('PSI_CODESTAT').AsString);
                  SetField('PFO_NOMSALARIE',	Q.FindField('PSI_LIBELLE').AsString);
                  SetField('PFO_PRENOM',		Q.FindField('PSI_PRENOM').AsString);
                  SetField('PFO_ETABLISSEMENT',	Q.FindField('PSI_ETABLISSEMENT').AsString);
                  SetField('PFO_LIBEMPLOIFOR',	Q.FindField('PSI_LIBELLEEMPLOI').AsString);
                  If TypeSaisie = 'CWASSAISIEFORMATION' then SetField('PFO_RESPONSFOR',V_PGI.UserSalarie)
                  else SetField('PFO_RESPONSFOR',Q.FindField('PSE_RESPONSFOR').AsString);
                  SetField('PFO_LIBREPCMB1',	Q.findField('PSI_LIBREPCMB1').AsString);
                  SetField('PFO_LIBREPCMB2',	Q.findField('PSI_LIBREPCMB2').AsString);
                  SetField('PFO_LIBREPCMB3',	Q.findField('PSI_LIBREPCMB3').AsString);
                  SetField('PFO_LIBREPCMB4',	Q.findField('PSI_LIBREPCMB4').AsString);
                  SetField('PFO_NODOSSIER',		V_PGI.NoDossier);
                  SetField('PFO_PREDEFINI',		'DOS');
          end
          Else
          begin
                  PGIBox('Ce salarié n''existe pas, veuillez modifier le matricule saisi',Ecran.Caption);
                  SetFocusControl('PFO_SALARIE');
                  SetField('PFO_SALARIE','');
                  Exit;
          end;

          SetField('PFO_NODOSSIER',Q.FindField('PSI_NODOSSIER').AsString);
          SetField('PFO_PREDEFINI','DOS');
          Ferme(Q);
        end
        else
        begin
          Q := OpenSQL('SELECT PSA_DATESORTIE,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_PRENOM,PSE_RESPONSFOR'+
           ',PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4 FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True);
          if not Q.eof then
          begin
                  If (Q.FindField('PSA_DATESORTIE').AsDateTime>10) and  (Q.FindField('PSA_DATESORTIE').AsDateTime<=DateSession) then
                  begin
                          PGIBox('Saisie impossible pour ce salarié, #13#10 sa date de sortie est supérieure à la date de début de la session',Ecran.Caption);
                          SetFocusControl('PFO_SALARIE');
                          SetField('PFO_SALARIE','');
                          Exit;
                  end;
                  SetField('PFO_TRAVAILN1',Q.FindField('PSA_TRAVAILN1').AsString);
                  SetField('PFO_TRAVAILN2',Q.FindField('PSA_TRAVAILN2').AsString);
                  SetField('PFO_TRAVAILN3',Q.FindField('PSA_TRAVAILN3').AsString);
                  SetField('PFO_TRAVAILN4',Q.FindField('PSA_TRAVAILN4').AsString);
                  SetField('PFO_CODESTAT',Q.FindField('PSA_CODESTAT').AsString);
                  SetField('PFO_NOMSALARIE',Q.FindField('PSA_LIBELLE').AsString);
                  SetField('PFO_PRENOM',Q.FindField('PSA_PRENOM').AsString);
                  SetField('PFO_ETABLISSEMENT',Q.FindField('PSA_ETABLISSEMENT').AsString);
                  SetField('PFO_LIBEMPLOIFOR',Q.FindField('PSA_LIBELLEEMPLOI').AsString);
                  If TypeSaisie = 'CWASSAISIEFORMATION' then SetField('PFO_RESPONSFOR',V_PGI.UserSalarie)
                  else SetField('PFO_RESPONSFOR',Q.FindField('PSE_RESPONSFOR').AsString);
                  SetField('PFO_LIBREPCMB1',Q.findField('PSA_LIBREPCMB1').AsString);
                  SetField('PFO_LIBREPCMB2',Q.findField('PSA_LIBREPCMB2').AsString);
                  SetField('PFO_LIBREPCMB3',Q.findField('PSA_LIBREPCMB3').AsString);
                  SetField('PFO_LIBREPCMB4',Q.findField('PSA_LIBREPCMB4').AsString);
                  SetField('PFO_NODOSSIER',V_PGI.NoDossier);
                  SetField('PFO_PREDEFINI','STD'); //PT23
          end
          Else
          begin
                  PGIBox('Ce salarié n''existe pas, veuillez modifier le matricule saisi',Ecran.Caption);
                  SetFocusControl('PFO_SALARIE');
                  SetField('PFO_SALARIE','');
                  Exit;
          end;
          Ferme(Q);
        end;
        //DEBUT PT15
        Stage := LeStage;
        Session := LaSession;
        Millesime := Lemillesime;
        //FIN PT15
        SetField('PFO_CODESTAGE',Stage);
        SetField('PFO_ORDRE',Session);
        SetField('PFO_MILLESIME',Millesime);
        Q := OpenSQL('SELECT PSS_PGTYPESESSION,PSS_NOSESSIONMULTI,PSS_DATEDEBUT,PSS_DATEFIN,PSS_DUREESTAGE,PSS_ORGCOLLECTSGU,PSS_DEBUTDJ,PSS_DEBUTDJ,PSS_LIEUFORM,PSS_CENTREFORMGU,'+
           'PSS_HEUREDEBUT,PSS_HEUREFIN,PSS_FORMATION1,PSS_FORMATION2,PSS_FORMATION3,PSS_FORMATION4,PSS_FORMATION5,PSS_FORMATION6,PSS_FORMATION7,PSS_FORMATION8'+
           ',PSS_INCLUSDECL'+
           ',PSS_NATUREFORM,PSS_FINDJ,PSS_DEBUTDJ FROM SESSIONSTAGE LEFT JOIN STAGE ON '+
           ' PSS_MILLESIME=PST_MILLESIME AND PST_CODESTAGE=PSS_CODESTAGE'+
           ' WHERE PSS_ORDRE='+Session+' AND PSS_CODESTAGE="'+Stage+'" AND PSS_MILLESIME="'+Millesime+'"',True);  //DB2
        if not Q.eof then
        begin
                SetField('PFO_DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsDateTime);
                SetField('PFO_DATEFIN',Q.FindField('PSS_DATEFIN').AsDateTime);
//                SetField('PFO_DATEPAIE',V_PGI.DateEntree);
                SetField('PFO_NATUREFORM',Q.FindField('PSS_NATUREFORM').AsString);
                SetField('PFO_LIEUFORM',Q.FindField('PSS_LIEUFORM').AsString);
                SetField('PFO_CENTREFORMGU',Q.FindField('PSS_CENTREFORMGU').AsString);//DB2
                SetField('PFO_DEBUTDJ',Q.FindField('PSS_DEBUTDJ').AsString);
                SetField('PFO_FINDJ',Q.FindField('PSS_FINDJ').AsString);
                SetField('PFO_ORGCOLLECTGU',Q.FindField('PSS_ORGCOLLECTSGU').AsString);
                SetField('PFO_HEUREDEBUT',Q.FindField('PSS_HEUREDEBUT').AsDateTime);
                SetField('PFO_HEUREFIN',Q.FindField('PSS_HEUREFIN').AsDateTime);
                SetField('PFO_FORMATION1',Q.FindField('PSS_FORMATION1').AsString);
                SetField('PFO_FORMATION2',Q.FindField('PSS_FORMATION2').AsString);
                SetField('PFO_FORMATION3',Q.FindField('PSS_FORMATION3').AsString);
                SetField('PFO_FORMATION4',Q.FindField('PSS_FORMATION4').AsString);
                SetField('PFO_FORMATION5',Q.FindField('PSS_FORMATION5').AsString);
                SetField('PFO_FORMATION6',Q.FindField('PSS_FORMATION6').AsString);
                SetField('PFO_FORMATION7',Q.FindField('PSS_FORMATION7').AsString);
                SetField('PFO_FORMATION8',Q.FindField('PSS_FORMATION8').AsString);
                SetField('PFO_INCLUSDECL',Q.FindField('PSS_INCLUSDECL').AsString);
                SetField('PFO_PGTYPESESSION',Q.FindField('PSS_PGTYPESESSION').AsString);
                SetField('PFO_NOSESSIONMULTI',Q.FindField('PSS_NOSESSIONMULTI').AsInteger);
                If TypeSaisie = 'SAISIEFP' Then
                begin
                        SetField('PFO_EFFECTUE','X');
                        SetField('PFO_NBREHEURE',Q.FindField('PSS_DUREESTAGE').AsFloat);
                        if (DefautHTpsT) or (GetField('PFO_TYPEPLANPREV') = 'PLF') then
                        begin
                             SetField('PFO_HTPSTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                             SetField('PFO_HTPSNONTRAV',0);
                        end
                        else
                        begin
                             SetField('PFO_HTPSTRAV',0);
                             SetField('PFO_HTPSNONTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                        end;
                        //DEBUT PT8
{                        TauxHoraire := 0;
                        if VH_Paie.PGForValoSalaire = 'VCR' then TauxHoraire := ForTauxHoraireReel(GetField('PFO_SALARIE'));
                        if VH_Paie.PGForValoSalaire = 'VCC' then TauxHoraire := ForTauxHoraireCategoriel(GetField('PFO_LIBEMPLOIFOR'),MillesimeEC);
                        Montant := TauxHoraire*(Q.FindField('PSS_DUREESTAGE').AsFloat);
                        SetField('PFO_COUTREELSAL',Montant);}
                        MettreAjourLeSalaire;
                        //FIN PT8
                end;
        end;
        Ferme(Q);
        If (TypeSaisie = 'VALIDINSC') or (TypeSaisie = 'CWASVALIDINSC') or (TypeSaisie = 'SAISIEFP') Then SetField('PFO_DATEACCEPT',Date);
end;

procedure TOM_FORMATIONS.BSelectionClick (Sender : TObject);
var BSelect : TToolBarButton97;
begin
        BSelect := TToolBarButton97(GetControl('BSELECTALL'));
        If BSelect.Down then TF.LaGrid.AllSelected:=True
        else TF.LaGrid.AllSelected:=False;
end;

procedure TOM_FORMATIONS.SalarieElipsisClick (Sender : TObject);
var Where,NomSal : String;
begin
  If (PGBundleHierarchie) and (PGDroitMultiForm) then
  begin
    NomSal := GetControlText('PFO_NOMSALARIE');
    //PT21
    (*
    Where := 'PSI_LIBELLE LIKE"'+NomSal+'%"';
    Where := RecupClauseHabilitationLookupList (Where);  //PT16
    SetControltext('PFO_NOMSALARIE','');
    LookupList(THEdit(Sender),'Liste des salariés','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM,PSI_NODOSSIER',Where,'PSI_INTERIMAIRE', True,-1);
    *)
    ElipsisSalarieMultidos(Sender);
    If IsNumeric(GetControlText('PFO_NOMSALARIE')) then
    begin
      SetField('PFO_SALARIE',GetControlText('PFO_NOMSALARIE'));
      If GetField('PFO_SALARIE') <> '' then RecupInfosSalarie(typeSaisie);
      If GetField('PFO_SALARIE') = '' then SetControltext('PFO_NOMSALARIE',NomSal);
    end
    Else SetControlText('PFO_NOMSALARIE',NomSal);
  end
  else
  begin
      NomSal := GetControlText('PFO_NOMSALARIE');
      Where := 'PSA_LIBELLE LIKE"'+NomSal+'%"';
      Where := RecupClauseHabilitationLookupList (Where);  //PT16
      SetControltext('PFO_NOMSALARIE','');
      LookupList(THEdit(Sender),'Liste des salariés','SALARIES','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',Where,'PSA_LIBELLE', True,-1);
      If IsNumeric(GetControlText('PFO_NOMSALARIE')) then
      begin
              SetField('PFO_SALARIE',GetControlText('PFO_NOMSALARIE'));
              If GetField('PFO_SALARIE') <> '' then RecupInfosSalarie(typeSaisie);
              If GetField('PFO_SALARIE') = '' then SetControltext('PFO_NOMSALARIE',NomSal);
      end
      Else SetControlText('PFO_NOMSALARIE',NomSal);
  end;
end;

procedure TOM_FORMATIONS.CodeSalarieElipsisClick (Sender : TObject);
var StFrom,StWhere : String;
begin
  //DEBUT PT20
  If (PGBundleHierarchie) and (PGDroitMultiForm) then
  begin
    If IsValidDate(GetControlText('DATEDEBUT')) then StWhere :=  ' AND (PSI_DATESORTIE IS NULL OR PSI_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSI_DATESORTIE >="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'")'
    Else StWhere := '';
    //StFrom := 'INTERIMAIRES';
    StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT16
    //LookupList(THEdit(Sender), 'Liste des salariés',StFrom,'PSI_INTERIMAIRE', 'PSI_LIBELLE,PSI_PRENOM,PSI_ETABLISSEMENT,PSI_NODOSSIER', StWhere, 'PSI_INTERIMAIRE', TRUE, -1);
    ElipsisSalarieMultiDos(Sender, StWhere);
  end
  else
  begin
    If IsValidDate(GetControlText('DATEDEBUT')) then StWhere :=  '(PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE >="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'")'
    Else StWhere := '';
    StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
    {$IFDEF EMANAGER}
    If StWhere <> '' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'
    else StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
    {$ENDIF}
    StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT16
    LookupList(THEdit(Sender), 'Liste des salariés',StFrom,'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT', StWhere, 'PSA_SALARIE', TRUE, -1);
  end;
  //FIN PT20
end;

//PT22
(*procedure TOM_FORMATIONS.MajCompetences ( Matricule : String;Present : Boolean);
var Stage,Competence,Session : String;
    Q : TQuery;
    IMax,i : Integer;
    TobCompetStage,T : Tob;
    DateValid : TDateTime;
begin
        //DEBUT PT15
        Q := OpenSQL('SELECT PSS_DATEFIN FROM SESSIONSTAGE WHERE PSS_ORDRE='+LaSession+' AND PSS_CODESTAGE="'+LeStage+'" AND PSS_MILLESIME="'+LeMillesime+'"',True);
        If Not Q.Eof then DateValid := Q.FindField('PSS_DATEFIN').AsDateTime
        Else DateValid := IDate1900;
        Ferme(Q);
        Stage := LeStage;
        Session := LaSession;
        //FIN PT15
        If Present = True then
        begin
                Q := OpenSQL('SELECT * FROM STAGEOBJECTIF WHERE POS_NATOBJSTAGE="FOR" AND POS_CODESTAGE="'+Stage+'"',True);
                TobCompetStage := Tob.Create('CompetencesStages',Nil,-1);
                TobCompetStage.LoadDetailDB('CompetencesStages','','',Q,False);
                Ferme(Q);
//              Q := OpenSQL('SELECT * FROM RHCOMPETRESSOURCE WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="'+Matricule+'"',True);
//              TobCompetSalarie := Tob.Create('RHCOMPETRESSOURCE',Nil,-1);
//              TobCompetSalarie.LoadDetailDB('RHCOMPETRESSOURCE','','',Q,False);
//              Ferme(Q);
                For i := 0 to TobCompetStage.Detail.Count - 1 do
                begin
                        Competence := TobCompetStage.Detail[i].Getvalue('POS_COMPETENCE');
                        T := Tob.Create('RHCOMPETRESSOURCE',Nil,-1);
                        Q:=OpenSQL('SELECT MAX(PCH_RANG) FROM RHCOMPETRESSOURCE WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="'+Matricule+'"',TRUE) ;
                        if Not Q.EOF then imax:=Q.Fields[0].AsInteger+1 else iMax:=1 ;
                        Ferme(Q) ;
                        T.PutValue('PCH_TYPERESSOURCE','SAL');
                        T.PutValue('PCH_SALARIE',Matricule);
                        T.PutValue('PCH_COMPETENCE',Competence);
                        T.PutValue('PCH_RANG',IMax);
                        T.PutValue('PCH_CODESTAGE',Stage);
                        T.PutValue('PCH_SESSIONSTAGE',StrToInt(Session));
                        T.PutValue('PCH_DATEVALIDATION',DateValid);
                        T.PutValue('PCH_DEGREMAITRISE',TobCompetStage.Detail[i].GetValue('POS_DEGREMAITRISE'));
                        T.PutValue('PCH_TABLELIBRECR1',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR1'));
                        T.PutValue('PCH_TABLELIBRECR2',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR2'));
                        T.PutValue('PCH_TABLELIBRECR3',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR3'));
                        T.InsertOrUpdateDB(False);
                        T.Free;
                end;
                TobCompetStage.Free;
//                TobCompetSalarie.Free;
        end;
end;*)

procedure TOM_FORMATIONS.InscriptionCursus;
var TobFormations,TobSessions,TF: Tob;
    Q : TQuery;
    j : Integer;
    Salarie : String;
    Nom,Prenom,Etab,LibEmploi : String;
    TN1,TN2,TN3,TN4,LC1,LC2,LC3,LC4,CS : String;
    DatePaie  : TDateTime;
    Nodossier : String;
begin
        Salarie := GetField('PFO_SALARIE');
        If Salarie='' then exit;
        Q := OpenSQL('SELECT MAX (PPU_DATEFIN) MAXDATE FROM PAIEENCOURS WHERE PPU_SALARIE="'+GetField('PFO_SALARIE')+'"',True);
        If Not Q.Eof then DatePaie := Q.FindField('MAXDATE').AsDateTime
        Else DatePaie := V_PGI.DateEntree;
        DatePaie := PlusMois(DatePaie,1);
        DatePaie := FinDeMois(DatePaie);
        Ferme(Q);
        If DatePaie < V_PGI.dateEntree then DatePaie := FinDeMois(V_PGI.DateEntree);
        //DEBUT PT20
(*        NoDossier := '';
        If PGBundleInscFormation then
        begin
          Q := OpenSQL('SELECT PSI_NODOSSIER FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+Salarie+'"',True);
          If Not Q.eof then Nodossier := Q.FindField('PSI_NODOSSIER').AsString;
          Ferme(Q);
        end;*)
        NoDossier := GetNoDossierSalarie(Salarie);
        //FINPT20
        
        Q := OpenSQL('SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_LIBELLE,PSA_PRENOM,'+
        'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4 '+
        'FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True,-1,'',false,NoDossier);
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
        'WHERE PCC_RANGCURSUS='+IntToStr(LeRangCursus)+' AND PCC_CURSUS="'+LeCursus+'" AND PCC_CODESTAGE<>"'+GetField('PFO_CODESTAGE')+'"',True);  //DB2
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
                TF.PutValue('PFO_DATEPAIE',DatePaie);
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
                TF.PutValue('PFO_PREDEFINI','DOS');
                TF.PutValue('PFO_NODOSSIER',NoDossier);
                TF.InsertOrUpdateDB(False);
        end;
        TobSessions.Free;
        TobFormations.Free;
end;

procedure TOM_FORMATIONS.MajCompetencesInitiales;
var Salarie,Stage,Competence : String;
    Q : TQuery;
    IMax,i,NumOrdre : Integer;
    TobCompetStage,T : Tob;
    DateValid : TDateTime;
begin
        Stage := GetField('PFO_CODESTAGE');
        DateValid := GetField('PFO_DATEFIN');
        NumOrdre := getField('PFO_ORDRE');
        Salarie := GetField('PFO_SALARIE');
        Q := OpenSQL('SELECT * FROM STAGEOBJECTIF WHERE POS_NATOBJSTAGE="FOR" AND POS_CODESTAGE="'+Stage+'"',True);
        TobCompetStage := Tob.Create('CompetencesStages',Nil,-1);
        TobCompetStage.LoadDetailDB('CompetencesStages','','',Q,False);
        Ferme(Q);
        For i := 0 to TobCompetStage.Detail.Count - 1 do
        begin
                Competence := TobCompetStage.Detail[i].Getvalue('POS_COMPETENCE');
                T := Tob.Create('RHCOMPETRESSOURCE',Nil,-1);
                Q:=OpenSQL('SELECT MAX(PCH_RANG) FROM RHCOMPETRESSOURCE WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="'+Salarie+'"',TRUE) ;
                if Not Q.EOF then imax:=Q.Fields[0].AsInteger+1 else iMax:=1 ;
                Ferme(Q) ;
                T.PutValue('PCH_TYPERESSOURCE','SAL');
                T.PutValue('PCH_SALARIE',Salarie);
                T.PutValue('PCH_COMPETENCE',Competence);
                T.PutValue('PCH_RANG',IMax);
                T.PutValue('PCH_CODESTAGE',Stage);
                T.PutValue('PCH_SESSIONSTAGE',NumOrdre);
                T.PutValue('PCH_DATEVALIDATION',DateValid);
                T.PutValue('PCH_DEGREMAITRISE',TobCompetStage.Detail[i].GetValue('POS_DEGREMAITRISE'));
                T.PutValue('PCH_TABLELIBRECR1',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR1'));
                T.PutValue('PCH_TABLELIBRECR2',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR2'));
                T.PutValue('PCH_TABLELIBRECR3',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR3'));
                T.InsertOrUpdateDB(False);
                T.Free;
        end;
        TobCompetStage.Free;
end;

procedure TOM_FORMATIONS.VerifReaPrevisionnel(Stage,Salarie,Effectue : String);
var Q:Tquery;
    i : Integer;
    TobInsc : Tob;
    SQL : String;
begin
// DEBUT PT5
   SQL := 'SELECT * FROM INSCFORMATION WHERE PFI_REALISE="-"'+ //PT10
     ' AND PFI_CODESTAGE="'+Stage+'" AND PFI_SALARIE="'+Salarie+'"';
// FIN PT5
   Q := OpenSQL(SQL,True);
     TobInsc := Tob.Create('INSCFORMATION',Nil,-1);
     TobInsc.LoadDetailDB('INSCFORMATION','','',Q,False);
     Ferme(Q);
For i := 0 to TobInsc.Detail.count - 1 do
begin
        TobInsc.Detail[i].PutValue('PFI_REALISE',Effectue);
        TobInsc.Detail[i].UpdateDB;
end;
TobInsc.Free;
end;

//PT22
(*procedure TOM_FORMATIONS.InsererDIFBulletin;
var Q : TQuery;
    NbhT,NbhNonT : Double;
    TypeFor : String;
    TobMvtDIF : Tob;
    i : Integer;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim : String;
    Salarie,Etab : String;
    DateD,DateF : TDateTime;
    DJ,FJ : String;
    HeuresTAlim,HeuresNonTAlim,Jours : Double;
    AbsOk : Integer;
    TypeForm,Predef : String;
    SuppAbsence,SuppRubrique : Boolean;
begin
     If (InitTypePlan <> GetField('PFO_TYPEPLANPREV')) or (GetField('PFO_DATEPAIE') <> InitDatePaie) then SupprimerDIF(True,SuppAbsence,SuppRubrique)
     else SupprimerDIF(False,SuppAbsence,SuppRubrique);
     If GetField('PFO_EFFECTUE') = '-' then Exit;
     NbhT := GetField('PFO_HTPSTRAV');
     NbhNonT := GetField('PFO_HTPSNONTRAV');
     TypeFor := GetField('PFO_TYPEPLANPREV');
     //DEBUT PT15
     Q := OpenSQL('SELECT PSS_JOURSTAGE FROM SESSIONSTAGE WHERE PSS_ORDRE='+LaSession+' AND PSS_CODESTAGE="'+LeStage+'" AND PSS_MILLESIME="'+LeMillesime+'"',True);
     If Not Q.Eof then Jours := Q.FindField('PSS_JOURSTAGE').AsFloat
     Else Jours := IDate1900;
     Ferme(Q);
     //FIN PT15
     Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI## PPF_TYPEPLANPREV="'+TypeFor+'"',True);
     TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
     TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
     Ferme(Q);
     Salarie := GetField('PFO_SALARIE');
     Etab := GetField('PFO_ETABLISSEMENT');
     DateD := GetField('PFO_DATEDEBUT');
     DateF := GetField('PFO_DATEFIN');
     DJ := GetField('PFO_DEBUTDJ');
     FJ := GetField('PFO_FINDJ');
     For i := 0 TO TobMvtDIF.Detail.Count - 1 do
     begin
          Travail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURETRAV');
          HorsTravail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURENONTRAV');
          Predef := TobMvtDIF.Detail[i].GetValue('PPF_PREDEFINI');
          If Predef <> 'DOS' then
          begin
               TypeForm := TobMvtDIF.Detail[i].GetValue('PPF_TYPEPLANPREV');
               If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['DOS',Travail,HorsTravail,TypeForm],False) <> Nil then Continue;
               If Predef = 'CEG' then
               begin
                    If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['STD',Travail,HorsTravail,TypeForm],False) <> Nil then Continue;
               end;
          end;
          HeuresTAlim := 0;
          HeuresNonTAlim := 0;
          If Travail = 'X' then HeuresTAlim := NbhT;
          If HorsTravail = 'X' then HeuresNonTAlim := NbhNonT;
          If (TobMvtDIF.Detail[i].GetValue('PPF_ALIMABS') = 'X') and (SuppAbsence) and (SuppRubrique) then
          begin
               TypeConge := TobMvtDIF.Detail[i].GetValue('PPF_TYPECONGE');
               AbsOk := PGDIFGenereAbs(HeuresTAlim,HeuresNonTAlim,Jours,Salarie,DJ,FJ,TypeConge,Etab,DateD,DateF);
               If AbsOk = -1 then
               begin
                    PGIBox('Impossible de créer l''absence correspondante car le salarié possède déjà une absence sur cette période',Ecran.Caption);
                    Exit;
               end;

          end;
          If (TobMvtDIF.Detail[i].GetValue('PPF_ALIMRUB') = 'X')  and (SuppAbsence) and (SuppRubrique) then
          begin
               Rubrique := TobMvtDIF.Detail[i].GetValue('PPF_RUBRIQUE');
               Travail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURETRAV');
               HorsTravail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURENONTRAV');
               TypeAlim := TobMvtDIF.Detail[i].GetValue('PPF_ALIMENT');
               RecupereLesFormationsDIF(Salarie,TypeFor,Rubrique,TypeAlim,GetField('PFO_DATEPAIE'),Travail = 'X',HorsTravail = 'X');
               //PGDIFGenereRub(HeuresTAlim,HeuresNonTAlim,Salarie,Etab,Rubrique,TypeAlim,GetField('PFO_CODESTAGE'),GetField('PFO_CONFIDENTIEL'),DateD,DateF,GetField('PFO_DATEPAIE'));
          end;
     end;
     TobMvtDIF.free;
end;

procedure TOM_FORMATIONS.SupprimerDIF(Supp : Boolean; var SuppAbsence,SuppRubrique : Boolean);
var Q : TQuery;
    TypeFor : String;
    TobMvtDIF : Tob;
    i : Integer;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim,Predef : String;
    DateDebRub,DateFinRub : TDateTime;
begin
     SuppRubrique := True;
     SuppAbsence := True;
     TypeFor := InitTypePlan;
     Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI## PPF_TYPEPLANPREV="'+TypeFor+'"',True);
     TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
     TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
     Ferme(Q);
     For i := 0 TO TobMvtDIF.Detail.Count - 1 do
     begin
          Travail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURETRAV');
          HorsTravail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURENONTRAV');
          Predef := TobMvtDIF.Detail[i].GetValue('PPF_PREDEFINI');
          If Predef <> 'DOS' then
          begin
               If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['DOS',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
               If Predef = 'CEG' then
               begin
                    If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['STD',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
               end;
          end;
          If TobMvtDIF.Detail[i].GetValue('PPF_ALIMABS') = 'X' then
          begin
               TypeConge := TobMvtDIF.Detail[i].GetValue('PPF_TYPECONGE');
               If ExisteSQL('SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS FROM ABSENCESALARIE WHERE '+
               'PCN_TYPECONGE="'+TypeConge+'" AND '+
               'PCN_SALARIE="'+GetField('PFO_SALARIE')+'" AND '+
               '(PCN_CODETAPE="P" OR PCN_CODETAPE="S") AND '+
               'PCN_DATEDEBUTABS="'+UsDateTime(GetField('PFO_DATEDEBUT'))+'" AND '+
               'PCN_DATEFINABS="'+UsDateTime(GetField('PFO_DATEFIN'))+'"') then
               begin
                    SuppAbsence := True;
                    PGIBox('L''absence pour le salarié '+RechDom('PGSALARIE',GetField('PFO_SALARIE'),False)+
                    ' du '+DateToStr(GetField('PFO_DATEDEBUT'))+' au '+DateToStr(GetField('PFO_DATEFIN'))+
                    '#13#10 a déjà été intégrée dans le bulletin. Aucune modification ne sera effectuée','Génération absence');
               end
               else PGDIFSuppAbs(GetField('PFO_SALARIE'),TypeConge,GetField('PFO_DATEDEBUT'),GetField('PFO_DATEFIN'));
          end;
          If TobMvtDIF.Detail[i].GetValue('PPF_ALIMRUB') = 'X' then
          begin
               Rubrique := TobMvtDIF.Detail[i].GetValue('PPF_RUBRIQUE');
               TypeAlim := TobMvtDIF.Detail[i].GetValue('PPF_ALIMENT');
               DateDebRub := DebutDeMois(InitDatePaie);
               DateFinRub := FinDeMois(InitDatePaie);
               If ExisteSQL('SELECT PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE '+
               'PPU_DATEDEBUT>="'+UsDateTime(dateDebRub)+'" AND '+
               'PPU_DATEFIN<="'+UsDateTime(dateFinRub)+'" AND '+
               'PPU_SALARIE="'+GetField('PFO_SALARIE')+'"') then
               begin
                    SuppRubrique := False;
                    If InitDatePaie <> GetField('PFO_DATEPAIE') then
                    begin
                         PGIBox('Le bulletin pour le salarié '+RechDom('PGSALARIE',GetField('PFO_SALARIE'),False)+
                         ' du '+DateToStr(dateDebRub)+' au '+DateToStr(dateFinRub)+
                         '#13#10 a déjà été créé. La date de validité ne peut être modifiée et aucune modification ne sera effectuée pour les rubriques','Génération absence');
                         SetField('PFO_DATEPAIE',InitDatePaie);
                    end
                    else
                    begin
                         PGIBox('Le bulletin pour le salarié '+RechDom('PGSALARIE',GetField('PFO_SALARIE'),False)+
                         ' du '+DateToStr(dateDebRub)+' au '+DateToStr(dateFinRub)+
                         '#13#10 a déjà été créé. Aucune modification ne sera effectuée pour les rubriques','Génération absence');
                    end;
               end
               else
               begin
                    PGDIFSuppRub(GetField('PFO_SALARIE'),Rubrique,TypeAlim,InitDatePaie);
                    If Supp then RecupereLesFormationsDIF(GetField('PFO_SALARIE'),TypeFor,Rubrique,TypeAlim,InitDatePaie,Travail = 'X',HorsTravail = 'X',GetField('PFO_CODESTAGE'),GetField('PFO_ORDRE'));
               end;
          end;
     end;
     TobMvtDIF.free;
end;*)

procedure TOM_FORMATIONS.VoirMouvementsGenere(Sender : TObject);
var St : String;
begin
    //DEBUT PT15
     St := LeStage + ';' +
     LaSession + ';' +
     LeMillesime;
     //FIN PT15
     AGLLanceFiche('PAY','MVTSFORMPAIE','','',St);
end;

procedure TOM_FORMATIONS.VoirCompteursDIF(Sender : TObject);
begin
     AGLLanceFiche('PAY','COMPTEURSDIF','','',GetField('PFO_SALARIE'));
end;

procedure TOM_FORMATIONS.MettreAjourLeSalaire;
var TauxHoraire,Montant : Double;
    Q : TQuery;
begin
     If GetField('PFO_SALARIE')='' then Exit;
     TauxHoraire := 0;
     if VH_Paie.PGForValoSalaire='VCR' then TauxHoraire := ForTauxHoraireReel(GetField('PFO_SALARIE'),0,0,'',ValReel,GetField('PFO_DATEFIN')); //PT19
     if VH_Paie.PGForValoSalaire='VCC' then TauxHoraire := ForTauxHoraireCategoriel(GetField('PFO_LIBEMPLOIFOR'),MillesimeEC);
     Montant := TauxHoraire*GetField('PFO_HTPSTRAV');
     //PT3 suppression contrôle taux cadre <> 1
     
     //PT21
     If PGBundleHierarchie Then
    	Q := OpenSQL('SELECT PSA_DADSCAT FROM '+GetBase(GetBaseSalarie(GetField('PFO_SALARIE')),'SALARIES')+' WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True)
     Else
    	Q := OpenSQL('SELECT PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+GetField('PFO_SALARIE')+'"',True);  //PT13
     If (Not Q.EOF) And ((Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02')) then 
		Montant := Arrondi(Montant * TauxChargeC,2)
     else 
     	Montant := Arrondi(Montant * TauxChargeNC,2);
     Ferme(Q);
     SetField('PFO_COUTREELSAL',Montant);
end;

Initialization
  registerclasses ( [ TOM_FORMATIONS ] ) ;
end.

