{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : TOM Gestion des sessions des stages
Mots clefs ... : PAIE
*****************************************************************
PT1  | 15/05/2003 | V_42  | JL | Gestion investissement
PT2  | 27/07/2004 | V_50  | JL | FQ 11411 Correction requête pour oracle
PT3  | 01/09/2004 | V_50  | JL | FQ 11535 Saisie sur plusieurs exercices
PT4  | 15/02/2005 | V_60  | JL | Maj table formation lors de laz validation
PT5  | 20/07/2005 | V_60  | JL | FQ 12447 Autorisation de saisie seulement si millésime non cloturé
PT6  | 20/07/2005 | V_60  |    | Maj colonne investissement si choix = <<AUCUN>>
PT7  | 20/07/2005 | V_650 | JL | FQ 11403 Ajout édition
PT8  | 16/03/2006 | V_650 | FL | FQ 12978 Ajout édition feuille présence
---  | 20/03/2006 |       | JL | modification clé annuaire ----
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT9  | 10/04/2007 | V_702 | FL | FQ 13695 Ajout du bouton Dupliquer
PT10 | 11/04/2007 | V_702 | FL | FQ 13996 Pas d'affichage d'alerte sur les dates en cas de sessions pluri-annuelles
PT11 | 14/06/2007 | V_720 | JL | Gestion partage formation
PT12 | 10/08/2007 | V_80  | JL | Ajout fichier joint au mail de convocation
PT13 | 14/11/2007 | V_80  | FL | CWAS : Chargement de la grille d'investissement sur affichage de l'onglet coût
     |            |       |    |        afin d'éviter le message "Impossible de focaliser une fenêtre inactive".
PT14 | 28/09/2007 | V_7   | FL | Emanager / Report / Adaptation cursus + accès assistant
PT15 | 02/04/2008 | V_80  | FL | Correctif pour les heures
PT16 | 02/04/2008 | V_80  | FL | Restriction sur les prédéfinis
PT17 | 02/04/2008 | V_80  | FL | Adaptation aux groupes de travail
PT18 | 02/04/2008 | V_80  | FL | Amélioration des mails
PT19 | 02/04/2008 | V_80  | FL | Génération de l'ID de session à chaque fois
PT20 | 07/04/2008 | V_80  | FL | Mise à zéro des heures, ajout de la balise html pour les mails formatés 
     |            |       |    | (pour envoi smtp) et copie du bloc note dans une variable pour éviter
     |            |       |    | la mise à jour du champ de la fiche lors de la création du mail
PT21 | 15/04/2008 | V_804 | FL | Suppression du fichier temporaire lors d'une convocation formation
PT22 | 15/04/2008 | V_804 | FL | Récupération du mail d'abord depuis la table UTILISAT, puis ensuite depuis DEPORTSAL
PT23 | 23/04/2008 | V_804 | FL | Restriction des stages sélectionnables en fonction du prédéfini/no dossier
PT24 | 30/04/2008 | V_804 | FL | Récupérer les adresses email à partir des salariés inscrits et non des salariés affectés (deportsal)
PT25 | 23/05/2008 | V_804 | FL | FQ 14052 Adaptation pour la fiche de session simplifiée
PT26 | 10/06/2008 | V_804 | FL | Correction du lancement de la feuille de présence
PT27 | 17/06/2008 | V_804 | FL | FQ 14055 Gestion des sous-sessions
PT29 | 11/09/2008 |       | VG | En CWAS, Erreur en création d'une session de
                                 formation
 }
unit UTOMSESSIONSTAGE;

interface
uses  Windows,Controls,Classes,sysutils,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Fe_Main,Fiche,EdtREtat,
{$ELSE}
       MaineAgl,eFiche,UtileAGL,
{$ENDIF}
      HCtrls,ComCtrls,HEnt1,HMsgBox,UTOM,UTOB,HTB97,LookUp,Menus,MailOL,PgOutilsFormation,EntPaie,ParamSoc,HRichOLE,UYFileSTD ;

Type
     TOM_SESSIONSTAGE = Class(TOM)
       procedure OnArgument (stArgument : String )  ; override ;
       procedure OnNewRecord  						; override ;
       procedure OnUpdateRecord 					; override ;
       procedure OnLoadRecord 						; override ;
       procedure OnChangeField(F: TField)  			; override ;
       procedure OnDeleteRecord						; override ;
       procedure OnAfterUpdateRecord				; override ;
       procedure OnClose            				; override ; //PT25

       private
         LeStage,Millesime,TypeSaisie,InitStage,MillesimeConsult,LaNature : String;
         LaDuree : Double;
         MillesimeEC : String;
         InitAvecClient : boolean;
         GInvest : THGrid;
         InitDDSession,InitDFSession : TDateTime;
         DuplSession, DuplOrdre, DuplMillesime : String; //PT9
         ADupliquer : boolean; //PT9
         GrilleChargee : Boolean;  //PT13
         GrilleSal   : THGrid;     //PT25
		 NewLine : Boolean;        //PT25
         TobInscrits : Tob;        //PT25
         NumSessionMaitre : String;//PT27
         TobModifiedFields :TOB;   //PT27
         MAJSousSessions : Boolean;//PT27
         
       procedure RendDureeStage;
       procedure ActiveFiche(LaFiche : String);
       procedure AnimatSession (Sender : TObject);
       procedure StageElipsisClick(Sender : TObject);
       procedure BFormationClick(Sender : TObject);
       procedure LanceMailSalarie(Sender : TObject);
       procedure LanceMailResponsables(Sender : TObject);
       procedure LanceMailAvertirSession(Sender : TObject);
       procedure ConvocSession(Sender : TObject);
       procedure CreerNouveauStage(Sender : TObject);
       procedure SaisirFeuillePresence(Sender : TObject);
       procedure OnGICellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
       procedure MajInvest;
       procedure GestionDesDif;
       procedure SupprimerDIF(Salarie,TypeFor : String;DateD,DateF,DatePaie : TDateTime; var SuppAbsence,SuppRubrique : Boolean);
       procedure EditionSession(Sender : TObject); //PT7
       procedure EditionPresenceSession(Sender : TObject); //PT8
       Procedure BSousSessionClick (Sender : TObject);
       procedure DupliquerSession (Sender: Tobject); //PT9
       procedure RecupInfosSession; //PT9
       procedure OnEnterOngletCouts (Sender : TObject); //PT13
       procedure OnChangeHeure (Sender : TObject); //PT15
       //PT25 - Début
       Procedure AccesBoutonsSousSession;
       procedure AccesBoutons (State : Boolean);
       procedure OnEnterOngletInscriptions(Sender: TObject);
       procedure OnExitOngletInscriptions(Sender: TObject);
       
       Function  InscriptionFormation(Salarie,TypePlan,Typologie : String; FormationEffectuee:boolean=True;NumSession:String='';NoSessionMulti:String='') : Boolean;
       procedure SaisieEnMasse (Sender : TObject);
       procedure AfficheStagiaires;
       
       procedure BRecupPrevClick      (Sender : TObject);
       Procedure BPlusInfosClick      (Sender : TObject);
       Procedure BAjoutStagiaireClick (Sender : Tobject);
       Procedure BSupprStagiaireClick (Sender : Tobject);
       Procedure BValidStagiaireClick (Sender : TObject);
       
       // Gestion de la grille des inscrits
       Procedure OnKeyDown    (Sender : TObject; var Key : Word; Shift : TShiftState);
       procedure OnGSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
       procedure OnGSClick    (Sender : TObject);
       procedure OnGSRowExit  (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
       procedure OnGSRowEnter (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
       procedure OnGSElipsisClick(Sender : TObject);
       //PT25 - Fin
     End ;

implementation

Uses PgOutils, UTOFPGMul_InscFor, PGOutilsGrids, UtilPGI;
//PT25
Const COL_SALARIE  = 0;
      COL_TYPEPLAN = 3;
      COL_NBHEURES = 4;
      COL_PRESENCE = 5;

{ TOM_SESSIONSTAGE }

procedure TOM_SESSIONSTAGE.OnAfterUpdateRecord;
var i               : Integer;
    TobSousSessions : TOB;
    St              : String;
begin
  inherited;
  
  	//PT27 Mise à jour des sous-sessions avec les modifications apportées à l'en-tête
  	If MAJSousSessions Then
  	Begin
  		// Chargement des sous-sessions
  		TobSousSessions := TOB.Create('LesSousSessions', Nil, -1);
  		TobSousSessions.LoadDetailDBFromSQL('SESSIONSTAGE','SELECT * FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'"');
  		
  		// Modification des valeurs dans les TOBs filles à part pour les champs qui sont propres à la session en-tête
  		For i:= 1 To TobModifiedFields.NbChamps Do
        Begin
        	If TobModifiedFields.IsFieldModified(TobModifiedFields.GetNomChamp(i))
        	   And (TobModifiedFields.GetNomChamp(i) <> 'PSS_LIBELLE')
        	   And (TobModifiedFields.GetNomChamp(i) <> 'PSS_DUREESTAGE')
        	   And (TobModifiedFields.GetNomChamp(i) <> 'PSS_JOURSTAGE')
        	   And (TobModifiedFields.GetNomChamp(i) <> 'PSS_DATEDEBUT')
        	   And (TobModifiedFields.GetNomChamp(i) <> 'PSS_DATEFIN')
        	   And (TobModifiedFields.GetNomChamp(i) <> 'PSS_DATEMODIF') Then
        	Begin
        		TobSousSessions.PutValueAllFille(TobModifiedFields.GetNomChamp(i),TobModifiedFields.GetValue(TobModifiedFields.GetNomChamp(i)));
        	End;
        End;
        Try
            TobSousSessions.UpdateDB();
        Except
            PGIError(TraduireMemoire('Une erreur s''est produite durant la mise à jour des sous-sessions'));
        End;
  	End;
  
    If TypeSaisie <> 'CURSUS' then
    begin
        If GetField('PSS_PGTYPESESSION') <> 'ENT' then
        begin
            SetControlEnabled('BANIMAT',True);
            SetControlEnabled('BFORMATION',True);
            SetControlEnabled('BCONVOC',True);
            SetControlEnabled('BPRESENCE',True);
        end;
        MajInvest;
        If isFieldModified( 'PSS_COUTPEDAG' ) then CalcCtInvestSession('PEDAGOGIQUE',GetField('PSS_CODESTAGE'),GetField('PSS_MILLESIME'),GetField('PSS_ORDRE'));
    end;
    
    If (IsFieldModified ('PSS_HEUREDEBUT')) or (IsFieldModified('PSS_HEUREFIN')) then
    begin
    	If GetField('PSS_PGTYPESESSION') <> 'ENT' Then //PT27
        	ExecuteSQL('UPDATE FORMATIONS SET PFO_HEUREDEBUT="'+UsTime(GetField('PSS_HEUREDEBUT'))+'",PFO_HEUREFIN="'+UsTime(GetField('PSS_HEUREFIN'))+'"'+
            	       'WHERE PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'"')
        Else //PT27
        Begin
        	If MAJSousSessions Then
        	Begin
        		Try
	        		BEGINTRANS;
	        		For i:=0 To TobSousSessions.Detail.Count-1 Do
	        		Begin
	        			ExecuteSQL('UPDATE FORMATIONS SET PFO_HEUREDEBUT="'+UsTime(TobSousSessions.Detail[i].GetValue('PSS_HEUREDEBUT'))+'",PFO_HEUREFIN="'+UsTime(TobSousSessions.Detail[i].GetValue('PSS_HEUREFIN'))+'"'+
	            			       'WHERE PFO_CODESTAGE="'+TobSousSessions.Detail[i].GetValue('PSS_CODESTAGE')+'" AND PFO_ORDRE='+IntToStr(TobSousSessions.Detail[i].GetValue('PSS_ORDRE'))+' AND PFO_MILLESIME="'+TobSousSessions.Detail[i].GetValue('PSS_MILLESIME')+'"');
	            	End;
	            	COMMITTRANS;
	            Except
	            	ROLLBACK;
	            End;
        	End;
        End;
    end;
    
    If (IsFieldModified ('PSS_DATEDEBUT')) or (IsFieldModified('PSS_DATEFIN')) And (GetField('PSS_PGTYPESESSION') <> 'ENT') Then //PT27
    begin
        ExecuteSQL('UPDATE FORMATIONS SET PFO_DATEDEBUT="'+UsDateTime(GetField('PSS_DATEDEBUT'))+'",PFO_DATEFIN="'+UsDateTime(GetField('PSS_DATEFIN'))+'"'+
                   'WHERE PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'"');  //DB2

        If ExisteSQL('SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' '+  //DB2
                     'AND PAN_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PAN_MILLESIME="'+GetField('PSS_MILLESIME')+'"') then
        begin
            PGIBox('Attention il existe des animateurs, vous devrez modifier leurs dates',Ecran.Caption);
        end;
        GestionDesDIF;
    end;
    
    TFFiche(Ecran).Retour:='MODIF';
    
    // DEBUT PT4
    St := 'UPDATE FORMATIONS SET PFO_FORMATION1="'+GetField('PSS_FORMATION1')+'",PFO_FORMATION2="'+GetField('PSS_FORMATION2')+'"'+
	               ',PFO_FORMATION3="'+GetField('PSS_FORMATION3')+'",PFO_FORMATION4="'+GetField('PSS_FORMATION4')+'",PFO_FORMATION5="'+GetField('PSS_FORMATION5')+'"'+
	               ',PFO_FORMATION6="'+GetField('PSS_FORMATION6')+'",PFO_FORMATION7="'+GetField('PSS_FORMATION7')+'",PFO_FORMATION8="'+GetField('PSS_FORMATION8')+'"'+
	               ',PFO_INCLUSDECL="'+GetField('PSS_INCLUSDECL')+'",PFO_NATUREFORM="'+GetField('PSS_NATUREFORM')+'"'+
	               ',PFO_CENTREFORMGU="'+GetField('PSS_CENTREFORMGU')+'"'+
	               ' WHERE PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'"';

    If GetField('PSS_PGTYPESESSION') <> 'ENT' Then //PT27
	    ExecuteSQL(St+' AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE')))
	Else If MAJSousSessions Then
		ExecuteSQL(St+' AND PFO_ORDRE IN (SELECT PSS_ORDRE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'")'); //PT27
    //FIN PT4

    If Assigned(TobSousSessions) Then FreeAndNil(TobSousSessions); //PT27
end;

procedure TOM_SESSIONSTAGE.OnDeleteRecord;
var Erreur:Boolean;
begin
  inherited;	
    Erreur := ExisteSQL('SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PAN_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PAN_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+'');  //DB2
    If Erreur=True Then
    begin
        LastError := 1;
        PGIBox(TraduireMemoire('Vous ne pouvez pas supprimer cette session de formation car des animateurs lui sont affectées'),TFFiche(Ecran).Caption);
        Exit;
    end;
    Erreur := ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+'');  //DB2
    If Erreur=True Then
    begin
        LastError := 1;
        PGIBox(TraduireMemoire('Vous ne pouvez pas supprimer cette session de formation car des inscriptions lui sont affectées'),TFFiche(Ecran).Caption);
        Exit;
    end;
    
  	//PT27
  	If GetField('PSS_PGTYPESESSION') = 'ENT' Then
  	Begin
  		Erreur := ExisteSQL('SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PAN_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PAN_ORDRE IN (SELECT PSS_ORDRE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'")');
  		If Not Erreur then	
  			Erreur := ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFO_ORDRE IN (SELECT PSS_ORDRE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'")');
  		If Erreur Then
  		Begin
	        LastError := 1;
	        PGIBox(TraduireMemoire('Vous ne pouvez pas supprimer cette session et les sous-sessions dépendantes car des animateurs et/ou inscriptions leur sont affectés'),TFFiche(Ecran).Caption);
	        Exit;
  		End
  		Else
  		Begin
            If ExisteSQL('SELECT 1 FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'"') Then
            Begin
	  			If PGIAsk(TraduireMemoire('Attention : toutes les sous-sessions seront également supprimées.#13#10Êtes-vous sûr de vouloir continuer?'),TFFiche(Ecran).Caption) = mrYes Then
	  			Begin
	  				Try
		  				BEGINTRANS;
		  				ExecuteSQL('DELETE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'"');
		  				COMMITTRANS;
	  				Except
	  					ROLLBACK;
	  				End;
	  			End
	            Else
	            Begin
	                LastError := 1;
	                Exit;
	            End;
            End;
  		End;
  	End;
    TFFiche(Ecran).Retour:='SUPPR';
end;

procedure TOM_SESSIONSTAGE.AnimatSession(Sender : TObject);
begin
    ActiveFiche('MUL_SESSIONANIMAT');
    RefreshDB;
end;

procedure TOM_SESSIONSTAGE.ActiveFiche(LaFiche : String);
var St    : String;
    I     : Integer;
    DD,DF : TDateTime;
begin
	//PT25
	// Plus besoin de ce test étant donné que les boutons sont cachés
    //if (GetField('PSS_CODESTAGE') <> '') AND (GetField ('PSS_ORDRE') <> 0) then
	//Begin
	I  := GetField ('PSS_ORDRE');
	DD := GetField('PSS_DATEDEBUT');
	DF := GetField('PSS_DATEFIN');
	st := 'SESSION;'+GetField('PSS_CODESTAGE')+';'+IntToStr (I)+';'
	+DateToStr (DD)+';'+DateToStr (DF)+';'+GetField('PSS_MILLESIME');
	AglLanceFiche ('PAY',LaFiche, '', '' , st);
    //end
    //else PgiBox (TraduireMemoire('Vous devez d''abord créer la session de stage'), Ecran.caption );
end;

procedure TOM_SESSIONSTAGE.OnArgument(stArgument : String);
var St     : String;
    BT,BFormation,BConvoc,BStage,BPresence,BImprimer,BDupliquer : TToolbarButton97;
    {$IFNDEF EAGLCLIENT}
    THStage : THDBEdit;
    {$ELSE}
    THStage : THEdit;
    {$ENDIF}
    MenuMail : TPopUpMenu;
    Num : Integer;
    DD,DF : TDateTime;
    Heures : THEdit; //PT15
    Compon : TComponent;
    //Action : String;
begin
  inherited;
    TypeSaisie := Trim(ReadTokenPipe(StArgument,';'));
    
    If (TypeSaisie = 'RECUPPREVPLA') or (TypeSaisie='RECUPPREVREA') then
    begin
        TypeSaisie := 'SAISIE';
    end
    Else If TypeSaisie = 'CARRIERE' then LaNature := Trim (ReadTokenPipe(StArgument,';'))
    else
    begin
        st := Trim (ReadTokenPipe(StArgument,';'));
        Millesime := Trim (ReadTokenPipe(StArgument,';'));
        MillesimeConsult := Trim (ReadTokenPipe(StArgument,';'));
        LeStage := Trim (ReadTokenSt(St));// recup du code du stage
    end;
    
    if LeStage <> '' then SetControlEnabled ('PSS_CODESTAGE',FALSE); //PT27
    SetControlEnabled ('PSS_ORDRE',FALSE);
    
    If TypeSaisie <> 'CURSUS' then
    begin
        BT := TToolbarButton97 (GetControl ('BANIMAT'));
        if BT <> NIL then BT.OnClick := AnimatSession;
        {$IFNDEF EAGLCLIENT}
        THStage := THDBEdit(GetControl('PSS_CODESTAGE'));
        {$ELSE}
        THStage := THEdit(GetControl('PSS_CODESTAGE'));
        {$ENDIF}
        If THStage <> Nil Then THStage.OnElipsisClick := StageElipsisClick;
        BFormation := TToolBarButton97(GetControl('BFORMATION'));
        If BFormation <> Nil Then BFormation.onclick := BFormationClick;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
            if Num >8 then Break;
            VisibiliteChampFormation (IntToStr(Num),GetControl ('PSS_FORMATION'+IntToStr(Num)),GetControl ('TPSS_FORMATION'+IntToStr(Num)));
        end;
        Compon := GetControl('TPMAIL') As TComponent;
        MenuMail := Compon As TPopupMenu;
        MenuMail.Items[0].OnClick := LanceMailSalarie;
        MenuMail.Items[1].OnClick := LanceMailResponsables;
        MenuMail.Items[2].OnClick := LanceMailAvertirSession;
        If VH_Paie.PgResponsables = False then
        begin
            MenuMail.Items[1].Visible := False;
            MenuMail.Items[2].Visible := False;
        end;
        BConvoc := TToolBarButton97(GetControl('BCONVOC'));
        If BConvoc <> Nil Then BConvoc.OnClick := ConvocSession;
        BStage := TtoolBarButton97(GetControl('BTNFORMATION'));
        If BStage <> Nil Then BStage.OnClick := CreerNouveauStage;
        BPresence := TToolBarButton97(GetControl('BPRESENCE'));
        If BPresence <> Nil Then BPresence.OnClick := SaisirFeuillePresence;
        SetControlEnabled('PSS_LIEUFORM',True);
        //DEBUT PT1
        GInvest := THGrid(GetControl('GINVESTSESSION'));
        //PT13 - Début
        If (TFFiche(Ecran).Name <> 'SESSIONSIMPLE') And (TypeSaisie <> 'SOUSSESSION') Then TTabSheet(GetControl('PCOUTS',true)).OnShow := onEnterOngletCouts; //PT25
        GrilleChargee := False;
        //If GInvest <> nil then
        //begin
        //        GInvest.ColFormats[0] := 'CB=PGTYPEPLANPREV';
        //        GInvest.ColFormats[1] := 'CB=PGOUINON';
        //        GInvest.ColFormats[2] := '# ##0.00';
        //        GInvest.ColFormats[3] := 'CB=PGOUINON';
        //        GInvest.ColFormats[4] := '# ##0.00';
        //        GInvest.ColFormats[5] := 'CB=PGOUINON';
        //        GInvest.ColFormats[6] := '# ##0.00';
        //        GInvest.ColFormats[7] := '# ##0';
        //        GInvest.OnCellExit := OnCellExit;
        //end;
        //PT13 - Fin
        //FIN PT1
        SetControLCaption('TPSS_COUTSALAIRE','Coût(s) animateur(s)');
    end;
    If TypeSaisie = 'CONSULTATION' Then
    begin
        SetControlVisible('BFORMATION',False);
        TFFiche(Ecran).FTypeAction := TaConsult;
    end
    //PT27
    Else If TypeSaisie = 'SOUSSESSION' Then
    begin
    	ReadTokenSt(StArgument);
    	NumSessionMaitre := ReadTokenPipe(StArgument,';');
    end
    Else
    begin
        If TypeSaisie = 'SAISIEWEB' Then TFFiche(Ecran).FTypeAction := TaConsult
        Else
        begin
            If TFFiche(Ecran).FTypeAction <> TaCreat Then
            begin
                If Not (ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+MillesimeConsult+'" AND PFE_CLOTURE="-"')) Then //PT5
                TFFiche(Ecran).FTypeAction := TaConsult
                Else TFFiche(Ecran).FTypeAction := TaModif;
            end;
        end;
    end;
    
    MillesimeEC := RendMillesimeRealise(DD,DF);;
    BImprimer := TToolBarButton97(GetControl('BImprimer'));
    If BImprimer <> Nil then BImprimer.OnClick := EditionSession;
    BImprimer := TToolBarButton97(GetControl('BIMPPRESENCE'));
    If BImprimer <> Nil then BImprimer.OnClick := EditionPresenceSession;
    //PT9 - Debut
    BDupliquer := TToolBarButton97(GetControl('BDUPLIQUER'));
    If BDupliquer <> nil then BDupliquer.OnClick := DupliquerSession;
    Bt := TToolBarButton97(GetControl('BSOUSSESSION'));
    If Bt <> Nil then Bt.OnClick := BSousSessionClick;
    DuplSession   := '';
    DuplOrdre     := '';
    DuplMillesime := '';
    ADupliquer    := false;
    //PT9 - Fin
    //DEBUT PT11
    If PGBundleCatalogue then
    begin
        If (not PGDroitMultiForm) Or (StArgument <> 'ACTION=CREATION') then
        begin
            SetControlEnabled('PSS_PREDEFINI',False);
            SetControlEnabled('PSS_NODOSSIER',False);
        end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('PSS_NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT17
    end
    else
    begin
        SetControlVisible('PSS_PREDEFINI',False);
        SetControlVisible('TPSS_PREDEFINI',False);
        SetControlVisible('PSS_NODOSSIER',False);
        SetControlVisible('TPSS_NODOSSIER',False);
    end;
    //FIN PT11

    GereAccesPredefini (THValComboBox(GetControl('PSS_PREDEFINI'))); //PT16

	//PT15 - Début
	Heures := THEdit(GetControl('HEUREDEBUT'));
	If Heures <> Nil Then Heures.OnExit := OnChangeHeure;
	Heures := THEdit(GetControl('HEUREFIN'));
	If Heures <> Nil Then Heures.OnExit := OnChangeHeure;
	//PT15 - Fin

    //PT25
    If TFFiche(Ecran).Name = 'SESSIONSIMPLE' Then
    Begin
        SetControlVisible('BPRESENCE', False);
        SetControlVisible('BFORMATION', False);
        //SetControlVisible('BSOUSSESSION', False); //PT27
        TTabSheet(GetControl('PINSCRIPTIONS',True)).OnShow := onEnterOngletInscriptions;
        TTabSheet(GetControl('PINSCRIPTIONS',True)).OnHide := onExitOngletInscriptions;

	    GrilleSal := THGrid(GetControl('GRILLESAL'));
	    GrilleSal.OnCellEnter              := OnGSCellEnter;
	    GrilleSal.OnClick                  := OnGSClick;
	    GrilleSal.OnRowExit                := OnGSRowExit;
	    GrilleSal.OnRowEnter               := OnGSRowEnter;
	    GrilleSal.OnElipsisClick           := OnGSElipsisClick;
	    GrilleSal.ColFormats[COL_TYPEPLAN] := 'CB=PGTYPEPLANPREV';
	    GrilleSal.ColFormats[COL_NBHEURES] := '##0';
	    GrilleSal.ColTypes  [COL_NBHEURES] := 'R';
	    GrilleSal.ColAligns [COL_NBHEURES] := taRightJustify;
	    GrilleSal.ColTypes  [COL_PRESENCE] := 'B';
	    GrilleSal.ColAligns [COL_PRESENCE] := taCenter;
        GrilleSal.ColWidths [COL_PRESENCE] := -1;
	    GrilleSal.ColFormats[COL_PRESENCE] := IntToStr(Ord(csCheckBox));
        GrilleSal.ColEditables[1] := False; GrilleSal.ColEditables[2] := False;
        // Attention : col non éditable pour la checkbox
        GrilleSal.ColEditables[COL_PRESENCE] := False;
        
        TFFiche(Ecran).OnKeyDown := OnKeyDown;
        
        Bt := TToolBarButton97(GetControl('BRECUPERATION'));
        If Bt <> Nil then Bt.OnClick := BRecupPrevClick;
      	Bt := TToolbarButton97(GetControl('BMULTISAL'));
       	If Bt <> Nil then Bt.onClick := SaisieEnMasse;
      	Bt := TToolbarButton97(GetControl('BPLUSDINFOS'));
       	If Bt <> Nil then Bt.onClick := BPlusInfosClick;
        Bt := TToolBarButton97(GetControl('BAJOUTSTAG'));
        If Bt <> Nil then Bt.OnClick := BAjoutStagiaireClick;
        Bt := TToolBarButton97(GetControl('BSUPPRSTAG'));
        If Bt <> Nil then Bt.OnClick := BSupprStagiaireClick;
        Bt := TToolBarButton97(GetControl('BVALIDERINSC'));
        If Bt <> Nil then Bt.OnClick := BValidStagiaireClick;
    End;

    //PT27
    If TypeSaisie = 'SOUSSESSION' Then
    Begin
      	Bt := TToolbarButton97(GetControl('BPLUSDINFOS'));
       	If Bt <> Nil then Bt.onClick := BPlusInfosClick;
    End
    Else If (TypeSaisie = 'SAISIE') Then
    Begin
        Bt := TToolBarButton97(GetControl('BRECUPERATIONSS'));
        If Bt <> Nil then Bt.OnClick := BRecupPrevClick;
      	Bt := TToolbarButton97(GetControl('BMULTISALSS'));
       	If Bt <> Nil then Bt.onClick := SaisieEnMasse;
    End;

    //PT27
    TobModifiedFields := TOB.Create('SESSIONSTAGE', Nil, -1);
end;

procedure TOM_SESSIONSTAGE.StageElipsisClick(Sender : TObject);
var  St,StWhere,StOrder : String;
     ceg : Boolean;
begin
    If Sender = Nil then Exit;
    Ceg := GetParamSocSecur('SO_IFDEFCEGID',False);
    St := 'SELECT PST_CODESTAGE,PST_LIBELLE,PST_LIBELLE1,PLF_LIBELLE FROM STAGE'+
          ' LEFT JOIN LIEUFORMATION ON PST_LIEUFORM=PLF_LIEUFORM ';
    If TypeSaisie= 'CARRIERE' then
    begin
        If LaNature = 'INITIALE' then StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="0000" AND PST_NATUREFORM="004"';
        If LaNature = 'EXTERNE' then StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="0000" AND PST_NATUREFORM="002"';
        If LaNature = 'INTERNE' then StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="0000" AND PST_NATUREFORM="001"';
    end
    else StWhere := 'PST_ACTIF="X" AND'+  //PT2
                    ' PST_MILLESIME="0000" AND PST_NATUREFORM<>"004"';
    
    If PGBundleCatalogue then
    Begin
        //PT23
        {If Not PGDroitMultiForm Then
            StWhere := StWhere + DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)
        Else
            StWhere := StWhere + DossiersAInterroger('','','PST',True,True);}
        StWhere := StWhere + DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('PSS_NODOSSIER'),'PST',True,True);
    End;
    
    StOrder := 'PST_LIBELLE,PST_MILLESIME';
    If ceg then
    begin
        {$IFNDEF EAGLCLIENT}
        LookupList(THDBEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,PST_FORMATION3',StWhere,StOrder, True,-1);
        {$ELSE}
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_LIBELLE1,PST_FORMATION3',StWhere,StOrder, True,-1);
        {$ENDIF}
    end
    else
    begin
        If PGBundleCatalogue then
        begin
            {$IFNDEF EAGLCLIENT}
            LookupList(THDBEdit(Sender),'Liste des stages','STAGE LEFT JOIN LIEUFORMATION ON PST_LIEUFORM=PLF_LIEUFORM LEFT JOIN DOSSIER ON PST_NODOSSIER=DOS_NODOSSIER','PST_CODESTAGE','(PST_LIBELLE||" "||PST_LIBELLE1) AS Libelle,PLF_LIBELLE,DOS_LIBELLE',StWhere,StOrder, True,-1);
            {$ELSE}
            LookupList(THEdit(Sender),'Liste des stages','STAGE LEFT JOIN LIEUFORMATION ON PST_LIEUFORM=PLF_LIEUFORM LEFT JOIN DOSSIER ON PST_NODOSSIER=DOS_NODOSSIER','PST_CODESTAGE','(PST_LIBELLE||" "||PST_LIBELLE1) AS Libelle,PLF_LIBELLE,DOS_LIBELLE',StWhere,StOrder, True,-1);
            {$ENDIF}
        end
        else
        begin
            {$IFNDEF EAGLCLIENT}
            LookupList(THDBEdit(Sender),'Liste des stages','STAGE LEFT JOIN LIEUFORMATION ON PST_LIEUFORM=PLF_LIEUFORM','PST_CODESTAGE','(PST_LIBELLE||" "||PST_LIBELLE1) AS Libelle,PLF_LIBELLE',StWhere,StOrder, True,-1);
            {$ELSE}
            LookupList(THEdit(Sender),'Liste des stages','STAGE LEFT JOIN LIEUFORMATION ON PST_LIEUFORM=PLF_LIEUFORM','PST_CODESTAGE','(PST_LIBELLE||" "||PST_LIBELLE1) AS Libelle,PLF_LIBELLE',StWhere,StOrder, True,-1);
            {$ENDIF}
        end;
    end;
end;

procedure TOM_SESSIONSTAGE.BFormationClick(Sender : TObject);
begin
    If TypeSaisie = 'SAISIEWEB' Then AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','SAISIEWEBSESSION;'+GetField('PSS_CODESTAGE')+';'+IntToStr(GetField('PSS_ORDRE'))+';'+GetField('PSS_MILLESIME'));
    If (TypeSaisie = 'SAISIE') Or (TypeSaisie = 'SOUSSESSION') Then //PT27
    begin
        If GetField('PSS_CLOTUREINSC') = 'X' then AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','SAISIESESSION;'+GetField('PSS_CODESTAGE')+';'+IntToStr(GetField('PSS_ORDRE'))+';'+GetField('PSS_MILLESIME')+';'+'ACTION=CONSULTATION')
        Else AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','SAISIESESSION;'+GetField('PSS_CODESTAGE')+';'+IntToStr(GetField('PSS_ORDRE'))+';'+GetField('PSS_MILLESIME'));
    end;
end;

procedure TOM_SESSIONSTAGE.OnChangeField(F : TField);
var LaDate : TDateTime;
    Libelle : THLabel;
    Q,QQ : TQuery;
    Nbj,NbJours,i : Integer ;
    St,StDate      : String;
    SystemTime0 : TSystemTime;
    TotalCout : Double;
    DD,DF : Variant;
    MenuMail : TPopuPMenu;
    TobInvest : Tob;
    PluriAnnu : String; //PT10
    Compon : TComponent;
    IMax : integer; //PT25
begin
  inherited;
  	//PT27 Sauvegarde de la nouvelle valeur
  	If DS.State in [dsEdit] Then TobModifiedFields.PutValue(F.FieldName, GetField(F.FieldName));

  
    if F.FieldName = 'PSS_PGTYPESESSION' then
    Begin
        If DS.State in [dsEdit,dsBrowse] Then AccesBoutonsSousSession; //PT25
    End
    Else if F.FieldName = 'PSS_CODESTAGE' then
    begin
        If DS.State in [dsInsert] then
        begin
            RendDureeStage;
            If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'" AND PST_CODESTAGE="'+GetField('PSS_CODESTAGE')+'"') Then
            SetField('PSS_MILLESIME',MillesimeEC)
            Else SetField('PSS_MILLESIME','0000');
            //PT23
            // On bloque les champs prédéfini et nodossier à la sélection d'un stage pour empêcher
            // de créer une session standard basée sur un stage dossier et vice versa
            If PGBundleCatalogue Then
            Begin
                SetControlEnabled('PSS_PREDEFINI', GetField('PSS_CODESTAGE')='');
                SetControlEnabled('PSS_NODOSSIER', GetField('PSS_CODESTAGE')='');
            End;
        end;
        Libelle := THLabel(GetControl('LSTAGE'));
        If GetField('PST_CODESTAGE') <> '' Then
        begin
            Q := OpenSQL('SELECT PST_LIBELLE,PST_LIBELLE1 FROM STAGE WHERE PST_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PST_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);
            if not Q.eof then Libelle.Caption := Q.FindField('PST_LIBELLE').AsString+' '+Q.FindField('PST_LIBELLE1').AsString;
            Ferme(Q);
        end
        Else Libelle.Caption := '';
        If (InitStage = '') and (GetField('PSS_CODESTAGE') <> '') Then
        begin
		    //PT25
		    If TFFiche(Ecran).Name = 'SESSIONSIMPLE' Then
		    Begin
			    // Récupération du PSS_ORDRE immédiatement dans le but de pouvoir saisir les inscriptions sans fermer puis rouvrir la fiche
			    If (DS.State in [dsInsert]) then
			    begin
			        st  :=  'SELECT MAX(PSS_ORDRE) FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'"';
			        QQ := OpenSQL(st,TRUE) ;
			        if Not QQ.EOF then
			        begin
			            IMax := QQ.Fields[0].AsInteger;
			            if IMax <> 0 then
			            IMax := IMax + 1
			            else
			            IMax := 1;
			        end
			        else IMax := 1 ;
			        Ferme(QQ) ;
			        SetField ('PSS_ORDRE', IMax);
			        SetField ('PSS_IDSESSIONFOR', GetField('PSS_CODESTAGE')+IntToStr(IMax)); //PT19
			    end;
		    end;
		    
            Nbj := 0;
            TFFiche(Ecran).caption := 'Saisie session pour la formation : '+Rechdom('PGSTAGEFORM',GetField('PSS_CODESTAGE'),FALSE);
            UpdateCaption(Ecran);
            GetLocalTime(SystemTime0);
            LaDate  :=  SystemTimeToDateTime(SystemTime0);
            StDate := DateToStr(LaDate);
            LaDate := StrToDate(StDate);
            
            If TypeSaisie <> 'CURSUS' then SetField ('PSS_DATEDEBUT',LaDate);
            st  :=  'SELECT PST_PREDEFINI,PST_COUTBUDGETE,PST_COUTFONCT,PST_COUTORGAN,PST_COUTUNITAIRE,PST_COUTMOB,PST_COUTEVAL,'+
                    'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,'+
                    'PST_COMPTERENDU,PST_ATTESTPRESENC,PST_QUESTAPPREC,PST_QUESTEVALUAT,PST_SUPPCOURS,PST_VIDEOPROJ,PST_RETROPROJ,'+
                    'PST_JEUXROLE,PST_ETUDECAS,PST_AUTREEVALUAT,PST_AUTREMOYEN,'+
                    'PST_INCLUSDECL,PST_ACTIONFORM,PST_DUREESTAGE,'+
                    'PST_NBSTAMIN,PST_NBSTAMAX,'+ //PT25
                    'PST_JOURSTAGE,PST_CENTREFORMGU,PST_LIEUFORM,PST_NATUREFORM,PST_TYPCONVFORM,PST_ORGCOLLECTPGU,PST_ORGCOLLECTSGU'+
                    ' FROM STAGE WHERE PST_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PST_MILLESIME="'+GetField('PSS_MILLESIME')+'"';
            QQ := OpenSQL(st,TRUE) ;
            if Not QQ.EOF then
            begin
                Nbj  :=  Trunc(QQ.FindField('PST_JOURSTAGE').AsFloat);
                SetField ('PSS_CENTREFORMGU',QQ.FindField('PST_CENTREFORMGU').AsString);
                If (PGDroitMultiForm) And (QQ.FindField('PST_PREDEFINI').AsString <> '') then SetField ('PSS_PREDEFINI',QQ.FindField('PST_PREDEFINI').AsString); //PT27
                SetField('PSS_LIEUFORM',QQ.FindField('PST_LIEUFORM').AsString);
                SetField('PSS_ORGCOLLECTPGU',QQ.FindField('PST_ORGCOLLECTPGU').AsString);
                SetField('PSS_ORGCOLLECTSGU',QQ.FindField('PST_ORGCOLLECTSGU').AsString);
                SetField('PSS_NATUREFORM',QQ.FindField('PST_NATUREFORM').AsString);
                SetField('PSS_TYPCONVFORM',QQ.FindField('PST_TYPCONVFORM').AsString);
                SetField('PSS_COUTPEDAG',QQ.FindField('PST_COUTBUDGETE').AsFloat);
                SetField('PSS_COUTFONCT',QQ.FindField('PST_COUTFONCT').AsFloat);
                SetField('PSS_COUTORGAN',QQ.FindField('PST_COUTORGAN').AsFloat);
                SetField('PSS_COUTUNITAIRE',QQ.FindField('PST_COUTUNITAIRE').AsFloat);
                SetField('PSS_COUTEVAL',QQ.FindField('PST_COUTEVAL').AsFloat);
                SetField('PSS_COUTMOB',QQ.FindField('PST_COUTMOB').AsFloat);
                SetField('PSS_FORMATION1',QQ.FindField('PST_FORMATION1').AsString);
                SetField('PSS_FORMATION2',QQ.FindField('PST_FORMATION2').AsString);
                SetField('PSS_FORMATION3',QQ.FindField('PST_FORMATION3').AsString);
                SetField('PSS_FORMATION4',QQ.FindField('PST_FORMATION4').AsString);
                SetField('PSS_FORMATION5',QQ.FindField('PST_FORMATION5').AsString);
                SetField('PSS_FORMATION6',QQ.FindField('PST_FORMATION6').AsString);
                SetField('PSS_FORMATION7',QQ.FindField('PST_FORMATION7').AsString);
                SetField('PSS_FORMATION8',QQ.FindField('PST_FORMATION8').AsString);
                SetField('PSS_COMPTERENDU',QQ.FindField('PST_COMPTERENDU').AsString);
                SetField('PSS_ATTESTPRESENC',QQ.FindField('PST_ATTESTPRESENC').AsString);
                SetField('PSS_QUESTAPPREC',QQ.FindField('PST_QUESTAPPREC').AsString);
                SetField('PSS_QUESTEVALUAT',QQ.FindField('PST_QUESTEVALUAT').AsString);
                SetField('PSS_SUPPCOURS',QQ.FindField('PST_SUPPCOURS').AsString);
                SetField('PSS_VIDEOPROJ',QQ.FindField('PST_VIDEOPROJ').AsString);
                SetField('PSS_RETROPROJ',QQ.FindField('PST_RETROPROJ').AsString);
                SetField('PSS_JEUXROLE',QQ.FindField('PST_JEUXROLE').AsString);
                SetField('PSS_ETUDECAS',QQ.FindField('PST_ETUDECAS').AsString);
                SetField('PSS_AUTREEVALUAT',QQ.FindField('PST_AUTREEVALUAT').AsString);
                SetField('PSS_AUTREMOYEN',QQ.FindField('PST_AUTREMOYEN').AsString);
                SetField('PSS_INCLUSDECL',QQ.FindField('PST_INCLUSDECL').AsString);
                SetField('PSS_ACTIONFORM',QQ.FindField('PST_ACTIONFORM').AsString);
                SetField('PSS_JOURSTAGE',QQ.FindField('PST_JOURSTAGE').AsFloat);      //DB2
                SetField('PSS_DUREESTAGE',QQ.FindField('PST_DUREESTAGE').AsFloat);     //DB2
                SetField('PSS_NBSTAMIN',QQ.FindField('PST_NBSTAMIN').AsFloat);       //PT25
                SetField('PSS_NBRESTAGPREV',QQ.FindField('PST_NBSTAMAX').AsFloat);   //PT25
            end
            Else PGIBox('Une erreur est sruvenue dans la récupération de la formation #13#10 veuillez vérifier la saisie du code de la formation',Ecran.Caption);
            Ferme (QQ);

            If (NbJ>0) and (LaDate <> IDate1900) Then LaDate   :=  PLUSDATE (LaDate, NbJ-1,'J');
            LaDuree  :=  Nbj;
            SetField ('PSS_DATEFIN',LaDate);
            SetField ('PSS_DEBUTDJ','MAT');
            SetField ('PSS_FINDJ','PAM');
            //DEBUT PT1
            If TypeSaisie <> 'CURSUS' then
            begin
                Q := OpenSQL('SELECT PIS_INVESTFORM,PIS_COUTPEDAG,PIS_COUTSALAIRE,PIS_FRAISFORM,PIS_TAUXPEDAG,PIS_TAUXSALFRAIS,PIS_TAUXSALAIRE,PIS_MONTANT FROM INVESTSESSION '+
                             'WHERE PIS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PIS_ORDRE=-1 AND PIS_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);
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
            end;
            //FIN PT1
        end;
        //PT25
        // On cache / affiche certains éléments suivant si le codestage est renseigné
        If TFFiche(Ecran).Name = 'SESSIONSIMPLE' Then
        Begin
            SetControlEnabled('BPLUSDINFOS', GetField('PSS_CODESTAGE') <> '');
            If (TTabSheet(GetControl('PCHAMPLIBRE')).TabVisible) And (GetField('PSS_CODESTAGE') = '') Then TTabSheet(GetControl('PCHAMPLIBRE')).TabVisible := False;
            If (GetField('PSS_CODESTAGE') = '') And (GrilleSal.CellValues[COL_SALARIE,1] = '') Then TTabSheet(GetControl('PINSCRIPTIONS')).TabVisible := False;
        End;
    end
    Else if (F.FieldName  = 'PSS_DATEDEBUT') and (DS.State in [dsInsert]) then
    begin
        If GetField('PSS_DATEDEBUT') <> IDate1900 then
        begin
            If LaDuree>0 Then LaDate  :=  PLUSDATE (GetField ('PSS_DATEDEBUT'),Trunc(LaDuree-1),'J')
            Else LaDate  :=  GetField ('PSS_DATEDEBUT');
            SetField ('PSS_DATEFIN',LaDate);
        end;
    end;

    If TFFiche(Ecran).FTypeAction <> TaConsult Then
    begin
        If F.FieldName = 'PSS_COUTMOB' Then
        begin
            TotalCout := GetField('PSS_COUTMOB')+GetField('PSS_COUTORGAN')+GetField('PSS_COUTFONCT');
            If TotalCout <> 0 Then
            begin
                SetField('PSS_COUTPEDAG',TotalCout);
                SetControlEnabled('PSS_COUTPEDAG',False);
            end
            Else SetControlEnabled('PSS_COUTPEDAG',True);
        end
        Else If F.FieldName = 'PSS_COUTORGAN' Then
        begin
            TotalCout := GetField('PSS_COUTMOB')+GetField('PSS_COUTORGAN')+GetField('PSS_COUTFONCT');
            If TotalCout <> 0 Then
            begin
                SetField('PSS_COUTPEDAG',TotalCout);
                SetControlEnabled('PSS_COUTPEDAG',False);
            end
            Else SetControlEnabled('PSS_COUTPEDAG',True);
        end
        Else If F.FieldName = 'PSS_COUTFONCT' Then
        begin
            TotalCout := GetField('PSS_COUTMOB')+GetField('PSS_COUTORGAN')+GetField('PSS_COUTFONCT');
            If TotalCout <> 0 Then
            begin
                SetField('PSS_COUTPEDAG',TotalCout);
                SetControlEnabled('PSS_COUTPEDAG',False);
            end
            Else SetControlEnabled('PSS_COUTPEDAG',True);
        end;
    end;

    If (F.FieldName = 'PSS_DATEFIN') Then
    begin
        If not AfterInserting then
        begin
            DD := GetField('PSS_DATEDEBUT');
            DF := GetField('PSS_DATEFIN');
            NbJours := GetField('PSS_JOURSTAGE');
            PluriAnnu := GetField('PSS_TYPCONVFORM'); //PT10
            If (DD <> IDate1900) and (NbJours <> (DF-DD+1)) and (NbJours>0) And (PluriAnnu <> 'PLU') Then      //PT10
            PGIBox('Attention les dates ne correspondent pas avec le nombre de jours du stage',TFFiche(Ecran).Caption);
        end;
    end
    Else If (F.FieldName = 'PSS_CLOTUREINSC') and (TypeSaisie <> 'CURSUS') Then
    begin
        Compon := GetControl('TPMAIL') As TComponent;
        MenuMail := Compon As TPopUpMenu;
        If GetField('PSS_CLOTUREINSC') = 'X' Then
        begin
            MenuMail.Items[0].Enabled := True;
            MenuMail.Items[1].Enabled := True;
            MenuMail.Items[2].Enabled := False;
        end
        Else
        begin
            MenuMail.Items[0].Enabled := False;
            MenuMail.Items[1].Enabled := False;
            MenuMail.Items[2].Enabled := True;
        end;
    end
    Else If F.FieldName = 'PSS_JOURSTAGE' then
    begin
        If isFieldModified('PSS_JOURSTAGE') then
        begin
            LaDuree := GetField('PSS_JOURSTAGE');
            If GetField('PSS_DATEDEBUT') <> IDate1900 then
            begin
                If LaDuree>0 Then LaDate  :=  PLUSDATE (GetField ('PSS_DATEDEBUT'),Trunc(LaDuree-1),'J')
                Else LaDate  :=  GetField ('PSS_DATEDEBUT');
                SetField ('PSS_DATEFIN',LaDate);
            end;
        end;
    end
    //PT25
    Else if F.FieldName = 'PSS_ORDRE' then
    begin
        if (TFFiche(Ecran).Name = 'SESSIONSIMPLE') And (DS.State in [dsInsert]) and (GetField('PSS_ORDRE') <> 0) Then
            SetControlVisible('PINSCRIPTIONS', True);
    end;
end;

procedure TOM_SESSIONSTAGE.OnLoadRecord;
var Q : TQuery;
    //TobInvest : Tob;
    //i : Integer;
    //GInvest : THGrid;
    TobStage : Tob;
    {$IFDEF EMANAGER}NbInsc : Integer;{$ENDIF} //PT14
    HeureEntre : String;
begin
	inherited;

	MAJSousSessions := False; //PT27

	If (TypeSaisie <> 'SOUSSESSION') And (GetField('PSS_CODESTAGE') <> '') then //PT27
	begin
		Q := openSQL('SELECT PST_BLOCNOTE FROM STAGE WHERE PST_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PST_MILLESIME="0000"',True);
		TobStage := Tob.Create('MaTob',Nil,-1);
		TobStage.LoadDetailDB('Table','','',Q,False);
		Ferme(Q);
		SetControlText('BLOCNOTESTAGE',TobStage.Detail[0].GetValue('PST_BLOCNOTE'));
		TobStage.Free;
	end;
	//PT13 - Début
	//GInvest := THGrid(GetControl('GINVESTSESSION'));
	//If GInvest <> nil then
	//begin
	//     If ExisteSQL('SELECT PIS_INVESTFORM FROM INVESTSESSION WHERE PIS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'"'+
	//      ' AND PIS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+
	//     ' AND (PIS_INVESTFORM="PDF" OR PIS_INVESTFORM="ALT")') then
	//        GInvest.ColFormats[0] := 'CB=PGFINANCEFORMATION';
	//end;
	//PT13 - Fin
	SetControLCaption('TPSS_COUTSALAIRE','Coût(s) animateur(s)');
	InitDDSession := GetField('PSS_DATEDEBUT');
	InitDFSession := GetField('PSS_DATEFIN');
	{$IFNDEF EMANAGER} //PT14
	If TypeSaisie <> 'CURSUS' then
	begin
		SetControlEnabled('BFirst',False);
		SetControlEnabled('BPrev',False);
		SetControlEnabled('BNext',False);
		SetControlEnabled('BLast',False);
		If AfterInserting then
		begin
			//PT25 Appel de la fonction
			//SetControlEnabled('BANIMAT',False);
			//SetControlEnabled('BFORMATION',False);
			//SetControlEnabled('BMAIL',False);
			//SetControlEnabled('BCONVOC',False);
			//SetControlEnabled('BPRESENCE',False);
			//SetControlEnabled('BTNFORMATION',True);
            AccesBoutons(False);
			SetControlEnabled('BDUPLIQUER',False); //PT9
		end
		Else
		begin
			If GetField('PSS_PGTYPESESSION') <> 'ENT' then
			begin
				//PT25 Appel de la fonction
				//SetControlEnabled('BANIMAT',True);
				//SetControlEnabled('BFORMATION',True);
				//SetControlEnabled('BCONVOC',True);
				//SetControlEnabled('BPRESENCE',True);
				//SetControlEnabled('BTNFORMATION',False);*)
                AccesBoutons(True);
				SetControlEnabled('BDUPLIQUER',True); //PT9
			end
			//PT27 Empêche de changer le type si des sous-sessions existent
			Else If (DS.State in [dsBrowse]) Then 
			Begin
				If ExisteSQL('SELECT 1 FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'"') Then
				Begin
					SetControlEnabled('PSS_PGTYPESESSION', False);
				   	SetControlVisible('BMULTISALSS',       True);
				   	SetControlVisible('BRECUPERATIONSS',   True);
				End
                Else
                Begin
                    SetControlEnabled('BMULTISALSS',       False);
                    SetControlEnabled('BRECUPERATIONSS',   False);
                End;
			End;
		end;
		If GetField('PSS_CLOTUREINSC') = 'X' Then SetControlEnabled('BCLOTUREINSC',False)
		else SetControlEnabled('BCLOTUREINSC',True);
		//DEBUT PT1
		//PT13 - Début
		//Q := OpenSQL('SELECT PIS_INVESTFORM,PIS_COUTPEDAG,PIS_COUTSALAIRE,PIS_FRAISFORM,PIS_TAUXPEDAG,PIS_TAUXSALAIRE,PIS_TAUXSALFRAIS,PIS_MONTANT FROM INVESTSESSION '+
		//'WHERE PIS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PIS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PIS_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);  //DB2
		//TobInvest := Tob.Create('Les investissements',Nil,-1);
		//TobInvest.LoadDetailDB('INVESTSESSION','','',Q,False);
		//Ferme(Q);
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
		//        GInvest.CellValues[7,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_MONTANT'));
		//end;
		//TobInvest.Free;
		//PT13 - Fin
		//FIN PT1
	end;
	{$ENDIF}
	RendDureeStage;
	InitStage := GetField('PSS_CODESTAGE');
	If GetField('PSS_AVECCLIENT') = 'X' Then InitAvecClient := True
	Else InitAvecClient := False;
	
	If (TypeSaisie <> 'SOUSSESSION') And (GetField('PSS_CODESTAGE') <> '') then //PT27
	begin
		If TypeSaisie <> 'CURSUS' then
		begin
			TFFiche(Ecran).caption := 'Saisie session pour la formation : '+Rechdom('PGSTAGEFORM',GetField('PSS_CODESTAGE'),FALSE);
			UpdateCaption(Ecran);
		end;
	end;
	
	//PT14 - Début
	{$IFDEF EMANAGER}
	Q := OpenSQL ('SELECT COUNT(*) AS NBINSC FROM FORMATIONS WHERE PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFO_ORDRE="'+IntToStr(GetField('PSS_ORDRE'))+'" AND PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'"', True);
	If Not Q.EOF Then NbInsc := Q.FindField('NBINSC').AsInteger;
	Ferme(Q);
	SetControlText('NBRESTAGEDISPO',IntToStr(GetField('PSS_NBRESTAGPREV')-NbInsc));
	{$ENDIF}
	//PT14 - Fin

	//PT15 - Début
	if TFFiche(Ecran).FTypeAction <> taCreat then
	begin
		HeureEntre := FormatDateTime('hh:mm', GetField('PSS_HEUREDEBUT'));
		SetControlText('HEUREDEBUT', HeureEntre);
		HeureEntre := FormatDateTime('hh:mm', GetField('PSS_HEUREFIN'));
		SetControlText('HEUREFIN', HeureEntre);
	End;
	//PT15 - Fin
	
	//PT27
	If TypeSaisie = 'SOUSSESSION' Then
	Begin
		Q := OpenSQL('SELECT PSS_LIBELLE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_ORDRE="'+IntToStr(GetField('PSS_NUMSESSION'))+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'"', True);
		If Not Q.EOF Then SetControlText('LSESSIONENTETE', Q.FindField('PSS_LIBELLE').AsString);
		Ferme(Q);
		SetControlText('CODESTAGE', GetField('PSS_CODESTAGE'));
	End;

	// Initialise la grille des inscrits
    If TFFiche(Ecran).Name = 'SESSIONSIMPLE' Then AfficheStagiaires; //PT25
    
    //PT27 Au chargement, on sauvegarde les données actuelles en vue de connaître tous les champs modifiés lors de l'update
    if (not(DS.State in [dsInsert])) then   //PT29
       TobModifiedFields.GetEcran(Ecran);
    TobModifiedFields.SetAllModifie(False);
end;

procedure TOM_SESSIONSTAGE.OnNewRecord;
Var QQ      : TQuery ;
    Nbj : Integer ;
    St,StDate      : String;
    LaDate  : TDateTime;
    SystemTime0 : TSystemTime;
begin
inherited;
    If (PGDroitMultiForm) Or (Not PGBundleCatalogue) then SetField('PSS_PREDEFINI','STD')
    else SetField('PSS_PREDEFINI','DOS');
    
    SetField('PSS_NODOSSIER',V_PGI.NoDossier);
    //SetField('PSS_PREDEFINI','DOS');
    SetField('PSS_NUMENVOI',0);
    If TypeSaisie = 'SOUSSESSION' Then 
    Begin
    	SetField('PSS_PGTYPESESSION','SOS');
    	SetField('PSS_NUMSESSION',    NumSessionMaitre);
    End
    Else
    	SetField('PSS_PGTYPESESSION','AUC');
    SetControlEnabled ('PSS_ORDRE',FALSE);
    If LeStage <> '' Then SetField ('PSS_CODESTAGE', LeStage);
    GetLocalTime(SystemTime0);
    LaDate  :=  SystemTimeToDateTime(SystemTime0);
    StDate := DateToStr(LaDate);
    LaDate := StrToDate(StDate);
    SetField ('PSS_DATEDEBUT',LaDate);
    
    If Not (ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+Millesime+'" AND PST_CODESTAGE="'+LeStage+'"')) then
    begin
        Millesime := '0000';
    end;
    SetField('PSS_MILLESIME',Millesime);
    
    Nbj := 0;
    if (TypeSaisie <> 'SOUSSESSION') And (GetField ('PSS_CODESTAGE')  <>  '')  then //PT27
    begin
        st  :=  'SELECT PST_COUTBUDGETE,PST_COUTFONCT,PST_COUTORGAN,PST_COUTUNITAIRE,PST_COUTMOB,'+
                'PST_JOURSTAGE,PST_CENTREFORMGU,PST_LIEUFORM,PST_NATUREFORM,PST_TYPCONVFORM,PST_ORGCOLLECTPGU,PST_ORGCOLLECTSGU,'+
                'PST_NBSTAMIN,PST_NBSTAMAX '+ //PT25
                'FROM STAGE WHERE PST_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PST_MILLESIME="'+Millesime+'"';
        QQ := OpenSQL(st,TRUE) ;
        if Not QQ.EOF then
        begin
            Nbj  :=  Trunc(QQ.FindField('PST_JOURSTAGE').AsFloat);
            SetField ('PSS_CENTREFORMGU',QQ.FindField('PST_CENTREFORMGU').AsString);
            SetField('PSS_LIEUFORM',QQ.FindField('PST_LIEUFORM').AsString);
            SetField('PSS_ORGCOLLECTPGU',QQ.FindField('PST_ORGCOLLECTPGU').AsString);
            SetField('PSS_ORGCOLLECTSGU',QQ.FindField('PST_ORGCOLLECTSGU').AsString);
            SetField('PSS_NATUREFORM',QQ.FindField('PST_NATUREFORM').AsString);
            SetField('PSS_TYPCONVFORM',QQ.FindField('PST_TYPCONVFORM').AsString);
            SetField('PSS_COUTPEDAG',QQ.FindField('PST_COUTBUDGETE').AsFloat);
            SetField('PSS_COUTFONCT',QQ.FindField('PST_COUTFONCT').AsFloat);
            SetField('PSS_COUTORGAN',QQ.FindField('PST_COUTORGAN').AsFloat);  //DB2
            SetField('PSS_COUTUNITAIRE',QQ.FindField('PST_COUTUNITAIRE').AsFloat);
            SetField('PSS_COUTMOB',QQ.FindField('PST_COUTMOB').AsFloat);
            //PT25
            SetField('PSS_NBSTAMIN',QQ.FindField('PST_NBSTAMIN').AsFloat);
            SetField('PSS_NBRESTAGPREV',QQ.FindField('PST_NBSTAMAX').AsFloat);
        end;
        Ferme (QQ);
    End
    //PT27
    // Dans le cas d'une sous-session, on reprend les mêmes infos que la session en-tête
    Else If TypeSaisie = 'SOUSSESSION' Then
    Begin
	    DuplSession   := LeStage;
	    DuplOrdre     := NumSessionMaitre;
	    DuplMillesime := Millesime;
	    ADupliquer    := True;
    End;
    
    If NbJ>0 Then LaDate   :=  PLUSDATE (LaDate, NbJ-1,'J');
    LaDuree  :=  Nbj;

    SetField ('PSS_DATEFIN',LaDate);
    SetField ('PSS_DEBUTDJ','MAT');
    SetField ('PSS_FINDJ','PAM');
    SetField ('PSS_FORMPREVIS','REA');
    SetField('PSS_EFFECTUE','-');
    SetField('PSS_CLOTUREINSC','-');
    SetField('PSS_VALIDORG','-');
    SetField('PSS_HEUREDEBUT',0);
    SetField('PSS_HEUREFIN',0);
    SetField('PSS_AVECCLIENT','-');
    
    //PT15 - Début
    SetControlText('HEUREDEBUT','00:00');  //PT20
    SetControlText('HEUREFIN',  '00:00');
    //PT15 - Fin
    
    //PT9 - Début
    If ADupliquer = true Then
    begin
        { On récupère les informations actuelles et on repositionne le booléen à faux afin de ne
        plus effectuer de duplication si clic sur le bouton Nouveau.
        }
        RecupInfosSession;
        ADupliquer := false;
    end;
    //PT9 - Fin

    //PT25 - Début
    If TFFiche(Ecran).Name = 'SESSIONSIMPLE' Then
    Begin
        SetControlEnabled('BPLUSDINFOS',  False);
    End;

    AccesBoutons(False);
    SetControlEnabled('BSOUSSESSION', False);
    //PT25 - Fin
end;

procedure TOM_SESSIONSTAGE.OnUpdateRecord;
Var QQ,Q      : TQuery ;
    IMax,i,rep    : Integer ;
    St,Question,Mess,ExerForm,LibEmploi : String;
    TobAnim,T : Tob;
    Salaire : Double;
    DebExerForm,FinExerForm : TDateTime;
begin
 inherited;
    If GetField('PSS_PREDEFINI') = 'STD' then SetField('PSS_NODOSSIER','000000');
    Mess := '';
    If existeSQL('SELECT PSS_ORDRE FROM SESSIONSTAGE WHERE PSS_ORDRE <> '+IntToStr(GetField('PSS_ORDRE'))+''+  //DB2
                 ' AND PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND '+
                 'PSS_DATEDEBUT="'+UsDateTime(GetField('PSS_DATEDEBUT'))+'" AND PSS_DATEFIN="'+UsDatetime(GetField('PSS_DATEFIN'))+'"') then
    begin
        If PGIAsk('Attention il existe déja une session pour la même formation avec des dates identiques,#13#10 Voulez-vous continuer ?',Ecran.Caption)=MrNo then
        begin
            LastError := 1;
            Exit;
        end;
    end;
    If (GetField('PSS_CODESTAGE') = '') then
    begin
        LastError := 1;
        LastErrorMsg := 'Vous devez renseigner le code du stage';
        exit;
    end;
    RendMillesimeRealise(DebExerForm,FinExerForm);
    //DEBUT PT3
    If GetField('PSS_NATUREFORM') <> '004' then
    begin
        If (GetField('PSS_DATEDEBUT') <> IDate1900) and ((GetField ('PSS_DATEDEBUT')<DebExerForm) or (GetField ('PSS_DATEDEBUT')>FinExerForm)) then
        begin
            If DebExerForm = IDate1900 then Mess := 'Saisie impossible car aucun exercice de formation n''est ouvert';
        end
        else
        begin
            If (GetField('PSS_DATEFIN') <> IDate1900) and ((GetField ('PSS_DATEFIN')<DebExerForm) or (GetField ('PSS_DATEFIN')>FinExerForm)) then
            begin
                If DebExerForm <> IDate1900 then
                begin
                    Rep := PGIAsk('Attention, la date de fin n''est pas comprise dans l''éxercie de formation actuel du '+DateToStr(DebExerForm)
                    +' au '+DateToStr(FinExerForm)+',#13#10voulez-vous continuer',Ecran.Caption);
                    If Rep <> MrYes then
                    begin
                        LastError := 1;
                        SetFocusControl('PSS_DATEFIN');
                        Exit;
                    end;
                end
                else If Mess = '' then Mess := 'Saisie impossible car aucun exercice de formation n''est ouvert';
            end;
        end;
    end;
    //FIN PT3
    if GetField ('PSS_DATEDEBUT') > GetField ('PSS_DATEFIN') then
    begin
        Mess := Mess+'#13#10La date de début est supérieure à la date de fin';
        SetFocusControl ('PSS_DATEBUT');
    end;
    If Mess <> '' then
    begin
        PGIBox(Mess,Ecran.Caption);
        LastError := 1;
        Exit;
    end;
    If (InitAvecClient=False) and (GetField('PSS_AVECCLIENT') = 'X') Then
    begin
        Q := OpenSQL('SELECT * FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PAN_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' '+  //DB2
                     'AND PAN_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);
        TobAnim := Tob.Create('SESSIONANIMAT',Nil,-1);
        TobAnim.LoadDetailDB('SESSIONANIMAT','','',Q,False);
        If TobAnim.Detail.Count>0 then
        begin
            Question := 'Attention, vous avez coché formation avec clients,#13#10 les salaires des animateurs pour cette session vont être mis à 0, #13#10 Voulez-vous continuer ?';
            If PGiAsk(Question,Ecran.Caption) <> MrYes then
            begin
                TobAnim.Free;
                Ferme(Q);
                SetFocusControl('PSS_AVECCLIENT');
                LastError := 1;
                Exit;
            end;
        end;
        For i := 0 To TobAnim.Detail.Count-1 do
        begin
            T := TobAnim.Detail[i];
            T.PutValue('PAN_SALAIREANIM',0);
            T.UpdateDB(False);
        end;
        TobAnim.Free;
        Ferme(Q);
    end;
    If (InitAvecClient=True) and (GetField('PSS_AVECCLIENT') = '-') Then
    begin
        Q := OpenSQL('SELECT * FROM SESSIONANIMAT WHERE PAN_SALARIE <> "" AND PAN_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PAN_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' '+  //DB2
                     'AND PAN_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);
        TobAnim := Tob.Create('SESSIONANIMAT',Nil,-1);
        TobAnim.LoadDetailDB('SESSIONANIMAT','','',Q,False);
        If TobAnim.Detail.Count>0 then
        begin
            Question := 'Attention, vous avez enlevé formation avec clients,#13#10 voulez-vous refaire le calcul des salaires des animateurs ?';
            If PGiAsk(Question,Ecran.Caption) <> MrYes then
            begin
                TobAnim.Free;
                Ferme(Q);
                SetFocusControl('PSS_AVECCLIENT');
                LastError := 1;
                Exit;
            end;
        end;
        For i := 0 To TobAnim.Detail.Count-1 do
        begin
            T := TobAnim.Detail[i];
            Salaire := 0;
            if VH_Paie.PGForValoSalaire = 'VCR' then Salaire := ForTauxHoraireReel(T.GetValue('PAN_SALARIE'));
            if VH_Paie.PGForValoSalaire = 'VCC' then
            begin
                ExerForm := '';
                LibEmploi := '';
                Mess := '';
                Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_DATEDEBUT<="'+UsDateTime(GetField('PSS_DATEDEBUT'))+'" '+
                             'AND PFE_DATEFIN>="'+UsDateTime(GetField('PSS_DATEFIN'))+'"',True);
                If Not Q.eof then ExerForm := Q.FindField('PFE_MILLESIME').AsString
                Else Mess := Mess+'#13#10- Aucun exercice de formation n''existe pour les dates saisies';
                Ferme(Q);
                Q := OpenSQL('SELECT PSA_LIBELLEEMPLOI FROM SALARIES WHERE PSA_SALARIE="'+T.GetValue('PAN_SALARIE')+'"',True);
                If not Q.eof then LibEmploi := Q.FindField('PSA_LIBELLEEMPLOI').AsString
                Else Mess := Mess+'#13#10- Le libellé emploi du salarié formateur n''est pas renseigné dans la fiche salarié';
                Ferme(Q);
                If (LibEmploi <> '') and (ExerForm <> '') then Salaire := ForTauxHoraireCategoriel(LibEmploi,ExerForm)
                Else PGIBox('Calcul du salaire impossible car '+Mess,T.GetValue('PAN_LIBELLE'));
            end;
            T.PutValue('PAN_SALAIREANIM',Salaire*T.GetValue('PAN_NBREHEURE'));
            T.UpdateDB(False);
        end;
        TobAnim.Free;
        Ferme(Q);
    end;
    If (DS.State in [dsInsert]) then
    begin     // increment automatique du numero d'ordre au moment de la creation
        st  :=  'SELECT MAX(PSS_ORDRE) FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'"';
        QQ := OpenSQL(st,TRUE) ;
        if Not QQ.EOF then
        begin
            IMax := QQ.Fields[0].AsInteger;
            if IMax <> 0 then
            IMax := IMax + 1
            else
            IMax := 1;
        end
        else IMax := 1 ;
        Ferme(QQ) ;
        SetField ('PSS_ORDRE', IMax);
        SetField ('PSS_IDSESSIONFOR', GetField('PSS_CODESTAGE')+IntToStr(IMax)); //PT19

        //PT27
        If GetField('PSS_PGTYPESESSION') = 'SOS' Then
        Begin
              // increment automatique du numero de sous-session au moment de la creation
            st  :=  'SELECT MAX(PSS_NOSESSIONMULTI) FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+Millesime+'" AND PSS_PGTYPESESSION="SOS" AND PSS_NUMSESSION="'+NumSessionMaitre+'"';
            QQ := OpenSQL(st,TRUE) ;
            if Not QQ.EOF then
            begin
                IMax := QQ.Fields[0].AsInteger;
                if IMax <> 0 then
                IMax := IMax + 1
                else
                IMax := 1;
            end
            else IMax := 1 ;
            Ferme(QQ) ;
            SetField ('PSS_NOSESSIONMULTI', IMax);
            SetField ('PSS_IDSESSIONFORMS', GetField('PSS_CODESTAGE')+NumSessionMaitre); //PT19
        End
        Else 
        Begin
        	SetField('PSS_NOSESSIONMULTI',   0);
        	SetField ('PSS_IDSESSIONFORMS', '');
        End;
    end;

    //PT15
	if GetControlText('HEUREDEBUT') <> '' then SetField('PSS_HEUREDEBUT', StrToTime(GetControlText('HEUREDEBUT')));
	if GetControlText('HEUREFIN') <> '' then SetField('PSS_HEUREFIN', StrToTime(GetControlText('HEUREFIN')));
	
	//PT27
	If (DS.State in [dsEdit]) And (GetField('PSS_PGTYPESESSION')='ENT') And (Not IsFieldModified('PSS_PGTYPESESSION')) And (TobModifiedFields.IsOneModifie) Then
	Begin
		If PGIAsk(TraduireMemoire('Voulez-vous mettre à jour les sous-sessions dépendantes?'), Ecran.Caption) = mrYes Then
		Begin
			MAJSousSessions := True;
		End;
	End;
end;

procedure TOM_SESSIONSTAGE.RendDureeStage;
var Nbj : Integer;
    QQ  : TQuery;
    St  : String;
begin
    Nbj  :=  0; // PORTAGECWAS
    if GetField ('PSS_CODESTAGE')  <>  ''  then
    begin
        st  :=  'SELECT PST_JOURSTAGE FROM STAGE WHERE PST_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PST_MILLESIME="0000"';
        QQ := OpenSQL(st,TRUE) ;
        if Not QQ.EOF then Nbj  :=  Trunc(QQ.FindField('PST_JOURSTAGE').AsFloat);
        Ferme (QQ);
    end
    else Nbj  :=  0;
    LaDuree  :=  Nbj;
end;

procedure TOM_SESSIONSTAGE.LanceMailSalarie(Sender : TObject);
var Q:TQuery;
    TobSalaries,T:Tob;
    Destinataire,Sujet,HeureDebut,HeureFin,Fichier:String;
    i:Integer;
    Texte:HTStrings;
    DateDebut,DateFin:TDateTime;
    NomF,Langue,Predef : String;
    Blob      : TRichEdit;
    Ret       : String;
    Req       : String;
    DestCopie : String;
    Contenu   : TStrings; //PT20
begin
	//PT18 - Début
	// Affichage de la liste des salariés à qui envoyer la convocation
	Ret := AGLLanceFiche('PAY', 'SALARIESCONVOC', '', '', 'SALARIES_SESSION;'+GetField('PSS_CODESTAGE')+';'+GetField('PSS_MILLESIME')+';'+IntToStr(GetField('PSS_ORDRE')));

	If Ret = '' Then Exit;
	//PT18 - Fin

    // Création de la liste des destinataires
	Req := 'SELECT DISTINCT(PFO_SALARIE),US_EMAIL,PSE_EMAILPROF FROM FORMATIONS LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE '+ //PT24
		   'LEFT JOIN UTILISAT ON PFO_SALARIE=US_AUXILIAIRE '+ //PT22
           'WHERE PFO_REFUSELE="'+UsDateTime(IDate1900)+'" AND PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" '+
           'AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'"'; //DB2

    // Prise en compte des salariés sélectionnés pour la convocation
    Req := Req + ' AND PFO_SALARIE IN ('+Ret+')';

    Q := OpenSQL(Req,True);

    TobSalaries := Tob.Create('Les mails salaries',Nil,-1);
    TobSalaries.LoadDetailDB('DEPORTSAL','','',Q,False);
    Ferme(Q);

    DateDebut := GetField('PSS_DATEDEBUT');
    DateFin := GetField('PSS_DATEFIN');
    HeureDebut := FormatDateTime('hh"h"mm',GetField('PSS_HEUREDEBUT'));
    HeureFin := FormatDateTime('hh"h"mm',GetField('PSS_HEUREFIN'));
    Destinataire := '';

    For i := 0 To TobSalaries.Detail.Count-1 Do
    Begin
        T := TobSalaries.Detail[i];
        If Destinataire <> '' Then Destinataire := Destinataire+';';
        If T.GetValue('US_EMAIL') <> '' Then //PT22
        	Destinataire := Destinataire+T.GetValue('US_EMAIL')
        Else
        	Destinataire := Destinataire+T.GetValue('PSE_EMAILPROF');
    End;

    //PT18 - Début
    FreeAndNil(TobSalaries);

    DestCopie := '';

    // Création de la liste des membres en copie
    If GetParamSocSecur('SO_PGFORMAILCOPIECONVOC', False) Then
    Begin
		Req := 'SELECT PSE_SALARIE,US_EMAIL,PSE_EMAILPROF FROM DEPORTSAL '+
			   'LEFT JOIN UTILISAT ON PSE_SALARIE=US_AUXILIAIRE WHERE PSE_SALARIE IN ('+ //PT22
	            'SELECT DISTINCT (PSE_RESPONSFOR) FROM DEPORTSAL WHERE PSE_SALARIE IN ('+Ret+'))';
	    TobSalaries := Tob.Create('Les mail responsables',Nil,-1);
	    TobSalaries.LoadDetailFromSQL(Req);
	
	    For i := 0 To TobSalaries.Detail.Count - 1 Do
	    Begin
	        If DestCopie <> '' Then DestCopie := DestCopie+';';
	        If TobSalaries.Detail[i].GetValue('US_EMAIL') <> '' Then //PT22
	        	DestCopie := DestCopie + TobSalaries.Detail[i].GetValue('US_EMAIL')
	        Else
	        	DestCopie := DestCopie+TobSalaries.Detail[i].GetValue('PSE_EMAILPROF');
	    End;
	    
	    If (VH_PAIE.PGForMailAdr <> '') Then DestCopie := DestCopie + ';' + VH_PAIE.PGForMailAdr;
    End;
    //PT18 - Fin

    Sujet := 'Convocation pour la formation '+RechDom('PGSTAGEFORM',GetField('PSS_CODESTAGE'),False);
    Texte :=  HTStringList.Create ;

    Texte.Add('<html>');  //PT20
    Texte.Add('Bonjour,');
    Texte.Add('<br><br>');
    Texte.Add('Nous vous confirmons, par la présente, votre participation à la formation ');
    Texte.Add('<br><h2 align="center">                '+RechDom('PGSTAGEFORM',GetField('PSS_CODESTAGE'),False)+'</h2>');
    // Libellé de la session //PT18
    If GetField('PSS_LIBELLE') <> '' Then Texte.Add('<h3 align="center"><i>'+GetField('PSS_LIBELLE')+'</i></h3>');
    Texte.Add('<br>');

    Texte.Add('<u><b>Date(s) :</b></u>&nbsp;');
    If GetField('PSS_DATEDEBUT')=GetField('PSS_DATEFIN') Then
    begin
        If GetField('PSS_HEUREDEBUT') <> 0 Then   Texte.Add('le '+DateToStr(DateDebut)+' de '+HeureDebut+' à '+HeureFin)
        Else Texte.Add('le '+DateToStr(DateDebut));
    end
    Else
    begin
        If GetField('PSS_HEUREDEBUT') <> 0 Then   Texte.Add('du '+DateToStr(DateDebut)+' à '+HeureDebut+' au '+
        DateToStr(DateFin)+' à '+HeureFin+'.')
        Else Texte.Add('du '+DateToStr(DateDebut)+' '+RechDom('PGDEMIJOURNEE',GetField('PSS_DEBUTDJ'),False)+' au '+
        DateToStr(DateFin)+' '+RechDom('PGDEMIJOURNEE',GetField('PSS_FINDJ'),False)+'.');
    end;

    //PT18 - Début
    // Récupération du ou des animateurs de stage
    Q := OpenSQL('SELECT PAN_LIBELLE FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PAN_ORDRE="'+IntToStr(GetField('PSS_ORDRE'))+'" AND PAN_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);
    If Not Q.EOF Then
    Begin
    	Texte.Add('<br><br>');
    	Texte.Add('<u><b>Animée par :</b></u> '+Q.FindField('PAN_LIBELLE').AsString);
    	If Q.RecordCount > 1 Then
    	Begin
    		Q.Next;
    		While Not Q.EOF Do
    		Begin
    			Texte.Add(' / '+Q.FindField('PAN_LIBELLE').AsString);
    			Q.Next;
    		End;
    	End;
    End;
    Ferme(Q);

    // Récupération du lieu de formation
    If GetField('PSS_LIEUFORM') <> '' Then
    Begin
    	Texte.Add('<br><br>');
	    //Texte.Add('Lieu de formation : '+RechDom('PGLIEUFORMATION',GetField('PSS_LIEUFORM'),False));
	    Q := OpenSQL('SELECT * FROM LIEUFORMATION WHERE PLF_LIEUFORM="'+GetField('PSS_LIEUFORM')+'"', True);
	    If Not Q.EOF Then
	    Begin
	    	Texte.Add('<u><b>Lieu de formation :</b></u> '+Q.FindField('PLF_LIBELLE').AsString);
	    	If Q.FindField('PLF_ADRESSE1').AsString <> '' Then Texte.Add('<br>                    '+Q.FindField('PLF_ADRESSE1').AsString);
	    	If Q.FindField('PLF_ADRESSE2').AsString <> '' Then Texte.Add('<br>                    '+Q.FindField('PLF_ADRESSE2').AsString);
	    	If Q.FindField('PLF_ADRESSE3').AsString <> '' Then Texte.Add('<br>                    '+Q.FindField('PLF_ADRESSE3').AsString);
	    	Texte.Add('<br>                    '+Q.FindField('PLF_CODEPOSTAL').AsString+' '+Q.FindField('PLF_VILLE').AsString);
	    	If Q.FindField('PLF_TELEPHONE').AsString <> '' Then Texte.Add('<br>                    Tél : '+Q.FindField('PLF_TELEPHONE').AsString);
	    	If Q.FindField('PLF_FAX').AsString <> '' Then Texte.Add('<br>                    Fax : '+Q.FindField('PLF_FAX').AsString);
	    End;
	    Ferme(Q);
	End;
	
	// Récupération des compléments
	If GetField('PSS_BLOCNOTE') <> '' Then
    Begin
    	Texte.Add('<br><br>');
    	Texte.Add('<u><b>Informations complémentaires :</u></b>');
        Blob := TRichEdit(GetControl('PSS_BLOCNOTE'));
        Contenu := TStringList.Create(); //PT20
        Contenu.AddStrings(Blob.Lines);
        For i := 0 To Contenu.Count-1 Do Contenu[i] := '<br>'+Contenu[i];
        Texte.AddStrings(Contenu);
    End;
    //PT18 - Fin
    
    //DEBUT PT12
    Fichier := '';
    Q := OpenSQL('SELECT YFS_NOM,YFS_PREDEFINI,YFS_LANGUE FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_CRIT1="MAW" AND YFS_CRIT2="COE" '+
                 ' AND YFS_BCRIT1="X"',True);
    If Not Q.Eof then
    begin
        NomF := Q.FindField('YFS_NOM').AsString;
        Predef := Q.FindField('YFS_PREDEFINI').AsString;
        Langue := Q.FindField('YFS_LANGUE').AsString;
        AGL_YFILESTD_EXTRACT (Fichier, 'PAIE', NomF , 'MAW', 'COE','','','',False,Langue,Predef);
    end;
    Ferme(Q);

    Texte.Add('</html>'); //PT20

    SendMail(Sujet, Destinataire, DestCopie, Texte, Fichier,  V_PGI.MailMethod<>mmOutLook, 1,'','',False,False) ;

	If (Fichier <> '') And FileExists(Fichier) Then DeleteFile(Fichier); //PT21
	    
    //FIN PT12
    
    //PT18 - Début
    // Enregistrement en base de la liste des salariés à qui on a envoyé la convocation
    If TobSalaries <> Nil Then FreeAndNil(TobSalaries);
    
    TobSalaries := TOB.Create('$LesSalariesConvoc', Nil, -1);
    
    Try
	    Q := OpenSQL('SELECT MAX(YMA_MAILID) AS MAXI FROM YMAILS WHERE YMA_UTILISATEUR="'+USER_MAIL_FORMATION+'" AND YMA_FROM="'+GetField('PSS_CODESTAGE')+IntToStr(GetField('PSS_ORDRE'))+'"', True);
	    IF Not Q.EOF Then i := Q.FindField('MAXI').AsInteger + 1 Else i := 0;
	    Ferme(Q);
	Except
		i := 0;
	End;
    
    While Ret <> '' Do
    Begin
    	Destinataire := ReadTokenPipe(Ret, ',');

    	// Extraction du numéro de salarié sans les guillemets
    	If Destinataire <> '' Then
    		Destinataire := Copy (Destinataire, 2, Length(Destinataire)-2)
    	Else Continue;
    	
    	T := TOB.Create('YMAILS', TobSalaries, -1);
    	T.PutValue('YMA_UTILISATEUR', USER_MAIL_FORMATION);
    	T.PutValue('YMA_MAILID',      i);
    	T.PutValue('YMA_EXPEDITEUR',  '');
    	T.PutValue('YMA_STRDATE',     DateToStr(Date));
        T.PutValue('YMA_FROM',        GetField('PSS_CODESTAGE')+IntToStr(GetField('PSS_ORDRE')));
    	T.PutValue('YMA_DEST',        Destinataire);
    	T.PutValue('YMA_SUBJECT',     'SAL|Convocation formation');
    	T.PutValue('YMA_READED',      '-');
    	
    	i := i + 1;
	End;

	Try
		If TobSalaries.Detail.Count > 0 Then TobSalaries.InsertOrUpdateDB;
	Finally
		FreeAndNil(TobSalaries);
	End;
    //PT18 - Fin
    
    Texte.free;
end;

procedure TOM_SESSIONSTAGE.LanceMailResponsables(Sender : TObject);
var
    TobResp,TR,TobSalaries,TS : Tob;
    Destinataire,Sujet,HeureDebut,HeureFin : String;
    Texte : HTStrings;
    Salarie,Resp : String;
    i : Integer;
    DateDebut,DateFin : TDateTime;
    DestCopie : String;
begin
    // Adresses des responsables
    TobResp := Tob.Create('Les responsables',Nil,-1);
    TobResp.LoadDetailFromSQL('SELECT DISTINCT(PFO_RESPONSFOR),PSE_EMAILPROF,US_EMAIL FROM FORMATIONS '+ //PT18
    			 'LEFT JOIN DEPORTSAL ON PFO_RESPONSFOR=PSE_SALARIE '+
    			 'LEFT JOIN UTILISAT ON PFO_SALARIE=US_AUXILIAIRE '+ //PT22 //PT24
                 'WHERE PFO_REFUSELE="'+UsDateTime(IDate1900)+'" AND PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+
                 '" AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND '+
                 'PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PFO_RESPONSFOR<>""');

    // Liste des salariés qui vont participer à la formation
    TobSalaries := Tob.Create('Les mail salaries',Nil,-1);
    TobSalaries.LoadDetailFromSQL('SELECT PFO_NOMSALARIE,PFO_PRENOM,PFO_SALARIE,PFO_RESPONSFOR FROM FORMATIONS '+ //PT18
                 'WHERE PFO_REFUSELE="'+UsDateTime(IDate1900)+'" AND PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+
                 '" AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PFO_RESPONSFOR<>""');

    Sujet := 'Salariés convoqués pour la formation '+RechDom('PGSTAGEFORM',GetField('PSS_CODESTAGE'),False);

    DateDebut := GetField('PSS_DATEDEBUT');
    DateFin := GetField('PSS_DATEFIN');
    HeureDebut := FormatDateTime('hh"h"mm',GetField('PSS_HEUREDEBUT'));
    HeureFin := FormatDateTime('hh"h"mm',GetField('PSS_HEUREFIN'));

    For i := 0 to  TobResp.detail.Count-1 do
    begin
        TR := TobResp.Detail[i];
        Resp := TR.GetValue('PFO_RESPONSFOR'); //PT18
        
        If  TR.GetValue('US_EMAIL') <> '' Then //PT22
        	Destinataire := TR.GetValue('US_EMAIL')
        Else
        	Destinataire := TR.GetValue('PSE_EMAILPROF');

        Texte  :=  HTStringList.Create ;

        Texte.Add('<html>');  //PT20

        //PT18 - Début
        Texte.Add('Bonjour,');
        Texte.Add('<br><br>');
        Texte.Add('Vous trouverez ci-dessous la liste des salariés qui participeront à la session ');

        If GetField('PSS_LIBELLE') <> '' Then
        Texte.Add('<br><h3 align="center"><i>'+GetField('PSS_LIBELLE')+'</i></h3><br>');
        Texte.Add(' qui se déroulera ');
        //PT18 - Fin

        If GetField('PSS_DATEDEBUT')=GetField('PSS_DATEFIN') Then
        begin
            If GetField('PSS_HEUREDEBUT') <> 0 Then
                Texte.Add('du '+DateToStr(DateDebut)+' de '+HeureDebut+' à '+HeureFin)
            Else
                Texte.Add('du '+DateToStr(DateDebut));
        end
        Else
        begin
            If GetField('PSS_HEUREDEBUT') <> 0 Then
                Texte.Add('du '+DateToStr(DateDebut)+' à '+HeureDebut+' au '+ DateToStr(DateFin)+' à '+HeureFin+'.')
            Else
                Texte.Add('du '+DateToStr(DateDebut)+' '+RechDom('PGDEMIJOURNEE',GetField('PSS_DEBUTDJ'),False)+' au '+
                    DateToStr(DateFin)+' '+RechDom('PGDEMIJOURNEE',GetField('PSS_FINDJ'),False)+'.');
        end;
        Texte.Add(':<br><br><ul type="circle">');

        TS := TobSalaries.FindFirst(['PFO_RESPONSFOR'],[Resp],False);
        While TS <> Nil do
        begin
            Salarie := TS.GetValue('PFO_NOMSALARIE')+' '+TS.GetValue('PFO_PRENOM');
            Texte.Add('<li>'+Salarie+'</li>');
            TS := TobSalaries.FindNext(['PFO_RESPONSFOR'],[Resp],False);
        end;
        Texte.Add('</ul>');

        Texte.Add('</html>');  //PT20

		If (VH_PAIE.PGForMailAdr <> '') Then DestCopie := VH_PAIE.PGForMailAdr;
		
        SendMail(Sujet, Destinataire, DestCopie, Texte, '', V_PGI.MailMethod<>mmOutLook, 1, '', '', False, False) ;
        Texte.Free;
    end;
    TobSalaries.Free;
    TobResp.Free;
end;

procedure TOM_SESSIONSTAGE.LanceMailAvertirSession(Sender : TObject);
var 
    TobSalaries,Ts,TobResp,TR : Tob;
    Resp,Salarie,Destinataire,Sujet,LibInsc,HeureDebut,HeureFin : String;
    Texte : HTStrings;
    i : Integer;
    DateDebut,DateFin : TDatetime;
    DestCopie : String;
begin
    // Récupération des responsables
    TobResp := Tob.Create('Les responsables',Nil,-1);
    TobResp.LoadDetailFromSQL('SELECT DISTINCT(PFI_RESPONSFOR),PSE_EMAILPROF,US_EMAIL FROM INSCFORMATION '+ //PT18
    			 'LEFT JOIN DEPORTSAL ON PFI_RESPONSFOR=PSE_SALARIE '+
    			 'LEFT JOIN UTILISAT ON PFI_SALARIE=US_AUXILIAIRE '+ //PT22 //PT24
                 'WHERE PFI_REFUSELE="'+UsDateTime(IDate1900)+'" AND PFI_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND '+
                 'PFI_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PFI_RESPONSFOR<>""');

    // Récupération des salariés inscrits au budget
    TobSalaries := Tob.Create('Les salaries inscrits',Nil,-1);
    TobSalaries.LoadDetailFromSQL('SELECT PFI_RESPONSFOR,PFI_LIBELLE,PFI_ETABLISSEMENT,PFI_SALARIE,PFI_LIBEMPLOIFOR FROM INSCFORMATION '+ //PT18
                 'WHERE PFI_REFUSELE="'+UsDateTime(IDate1900)+'" AND PFI_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND '+
                 'PFI_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PFI_RESPONSFOR<>""');

    Sujet := 'Ouverture d''une session pour la formation '+RechDom('PGSTAGEFORM',GetField('PSS_CODESTAGE'),False);

    DateDebut := GetField('PSS_DATEDEBUT');
    DateFin := GetField('PSS_DATEFIN');
    HeureDebut := FormatDateTime('hh"h"mm',GetField('PSS_HEUREDEBUT'));
    HeureFin := FormatDateTime('hh"h"mm',GetField('PSS_HEUREFIN'));

    For i := 0 to  TobResp.detail.Count-1 do
    begin
        TR := TobResp.Detail[i];
        Resp := TR.GetValue('PFI_RESPONSFOR');
        
        If TR.GetValue('US_EMAIL') <> '' Then //PT22
        	Destinataire := TR.GetValue('US_EMAIL')
        Else
        	Destinataire := TR.GetValue('PSE_EMAILPROF');
        Texte  :=  HTStringList.Create ;

        Texte.Add('<html>');  //PT20

        //PT18
        Texte.Add('Bonjour,<br><br>');

        If GetField('PSS_LIBELLE') <> '' Then
            Texte.Add('Nous vous informons de la création de la session <br><h3 align="center"><i>'+GetField('PSS_LIBELLE')+'</i></h3><br>qui se déroulera')
        Else
            Texte.Add('Nous vous informons de la création d''une session qui se déroulera');
        //PT18 - Fin

        If GetField('PSS_DATEDEBUT')=GetField('PSS_DATEFIN') Then
        begin
            If GetField('PSS_HEUREDEBUT') <> 0 Then
                Texte.Add(' le '+DateToStr(DateDebut)+' de '+HeureDebut+' à '+HeureFin)
            Else
                Texte.Add(' le '+DateToStr(DateDebut));
        end
        Else
        begin
            If GetField('PSS_HEUREDEBUT') <> 0 Then
                Texte.Add(' du '+DateToStr(DateDebut)+' à '+HeureDebut+' au '+DateToStr(DateFin)+' à '+HeureFin+'.')
            Else
                Texte.Add(' du '+DateToStr(DateDebut)+' '+RechDom('PGDEMIJOURNEE',GetField('PSS_DEBUTDJ'),False)+' au '+
                DateToStr(DateFin)+' '+RechDom('PGDEMIJOURNEE',GetField('PSS_FINDJ'),False)+'.');
        end;

		//PT18
        Texte.Add('<br><br>');
        Texte.Add('<b><u>Rappel de vos inscriptions au budget :</u></b><ul type="circle">');
        TS := TobSalaries.FindFirst(['PFI_RESPONSFOR'],[Resp],False);
        While TS <> Nil do
        begin
            Salarie := TS.GetValue('PFI_SALARIE');
            If Salarie <> '' Then LibInsc := TS.GetValue('PFI_LIBELLE')
            Else libInsc := TS.GetValue('PFI_LIBELLE')+' , '+TS.GetValue('PFI_ETABLISSEMENT')+' , '+TS.GetValue('PFI_LIBEMPLOIFOR');
            Texte.Add('<li>'+LibInsc+'</li>');
            TS := TobSalaries.FindNext(['PFI_RESPONSFOR'],[Resp],False);
        end;
        Texte.Add('</ul>');
        Texte.Add('</html>');  //PT20

		If (VH_PAIE.PGForMailAdr <> '') Then DestCopie := VH_PAIE.PGForMailAdr;
		
        SendMail(Sujet, Destinataire, DestCopie, Texte, '', V_PGI.MailMethod<>mmOutLook, 1, '', '', False, False) ;
        Texte.Free;
    end;
    TobSalaries.Free;
    TobResp.Free;
end;

procedure TOM_SESSIONSTAGE.ConvocSession(Sender : TObject);
begin
    AglLanceFiche ('PAY','SALARIE_RTF', '', '' , 'FORMATION;'+GetField('PSS_CODESTAGE')+';'+IntToStr(GetField('PSS_ORDRE'))+';'+GetField('PSS_MILLESIME'));
end;


procedure TOM_SESSIONSTAGE.CreerNouveauStage(Sender : TObject);
var NouveauStage : String;
begin
    NouveauStage := AglLanceFiche('PAY','STAGE','','','SAISIECAT;ACTION=CREATION');
    If NouveauStage <> '' then SetField('PSS_CODESTAGE',NouveauStage);
    AvertirTable('PGSTAGEFORM');
end;

procedure TOM_SESSIONSTAGE.SaisirFeuillePresence(Sender : TObject);
var St : String;
begin
    if (DS.State in [dsEdit]) then
    begin
        PGIBox('Vous avez des modifications en cours veuillez valider ou les annuler avant la saisie de la feuille de présence',Ecran.Caption);
        Exit;
    end;
    st := GetField('PSS_CODESTAGE')+';'+IntToStr(GetField('PSS_ORDRE'))+';'+GetField('PSS_MILLESIME');
    AglLanceFiche('PAY','PRESENCEFORM','','','SAISIEFP;'+St+';ACTION=CREATION');
    RefreshDB;
end;

procedure TOM_SESSIONSTAGE.OnGICellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
    ForceUpdate;
    If ACol = 0 then
    begin
        If (GInvest.CellValues[0,ARow] <> '') AND (GInvest.CellValues[0,ARow] <> 'AUC') then     //PT6
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

procedure TOM_SESSIONSTAGE.MajInvest;
var Q:Tquery;
    TobInvest,TI : tob;
    i : Integer;
begin
    ExecuteSQL('DELETE FROM INVESTSESSION WHERE PIS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PIS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PIS_MILLESIME="'+GetField('PSS_MILLESIME')+'"');  //DB2
    Q := OpenSQL('SELECT * FROM INVESTSESSION '+
                 'WHERE PIS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PIS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PIS_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);  //DB2
    TobInvest := Tob.Create('INVESTSESSION',Nil,-1);
    TobInvest.LoadDetailDB('INVESTSESSION','','',Q,False);
    Ferme(Q);
    For i := 1 to 2 do
    begin
        If (GInvest.CellValues[0,i] <> '') AND (GInvest.CellValues[0,i] <> 'AUC') then
        begin
            TI := Tob.Create('INVESTSESSION',TobInvest,-1);
            TI.PutValue('PIS_CODESTAGE',GetField('PSS_CODESTAGE'));
            TI.PutValue('PIS_ORDRE',GetField('PSS_ORDRE'));
            TI.PutValue('PIS_MILLESIME',GetField('PSS_MILLESIME'));
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

procedure TOM_SESSIONSTAGE.GestionDesDif;
var TobStagiaire,TobMvtDIF : Tob;
    i,a : Integer;
    Q : TQuery;
    Salarie,etab,FJ,DJ,TypeFor : String;
    DateD,DateF : TDateTime;
    AncDateD,AncDateF : TDateTime;
    NbhT,NbHNonT : Double;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim : String;
    HeuresTAlim,HeuresNonTAlim,Jours : Double;
    AbsOk : Integer;
    Predef : String;
    DatePaie : TDateTime;
    SuppAbsence,SuppRubrique : Boolean;
begin
    DateD := GetField('PSS_DATEDEBUT');
    DateF := GetField('PSS_DATEFIN');
    DJ := GetField('PSS_DEBUTDJ');
    FJ := GetField('PSS_FINDJ');
    Jours := GetField('PSS_JOURSTAGE');
    Q := OpenSQL('SELECT * FROM FORMATIONS WHERE '+
                 'PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND '+
                 'PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND '+
                 'PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND '+
                 'PFO_EFFECTUE="X"',True);
    TobStagiaire :=Tob.Create('LesDif',Nil,-1);
    TobStagiaire.LoadDetailDB('LesDif','','',Q,False);
    ferme(Q);
    For i := 0 to TobStagiaire.Detail.Count - 1 do
    begin
        Salarie := TobStagiaire.Detail[i].GetValue('PFO_SALARIE');
        Etab := TobStagiaire.Detail[i].GetValue('PFO_ETABLISSEMENT');
        NbhT := TobStagiaire.Detail[i].GetValue('PFO_HTPSTRAV');
        NbhNonT := TobStagiaire.Detail[i].GetValue('PFO_HTPSNONTRAV');
        TypeFor := TobStagiaire.Detail[i].GetValue('PFO_TYPEPLANPREV');
        AncDateD := TobStagiaire.Detail[i].GetValue('PFO_DATEDEBUT');
        AncDateF := TobStagiaire.Detail[i].GetValue('PFO_DATEFIN');
        DatePaie := TobStagiaire.Detail[i].GetValue('PFO_DATEPAIE');
        SupprimerDIF(Salarie,TypeFor,AncDateD,AncDateF,DatePaie,SuppAbsence,SuppRubrique);
        Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI## PPF_TYPEPLANPREV="'+TypeFor+'"',True);
        TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
        TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
        Ferme(Q);
        For a := 0 TO TobMvtDIF.Detail.Count - 1 do
        begin
            Travail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURETRAV');
            HorsTravail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURENONTRAV');
            Predef := TobMvtDIF.Detail[a].GetValue('PPF_PREDEFINI');
            If Predef <> 'DOS' then
            begin
                If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['DOS',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
                If Predef = 'CEG' then
                begin
                    If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['STD',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
                end;
            end;
            HeuresTAlim := 0;
            HeuresNonTAlim := 0;
            If Travail = 'X' then HeuresTAlim := NbhT;
            If HorsTravail = 'X' then HeuresNonTAlim := NbhNonT;
            If (TobMvtDIF.Detail[a].GetValue('PPF_ALIMABS') = 'X') and (SuppAbsence) and (SuppRubrique) then
            begin
                TypeConge := TobMvtDIF.Detail[a].GetValue('PPF_TYPECONGE');
                AbsOk := PGDIFGenereAbs(HeuresTAlim,HeuresNonTAlim,Jours,Salarie,DJ,FJ,TypeConge,Etab,DateD,DateF);
                If AbsOk = -1 then
                begin
                    PGIBox('Impossible de créer l''absence correspondante car le salarié possède déja une absence sur cette période',Ecran.Caption);
                    Exit;
                end;
            end;
            If (TobMvtDIF.Detail[a].GetValue('PPF_ALIMRUB') = 'X') and (SuppRubrique) and (SuppAbsence) then
            begin
                Rubrique := TobMvtDIF.Detail[a].GetValue('PPF_RUBRIQUE');
                Travail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURETRAV');
                HorsTravail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURENONTRAV');
                TypeAlim := TobMvtDIF.Detail[a].GetValue('PPF_ALIMENT');
                RecupereLesFormationsDIF(Salarie,TypeFor,Rubrique,TypeAlim,DatePaie,Travail = 'X',HorsTravail = 'X');
                //                    PGDIFGenereRub(HeuresTAlim,HeuresNonTAlim,Salarie,Etab,Rubrique,TypeAlim,GetField('PSS_CODESTAGE'),GetField('PFO_CONFIDENTIEL'),DateD,DateF,DatePaie);
            end;
        end;
        TobMvtDIF.Free;
    end;
    TobStagiaire.Free;
end;

procedure TOM_SESSIONSTAGE.SupprimerDIF(Salarie,TypeFor : String;DateD,DateF,DatePaie : TDateTime; var SuppAbsence,SuppRubrique : Boolean);
var Q : TQuery;
    TobMvtDIF : Tob;
    i : Integer;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim,Predef : String;
    DateDebRub,DateFinRub : TDateTime;
begin
    SuppAbsence := True;
    SuppRubrique := True;
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
                         'PCN_SALARIE="'+Salarie+'" AND '+
                         '(PCN_CODETAPE="P" OR PCN_CODETAPE="S") AND '+
                         'PCN_DATEDEBUTABS="'+UsDateTime(InitDDSession)+'" AND '+
                         'PCN_DATEFINABS="'+UsDateTime(InitDFSession)+'"') then
            begin
                PGIBox('L''absence pour le salarié '+RechDom('PGSALARIE',Salarie,False)+
                ' du '+DateToStr(InitDDSession)+' au '+DateToStr(InitDFSession)+
                '#13#10 a déja été intégré dans le bulletin, aucune modification ne sera effectué','Génération absence');
                SuppAbsence := False;
            end
            else PGDIFSuppAbs(Salarie,TypeConge,InitDDSession,InitDFSession);
        end;
        If TobMvtDIF.Detail[i].GetValue('PPF_ALIMRUB') = 'X' then
        begin
            Rubrique := TobMvtDIF.Detail[i].GetValue('PPF_RUBRIQUE');
            TypeAlim := TobMvtDIF.Detail[i].GetValue('PPF_ALIMENT');
            DateDebRub := DebutDeMois(DatePaie);
            DateFinRub := FinDeMois(DatePaie);
            If ExisteSQL('SELECT PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE '+
                         'PPU_DATEDEBUT>="'+UsDateTime(DateDebRub)+'" AND '+
                         'PPU_DATEFIN<="'+UsDateTime(dateFinRub)+'" AND '+
                         'PPU_SALARIE="'+Salarie+'"') then
            begin
                SuppRubrique := False;
                PGIBox('La bulletin pour le salarié '+RechDom('PGSALARIE',GetField('PFO_SALARIE'),False)+
                       ' du '+DateToStr(dateDebRub)+' au '+DateToStr(dateFinRub)+
                       '#13#10 a déja été créé, aucune modification ne sera effectuée pour les rubriques','Génération absence');
            end
            else
            PGDIFSuppRub(Salarie,Rubrique,TypeAlim,DatePaie);
        end;
    end;
    TobMvtDIF.free;
end;

procedure TOM_SESSIONSTAGE.EditionSession(Sender: TObject);
var Where : String;
{$IFDEF EAGLCLIENT}
var StPages : String;
    Pages : TPageControl;
{$ENDIF}
begin
    Where := 'PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'"';
    {$IFNDEF EAGLCLIENT}
    LanceEtat('E', 'PFO', 'PSS', True, False, False, nil, Where, '', False);
    {$ELSE}
    Pages := TPageControl(GetControl('PAGES'));
    StPages := AglGetCriteres (Pages, FALSE);
    LanceEtat('E', 'PFO', 'PSS', True, False, False, nil, Where, '', False,0,StPages);
    {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Attention : Si modification de cette fonction, modifier également 
Suite ........ : TOF_PGFEUILLEPRESENCEFOR.EditerFeuillePresence
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.EditionPresenceSession(Sender : TObject); //PT8
var Q : TQuery;
    a,i,Ordre,j : Integer;
    TobAnim,TobSta,TobSession,TobEtat,T : Tob;
    Where,Stage,Libelle,Libelle1,Lieu : String;
    Pages : TPageControl;
    NbJours,NbHeures : Double;
    NbJourInt : Integer;
    DateDebut,DateFin : TDateTime;
    WherePredef : String;  //PT26
begin
	TobEtat := Tob.Create('Edition',Nil,-1);
	Pages := TPageControl(GetControl('PAGES'));
    Where := 'WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'"';
	WherePredef := '';
	If PGBundleInscFormation Then WherePredef := WherePredef + DossiersAInterroger('',V_PGI.NoDossier,'PSS',True,True);  //PT26

	Q := OpenSQL('SELECT * FROM SESSIONSTAGE LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME '+Where+WherePredef,True);  //PT26
	TobSession := Tob.Create('LesSessions',Nil,-1);
	TobSession.LoadDetailDB('LesSessions','','',Q,False);
	Ferme(Q);

	For i := 0 To TobSession.Detail.Count - 1 do
	begin
		Stage := TobSession.Detail[i].GetValue('PSS_CODESTAGE');
		Ordre := TobSession.Detail[i].GetValue('PSS_ORDRE');
		Libelle := TobSession.Detail[i].GetValue('PST_LIBELLE');
		Libelle1 := TobSession.Detail[i].GetValue('PST_LIBELLE1');
		NbJours := TobSession.Detail[i].GetValue('PSS_JOURSTAGE');
		NbJourInt := TobSession.Detail[i].GetInteger('PSS_JOURSTAGE');
		NbHeures := TobSession.Detail[i].GetValue('PSS_DUREESTAGE');
		DateDebut := TobSession.Detail[i].GetValue('PSS_DATEDEBUT');
		DateFin := TobSession.Detail[i].GetValue('PSS_DATEFIN');
		Lieu := TobSession.Detail[i].GetValue('PSS_LIEUFORM');

        //PT26
		// Récupération des animateurs
		If PGBundleHierarchie Then
		Q := OpenSQL('SELECT PAN_SALARIE,PAN_LIBELLE,PSI_LIBELLE,PSI_PRENOM FROM SESSIONANIMAT '+
		'LEFT JOIN INTERIMAIRES ON PAN_SALARIE=PSI_INTERIMAIRE '+
		'WHERE PAN_CODESTAGE="'+Stage+'" AND PAN_ORDRE='+IntToStr(Ordre),True)
		Else
		Q := OpenSQL('SELECT PAN_SALARIE,PAN_LIBELLE,PSA_LIBELLE,PSA_PRENOM FROM SESSIONANIMAT '+
		'LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE '+
		'WHERE PAN_CODESTAGE="'+Stage+'" AND PAN_ORDRE='+IntToStr(Ordre),True);

		TobAnim := Tob.Create('LesAnimateurs',Nil,-1);
		TobAnim.LoadDetailDB('LesAnimateurs','','',Q,False);
		Ferme(Q);

		For a := 0 To TobAnim.Detail.Count - 1 do
		begin
			For j := 1 to NbJourInt do
			begin
				T := Tob.Create('Fille',TobEtat,-1);
				T.AddChampSupValeur('STAGE',Stage);
				T.AddChampSupValeur('ORDRE',Ordre);
				T.AddChampSupValeur('NUMJOUR',j);
				T.AddChampSupValeur('SALARIE',TobAnim.Detail[a].GetValue('PAN_SALARIE'));
                //PT26
				If PGBundleHierarchie Then
				Begin
					If TobAnim.Detail[a].GetValue('PAN_SALARIE') <> '' Then
					Begin
						T.AddChampSupValeur('NOM',    TobAnim.Detail[a].GetValue('PSI_LIBELLE'));
						T.AddChampSupValeur('PRENOM', TobAnim.Detail[a].GetValue('PSI_PRENOM'));
					End
					Else
					Begin
						T.AddChampSupValeur('NOM',    TobAnim.Detail[a].GetValue('PAN_LIBELLE'));
						T.AddChampSupValeur('PRENOM', '');
					End;
				End
				Else
				Begin
					If TobAnim.Detail[a].GetValue('PAN_SALARIE') <> '' Then
					Begin
						T.AddChampSupValeur('NOM',    TobAnim.Detail[a].GetValue('PSA_LIBELLE'));
						T.AddChampSupValeur('PRENOM', TobAnim.Detail[a].GetValue('PSA_PRENOM'));
					End
					Else
					Begin
						T.AddChampSupValeur('NOM',    TobAnim.Detail[a].GetValue('PAN_LIBELLE'));
						T.AddChampSupValeur('PRENOM', '');
					End;
				End;
				T.AddChampSupValeur('TYPE','ANIM');
				T.AddChampSupValeur('LIBELLE',Libelle);
				T.AddChampSupValeur('LIBELLE1',Libelle1);
				T.AddChampSupValeur('NBJOURS',NbJours);
				T.AddChampSupValeur('NBHEURES',NbHeures);
				T.AddChampSupValeur('DATEDEBUT',DateDebut);
				T.AddChampSupValeur('DATEFIN',DateFin);
				T.AddChampSupValeur('PSS_LIEUFORM',Lieu);
			end;
		end;
		TobAnim.Free;

        //PT26
		// Récupération des stagiaires
		If PGBundleHierarchie Then
			Q := OpenSQL('SELECT PFO_SALARIE,PSI_LIBELLE,PSI_PRENOM FROM FORMATIONS '+
			'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFO_SALARIE '+
			'WHERE PFO_CODESTAGE="'+Stage+'" AND PFO_ORDRE='+IntToStr(Ordre),True)
		Else
			Q := OpenSQL('SELECT PFO_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM FORMATIONS '+
			'LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE '+
			'WHERE PFO_CODESTAGE="'+Stage+'" AND PFO_ORDRE='+IntToStr(Ordre),True);
		TobSta := Tob.Create('LesStagiaires',Nil,-1);
		TobSta.LoadDetailDB('LesStagiaires','','',Q,False);
		Ferme(Q);
		For a := 0 To TobSta.Detail.Count - 1 do
		begin
			For j := 1 to NbJourInt do
			begin
				T := Tob.Create('Fille',TobEtat,-1);
				T.AddChampSupValeur('STAGE',Stage);
				T.AddChampSupValeur('ORDRE',Ordre);
				T.AddChampSupValeur('NUMJOUR',j);
				T.AddChampSupValeur('SALARIE',TobSta.Detail[a].GetValue('PFO_SALARIE'));
                //PT26
				If PGBundleHierarchie Then 
				Begin
					T.AddChampSupValeur('NOM',TobSta.Detail[a].GetValue('PSI_LIBELLE'));
					T.AddChampSupValeur('PRENOM',TobSta.Detail[a].GetValue('PSI_PRENOM'));
				End
				Else
				Begin
					T.AddChampSupValeur('NOM',TobSta.Detail[a].GetValue('PSA_LIBELLE'));
					T.AddChampSupValeur('PRENOM',TobSta.Detail[a].GetValue('PSA_PRENOM'));
				End;
				T.AddChampSupValeur('TYPE','STA');
				T.AddChampSupValeur('LIBELLE',Libelle);
				T.AddChampSupValeur('LIBELLE1',Libelle1);
				T.AddChampSupValeur('NBJOURS',NbJours);
				T.AddChampSupValeur('NBHEURES',NbHeures);
				T.AddChampSupValeur('DATEDEBUT',DateDebut);
				T.AddChampSupValeur('DATEFIN',DateFin);
				T.AddChampSupValeur('PSS_LIEUFORM',Lieu);
			end;
		end;
		TobSta.Free;
	end;
	TobSession.Free;
	TobEtat.Detail.Sort('STAGE;ORDRE;NUMJOUR;TYPE;NOM');
    LanceEtatTOB('E','PFO','PFP',TobEtat,True,False,False,Pages,'','',False);
    TobEtat.Free;
end;

Procedure TOM_SESSIONSTAGE.BSousSessionClick (Sender : TObject);
var Action : String;
begin
	If TFFiche(Ecran).FTypeAction = TaConsult Then Action := 'CONSULTATION' Else Action := 'MODIFICATION';
    AGLLanceFiche('PAY', 'MULSOUSSESSION', '', '', GetField('PSS_CODESTAGE')+';'+GetField('PSS_MILLESIME')+';'+IntToStr(GetField('PSS_ORDRE'))+';'+Action);

    If ExisteSQL('SELECT 1 FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'"') Then
    Begin
	    SetControlEnabled('PSS_PGTYPESESSION', False);
	   	SetControlVisible('BMULTISALSS',       True);
	   	SetControlVisible('BRECUPERATIONSS',   True);
        SetControlEnabled('BMULTISALSS',       True);
        SetControlEnabled('BRECUPERATIONSS',   True);
    End
    Else
    Begin
	    SetControlEnabled('PSS_PGTYPESESSION', True);
        SetControlEnabled('BMULTISALSS',       False);
        SetControlEnabled('BRECUPERATIONSS',   False);
    End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 10/04/2007 / PT9
Modifié le ... :   /  /
Description .. : Duplication d'une session de formation
Mots clefs ... : DUPLICATION;SESSION
*****************************************************************}
procedure TOM_SESSIONSTAGE.DupliquerSession(Sender: Tobject);
begin
    DuplSession   := GetField('PSS_CODESTAGE');
    If GetField('PSS_PGTYPESESSION') = 'SOS' Then //PT27
    	DuplOrdre := GetField('PSS_NUMSESSION')
    Else
    	DuplOrdre := GetField('PSS_ORDRE');
    DuplMillesime := GetField('PSS_MILLESIME');
    ADupliquer    := true;
    
    //PT27
    If GetField('PSS_PGTYPESESSION') = 'SOS' Then
    Begin
    	LeStage := GetField('PSS_CODESTAGE');
    	Millesime := GetField('PSS_MILLESIME');
    	NumSessionMaitre := GetField('PSS_NUMSESSION');
    End;
    
    TFFiche(Ecran).BinsertClick(TFFiche(Ecran).BInsert);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 10/04/2007 / PT9
Modifié le ... :   /  /
Description .. : Récupération des informations d'une session afin de la dupliquer
                 N.B. : Les dates de début et de fin ne sont pas récupérées ainsi que les tops de clôture
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.RecupInfosSession;
var Q : TQuery;
begin
    Q := OpenSQL('SELECT * FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+DuplSession+'" AND PSS_ORDRE="'+DuplOrdre+'" AND PSS_MILLESIME="'+DuplMillesime+'"',True);
    if not Q.eof then
    begin
        SetField('PSS_CODESTAGE',Q.FindField('PSS_CODESTAGE').AsString);

		If TypeSaisie <> 'SOUSSESSION' Then  //PT27
		Begin
        	SetField('PSS_DEBUTDJ',Q.FindField('PSS_DEBUTDJ').AsString);
        	SetField('PSS_FINDJ',Q.FindField('PSS_FINDJ').AsString);
        	SetField('PSS_DUREESTAGE',Q.FindField('PSS_DUREESTAGE').AsString);
        	SetField('PSS_JOURSTAGE',Q.FindField('PSS_JOURSTAGE').AsString);
        	SetField('PSS_HEUREDEBUT',Q.FindField('PSS_HEUREDEBUT').AsString);
        	SetField('PSS_HEUREFIN',Q.FindField('PSS_HEUREFIN').AsString);
        	//PT15 - Début
        	SetControlText('HEUREDEBUT', FormatDateTime('hh:mm', GetField('PSS_HEUREDEBUT')));
        	SetControlText('HEUREFIN',   FormatDateTime('hh:mm', GetField('PSS_HEUREDEBUT')));
        	//PT15 - Fin
        End;

        SetField('PSS_CENTREFORMGU',Q.FindField('PSS_CENTREFORMGU').AsString);
        SetField('PSS_LIEUFORM',Q.FindField('PSS_LIEUFORM').AsString);
        If TypeSaisie <> 'SOUSSESSION' Then SetField('PSS_LIBELLE',Q.FindField('PSS_LIBELLE').AsString); //PT27
        SetField('PSS_INCLUSDECL',Q.FindField('PSS_INCLUSDECL').AsString);
        SetField('PSS_COMPTERENDU',Q.FindField('PSS_COMPTERENDU').AsString);
        SetField('PSS_ATTESTPRESENC',Q.FindField('PSS_ATTESTPRESENC').AsString);
        SetField('PSS_QUESTAPPREC',Q.FindField('PSS_QUESTAPPREC').AsString);
        SetField('PSS_QUESTEVALUAT',Q.FindField('PSS_QUESTEVALUAT').AsString);
        SetField('PSS_SUPPCOURS',Q.FindField('PSS_SUPPCOURS').AsString);
        SetField('PSS_VIDEOPROJ',Q.FindField('PSS_VIDEOPROJ').AsString);
        SetField('PSS_RETROPROJ',Q.FindField('PSS_RETROPROJ').AsString);
        SetField('PSS_JEUXROLE',Q.FindField('PSS_JEUXROLE').AsString);
        SetField('PSS_ETUDECAS',Q.FindField('PSS_ETUDECAS').AsString);
        SetField('PSS_AUTREMOYEN',Q.FindField('PSS_AUTREMOYEN').AsString);
        SetField('PSS_NATUREFORM',Q.FindField('PSS_NATUREFORM').AsString);
        SetField('PSS_ACTIONFORM',Q.FindField('PSS_ACTIONFORM').AsString);
        SetField('PSS_TYPCONVFORM',Q.FindField('PSS_TYPCONVFORM').AsString);
        SetField('PSS_NBSTAMIN',Q.FindField('PSS_NBSTAMIN').AsString);
        SetField('PSS_NBRESTAGPREV',Q.FindField('PSS_NBRESTAGPREV').AsString);
        SetField('PSS_AUTREEVALUAT',Q.FindField('PSS_AUTREEVALUAT').AsString);
        SetField('PSS_VOIRPORTAIL',Q.FindField('PSS_VOIRPORTAIL').AsString);
        SetField('PSS_VALIDORG',Q.FindField('PSS_VALIDORG').AsString); 
        SetField('PSS_PROGRAMME',Q.FindField('PSS_PROGRAMME').AsString);

        SetField('PSS_COUTPEDAG',Q.FindField('PSS_COUTPEDAG').AsString);
        SetField('PSS_COUTSALAIR',Q.FindField('PSS_COUTSALAIR').AsString);
        SetField('PSS_COUTBUDGETE',Q.FindField('PSS_COUTBUDGETE').AsString);
        SetField('PSS_COUTMOB',Q.FindField('PSS_COUTMOB').AsString);
        SetField('PSS_COUTFONCT',Q.FindField('PSS_COUTFONCT').AsString);
        SetField('PSS_COUTORGAN',Q.FindField('PSS_COUTORGAN').AsString);
        SetField('PSS_COUTUNITAIRE',Q.FindField('PSS_COUTUNITAIRE').AsString);
        SetField('PSS_COUTEVAL',Q.FindField('PSS_COUTEVAL').AsString);
        SetField('PSS_TOTALHT',Q.FindField('PSS_TOTALHT').AsString);

        SetField('PSS_NUMOCP',Q.FindField('PSS_NUMOCP').AsString);
        SetField('PSS_ORGCOLLECTPGU',Q.FindField('PSS_ORGCOLLECTPGU').AsString);
        SetField('PSS_ORGCOLLECTSGU',Q.FindField('PSS_ORGCOLLECTSGU').AsString);
        SetField('PSS_NUMOCS',Q.FindField('PSS_NUMOCS').AsString);
        SetField('PSS_NUMFACTURE',Q.FindField('PSS_NUMFACTURE').AsString);

        SetField('PSS_FORMATION1',Q.FindField('PSS_FORMATION1').AsString);
        SetField('PSS_FORMATION2',Q.FindField('PSS_FORMATION2').AsString);
        SetField('PSS_FORMATION3',Q.FindField('PSS_FORMATION3').AsString);
        SetField('PSS_FORMATION4',Q.FindField('PSS_FORMATION4').AsString);
        SetField('PSS_FORMATION5',Q.FindField('PSS_FORMATION5').AsString);
        SetField('PSS_FORMATION6',Q.FindField('PSS_FORMATION6').AsString);
        SetField('PSS_FORMATION7',Q.FindField('PSS_FORMATION7').AsString);
        SetField('PSS_FORMATION8',Q.FindField('PSS_FORMATION8').AsString);

        SetField('PSS_ENVMAILSAL',Q.FindField('PSS_ENVMAILSAL').AsString);
        SetField('PSS_ENVMAILRESP',Q.FindField('PSS_ENVMAILRESP').AsString);
        SetField('PSS_ENVMAILAVER',Q.FindField('PSS_ENVMAILAVER').AsString);
        SetField('PSS_EDITCONVOC',Q.FindField('PSS_EDITCONVOC').AsString);

        SetField('PSS_AVECCLIENT',Q.FindField('PSS_AVECCLIENT').AsString);
        SetField('PSS_REGLEMENTAIRE',Q.FindField('PSS_REGLEMENTAIRE').AsString);
        SetField('PSS_NUMENVOI',Q.FindField('PSS_NUMENVOI').AsString);

        SetField('PSS_TYPESESSSTAGE',Q.FindField('PSS_TYPESESSSTAGE').AsString);
        SetField('PSS_PREDEFINI',Q.FindField('PSS_PREDEFINI').AsString);
        SetField('PSS_NODOSSIER',Q.FindField('PSS_NODOSSIER').AsString);
        If TypeSaisie <> 'SOUSSESSION' Then  //PT27
        Begin
        	SetField('PSS_IDSESSIONFOR',Q.FindField('PSS_IDSESSIONFOR').AsString);
        	SetField('PSS_NOSESSIONMULTI',Q.FindField('PSS_NOSESSIONMULTI').AsString);
        	SetField('PSS_PGTYPESESSION',Q.FindField('PSS_PGTYPESESSION').AsString);
        End;
        SetField('PSS_BLOCNOTE',Q.FindField('PSS_BLOCNOTE').AsString);
    end;
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Affiche ou cache les boutons en fonction du type de session
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.AccesBoutonsSousSession;
begin
    AccesBoutons (Not (GetField('PSS_PGTYPESESSION') = 'ENT'));
    SetControlEnabled('BSOUSSESSION',(GetField('PSS_PGTYPESESSION') = 'ENT'));
    If GetField('PSS_PGTYPESESSION') = 'ENT' Then //PT27
    Begin
	   	SetControlVisible('BFORMATION',      False);
	   	SetControlVisible('BPRESENCE',       False);
	   	If TFFiche(Ecran).Name = 'SESSIONSIMPLE' Then TTabSheet(GetControl('PINSCRIPTIONS')).TabVisible := False;
	End
	Else If GetField('PSS_PGTYPESESSION') = 'AUC' Then
	Begin
	   	SetControlVisible('BFORMATION',      True);
	   	SetControlVisible('BPRESENCE',       True);
	   	SetControlVisible('BMULTISALSS',     False);
	   	SetControlVisible('BRECUPERATIONSS', False);
	End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Affiche ou cache les boutons de gestion des éléments dépendants de la session
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.AccesBoutons (State : Boolean);
begin
    SetControlEnabled('BANIMAT',        State);
    SetControlEnabled('BFORMATION',     State);
    SetControlEnabled('BPRESENCE',      State);
    SetControlEnabled('BMAIL',          State);
    SetControlEnabled('BCONVOC',        State);
    SetControlEnabled('BIMPPRESENCE',   State);
    If TFFiche(Ecran).Name = 'SESSIONSIMPLE' Then 
    	TTabSheet(GetControl('PINSCRIPTIONS')).TabVisible := State;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/11/2007 / PT13
Modifié le ... :   /  /
Description .. : Mise en forme de la grille d'investissement
Suite ........ : et chargement des données associées au stage
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnEnterOngletCouts(Sender: TObject);
var i : Integer;
    Q : TQuery;
    TobInvest : Tob;
begin
    If (Not GrilleChargee) And (GInvest <> Nil) Then
    Begin
        If ExisteSQL('SELECT PIS_INVESTFORM FROM INVESTSESSION WHERE PIS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'"'+
                     ' AND PIS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+
                     ' AND (PIS_INVESTFORM="PDF" OR PIS_INVESTFORM="ALT")') then
            GInvest.ColFormats[0] := 'CB=PGFINANCEFORMATION'
        Else
            GInvest.ColFormats[0] := 'CB=PGTYPEPLANPREV';
        GInvest.ColFormats[1] := 'CB=PGOUINON';
        GInvest.ColFormats[2] := '# ##0.00';
        GInvest.ColFormats[3] := 'CB=PGOUINON';
        GInvest.ColFormats[4] := '# ##0.00';
        GInvest.ColFormats[5] := 'CB=PGOUINON';
        GInvest.ColFormats[6] := '# ##0.00';
        GInvest.ColFormats[7] := '# ##0';
        GInvest.OnCellExit    := OnGICellExit;

        Q := OpenSQL('SELECT PIS_INVESTFORM,PIS_COUTPEDAG,PIS_COUTSALAIRE,PIS_FRAISFORM,PIS_TAUXPEDAG,PIS_TAUXSALAIRE,PIS_TAUXSALFRAIS,PIS_MONTANT FROM INVESTSESSION '+
                     'WHERE PIS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PIS_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PIS_MILLESIME="'+GetField('PSS_MILLESIME')+'"',True);  //DB2
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
            GInvest.CellValues[7,i+1] :=  FloatToStr(TobInvest.Detail[i].GetValue('PIS_MONTANT'));
        end;
        TobInvest.Free;

        GrilleChargee := True;
    End;
end;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008
Modifié le ... :   /  /    
Description .. : Répercute un changement d'heure dans le champ basé
Mots clefs ... : 
*****************************************************************}
Procedure TOM_SESSIONSTAGE.OnChangeHeure (Sender : TObject);
Begin
        SetField('PSS_'+THEdit(Sender).Name, StrToTime(THEdit(Sender).Text));
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Entrée sur l'onglet des inscriptions en SESSIONSIMPLE
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnEnterOngletInscriptions(Sender: TObject);
var TempoBool : Boolean;
    Row,Col   : Integer;
begin
    // Déclenchement manuel des évènements nécessaires à l'initialisation des données de la grille
    OnGSRowEnter (Sender,1,TempoBool,False);
    Col := 0; Row := 1;
    OnGSCellEnter(Sender,Col,Row,TempoBool);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Sortie de l'onglet des inscriptions en SESSIONSIMPLE
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnExitOngletInscriptions(Sender: TObject);
var tempoBool : Boolean;
begin
    GrilleSal.OnRowExit(Sender, GrilleSal.Row,tempoBool,False);
    NewLine := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Clic sur le bouton "Récupération du prévisionnel"
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.BRecupPrevClick (Sender : TObject);
var i,j              : Integer;
    TobSousSessions  : TOB;
begin
	AGLLanceFiche('PAY','MUL_INSCFOR','','','RECUPPREVISIONNEL;'+GetField('PSS_CODESTAGE')+';'+GetField('PSS_MILLESIME'));
	If PGTobRecupPrev <> Nil then
	begin
		//PT27
		// Chargement des sous-sessions
		If GetField('PSS_PGTYPESESSION') = 'ENT' Then
		Begin
	  		TobSousSessions := TOB.Create('LesSousSessions', Nil, -1);
	  		TobSousSessions.LoadDetailDBFromSQL('LesSousSessions','SELECT PSS_ORDRE,PSS_NOSESSIONMULTI FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'"');
	  	End;
  	
		For i := 0 to PGTobRecupPrev.Detail.Count - 1 do
		begin
			If GetField('PSS_PGTYPESESSION') <> 'ENT' Then
				{ Appel de la récupération du prévisionnel avec False pour ne pas rechercher les salaires et impacter
				les jours de formation }
				InscriptionFormation(PGTobRecupPrev.Detail[i].GetValue('SALARIE'),PGTobRecupPrev.Detail[i].GetValue('TYPEPLAN'),'', GetField('PSS_EFFECTUE')='X')
			Else
			Begin //PT27
				For j := 0 To TobSousSessions.Detail.Count - 1 Do
                Begin
					If Not InscriptionFormation(PGTobRecupPrev.Detail[i].GetValue('SALARIE'),PGTobRecupPrev.Detail[i].GetValue('TYPEPLAN'),'', GetField('PSS_EFFECTUE')='X', IntToStr(TobSousSessions.Detail[j].GetValue('PSS_ORDRE')),IntToStr(TobSousSessions.Detail[j].GetValue('PSS_NOSESSIONMULTI'))) Then
						Break;
                End;
			End;
		end;
		PGTobRecupPrev.Free;

		//PT27
		If GetField('PSS_PGTYPESESSION') <> 'ENT' Then
			MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(GetField('PSS_ORDRE')), Nil)
		Else
		Begin
			Try //PT27
				For i := 0 To TobSousSessions.Detail.Count - 1 Do
					MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(TobSousSessions.Detail[i].GetValue('PSS_ORDRE')), Nil);
			Except
			End;
		End;

		If (TFFiche(Ecran).Name = 'SESSIONSIMPLE') Then AfficheStagiaires;

		If DS.State in [dsInsert] Then SetControlEnabled('PSS_CODESTAGE', False);
		
		If Assigned(TobSousSessions) Then FreeAndNil(TobSousSessions); //PT27
	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Inscription d'un salarié à la session
Mots clefs ... :
*****************************************************************}
Function TOM_SESSIONSTAGE.InscriptionFormation(Salarie,TypePlan,Typologie : String; FormationEffectuee:Boolean=True; NumSession:String='';NoSessionMulti:String='') : Boolean;
var Stage,Session,Millesime : String;
    Q : TQuery;
    DatePaie : TDateTime;
    DefautHTpsT : Boolean;
    Req,Pfx,Tabt: String;
    TF          : TOB;
begin
        Result := True;

        If Salarie = '' Then Exit;

        //PT27
        If PGBundleHierarchie Then
            Tabt := 'PGSALARIEINT'
        Else
            Tabt := 'PGSALARIE';
        
        DefautHTpsT := GetParamSoc('SO_PGDIFTPSTRAV');
        Stage       := GetField('PSS_CODESTAGE');
        Millesime   := GetField('PSS_MILLESIME');
        If NumSession <> '' Then Session := NumSession //PT27
        Else Session:= GetField('PSS_ORDRE');
        
        If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" AND PFO_CODESTAGE="'+Stage+'" '+
        ' AND PFO_ORDRE='+Session+' AND PFO_MILLESIME="'+Millesime+'"') then  //DB2
        begin
                PGIBox(Format(TraduireMemoire('Inscription impossible car le salarié %s est déjà inscrit'),[RechDom('PGSALARIE',Salarie,False)]),Ecran.Caption);
                Result := False;
                Exit;
        end;

        TF := TOB.Create('FORMATIONS', TobInscrits, -1);

        TF.PutValue('PFO_SALARIE',Salarie);
        If Typologie <> '' then TF.PutValue('PFO_TYPOFORMATION',Typologie)
        else TF.PutValue('PFO_TYPOFORMATION','001');

        // One ne recherche la date du bulletin que si la formation a été effectuée (donc pas dans le cas d'une inscription).
        // Sinon, on considère 01/01/1900.
        If (FormationEffectuee) Then
        Begin
        	If PGBundleHierarchie Then
        		Q := OpenSQL('SELECT MAX (PPU_DATEFIN) MAXDATE FROM '+GetBase(GetBaseSalarie(Salarie),'PAIEENCOURS')+' WHERE PPU_SALARIE="'+Salarie+'"',True)
        	Else
             	Q := OpenSQL('SELECT MAX (PPU_DATEFIN) MAXDATE FROM PAIEENCOURS WHERE PPU_SALARIE="'+Salarie+'"',True);
             If Not Q.Eof then DatePaie := Q.FindField('MAXDATE').AsDateTime
             Else DatePaie := V_PGI.DateEntree;
             DatePaie := PlusMois(DatePaie,1);
             DatePaie := FinDeMois(DatePaie);
             Ferme(Q);
             If DatePaie < V_PGI.dateEntree then DatePaie := FinDeMois(V_PGI.DateEntree);
             TF.PutValue('PFO_DATEPAIE',DatePaie);
			 TF.PutValue('PFO_EFFECTUE','X');
        End
        Else
        Begin
             DatePaie := iDate1900;
             TF.PutValue('PFO_DATEPAIE',DatePaie);
             TF.PutValue('PFO_EFFECTUE','-');
        End;

        TF.PutValue('PFO_TYPEPLANPREV',TypePlan);
        TF.PutValue('PFO_CODESTAGE',Stage);
        TF.PutValue('PFO_ORDRE',StrToInt(Session));
        TF.PutValue('PFO_MILLESIME',Millesime);

        Req := 'SELECT PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_PRENOM,PSE_RESPONSFOR'+
        ',PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4 FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE WHERE PSA_SALARIE="'+Salarie+'"';

        Req := AdaptMultiDosForm (Req, False);
        If PGBundleHierarchie Then Pfx := 'PSI' Else Pfx := 'PSA';

        Q := OpenSQL(Req,True);
        if not Q.eof then
        begin
                TF.PutValue('PFO_TRAVAILN1',Q.FindField(Pfx+'_TRAVAILN1').AsString);
                TF.PutValue('PFO_TRAVAILN2',Q.FindField(Pfx+'_TRAVAILN2').AsString);
                TF.PutValue('PFO_TRAVAILN3',Q.FindField(Pfx+'_TRAVAILN3').AsString);
                TF.PutValue('PFO_TRAVAILN4',Q.FindField(Pfx+'_TRAVAILN4').AsString);
                TF.PutValue('PFO_CODESTAT',Q.FindField(Pfx+'_CODESTAT').AsString);
                TF.PutValue('PFO_NOMSALARIE',Q.FindField(Pfx+'_LIBELLE').AsString);
                TF.PutValue('PFO_PRENOM',Q.FindField(Pfx+'_PRENOM').AsString);
                TF.PutValue('PFO_ETABLISSEMENT',Q.FindField(Pfx+'_ETABLISSEMENT').AsString);
                TF.PutValue('PFO_LIBEMPLOIFOR',Q.FindField(Pfx+'_LIBELLEEMPLOI').AsString);
                TF.PutValue('PFO_RESPONSFOR',Q.FindField('PSE_RESPONSFOR').AsString);
                TF.PutValue('PFO_LIBREPCMB1',Q.findField(Pfx+'_LIBREPCMB1').AsString);
                TF.PutValue('PFO_LIBREPCMB2',Q.findField(Pfx+'_LIBREPCMB2').AsString);
                TF.PutValue('PFO_LIBREPCMB3',Q.findField(Pfx+'_LIBREPCMB3').AsString);
                TF.PutValue('PFO_LIBREPCMB4',Q.findField(Pfx+'_LIBREPCMB4').AsString);
        end;
        Ferme(Q);

        TF.PutValue('PFO_DATEDEBUT',GetField('PSS_DATEDEBUT'));
        TF.PutValue('PFO_DATEFIN',GetField('PSS_DATEFIN'));
        TF.PutValue('PFO_NATUREFORM',GetField('PSS_NATUREFORM'));
        TF.PutValue('PFO_LIEUFORM',GetField('PSS_LIEUFORM'));
        TF.PutValue('PFO_CENTREFORMGU',GetField('PSS_CENTREFORMGU'));//DB2
        TF.PutValue('PFO_DEBUTDJ',GetField('PSS_DEBUTDJ'));
        TF.PutValue('PFO_FINDJ',GetField('PSS_FINDJ'));
        TF.PutValue('PFO_ORGCOLLECT',   -1);
        TF.PutValue('PFO_ORGCOLLECTGU',GetField('PSS_ORGCOLLECTSGU'));
        TF.PutValue('PFO_NBREHEURE',GetField('PSS_DUREESTAGE'));
        if (DefautHTpsT) or (TypePlan = 'PLF') then
        begin
             TF.PutValue('PFO_HTPSTRAV', GetField('PSS_DUREESTAGE'));
             TF.PutValue('PFO_HTPSNONTRAV', 0);
        end
        else
        begin
             TF.PutValue('PFO_HTPSTRAV', 0);
             TF.PutValue('PFO_HTPSNONTRAV', GetField('PSS_DUREESTAGE'));
        end;

        TF.PutValue('PFO_HEUREDEBUT',    GetField('PSS_HEUREDEBUT'));
        TF.PutValue('PFO_HEUREFIN',      GetField('PSS_HEUREFIN'));
        TF.PutValue('PFO_FORMATION1',    GetField('PSS_FORMATION1'));
        TF.PutValue('PFO_FORMATION2',    GetField('PSS_FORMATION2'));
        TF.PutValue('PFO_FORMATION3',    GetField('PSS_FORMATION3'));
        TF.PutValue('PFO_FORMATION4',    GetField('PSS_FORMATION4'));
        TF.PutValue('PFO_FORMATION5',    GetField('PSS_FORMATION5'));
        TF.PutValue('PFO_FORMATION6',    GetField('PSS_FORMATION6'));
        TF.PutValue('PFO_FORMATION7',    GetField('PSS_FORMATION7'));
        TF.PutValue('PFO_FORMATION8',    GetField('PSS_FORMATION8'));
        TF.PutValue('PFO_INCLUSDECL',    GetField('PSS_INCLUSDECL'));
        If GetParamSocSecur('SO_PGFORMVALIDREA', False) Then
            TF.PutValue('PFO_ETATINSCFOR',   'ATT')
        Else
            TF.PutValue('PFO_ETATINSCFOR',   'VAL');

		// Dans le cas d'inscription à la session en-tête, le NumSession est renseigné
		If NumSession = '' Then //PT27
		Begin
	        TF.PutValue('PFO_NOSESSIONMULTI', GetField('PSS_NOSESSIONMULTI'));
	        TF.PutValue('PFO_IDSESSIONFOR',   GetField('PSS_IDSESSIONFOR'));
	        TF.PutValue('PFO_PGTYPESESSION',  GetField('PSS_PGTYPESESSION'));
	    End
	    Else
	    Begin
	        TF.PutValue('PFO_NOSESSIONMULTI', NoSessionMulti);
	        TF.PutValue('PFO_IDSESSIONFOR',   GetField('PSS_CODESTAGE')+NumSession);
	        TF.PutValue('PFO_PGTYPESESSION',  'SOS');
	    End;
                
        TF.PutValue('PFO_DATECOINVES',	IDate1900);
        TF.PutValue('PFO_DATEACCEPT',	IDate1900);
        TF.PutValue('PFO_REFUSELE',		IDate1900);
        TF.PutValue('PFO_REPORTELE',	IDate1900);

        If PGBundleInscFormation Then
        Begin
        	TF.PutValue('PFO_NODOSSIER', GetNoDossierSalarie(Salarie));
        	TF.PutValue('PFO_PREDEFINI', 'DOS');
        End
        Else
        Begin
        	TF.PutValue('PFO_NODOSSIER', V_PGI.NoDossier);
        	TF.PutValue('PFO_PREDEFINI', 'STD');
        End;
        
        TF.InsertOrUpdateDB();
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Clic sur le bouton "Inscriptions multiples"
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.SaisieEnMasse (Sender : TObject);
var Stage,Session,Millesime     : String;
    ListeSalaries,RetSauve,Ret  : String;
    TobSousSessions             : TOB;
    i                           : Integer;
    Sal                         : String;
begin
	Stage     := GetField('PSS_CODESTAGE');
	Session   := GetField('PSS_ORDRE');
	Millesime := GetField('PSS_MILLESIME');

	Ret := AGLLanceFiche('PAY','INTERIMINSC_MUL','','','SESSION;PFO_CODESTAGE="'+Stage+'" AND PFO_MILLESIME="'+Millesime+'"');

	If Ret = '' then exit;

	// Chargement des données des salariés
	RetSauve := Ret;
	While RetSauve<>'' Do
	Begin
		ListeSalaries := ListeSalaries + '"' + ReadTokenSt(RetSauve) + '"';
		If RetSauve <> '' Then ListeSalaries := ListeSalaries + ',';
	End;
	
	//PT27
	// Chargement des sous-sessions
	If GetField('PSS_PGTYPESESSION') = 'ENT' Then
	Begin
  		TobSousSessions := TOB.Create('LesSousSessions', Nil, -1);
  		TobSousSessions.LoadDetailDBFromSQL('LesSousSessions','SELECT PSS_ORDRE,PSS_NOSESSIONMULTI FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PSS_MILLESIME')+'" AND PSS_NUMSESSION="'+IntToStr(GetField('PSS_ORDRE'))+'"');
  	End;

	// Boucle sur les stagiaires sélectionnés
	While Ret <> '' Do
	Begin
		If GetField('PSS_PGTYPESESSION') <> 'ENT' Then
			InscriptionFormation(ReadTokenSt(Ret),'PLF','', GetField('PSS_EFFECTUE')='X')
		Else
		Begin //PT27
            Sal := ReadTokenSt(Ret);
			For i := 0 To TobSousSessions.Detail.Count - 1 Do
				If Not InscriptionFormation(Sal,'PLF','', GetField('PSS_EFFECTUE')='X', IntToStr(TobSousSessions.Detail[i].GetValue('PSS_ORDRE')),IntToStr(TobSousSessions.Detail[i].GetValue('PSS_NOSESSIONMULTI'))) Then
					Break;
		End;
	End;

	//PT27
	If GetField('PSS_PGTYPESESSION') <> 'ENT' Then
		MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(GetField('PSS_ORDRE')), Nil)
	Else
	Begin
		Try //PT27
			For i := 0 To TobSousSessions.Detail.Count - 1 Do
				MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(TobSousSessions.Detail[i].GetValue('PSS_ORDRE')), Nil);
		Except
		End;
	End;

	If (TFFiche(Ecran).Name = 'SESSIONSIMPLE') Then AfficheStagiaires;

	If DS.State in [dsInsert] Then SetControlEnabled('PSS_CODESTAGE', False);
	
	If Assigned(TobSousSessions) Then FreeAndNil(TobSousSessions); //PT27
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Récupération des salariés inscrits à la session
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.AfficheStagiaires;
begin
	If GrilleSal = Nil Then Exit;

	// Vidage de la grille
	GrilleSal.RowCount := 2;
	GrilleSal.Rows[1].Clear;

	If TobInscrits = Nil Then
	begin
		TobInscrits := Tob.Create('les stagiaires',Nil,-1);
		TobInscrits.LoadDetailDBFromSQL('FORMATIONS','SELECT * FROM FORMATIONS WHERE '+
										'PFO_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFO_ORDRE='+IntToStr(GetField('PSS_ORDRE'))+' AND PFO_MILLESIME="'+GetField('PSS_MILLESIME')+'" ORDER BY PFO_NOMSALARIE,PFO_PRENOM');
	end;

	If TobInscrits.Detail.Count > 0 Then
    Begin
		TobInscrits.PutGridDetail(GrilleSal,False,False,'PFO_SALARIE;PFO_NOMSALARIE;PFO_PRENOM;PFO_TYPEPLANPREV;PFO_NBREHEURE;PFO_EFFECTUE');
        If TobInscrits.FindFirst(['PFO_EFFECTUE'],['-'],False) <> Nil Then
            SetControlEnabled('BVALIDERINSC', True)
        Else
            SetControlEnabled('BVALIDERINSC', False);
        If TobInscrits.FindFirst(['PFO_EFFECTUE'],['X'],False) <> Nil Then
            GrilleSal.ColWidths[COL_PRESENCE] := 60
        Else
            GrilleSal.ColWidths[COL_PRESENCE] := -1;
    End
    Else
        SetControlEnabled('BVALIDERINSC', False);

    If GetField('PSS_CLOTUREINSC') = 'X' Then
    Begin
        SetControlEnabled('BRECUPERATION', False);
        SetControlEnabled('BMULTISAL',     False);
        SetControlEnabled('BAJOUTSTAG',    False);
        SetControlEnabled('BSUPPRSTAG',    False);
    End;
    If GetField('PSS_EFFECTUE') = 'X' Then
    Begin
        SetControlEnabled('BVALIDERINSC',  False);
    End;

	TFFiche(Ecran).HMTrad.ResizeGridColumns(GrilleSal);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Gestion des raccourcis clavier pour la saisie des inscriptions 
Suite ........ : en SESSIONSIMPLE
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    If GrilleSal = Nil Then Exit;

	// Gestion de la grille
    If THPageControl2(GetControl('Pages')).ActivePage = TTabSheet(GetControl('PINSCRIPTIONS')) Then
    Begin
    	// Suppression d'une ligne
	    If (Key = VK_DELETE) And (ssCtrl in Shift) Then
	    Begin
            BSupprStagiaireClick(Nil);
	    End
	    // Création d'une nouvelle ligne d'inscription
	    Else If ((Key = VK_DOWN) Or ((Key = VK_TAB) And (GrilleSal.Col = GrilleSal.ColCount-1))) And (GrilleSal.Row = GrilleSal.RowCount-1) Then
	    Begin
			If Not NewLine Then
            Begin
                BAjoutStagiaireClick(Nil);
            End;
		End
	    // Annulation création d'une nouvelle ligne d'inscription
	    Else If ((Key = VK_UP) Or ((Key = VK_TAB) And (ssShift in Shift) And (GrilleSal.Col = 0))) And (GrilleSal.Row = GrilleSal.RowCount-1) Then
	    Begin
            If (GrilleSal.Row > 1) And (GrilleSal.CellValues[COL_SALARIE,GrilleSal.Row] = '') Then
            Begin
                BSupprStagiaireClick(Nil);
                Key := 0;
            End;
	    End;
	End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Fermeture de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnClose;
var TempoBool : Boolean;
    i : Integer;
    T : TOB;
begin
  inherited;
  	If Assigned(TobModifiedFields) Then FreeAndNil(TobModifiedFields); //PT27
  	
	If (TFFiche(Ecran).Name = 'SESSIONSIMPLE') Then
	Begin
		If (DS.State in [dsInsert]) Then
		Begin
			If (GrilleSal.CellValues[COL_SALARIE,1] <> '') And (Not NewLine) Then
			Begin
				LastError := 1;
				LastErrorMsg := TraduireMemoire('Vous ne pouvez pas annuler la création de cette session car des inscriptions lui sont affectées.');
				Exit;
			End;
		End;

		If THPageControl2(GetControl('Pages')).ActivePage = TTabSheet(GetControl('PINSCRIPTIONS')) Then
		Begin
			If GrilleSal.CellValues[COL_SALARIE,GrilleSal.Row] <> '' Then
			Begin
				ValideComboDansGrid (GrilleSal, GrilleSal.Col, 0, GrilleSal.Row);
				GrilleSal.OnRowExit(Nil, GrilleSal.Row,tempoBool,False);
			End;
		End;

        // Mise à jour des données possiblement modifiées
        For i := 1 To GrilleSal.RowCount-1 Do
        Begin
        	T := TobInscrits.FindFirst(['PFO_SALARIE'],[GrilleSal.CellValues[COL_SALARIE,i]],True);
        	If T <> Nil Then
        	Begin
        		T.PutValue('PFO_TYPEPLANPREV', GrilleSal.CellValues[COL_TYPEPLAN,i]);
        		T.PutValue('PFO_NBREHEURE',    GrilleSal.CellValues[COL_NBHEURES,i]);
        		T.PutValue('PFO_EFFECTUE',     GrilleSal.CellValues[COL_PRESENCE,i]);
                if (GetParamSoc('SO_PGDIFTPSTRAV')) or (GrilleSal.CellValues[COL_TYPEPLAN,i] = 'PLF') then
                begin
                     T.PutValue('PFO_HTPSTRAV', GrilleSal.CellValues[COL_NBHEURES,i]);
                     T.PutValue('PFO_HTPSNONTRAV', 0);
                end
                else
                begin
                     T.PutValue('PFO_HTPSTRAV', 0);
                     T.PutValue('PFO_HTPSNONTRAV', GrilleSal.CellValues[COL_NBHEURES,i]);
                end;
                If (T.IsFieldModified('PFO_TYPEPLANPREV') Or T.IsFieldModified('PFO_NBREHEURE') Or T.IsFieldModified('PFO_EFFECTUE'))
                   And (T.GetValue('PFO_EFFECTUE') = 'X') Then
                Begin
                	MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(GetField('PSS_ORDRE')), Nil); //PT27
                    MajCompetences(T.GetValue('PFO_SALARIE'), GetField('PSS_CODESTAGE'), GetField('PSS_ORDRE'), GetField('PSS_MILLESIME'), True);
                    InsererDIFBulletin (T, T.GetValue('PFO_DATEPAIE'));
                End;
        	End;
        End;
  		If TobInscrits.IsOneModifie Then TobInscrits.UpdateDB();
  		
  		If Assigned(TobInscrits) Then FreeAndNil(TobInscrits);
	End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Clic sur le bouton "Afficher plus d'informations"
Suie ......... : en SESSIONSIMPLE
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.BPlusInfosClick(Sender: TObject);
begin
    TTabSheet(GetControl('PCHAMPLIBRE')).TabVisible := True;
    TToolbarButton97(Sender).Enabled := False;
    THPageControl2(GetControl('Pages')).ActivePageIndex := 1;
end;

procedure TOM_SESSIONSTAGE.BAjoutStagiaireClick(Sender: Tobject);
var TempoBool : Boolean;
    Col,Row : Integer;
begin
    If GetField('PSS_CLOTUREINSC') = 'X' Then Exit;

	// Evite de créer une ligne vide
	If (GrilleSal.Row=1) And (Trim(GrilleSal.CellValues[COL_SALARIE,GrilleSal.Row]) = '') Then
	Begin
		GrilleSal.CellValues[COL_TYPEPLAN,GrilleSal.Row] := 'PLF';
		GrilleSal.CellValues[COL_NBHEURES,GrilleSal.Row] := GetField('PSS_DUREESTAGE');
        If GrilleSal.ColWidths[COL_PRESENCE] > 0 Then
    		GrilleSal.CellValues[COL_PRESENCE,GrilleSal.Row] := 'X'
        Else
    		GrilleSal.CellValues[COL_PRESENCE,GrilleSal.Row] := '-';
		OnGSRowEnter (Sender,GrilleSal.Row,TempoBool,False);
		Col := 0;
		Row := GrilleSal.Row;
		GrilleSal.Col := 0;
		OnGSCellEnter(Sender,Col,Row,TempoBool);
		Exit;
	End
	Else
	Begin
		GrilleSal.InsertRow(GrilleSal.RowCount);
		GrilleSal.CellValues[COL_TYPEPLAN,GrilleSal.RowCount-1] := 'PLF';
		GrilleSal.CellValues[COL_NBHEURES,GrilleSal.RowCount-1] := GetField('PSS_DUREESTAGE');
        If GrilleSal.ColWidths[COL_PRESENCE] > 0 Then
    		GrilleSal.CellValues[COL_PRESENCE,GrilleSal.RowCount-1] := 'X'
        Else
    		GrilleSal.CellValues[COL_PRESENCE,GrilleSal.RowCount-1] := '-';
		Col := 0;
		Row := GrilleSal.RowCount - 1;
		GrilleSal.Col := Col;
		If Sender <> Nil Then
		Begin
			GrilleSal.Row := Row;
			OnGSRowEnter (Sender,GrilleSal.Row,TempoBool,False);
			OnGSCellEnter(Sender,Col,Row,TempoBool);
		End;
	End;
	TToolBarButton97(GetControl('BAJOUTSTAG')).Enabled    := False;
	TToolBarButton97(GetControl('BRECUPERATION')).Enabled := False;
	TToolBarButton97(GetControl('BMULTISAL')).Enabled     := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Suppression d'un stagiaire inscrit
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.BSupprStagiaireClick(Sender: Tobject);
var Col,Row : Integer;
    Cancel : Boolean;
    T : TOB;
begin
    If GetField('PSS_CLOTUREINSC') = 'X' Then Exit;

	If (GrilleSal.Row = 1) And (Trim(GrilleSal.CellValues[COL_SALARIE,GrilleSal.Row]) = '') Then
	Begin
		GrilleSal.CellValues[COL_TYPEPLAN,GrilleSal.Row] := '';
		GrilleSal.CellValues[COL_NBHEURES,GrilleSal.Row] := '';
		Exit;
	End;
	
	If (NewLine) Or (Trim(GrilleSal.CellValues[COL_SALARIE,GrilleSal.RowCount-1]) = '') Then
	Begin
		GrilleSal.DeleteRow(GrilleSal.RowCount - 1);
		NewLine := False;
		Col := GrilleSal.Col; Row := GrilleSal.Row;
		OnGSCellEnter(Sender,Col,Row,Cancel);
	End
	Else If PGIAsk(TraduireMemoire('Confirmez-vous la suppression de cet enregistrement?'),TraduireMemoire('Suppression de l''inscription')) = mrYes Then
	Begin
		BEGINTRANS;
		Try
            T := TobInscrits.FindFirst(['PFO_SALARIE'],[GrilleSal.CellValues[COL_SALARIE, GrilleSal.Row]], False);
            If T <> Nil Then
            Begin
                // Suppression de l'enregistrement en cours
                T.DeleteDB();
                FreeAndNil(T);

                // Suppression des frais
    			ExecuteSQL('DELETE FROM FRAISSALFORM WHERE PFS_SALARIE="'+GrilleSal.CellValues[COL_SALARIE, GrilleSal.Row]+'" AND '+
	    		'PFS_CODESTAGE="'+GetField('PSS_CODESTAGE')+'" AND PFS_ORDRE='+IntToStr(GetField('PSS_ORDRE')));

                // Mise à jour des coûts
		    	MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(GetField('PSS_ORDRE')), Nil);

			    COMMITTRANS;
            End;
			Except
			On E: Exception do
			begin
				PGIBox(E.Message);
				ROLLBACK;
			end;
		End;

		// Rafraîchissement de la liste
		AfficheStagiaires;
	End;
	TToolBarButton97(GetControl('BAJOUTSTAG')).Enabled    := True;
	TToolBarButton97(GetControl('BRECUPERATION')).Enabled := True;
	TToolBarButton97(GetControl('BMULTISAL')).Enabled     := True;

    If TobInscrits.Detail.Count = 0 Then
        SetControlEnabled('BVALIDERINSC', False);
    TFFiche(Ecran).HMTrad.ResizeGridColumns(GrilleSal);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 27/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Validation de la présence des stagiaires inscrits
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.BValidStagiaireClick(Sender: TObject);
var i  : Integer;
begin
    For i := 0 To TobInscrits.Detail.Count-1 Do
    Begin
        TobInscrits.Detail[i].PutValue('PFO_EFFECTUE', 'X');
        MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(GetField('PSS_ORDRE')), Nil); //PT27
        MajCompetences(TobInscrits.Detail[i].GetValue('PFO_SALARIE'), GetField('PSS_CODESTAGE'), GetField('PSS_ORDRE'), GetField('PSS_MILLESIME'), True);
        InsererDIFBulletin (TobInscrits.Detail[i], TobInscrits.Detail[i].GetValue('PFO_DATEPAIE'));
    End;

    GrilleSal.ColWidths [COL_PRESENCE] := 60;

    SetField('PSS_EFFECTUE',   'X');
    SetField('PSS_CLOTUREINSC','X');
    
    AfficheStagiaires;

    TToolbarButton97(Sender).Enabled := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Entrée sur une cellule de la grille des salariés inscrits
Suite ........ : Rend éditable ou non la colonne. Cache la combo (bug CBP)
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnGSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if (GrilleSal.col = COL_SALARIE) then
  Begin
    If (NewLine) And (GrilleSal.CellValues[COL_TYPEPLAN,GrilleSal.Row] <> '') Then
    Begin
        GrilleSal.ColEditables[COL_SALARIE] := True;
        GrilleSal.ElipsisButton := True;
    End
    Else
    Begin
        GrilleSal.ColEditables[COL_SALARIE] := False;
        GrilleSal.ElipsisButton := False;
    End;
  End
  Else GrilleSal.ElipsisButton := False;

  If GrilleSal.Col <> COL_TYPEPLAN Then
  Begin
      if Assigned(GrilleSal.ValCombo) then
        GrilleSal.ValCombo.Hide;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Clic elipsis salarié de la grille des salariés inscrits
Suite ........ : Inscrit automatiquement le salarié, ou vide le champ s'il est déjà inscrit
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnGSElipsisClick(Sender : TObject);
var Salarie,Nom,Prenom,StWhere : String;
    Q : TQuery;
begin
	If PGBundleHierarchie then
	Begin
		StWhere := RecupClauseHabilitationLookupList(StWhere);
		ElipsisSalarieMultiDos(Sender, StWhere);
	End
	Else
	LookupList(THEdit(Sender),'Liste des salariés','SALARIES','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM','','PSA_LIBELLE', True,-1);

	Salarie := GrilleSal.CellValues[COL_SALARIE,GrilleSal.Row];

	If Salarie = '' Then Exit;

	If Not InscriptionFormation(Salarie,'PLF','', GetField('PSS_EFFECTUE')='X') Then
	Begin
		// Si l'inscription ne fonctionne pas, on supprime les valeurs pour éviter les problèmes
		GrilleSal.CellValues[COL_SALARIE,GrilleSal.Row] := '';
	End
	Else
	Begin
		Nom := '';
		Prenom := '';
		Q := openSQL('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
		If Not Q.eof then
		begin
			Nom := Q.FindField('PSA_LIBELLE').AsString;
			Prenom := Q.FindField('PSA_PRENOM').AsString;
		end;
		Ferme(Q);
		GrilleSal.CellValues[1,GrilleSal.Row] := Nom;
		GrilleSal.CellValues[2,GrilleSal.Row] := Prenom;

		MAJCoutsFormation(GetField('PSS_MILLESIME'), GetField('PSS_CODESTAGE'), IntToStr(GetField('PSS_ORDRE')), Nil);

		If DS.State in [dsInsert] Then SetControlEnabled('PSS_CODESTAGE', False);

        NewLine := False;

		TToolBarButton97(GetControl('BAJOUTSTAG')).Enabled    := True;
		TToolBarButton97(GetControl('BRECUPERATION')).Enabled := True;
		TToolBarButton97(GetControl('BMULTISAL')).Enabled     := True;

    	If TobInscrits.Detail.Count > 0 Then
            SetControlEnabled('BVALIDERINSC', True);
    	TFFiche(Ecran).HMTrad.ResizeGridColumns(GrilleSal);
	End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Sortie de ligne de la grille des salariés inscrits.
Suite ........ : Mise à jour de la valeur du plan de formation ou suppression de la ligne
Suite ........ : s'il s'agit d'une nouvelle ligne vide
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnGSRowExit (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
Begin
    If ((NewLine) Or (GrilleSal.CellValues[COL_SALARIE,Ou] = '')) And (Ou > 1) Then
		BSupprStagiaireClick(Nil);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/05/2008 / PT25
Modifié le ... :   /  /
Description .. : Entrée sur une ligne de la grille des salariés inscrits.
Mots clefs ... :
*****************************************************************}
procedure TOM_SESSIONSTAGE.OnGSRowEnter (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
Begin
    If GrilleSal.CellValues[COL_SALARIE,Ou] = '' Then
		NewLine := True;
End;

procedure TOM_SESSIONSTAGE.OnGSClick(Sender: TOBject);
begin
    If (GrilleSal.Col = COL_PRESENCE) Then
    Begin
        if GrilleSal.Cells[GrilleSal.Col,GrilleSal.Row]='X' then
            GrilleSal.Cells[GrilleSal.Col,GrilleSal.Row]:='-'
        else
            GrilleSal.Cells[GrilleSal.Col,GrilleSal.Row]:='X';
    End;
end;

Initialization
    registerclasses([TOM_SESSIONSTAGE]) ;
end.


