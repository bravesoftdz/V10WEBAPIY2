{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 16/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITCOMPARATIFFOR ()
Mots clefs ... : TOF;PGEDITCOMPARATIFFOR
Contient les sources TOF pour les éditions de la formation
*****************************************************************
PT1  | 24/04/2003 | V_42  | JL | Développement pour CWAS : Modification de tous els lookuplist
PT2  | 28/02/2005 | V_60  | JL | FQ 12035 Correction nom des THEdit pour dates
---  | 20/03/2006 |       | JL | modification clé annuaire ----
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT3  | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
     |            |       |    |  pour les critères code salarié uniquement
PT4  | 19/04/2007 | V_720 | FL | FQ 14054 Ajout de la TOF associée à l'édition des cursus
PT5  | 24/04/2007 | V_720 | FL | Editions budget collectif et infividuel dans la fenêtre principale.
     |            |       |    |  Niveaux de rupture dynamiques.
PT6  | 24/04/2007 | V_720 | FL | FQ 12255 Gestion du comparatif pluri-annuel
PT7  | 26/04/2007 | V_720 | FL | Utilisation d'une TOB pour l'édition des frais salariés et animateurs
PT8  | 16/05/2007 | V_720 | JL | exclusion des sous sessions
PT9  | 23/05/2007 | V_720 | JL | Modif clé table fraissalform
PT10 | 05/06/2007 | V_720 | FL | Edition du suivi des formations par salarié
PT11 | 12/06/2007 | V_720 | FL | FQ 14054 Evolution : Rattachement des formations aux cursus réalisés
     |            |       |    |  même si le top n'a pas été indiqué par l'utilisateur.
PT12 | 10/09/2007 | V_800 | FL | FQ 11531 Plus de lanceEtatTob pour les feuilles de présence
PT13 | 01/10/2007 | V_800 | FL | Adaptation accès multi dossiers                         
PT14 | 26/10/2007 | V_800 | FL | Emanager / Report / Adaptation    
PT15 | 08/01/2008 | V_802 | FL | FQ 14790 Récupération du nom de l'animateur s'il est externe                     
PT16 | 09/01/2008 | V_802 | FL | FQ 14793 Affichage de la coche rupture dès la première valeur saisie
PT17 | 19/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT18 | 03/04/2008 | V_803 | FL | Adaptation partage formation
PT19 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT20 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue + Gestion elipsis salarié et responsable
PT21 | 19/05/2008 | V_804 | FL | Correction de la FQ 14054 : doivent apparaître les inscriptions prévues non encore réalisées
PT22 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT23 | 10/06/2008 | V_804 | FL | Adaptation en multidossier des éditions prévisionnels individuel et collectif
PT24 | 04/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
PT25 | 09/07/2008 | V_804 | FL | Suivi plan d'intégration
}

unit UtofPGEditionsFormation;

interface
Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtREtat,mul,uTobDebug,
{$ELSE}
     eQRS1,UtilEAgl,emul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,PgoutilsFormation,EntPaie,HTB97,LookUp,ParamDat,HQry,P5DEF,Paramsoc,PGEdtEtat,PGOutils2,ed_tools,Hstatus; //PT4

//PT7 - Début
Const PfxFOR = 'PFO';
Const PfxSAL = 'PSA';
Const PfxINT = 'PSI';
Const PfxANI = 'PAN';
//PT7 - Fin

Type
  TOF_PGEDITCOMPARATIFFOR = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure StageElipsisClick(Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure ChangeMillesime(Sender : TObject);
  end ;

  TOF_PGEDITFACTUREFORM = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure StageElipsisClick(Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure ChangeMillesime(Sender : TObject);
    procedure StageChange(Sender : TObject);
  end ;

  TOF_PGFEUILLEPRESENCEFOR = Class (TOF)
    procedure OnUpdate                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Procedure OnClose                  ; override ; //PT12
    private
    TobEtat : TOB;                                  //PT12
    procedure StageElipsisClick(Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure ChangeMillesime(Sender : TObject);
    procedure StageChange(Sender : TObject);
    procedure EditerFeuillePresence(Sender : TObject);
  end ;

  TOF_PGEDITSUIVIBUDFORM = Class (TOF)
    procedure OnUpdate                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFDEF EMANAGER}TypeUtilisat : String;{$ENDIF}
    procedure ExitMotifInsc(Sender : TObject);
    procedure StageElipsisClick(Sender : TObject);
    procedure RespElipsisClick(Sender : TObject);
    procedure SalarieElipsisClick(Sender : TObject);
    procedure ExitEdit(Sender : TObject);
  end;

  TOF_PGEDITFRAISSAL = Class (TOF)
    procedure OnUpdate ; override;
    procedure OnArgument (S : String ) ; override ;
    Procedure OnClose                  ; override ;
    private
    TypeEdition : String;
    SuffixeFor, SuffixeSal : String; //PT7
    TobEtat : TOB; //PT7
    Tablettes : Array of Array[0..1] of String; //PT7
    Function  RechTablette (NomChamp : String) : String;//PT7
    procedure StageElipsisClick(Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure ChangeMillesime(Sender : TObject);
    procedure StageChange(Sender : TObject);
    //{$IFDEF EMANAGER}
    procedure SalarieElipsisClick(Sender : TObject);
    //{$ENDIF}
    Procedure RespElipsisClick(Sender : TObject); //PT18
    Procedure OnChangeRuptures (Sender : TObject); //PT7
  end;

  TOF_PGEDITBUDGETFOR = Class (TOF)
    procedure OnUpdate           ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    OrderBy : String;
    Arg : String;
    {$IFDEF EMANAGER}
    MultiNiveau : Boolean;
    {$ENDIF}
    procedure SalarieElipsisClick(Sender : TObject);
    procedure StageElipsisClick(Sender : TObject);
    procedure CreerRupture(Champ : String;Num : Integer);
    procedure RespElipsisClick(Sender : TObject);
  end;

  TOF_PGEDITCOMBUDGET  = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure StageElipsisClick(Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure ChangeMillesime(Sender : Tobject);
  end;

  TOF_PGEDITCOMPSESSION = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure StageElipsisClick(Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure ChangeMillesime(Sender : Tobject);
    procedure StageChange(Sender : TObject);
  end;

  TOF_PGEDITFORREABUD = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    Procedure OnClose                  ; override ;  //PT5
    private
    LeMenu : String;
    TobEtat : Tob;  //PT5
    procedure EditionComparCollectif; //PT5
    procedure EditionComparIndividuel; //PT5
    procedure EditionComparPluriAnnuel; //PT6
    Procedure OnChangeRuptures (Sender : TObject); //PT5
    Function ChampRuptureForm(TypeRupt,ValeurRupt : String) : String;
    procedure ExitEdit(Sender : TObject);
    Function GetMillesimes () : String; //PT6
	procedure SalarieElipsisClick(Sender : TObject); //PT18
  end ;

  TOF_PGEDITPLANFORMATION = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    LeMenu : String;
    {$IFDEF EMANAGER}
    MultiNiveau : Boolean;
    {$ENDIF}
    Procedure RespElipsisClick(Sender : TObject); //PT18
    procedure SalarieElipsisClick(Sender : TObject);
    procedure ExitEdit(Sender : TObject);
  end ;

  // Suivi formation - Catalogue au réalisé
  TOF_PGETATLIBREFOR = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;
  end;

  //PT4 - Début
  TOF_PGEDITSUIVICURSUS = Class (TOF)
    Procedure OnUpdate                 ; override ;
    Procedure OnArgument (S : String ) ; override ;
    Procedure OnClose                  ; override ;

    private
    TobEdition : Tob;
    {$IFDEF EMANAGER}
    TypeUtilisat : String; //PT14
    {$ELSE}
    Procedure ExitEdit(Sender : TObject);
    {$ENDIF}
    Procedure RecupData (StWhere : String);
    Procedure CopieData (Donnees : TOB; LibelleCursus, Stage, LibelleStage : String; DateDebut,DateFin : TDateTime; Realise : String);
//    {$IFDEF EMANAGER}
    Procedure SalarieElipsisClick (Sender : TObject); //PT14
//    {$ENDIF}
	{$IFNDEF EMANAGER}
	Procedure RespElipsisClick (Sender : TObject); //PT18
	{$ENDIF}
  end ;
  //PT4 - Fin

  //PT10 - Début
  TOF_PGEDITSUIVIFORM = Class (TOF)
    Procedure OnUpdate                 ; override ;
    Procedure OnArgument (S : String ) ; override ;
    Procedure OnClose                  ; override ;
    private
    TobEdition : Tob;
    procedure CkSortieClick (Sender : TObject);
    procedure DateElipsisClick (Sender : TObject);
    procedure ExitEdit (Sender : TObject);
    Procedure SalarieElipsisClick (Sender : TObject); //PT18
  end ;
  //PT10 - Fin
  
//PT25
  TOF_PGEDTSUIVIINTEGR = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    Procedure OnClose                  ; override ;
    private
{$IFDEF EMANAGER}
    TypeUtilisat : String;
{$ENDIF}
    TobEtat : Tob;
    Tablettes : Array of Array[0..1] of String;
    Procedure OnChangeRuptures   (Sender : TObject);
//    Function  ChampRuptureForm   (TypeRupt,ValeurRupt : String) : String;
    Function  RechTablette (NomChamp : String) : String;
	procedure SalarieElipsisClick(Sender : TObject);
	procedure RespElipsisClick   (Sender : TObject);
	procedure ChargeEdition;
  end ;
//FIN PT25

implementation

uses pgOutils, TntStdCtrls, TypInfo, PgEditOutils2, utilPGI, GalOutil, DateUtils;

{TOF_PGEDITCOMPARATIFFOR}
// Edition du comparatif prévisionnel/réalisé
procedure TOF_PGEDITCOMPARATIFFOR.OnArgument (S : String ) ;
var Stage : THEdit;
    Millesime : THValComboBox;
    Datedebut,DateFin : THEdit;
    DD,DF : TdateTime;
    Num : Integer;
begin
  Inherited ;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        Stage := THEdit(Getcontrol('PST_CODESTAGE'));
        Millesime := THValComboBox(GetControl('PST_MILLESIME'));
        If Millesime <> Nil Then Millesime.OnChange := ChangeMillesime;
        dateDebut := THEdit(GetControl('DATEDEBUT'));
        DateFin := THEdit(GetControl('DATEFIN'));
        If DateFin <> Nil Then DateFin.OnElipsisClick := DateElipsisClick;
        If DateDebut <> Nil Then DateDebut.OnElipsisClick := DateElipsisClick;
        If Stage <> Nil then Stage.OnElipsisClick := StageElipsisClick;
        Millesime.Value := RendMillesimeRealise(DD,DF);
        DateDebut.Text := DateToStr(DD);
        DateFin.Text := DateToStr(DF);
        SetControlCaption('LIBSTAGE','');
end ;

procedure TOF_PGEDITCOMPARATIFFOR.StageElipsisClick;
var  St,StWhere : String;
     Millesime : THValComboBox;
begin
        Millesime := THValComboBox(GetControl('PST_MILLESIME'));
        St := ' SELECT PST_CODESTAGE,PST_MILLESIME,PST_LIBELLE FROM STAGE ';
        StWhere := 'PST_ACTIF="X" AND'+
        ' PST_MILLESIME="'+Millesime.Value+'"';
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,'', True,-1);
end;

Procedure TOF_PGEDITCOMPARATIFFOR.DateElipsisClick(Sender : TObject);
var key : char;
begin
        key := '*';
        ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITCOMPARATIFFOR.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+THValComboBox(Sender).Value+'"',True);
        If Not Q.Eof Then
        begin
                SetControlText('DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

{TOF_PGEDITFACTUREFORM}
// Edition des factures de paie formation
procedure TOF_PGEDITFACTUREFORM.OnArgument (S : String ) ;
var Stage : THEdit;
    Millesime,Session : THValComboBox;
    Datedebut,DateFin : THEdit;
    DD,DF : TDateTime;
begin
  Inherited ;
        SetControlCaption('LIBSTAGE','');
        Stage := THEdit(Getcontrol('PSS_CODESTAGE'));
        Session := THValComboBox(GetControl('PSS_ORDRE'));
        If Session <> Nil Then Session.Value := '';
        Millesime := THValComboBox(GetControl('MILLESIME'));
        If Millesime <> Nil Then Millesime.OnChange := ChangeMillesime;
        dateDebut := THEdit(GetControl('DATEDEBUT'));
        DateFin := THEdit(GetControl('DATEFIN'));
        If DateFin <> Nil Then DateFin.OnElipsisClick := DateElipsisClick;
        If DateDebut <> Nil Then DateDebut.OnElipsisClick := DateElipsisClick;
        If Stage <> Nil then
        begin
                Stage.OnElipsisClick := StageElipsisClick;
                Stage.OnChange := StageChange;
        end;
        Millesime.Value := RendMillesimeRealise(DD,DF);
        DateDebut.Text := DateToStr(DD);
        DateFin.Text := DateToStr(DF);
        SetControlCaption('LIBSTAGE','');
end ;

procedure TOF_PGEDITFACTUREFORM.StageChange(Sender : TObject);
var Session : THValComboBox;
begin
        Session := THValComboBox(GetControl('PSS_ORDRE'));
        If GetControlText('PSS_CODESTAGE') <> '' Then
        begin
                Session.Enabled := True;
                Session.Plus := 'PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"'+
                ' AND PSS_CODESTAGE="'+GetControlText('PSS_CODESTAGE')+'"';
        end
        Else Session.enabled := False;
end;

procedure TOF_PGEDITFACTUREFORM.StageElipsisClick(Sender : TObject);
var  St,StWhere : String;
     Session : THValComboBox;
begin
        St := ' SELECT PST_CODESTAGE,PST_MILLESIME,PST_LIBELLE FROM STAGE ';
        StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="0000" AND'+
        ' PST_CODESTAGE IN (SELECT PSS_CODESTAGE FROM SESSIONSTAGE WHERE PSS_DATEDEBUT>="'+UsDatetime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
        ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'")';
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE',StWhere,'', True,-1);
        Session := THValComboBox(GetControl('PSS_ORDRE'));
        If GetControlText('PSS_CODESTAGE') <> '' Then
        begin
                Session.Enabled := True;
                Session.Value := '';
                Session.Plus := 'PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"'+
                ' AND PSS_CODESTAGE="'+GetControlText('PSS_CODESTAGE')+'"';
        end
        Else
        begin
                Session.Enabled := False;
                Session.Value := '';
                Session.Plus := '';
        end;
end;

Procedure TOF_PGEDITFACTUREFORM.DateElipsisClick(Sender : TObject);
var key : char;
begin
        key := '*';
        ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITFACTUREFORM.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+THValComboBox(Sender).Value+'"',True);
        If Not Q.Eof Then
        begin
                SetControlText('DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

{TOF_PGFEUILLEPRESENCEFOR}
// Edition des feuilles de présences
procedure TOF_PGFEUILLEPRESENCEFOR.OnUpdate;
begin
Inherited ;
     //PT12 - Début
{        TFQRS1(Ecran).WhereSQl := 'SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_DATEDEBUT,PSS_DATEFIN'+
        ',PSS_LIEUFORM,PST_DUREESTAGE,PSS_JOURSTAGE,PST_LIBELLE,PST_LIBELLE1 FROM SESSIONSTAGE'+
        ' LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME'+
        ' WHERE PSS_CLOTUREINSC="X" AND PSS_DATEDEBUT>"'+UsDateTime(StrToDate(GetControlText('PSS_DATEDEBUT')))+'" AND PSS_DATEDEBUT<="'+UsDatetime(StrToDate(GetControltext('PSS_DATEFIN')))+'"';//PT4
}
     EditerFeuillePresence(Nil);
     //PT12 - Fin        
end;

procedure TOF_PGFEUILLEPRESENCEFOR.OnClose;
Begin
	If TobEtat <> Nil Then FreeAndNil(TobEtat); //PT12
End;

procedure TOF_PGFEUILLEPRESENCEFOR.OnArgument (S : String ) ;
var Stage : THEdit;
    Millesime,Session : THValComboBox;
    Datedebut,DateFin : THEdit;
    DD,DF : TDateTime;
    //BEtat : TToolBarButton97;
begin
  Inherited ;
	SetControlCaption('LIBSTAGE','');
	Stage := THEdit(Getcontrol('PSS_CODESTAGE'));
	Session := THValComboBox(GetControl('PSS_ORDRE'));
	If Session <> Nil Then Session.Value := '';
	Millesime := THValComboBox(GetControl('MILLESIME'));
	If Millesime <> Nil Then Millesime.OnChange := ChangeMillesime;
	dateDebut := THEdit(GetControl('PSS_DATEDEBUT'));
	DateFin := THEdit(GetControl('PSS_DATEFIN'));
	If DateFin <> Nil Then DateFin.OnElipsisClick := DateElipsisClick;
	If DateDebut <> Nil Then DateDebut.OnElipsisClick := DateElipsisClick;
	If Stage <> Nil then
	begin
		Stage.OnElipsisClick := StageElipsisClick;
		Stage.OnChange := StageChange;
	end;
	Millesime.Value := RendMillesimeRealise(DD,DF);
	DateDebut.Text := DateToStr(DD);
	DateFin.Text := DateToStr(DF);        SetControlCaption('LIBSTAGE','');
	//Betat  := TToolBarButton97(GetControl('BVALIDER'));
	//If BEtat <> Nil then BEtat.OnClick  := EditerFeuillePresence; //PT12
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Attention : Si modification de cette fonction, modifier également 
Suite ........ : TOM_SESSIONSTAGE.EditionPresenceSession
Mots clefs ... :
*****************************************************************}
procedure TOF_PGFEUILLEPRESENCEFOR.EditerFeuillePresence(Sender : TObject);
var Q : TQuery;
    a,i,Ordre,j : Integer;
    TobAnim,TobSta,TobSession,T : Tob;
    Where,Stage,Libelle,Libelle1,Lieu : String;
    Pages : TPageControl;
    NbJours,NbHeures : Double;
    NbJourInt : Integer;
    DateDebut,DateFin : TDateTime;
    WherePredef : String; //PT18
begin
	If TobEtat <> Nil Then FreeAndNil(TobEtat); //PT12

	TobEtat := Tob.Create('Edition',Nil,-1);
	Pages := TPageControl(GetControl('PAGES'));
	Where := RecupWhereCritere(Pages);
	//PT18 - Début
	WherePredef := '';
	If PGBundleInscFormation Then WherePredef := WherePredef + DossiersAInterroger('',V_PGI.NoDossier,'PSS',True,True);
	//PT18 - Fin

	Q := OpenSQL('SELECT * FROM SESSIONSTAGE LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME '+Where+WherePredef,True); //PT18
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

		// Récupération des animateurs
		If PGBundleHierarchie Then   //PT13
		Q := OpenSQL('SELECT PAN_SALARIE,PAN_LIBELLE,PSI_LIBELLE,PSI_PRENOM FROM SESSIONANIMAT '+ //PT15
		'LEFT JOIN INTERIMAIRES ON PAN_SALARIE=PSI_INTERIMAIRE '+
		'WHERE PAN_CODESTAGE="'+Stage+'" AND PAN_ORDRE='+IntToStr(Ordre),True)
		Else
		Q := OpenSQL('SELECT PAN_SALARIE,PAN_LIBELLE,PSA_LIBELLE,PSA_PRENOM FROM SESSIONANIMAT '+ //PT15
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
				If PGBundleHierarchie Then //PT13
				Begin
					//PT15 - Début
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
					//PT15 - Fin
				End
				Else
				Begin
					//PT15 - Début
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
					//PT15 - Fin
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

		// Récupération des stagiaires
		If PGBundleHierarchie Then   //PT13
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
				If PGBundleHierarchie Then //PT13
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
	//PT12 - Début
	// Attribution de la TOB à l'état
	TFQRS1(Ecran).LaTob:= TobEtat;
	//If GetCheckBoxState('FLISTE') = CbChecked then LanceEtatTOB('E','PFO','PFP',TobEtat,True,True,False,Pages,'','',False)
	//else LanceEtatTOB('E','PFO','PFP',TobEtat,True,False,False,Pages,'','',False);
	//TobEtat.Free;
	//PT12 - Fin
end;

procedure TOF_PGFEUILLEPRESENCEFOR.StageChange(Sender : TObject);
var Session : THValComboBox;
begin
	Session := THValComboBox(GetControl('PSS_ORDRE'));
	If GetControlText('PSS_CODESTAGE') <> '' Then
	begin
		Session.Enabled := True;
		Session.Plus := 'PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('PSS_DATEDEBUT')))+'"'+
		' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PSS_DATEFIN')))+'"'+
		' AND PSS_CODESTAGE="'+GetControlText('PSS_CODESTAGE')+'"';
		If PGBundleInscFormation Then Session.Plus := Session.Plus + ' AND (PSS_PREDEFINI="STD" OR (PSS_PREDEFINI="DOS" AND PSS_NODOSSIER="'+V_PGI.NoDossier+'"))'; //PT18
	end
	Else Session.enabled := False;
end;


procedure TOF_PGFEUILLEPRESENCEFOR.StageElipsisClick(Sender : TObject);
var  St,StWhere : String;
     Session : THValComboBox;
begin
	St := ' SELECT PST_CODESTAGE,PST_MILLESIME,PST_LIBELLE FROM STAGE ';
	StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="0000" AND'+
	' PST_CODESTAGE IN (SELECT PSS_CODESTAGE FROM SESSIONSTAGE WHERE PSS_DATEDEBUT>="'+UsDatetime(StrToDate(GetControlText('PSS_DATEDEBUT')))+'"'+//PT4
	' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PSS_DATEFIN')))+'")';
	If PGBundleInscFormation Then StWhere := StWhere + ' AND (PST_PREDEFINI="STD" OR (PST_PREDEFINI="DOS" AND PST_NODOSSIER="'+V_PGI.NoDossier+'"))'; //PT18

	LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE',StWhere,'', True,-1);
	Session := THValComboBox(GetControl('PSS_ORDRE'));
	If GetControlText('PSS_CODESTAGE') <> '' Then
	begin
		Session.Enabled := True;
		Session.Value := '';
		Session.Plus := 'PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('PSS_DATEDEBUT')))+'"'+
		' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PSS_DATEFIN')))+'"'+
		' AND PSS_CODESTAGE="'+GetControlText('PSS_CODESTAGE')+'"';
		If PGBundleInscFormation Then Session.Plus := Session.Plus + ' AND (PSS_PREDEFINI="STD" OR (PSS_PREDEFINI="DOS" AND PSS_NODOSSIER="'+V_PGI.NoDossier+'"))'; //PT18
	end
	Else
	begin
		Session.Enabled := False;
		Session.Value := '';
		Session.Plus := '';
	end;
end;

Procedure TOF_PGFEUILLEPRESENCEFOR.DateElipsisClick(Sender : TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGFEUILLEPRESENCEFOR.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+THValComboBox(Sender).Value+'"',True);
        If Not Q.Eof Then
        begin
                SetControlText('PSS_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PSS_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

{TOF_PGEDITSUIVIBUDFORM}
// Edition du suivi du budget

procedure TOF_PGEDITSUIVIBUDFORM.SalarieElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StWhere,StOrder,StListe : String;
{$ENDIF}
begin
	{$IFDEF EMANAGER}
    If typeUtilisat = 'R' then
      StWhere := 'PSE_CODESERVICE IN '+
      '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
      ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
      ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))'
    else
      StWhere := 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
      'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
    
    If GetControlText('RESPONSFOR') <> '' then
      StWhere := StWhere + ' AND PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'"';
    
    If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then
      StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
      
    StOrder := 'PSA_SALARIE';
    StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
    
    LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
         
    {$ELSE} //PT18
    //If PGBundleInscFormation Then //PT20
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
    {$ENDIF}
end;

procedure TOF_PGEDITSUIVIBUDFORM.ExitEdit(Sender: TObject);
var edit : thedit;
begin
	edit:=THEdit(Sender);
	if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
	if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
	edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDITSUIVIBUDFORM.RespElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StOrder,StWhere : String;
    St : String;
{$ENDIF}
begin
	{$IFDEF EMANAGER}
	If TypeUtilisat = 'R' then StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
		' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
	else StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
		'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
	StOrder := 'PSI_LIBELLE';
	LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
	{$ELSE}
	ElipsisResponsableMultidos (Sender); //PT18	
	{$ENDIF}
end;

procedure TOF_PGEDITSUIVIBUDFORM.OnUpdate ;
var StWhere,StOrder,StSelect : String;
    Pages : TPageControl;
    Rupt1,Rupt2,Rupt3,Rupt4 : String;
    {$IFDEF EMANAGER}MultiNiveau : Boolean;{$ENDIF}
begin
  Inherited ;
	Pages := TPageControl(GetControl('Pages'));
	StWhere := RecupWhereCritere(Pages);
	If StWhere <> '' then StWhere := StWhere+' AND PFI_CODESTAGE<>"--CURSUS--"'
	else StWhere := 'WHERE PFI_CODESTAGE<>"--CURSUS--"';
	
	{$IFDEF EMANAGER}
	MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
	If TypeUtilisat = 'R' then
	begin
		If MultiNiveau then
		begin
			If StWhere <> '' then StWhere := StWhere + ' AND (PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
			' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
			else StWhere := 'WHERE (PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
			'WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
		end
		else
		begin
			If StWhere <> '' then StWhere := StWhere + ' AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"'
			else StWhere := 'WHERE PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"';
		end;
	end
	else
	begin
		If StWhere <> '' then StWhere := StWhere + ' AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
		'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))'
		else StWhere := 'WHERE PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
		'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
	end;
	{$ENDIF}
	
	StSelect := 'SELECT PFI_HTPSTRAV,PFI_HTPSNONTRAV,PFI_TYPOFORMATION,PFI_NATUREFORM,PFI_TYPEPLANPREV,'+
	'PFI_CURSUS,PFI_NBINSC,PFI_CODESTAGE,PFI_SALARIE,PFI_LIBELLE,PFI_RESPONSFOR,PFI_ETABLISSEMENT,PFI_LIBEMPLOIFOR,PFI_ETATINSCFOR,PFI_LIEUFORM,'+
	'PFI_MOTIFETATINSC,PFI_MOTIFINSCFOR,PFI_FORMATION1,PFI_FORMATION2,PFI_FORMATION3,PFI_FORMATION4,PFI_FORMATION5,PFI_FORMATION6,PFI_FORMATION7,PFI_FORMATION8,'+
	'PST_NATUREFORM,PFI_COUTREELSAL,PFI_FRAISFORFAIT,PFI_AUTRECOUT,PFI_TRAVAILN1,PFI_TRAVAILN2,PFI_TRAVAILN3,PFI_TRAVAILN4 '+
	'FROM INSCFORMATION LEFT JOIN STAGE ON PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME=PFI_MILLESIME';
	Rupt1 := GetControlText('THVALRUPTURE1');
	Rupt2 := GetControlText('THVALRUPTURE2');
	Rupt3 := GetControlText('THVALRUPTURE3');
	Rupt4 := GetControlText('THVALRUPTURE4');
	SetControlText('XX_RUPTURE1',Rupt1);
	SetControlText('XX_RUPTURE2',Rupt2);
	SetControlText('XX_RUPTURE3',Rupt3);
	SetControlText('XX_RUPTURE4',Rupt4);
	StOrder := '';
	If Rupt1 <> '' then StOrder := 'ORDER BY '+Rupt1;
	If Rupt2 <> '' then StOrder := StOrder+','+Rupt2;
	If Rupt3 <> '' then StOrder := StOrder+','+Rupt3;
	If Rupt4 <> '' then StOrder := StOrder+','+Rupt4;

	//PT22
	// N'affiche pas les salariés confidentiels
	If StWhere <> '' Then StWhere := StWhere + ' AND ';
	If PGBundleHierarchie Then
		StWhere := StWhere + ' (PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))'  //PT24
	Else
		StWhere := StWhere + ' (PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))';  //PT24
		
	//PT18 - Début
	If PGBundleInscFormation Then
	Begin
		StWhere := StWhere + DossiersAInterroger(GetControlText('PFI_PREDEFINI'),GetControlText('NODOSSIER'),'PFI',False,True);
	End;
	//PT18 - Fin

	TFQRS1(Ecran).WhereSQL := StSelect +' '+StWhere+' '+StOrder;
end;

procedure TOF_PGEDITSUIVIBUDFORM.OnArgument (S : String ) ;
var Num,i : Integer;
    Combo : THValComboBox;
    Stage,Edit : THEdit;
    Q : TQuery;
    Arg : String;
begin
  Inherited ;
        Arg := S;
        SetControlCaption('LIBSTAGE','');
        Stage := THEdit(GetControl('PFI_CODESTAGE'));
        If Stage <> Nil Then Stage.OnElipsisClick := StageElipsisClick;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PFI_FORMATION'+IntToStr(Num)),GetControl ('TPFI_FORMATION'+IntToStr(Num)));
        end;
        SetControlText('PFI_MILLESIME',RendMillesimePrevisionnel);
        Ferme(Q);
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('THVALRUPTURE'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                        Combo.Items.Add('<<Aucun>>');
                        Combo.Values.Add('');
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('PFI_CODESTAGE');
                        if VH_Paie.PGResponsables = True then
                        begin
                                Combo.Items.Add('Responsable formation');
                                Combo.Values.Add('PFI_RESPONSFOR');
                        end;
                        Combo.Items.Add('Type de plan');
                        Combo.Values.Add('PFI_TYPEPLANPREV');
                        Combo.Items.Add('Etablissement');
                        Combo.Values.Add('PFI_ETABLISSEMENT');
                        Combo.Items.Add('Libellé emploi');
                        Combo.Values.Add('PFI_LIBEMPLOIFOR');
                        Combo.Items.Add('Lieu formation');
                        Combo.Values.Add('PFI_LIEUFORM');
                        Combo.Items.Add('Type (Interne/Externe)');
                        Combo.Values.Add('PST_NATUREFORM');
                        Combo.Items.Add('Motif inscription');
                        Combo.Values.Add('PFI_MOTIFINSCFOR');
                        Combo.Items.Add('Etat inscription');
                        Combo.Values.Add('PFI_ETATINSCFOR');
                        Combo.Items.Add('Motif décision');
                        Combo.Values.Add('PFI_MOTIFETATINSC');
                        For Num  :=  1 to VH_Paie.NBFormationLibre do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PFI_FORMATION1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PFI_FORMATION2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PFI_FORMATION3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PFI_FORMATION4');end;
                                If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PFI_FORMATION5');end;
                                If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PFI_FORMATION6');end;
                                If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PFI_FORMATION7');end;
                                If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PFI_FORMATION8');end;
                        end;
                        For Num := 1 to VH_Paie.PGNbreStatOrg do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('PFI_TRAVAILN1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('PFI_TRAVAILN2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('PFI_TRAVAILN3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('PFI_TRAVAILN4');end;
                        end;
                end;
        end;
        SetControlProperty('THVALRUPTURE1','Value','');
        SetControlProperty('THVALRUPTURE2','Value','');
        SetControlProperty('THVALRUPTURE3','Value','');
        SetControlProperty('THVALRUPTURE4','Value','');
        SetControlText('PFI_ETATINSCFOR','');
        SetControlText('PFI_MOTIFINSCFOR','');
        SetControlText('PFI_MOTIFETATINSC','');
        Combo := THValComboBox(GetControl('PFI_ETATINSCFOR'));
        SetControlCaption('LIBRESP','');
        If Combo <> Nil then Combo.OnExit := ExitMotifInsc;
        {$IFNDEF EMANAGER}
        if VH_Paie.PGResponsables = False then
        begin
                SetControlVisible ('PFI_RESPONSFOR',False);
                SetControlVisible('LIBRESP',False);
                SetControlVisible('TRESPONSFOR',False);
        end;
        {$ENDIF}
    For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFI_TRAVAILN'+IntToStr(Num)),GetControl ('TPFI_TRAVAILN'+IntToStr(Num)));
    end;

        {$IFDEF EMANAGER}
        SetControlVisible('CMULTINIVEAU',True);
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        Ecran.Caption := 'Edition du prévisionnel';
        UpdateCaption(Ecran);
//        SetControlEnabled('PFI_MILLESIME',False);
        SetControlProperty('Option','TabVisible',False);
    {    If MultiNiveau = False then
        begin
                SetControlVisible('PFI_RESPONSFOR',False);
                SetControlVisible('TPFI_RESPONSFOR',False);
                SetControlVisible('LIBRESP',False);
        end;}

        SetControlVisible('PRUPTURE',False);
        SetControlProperty('PRUPTURE','TabVisible',False);
        SetControlText('THVALRUPTURE1','PFI_RESPONSFOR');
        SetControlText('THVALRUPTURE2','PST_NATUREFORM');
        SetControlText('THVALRUPTURE3','PFI_ETATINSCFOR');
        SetControlText('THVALRUPTURE4','PFI_CODESTAGE');
       {$ENDIF}
       	//PT3
		If PGBundleHierarchie Then //PT20
		Begin
	        Edit := THEdit(getControl('PFI_RESPONSFOR'));
	        If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick;
	   	End;
        
        Edit := THEdit(getControl('PFI_SALARIE'));
        If Edit <> Nil then
        Begin
        	{$IFNDEF EMANAGER}If PGBundleHierarchie Then{$ENDIF} //PT20
            	Edit.OnElipsisClick := SalarieElipsisClick;
            Edit.OnExit := ExitEdit;
        End;
       
	//PT18 - Début
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFI_PREDEFINI',False);
			SetControlText   ('PFI_PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT19
	end
	else
	begin
		SetControlVisible('PFI_PREDEFINI', False);
		SetControlVisible('TPFI_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT18 - Fin
end ;

procedure TOF_PGEDITSUIVIBUDFORM.StageElipsisClick(Sender : TObject);
var StWhere : String;
begin
        StWhere :=  'PST_ACTIF="X" AND PST_MILLESIME="'+GetControlText('PFI_MILLESIME')+'" ';
		//PT18 - Début
        If PGBundleInscFormation Then
        Begin
        	If Not PGDroitMultiForm Then 
        		StWhere := StWhere + DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True) //PT18
        	Else
        		StWhere := StWhere + DossiersAInterroger('','','PST',True,True); //PT18
        End;
		//PT18 - Fin
        {$IFDEF EAGLCLIENT}
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,'', True,-1);
        {$ELSE}
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,'', True,-1);
        {$ENDIF}
end;

procedure TOF_PGEDITSUIVIBUDFORM.ExitMotifInsc(Sender : TObject);
begin
        If GetControlText('PFI_ETATINSCFOR') = '' then SetControlProperty('PFI_MOTIFETATINSC','Plus','');
        If GetControlText('PFI_ETATINSCFOR') = 'ATT' then SetControlProperty('PFI_MOTIFETATINSC','Plus','');
        If GetControlText('PFI_ETATINSCFOR') = 'NAN' then SetControlProperty('PFI_MOTIFETATINSC','Plus','');
        If GetControlText('PFI_ETATINSCFOR') = 'REF' then SetControlProperty('PFI_MOTIFETATINSC','Plus',' AND CC_LIBRE="REF"');
        If GetControlText('PFI_ETATINSCFOR') = 'REP' then SetControlProperty('PFI_MOTIFETATINSC','Plus',' AND CC_LIBRE="REP"');
        If GetControlText('PFI_ETATINSCFOR') = 'VAL' then SetControlProperty('PFI_MOTIFETATINSC','Plus',' AND CC_LIBRE="VAL"');
end;

{TOF_PGEDITFRAISSAL}
{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 30/04/2007  / PT7
Modifié le ... :   /  /    
Description .. : Recherche de la tablette associée à un champ
Mots clefs ... : 
*****************************************************************}
Function TOF_PGEDITFRAISSAL.RechTablette (NomChamp : String) : String;
Var
  Champ,Valeur : String;
  Q : TQuery;
  i : Integer;
Begin
  Valeur := '';
  If (NomChamp <> '') And (Length(NomChamp) > 5) Then
  Begin
    Champ := Copy(NomChamp,5,255);

    // Afin de limiter les accès en base, on sauvegarde les résultats précédents de recherche de tablette
    If (Length(Tablettes) > 0) Then
    Begin
          For i:=0 To Length(Tablettes)-1 Do
          Begin
                If (Tablettes[i][0] = Champ) Then Begin Valeur := Tablettes[i][1]; Break; End;
          End;
    End;

    // Recherche de la tablette concernée
    If (Valeur = '') Then
    Begin
          Q := OpenSql('SELECT DO_COMBO FROM DECOMBOS WHERE DO_NOMCHAMP like "%' + Champ + '%" ', True);
          Valeur := Q.FindField('DO_COMBO').AsString;
          Ferme(Q);
          // Sauvegarde de la valeur lue
          SetLength(Tablettes, Length(Tablettes)+1);
          Tablettes[Length(Tablettes)-1][0] := Champ;
          Tablettes[Length(Tablettes)-1][1] := Valeur;
    End;
  End;
  Result := Valeur;
End;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 30/04/2007 / PT7
Description .. : Edition des frais des salariés a une formation
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITFRAISSAL.OnUpdate;
var OrderBy,Where,Select : String;
    i : Integer;
    Pages : TPageControl;
    AvecLibSession,ForceCodeStage : Boolean;
    //PT7 - Début
    Q : TQuery;
    TobFraisFor, T : TOB;
    Salarie,Stage,Millesime,Ordre : String;
    Tablette, Requete : String;
    code : Integer;
    LesFrais: array[1..15] of string;
    //PT7 - Fin
begin
Inherited ;
        TobEtat := Tob.Create('Edition',Nil,-1);

        ForceCodeStage := True;
        AvecLibSession := False;

        // Construction dynamique du SELECT pour les ruptures
        For i := 1 to 4 do
        begin
                SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
                If GetControlText('VRUPT'+IntToStr(i)) <> '' then
                begin
                        If GetControlText('VRUPT'+IntToStr(i)) = SuffixeFor+'_ORDRE' Then AvecLibSession := True;

                        If OrderBy <> '' then OrderBy := OrderBy + ',' + GetControlText('VRUPT'+IntToStr(i))
                        else OrderBy := ' ORDER BY '+ GetControlText('VRUPT'+IntToStr(i));

                        If (GetControlText('VRUPT'+IntToStr(i)) <> SuffixeFor+'_CODESTAGE') And (GetControlText('VRUPT'+IntToStr(i)) <> SuffixeFor+'_ORDRE') Then
                             Select := Select+','+GetControlText('VRUPT'+IntToStr(i));
                end;
        end;

        If AvecLibSession = True then
        Begin
                For i := 1 to 4 do
                Begin
                        If (GetControlText('VRUPT'+IntToStr(i)) = SuffixeFor+'_CODESTAGE') Or 
                           (GetControlText('VRUPT'+IntToStr(i)) = SuffixeFor+'_ORDRE') Then
                             ForceCodeStage := False;
                End;
        End
        Else ForceCodeStage := False;

        // Construction dynamique du ORDER BY pour les ruptures
        If ForceCodeStage = True then
        begin
                If OrderBy <> '' then OrderBy := OrderBy + ','
                Else OrderBy :=  ' ORDER BY ';

                OrderBy := OrderBy + 'PFO_CODESTAGE,PFO_ORDRE';
                SetControltext('XX_RUPTURE5',SuffixeFor+'_CODESTAGE');
        End
        Else SetControlText('XX_RUPTURE5','');
        If OrderBy <> '' Then OrderBy := OrderBy + ',PFS_FRAISSALFOR'
        Else OrderBy := 'ORDER BY FRAISSALFOR';

        Pages := TPageControl(GetControl('PAGES'));
        Where := RecupWhereCritere(Pages);

        //PT22
	    // N'affiche pas les salariés confidentiels
	    If Where <> '' Then Where := Where + ' AND ';
	    If PGBundleHierarchie Then
	    	Where := Where + 'PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
	    Else
	    	Where := Where + 'PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
			    
        // Récupération des types de frais possibles
        //PT18
        If PGBundleInscFormation Then
        	Q := OpenSQL('SELECT CC_CODE FROM '+GetBase(GetBasePartage(BUNDLE_FORMATION),'CHOIXCOD')+' WHERE CC_TYPE="PFA"', True)
        Else
        	Q := OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PFA"', True);
        i := 1;
        While not Q.Eof do
        Begin
               LesFrais[i] := Q.FindField('CC_CODE').AsString;
               i := i + 1;
               Q.Next;
        End;
        Ferme(Q);

        // Récupération des libellés de colonnes
        For i := 1 To VH_PAIE.PGFNbFraisLibre Do   //PT18
        Begin
               If LesFrais[i] <> '' Then
                     SetControlText('LIBELLE'+IntToStr(i),RechDom('PGFRAISSALFORM', LesFrais[i], False))
               Else
                     SetControlText('LIBELLE'+IntToStr(i),'');
        End;

        { Construction de la requête }

        Requete := 'SELECT PFO_SALARIE,PFO_CODESTAGE,PSS_LIBELLE,PFO_MILLESIME,PFO_ORDRE,PFO_DATEDEBUT,PFO_DATEFIN,PFS_FRAISSALFOR,PFS_MONTANT,PFS_MONTANTPLAF' + Select + ' FROM FRAISSALFORM ';

        // Spécificités des frais salariés
        If (SuffixeFor = PfxFOR) Then
        Begin
                Requete := Requete + ' LEFT JOIN FORMATIONS ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_SALARIE=PFO_SALARIE AND PFS_CODESTAGE=PFO_CODESTAGE AND PFS_ORDRE=PFO_ORDRE AND PFS_MILLESIME=PFO_MILLESIME ';//PT9
        End
        // Spécificités des frais par animateurs
        Else If (SuffixeFor = PfxANI) Then
        Begin
             If PGBundleHierarchie Then //PT13
               	Requete := Requete + ' LEFT JOIN SESSIONANIMAT ON PFS_SALARIE=PAN_SALARIE LEFT JOIN INTERIMAIRES ON PAN_SALARIE=PSI_INTERIMAIRE '
             Else
                Requete := Requete + ' LEFT JOIN SESSIONANIMAT ON PFS_SALARIE=PAN_SALARIE LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE ';
        End;
        Requete := Requete + ' LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME '; // Pour l'onglet Formation
        Requete := Requete + ' LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME ';
        Requete := Requete + Where;
        
		//PT18 - Début
        If PGBundleInscFormation Then
        Begin
			Requete := Requete + DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PST',False,True);
        End;
		//PT18 - Fin

        Requete := Requete + OrderBy;

        // Conversion des préfixes pour adapter si besoin aux animateurs
        If (SuffixeFor <> PfxFOR) Then Requete := ConvertPrefixe(Requete,PfxFOR,SuffixeFor);

        // Création de la TOB
        Q := OpenSQL(Requete, True);
        TobFraisFor := Tob.Create('LesFrais', nil, -1);
        TobFraisFor.LoadDetailDB('LesFrais', '', '', Q, False);
        Ferme(Q);

        // Intégration des frais en une seule ligne dans la TOB finale
        For i := 0 To TobFraisFor.Detail.Count-1 Do
        Begin
                // Clé de l'inscription
                Salarie := TobFraisFor.Detail[i].GetValue(SuffixeFor+'_SALARIE');
                Stage := TobFraisFor.Detail[i].GetValue(SuffixeFor+'_CODESTAGE');
                Millesime := TobFraisFor.Detail[i].GetValue(SuffixeFor+'_MILLESIME');
                Ordre := TobFraisFor.Detail[i].GetValue(SuffixeFor+'_ORDRE');

                T := TobEtat.FindFirst([SuffixeFor+'_SALARIE',SuffixeFor+'_CODESTAGE',SuffixeFor+'_MILLESIME',SuffixeFor+'_ORDRE'],[Salarie,Stage,Millesime,Ordre],True);
                If T = Nil Then
                Begin
                      // La ligne n'existe pas : on la crée
                      T := Tob.Create('LaFille', TobEtat, -1);
                      T.AddChampSupValeur('PFO_SALARIE',Salarie);
                      T.AddChampSupValeur('PFO_CODESTAGE',Stage);
                      T.AddChampSupValeur('PFO_MILLESIME',Millesime);
                      T.AddChampSupValeur('PFO_ORDRE',Ordre);
                      T.AddChampSupValeur('PSS_LIBELLE',TobFraisFor.Detail[i].GetValue('PSS_LIBELLE'));
                      T.AddChampSupValeur('PFO_DATEDEBUT',TobFraisFor.Detail[i].GetValue(SuffixeFor+'_DATEDEBUT'));
                      T.AddChampSupValeur('PFO_DATEFIN',TobFraisFor.Detail[i].GetValue(SuffixeFor+'_DATEFIN'));

                      // Ajout des ruptures
                      T.AddChampSupValeur('RUPTURE1',''); T.AddChampSupValeur('LIBRUPTURE1','');
                      T.AddChampSupValeur('RUPTURE2',''); T.AddChampSupValeur('LIBRUPTURE2','');
                      T.AddChampSupValeur('RUPTURE3',''); T.AddChampSupValeur('LIBRUPTURE3','');
                      T.AddChampSupValeur('RUPTURE4',''); T.AddChampSupValeur('LIBRUPTURE4','');
                      T.AddChampSupValeur('RUPTURE5',''); T.AddChampSupValeur('LIBRUPTURE5','');

                      If GetControlText('XX_RUPTURE1') <> '' Then
                      Begin
                          T.PutValue('RUPTURE1',TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE1')));
                          Tablette := RechTablette(GetControlText('XX_RUPTURE1'));
                          If (Tablette <> '') Then T.PutValue('LIBRUPTURE1',RechDom(Tablette, TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE1')), False));
                      End;
                      If GetControlText('XX_RUPTURE2') <> '' Then
                      Begin
                          T.PutValue('RUPTURE2',TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE2')));
                          Tablette := RechTablette(GetControlText('XX_RUPTURE2'));
                          If (Tablette <> '') Then T.PutValue('LIBRUPTURE2',RechDom(Tablette, TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE2')), False));
                      End;
                      If GetControlText('XX_RUPTURE3') <> '' Then
                      Begin
                          T.PutValue('RUPTURE3',TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE3')));
                          Tablette := RechTablette(GetControlText('XX_RUPTURE3'));
                          If (Tablette <> '') Then T.PutValue('LIBRUPTURE3',RechDom(Tablette, TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE3')), False));
                      End;
                      If GetControlText('XX_RUPTURE4') <> '' Then
                      Begin
                          T.PutValue('RUPTURE4',TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE4')));
                          Tablette := RechTablette(GetControlText('XX_RUPTURE4'));
                          If (Tablette <> '') Then T.PutValue('LIBRUPTURE4',RechDom(Tablette, TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE4')), False));
                      End;
                      If GetControlText('XX_RUPTURE5') <> '' Then
                      Begin
                          T.PutValue('RUPTURE5',TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE5')));
                          Tablette := RechTablette(GetControlText('XX_RUPTURE5'));
                          If (Tablette <> '') Then T.PutValue('LIBRUPTURE5',RechDom(Tablette, TobFraisFor.Detail[i].GetValue(GetControlText('XX_RUPTURE5')), False));
                      End;

                      // Création à vide de tous les champs correspondant aux frais
                      T.AddChampSupValeur('MTT1',0);T.AddChampSupValeur('MTTPLAF1',0);
                      T.AddChampSupValeur('MTT2',0);T.AddChampSupValeur('MTTPLAF2',0);
                      T.AddChampSupValeur('MTT3',0);T.AddChampSupValeur('MTTPLAF3',0);
                      T.AddChampSupValeur('MTT4',0);T.AddChampSupValeur('MTTPLAF4',0);
                      T.AddChampSupValeur('MTT5',0);T.AddChampSupValeur('MTTPLAF5',0);
                      T.AddChampSupValeur('MTT6',0);T.AddChampSupValeur('MTTPLAF6',0);
                      T.AddChampSupValeur('MTT7',0);T.AddChampSupValeur('MTTPLAF7',0);
                      T.AddChampSupValeur('MTT8',0);T.AddChampSupValeur('MTTPLAF8',0);
                      T.AddChampSupValeur('MTT9',0);T.AddChampSupValeur('MTTPLAF9',0);
                      T.AddChampSupValeur('MTT10',0);T.AddChampSupValeur('MTTPLAF10',0);
                      T.AddChampSupValeur('MTT11',0);T.AddChampSupValeur('MTTPLAF11',0);
                      T.AddChampSupValeur('MTT12',0);T.AddChampSupValeur('MTTPLAF12',0);
                      T.AddChampSupValeur('MTT13',0);T.AddChampSupValeur('MTTPLAF13',0);
                      T.AddChampSupValeur('MTT14',0);T.AddChampSupValeur('MTTPLAF14',0);
                      T.AddChampSupValeur('MTT15',0);T.AddChampSupValeur('MTTPLAF15',0);
                End;
			    
                // Recherche de la colonne à mettre à jour
                code := 1;
                While code <= Length(LesFrais) do
                Begin
                      If LesFrais[code] = TobFraisFor.Detail[i].GetValue('PFS_FRAISSALFOR') Then Break
                      Else code := code + 1;
                End;

                // Mise à jour des valeurs de la colonne
                If (code <= Length(LesFrais)) And (Code <= VH_PAIE.PGFNbFraisLibre) Then //PT18
                Begin
                      T.PutValue('MTT'+IntToStr(code),TobFraisFor.Detail[i].GetValue('PFS_MONTANT'));
                      T.PutValue('MTTPLAF'+IntToStr(code),TobFraisFor.Detail[i].GetValue('PFS_MONTANTPLAF'));
                End;
        End;

        // Libération de la TOB intermédiaire
        FreeAndNil (TobFraisFor);

        // Attribution de la TOB à l'état
        TFQRS1(Ecran).LaTob:= TobEtat;
end;

//PT7 - Début
{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 27/04/2007
Modifié le ... :   /  /    
Description .. : Libérations à la fermeture de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITFRAISSAL.OnClose;
begin
Inherited ;
        FreeAndNil(TobEtat);
        SetLength(Tablettes,0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 27/04/2007
Modifié le ... :   /  /    
Description .. : Rend (in)visible la checkbox du niveau de rupture
Suite ........ : précédent en fonction de la valeur saisie. Effectue également
Suite ........ : le contrôle de cohérence des ruptures.
Mots clefs ... : 
*****************************************************************}
Procedure TOF_PGEDITFRAISSAL.OnChangeRuptures (Sender : TObject);
Var
  i,j : Integer;
  Affiche : boolean;
  Valeur : String;
Begin
     // Récupération du niveau de rupture
     i := StrToInt(Copy(TControl(Sender).Name,6,1));
     // Détermine s'il faut cacher ou afficher la checkbox de saut de page
     Affiche := (THValComboBox(GetControl('VRUPT'+IntToStr(i))).Value <> '');

     // Contrôle de cohérence des ruptures
     For j:= 2 To 3 Do
     Begin
        If (THValComboBox(GetControl('VRUPT'+IntToStr(j))).Value = '') And
           (THValComboBox(GetControl('VRUPT'+IntToStr(j+1))).Value <> '') Then
          Begin
               PGIBox('Le niveau de rupture '+IntToStr(j)+' doit être renseigné',Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End
     End;

     Valeur := THValComboBox(GetControl('VRUPT'+IntToStr(i))).Value;
     For j:=1 To 4 Do
     Begin
          If (i <> j) And (Valeur = THValComboBox(GetControl('VRUPT'+IntToStr(j))).Value) And (Valeur <> '') Then
          Begin
               PGIBox('La rupture '+IntToStr(i)+' doit être différente de la '+IntToStr(j),Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End;
     End;

     GetControl('BValider').Enabled := True;

     // Afficher/cacher la checkbox
     //If (i>1) Then //PT16
     Begin
          TCheckBox(GetControl('CSAUT'+IntToStr(i))).Visible := Affiche; //PT16
          // Si on cache, on remet également l'état à "décoché"
          If (Affiche = False) Then TCheckBox(GetControl('CSAUT'+IntToStr(i))).Checked := Affiche; //PT16
     End;
End;

procedure TOF_PGEDITFRAISSAL.OnArgument (S : String ) ;
var DD,DF : TDateTime;
    Millesime,Combo : THVaLComboBox;
    Stage,DateDebut,DateFin,Edit : THEdit;
    Num,i : Integer;
    CkBox : TCheckBox;
begin
  Inherited ;
	SetControlVisible('TYPEMONTANT',False);
	SetControlVisible('TTYPEMONTANT',False);
	TypeEdition := ReadTokenPipe(S,';');

	If TypeEdition = 'FRAISANIM' then
	begin
		TFQRS1(Ecran).Caption := 'Edition des frais des animateurs de formation';
		//PT7 - Début
        If PGBundleHierarchie Then //PT18
            SuffixeSal := PfxINT
        Else
    		SuffixeSal := PfxSAL;
		SuffixeFor := PfxANI;
		SetControlText('QUI','animateur');
		SetControlVisible('TPFO_RESPONSFOR', False);
		SetControlVisible('PFO_RESPONSFOR', False);
		//PT7 - Fin
	end
	else
	begin   //CONSULTFRAIS
		TFQRS1(Ecran).Caption := 'Edition des frais des stagiaires';
		//PT7 - Début
		SuffixeSal := PfxFOR;
		SuffixeFor := PfxFOR;
		SetControlText('QUI','stagiaire');
		//PT7 - Fin
	end;

	// Adaptation des ruptures
	For i := 1 To 4 do
	begin
		Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
		If Combo <> Nil Then
		begin
			//PT7 - Début
			// Empêcher la suppression de la rupture principale
			If i > 1 Then
			Begin
				Combo.Items.Add('<<Aucun>>');
				Combo.Values.Add('');
			End;

			// Rendre invisible toutes les check-boxes de saut de page par défaut
			CkBox := TCheckBox(GetControl('CSAUT'+IntToStr(i)));
			CkBox.Visible := False;
			// Sur le changement de valeur d'une combo, il faut afficher le saut de page précédent
			Combo.OnChange := OnChangeRuptures;
			//PT7 - Fin

			Combo.Items.Add('Formation');
			Combo.Values.Add(SuffixeFor+'_CODESTAGE');
			Combo.Items.Add('Session');
			Combo.Values.Add(SuffixeFor+'_ORDRE');
			if (VH_Paie.PGResponsables = True) And (TypeEdition = 'CONSULTFRAIS') then //PT7
			begin
				Combo.Items.Add('Responsable formation');
				Combo.Values.Add('PFO_RESPONSFOR');
			end;
			Combo.Items.Add('Etablissement');
			Combo.Values.Add(SuffixeSal+'_ETABLISSEMENT');
			Combo.Items.Add('Libellé emploi');
			//PT7 - Début
			If TypeEdition = 'FRAISANIM' then
			Begin
//				If PGBundleInscFormation Then //PT13
				Combo.Values.Add(SuffixeSal+'_LIBELLEEMPLOI')
//				Else
//				Combo.Values.Add('PSA_LIBELLEEMPLOI')
			End
			Else
				Combo.Values.Add('PFO_LIBEMPLOIFOR');

			If TypeEdition = 'CONSULTFRAIS' then
			Begin
				Combo.Items.Add('Type (Interne/Externe)');
				Combo.Values.Add(SuffixeFor+'_NATUREFORM');

				For Num  :=  1 to VH_Paie.NBFormationLibre do
				begin
					If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PFO_FORMATION1');end;
					If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PFO_FORMATION2');end;
					If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PFO_FORMATION3');end;
					If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PFO_FORMATION4');end;
					If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PFO_FORMATION5');end;
					If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PFO_FORMATION6');end;
					If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PFO_FORMATION7');end;
					If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PFO_FORMATION8');end;
				end;
			End;
			//PT7 - Fin
			For Num := 1 to VH_Paie.PGNbreStatOrg do
			begin
				If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add(SuffixeSal+'_TRAVAILN1');end;
				If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add(SuffixeSal+'_TRAVAILN2');end;
				If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add(SuffixeSal+'_TRAVAILN3');end;
				If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add(SuffixeSal+'_TRAVAILN4');end;
			end;
		end;
	end;

	UpdateCaption(TFQRS1(Ecran)) ;
	SetControlCaption('LIBSTAGE','');
	SetcontrolCaption('LIBRESP','');
	Millesime := THValComboBox(GetControl('MILLESIME'));
	Stage := THEdit(GetControl('PFO_CODESTAGE'));
	DateDebut := THEdit(GetControl('PFO_DATEDEBUT'));
	DateFin := THEdit(GetControl('PFO_DATEFIN'));
	Millesime.Value := RendMillesimeRealise(DD,DF);
	DateDebut.Text := DateToStr(DD);
	DateFin.Text := DateToStr(DF);
	If DateFin <> Nil Then DateFin.OnElipsisClick := DateElipsisClick;
	If DateDebut <> Nil Then DateDebut.OnElipsisClick := DateElipsisClick;
	If Stage <> Nil then
	begin
		Stage.OnElipsisClick := StageElipsisClick;
		Stage.OnChange := StageChange;
	end;
	If Millesime <> Nil Then Millesime.OnChange := ChangeMillesime;
	For Num  :=  1 to VH_Paie.NBFormationLibre do
	begin
		if Num >8 then Break;
		VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
	end;
	{$IFDEF EMANAGER}
	If TypeEdition = 'CONSULTFRAIS' then SetControlText('PFO_RESPONSFOR',V_PGI.UserSalarie); //PT7
	{$ENDIF}
	
	//PT18
	{$IFNDEF EMANAGER}
	If PGBundleHierarchie Then //PT20
    Begin
    {$ENDIF}
		Edit := THEdit(GetControl('PFO_SALARIE'));
		If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
		Edit := THEdit(GetControl('PFO_SALARIE_'));
		If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
		Edit := THEdit(GetControl('PFO_RESPONSFOR'));
		If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick;
	{$IFNDEF EMANAGER}
	End;
	{$ENDIF}
	//PT18

	// Adaptations particulières pour les animateurs
	If TypeEdition = 'FRAISANIM' then
	begin
		SetControlProperty('PFO_ETABLISSEMENT','Name',SuffixeSal+'_ETABLISSEMENT');
		SetControlProperty('PFO_LIBEMPLOIFOR','Name',SuffixeSal+'_LIBELLEEMPLOI');
		For Num := 1 to 4 do
		begin
			SetControlProperty('PFO_TRAVAILN'+IntToStr(Num),'Name',SuffixeSal+'_TRAVAILN'+IntToStr(Num));
		end;
	end;
	SetControlText('VRUPT1',SuffixeFor+'_CODESTAGE'); //PT7

	//PT18 - Début
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFO_PREDEFINI',False);
			SetControlText   ('PFO_PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT19
	end
	else
	begin
		SetControlVisible('PFO_PREDEFINI', False);
		SetControlVisible('TPFO_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT18 - Fin
end ;

procedure TOF_PGEDITFRAISSAL.StageChange(Sender : TObject);
var Session : THValComboBox;
begin
	Session := THValComboBox(GetControl('PFO_ORDRE'));
	If GetControlText('PFO_CODESTAGE') <> '' Then
	begin
		Session.Enabled := True;
		Session.Plus := 'PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('PFO_DATEDEBUT')))+'"'+  //PT7
		' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PFO_DATEFIN')))+'"'+
		' AND PSS_CODESTAGE="'+GetControlText('PFO_CODESTAGE')+'"';

		//PT18
		Session.Plus := Session.Plus + DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PSS',True,True);
	end
	Else
	begin
		Session.enabled := False;
		Session.Value := '';
	end;
end;

procedure TOF_PGEDITFRAISSAL.StageElipsisClick;
var  StWhere : String;
begin
        StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="0000" AND'+
        ' PST_CODESTAGE IN (SELECT PSS_CODESTAGE FROM SESSIONSTAGE WHERE PSS_DATEDEBUT>="'+UsDatetime(StrToDate(GetControlText('PFO_DATEDEBUT')))+'"'+ //PT7
        ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PFO_DATEFIN')))+'")';
        
		//PT18
		StWhere := StWhere + DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PST',True,True);
        
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWHere,'', True,-1);
end;

Procedure TOF_PGEDITFRAISSAL.DateElipsisClick;
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITFRAISSAL.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+THValComboBox(Sender).Value+'"',True);
        If Not Q.Eof Then
        begin
                SetControlText('PFO_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PFO_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

procedure TOF_PGEDITFRAISSAL.SalarieElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var St,StWhere,StOrder : String;
{$ENDIF}
begin
		{$IFDEF EMANAGER}
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';

        (*If PGBundleInscFormation Then //PT13
        Begin
             St := ' SELECT PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES'+
             ' LEFT JOIN DEPORTSAL ON PSI_INTERIMAIRE=PSE_SALARIE ';
             StOrder := ' PSI_INTERIMAIRE';
             StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
             LookupList(THEdit(Sender),'Liste des stages','INTERIMAIRES LEFT JOIN DEPORTSAL ON PSI_INTERIMAIRE=PSE_SALARIE','PSI_INTERIMAIRE','PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
        End
        Else
        Begin*)
             St := ' SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES'+
             ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE ';
             StOrder := ' PSA_SALARIE';
             StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
             LookupList(THEdit(Sender),'Liste des stages','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
        (*End;*)
        {$ELSE}
    //If PGBundleInscFormation Then //PT20
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
        {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT18
Modifié le ... :   /  /
Description .. : Clic elipsis responsable
Mots clefs ... :
*****************************************************************}
Procedure TOF_PGEDITFRAISSAL.RespElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}var StWhere : String;{$ENDIF}
begin
        {$IFDEF EMANAGER}
        StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
        LookupList(THEdit(Sender),'Liste des responsables de formation','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWHere,'', True,-1);
        {$ELSE}
        ElipsisResponsableMultidos (Sender); //PT18
        {$ENDIF}
end;

{TOF_PGEDITBUDGETFOR
Edition du budget}

procedure TOF_PGEDITBUDGETFOR.SalarieElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StWhere,StOrder : String;
{$ENDIF}
begin
	{$IFDEF EMANAGER}
        If GetControltext('PFI_RESPONSFOR') <> '' then StWhere := 'PSE_RESPONSFOR="'+GetcontrolText('PFI_RESPONSFOR')+'"'
        else
            StWhere := 'PSE_CODESERVICE IN '+
            '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
            ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
        StOrder := 'PSA_LIBELLE';
        LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
    {$ELSE}    
    //If PGBundleInscFormation Then //PT20
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
    {$ENDIF}
end;

procedure TOF_PGEDITBUDGETFOR.OnUpdate;
var i : Integer;
    Combo : THValComboBox;
    SQL,Where : String;
    Pages : TPageControl;
begin
  Inherited ;
	Pages := TPageControl(GetControl('Pages'));
	Where := RecupWhereCritere(Pages);
	
	//PT18 - Début
	If PGBundleHierarchie Then
	Begin
		Where := Where + DossiersAInterroger(GetControlText('PFI_PREDEFINI'),GetControlText('NODOSSIER'),'PFI',False,True); //PT23
	End;
	//PT18 - Fin
	
	OrderBy := '';
	
	If Arg = 'INDIVIDUEL' then
	begin
		If GetCheckBoxState('SYNTHETIQUE') = CbChecked then
		begin
			If (GetCheckBoxState('CAFFICHECOUTP') = CbChecked) or (GetCheckBoxState('CAFFICHECOUTS') = CbChecked) then
			begin
				TFQRS1(Ecran).CodeEtat := 'PPX';
				TFQRS1(Ecran).FCodeEtat := 'PPX';
			end
			else
			begin
				TFQRS1(Ecran).CodeEtat := 'PPY';
				TFQRS1(Ecran).FCodeEtat := 'PPY';
			end;
		end
		else
		begin
			TFQRS1(Ecran).CodeEtat := 'PCG';
			TFQRS1(Ecran).FCodeEtat := 'PCG';
		end;
		{$IFDEF EMANAGER}
		If MultiNiveau then
		begin
			If Where <> '' then Where := Where + 'AND (PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
				' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
			else Where := '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
				' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
		end;
		{$ENDIf}
		OrderBy := 'ORDER BY PFI_SALARIE,PFI_CODESTAGE';
        
	    //PT22
	    // N'affiche pas les salariés confidentiels
	    If Where <> '' Then Where := Where + ' AND ';
	    If PGBundleHierarchie Then
	    	Where := Where + '(PSI_CONFIDENTIEL IS NULL OR PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'  //PT24
	    Else
	    	Where := Where + '(PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';  //PT24
        
		SQL := 'SELECT PFI_TYPOFORMATION,PFI_TYPEPLANPREV,PFI_DATEDIF,PFI_HTPSTRAV,PFI_HTPSNONTRAV,PFI_NATUREFORM,'+
			   'PFI_CURSUS,PFI_CODESTAGE,PFI_MILLESIME,PFI_SALARIE,PFI_LIBELLE,SUM(PFI_NBINSC) PFI_NBINSC,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,'+
		 	   'PFI_RESPONSFOR,SUM(PFI_COUTREELSAL) PFI_COUTREELSAL,SUM(PFI_FRAISFORFAIT) PFI_FRAISFORFAIT,SUM(PFI_DUREESTAGE) PFI_DUREESTAGE,PFI_LIEUFORM';
		
		If PGBundleHierarchie Then //PT13
			SQL := SQL + ',PSI_DATEENTREE'
		Else
			SQL := SQL + ',PSA_DATEENTREE';
		
		SQL := SQL + ',PST_LIBELLE,PST_NATUREFORM,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
					 ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_DUREESTAGE,PFI_TRAVAILN1,PFI_TRAVAILN2,PFI_TRAVAILN3,PFI_TRAVAILN4'+
					 ' FROM INSCFORMATION ';
		
		If PGBundleHierarchie Then //PT13
			SQL := SQL + 'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFI_SALARIE '
		Else
			SQL := SQL + 'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE ';
		
		SQL := SQL + 'LEFT JOIN STAGE '+
					 'ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+Where+
					 ' GROUP BY PFI_TYPOFORMATION,PFI_TYPEPLANPREV,PFI_DATEDIF,PFI_HTPSTRAV,PFI_HTPSNONTRAV,PFI_NATUREFORM,PFI_CURSUS,PFI_CODESTAGE,PFI_MILLESIME,PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,'+
					 'PFI_RESPONSFOR,PFI_LIEUFORM';
		
		If PGBundleHierarchie Then //PT13
			SQL := SQL + ',PSI_DATEENTREE'
		Else
			SQL := SQL + ',PSA_DATEENTREE';
		
		SQL := SQL + ',PST_LIBELLE,PST_NATUREFORM,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
					 ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_DUREESTAGE,PFI_TRAVAILN1,PFI_TRAVAILN2,PFI_TRAVAILN3,PFI_TRAVAILN4'+' '+OrderBy;
		TFQRS1(Ecran).WhereSQl := SQL;
	end
	else
	begin
		If GetControlText('CSAUT4') = 'X' Then SetControlText('CSAUT3','X');
		If GetControlText('CSAUT3') = 'X' Then SetControlText('CSAUT2','X');
		If GetControlText('CSAUT2') = 'X' Then SetControlText('CSAUT1','X');

		For i := 1 to 4 do
		begin
			SetControltext('XX_RUPTURE'+IntToStr(i),'');
		end;
		For i := 1 to 4 do
		begin
			Combo := THValComboBox(GetControl('THVRUPTURE'+IntToStr(i)));
			If Combo.Value <> '' Then
			begin
				If Combo.Value = 'F1' Then CreerRupture('PST_FORMATION1',i);
				If Combo.Value = 'F2' Then CreerRupture('PST_FORMATION2',i);
				If Combo.Value = 'F3' Then CreerRupture('PST_FORMATION3',i);
				If Combo.Value = 'F4' Then CreerRupture('PST_FORMATION4',i);
				If Combo.Value = 'F5' Then CreerRupture('PST_FORMATION5',i);
				If Combo.Value = 'F6' Then CreerRupture('PST_FORMATION6',i);
				If Combo.Value = 'F7' Then CreerRupture('PST_FORMATION7',i);
				If Combo.Value = 'F8' Then CreerRupture('PST_FORMATION8',i);
				If Combo.Value = 'ET' Then CreerRupture('PFI_ETABLISSEMENT',i);
				If Combo.Value = 'LE' Then CreerRupture('PFI_LIBEMPLOIFOR',i);
				If Combo.Value = 'RF' Then CreerRupture('PFI_RESPONSFOR',i);
				If Combo.Value = 'TY' Then CreerRupture('PST_NATUREFORM',i);
				If Combo.Value = 'ST' Then CreerRupture('PFI_CODESTAGE',i);
				If Combo.Value = 'T1' Then CreerRupture('PFI_TRAVAILN1',i);
				If Combo.Value = 'T2' Then CreerRupture('PFI_TRAVAILN2',i);
				If Combo.Value = 'T3' Then CreerRupture('PFI_TRAVAILN3',i);
				If Combo.Value = 'T4' Then CreerRupture('PFI_TRAVAILN4',i);
				If Combo.Value = 'LI' Then CreerRupture('PFI_LIEUFORM',i);
				If Combo.Value = 'TYP' Then CreerRupture('PFI_TYPEPLANPREV',i);
			end
			Else SetControlText('XX_RUPTURE'+IntToStr(i),'');
		end;
		
		//PT22
		// N'affiche pas les salariés confidentiels
		If Where <> '' Then Where := Where + ' AND ';
		If PGBundleHierarchie Then
			Where := Where + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))'  //PT24
		Else
			Where := Where + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))';  //PT24
        
		SQL := 'SELECT PFI_TYPOFORMATION,PFI_TYPEPLANPREV,PFI_DATEDIF,PFI_HTPSTRAV,PFI_HTPSNONTRAV,PFI_NATUREFORM,'+
		'PFI_CURSUS,PFI_CODESTAGE,PFI_MILLESIME,PFI_SALARIE,PFI_LIBELLE,PFI_NBINSC,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,'+
		'PFI_RESPONSFOR,PFI_COUTREELSAL,PFI_FRAISFORFAIT,PFI_DUREESTAGE,PFI_LIEUFORM'+
		',PST_LIBELLE,PST_NATUREFORM,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
		',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_DUREESTAGE,PFI_TRAVAILN1,PFI_TRAVAILN2,PFI_TRAVAILN3,PFI_TRAVAILN4'+
		' FROM INSCFORMATION LEFT JOIN STAGE '+
		'ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME ';
		
		If PGBundleHierarchie Then SQL := SQL + 'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFI_SALARIE '; //PT18
		
		SQL := SQL +Where+' '+OrderBy;
		
		TFQRS1(Ecran).WhereSQl := SQL;
	end;
end;

procedure TOF_PGEDITBUDGETFOR.CreerRupture(Champ:String;Num:Integer);
begin
        SetControlText('XX_RUPTURE'+IntToStr(Num),Champ);
        If OrderBy <> '' Then OrderBy := OrderBy+','+Champ
        Else OrderBy := 'ORDER BY '+Champ;
end;


procedure TOF_PGEDITBUDGETFOR.OnArgument(S : String);
var Num,i : Integer;
    Millesime,Combo : THValComboBox;
    Stage,Resp,edit : THEdit;
begin
  Inherited ;
        Arg := S;
        {$IFDEF EMANAGER}
//        If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
//        'WHERE PHO_NIVEAUH<=2 AND PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"') then MultiNiveau := True
        If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
        'AND PGS_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE)') then MultiNiveau := True
        else MultiNiveau := False;
        If MultiNiveau = False then
        begin
                SetControlText('PFI_RESPONSFOR',V_PGI.UserSalarie);
                SetControlVisible('PFI_RESPONSFOR',False);
                SetControlVisible('TRESPONSFOR',False);
        end;
        {$ENDIF}

		{$IFNDEF EMANAGER}
		If PGBundleHierarchie Then //PT20
		Begin
		{$ENDIF}
	        Edit := THEdit(getControl('PFI_SALARIE'));
	        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        {$IFNDEF EMANAGER}
    	End;
        {$ENDIF}

        If Arg = 'CWASBUDGET' then
        begin
                SetControlText('PFI_RESPONSFOR',V_PGI.UserSalarie);
                SetControlVisible('PFI_RESPONSFOR',False);
                SetControlVisible('TRESPONSFOR',False);
                SetControlCaption('LIBRESP',RechDom ('PGSALARIE',V_PGI.UserSalarie,FALSE));
                SetControlVisible('LIBRESP',False);
        end
        else SetcontrolCaption('LIBRESP','');
        SetControlCaption('LIBSTAGE','');
        Stage := THEdit(Getcontrol('PST_CODESTAGE'));
        If Stage <> Nil Then Stage.OnElipsisClick := StageElipsisCLick;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        Millesime := THValComboBox(GetControl('PST_MILLESIME'));
        Millesime.Value := RendMillesimePrevisionnel;
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('THVRUPTURE'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                        Combo.Items.Add('<<Aucun>>');
                        Combo.Values.Add('');
                        if VH_Paie.PGResponsables = true then
                        begin
                                Combo.Items.Add('Responsable formation');
                                Combo.Values.Add('RF');
                        end;
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('ST');
                        Combo.Items.Add('Type de plan');
                        Combo.Values.Add('TYP');
                        Combo.Items.Add('Etablissement');
                        Combo.Values.Add('ET');
                        Combo.Items.Add('Libellé emploi');
                        Combo.Values.Add('LE');
                        Combo.Items.Add('Lieu formation');
                        Combo.Values.Add('LI');
                        Combo.Items.Add('Type (Interne/Externe)');
                        Combo.Values.Add('TY');
                        For Num  :=  1 to VH_Paie.NBFormationLibre do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('F1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('F2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('F3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('F4');end;
                                If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('F5');end;
                                If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('F6');end;
                                If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('F7');end;
                                If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('F8');end;
                        end;
                        For Num := 1 to VH_Paie.PGNbreStatOrg do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('T1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('T2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('T3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('T4');end;
                        end;
                end;
        end;
        SetControlProperty('THVRUPTURE1','Value','');
        SetControlProperty('THVRUPTURE2','Value','');
        SetControlProperty('THVRUPTURE3','Value','');
        SetControlProperty('THVRUPTURE4','Value','');
        If PGBundleHierarchie Then //PT20
        Begin
	        Resp := THEdit(GetControl('PFI_RESPONSFOR'));
	        If Resp <> Nil then Resp.OnElipsisClick := RespElipsisClick;
	    End;
	    
        SetControlCaption('LIBELLEEMPLOI','');
        SetControlCaption('LIBETAB','');
        if VH_Paie.PGResponsables = False then
        begin
                SetControlVisible ('PFI_RESPONSFOR',False);
                SetControlVisible('LIBRESP',False);
                SetControlVisible('TPFI_RESPONSFOR',False);
        end;
        SetControlText('PFI_ETATINSCFOR','VAL');
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFI_TRAVAILN'+IntToStr(Num)),GetControl ('TPFI_TRAVAILN'+IntToStr(Num)));
        end;
        If Arg = 'INDIVIDUEL' then
        begin
                SetControlProperty('PRUPTURES','TabVisible',False);
                SetControlvisible('PRUPTURES',False);
                SetControlVisible('CSANSDETAIL',False);
                Ecran.Caption := 'Edition du prévisionnel individuel';
                SetControlVisible('CAFFICHECOUTP',True);
                SetControlVisible('CAFFICHECOUTS',True);
                SetControlVisible('SYNTHETIQUE',True);
                SetControlVisible('CSAUTSAL',True);
                SetControlChecked('CSAUTSAL',True);
                UpdateCaption(Ecran);
        end;
        
	//PT23 - Début
	If PGBundleCatalogue Then
	Begin
		If not PGDroitMultiForm then
			SetControlProperty ('PST_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
			SetControlProperty ('PST_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); 
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFI_PREDEFINI',False);
			SetControlText   ('PFI_PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); 
		End;
	end
	else
	begin
		SetControlVisible('PFI_PREDEFINI', False);
		SetControlVisible('TPFI_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT23 - Fin
end;

procedure TOF_PGEDITBUDGETFOR.StageElipsisClick(Sender : TObject);
var StWhere : String;
begin
        StWHere := 'PST_ACTIF="X" AND PST_MILLESIME="'+getControlText('PST_MILLESIME')+'"';
		//PT18 - Début
        If PGBundleInscFormation Then
        Begin
        	If Not PGDroitMultiForm Then 
        		StWhere := StWhere + DossiersAInterroger('',V_PGI.NoDossier, 'PST', True, True) //PT18
        	Else
        		StWhere := StWhere + DossiersAInterroger('','', 'PST', True, True); //PT18
        End;
		//PT18 - Fin
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,'', True,-1);
end;

procedure TOF_PGEDITBUDGETFOR.RespElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}var StWhere : String;{$ENDIF}
begin
        {$IFDEF EMANAGER}
        StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
        LookupList(THEdit(Sender),'Liste des responsables de formation','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWHere,'', True,-1);
        {$ELSE}
        ElipsisResponsableMultidos (Sender); //PT18
        {$ENDIF}
end;

{TOF_PGEDITCOMBUDGET}
procedure TOF_PGEDITCOMBUDGET.OnArgument (S : String ) ;
var Millesime : THValComboBox;
    DD,DF : TDateTime;
    Stage,DateDebut,DateFin : THEdit;
begin
        SetControlCaption('LIBSTAGE','');
        Stage := THEdit(GetControl('PFI_CODESTAGE'));
        DateDebut := THEdit(GetControl('DATEDEBUT'));
        DateFin := THEdit(GetControl('DATEFIN'));
        Millesime := THValComboBox(GetControl('PFI_MILLESIME'));
        Millesime.Value := RendMillesimeRealise(DD,DF);
        DateDebut.Text := DateToStr(DD);
        DateFin.Text := DateToStr(DF);
        If DateDebut <> Nil Then DateDebut.OnElipsisClick := DateElipsisclick;
        If DateFin <> Nil Then DateFin.OnElipsisClick := DateElipsisClick;
        If Stage <> Nil Then Stage.OnElipsisClick := StageElipsisClick;
        Millesime.OnChange := ChangeMillesime;
end;

procedure TOF_PGEDITCOMBUDGET.StageElipsisClick(Sender : TObject);
var StWhere : String;
    Millesime : THValComboBox;
begin
        Millesime := THValComboBox(GetControl('PFI_MILLESIME'));
        StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="'+Millesime.Value+'" ';
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,'', True,-1);
end;

procedure TOF_PGEDITCOMBUDGET.DateElipsisClick(Sender:TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITCOMBUDGET.ChangeMillesime(Sender : Tobject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+THValComboBox(Sender).Value+'"',True);
        If Not Q.Eof Then
        begin
                SetControlText('DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

{TOF_PGEDITCOMPSESSION}
procedure TOF_PGEDITCOMPSESSION.OnArgument (S : String ) ;
var Millesime : THValComboBox;
    Stage,DateDebut,DateFin : THEdit;
    DD,DF : TDateTime;
begin
        SetControlCaption('LIBSTAGE','');
        Stage := THEdit(GetControl('PFO_CODESTAGE'));
        DateDebut := THEdit(GetControl('PFO_DATEDEBUT'));
        DateFin := THEdit(GetControl('PFO_DATEFIN'));
        Millesime := THValComboBox(GetControl('MILLESIME'));
        Millesime.Value := RendMillesimeRealise(DD,DF);
        DateDebut.Text := DateToStr(DD);
        DateFin.Text := DateToStr(DF);
        If DateDebut <> Nil Then DateDebut.OnElipsisClick := DateElipsisclick;
        If DateFin <> Nil Then DateFin.OnElipsisClick := DateElipsisClick;
        If Stage <> Nil Then
        begin
                Stage.OnElipsisClick := StageElipsisClick;
                Stage.OnChange := StageChange;
        end;
        If Millesime <> Nil Then Millesime.OnChange := ChangeMillesime;
end;

procedure TOF_PGEDITCOMPSESSION.DateElipsisClick(Sender : TObject);
var key  :  char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITCOMPSESSION.ChangeMillesime(Sender : Tobject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+THValComboBox(Sender).Value+'"',True);
        If Not Q.Eof Then
        begin
                SetControlText('PFO_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PFO_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

procedure TOF_PGEDITCOMPSESSION.StageElipsisClick(Sender : TObject);
var StWhere : String;
    Millesime : THValComboBox;
begin
        Millesime := THValComboBox(GetControl('MILLESIME'));
        StWhere := 'PST_ACTIF="X" AND PST_MILLESIME="'+Millesime.Value+'" ';
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,'', True,-1);
end;

procedure TOF_PGEDITCOMPSESSION.StageChange(Sender : TObject);
var Session : THValComboBox;
begin
        Session := THValComboBox(GetControl('PFO_ORDRE'));
        If GetControlText('PFO_CODESTAGE') <> '' Then
        begin
                Session.Enabled := True;
                Session.Plus := 'PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('PFO_DATEDEBUT')))+'"'+
                ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('PFO_DATEFIN')))+'"'+
                ' AND PSS_CODESTAGE="'+GetControlText('PFO_CODESTAGE')+'"';
        end
        Else Session.enabled := False;
end;

{TOF_PGEDITFORREABUD}
Procedure TOF_PGEDITFORREABUD.OnUpdate ;
var Pages : TPageControl;
    Where,OrderBy : String;
    i : Integer;
begin
  Inherited ;
  OrderBy := '';
  Pages := TPageControl(GetControl('Pages'));
  
  Where := RecupWhereCritere(Pages);

  //PT5 - Début
  If LeMenu = '47572' then EditionComparCollectif;
  If LeMenu = '47571' then EditionComparIndividuel;
  If LeMenu = '47575' then EditionComparPluriAnnuel; //PT6
  //PT5 - Fin

  For i := 1 to 4 do
  begin
        If GetControlText('VRUPT'+IntToStr(i)) <> '' then
        begin
                If OrderBy = '' then OrderBy := ' ORDER BY '+GetControlText('VRUPT'+IntToStr(i))
                else OrderBy := OrderBy+','+GetControltext('VRUPT'+IntToStr(i));
                SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
        end
        else SetControlText('XX_RUPTURE'+IntToStr(i),'');
  end;
end ;

procedure TOF_PGEDITFORREABUD.OnArgument (S : String ) ;
var Num,i : Integer;
    Combo : THValComboBox;
    CkBox : TCheckBox;
    Edit : THEdit;
    Millesime : String;
    DD,DF : TDateTime;
begin
    Inherited ;
    TobEtat := nil; //PT5
    
    Edit := THEdit(GetControl('PFO_SALARIE'));
    If Edit <> Nil then 
    Begin
    	Edit.OnExit := ExitEdit;
		If PGBundleHierarchie Then Edit.OnElipsisClick := SalarieElipsisClick; //PT18 /PT20
	End;
    Edit := THEdit(GetControl('PFO_SALARIE_'));
    If Edit <> Nil then 
    Begin
    	Edit.OnExit := ExitEdit;
    	If PGBundleHierarchie Then Edit.OnElipsisClick := SalarieElipsisClick; //PT18 //PT20
    End;
    
    LeMenu := ReadTokenPipe(S,';');
    SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
    Millesime := RendMillesimeRealise(DD,DF);
    SetControlText('DATEDEBUT',DateToStr(DD));
    SetControlText('DATEFIN',DateToStr(DF));
    
    For Num  :=  1 to VH_Paie.NBFormationLibre do
    begin
            if Num >8 then Break;
            VisibiliteChampFormation (IntToStr(Num),GetControl ('PFO_FORMATION'+IntToStr(Num)),GetControl ('TPFO_FORMATION'+IntToStr(Num)));
    end;
    
    For Num  :=  1 to VH_Paie.PGNbreStatOrg do
    begin
            if Num >4 then Break;
            VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
            SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
    end;
    VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
    
    If LeMenu = '47572' then     //Suivi comparatif collectif
    begin
            SetControlVisible('SANSDETAIL',False);
//                SetControlProperty('PRUPT','TabVisible',False);
            TFQRS1(Ecran).CodeEtat := 'PCC';  //PT5
            TFQRS1(Ecran).FCodeEtat := 'PCC';
            Ecran.Caption := 'Suivi comparatif collectif';
    end
    else If LeMenu = '47571' then     //Suivi comparatif individuel
    begin
            SetControlVisible('SANSDETAIL',False);
//                SetControlProperty('PRUPT','TabVisible',False);
            TFQRS1(Ecran).CodeEtat := 'PL1';  //PT5
            TFQRS1(Ecran).FCodeEtat := 'PL1';
            Ecran.Caption := 'Suivi comparatif individuel';
            TTabSheet(GetControl('PRUPT')).TabVisible := False;
    end
    //PT6 - Début
    else If LeMenu = '47575' Then     // Comparatif pluri-annuel
    Begin
            TFQRS1(Ecran).CodeEtat := 'PCE';
            TFQRS1(Ecran).FCodeEtat := 'PCE';
            Ecran.Caption := 'Suivi comparatif pluri-annuel';
            SetControlVisible('SANSDETAIL',False);
            SetControlVisible('NBEXERCICES',True);
            SetControlVisible('TNBEXERCICES',True);
            SetControlVisible('MILLESIMEUNIQUE',True);
            SetControlVisible('MILLESIME',False);
            SetControlVisible('DATEDEBUT',False);SetControlVisible('TPFO_DATEDEBUT',False);
            SetControlVisible('DATEFIN',False);SetControlVisible('TPFO_DATEFIN',False);
            // Présélections
            SetControlText('MILLESIMEUNIQUE',Millesime);
            SetControlText('NBEXERCICES','2');
    End;
    //PT6 - Fin

    UpdateCaption(Ecran);
    For i := 1 To 4 do
    begin
            Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
            If Combo <> Nil Then
            begin
                   //PT5 - Début
                   // Rendre invisible toutes les check-boxes de saut de page par défaut
                   CkBox := TCheckBox(GetControl('CSAUTRUPT'+IntToStr(i)));
                   CkBox.Visible := False;
                   // Sur le changement de valeur d'une combo, il faut afficher le saut de page précédent
                   Combo.OnChange := OnChangeRuptures;
                   //PT5 - Fin

                    If i <> 1 Then
                    begin
                            Combo.Items.Add('<<Aucune>>');
                            Combo.Values.Add('');
                    end;
                    Combo.Items.Add('Salarié');
                    Combo.Values.Add('PFO_SALARIE');
                    Combo.Items.Add('Responsable formation');
                    Combo.Values.Add('PFO_RESPONSFOR');
                    Combo.Items.Add('Formation');
                    Combo.Values.Add('PFO_CODESTAGE');
                    Combo.Items.Add('Etablissement');
                    Combo.Values.Add('PFO_ETABLISSEMENT');
                    Combo.Items.Add('Libellé emploi');
                    Combo.Values.Add('PFO_LIBEMPLOIFOR');
                    Combo.Items.Add('Nature (Interne/Externe)');
                    Combo.Values.Add('PFO_NATUREFORM');
                    Combo.Items.Add('Lieu');
                    Combo.Values.Add('PFO_LIEUFORM');
                    Combo.Items.Add('Centre de formation');
                    Combo.Values.Add('PFO_CENTREFORMGU');
                    For Num  :=  1 to VH_Paie.NBFormationLibre do
                    begin
                            If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PFO_FORMATION1');end;
                            If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PFO_FORMATION2');end;
                            If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PFO_FORMATION3');end;
                            If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PFO_FORMATION4');end;
                            If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PFO_FORMATION5');end;
                            If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PFO_FORMATION6');end;
                            If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PFO_FORMATION7');end;
                            If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PFO_FORMATION8');end;
                    end;
                    For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                    begin
                            If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('PFO_TRAVAILN1');end;
                            If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('PFO_TRAVAILN2');end;
                            If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('PFO_TRAVAILN3');end;
                            If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('PFO_TRAVAILN4');end;
                    end;
                    If VH_Paie.PGLibCodeStat <> '' Then begin Combo.Items.Add(VH_Paie.PGLibCodeStat); Combo.Values.Add('PFO_CODESTAT');end;
            end;
    end;
    SetControlText('VRUPT1','PFO_CODESTAGE');
    
    {$IFDEF EMANAGER}
    SetControlText('PFO_RESPONSFOR',V_PGI.UserSalarie);
    {$ENDIF}
        
	//PT18 - Début
	If PGBundleCatalogue Then //PT20
	Begin
		If not PGDroitMultiForm then
			SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
			SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT18
	End;
	
    //PT22
    // N'affiche pas les salariés confidentiels
    If PGBundleHierarchie Then
    	SetControlText('XX_WHERE', 'PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")')
    Else
    	SetControlText('XX_WHERE', 'PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")');

	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFO_PREDEFINI',False);
			SetControlText   ('PFO_PREDEFINI','');
			//SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT20
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT19
			//SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT18 //PT20
		End;
	end
	else
	begin
		SetControlVisible('PFO_PREDEFINI', False);
		SetControlVisible('TPFO_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT18 - Fin        
end ;

procedure TOF_PGEDITFORREABUD.EditionComparCollectif;
var TobNbInscBud,TobNbInscRea,TobBudget,TobRealise : Tob;
     TNbIB,TNbIR,T : Tob;
     Q : TQuery;
     DateDebut,DateFin : TDateTime;
     i,NbRupt : Integer;
     Salarie,Where,WherePFI,WherePFO,Millesime,CodeStage,StPages,WhereMillesime,StrTemp : String; //PT6
     Rupt1,Rupt2,Rupt3,Rupt4 : String;
     ChampRuptRea1,ChampRuptRea2,ChampRuptRea3,ChampRuptRea4 : String;
     ChampRuptBud1,ChampRuptBud2,ChampRuptBud3,ChampRuptBud4 : String;
     ValeurRupt1,ValeurRupt2,ValeurRupt3,ValeurRupt4 : String;
     Tablette1,Tablette2,Tablette3,Tablette4,Libelle : String;
     NbStaMax,NbInscSal,NbInsc : Integer;
     NbHeuresR,NbHeuresB,CoutPedag,CoutAnim,CoutUnitaire,CoutSal,Frais,NbHSession : Double;
     TotalRea,TotalBud,NbSession : Double;
     Pages : TPageControl;
     ListeChampRea,ListeChampBud : String;
begin
        NbInsc := 0;
        NbHSession := 0;
        Pages := TPageControl(GetControl('Pages'));
		SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFO')); //PT18
		
        Rupt1 := '';
        Rupt2 := '';
        Rupt3 := '';
        Rupt4 := '';
        If GetControltext('VRUPT4') <> '' then
        begin
                Rupt4 := GetControltext('VRUPT4');
                Rupt3 := GetControltext('VRUPT3');
                Rupt2 := GetControltext('VRUPT2');
                Rupt1 := GetControltext('VRUPT1');
                SetControlText('XX_RUPTURE1','RUPTURE1');
                SetControlText('XX_RUPTURE2','RUPTURE2');
                SetControlText('XX_RUPTURE3','RUPTURE3');
                SetControlText('XX_RUPTURE4','RUPTURE4');
                NbRupt := 4;
//                If GetControlText('VRUPT3') = '' then begin PGIBox('Le niveau de rupture 3 doit être renseigné',Ecran.Caption);exit;end;//PT5
//                If GetControlText('VRUPT2') = '' then begin PGIBox('Le niveau de rupture 2 doit être renseigné',Ecran.Caption);exit;end;//PT5
        end
        Else
        begin
                If GetControlText('VRUPT3') <> '' then
                begin
                        Rupt3 := GetControltext('VRUPT3');
                        Rupt2 := GetControltext('VRUPT2');
                        Rupt1 := GetControltext('VRUPT1');
                        SetControlText('XX_RUPTURE1','RUPTURE1');
                        SetControlText('XX_RUPTURE2','RUPTURE2');
                        SetControlText('XX_RUPTURE3','RUPTURE3');
                        SetControlText('XX_RUPTURE4','');
                        NbRupt := 3;
//                        If GetControlText('VRUPT2') = '' then begin PGIBox('Le niveau de rupture 2 doit être renseigné',Ecran.Caption);exit;end;//PT5
                end
                Else
                begin
                        If GetControlText('VRUPT2') <> '' then
                        begin
                                Rupt2 := GetControltext('VRUPT2');
                                Rupt1 := GetControltext('VRUPT1');
                                SetControlText('XX_RUPTURE1','RUPTURE1');
                                SetControlText('XX_RUPTURE2','RUPTURE2');
                                SetControlText('XX_RUPTURE3','');
                                SetControlText('XX_RUPTURE4','');
                                NbRupt := 2;
                        end
                        Else
                        begin
                                Rupt1 := GetControltext('VRUPT1');
                                SetControlText('XX_RUPTURE1','RUPTURE1');
                                SetControlText('XX_RUPTURE2','');
                                SetControlText('XX_RUPTURE3','');
                                SetControlText('XX_RUPTURE4','');
                                NbRupt := 1;
                        end;
                end;
        end;
//        SetControlText('NBRUPT',IntToStr(NbRupt));
        ChampRuptRea1 := Rupt1;
        ChampRuptRea2 := Rupt2;
        ChampRuptRea3 := Rupt3;
        ChampRuptRea4 := Rupt4;
        ChampRuptBud1 := ConvertPrefixe(Rupt1,'PSS','PST');
        ChampRuptBud2 := ConvertPrefixe(Rupt2,'PSS','PST');
        ChampRuptBud3 := ConvertPrefixe(Rupt3,'PSS','PST');
        ChampRuptBud4 := ConvertPrefixe(Rupt4,'PSS','PST');
        ChampRuptBud1 := ConvertPrefixe(ChampRuptBud1,'PFO','PFI');
        ChampRuptBud2 := ConvertPrefixe(ChampRuptBud2,'PFO','PFI');
        ChampRuptBud3 := ConvertPrefixe(ChampRuptBud3,'PFO','PFI');
        ChampRuptBud4 := ConvertPrefixe(ChampRuptBud4,'PFO','PFI');
        If Rupt1 <> '' then ListeChampRea := ListeChampRea+','+Rupt1;
        If Rupt2 <> '' then ListeChampRea := ListeChampRea+','+Rupt2;
        If Rupt3 <> '' then ListeChampRea := ListeChampRea+','+Rupt3;
        If Rupt4 <> '' then ListeChampRea := ListeChampRea+','+Rupt4;
        If ChampRuptBud1 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud1;
        If ChampRuptBud2 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud2;
        If ChampRuptBud3 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud3;
        If ChampRuptBud4 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud4;
        Tablette1 := ChampRuptureForm('TAB',Rupt1);
        Tablette2 := ChampRuptureForm('TAB',Rupt2);
        Tablette3 := ChampRuptureForm('TAB',Rupt3);
        Tablette4 := ChampRuptureForm('TAB',Rupt4);

        DateDebut := StrToDate(GetControlText('DATEDEBUT'));
        dateFin := StrToDate(GetControlText('DATEFIN'));

        Where := '';
        Where := RecupWhereCritere(Pages);
        WherePFO := Where + ' AND PSS_PGTYPESESSION<>"SOS"'; //PT8
        WherePFI := ConvertPrefixe(Where,'PFO','PFI');
        WherePFI := ConvertPrefixe(WherePFI,'PSS','PST');

        // Cas particulier du millesime en multisélection
        WhereMillesime := GetMillesimes; //PT6

        If WherePFO <> '' then WherePFO := WherePFO+' AND PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PFO_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'" AND PFO_EFFECTUE="X"'
        else WherePFO := 'WHERE PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PFO_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'" AND PFO_EFFECTUE="X"';
        If WherePFI <> '' then WherePFI := WherePFI + 'AND PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--"'+WhereMillesime
        Else  WHerePFI := 'WHERE PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--"'+WhereMillesime;
        If (ChampRuptBud1 = 'PFI_SALARIE') or (ChampRuptBud2 = 'PFI_SALARIE') or (ChampRuptBud3 = 'PFI_SALARIE') or (ChampRuptBud4 = 'PFI_SALARIE') then
        Where := Where + ' AND PFI_SALARIE <> ""';
        
        Q := OpenSQL('SELECT PFI_MILLESIME,PFI_CODESTAGE,PFI_NBINSC,PFI_COUTREELSAL,PFI_FRAISFORFAIT,PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,PFI_RESPONSFOR'+
        ',PST_DUREESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX'+ListeChampBud+
        ' FROM INSCFORMATION LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+WherePFI,True);
        TobBudget := Tob.Create('Les inscriptions au budget',Nil,-1);
        TobBudget.LoadDetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);
        
        If PGBundleInscFormation Then //PT18
        	Where := DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFI',False,True)
        Else
        	Where := '';
        	
        Q := OpenSQL('SELECT SUM(PFI_NBINSC) AS NBINSC,PFI_MILLESIME,PFI_CODESTAGE FROM INSCFORMATION'+
        ' WHERE PFI_CODESTAGE<>"--CURSUS--" ' + WhereMillesime + Where +
        ' GROUP BY PFI_CODESTAGE,PFI_MILLESIME',True);
        TobNbInscBud := Tob.Create('Nombre insc',Nil,-1);
        TobNbInscBud.LoadDetailDB('Nombre insc','','',Q,False);
        Ferme(Q);
        
        TobEtat := Tob.Create('Les salariés pour édition',Nil,-1);
        For i := 0 to TobBudget.Detail.Count-1 do
        begin
                Salarie := TobBudget.Detail[i].GetValue('PFI_SALARIE');
                CodeStage := TobBudget.Detail[i].GetValue('PFI_CODESTAGE');
                Millesime := TobBudget.Detail[i].GetValue('PFI_MILLESIME');
                TNbIB := TobNbInscBud.FindFirst(['PFI_CODESTAGE','PFI_MILLESIME'],[CodeStage,Millesime],False);
                If TNbIB <> Nil then NbInsc := TNbIB.GetValue('NBINSC')
                Else Exit;
                NbStaMax := TobBudget.Detail[i].GetValue('PST_NBSTAMAX');
                NbSession := NombreDeSessionFormationAPrevoir(NbStaMax,NbInsc);
                NbInscSal := TobBudget.Detail[i].GetValue('PFI_NBINSC');
                NbHeuresB := TobBudget.Detail[i].GetValue('PST_DUREESTAGE');
                NbHeuresB := NbHeuresB * NbInscSal;
                CoutPedag := TobBudget.Detail[i].GetValue('PST_COUTBUDGETE');
                CoutPedag := CoutPedag*NbSession*NbInscSal;
                If NbInsc > 0 then CoutPedag := arrondi(CoutPedag/NbInsc,2)
                else CoutPedag :=0;
                CoutAnim := TobBudget.Detail[i].GetValue('PST_COUTSALAIR');
                CoutAnim := CoutAnim*NbSession*NbInscSal;
                CoutAnim := arrondi(CoutAnim/NbInsc,2);
                CoutUnitaire := TobBudget.Detail[i].GetValue('PST_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire*NbInscSal;
                CoutSal := TobBudget.Detail[i].GetValue('PFI_COUTREELSAL');
                Frais := TobBudget.Detail[i].GetValue('PFI_FRAISFORFAIT');
                TotalBud := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                ValeurRupt1 := '';
                ValeurRupt2 := '';
                ValeurRupt3 := '';
                ValeurRupt4 := '';
                If NbRupt > 3 then ValeurRupt4 := TobBudget.Detail[i].GetValue(ChampRuptBud4);
                If NbRupt > 2 then ValeurRupt3 := TobBudget.Detail[i].GetValue(ChampRuptBud3);
                If NbRupt > 1 then ValeurRupt2 := TobBudget.Detail[i].GetValue(ChampRuptBud2);
                ValeurRupt1 := TobBudget.Detail[i].GetValue(ChampRuptBud1);
                {If NbRupt=4 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt4,ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=3 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE2','RUPTURE1'],[ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=2 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE1'],[ValeurRupt2,ValeurRupt1],False)
                Else T := TobEtat.FindFirst(['RUPTURE4'],[ValeurRupt1],False);       }
                If NbRupt=4 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt4,ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=3 then T := TobEtat.FindFirst(['RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=2 then T := TobEtat.FindFirst(['RUPTURE2','RUPTURE1'],[ValeurRupt2,ValeurRupt1],False)
                Else T := TobEtat.FindFirst(['RUPTURE1'],[ValeurRupt1],False);
                If T <> Nil then
                begin
                        T.PutValue('HEURESNR',T.GetValue('HEURESNR')+NbHeuresB);
                        T.PutValue('COUTSNR',T.GetValue('COUTSNR')+TotalBud);
                        T.PutValue('HEURESPREV',T.GetValue('HEURESPREV')+NbHeuresB);
                        T.PutValue('COUTSPREV',T.GetValue('COUTSPREV')+TotalBud);
                end
                Else
                begin
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSupValeur('RUPTURE1','');T.AddChampSupValeur('RUPTURE2','');
                        T.AddChampSupValeur('RUPTURE3','');T.AddChampSupValeur('RUPTURE4','');
                        T.AddChampSupValeur('LIBRUPTURE1','');T.AddChampSupValeur('LIBRUPTURE2','');
                        T.AddChampSupValeur('LIBRUPTURE3','');T.AddChampSupValeur('LIBRUPTURE4','');
                        T.AddChampSupValeur('HEURESR',0);T.AddChampSupValeur('COUTSR',0);
                        T.AddChampSupValeur('HEURESNR',0);T.AddChampSupValeur('COUTSNR',0);
                        T.AddChampSupValeur('HEURESNP',0);T.AddChampSupValeur('COUTSNP',0);
                        T.AddChampSupValeur('HEURESPREV',0);T.AddChampSupValeur('COUTSPREV',0);
                        T.AddChampSupValeur('HEURESREAL',0);T.AddChampSupValeur('COUTSREAL',0);
                        T.PutValue('HEURESNR',NbHeuresB);
                        T.PutValue('COUTSNR',TotalBud);
                        T.PutValue('HEURESPREV',NbHeuresB);
                        T.PutValue('COUTSPREV',TotalBud);
                        If NbRupt > 3 then T.PutValue('RUPTURE4',TobBudget.Detail[i].GetValue(ChampRuptBud4));
                        If NbRupt > 2 then T.PutValue('RUPTURE3',TobBudget.Detail[i].GetValue(ChampRuptBud3));
                        If NbRupt > 1 then T.PutValue('RUPTURE2',TobBudget.Detail[i].GetValue(ChampRuptBud2));
                        T.PutValue('RUPTURE1',TobBudget.Detail[i].GetValue(ChampRuptBud1));
                end;
        end;
        TobNbInscBud.Free;
        TobBudget.Free;
        
        Q := OpenSQL('SELECT PSS_COUTPEDAG,PSS_COUTSALAIR,PSS_COUTUNITAIRE,PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_RESPONSFOR'+
			        ',PFO_NBREHEURE,PFO_FRAISREEL,PFO_COUTREELSAL,PFO_NBREHEURE'+ListeChampRea+
			        ' FROM FORMATIONS LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME '+WherePFO,True);
        TobRealise := Tob.Create('Les inscriptions au formations',Nil,-1);
        TobRealise.LoadDetailDB('FORMATIONS','','',Q,False);
        Ferme(Q);
        
        If PGBundleInscFormation Then //PT18
        	Where := DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFO',False,True)
        Else
        	Where := '';
        	
        Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) AS NBSAL,PFO_ORDRE,PFO_MILLESIME,PFO_CODESTAGE FROM FORMATIONS WHERE '+
			        'PFO_DATEDEBUT>="'+UsDatEtime(DateDebut)+'" AND PFO_PGTYPESESSION<>"SOS" AND PFO_DATEDEBUT<="'+UsDateTime(Datefin)+'" AND PFO_EFFECTUE="X"'+ //PT8
			        Where+
			        ' GROUP BY PFO_ORDRE,PFO_CODESTAGE,PFO_MILLESIME',True);
        TobNbInscRea := Tob.Create('Les Sessions',Nil,-1);
        TobNbInscRea.LoadDetailDB('SESSIONSTAGE','','',Q,False);
        Ferme(Q);
        
        For i := 0 to TobRealise.detail.Count-1 do
        begin
                Salarie := TobRealise.Detail[i].GetValue('PFO_SALARIE');
                CodeStage := TobRealise.Detail[i].GetValue('PFO_CODESTAGE');
                Millesime := TobRealise.Detail[i].GetValue('PFO_MILLESIME');
                NbHeuresR := TobRealise.Detail[i].GetValue('PFO_NBREHEURE');
                TNbIR := TobNbInscRea.FindFirst(['PFO_CODESTAGE','PFO_MILLESIME','PFO_ORDRE'],[CodeStage,Millesime,TobRealise.Detail[i].GetValue('PFO_ORDRE')],False);
                If TNbIR <> Nil then NbHSession := TNbIR.GetValue('NBSAL')
                Else Continue;
                CoutPedag := TobRealise.Detail[i].GetValue('PSS_COUTPEDAG');
                If NbHSession> 0 then CoutPedag := arrondi(NbHeuresR*(CoutPedag/NbHSession),2)
                else CoutPedag := 0;
                CoutAnim := TobRealise.Detail[i].GetValue('PSS_COUTSALAIR');
                If NbHSession> 0 then CoutAnim := arrondi(NbHeuresR*(CoutAnim/NbHSession),2)
                else CoutAnim := 0;
                CoutUnitaire := TobRealise.Detail[i].GetValue('PSS_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire;
                CoutSal := TobRealise.Detail[i].GetValue('PFO_COUTREELSAL');
                Frais := TobRealise.Detail[i].GetValue('PFO_FRAISREEL');
                TotalRea := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                ValeurRupt1 := '';
                ValeurRupt2 := '';
                ValeurRupt3 := '';
                ValeurRupt4 := '';
                If NbRupt > 3 then ValeurRupt4 := TobRealise.Detail[i].GetValue(ChampRuptRea4);
                If NbRupt > 2 then ValeurRupt3 := TobRealise.Detail[i].GetValue(ChampRuptRea3);
                If NbRupt > 1 then ValeurRupt2 := TobRealise.Detail[i].GetValue(ChampRuptRea2);
                ValeurRupt1 := TobRealise.Detail[i].GetValue(ChampRuptRea1);
                If NbRupt=4 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt4,ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=3 then T := TobEtat.FindFirst(['RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=2 then T := TobEtat.FindFirst(['RUPTURE2','RUPTURE1'],[ValeurRupt2,ValeurRupt1],False)
                Else T := TobEtat.FindFirst(['RUPTURE1'],[ValeurRupt1],False);
                If T <> Nil then
                begin
                        NbHeuresB := T.GetValue('HEURESNR');
                        TotalBud := T.GetValue('COUTSNR');
                        T.PutValue('HEURESREAL',T.GetValue('HEURESREAL')+NbHeuresR);
                        T.PutValue('COUTSREAL',T.GetValue('COUTSREAL')+TotalRea);
                        If NbHeuresR>0 then
                        begin
                                if NbHeuresB>=NbHeuresR then
                                begin
                                        T.PutValue('HEURESNR',NbHeuresB-NbHeuresR);
                                        T.PutValue('HEURESR',T.GetValue('HEURESR')+NbHeuresR);
                                end
                                Else
                                begin
                                        T.PutValue('HEURESNP',T.GetValue('HEURESNP') + NbHeuresR - T.GetValue('HEURESNR'));
                                        T.PutValue('HEURESNR',0);
                                        T.PutValue('HEURESR',T.GetValue('HEURESR')+NbHeuresB);
                                end;
                        end;
                        If TotalRea>0 then
                        begin
                                if TotalBud>=TotalRea then
                                begin
                                        T.PutValue('COUTSNR',TotalBud-TotalRea);
                                        T.PutValue('COUTSR',T.GetValue('COUTSR')+TotalRea);
                                end
                                Else
                                begin
                                        T.PutValue('COUTSNP',T.GetValue('COUTSNP') + TotalRea - T.GetValue('COUTSNR'));
                                        T.PutValue('COUTSNR',0);
                                        T.PutValue('COUTSR',T.GetValue('COUTSR')+TotalBud);
                                end;
                        end;
                end
                Else
                begin
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSupValeur('RUPTURE1','');T.AddChampSupValeur('RUPTURE2','');
                        T.AddChampSupValeur('RUPTURE3','');T.AddChampSupValeur('RUPTURE4','');
                        T.AddChampSupValeur('LIBRUPTURE1','');T.AddChampSupValeur('LIBRUPTURE2','');
                        T.AddChampSupValeur('LIBRUPTURE3','');T.AddChampSupValeur('LIBRUPTURE4','');
                        T.AddChampSupValeur('HEURESR',0);T.AddChampSupValeur('COUTSR',0);
                        T.AddChampSupValeur('HEURESNR',0);T.AddChampSupValeur('COUTSNR',0);
                        T.AddChampSupValeur('HEURESNP',0);T.AddChampSupValeur('COUTSNP',0);
                        T.AddChampSupValeur('HEURESPREV',0);T.AddChampSupValeur('COUTSPREV',0);
                        T.AddChampSupValeur('HEURESREAL',0);T.AddChampSupValeur('COUTSREAL',0);
                        If NbRupt > 3 then T.PutValue('RUPTURE4',TobRealise.Detail[i].GetValue(ChampRuptRea4));
                        If NbRupt > 2 then T.PutValue('RUPTURE3',TobRealise.Detail[i].GetValue(ChampRuptRea3));
                        If NbRupt > 1 then T.PutValue('RUPTURE2',TobRealise.Detail[i].GetValue(ChampRuptRea2));
                        T.PutValue('RUPTURE1',TobRealise.Detail[i].GetValue(ChampRuptRea1));
                        T.PutValue('HEURESNP',NbHeuresR);T.PutValue('COUTSNP',TotalRea);
                        T.PutValue('HEURESREAL',NbHeuresR);
                        T.PutValue('COUTSREAL',TotalRea);
                end;
        end;
        TobRealise.Free;
        TobNbInscRea.Free;
        For i := 0 to TobEtat.detail.Count-1 do
        begin
                If NbRupt > 3 then
                begin
                        Libelle := RechDom(Tablette4,TobEtat.Detail[i].GetValue('RUPTURE4'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE4',Libelle);
                end;
                If NbRupt > 2 then
                begin
                        Libelle := RechDom(Tablette3,TobEtat.Detail[i].GetValue('RUPTURE3'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE3',Libelle);
                end;
                If NbRupt > 1 then
                begin
                        Libelle := RechDom(Tablette2,TobEtat.Detail[i].GetValue('RUPTURE2'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE2',Libelle);
                end;
                Libelle := RechDom(Tablette1,TobEtat.Detail[i].GetValue('RUPTURE1'),False);
                If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE1',Libelle);
        end;
        TobEtat.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;RUPTURE4');
        StPages := '';
       {$IFDEF EAGLCLIENT}
        StPages := AglGetCriteres (Pages, FALSE);
        {$ENDIF}
        //PT5 - Début
//        If GetCheckBoxState('FLISTE') = CbChecked then LanceEtatTOB('E','PFO','PCC',TobEtat,True,True,False,Pages,' ','',False,0,StPages)
//        else LanceEtatTOB('E','PFO','PCC',TobEtat,True,False,False,Pages,' ','',False,0,StPages);
        TFQRS1(Ecran).LaTob:= TobEtat;
//        TobEtat.free;
        //PT5 - Fin
end;

procedure TOF_PGEDITFORREABUD.EditionComparIndividuel;
var TobNbInscBud,TobNbInscRea,TobBudget,TobRealise : Tob;
     TNbIB,TNbIR,T : Tob;
     Q : TQuery;
     DateDebut,DateFin : TDateTime;
     i : Integer;
     Salarie,Where,WherePFI,WherePFO,Millesime,CodeStage,StPages : String;
     NbStaMax,NbInscSal,NbInsc : Integer;
     NbHeuresR,NbHeuresB,CoutPedag,CoutAnim,CoutUnitaire,CoutSal,Frais,NbHSession : Double;
     TotalRea,TotalBud,NbSession : Double;
     Pages : TPageControl;
     ListeChampRea,ListeChampBud : String;
begin
        NbInsc := 0;
        NbHSession := 0;
        Pages := TPageControl(GetControl('Pages'));
        DateDebut := StrToDate(GetControlText('DATEDEBUT'));
        dateFin := StrToDate(GetControlText('DATEFIN'));
        Where := '';
        Where := RecupWhereCritere(Pages);
        WherePFO := Where;
        WherePFI := ConvertPrefixe(Where,'PFO','PFI');
        WherePFI := ConvertPrefixe(WherePFI,'PSS','PST');
        If WherePFO <> '' then WherePFO := WherePFO+' AND PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PFO_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'" AND PFO_EFFECTUE="X"'
        else WherePFO := 'WHERE PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PFO_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'" AND PFO_EFFECTUE="X"';
        If WherePFI <> '' then WHerePFI := WherePFI + 'AND PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--" '+GetMillesimes //PT18
        Else WHerePFI := 'WHERE PFI_ETATINSCFOR="VAL" AND PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE<>"--CURSUS--" AND PFI_SALARIE<>""';
        WherePFO := WherePFO + ' AND PSS_PGTYPESESSION<>"SOS"'; //PT8
        
        If PGBundleHierarchie Then //PT13
	        Q := OpenSQL('SELECT PFI_MILLESIME,PFI_CODESTAGE,PFI_NBINSC,PFI_COUTREELSAL,PFI_FRAISFORFAIT,PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,PFI_RESPONSFOR'+
	        ',PST_DUREESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX'+
	        ',PSI_INTERIMAIRE,PSI_ETABLISSEMENT,PSI_LIBELLEEMPLOI,PSI_DATEENTREE,PSI_TRAVAILN1,PSI_TRAVAILN2,PSI_TRAVAILN3,PSI_TRAVAILN4'+
	        ',PST_CODESTAGE,PST_NATUREFORM,PST_CENTREFORMGU,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
	        ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8'+
	        ' FROM INSCFORMATION '+
	        'LEFT JOIN INTERIMAIRES ON PFI_SALARIE=PSI_INTERIMAIRE '+
	        'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+WherePFI,True)
        Else
	        Q := OpenSQL('SELECT PFI_MILLESIME,PFI_CODESTAGE,PFI_NBINSC,PFI_COUTREELSAL,PFI_FRAISFORFAIT,PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,PFI_RESPONSFOR'+
	        ',PST_DUREESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX'+
	        ',PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_DATEENTREE,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4'+
	        ',PST_CODESTAGE,PST_NATUREFORM,PST_CENTREFORMGU,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
	        ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8'+
	        ' FROM INSCFORMATION '+
	        'LEFT JOIN SALARIES ON PFI_SALARIE=PSA_SALARIE '+
	        'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+WherePFI,True);

        TobBudget := Tob.Create('Les inscriptions au budget',Nil,-1);
        TobBudget.LoadDetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);
        
        Q := OpenSQL('SELECT SUM(PFI_NBINSC) AS NBINSC,PFI_MILLESIME,PFI_CODESTAGE FROM INSCFORMATION'+
        ' WHERE PFI_CODESTAGE<>"--CURSUS--" '+GetMillesimes+' '+
        ' GROUP BY PFI_CODESTAGE,PFI_MILLESIME',True);
        TobNbInscBud := Tob.Create('Nombre insc',Nil,-1);
        TobNbInscBud.LoadDetailDB('Nombre insc','','',Q,False);
        Ferme(Q);
        
        TobEtat := Tob.Create('Les salariés pour édition',Nil,-1);
        For i := 0 to TobBudget.Detail.Count-1 do
        begin
                Salarie := TobBudget.Detail[i].GetValue('PFI_SALARIE');
                CodeStage := TobBudget.Detail[i].GetValue('PFI_CODESTAGE');
                Millesime := TobBudget.Detail[i].GetValue('PFI_MILLESIME');
                TNbIB := TobNbInscBud.FindFirst(['PFI_CODESTAGE','PFI_MILLESIME'],[CodeStage,Millesime],False);
                If TNbIB <> Nil then NbInsc := TNbIB.GetValue('NBINSC')
                Else Exit;
                NbStaMax := TobBudget.Detail[i].GetValue('PST_NBSTAMAX');
                NbSession := NombreDeSessionFormationAPrevoir(NbStaMax,NbInsc);
                NbInscSal := TobBudget.Detail[i].GetValue('PFI_NBINSC');
                NbHeuresB := TobBudget.Detail[i].GetValue('PST_DUREESTAGE');
                NbHeuresB := NbHeuresB * NbInscSal;
                CoutPedag := TobBudget.Detail[i].GetValue('PST_COUTBUDGETE');
                CoutPedag := CoutPedag*NbSession*NbInscSal;
                If NbInsc > 0 then CoutPedag := arrondi(CoutPedag/NbInsc,2)
                else CoutPedag :=0;
                CoutAnim := TobBudget.Detail[i].GetValue('PST_COUTSALAIR');
                CoutAnim := CoutAnim*NbSession*NbInscSal;
                CoutAnim := arrondi(CoutAnim/NbInsc,2);
                CoutUnitaire := TobBudget.Detail[i].GetValue('PST_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire*NbInscSal;
                CoutSal := TobBudget.Detail[i].GetValue('PFI_COUTREELSAL');
                Frais := TobBudget.Detail[i].GetValue('PFI_FRAISFORFAIT');
                TotalBud := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                T := TobEtat.FindFirst(['PFI_SALARIE','PFI_CODESTAGE'],[Salarie,CodeStage],False);
                If T <> Nil then
                begin
                        T.PutValue('HEURESP',T.GetValue('HEURESP')+NbHeuresB);
                        T.PutValue('COUTP',T.GetValue('COUTP')+TotalBud);
                end
                Else
                begin
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSupValeur('HEURESR',0);
                        T.AddChampSupValeur('COUTR',0);
                        T.AddChampSupValeur('HEURESP',NbHeuresB);
                        T.AddChampSupValeur('COUTP',TotalBud);
                        If PGBundleHierarchie Then //PT13
                        Begin
                             T.AddChampSupValeur('PSA_SALARIE',TobBudget.Detail[i].GetValue('PSI_INTERIMAIRE'));
                             T.AddChampSupValeur('PSA_ETABLISSEMENT',TobBudget.Detail[i].GetValue('PSI_ETABLISSEMENT'));
                             T.AddChampSupValeur('PSA_LIBELLEEMPLOI',TobBudget.Detail[i].GetValue('PSI_LIBELLEEMPLOI'));
                             T.AddChampSupValeur('PSA_DATEENTREE',TobBudget.Detail[i].GetValue('PSI_DATEENTREE'));
                             T.AddChampSupValeur('PSA_TRAVAILN1',TobBudget.Detail[i].GetValue('PSI_TRAVAILN1'));
                             T.AddChampSupValeur('PSA_TRAVAILN2',TobBudget.Detail[i].GetValue('PSI_TRAVAILN2'));
                             T.AddChampSupValeur('PSA_TRAVAILN3',TobBudget.Detail[i].GetValue('PSI_TRAVAILN3'));
                             T.AddChampSupValeur('PSA_TRAVAILN4',TobBudget.Detail[i].GetValue('PSI_TRAVAILN4'));
                        End
                        Else
                        Begin
                             T.AddChampSupValeur('PSA_SALARIE',TobBudget.Detail[i].GetValue('PSA_SALARIE'));
                             T.AddChampSupValeur('PSA_ETABLISSEMENT',TobBudget.Detail[i].GetValue('PSA_ETABLISSEMENT'));
                             T.AddChampSupValeur('PSA_LIBELLEEMPLOI',TobBudget.Detail[i].GetValue('PSA_LIBELLEEMPLOI'));
                             T.AddChampSupValeur('PSA_DATEENTREE',TobBudget.Detail[i].GetValue('PSA_DATEENTREE'));
                             T.AddChampSupValeur('PSA_TRAVAILN1',TobBudget.Detail[i].GetValue('PSA_TRAVAILN1'));
                             T.AddChampSupValeur('PSA_TRAVAILN2',TobBudget.Detail[i].GetValue('PSA_TRAVAILN2'));
                             T.AddChampSupValeur('PSA_TRAVAILN3',TobBudget.Detail[i].GetValue('PSA_TRAVAILN3'));
                             T.AddChampSupValeur('PSA_TRAVAILN4',TobBudget.Detail[i].GetValue('PSA_TRAVAILN4'));
                        End;
                        T.AddChampSupValeur('PST_CODESTAGE',TobBudget.Detail[i].GetValue('PST_CODESTAGE'));
                        T.AddChampSupValeur('PST_NATUREFORM',TobBudget.Detail[i].GetValue('PST_NATUREFORM'));
                        T.AddChampSupValeur('PST_CENTREFORMGU',TobBudget.Detail[i].GetValue('PST_CENTREFORMGU'));
                        T.AddChampSupValeur('PST_FORMATION1',TobBudget.Detail[i].GetValue('PST_FORMATION1'));
                        T.AddChampSupValeur('PST_FORMATION2',TobBudget.Detail[i].GetValue('PST_FORMATION2'));
                        T.AddChampSupValeur('PST_FORMATION3',TobBudget.Detail[i].GetValue('PST_FORMATION3'));
                        T.AddChampSupValeur('PST_FORMATION4',TobBudget.Detail[i].GetValue('PST_FORMATION4'));
                        T.AddChampSupValeur('PST_FORMATION5',TobBudget.Detail[i].GetValue('PST_FORMATION5'));
                        T.AddChampSupValeur('PST_FORMATION6',TobBudget.Detail[i].GetValue('PST_FORMATION6'));
                        T.AddChampSupValeur('PST_FORMATION7',TobBudget.Detail[i].GetValue('PST_FORMATION7'));
                        T.AddChampSupValeur('PST_FORMATION8',TobBudget.Detail[i].GetValue('PST_FORMATION8'));
                end;
        end;
        TobNbInscBud.Free;
        TobBudget.Free;
        
        If PGBundleHierarchie Then //PT13
        Begin
        	Where := DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PSI',False,True); //PT18
	        Q := OpenSQL('SELECT PSS_COUTPEDAG,PSS_COUTSALAIR,PSS_COUTUNITAIRE,PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_RESPONSFOR'+
	        ',PFO_NBREHEURE,PFO_FRAISREEL,PFO_COUTREELSAL,PFO_NBREHEURE'+
	        ',PSI_INTERIMAIRE,PSI_ETABLISSEMENT,PSI_LIBELLEEMPLOI,PSI_DATEENTREE,PSI_TRAVAILN1,PSI_TRAVAILN2,PSI_TRAVAILN3,PSI_TRAVAILN4'+
	        ',PST_CODESTAGE,PST_NATUREFORM,PST_CENTREFORMGU,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
	        ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8'+
	        ' FROM FORMATIONS '+
	        'LEFT JOIN INTERIMAIRES ON PFO_SALARIE=PSI_INTERIMAIRE '+
	        'LEFT JOIN STAGE ON PST_CODESTAGE=PFO_CODESTAGE AND PST_MILLESIME=PFO_MILLESIME '+
	        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME '+WherePFO+' '+Where,True)
	    End
        Else
	        Q := OpenSQL('SELECT PSS_COUTPEDAG,PSS_COUTSALAIR,PSS_COUTUNITAIRE,PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_RESPONSFOR'+
	        ',PFO_NBREHEURE,PFO_FRAISREEL,PFO_COUTREELSAL,PFO_NBREHEURE'+
	        ',PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_DATEENTREE,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4'+
	        ',PST_CODESTAGE,PST_NATUREFORM,PST_CENTREFORMGU,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
	        ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8'+
	        ' FROM FORMATIONS '+
	        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
	        'LEFT JOIN STAGE ON PST_CODESTAGE=PFO_CODESTAGE AND PST_MILLESIME=PFO_MILLESIME '+
	        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME '+WherePFO,True);
        TobRealise := Tob.Create('Les inscriptions au formations',Nil,-1);
        TobRealise.LoadDetailDB('FORMATIONS','','',Q,False);
        Ferme(Q);
        
        Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) AS NBSAL,PFO_ORDRE,PFO_MILLESIME,PFO_CODESTAGE FROM FORMATIONS WHERE '+
        'PFO_DATEDEBUT>="'+UsDatEtime(DateDebut)+'" AND PFO_DATEDEBUT<="'+UsDateTime(Datefin)+'" AND PFO_EFFECTUE="X"'+
        ' GROUP BY PFO_ORDRE,PFO_CODESTAGE,PFO_MILLESIME',True);
        TobNbInscRea := Tob.Create('Les Sessions',Nil,-1);
        TobNbInscRea.LoadDetailDB('SESSIONSTAGE','','',Q,False);
        Ferme(Q);
        
        For i := 0 to TobRealise.detail.Count-1 do
        begin
                Salarie := TobRealise.Detail[i].GetValue('PFO_SALARIE');
                CodeStage := TobRealise.Detail[i].GetValue('PFO_CODESTAGE');
                Millesime := TobRealise.Detail[i].GetValue('PFO_MILLESIME');
                NbHeuresR := TobRealise.Detail[i].GetValue('PFO_NBREHEURE');
                TNbIR := TobNbInscRea.FindFirst(['PFO_CODESTAGE','PFO_MILLESIME','PFO_ORDRE'],[CodeStage,Millesime,TobRealise.Detail[i].GetValue('PFO_ORDRE')],False);
                If TNbIR <> Nil then NbHSession := TNbIR.GetValue('NBSAL')
                Else Exit;
                CoutPedag := TobRealise.Detail[i].GetValue('PSS_COUTPEDAG');
                If NbHSession> 0 then CoutPedag := arrondi(NbHeuresR*(CoutPedag/NbHSession),2)
                else CoutPedag := 0;
                CoutAnim := TobRealise.Detail[i].GetValue('PSS_COUTSALAIR');
                If NbHSession> 0 then CoutAnim := arrondi(NbHeuresR*(CoutAnim/NbHSession),2)
                else CoutAnim := 0;
                CoutUnitaire := TobRealise.Detail[i].GetValue('PSS_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire;
                CoutSal := TobRealise.Detail[i].GetValue('PFO_COUTREELSAL');
                Frais := TobRealise.Detail[i].GetValue('PFO_FRAISREEL');
                TotalRea := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                If PGBundleHierarchie Then //PT13
                     T := TobEtat.FindFirst(['PSI_INTERIMAIRE','PST_CODESTAGE'],[Salarie,CodeStage],False)
                Else
                     T := TobEtat.FindFirst(['PSA_SALARIE','PST_CODESTAGE'],[Salarie,CodeStage],False);
                If T <> Nil then
                begin
                        T.PutValue('HEURESR',T.GetValue('HEURESR') + NbHeuresR);
                        T.PutValue('COUTR',T.GetValue('COUTR')+TotalRea);
                end
                Else
                begin
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSupValeur('HEURESR',NbHeuresR);
                        T.AddChampSupValeur('COUTR',TotalRea);
                        T.AddChampSupValeur('HEURESP',0);
                        T.AddChampSupValeur('COUTP',0);
                        If PGBundleHierarchie Then //PT13
                        Begin
                             T.AddChampSupValeur('PSA_SALARIE',TobRealise.Detail[i].GetValue('PSI_INTERIMAIRE'));
                             T.AddChampSupValeur('PSA_ETABLISSEMENT',TobRealise.Detail[i].GetValue('PSI_ETABLISSEMENT'));
                             T.AddChampSupValeur('PSA_LIBELLEEMPLOI',TobRealise.Detail[i].GetValue('PSI_LIBELLEEMPLOI'));
                             T.AddChampSupValeur('PSA_DATEENTREE',TobRealise.Detail[i].GetValue('PSI_DATEENTREE'));
                             T.AddChampSupValeur('PSA_TRAVAILN1',TobRealise.Detail[i].GetValue('PSI_TRAVAILN1'));
                             T.AddChampSupValeur('PSA_TRAVAILN2',TobRealise.Detail[i].GetValue('PSI_TRAVAILN2'));
                             T.AddChampSupValeur('PSA_TRAVAILN3',TobRealise.Detail[i].GetValue('PSI_TRAVAILN3'));
                             T.AddChampSupValeur('PSA_TRAVAILN4',TobRealise.Detail[i].GetValue('PSI_TRAVAILN4'));
                        End
                        Else
                        Begin
                             T.AddChampSupValeur('PSA_SALARIE',TobRealise.Detail[i].GetValue('PSA_SALARIE'));
                             T.AddChampSupValeur('PSA_ETABLISSEMENT',TobRealise.Detail[i].GetValue('PSA_ETABLISSEMENT'));
                             T.AddChampSupValeur('PSA_LIBELLEEMPLOI',TobRealise.Detail[i].GetValue('PSA_LIBELLEEMPLOI'));
                             T.AddChampSupValeur('PSA_DATEENTREE',TobRealise.Detail[i].GetValue('PSA_DATEENTREE'));
                             T.AddChampSupValeur('PSA_TRAVAILN1',TobRealise.Detail[i].GetValue('PSA_TRAVAILN1'));
                             T.AddChampSupValeur('PSA_TRAVAILN2',TobRealise.Detail[i].GetValue('PSA_TRAVAILN2'));
                             T.AddChampSupValeur('PSA_TRAVAILN3',TobRealise.Detail[i].GetValue('PSA_TRAVAILN3'));
                             T.AddChampSupValeur('PSA_TRAVAILN4',TobRealise.Detail[i].GetValue('PSA_TRAVAILN4'));
                        End;
                        T.AddChampSupValeur('PST_CODESTAGE',TobRealise.Detail[i].GetValue('PST_CODESTAGE'));
                        T.AddChampSupValeur('PST_NATUREFORM',TobRealise.Detail[i].GetValue('PST_NATUREFORM'));
                        T.AddChampSupValeur('PST_CENTREFORMGU',TobRealise.Detail[i].GetValue('PST_CENTREFORMGU'));
                        T.AddChampSupValeur('PST_FORMATION1',TobRealise.Detail[i].GetValue('PST_FORMATION1'));
                        T.AddChampSupValeur('PST_FORMATION2',TobRealise.Detail[i].GetValue('PST_FORMATION2'));
                        T.AddChampSupValeur('PST_FORMATION3',TobRealise.Detail[i].GetValue('PST_FORMATION3'));
                        T.AddChampSupValeur('PST_FORMATION4',TobRealise.Detail[i].GetValue('PST_FORMATION4'));
                        T.AddChampSupValeur('PST_FORMATION5',TobRealise.Detail[i].GetValue('PST_FORMATION5'));
                        T.AddChampSupValeur('PST_FORMATION6',TobRealise.Detail[i].GetValue('PST_FORMATION6'));
                        T.AddChampSupValeur('PST_FORMATION7',TobRealise.Detail[i].GetValue('PST_FORMATION7'));
                        T.AddChampSupValeur('PST_FORMATION8',TobRealise.Detail[i].GetValue('PST_FORMATION8'));
                end;
        end;
        TobRealise.Free;
        TobNbInscRea.Free;
        TobEtat.Detail.Sort('PSA_SALARIE');
        StPages := '';
       {$IFDEF EAGLCLIENT}
        StPages := AglGetCriteres (Pages, FALSE);
        {$ENDIF}
        //PT5 - Début
//        If GetCheckBoxState('FLISTE') = CbChecked then LanceEtatTOB('E','PFO','PL1',TobEtat,True,True,False,Pages,' ','',False,0,StPages)
//        else LanceEtatTOB('E','PFO','PL1',TobEtat,True,False,False,Pages,' ','',False,0,StPages);
//        TobEtat.free;
        TFQRS1(Ecran).LaTob:= TobEtat;
        //PT5 - Fin
end;

procedure TOF_PGEDITFORREABUD.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDITFORREABUD.SalarieElipsisClick(Sender : TObject);
begin
    //If PGBundleInscFormation Then //PT20
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
end;

Function TOF_PGEDITFORREABUD.ChampRuptureForm(TypeRupt,ValeurRupt : String) : String;
var Tablette : String;
    j : Integer;
begin
        If ValeurRupt = 'PFO_CODESTAGE' then Tablette := 'PGSTAGEFORM';
        If ValeurRupt = 'PFO_SALARIE' then 
        Begin
        	If PGBundleHierarchie Then //PT18
        		Tablette := 'PGSALARIEINT'
        	Else
        		Tablette := 'PGSALARIE';
        End;
        If ValeurRupt = 'PFO_ETABLISSEMENT' then Tablette := 'TTETABLISSEMENT';
        If ValeurRupt = 'PFO_LIBEMPLOIFOR' then Tablette := 'PGLIBEMPLOI';
        If ValeurRupt = 'PFO_NATUREFORM' then Tablette := 'PGNATUREFORM';
        If ValeurRupt = 'PFO_LIEUFORM' then Tablette := 'PGLIEUFORMATION';
        For j := 1 to 8 do
        begin
                If ValeurRupt = 'PFO_FORMATION'+IntToStr(j) then
                begin
                        Tablette := 'PGFORMATION'+IntToStr(j);
                end;
        end;
        For j := 1 to 4 do
        begin
                If ValeurRupt = 'PFO_TRAVAILN'+IntToStr(j) then
                begin
                Tablette := 'PGTRAVAILN'+IntToStr(j);
                end;
        end;
        Result := Tablette;
end;

//PT5 - Début
procedure TOF_PGEDITFORREABUD.OnClose;
begin
  Inherited ;
  If (TobEtat <> nil) then FreeAndNil (TobEtat);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 24/04/2007
Modifié le ... :   /  /    
Description .. : Rend (in)visible la checkbox du niveau de rupture
Suite ........ : précédent en fonction de la valeur saisie. Effectue également
Suite ........ : le contrôle de cohérence des ruptures.
Mots clefs ... : 
*****************************************************************}
Procedure TOF_PGEDITFORREABUD.OnChangeRuptures (Sender : TObject);
Var
  i,j : Integer;
  Affiche : boolean;
  Valeur : String;
Begin
     // Récupération du niveau de rupture
     i := StrToInt(Copy(TControl(Sender).Name,6,1));
     // Détermine s'il faut cacher ou afficher la checkbox de saut de page
     Affiche := (THValComboBox(GetControl('VRUPT'+IntToStr(i))).Value <> '');

     // Contrôle de cohérence des ruptures
     For j:= 2 To 3 Do
     Begin
        If (THValComboBox(GetControl('VRUPT'+IntToStr(j))).Value = '') And
           (THValComboBox(GetControl('VRUPT'+IntToStr(j+1))).Value <> '') Then
          Begin
               PGIBox('Le niveau de rupture '+IntToStr(j)+' doit être renseigné',Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End
     End;

     Valeur := THValComboBox(GetControl('VRUPT'+IntToStr(i))).Value;
     For j:=1 To 4 Do
     Begin
          If (i <> j) And (Valeur = THValComboBox(GetControl('VRUPT'+IntToStr(j))).Value) And (Valeur <> '') Then
          Begin
               PGIBox('La rupture '+IntToStr(i)+' doit être différente de la '+IntToStr(j),Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End;
     End;

     GetControl('BValider').Enabled := True;

     // Afficher/cacher la checkbox
     //If (i>1) Then  //PT16
     Begin
          TCheckBox(GetControl('CSAUTRUPT'+IntToStr(i))).Visible := Affiche; //PT16
          // Si on cache, on remet également l'état à "décoché"
          If (Affiche = False) Then TCheckBox(GetControl('CSAUTRUPT'+IntToStr(i))).Checked := Affiche; //PT16
     End;
End;
//PT5 - Fin

//PT6 - Début
{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 24/04/2007
Modifié le ... :   /  /    
Description .. : Edition du comparatif pluri-annuel
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITFORREABUD.EditionComparPluriAnnuel;
var TobNbInscBud,TobNbInscRea,TobBudget,TobRealise : Tob;
     TNbIB,TNbIR,T : Tob;
     Q : TQuery;
     DateDebut,DateFin : TDateTime;
     i,NbRupt : Integer;
     Salarie,Where,WherePFI,WherePFO,Millesime,CodeStage,StPages,StrTemp,WhereDos : String;
     Rupt1,Rupt2,Rupt3,Rupt4 : String;
     ChampRuptRea1,ChampRuptRea2,ChampRuptRea3,ChampRuptRea4 : String;
     ChampRuptBud1,ChampRuptBud2,ChampRuptBud3,ChampRuptBud4 : String;
     ValeurRupt1,ValeurRupt2,ValeurRupt3,ValeurRupt4 : String;
     Tablette1,Tablette2,Tablette3,Tablette4,Libelle : String;
     NbStaMax,NbInscSal,NbInsc : Integer;
     NbHeuresR,NbHeuresB,CoutPedag,CoutAnim,CoutUnitaire,CoutSal,Frais,NbHSession : Double;
     TotalRea,TotalBud,NbSession : Double;
     Pages : TPageControl;
     ListeChampRea,ListeChampBud : String;
     a : Integer;
begin
        NbInsc := 0;
        NbHSession := 0;
        Pages := TPageControl(GetControl('Pages'));
        SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFO')); //PT18
        Rupt1 := '';
        Rupt2 := '';
        Rupt3 := '';
        Rupt4 := '';

        // Démarrage de l'écran d'attente
        InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', StrToInt(GetControlText('NBEXERCICES'))+1, FALSE, TRUE);

        { Gestion des ruptures }

        NbRupt := 1;
        SetControlText('XX_RUPTURE2','');
        SetControlText('XX_RUPTURE3','');
        SetControlText('XX_RUPTURE4','');
        Rupt1 := GetControltext('VRUPT1');SetControlText('XX_RUPTURE1','RUPTURE1');
        If GetControlText('VRUPT2') <> '' then
        Begin
             Rupt2 := GetControltext('VRUPT2');SetControlText('XX_RUPTURE2','RUPTURE2');NbRupt := NbRupt + 1;
        End;
        If GetControlText('VRUPT3') <> '' then
        Begin
             Rupt3 := GetControltext('VRUPT3');SetControlText('XX_RUPTURE3','RUPTURE3');NbRupt := NbRupt + 1;
        End;
        If GetControlText('VRUPT4') <> '' then
        Begin
             Rupt4 := GetControltext('VRUPT4');SetControlText('XX_RUPTURE4','RUPTURE4');NbRupt := NbRupt + 1;
        End;
        ChampRuptRea1 := Rupt1;
        ChampRuptRea2 := Rupt2;
        ChampRuptRea3 := Rupt3;
        ChampRuptRea4 := Rupt4;
        ChampRuptBud1 := ConvertPrefixe(Rupt1,'PSS','PST');
        ChampRuptBud2 := ConvertPrefixe(Rupt2,'PSS','PST');
        ChampRuptBud3 := ConvertPrefixe(Rupt3,'PSS','PST');
        ChampRuptBud4 := ConvertPrefixe(Rupt4,'PSS','PST');
        ChampRuptBud1 := ConvertPrefixe(ChampRuptBud1,'PFO','PFI');
        ChampRuptBud2 := ConvertPrefixe(ChampRuptBud2,'PFO','PFI');
        ChampRuptBud3 := ConvertPrefixe(ChampRuptBud3,'PFO','PFI');
        ChampRuptBud4 := ConvertPrefixe(ChampRuptBud4,'PFO','PFI');
        If Rupt1 <> '' then ListeChampRea := ListeChampRea+','+Rupt1;
        If Rupt2 <> '' then ListeChampRea := ListeChampRea+','+Rupt2;
        If Rupt3 <> '' then ListeChampRea := ListeChampRea+','+Rupt3;
        If Rupt4 <> '' then ListeChampRea := ListeChampRea+','+Rupt4;
        If ChampRuptBud1 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud1;
        If ChampRuptBud2 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud2;
        If ChampRuptBud3 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud3;
        If ChampRuptBud4 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud4;
        Tablette1 := ChampRuptureForm('TAB',Rupt1);
        Tablette2 := ChampRuptureForm('TAB',Rupt2);
        Tablette3 := ChampRuptureForm('TAB',Rupt3);
        Tablette4 := ChampRuptureForm('TAB',Rupt4);

        // Détermination de la clause WHERE
        Where     := '';
        Where     := RecupWhereCritere(Pages);
        If (ChampRuptBud1 = 'PFI_SALARIE') or (ChampRuptBud2 = 'PFI_SALARIE') or (ChampRuptBud3 = 'PFI_SALARIE') or (ChampRuptBud4 = 'PFI_SALARIE') then
             If (Where <> '') Then
                  Where := Where + ' AND PFI_SALARIE <> ""'
             Else
                  Where := 'WHERE PFI_SALARIE <> ""';

        // Adaptation de la clause au prévisionnel (tables différentes)
        WherePFI  := ConvertPrefixe(Where,'PFO','PFI');
        WherePFI  := ConvertPrefixe(WherePFI,'PSS','PST');
        If WherePFI <> '' then WherePFI := WherePFI + 'AND '
        Else  WherePFI := 'WHERE ';
        WherePFI := WherePFI + 'PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--"';
        If THValComboBox(GetControl('MILLESIMEUNIQUE')).Value <> '' Then //PT18
        	WherePFO := WherePFI + ' AND PFI_MILLESIME="'+THValComboBox(GetControl('MILLESIMEUNIQUE')).Value+'"';

        // Mise à jour des champs cachés utilisés dans le titre de l'état
        SetControlText ('TXTMILLESIMES',THValComboBox(GetControl('MILLESIMEUNIQUE')).Value);
        SetControlText ('TXTANNEEINT',IntToStr(StrToInt(THValComboBox(GetControl('MILLESIMEUNIQUE')).Value)-1));
        SetControlText ('TXTANNEEDEPART',IntToStr(StrToInt(THValComboBox(GetControl('MILLESIMEUNIQUE')).Value)-2));

        { Calcul des coûts au prévisionnel }

        // Recherche des caractéristiques de chaque formation
        Q := OpenSQL('SELECT PFI_MILLESIME,PFI_CODESTAGE,PFI_NBINSC,PFI_COUTREELSAL,PFI_FRAISFORFAIT,PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,PFI_RESPONSFOR'+
        ',PST_DUREESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX'+ListeChampBud+
        ' FROM INSCFORMATION LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+WherePFI,True);
        TobBudget := Tob.Create('Les inscriptions au budget',Nil,-1);
        TobBudget.LoadDetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);
        
        // Recherche du nombre d'inscrits afin de calculer les coûts
        If PGBundleInscFormation Then //PT18
        	WhereDos := DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFI',False,True)
        Else
        	WhereDos := '';
        Q := OpenSQL('SELECT SUM(PFI_NBINSC) AS NBINSC,PFI_MILLESIME,PFI_CODESTAGE FROM INSCFORMATION'+
        ' WHERE PFI_CODESTAGE<>"--CURSUS--" AND PFI_MILLESIME="'+THValComboBox(GetControl('MILLESIMEUNIQUE')).Value+'"'+WhereDos+
        ' GROUP BY PFI_CODESTAGE,PFI_MILLESIME',True);
        TobNbInscBud := Tob.Create('Nombre insc',Nil,-1);
        TobNbInscBud.LoadDetailDB('Nombre insc','','',Q,False);
        Ferme(Q);

        MoveCurProgressForm('Données du prévisionnel');

        TobEtat := Tob.Create('Les salariés pour édition',Nil,-1);
        For i := 0 to TobBudget.Detail.Count-1 do
        begin
                Salarie := TobBudget.Detail[i].GetValue('PFI_SALARIE');
                CodeStage := TobBudget.Detail[i].GetValue('PFI_CODESTAGE');
                Millesime := TobBudget.Detail[i].GetValue('PFI_MILLESIME');
                TNbIB := TobNbInscBud.FindFirst(['PFI_CODESTAGE','PFI_MILLESIME'],[CodeStage,Millesime],False);
                If TNbIB <> Nil then NbInsc := TNbIB.GetValue('NBINSC')
                Else Begin FiniMoveProgressForm;Break;End; //PT18
                NbStaMax := TobBudget.Detail[i].GetValue('PST_NBSTAMAX');
                NbSession := NombreDeSessionFormationAPrevoir(NbStaMax,NbInsc);
                NbInscSal := TobBudget.Detail[i].GetValue('PFI_NBINSC');
                // Calcul du nombre d'heures
                NbHeuresB := TobBudget.Detail[i].GetValue('PST_DUREESTAGE');
                NbHeuresB := NbHeuresB * NbInscSal;
                // Calcul du coût
                CoutPedag := TobBudget.Detail[i].GetValue('PST_COUTBUDGETE');
                CoutPedag := CoutPedag*NbSession*NbInscSal;
                If NbInsc > 0 then CoutPedag := arrondi(CoutPedag/NbInsc,2)
                else CoutPedag :=0;
                CoutAnim := TobBudget.Detail[i].GetValue('PST_COUTSALAIR');
                CoutAnim := CoutAnim*NbSession*NbInscSal;
                CoutAnim := arrondi(CoutAnim/NbInsc,2);
                CoutUnitaire := TobBudget.Detail[i].GetValue('PST_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire*NbInscSal;
                CoutSal := TobBudget.Detail[i].GetValue('PFI_COUTREELSAL');
                Frais := TobBudget.Detail[i].GetValue('PFI_FRAISFORFAIT');
                TotalBud := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                // Détermination des ruptures afin de création l'enregistrement dans la TOB
                ValeurRupt1 := '';
                ValeurRupt2 := '';
                ValeurRupt3 := '';
                ValeurRupt4 := '';
                If NbRupt > 3 then ValeurRupt4 := TobBudget.Detail[i].GetValue(ChampRuptBud4);
                If NbRupt > 2 then ValeurRupt3 := TobBudget.Detail[i].GetValue(ChampRuptBud3);
                If NbRupt > 1 then ValeurRupt2 := TobBudget.Detail[i].GetValue(ChampRuptBud2);
                ValeurRupt1 := TobBudget.Detail[i].GetValue(ChampRuptBud1);
                If NbRupt=4 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt4,ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=3 then T := TobEtat.FindFirst(['RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=2 then T := TobEtat.FindFirst(['RUPTURE2','RUPTURE1'],[ValeurRupt2,ValeurRupt1],False)
                Else T := TobEtat.FindFirst(['RUPTURE1'],[ValeurRupt1],False);
                If T <> Nil then
                begin
                        // L'enregistrement existe déjà : on met à jour
                        T.PutValue('HEURESPREV',T.GetValue('HEURESPREV')+NbHeuresB);
                        T.PutValue('COUTSPREV',T.GetValue('COUTSPREV')+TotalBud);
                end
                Else
                begin
                        // On crée un nouvel enregistrement dans la TOB
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSupValeur('RUPTURE1','');T.AddChampSupValeur('RUPTURE2','');
                        T.AddChampSupValeur('RUPTURE3','');T.AddChampSupValeur('RUPTURE4','');
                        T.AddChampSupValeur('LIBRUPTURE1','');T.AddChampSupValeur('LIBRUPTURE2','');
                        T.AddChampSupValeur('LIBRUPTURE3','');T.AddChampSupValeur('LIBRUPTURE4','');
                        T.AddChampSupValeur('HEURESPREV',0);T.AddChampSupValeur('COUTSPREV',0);
                        T.PutValue('HEURESPREV',NbHeuresB);T.PutValue('COUTSPREV',TotalBud);
                        If NbRupt > 3 then T.PutValue('RUPTURE4',TobBudget.Detail[i].GetValue(ChampRuptBud4));
                        If NbRupt > 2 then T.PutValue('RUPTURE3',TobBudget.Detail[i].GetValue(ChampRuptBud3));
                        If NbRupt > 1 then T.PutValue('RUPTURE2',TobBudget.Detail[i].GetValue(ChampRuptBud2));
                                           T.PutValue('RUPTURE1',TobBudget.Detail[i].GetValue(ChampRuptBud1));
                        // Création dès le départ des champs réalisés à vide car leur remplissage est
                        // variable selon le nombre d'exercices choisis par l'utilisateur
                        T.AddChampSupValeur('HEURESREAL1',0);T.AddChampSupValeur('COUTSREAL1',0);
                        T.AddChampSupValeur('HEURESREAL2',0);T.AddChampSupValeur('COUTSREAL2',0);
                        T.AddChampSupValeur('HEURESREAL3',0);T.AddChampSupValeur('COUTSREAL3',0);
                end;
        end;
        TobNbInscBud.Free;
        TobBudget.Free;

        { Calcul des coûts au réel }
        // Ici, on effectue l'ensemble du traitement de calcul des coûts pour chaque année demandée
        For a := 1 To StrToInt(GetControlText('NBEXERCICES')) Do
        Begin
           MoveCurProgressForm('Données de l''exercice '+IntToStr(StrToInt(GetControlText('TXTMILLESIMES'))-(a-1)));

           // Détermination de la clause WHERE
           If Where <> '' then
           Begin
               WherePFO := Where + ' AND ';
               WherePFO := ConvertPrefixe(WherePFO,'PFI','PFO'); // Pour les ruptures sur salarié
           End
           Else WherePFO := 'WHERE ';
           DateDebut := StrToDate('01/01/'+IntToStr(StrToInt(GetControlText('TXTMILLESIMES'))-(a-1)));
           DateFin := StrToDate('31/12/'+IntToStr(StrToInt(GetControlText('TXTMILLESIMES'))-(a-1)));
           WherePFO := WherePFO + ' PFO_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'" AND PFO_EFFECTUE="X"';
           WherePFO := WherePFO + ' AND PFO_PGTYPESESSION<>"SOS"'; //PT8
           
           // Recherche des caractéristiques de chaque stage
           Q := OpenSQL('SELECT PSS_COUTPEDAG,PSS_COUTSALAIR,PSS_COUTUNITAIRE,PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_RESPONSFOR'+
           ',PFO_NBREHEURE,PFO_FRAISREEL,PFO_COUTREELSAL,PFO_NBREHEURE'+ListeChampRea+
           ' FROM FORMATIONS LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+WherePFO,True);
           TobRealise := Tob.Create('Les inscriptions au formations',Nil,-1);
           TobRealise.LoadDetailDB('FORMATIONS','','',Q,False);
           Ferme(Q);
           
           // Recherche du nombre d'heures réellement effectuées pour chaque session
	        If PGBundleInscFormation Then //PT18
	        	WhereDos := DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFO',False,True)
	        Else
	        	WhereDos := '';
           Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) AS NBSAL,PFO_ORDRE,PFO_MILLESIME,PFO_CODESTAGE FROM FORMATIONS WHERE '+
           'PFO_DATEDEBUT>="'+UsDatetime(DateDebut)+'" AND PFO_DATEDEBUT<="'+UsDateTime(Datefin)+'" AND PFO_EFFECTUE="X"'+WhereDos+
           ' GROUP BY PFO_ORDRE,PFO_CODESTAGE,PFO_MILLESIME',True);
           TobNbInscRea := Tob.Create('Les Sessions',Nil,-1);
           TobNbInscRea.LoadDetailDB('SESSIONSTAGE','','',Q,False);
           Ferme(Q);

           For i := 0 to TobRealise.detail.Count-1 do
           begin
                Salarie := TobRealise.Detail[i].GetValue('PFO_SALARIE');
                CodeStage := TobRealise.Detail[i].GetValue('PFO_CODESTAGE');
                Millesime := TobRealise.Detail[i].GetValue('PFO_MILLESIME');
                // Calcul du nombre d'heures
                NbHeuresR := TobRealise.Detail[i].GetValue('PFO_NBREHEURE');
                TNbIR := TobNbInscRea.FindFirst(['PFO_CODESTAGE','PFO_MILLESIME','PFO_ORDRE'],[CodeStage,Millesime,TobRealise.Detail[i].GetValue('PFO_ORDRE')],False);
                If TNbIR <> Nil then NbHSession := TNbIR.GetValue('NBSAL')
                Else Begin FiniMoveProgressForm;Break;End; //PT18
                // Calcul des coûts
                CoutPedag := TobRealise.Detail[i].GetValue('PSS_COUTPEDAG');
                If NbHSession> 0 then CoutPedag := arrondi(NbHeuresR*(CoutPedag/NbHSession),2)
                else CoutPedag := 0;
                CoutAnim := TobRealise.Detail[i].GetValue('PSS_COUTSALAIR');
                If NbHSession> 0 then CoutAnim := arrondi(NbHeuresR*(CoutAnim/NbHSession),2)
                else CoutAnim := 0;
                CoutUnitaire := TobRealise.Detail[i].GetValue('PSS_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire;
                CoutSal := TobRealise.Detail[i].GetValue('PFO_COUTREELSAL');
                Frais := TobRealise.Detail[i].GetValue('PFO_FRAISREEL');
                TotalRea := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                // Détermination des ruptures afin de création l'enregistrement dans la TOB
                ValeurRupt1 := '';
                ValeurRupt2 := '';
                ValeurRupt3 := '';
                ValeurRupt4 := '';
                If NbRupt > 3 then ValeurRupt4 := TobRealise.Detail[i].GetValue(ChampRuptRea4);
                If NbRupt > 2 then ValeurRupt3 := TobRealise.Detail[i].GetValue(ChampRuptRea3);
                If NbRupt > 1 then ValeurRupt2 := TobRealise.Detail[i].GetValue(ChampRuptRea2);
                ValeurRupt1 := TobRealise.Detail[i].GetValue(ChampRuptRea1);
                If NbRupt=4 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt4,ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=3 then T := TobEtat.FindFirst(['RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=2 then T := TobEtat.FindFirst(['RUPTURE2','RUPTURE1'],[ValeurRupt2,ValeurRupt1],False)
                Else T := TobEtat.FindFirst(['RUPTURE1'],[ValeurRupt1],False);
                If T <> Nil then
                begin
                        // L'enregistrement existe déjà : on met à jour
                        T.PutValue('HEURESREAL'+IntToStr(a),T.GetValue('HEURESREAL'+IntToStr(a))+NbHeuresR);
                        T.PutValue('COUTSREAL'+IntToStr(a),T.GetValue('COUTSREAL'+IntToStr(a))+TotalRea);
                end
                Else
                begin
                        // On crée un nouvel enregistrement dans la TOB
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSupValeur('RUPTURE1','');T.AddChampSupValeur('RUPTURE2','');
                        T.AddChampSupValeur('RUPTURE3','');T.AddChampSupValeur('RUPTURE4','');
                        T.AddChampSupValeur('LIBRUPTURE1','');T.AddChampSupValeur('LIBRUPTURE2','');
                        T.AddChampSupValeur('LIBRUPTURE3','');T.AddChampSupValeur('LIBRUPTURE4','');
                        // Création des champs réalisés à vide car leur remplissage est
                        // variable selon le nombre d'exercices choisis par l'utilisateur
                        T.AddChampSupValeur('HEURESPREV',0);T.AddChampSupValeur('COUTSPREV',0);
                        T.AddChampSupValeur('HEURESREAL1',0);T.AddChampSupValeur('COUTSREAL1',0);
                        T.AddChampSupValeur('HEURESREAL2',0);T.AddChampSupValeur('COUTSREAL2',0);
                        T.AddChampSupValeur('HEURESREAL3',0);T.AddChampSupValeur('COUTSREAL3',0);
                        If NbRupt > 3 then T.PutValue('RUPTURE4',TobRealise.Detail[i].GetValue(ChampRuptRea4));
                        If NbRupt > 2 then T.PutValue('RUPTURE3',TobRealise.Detail[i].GetValue(ChampRuptRea3));
                        If NbRupt > 1 then T.PutValue('RUPTURE2',TobRealise.Detail[i].GetValue(ChampRuptRea2));
                                           T.PutValue('RUPTURE1',TobRealise.Detail[i].GetValue(ChampRuptRea1));
                        // Renseignement des valeurs en cours
                        T.PutValue('HEURESREAL'+IntToStr(a),NbHeuresR); T.PutValue('COUTSREAL'+IntToStr(a),TotalRea);
                end;
           end;

           TobRealise.Free;
           TobNbInscRea.Free;
        End; //for

        { Mise à jour de la TOB principale avec les libellés de ruptures }

        For i := 0 to TobEtat.detail.Count-1 do
        begin
                If NbRupt > 3 then
                begin
                        Libelle := RechDom(Tablette4,TobEtat.Detail[i].GetValue('RUPTURE4'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE4',Libelle);
                end;
                If NbRupt > 2 then
                begin
                        Libelle := RechDom(Tablette3,TobEtat.Detail[i].GetValue('RUPTURE3'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE3',Libelle);
                end;
                If NbRupt > 1 then
                begin
                        Libelle := RechDom(Tablette2,TobEtat.Detail[i].GetValue('RUPTURE2'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE2',Libelle);
                end;
                Libelle := RechDom(Tablette1,TobEtat.Detail[i].GetValue('RUPTURE1'),False);
                If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE1',Libelle);
        end;
        TobEtat.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;RUPTURE4');
        StPages := '';
        {$IFDEF EAGLCLIENT}
        StPages := AglGetCriteres (Pages, FALSE);
        {$ENDIF}

        TFQRS1(Ecran).LaTob:= TobEtat;

        // Fermeture de l'écran d'attente
        FiniMoveProgressForm;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 25/04/2007
Modifié le ... :   /  /
Description .. : Crée la clause where pour la sélection des millésimes
Suite ........ : et met à jour le champ caché TXTMILLESIMES utilisé par le 
Suite ........ : titre de l'état
Mots clefs ... : 
*****************************************************************}
function TOF_PGEDITFORREABUD.GetMillesimes(): String;
Var
  WhereMillesime,StrTemp,StrCopy,Titre : String;
begin
        StrCopy := GetControlText('MILLESIME');
        If (StrCopy = '') Or (StrCopy = '<<Tous>>') Then
        Begin
          WhereMillesime := '';
          Titre := 'complet';
        End
        Else
        Begin
            Titre := 'de ';
            While StrCopy <> '' Do
            Begin
               StrTemp := ReadTokenPipe (StrCopy,';');
               If WhereMillesime = '' Then
               Begin
                   WhereMillesime := 'AND (PFI_MILLESIME="'+StrTemp+'"';
                   Titre := Titre + StrTemp;
               End
               Else
               Begin
                   WhereMillesime := WhereMillesime + ' OR PFI_MILLESIME="'+StrTemp+'"';
                   Titre := Titre +','+StrTemp;
               End;
            End;
            WhereMillesime := WhereMillesime + ')';
        End;
        SetControlText('TXTMILLESIMES',Titre);
        Result := WhereMillesime;
end;
//PT6 - Fin

{TOF_PGEDITPLANFORMATION}
Procedure TOF_PGEDITPLANFORMATION.RespElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StWhere,StOrder : String;
{$ENDIF}
begin
	{$IFDEF EMANAGER}
        StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
        StOrder := 'PSI_LIBELLE';
        LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
	{$ELSE}
	ElipsisResponsableMultidos (Sender); //PT18
    {$ENDIF}
end;

procedure TOF_PGEDITPLANFORMATION.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

Procedure TOF_PGEDITPLANFORMATION.OnUpdate ;
var Pages : TPageControl;
    Where,SQL,OrderBy,Select : String;
    AvecLibSession,ForceCodeStage,AvecResp : Boolean;
    i : Integer;
begin
  Inherited ;
  OrderBy := '';
  AvecResp := False;
  If GetControlText('PSE_RESPONSFOR') <> '' then AvecResp := True;
  Pages := TPageControl(GetControl('Pages'));
  
  SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFO')); //PT18
  
  Where := RecupWhereCritere(Pages);
  If Where <> '' then Where := Where + ' AND PFO_EFFECTUE="X"'
  else Where := 'WHERE PFO_EFFECTUE="X"';
  
        For i := 1 to 4 do
        begin
                If GetControlText('VRUPT'+IntToStr(i)) <> '' then
                begin
                        If OrderBy = '' then OrderBy := ' ORDER BY '+GetControlText('VRUPT'+IntToStr(i))
                        else OrderBy := OrderBy+','+GetControltext('VRUPT'+IntToStr(i));
                        SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
                        If GetControlText('VRUPT'+IntToStr(i)) = 'PSE_RESPONSFOR' then AvecResp := True;
                end
                else SetControlText('XX_RUPTURE'+IntToStr(i),'');
        end;
        
        If LeMenu = '47525' then     //Suivi DIF
        begin
                //OrderBy := ' ORDER BY PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE';
                TFQRS1(Ecran).CodeEtat := 'PL2';
                TFQRS1(Ecran).FCodeEtat := 'PL2';
                Orderby := ' ORDER BY ';

                SQL := 'SELECT PFO_CODESTAGE,PFO_ORDRE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_NBREHEURE,PFO_SALARIE,PSS_LIBELLE,'+
                	   'PSA_ETABLISSEMENT AS ETABLISSEMENT,PSA_LIBELLEEMPLOI AS LIBELLEEMPLOI,PSA_TRAVAILN1 AS TRAVAILN1,PSA_TRAVAILN2 AS TRAVAILN2,PSA_TRAVAILN3 AS TRAVAILN3,PSA_TRAVAILN4 AS TRAVAILN4,PSA_DATEENTREE AS DATEENTREE '+
                	   'FROM FORMATIONS '+
	                   'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_MILLESIME=PSS_MILLESIME AND PSS_ORDRE=PFO_ORDRE ';
                
                (*If AvecResp = True then 
                	SQL := 'SELECT PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_NBREHEURE,PFO_SALARIE,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_TRAVAILN1,PFO_TRAVAILN2,PFO_TRAVAILN3,PFO_TRAVAILN4 FROM FORMATIONS '+
                        'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_MILLESIME=PSS_MILLESIME '+
                        'AND PSS_ORDRE=PFO_ORDRE '
                else 
                	SQL := 'SELECT * FROM FORMATIONS '+
                        'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_MILLESIME=PSS_MILLESIME '+
                        'AND PSS_ORDRE=PFO_ORDRE ';*)
                
				//PT22
				// N'affiche pas les salariés confidentiels
				If Where <> '' Then Where := Where + ' AND ';
				Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
				
                If PGBundleHierarchie Then //PT13
                Begin
                     SQL := SQL +'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFO_SALARIE ';
                     Orderby := OrderBy + 'PSI_INTERIMAIRE,';
                     SQL   := StringReplace (SQL,   'PSA_', 'PSI_', [rfIgnoreCase,rfReplaceAll]);
                     Where := StringReplace (Where, 'PSA_', 'PSI_', [rfIgnoreCase,rfReplaceAll]);
                End
                Else
                Begin
                     SQL := SQL +'LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE ';
                     Orderby := OrderBy + 'PSA_SALARIE,';
                End;
                
                Orderby := OrderBy +'PFO_CODESTAGE,PFO_ORDRE';
                     
                TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
        end;
        If LeMenu = '47531' then     //Suivi pédagogique stagiaire individuel
        begin
                {$IFDEF EMANAGER}
                If MultiNiveau then
                begin
                        AvecResp := True;
                        If Where <> '' then Where := Where + 'AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR AND PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                        ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
                        else Where := 'WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                        ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
                end;
                {$ENDIf}
                
                OrderBy := ' ORDER BY PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE';
                TFQRS1(Ecran).CodeEtat := 'PPI';
                TFQRS1(Ecran).FCodeEtat := 'PPI';
                
				//PT22
				// N'affiche pas les salariés confidentiels
				If Where <> '' Then Where := Where + ' AND ';
				Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
				
                //PT18 - Début
                If PGBundleHierarchie Then //PT13
                Begin
                     SQL := 'SELECT FORMATIONS.*,PSI_DATEENTREE AS DATEENTREE FROM FORMATIONS LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFO_SALARIE ';
                     Where := StringReplace (Where, 'PSA_', 'PSI_', [rfReplaceAll,rfIgnoreCase]);
                     // On change la propriété du dossier pour prendre en compte le salarié intérimaire du dossier 
                     // et pas seulement les formations du dossier
                     Where := StringReplace (Where, 'PFO_NODOSSIER', 'PSI_NODOSSIER', [rfReplaceAll,rfIgnoreCase]);
                End
                Else
                     SQL := 'SELECT FORMATIONS.*,PSA_DATEENTREE AS DATEENTREE FROM FORMATIONS LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE ';
                
                If AvecResp = True then SQL := SQL + ' LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE ';
                //PT18 - Fin
                     
                TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
        end;
        If LeMenu = '47532' then     //Suivi pédagogique stagiaire collectif
        begin
                TFQRS1(Ecran).CodeEtat := 'PPC';
                TFQRS1(Ecran).FCodeEtat := 'PPC';
                If AvecResp = True then SQL := 'SELECT * FROM FORMATIONS LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE '
                else SQL := 'SELECT * FROM FORMATIONS ';
                
				//PT22
				// N'affiche pas les salariés confidentiels
				If Where <> '' Then Where := Where + ' AND ';
				Where := Where + ' PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';

                If PGBundleHierarchie Then 
                Begin
                	Where := StringReplace (Where, 'PSA_SALARIE', 'PSI_INTERIMAIRE', [rfReplaceAll,rfIgnoreCase]); //PT22
                	Where := StringReplace (Where, 'PSA_', 'PSI_', [rfReplaceAll,rfIgnoreCase]); //PT18
                	Where := StringReplace (Where, 'SALARIES', 'INTERIMAIRES', [rfReplaceAll,rfIgnoreCase]);
                End;
                
                TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
        end;
        If LeMenu = '47551' then     //Suivi gestion individuel
        begin
                TFQRS1(Ecran).CodeEtat := 'PGI';
                TFQRS1(Ecran).FCodeEtat := 'PGI';
                OrderBy := ' ORDER BY PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE';
                
                SQL := 'SELECT * FROM FORMATIONS ';
                //PT22
				// N'affiche pas les salariés confidentiels
				If Where <> '' Then Where := Where + ' AND ';
				Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
				
                If PGBundleHierarchie Then //PT13
                Begin
                     SQL := SQL + 'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFO_SALARIE ';
                     Where := StringReplace (Where, 'PSA_', 'PSI_', [rfReplaceAll,rfIgnoreCase]);
                End
                Else
                     SQL := SQL + 'LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE ';

                If AvecResp = True then SQL := SQL + ' LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE ';
                
                
                TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
        end;
        If LeMenu = '47552' then     //Suivi gestion collectif
        begin
                TFQRS1(Ecran).CodeEtat := 'PGC';
                TFQRS1(Ecran).FCodeEtat := 'PGC';
                If AvecResp = True then SQL := 'SELECT * FROM FORMATIONS LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE '
                else SQL := 'SELECT * FROM FORMATIONS ';
                
                //PT22
				// N'affiche pas les salariés confidentiels
				If Where <> '' Then Where := Where + ' AND ';
				Where := Where + ' PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
				
                If PGBundleHierarchie Then 
                Begin
                	Where := StringReplace (Where, 'PSA_SALARIE', 'PSI_INTERIMAIRE', [rfReplaceAll,rfIgnoreCase]); //PT22
                	Where := StringReplace (Where, 'PSA_', 'PSI_', [rfReplaceAll,rfIgnoreCase]); //PT18
                	Where := StringReplace (Where, 'SALARIES', 'INTERIMAIRES', [rfReplaceAll,rfIgnoreCase]);
                End;
                
                TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
        end;
        If LeMenu = '47561' then //Frais stagiaires
        begin
                TFQRS1(Ecran).CodeEtat := 'PFS';
                TFQRS1(Ecran).FCodeEtat := 'PFS';
                ForceCodeStage := True;
                AvecLibSession := False;
                OrderBy := '';
                Select := '';
                
                If PGBundleHierarchie Then Where := StringReplace (Where, 'PSA_', 'PSI_', [rfReplaceAll,rfIgnoreCase]); //PT18
                
                For i := 1 to 4 do
                begin
                        SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
                        If GetControlText('VRUPT'+IntToStr(i)) <> '' then
                        begin
                                If GetControlText('VRUPT'+IntToStr(i)) = 'PFO_ORDRE' then  AvecLibSession := True;
                                If OrderBy <> '' then OrderBy := OrderBy + ',' + GetControlText('VRUPT'+IntToStr(i))
                                else OrderBy := ' ORDER BY '+ GetControlText('VRUPT'+IntToStr(i));
                                If (GetControlText('VRUPT'+IntToStr(i)) <> 'PFO_CODESTAGE') and (GetControlText('VRUPT'+IntToStr(i)) <> 'PFO_ORDRE') then
                                Select := Select+','+GetControlText('VRUPT'+IntToStr(i));

                        end;
                end;
                If AvecLibSession = True then
                begin
                        For i := 1 to 4 do
                        begin
                                If GetControlText('VRUPT'+IntToStr(i)) = 'PFO_CODESTAGE' then ForceCodeStage := False;
                        end;
                end
                else forceCodeStage := False;
                If ForceCodeStage = True then
                begin
                        If OrderBy <> '' then OrderBy := OrderBy + ',PFO_CODESTAGE'
                        else OrderBy := ' ORDER BY PFO_CODESTAGE';
                        SetControltext('XX_RUPTURE5','PFO_CODESTAGE');
                end
                Else SetControlText('XX_RUPTURE5','');
                
                //PT22
				// N'affiche pas les salariés confidentiels
				If Where <> '' Then Where := Where + ' AND ';
				If PGBundleHierarchie Then
					Where := Where + ' PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
				Else
					Where := Where + ' PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
                
                If not AvecLibSession then
                begin
                        If AvecResp = True then TFQRS1(Ecran).WhereSQL := 'SELECT PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_DATEDEBUT,PFO_DATEFIN'+Select+' FROM FORMATIONS LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE '+Where+OrderBy
                        else TFQRS1(Ecran).WhereSQL := 'SELECT PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_DATEDEBUT,PFO_DATEFIN'+Select+' FROM FORMATIONS '+Where+OrderBy;
                end
                else
                begin
                        If AvecResp = True then TFQRS1(Ecran).WhereSQL := 'SELECT PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_DATEDEBUT,PFO_DATEFIN,PSS_LIBELLE'+Select+' FROM FORMATIONS '+
                                'LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE '+
                                'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME '+Where+OrderBy
                        else TFQRS1(Ecran).WhereSQL := 'SELECT PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_DATEDEBUT,PFO_DATEFIN,PSS_LIBELLE'+Select+' FROM FORMATIONS '+
                                'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME '+Where+OrderBy;
                end;
        end;
        If LeMenu = '47563' then   //Couts détaillé
        begin
                TFQRS1(Ecran).CodeEtat := 'PSC';
                TFQRS1(Ecran).FCodeEtat := 'PSC';

                //PT22
				// N'affiche pas les salariés confidentiels
                If PGBundleHierarchie Then
    				TFQRS1(Ecran).WhereSQL  := TFQRS1(Ecran).WhereSQL  + ' AND PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
                Else
    				TFQRS1(Ecran).WhereSQL  := TFQRS1(Ecran).WhereSQL  + ' AND PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
        end;

end ;

procedure TOF_PGEDITPLANFORMATION.OnArgument (S : String ) ;
var Num,i : Integer;
    Combo : THValComboBox;
    Edit : THEdit;
    DF,DD : TdateTime;
begin
  Inherited ;
	{$IFDEF EMANAGER}
	//  If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
	//          'WHERE PHO_NIVEAUH<=2 AND PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"') then MultiNiveau := True
	If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
	'AND PGS_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE)') then MultiNiveau := True
	else MultiNiveau := False;
	If MultiNiveau = False then
	begin
		SetControlVisible('PSE_RESPONSFOR',False);
		SetControlVisible('TPSE_RESPONSFOR',False);
	end;
	SetControlText('PSE_RESPONSFOR',V_PGI.UserSalarie);
	SetControltext('LIBRESP',RechDom('PGSALARIE',V_PGI.UserSalarie,False));
	{$ENDIF}
	If PGBundleHierarchie Then //PT20
    Begin
		Edit := THEdit(getControl('PSE_RESPONSFOR'));
		If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick;
	End;

	LeMenu := ReadTokenPipe(S,';');
	SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
	RendMillesimeRealise(DD,DF);
	SetControlText('PFO_DATEDEBUT',DateToStr(DD));
	SetControlText('PFO_DATEFIN',DateToStr(DF));

	For Num  :=  1 to VH_Paie.NBFormationLibre do
	begin
		if Num >8 then Break;
		VisibiliteChampFormation (IntToStr(Num),GetControl ('PFO_FORMATION'+IntToStr(Num)),GetControl ('TPFO_FORMATION'+IntToStr(Num)));
	end;
	For Num  :=  1 to VH_Paie.PGNbreStatOrg do
	begin
		if Num >4 then Break;
		VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
		SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
	end;
	VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;

	If LeMenu = '47525' then
	begin
		SetControlVisible('SANSDETAIL',False);
		SetControlProperty('PRUPT','TabVisible',False);
		Ecran.Caption := 'Suivi du Droit Individuel à la Formation';
		SetControlText('PFO_TYPEPLANPREV','DIF');
	end;
	If LeMenu = '47531' then     //Suivi pédagogique individuel
	begin
		SetControlVisible('SANSDETAIL',False);
		SetControlProperty('PRUPT','TabVisible',False);
		{$IFDEF EMANAGER}
		Ecran.Caption := 'Edition du réalisé';
		{$ELSE}
		Ecran.Caption := 'Suivi pédagogique individuel';
		{$ENDIF}

	end;
	If LeMenu = '47532' then     //Suivi pédagogique collectif
	begin

		Ecran.Caption := 'Suivi pédagogique collectif';
	end;
	If LeMenu = '47551' then     //Suivi gestion individuel
	begin
		SetControlVisible('SANSDETAIL',False);
		SetControlProperty('PRUPT','TabVisible',False);
		Ecran.Caption := 'Suivi de gestion individuel';
		SetControlVisible('CCODEORGA',True);
	end;
	If LeMenu = '47552' then     //Suivi gestion collectif
	begin
		Ecran.Caption := 'Suivi de gestion collectif';
	end;
	If LeMenu = '47551' then     //Suivi de gestion individuel
	begin
		Ecran.Caption := 'Suivi de gestion individuel';
	end;
	If LeMenu = '47561' then  // Frais des stagiaires
	begin
		SetControlVisible('SANSDETAIL',False);
		Ecran.Caption := 'Edition des frais par stagiaire';
	end;
	If LeMenu = '47563' then // Coûts détaillés
	begin
		SetControlProperty('PRUPT','TabVisible',False);
		Ecran.Caption := 'Détail des coûts par session';
	end;
	UpdateCaption(Ecran);
	For i := 1 To 4 do
	begin
		Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
		If Combo <> Nil Then
		begin
			// If i <> 1 Then
			// begin
			Combo.Items.Add('<<Aucune>>');
			Combo.Values.Add('');
			//  end;
			Combo.Items.Add('Salarié');
			Combo.Values.Add('PFO_SALARIE');
			Combo.Items.Add('Responsable formation');
			Combo.Values.Add('PSE_RESPONSFOR');
			Combo.Items.Add('Formation');
			Combo.Values.Add('PFO_CODESTAGE');
			Combo.Items.Add('Etablissement');
			Combo.Values.Add('PFO_ETABLISSEMENT');
			Combo.Items.Add('Libellé emploi');
			Combo.Values.Add('PFO_LIBEMPLOIFOR');
			Combo.Items.Add('Nature (Interne/Externe)');
			Combo.Values.Add('PFO_NATUREFORM');
			Combo.Items.Add('Lieu');
			Combo.Values.Add('PFO_LIEUFORM');
			Combo.Items.Add('Centre de formation');
			Combo.Values.Add('PFO_CENTREFORMGU');
			For Num  :=  1 to VH_Paie.NBFormationLibre do
			begin
				If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PFO_FORMATION1');end;
				If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PFO_FORMATION2');end;
				If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PFO_FORMATION3');end;
				If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PFO_FORMATION4');end;
				If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PFO_FORMATION5');end;
				If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PFO_FORMATION6');end;
				If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PFO_FORMATION7');end;
				If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PFO_FORMATION8');end;
			end;
			For Num  :=  1 to VH_Paie.PGNbreStatOrg do
			begin
				If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('PFO_TRAVAILN1');end;
				If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('PFO_TRAVAILN2');end;
				If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('PFO_TRAVAILN3');end;
				If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('PFO_TRAVAILN4');end;
			end;
			If VH_Paie.PGLibCodeStat <> '' Then begin Combo.Items.Add(VH_Paie.PGLibCodeStat); Combo.Values.Add('PFO_CODESTAT');end;
		end;
	end;
	Edit := THEdit(GetControl('PFO_SALARIE'));
	if Edit <> Nil then
	begin
		Edit.OnElipsisClick := SalarieElipsisClick;
		Edit.OnExit := ExitEdit;
	end;
	Edit := THEdit(GetControl('PFO_SALARIE_'));
	if Edit <> Nil then
	begin
		Edit.OnElipsisClick := SalarieElipsisClick;
		Edit.OnExit := ExitEdit;
	end;
	
	//PT18 - Début
	If PGBundleCatalogue Then //PT20
	Begin
		If not PGDroitMultiForm then
			SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
       		SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT18
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFO_PREDEFINI',False);
			SetControlText   ('PFO_PREDEFINI','');
			//SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT18 //PT20
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT19
       		//SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT18 
       	End;
	end
	else
	begin
		SetControlVisible('PFO_PREDEFINI', False);
		SetControlVisible('TPFO_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT18 - Fin
end ;

procedure TOF_PGEDITPLANFORMATION.SalarieElipsisClick(Sender : TObject);
var StFrom,StWhere : String;
begin
	StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
	StWhere := '';
	{$IFDEF EMANAGER}
		If GetControltext('PSE_RESPONSFOR') <> '' then
			StWhere := 'PSE_RESPONSFOR="'+GetcontrolText('PSE_RESPONSFOR')+'"'
		else
			StWhere := 'PSE_CODESERVICE IN '+
						'(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
						' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
						' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))';
		StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
		LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,'', True,-1);
	{$ELSE}
		StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
		//PT18
		If PGBundleHierarchie Then
			ElipsisSalarieMultidos (Sender)
		Else
			LookupList(THEdit(Sender), 'Liste des salariés',StFrom,'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
	{$ENDIF}
end;

{ TOF_PGETATLIBREFOR }

procedure TOF_PGETATLIBREFOR.OnArgument (S : String ) ;
var DD,DF : TDatetime;
begin
     Inherited ;
     RendMillesimeRealise(DD,DF);
     SetControlText('PSS_DATEDEBUT',DateToStr(DD));
     SetControlText('PSS_DATEDEBUT_',DateToStr(DF));
     
	//PT18 - Début
	If PGBundleCatalogue Then //PT20
	Begin
		If not PGDroitMultiForm then
			SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
       		SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT18
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PSS_PREDEFINI',False);
			SetControlText   ('PSS_PREDEFINI','');
			//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT18 //PT20
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT19
       		//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT18 //PT20
       	End;
	end
	else
	begin
		SetControlVisible('PSS_PREDEFINI', False);
		SetControlVisible('TPSS_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT18 - Fin     
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT18
Modifié le ... :   /  /
Description .. : Lancement de l'état
Mots clefs ... :
*****************************************************************}
procedure TOF_PGETATLIBREFOR.OnUpdate ;
Begin
	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS'));
End;


{ TOF_PGEDITSUIVICURSUS }
//PT4 - Début

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 19/04/2007
Modifié le ... :   /  /    
Description .. : Comportement de la combo salarié sur la sortie de la sélection
Mots clefs ... : 
*****************************************************************}
{$IFNDEF EMANAGER}
procedure TOF_PGEDITSUIVICURSUS.ExitEdit(Sender: TObject);
var Edit : THEdit;
begin
     Edit := THEdit(Sender);
     If Edit <> Nil Then	//AffectDefautCode que si gestion du code salarié en Numérique
     If (VH_Paie.PgTypeNumSal='NUM') And (length(Edit.Text)<11) And (IsNumeric(Edit.Text)) Then
     Edit.Text := AffectDefautCode(Edit,10);
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 19/04/2007
Modifié le ... :   /  /    
Description .. : Initialisation des valeurs par défaut de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITSUIVICURSUS.OnArgument(S: String);
var {$IFNDEF EMANAGER}
    Num   : Integer;
    {$ENDIF}
    Edit  : THEdit;
    DF,DD : TdateTime;
begin
  Inherited ;

  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));

  // Dates de début et de fin par défaut : Millésime
  RendMillesimeRealise(DD,DF);
  SetControlText('PFO_DATEDEBUT',DateToStr(DD));
  SetControlText('PFO_DATEFIN',DateToStr(DF));

  {$IFNDEF EMANAGER}//PT14
  // Affichage des champs libres de l'onglet Formation
  For Num := 1 to VH_Paie.NBFormationLibre do
  begin
          if Num >8 then Break;
          VisibiliteChampFormation (IntToStr(Num),GetControl ('PFO_FORMATION'+IntToStr(Num)),GetControl ('TPFO_FORMATION'+IntToStr(Num)));
  end;

  // Affichage des champs libres de l'onglet Salarié
  For Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
          if Num >4 then Break;
          VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
          SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
  end;
  VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;

  // Comportement des ellipsis pour la sélection des salariés
  Edit := THEdit(GetControl('PFO_SALARIE')); // De
  if Edit <> Nil then Edit.OnExit := ExitEdit;
  Edit := THEdit(GetControl('PFO_SALARIE_')); // à
  if Edit <> Nil then Edit.OnExit := ExitEdit;
  
	//PT18 - Début
	If PGBundleCatalogue Then //PT20
	Begin
		If not PGDroitMultiForm then
			SetControlProperty ('PFO_CURSUS', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PCU',True,True))
		Else If V_PGI.ModePCL='1' Then
			SetControlProperty ('PFO_CURSUS', 'Plus', DossiersAInterroger('','','PCU',True,True)); //PT18
	End;

	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFO_PREDEFINI',False);
			SetControlText   ('PFO_PREDEFINI','');
			//SetControlProperty ('PFO_CURSUS', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PCU',True,True)); //PT18 //PT20
		end
		Else If V_PGI.ModePCL='1' Then
		Begin
			SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT19
			//SetControlProperty ('PFO_CURSUS', 'Plus', DossiersAInterroger('','','PCU',True,True)); //PT18 //PT20
		End;
	end
	else
	begin
		SetControlVisible('PFO_PREDEFINI', False);
		SetControlVisible('TPFO_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;

	If PGBundleHierarchie Then //PT20
    Begin	
		Edit := THEdit(GetControl('PFO_SALARIE'));
		if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
		Edit := THEdit(GetControl('PFO_SALARIE_'));
		if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
		Edit := THEdit(GetControl('PSE_RESPONSFOR'));
		if Edit <> nil then Edit.OnElipsisClick := RespElipsisClick;
	End
    Else
    	SetControlProperty('PSE_RESPONSFOR', 'Plus', 'SALARIE IN (SELECT PGS_RESPONSFOR FROM SERVICES)'); //PT20
	//PT18 - Fin
  
  {$ELSE}
  //PT14 - Début
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';

  SetControlProperty('PFO_LIBEMPLOIFOR','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
  'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND '+
  AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,False)+')');
  
  Edit := THEdit(GetControl('PFO_SALARIE'));
  if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
  //PT14 - Fin
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 26/10/2007 / PT14
Modifié le ... :   /  /    
Description .. : Elipsis pour salariés rattachés au manager
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITSUIVICURSUS.SalarieElipsisClick(Sender: TObject);
{$IFDEF EMANAGER}
var
  StFrom, StWhere: string;
{$ENDIF}
begin
	{$IFNDEF EMANAGER}
    //If PGBundleInscFormation Then //PT18 //PT20
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
	{$ELSE}
	StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
	StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
	StWhere := StWhere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
	LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
  	{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT18
Modifié le ... :   /  /
Description .. : Clic elipsis Responsable
Mots clefs ... :
*****************************************************************}
{$IFNDEF EMANAGER}
procedure TOF_PGEDITSUIVICURSUS.RespElipsisClick(Sender: TObject);
begin
	ElipsisResponsableMultidos (Sender);
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 20/04/2007
Modifié le ... :   /  /
Description .. : Fermeture de l'état
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITSUIVICURSUS.OnClose;
begin
  Inherited ;
  FreeAndNil (TobEdition);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 19/04/2007
Modifié le ... :   /  /    
Description .. : Construction de la requête de base pour la génération de 
Suite ........ : l'état
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITSUIVICURSUS.OnUpdate;
var Pages : TPageControl;
    Where : String;
begin
  Inherited ;
     Pages := TPageControl(GetControl('Pages'));
     
	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFO')); //PT18     

     // Récupération des critères de l'écran afin de construire la clause Where automatiquement
     Where := RecupWhereCritere(Pages);

     //PT22
	// N'affiche pas les salariés confidentiels
	If Where <> '' Then Where := Where + ' AND ';
	If PGBundleHierarchie Then
		Where := Where + ' PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
	Else
		Where := Where + ' PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';

     // Création de la TOB servant à imprimer les données
     If Assigned(TobEdition) Then FreeAndNil(TobEdition);
     TobEdition := Tob.Create('SuiviCursus',Nil,-1);

     // Démarrage de l'écran d'attente
     InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', 0, FALSE, TRUE);
     InitMove(0, '');

     // Récupération des données
     RecupData (Where);

     // Tri selon le nom du salarié, le libellé du cursus et le libellé du stage (les autres
     // colonnes ne servent que pour les niveaux de rupture)
     TobEdition.Detail.Sort('PFO_NOMSALARIE;PFO_SALARIE;PCU_LIBELLE;PFO_CURSUS;PST_LIBELLE;PFO_CODESTAGE');

     // Fermeture de l'écran d'attente
     FiniMoveProgressForm;

     // Attribution de la TOB à l'état
     TFQRS1(Ecran).LaTob:= TobEdition;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 20/04/2007
Modifié le ... :   /  /    
Description .. : Copie des informations de la TOB des salariés en formation 
Suite ........ : dans la TOB d'édition
Mots clefs ... : 
*****************************************************************}
Procedure TOF_PGEDITSUIVICURSUS.CopieData (Donnees : TOB; LibelleCursus,Stage,LibelleStage : String; DateDebut,DateFin : TDateTime; Realise : String);
Var
  T : TOB;
Begin
      T := Tob.Create('FilleEdition',TobEdition,-1);

      T.AddChampSupValeur('PFO_SALARIE', Donnees.GetValue('PFO_SALARIE'));
      T.AddChampSupValeur('PFO_NOMSALARIE', Donnees.GetValue('PFO_NOMSALARIE'));
      T.AddChampSupValeur('PFO_PRENOM', Donnees.GetValue('PFO_PRENOM'));
      T.AddChampSupValeur('PFO_CURSUS', Donnees.GetValue('PFO_CURSUS'));
      T.AddChampSupValeur('PCU_LIBELLE', LibelleCursus);
      T.AddChampSupValeur('PFO_CODESTAGE', Stage);
      T.AddChampSupValeur('PST_LIBELLE', LibelleStage);
      If (DateDebut <> iDate1900) Then
          T.AddChampSupValeur('PFO_DATEDEBUT', DateDebut)
      Else
          T.AddChampSupValeur('PFO_DATEDEBUT', '');
      If (DateFin <> iDate1900) Then
          T.AddChampSupValeur('PFO_DATEFIN', DateFin)
      Else
          T.AddChampSupValeur('PFO_DATEFIN', '');
      T.AddChampSupValeur('REALISE', Realise);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 19/04/2007
Modifié le ... : 12/06/2007 / PT11
Description .. : Création d'une TOB pour la liste des cursus et des 
Suite ........ : formations dépendantes. Création d'une TOB pour les
Suite ........ : salariés et formations effectuées.
Suite ........ : Au final, alimentation de la TOB générale avec les infos.
Paramètres ... : StWhere : Clause Where reprenant les critères sélectionnés par l'utilisateur
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITSUIVICURSUS.RecupData(StWhere: String);
var TCursus, TFormations, T, T2 : Tob;
    Q : TQuery;
    i : Integer;
    Salarie,Cursus,CodeStage,OrderBy,WhereResp,FromResp,Where : String;
    DateDebut,DateFin : TDateTime;
    TopRealise : String;
begin
     { Récupération des cursus }

	//PT18 - Début
	If PGBundleInscFormation Then
		Where := DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PCU',False,True)
	Else
		Where := '';
	//PT18 - Fin

     // Construction de la requête de récupération des cursus
     Q := OpenSQL('SELECT DISTINCT(PCU_CURSUS), PCU_LIBELLE, PCC_CODESTAGE, PST_LIBELLE '+
                  ' FROM CURSUS, CURSUSSTAGE, STAGE '+ //StWhere +
                  ' WHERE PCU_CURSUS = PCC_CURSUS AND PCC_CODESTAGE=PST_CODESTAGE' +Where+
                  ' ORDER BY PCU_LIBELLE,PST_LIBELLE', True);

     // Création de la TOB récupérant les résultats
     TCursus := Tob.Create('ListeCursus',Nil,-1);

     // Lecture des enregistrements en base
     TCursus.LoadDetailDB('ListeCursus','','',Q,False);
     Ferme(Q);

     { Récupération des formations effectuées }

     // Création de l'ordre de tri
     OrderBy := 'PFO_CURSUS, PFO_CODESTAGE';
     If (THCheckbox(GetControl('CALPHA')).Checked = True) Then
          OrderBy := 'PFO_NOMSALARIE, PFO_PRENOM, ' + OrderBy;

     // Cas particulier du responsable de formation
     WhereResp := '';
     FromResp := '';
     {$IFNDEF EMANAGER}//PT14
     If (THEdit(GetControl('PSE_RESPONSFOR')).Text <> '') Then
     Begin
       WhereResp := ' AND PFO_SALARIE = PSE_SALARIE';
       FromResp := ', DEPORTSAL ';
     End;
     {$ELSE}
     WhereResp := ' AND ' +AdaptByRespEmanager(TypeUtilisat, 'PFO', V_PGI.UserSalarie, False); //PT18
     {$ENDIF}

     // Construction de la requête de récupération des salariés avec les formations effectuées au sein de cursus
     Q := OpenSQL('SELECT PFO_SALARIE, PFO_NOMSALARIE, PFO_PRENOM, PFO_CURSUS, PFO_CODESTAGE, PFO_DATEDEBUT, PFO_DATEFIN, PFO_EFFECTUE '+ //PT21
                  ' FROM FORMATIONS '+ FromResp + StWhere +
                  ' AND ' + WhereResp + //AND PFO_CURSUS <> ""  / PT11 PFO_EFFECTUE = "X" //PT21
                  ' ORDER BY ' + OrderBy, True);

     // Création de la TOB récupérant les résultats
     TFormations := Tob.Create('ListeFormations',Nil,-1);

     // Lecture des enregistrements en base
     TFormations.LoadDetailDB('ListeFormations','','',Q,False);
     Ferme(Q);

     { Traitement des données}

     // 1e partie : Parcours des résultats rattachées à un cursus afin d'alimenter la TOB mère
     For i := 0 To TFormations.Detail.Count-1 Do
     Begin
          Cursus    := TFormations.Detail[i].GetValue('PFO_CURSUS');

          // La première passe ne considère que les cursus renseignés. On élimine donc du traitement
          // les formations réalisées non rattachées à un cursus.
          If Cursus = '' Then Continue;

          Salarie   := TFormations.Detail[i].GetValue('PFO_SALARIE');
          CodeStage := TFormations.Detail[i].GetValue('PFO_CODESTAGE');

          // Recherche du groupe [Salarié,Cursus,Stage] dans la TOB mère
          T := TobEdition.FindFirst(['PFO_SALARIE','PFO_CURSUS','PFO_CODESTAGE'],
                                    [Salarie,Cursus,CodeStage], True);

          { On n'a pas trouvé d'occurrence, on crée donc la ligne en question ainsi que
            tous les autres stages du cursus pour ce salarié (dates à vide et top réalisé à Non par défaut) }
          If (T = Nil) Then
          Begin
               With TFormations.Detail[i] Do
               Begin
                    // Création de tous les stages du cursus pour le salarié
                    T := TCursus.FindFirst(['PCU_CURSUS'],[Cursus],True);
                    While T <> Nil Do
                    Begin
                         // Cas de l'enregistrement courant
                         If (CodeStage = T.GetValue('PCC_CODESTAGE')) Then
                         Begin
                              DateDebut := GetValue('PFO_DATEDEBUT');
                              DateFin   := GetValue('PFO_DATEFIN');
                              TopRealise:= GetValue('PFO_EFFECTUE'); //PT21
                         End
                         // Cas des autres enregistrements
                         Else
                         Begin
                              DateDebut := iDate1900;
                              DateFin   := iDate1900;
                              TopRealise:= '-';
                         End;

                         // Création de la ligne dans la TOB d'édition
                         CopieData (TFormations.Detail[i],T.GetValue('PCU_LIBELLE'),T.GetValue('PCC_CODESTAGE'),T.GetValue('PST_LIBELLE'),DateDebut,DateFin,TopRealise);

                         T := TCursus.FindNext(['PCU_CURSUS'],[Cursus],False);
                    End;
               End;
          End
          Else
          { Le stage était déjà dans la TOB. On la passe à l'état Réalisé et on met à jour les dates }
          Begin
               T.PutValue('REALISE','X');
               T.PutValue('PFO_DATEDEBUT',TFormations.Detail[i].GetValue('PFO_DATEDEBUT'));
               T.PutValue('PFO_DATEFIN',TFormations.Detail[i].GetValue('PFO_DATEFIN'));
          End;
     End;

     // 2e Partie : Parcours des résultats non rattachées à un cursus
     For i := 0 To TFormations.Detail.Count-1 Do
     Begin
          Cursus    := TFormations.Detail[i].GetValue('PFO_CURSUS');

          // A présent, on ne garde que les formations non rattachées afin d'essayer de trouver une correspondance
          // parmi les informations déjà traitées
          If Cursus <> '' Then Continue;

          Salarie   := TFormations.Detail[i].GetValue('PFO_SALARIE');
          CodeStage := TFormations.Detail[i].GetValue('PFO_CODESTAGE');

          T := TobEdition.FindFirst(['PFO_SALARIE','PFO_CODESTAGE','REALISE'],
                                    [Salarie,CodeStage,'-'], True);
                                    
          { Si on ne trouve pas d'occurrence, c'est que le stage n'est pas compris dans un cursus.
            Si on trouve plusieurs occurrences, on ne peut pas déterminer le cursus de rattachement.
            Dans les 2 cas, on ne fait rien. }
          If (T <> Nil) Then
          Begin
               T2 := TobEdition.FindNext(['PFO_SALARIE','PFO_CODESTAGE','REALISE'],[Salarie,CodeStage,'-'], True);

               // Il n'y a qu'un seul enregistrement trouvé, on peut le rattacher
               If T2 = Nil Then
               Begin
                    T.PutValue('REALISE','X');
                    T.PutValue('PFO_DATEDEBUT',TFormations.Detail[i].GetValue('PFO_DATEDEBUT'));
                    T.PutValue('PFO_DATEFIN',TFormations.Detail[i].GetValue('PFO_DATEFIN'));
               End;
          End;
     End;

     // Libération des objets temporaires
     TCursus.Free;
     TFormations.Free;
end;
//PT4 - Fin

{ TOF_PGEDITSUIVIFORM }

//PT10 - Début
procedure TOF_PGEDITSUIVIFORM.CkSortieClick(Sender: TObject);
begin
     If Sender = Nil then Exit;
     If GetCheckBoxState('CKSORTIE') = CbChecked then
     begin
             SetControlVisible('DATESORTIE',True);
             SetControlText('DATESORTIE',DateToStr(Date));
             SetControlVisible('TDATEARRET',True);
     end
     else
     begin
             SetControlVisible('DATESORTIE',False);
             SetControlVisible('TDATEARRET',False);
     end;
end;

procedure TOF_PGEDITSUIVIFORM.DateElipsisClick(Sender: TObject);
var key  :  char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITSUIVIFORM.ExitEdit(Sender: TObject);
var edit : thedit;
begin
     edit:=THEdit(Sender);
     if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
          if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
               edit.text:=AffectDefautCode(edit,10);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT18
Modifié le ... :   /  /
Description .. : Clic elipsis salarie
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITSUIVIFORM.SalarieElipsisClick(Sender : TObject);
begin
    //If PGBundleInscFormation Then //PT20
    	ElipsisSalarieMultidos (Sender)
    //Else
		//Inherited;
end;

procedure TOF_PGEDITSUIVIFORM.OnArgument(S: String);
Var
    Check : TCheckBox;
    Min,Max : String;
    Edit : THEdit;
    DD,DF : TDateTime;
begin
  inherited;
     SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));

     // Combo de sortie
     Check := TCheckBox(GetControl('CKSORTIE'));
     If Check <> Nil Then Check.OnCLick := CkSortieClick;

     // Etablissements
     RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
     SetControlText('PSA_ETABLISSEMENT',Min);
     SetControlText('PSA_ETABLISSEMENT_',Max);

     // Salariés
     //PT18 - Début
    If Not PGBundleHierarchie Then
    Begin
		RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
     	SetControlText('PSA_SALARIE', Min);
     	SetControlText('PSA_SALARIE_',Max);
    End;
    //PT18 - Fin
     Edit := ThEdit(getcontrol('PSA_SALARIE'));
     If Edit <> nil then 
     Begin
     	Edit.OnExit:=ExitEdit;
     	If PGBundleHierarchie Then Edit.OnElipsisClick := SalarieElipsisClick; //PT18 //PT20
    End;
     
     
     Edit := ThEdit(getcontrol('PSA_SALARIE_'));
     If Edit <> nil then 
     Begin
     	Edit.OnExit:=ExitEdit;
     	If PGBundleHierarchie Then Edit.OnElipsisClick := SalarieElipsisClick; //PT18 //PT20
    End;

     // Dates d'entrée et de sortie
     SetControlText('PSA_DATEENTREE', DateToStr(IDate1900));
     SetControlText('PSA_DATEENTREE_',DateToStr(Date));
     Edit := THEdit(GetControl('PSA_DATEENTREE'));
     If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisClick;
     Edit := THEdit(GetControl('PSA_DATEENTREE_'));
     If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
     Edit := THEdit(GetControl('DATESORTIE'));
     If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;

     // Dates de formation
     RendMillesimeRealise(DD,DF);
     SetControlText('DATEDEBUT',DateToStr(DD));
     SetControlText('DATEFIN',  DateToStr(DF));

     TobEdition := Nil;
     
	//PT18 - Début
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT19
	end
	else
	begin
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT18 - Fin     
end;

procedure TOF_PGEDITSUIVIFORM.OnClose;
begin
  inherited;
     FreeAndNil(TobEdition);
end;

procedure TOF_PGEDITSUIVIFORM.OnUpdate;
var StWhere,StJoin,StSQL,Temp,Valeur : String;
    i : Integer;
    Pages : TPageControl;
    First : Boolean;
begin
	Inherited ;
	Pages := TPageControl(GetControl('Pages'));

	//PT18 - Début
	If PGBundleInscFormation Then
	Begin
		SetControlText('XX_WHEREPREDEF', DossiersAInterroger('',GetControlText('NODOSSIER'),'PSI',False,False)); //PT18
	End;

	//PT18
	If PGBundleHierarchie Then
	Begin
		StSQL := 'SELECT DISTINCT PSI_INTERIMAIRE AS SALARIE, PSI_LIBELLE AS LIBELLE, PSI_PRENOM AS PRENOM, '+
		'PSI_ETABLISSEMENT AS ETABLISSEMENT, PSI_LIBELLEEMPLOI AS LIBELLEEMPLOI, SUM(PFO_NBREHEURE) AS NBH FROM INTERIMAIRES LEFT JOIN FORMATIONS ON PSI_INTERIMAIRE=PFO_SALARIE ';
	End
	Else
	Begin
		StSQL := 'SELECT DISTINCT PSA_SALARIE AS SALARIE, PSA_LIBELLE AS LIBELLE, PSA_PRENOM AS PRENOM, '+
		'PSA_ETABLISSEMENT AS ETABLISSEMENT, PSA_LIBELLEEMPLOI AS LIBELLEEMPLOI, SUM(PFO_NBREHEURE) AS NBH FROM SALARIES LEFT JOIN FORMATIONS ON PSA_SALARIE=PFO_SALARIE ';
	End;

	// Clause Where
	StWhere := RecupWhereCritere(Pages);
	If GetCheckBoxState('CKSORTIE') = CbChecked then
	begin
		If StWhere <> ''  then StWhere := StWhere + ' AND '
		Else StWhere := 'WHERE ';
		If PGBundleHierarchie Then //PT13
		Begin
			StWhere := StWhere + '(PSI_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSI_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")';
			//StSQL := 'SELECT DISTINCT PSI_INTERIMAIRE AS SALARIE, PSI_LIBELLE AS LIBELLE, PSI_PRENOM AS PRENOM, '+
			//         'PSI_ETABLISSEMENT AS ETABLISSEMENT, PSI_LIBELLEEMPLOI AS LIBELLEEMPLOI, SUM(PFO_NBREHEURE) AS NBH FROM INTERIMAIRES LEFT JOIN FORMATIONS ON PSI_INTERIMAIRE=PFO_SALARIE ';
		End
		Else
		Begin
			StWhere := StWhere + '(PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")';
			//StSQL := 'SELECT DISTINCT PSA_SALARIE AS SALARIE, PSA_LIBELLE AS LIBELLE, PSA_PRENOM AS PRENOM, '+
			//         'PSA_ETABLISSEMENT AS ETABLISSEMENT, PSA_LIBELLEEMPLOI AS LIBELLEEMPLOI, SUM(PFO_NBREHEURE) AS NBH FROM SALARIES LEFT JOIN FORMATIONS ON PSA_SALARIE=PFO_SALARIE ';
		End;
	end;

	// Compléments pour la formation
	StJoin := '';
	If GetControlText('DATEDEBUT') <> '  /  /    ' Then StJoin := StJoin + ' AND PFO_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" ';
	If GetControlText('DATEFIN')   <> '  /  /    ' Then StJoin := StJoin + ' AND PFO_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'" ';
	If GetControlText('TYPEPLANPREV') <> '' Then
	Begin
		StJoin := StJoin + ' AND PFO_TYPEPLANPREV IN (';
		Temp := GetControlText('TYPEPLANPREV');
		Valeur := ReadTokenPipe(Temp,';') ;
		First := True;
		While (Valeur <> '') Do
		Begin
			If Not First Then StJoin := StJoin + ',';
			StJoin := StJoin + '"' + Valeur + '"';
			Valeur := ReadTokenPipe(Temp,';');
			First := False;
		End;
		StJoin := StJoin + ') ';
	End;

    //PT22
	// N'affiche pas les salariés confidentiels
	If StWhere <> '' Then StWhere := StWhere + ' AND ';
	If PGBundleHierarchie Then
		StWhere := StWhere + ' PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
	Else
		StWhere := StWhere + ' PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';


	if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
	if StWhere <> '' then
	StWhere := StWhere + ' AND ' + MonHabilitation.LeSQL
	else
	StWhere := MonHabilitation.LeSQL;

	// Création de la TOB servant à imprimer les données
	If Assigned(TobEdition) Then TobEdition.Free;
	TobEdition := Tob.Create('SuiviFormations',Nil,-1);

	//PT18
	If PGBundleHierarchie Then
	Begin
		StWhere := StringReplace(StWhere, 'PSA_SALARIE', 'PSI_INTERIMAIRE', [rfReplaceAll, rfIgnoreCase]);
		StWhere := StringReplace(StWhere, 'PSA_', 'PSI_', [rfReplaceAll, rfIgnoreCase]);
	End;

	// Récupération des données
	//PT18
	If PGBundleHierarchie Then //PT13
	TobEdition.LoadDetailFromSQL(StSQL + StJoin + StWHere + ' GROUP BY PSI_ETABLISSEMENT, PSI_INTERIMAIRE, PSI_LIBELLE, PSI_PRENOM, PSI_LIBELLEEMPLOI ORDER BY PSI_ETABLISSEMENT, PSI_LIBELLE')
	Else
	TobEdition.LoadDetailFromSQL(StSQL + StJoin + StWHere + ' GROUP BY PSA_ETABLISSEMENT, PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM, PSA_LIBELLEEMPLOI ORDER BY PSA_ETABLISSEMENT, PSA_LIBELLE');

	//     TobEdition.LoadDetailFromSQL(StSQL + StJoin + StWHere + ' GROUP BY ETABLISSEMENT, SALARIE, LIBELLE, PRENOM, LIBELLEEMPLOI ORDER BY ETABLISSEMENT, LIBELLE');

	// Mise à 0 du nombre d'heures si aucune formation donnée
	For i := 0 To TobEdition.Detail.Count - 1 Do
	Begin
		If TobEdition.Detail[i].GetString('NBH') = '' Then
		TobEdition.Detail[i].PutValue('NBH', 0);
	End;

	// Attribution de la TOB à l'état
	TFQRS1(Ecran).LaTob:= TobEdition;
end;
//PT10 - Fin

//PT25
{TOF_PGEDTSUIVIINTEGR}

procedure TOF_PGEDTSUIVIINTEGR.OnArgument (S : String ) ;
var Num,i : Integer;
    Combo : THValComboBox;
    CkBox : TCheckBox;
    Edit : THEdit;
    DD,DF : TDateTime;
begin
    Inherited ;
    TobEtat := nil;
    
    Edit := THEdit(GetControl('PSA_SALARIE'));
    If Edit <> Nil then 
    Begin
		If PGBundleHierarchie Then Edit.OnElipsisClick := SalarieElipsisClick; 
	End;
	
    Edit := THEdit(GetControl('PSE_RESPONSFOR'));
    If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick; 
	  
    // Par défaut : sur 6 mois, limité à la fin de l'exercice en cours
    SetControlText('PSA_DATEENTREE',   DateToStr(PlusMois(Date,-6)));
    SetControlText('PSA_DATEENTREE_',  DateToStr(Date));
	RendMillesimeRealise(DD,DF);
    SetControlText('DATEFINSUIVI', DateToStr(DF));
    
    For Num  :=  1 to VH_Paie.NBFormationLibre do
    begin
            if Num >8 then Break;
            VisibiliteChampFormation (IntToStr(Num),GetControl ('FORMATION'+IntToStr(Num)),GetControl ('TFORMATION'+IntToStr(Num)));
    end;
    
    For Num  :=  1 to VH_Paie.PGNbreStatOrg do
    begin
            if Num >4 then Break;
            VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
            SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
    end;
    VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;

    For i := 1 To 4 do
    begin
            Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
            If Combo <> Nil Then
            begin
                   // Rendre invisible toutes les check-boxes de saut de page par défaut
                   CkBox := TCheckBox(GetControl('CSAUTRUPT'+IntToStr(i)));
                   CkBox.Visible := False;
                   // Sur le changement de valeur d'une combo, il faut afficher le saut de page précédent
                   Combo.OnChange := OnChangeRuptures;

                    If i <> 1 Then
                    begin
                            Combo.Items.Add('<<Aucune>>');
                            Combo.Values.Add('');
                    end;
                    // Pas d'établissement en environnement partagé car les établissements ne sont pas mis en commun
                    If Not PGBundleHierarchie Then
                    Begin
                    	Combo.Items.Add('Etablissement');
                    	Combo.Values.Add('PFO_ETABLISSEMENT');
                    End;
                    Combo.Items.Add('Libellé emploi');
                    Combo.Values.Add('PFO_LIBEMPLOIFOR');
                    For Num  :=  1 to VH_Paie.NBFormationLibre do
                    begin
                            If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PFO_FORMATION1');end;
                            If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PFO_FORMATION2');end;
                            If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PFO_FORMATION3');end;
                            If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PFO_FORMATION4');end;
                            If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PFO_FORMATION5');end;
                            If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PFO_FORMATION6');end;
                            If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PFO_FORMATION7');end;
                            If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PFO_FORMATION8');end;
                    end;
                    For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                    begin
                            If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('PFO_TRAVAILN1');end;
                            If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('PFO_TRAVAILN2');end;
                            If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('PFO_TRAVAILN3');end;
                            If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('PFO_TRAVAILN4');end;
                    end;
                    If VH_Paie.PGLibCodeStat <> '' Then begin Combo.Items.Add(VH_Paie.PGLibCodeStat); Combo.Values.Add('PFO_CODESTAT');end;
            end;
    end;
    
    {$IFDEF EMANAGER}
    SetControlText   ('PSE_RESPONSFOR',  V_PGI.UserSalarie);
    SetControlVisible('PSE_RESPONSFOR',  False);
    SetControlVisible('TPSE_RESPONSFOR', False);
    SetControlVisible('DATEFINSUIVI',    False);
    SetControlVisible('TDATEFINSUIVI',   False);
    SetControlVisible('PSA_DATEENTREE',   False);
    SetControlVisible('PSA_DATEENTREE_',   False);
    SetControlVisible('TPSA_DATEENTREE',   False);
    SetControlVisible('TPSA_DATEENTREE_',   False);
    {$ELSE}
	
    // N'affiche pas les salariés confidentiels
    If PGBundleHierarchie Then
    	SetControlText('XX_WHERE', 'PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"')
    Else
    	SetControlText('XX_WHERE', 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"');

	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PREDEFINI',False);
			SetControlText   ('PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous);
		End;
	end
	else
	begin
		SetControlVisible('PREDEFINI', False);
		SetControlVisible('TPREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	{$ENDIF}
end ;

Procedure TOF_PGEDTSUIVIINTEGR.OnUpdate ;
var Pages : TPageControl;
    Where,OrderBy : String;
    i : Integer;
begin
  Inherited ;
  OrderBy := '';
  Pages := TPageControl(GetControl('Pages'));
  
  Where := RecupWhereCritere(Pages);

  For i := 1 to 4 do
  begin
        If GetControlText('VRUPT'+IntToStr(i)) <> '' then
        begin
                If OrderBy = '' then OrderBy := ' ORDER BY '+GetControlText('VRUPT'+IntToStr(i))
                else OrderBy := OrderBy+','+GetControltext('VRUPT'+IntToStr(i));
                SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
        end
        else SetControlText('XX_RUPTURE'+IntToStr(i),'');
  end;
  
  ChargeEdition;
end ;

procedure TOF_PGEDTSUIVIINTEGR.ChargeEdition;
var TobSalaries : Tob;
    T,T2,T3 : Tob;
    i : Integer;
    Where,WherePFI,WherePFO,StPages : String;
    Pages : TPageControl;
    Nb : Integer;
    Tablette,StSQL,WherePFOSuivi,OrderBy,TabletteRupt : String;
    TobInscReal,TobInscPrev,TSal : TOB;
    j,NbJoursPrem : Integer;
    PremDate : TDateTime;
    TotNbHeuresReal,TotNbHeuresPrev,TotCoutsPrev,TotCoutsReal,TotJoursReal,TotJoursPrev : Integer;

	Function CreateTobVide : TOB;
	var myTob : TOB;
	Begin
		myTob := TOB.Create('~LeSalarie', TobEtat, -1);
	
		myTob.AddChampSupValeur('SALARIE','');
    	myTob.AddChampSupValeur('PSA_TRAVAILN1','');
    	myTob.AddChampSupValeur('PSA_TRAVAILN2','');
    	myTob.AddChampSupValeur('PSA_TRAVAILN3','');
    	myTob.AddChampSupValeur('PSA_TRAVAILN4','');
    	myTob.AddChampSupValeur('LIBTRAVAILN1',VH_Paie.PGLibelleOrgStat1);
    	myTob.AddChampSupValeur('LIBTRAVAILN2',VH_Paie.PGLibelleOrgStat2);
    	myTob.AddChampSupValeur('LIBTRAVAILN3',VH_Paie.PGLibelleOrgStat3);
    	myTob.AddChampSupValeur('LIBTRAVAILN4',VH_Paie.PGLibelleOrgStat4);
		myTob.AddChampSupValeur('LIBSALARIE','');
		myTob.AddChampSupValeur('RESPONSABLE','');
		myTob.AddChampSupValeur('LIBRESPONSABLE','');
		myTob.AddChampSupValeur('LIBEMPLOI','');
		myTob.AddChampSupValeur('DATEENTREE',0);
		myTob.AddChampSupValeur('ETABLISSEMENT','');
		myTob.AddChampSupValeur('DATESUIVI',0);
		myTob.AddChampSupValeur('NBMOIS','');
	
		myTob.AddChampSupValeur('RUPTURE1', 		 '');
		myTob.AddChampSupValeur('RUPTURE2', 		 '');
		myTob.AddChampSupValeur('RUPTURE3', 		 '');
		myTob.AddChampSupValeur('RUPTURE4', 		 '');
	
		myTob.AddChampSupValeur('LIBRUPTURE1', 	 '');
		myTob.AddChampSupValeur('LIBRUPTURE2', 	 '');
		myTob.AddChampSupValeur('LIBRUPTURE3', 	 '');
		myTob.AddChampSupValeur('LIBRUPTURE4', 	 '');
	
		myTob.AddChampSupValeur('CODESTAGE',    '');
		myTob.AddChampSupValeur('LIBSTAGE',     '');
		myTob.AddChampSupValeur('DATEDEBUT',0);
		myTob.AddChampSupValeur('CURSUS',       '');
	
		myTob.AddChampSupValeur('NBHEURESPREV', 0);
		myTob.AddChampSupValeur('NBHEURESREAL', 0);
		myTob.AddChampSupValeur('NBHEURESECART',0);
	
		myTob.AddChampSupValeur('COUTPREV',     0);
		myTob.AddChampSupValeur('COUTREAL',     0);
		myTob.AddChampSupValeur('COUTECART',    0);
	
		myTob.AddChampSupValeur('NBJOURSPREV',    0);
		myTob.AddChampSupValeur('NBJOURSREAL',    0);
		myTob.AddChampSupValeur('NBJOURSECART',   0);
	
		myTob.AddChampSupValeur('NBJOURSPREM', '');
		myTob.AddChampSupValeur('NBJOURSPREMCOURT', '');
		
		myTob.AddChampSupValeur('NBTOTHEURESPREV', 0);
		myTob.AddChampSupValeur('NBTOTHEURESREAL', 0);
		myTob.AddChampSupValeur('NBTOTHEURESECART',0);
		
		myTob.AddChampSupValeur('TOTCOUTSPREV',     0);
		myTob.AddChampSupValeur('TOTCOUTSREAL',     0);
		myTob.AddChampSupValeur('TOTCOUTSECART',    0);
	
		Result := myTob;
	End;
begin
		{$IFNDEF EMANAGER}If PGBundleHierarchie Then Tablette := 'PGSALARIEINT' Else {$ENDIF}Tablette := 'PGSALARIE';
        Pages := TPageControl(GetControl('Pages'));

        Where := '';
        Where := RecupWhereCritere(Pages);

        // Chargement des salariés
        TobSalaries := TOB.Create('~LesSalaries', Nil, -1);
        StSQL := 'SELECT PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM, PSA_DATEENTREE, PSA_LIBELLEEMPLOI, PSA_ETABLISSEMENT, PSA_TRAVAILN1, PSA_TRAVAILN2, PSA_TRAVAILN3, PSA_TRAVAILN4, PSE_RESPONSFOR '+
        		 'FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE ';
        StSQL := StSQL + Where;
        {$IFNDEF EMANAGER}
        If PGBundleHierarchie Then StSQL := StSQL + DossiersAInterroger(GetControlText('PREDEFINI'),GetControlText('NODOSSIER'),'PSA',False,True);
        StSQL := AdaptMultiDosForm(StSQL);
        {$ENDIF}
        TobSalaries.LoadDetailFromSQL(StSQL+' ORDER BY PSE_RESPONSFOR,PSA_LIBELLE,PSA_PRENOM');

        // Chargement des inscriptions au plan de formation
        WherePFO := '';
        TobInscReal := TOB.Create('~LesInscReal', Nil, -1);
        StSQL := 'SELECT PFO_SALARIE,PFO_CODESTAGE, PST_LIBELLE, PFO_CURSUS, PCU_LIBELLE, PFO_NBREHEURE, PFO_AUTRECOUT, PFO_DATEDEBUT, '+
        		 'PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_TRAVAILN1,PFO_TRAVAILN2,PFO_TRAVAILN3,PFO_TRAVAILN4,PFO_FORMATION1,PFO_FORMATION2,PFO_FORMATION3,PFO_FORMATION4,PFO_FORMATION5,PFO_FORMATION6,PFO_FORMATION7,PFO_FORMATION8 '+ //ruptures
        		 'FROM FORMATIONS LEFT JOIN CURSUS ON PCU_CURSUS=PFO_CURSUS AND PCU_RANGCURSUS="0" '+
        		 'LEFT JOIN STAGE ON PST_CODESTAGE=PFO_CODESTAGE AND PST_MILLESIME="0000" '+
         		 'WHERE PFO_EFFECTUE="X" AND PFO_PGTYPESESSION<>"SOS" ';
	    If GetControlText('PSE_RESPONSFOR') <> '' Then WherePFO := WherePFO + ' AND PFO_RESPONSFOR="'+GetControlText('PSE_RESPONSFOR')+'" ';
	    If GetControlText('FORMATION1') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION1="'+GetControlText('FORMATION1')+'" ';
	    If GetControlText('FORMATION2') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION2="'+GetControlText('FORMATION2')+'" ';
	    If GetControlText('FORMATION3') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION3="'+GetControlText('FORMATION3')+'" ';
	    If GetControlText('FORMATION4') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION4="'+GetControlText('FORMATION4')+'" ';
	    If GetControlText('FORMATION5') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION5="'+GetControlText('FORMATION5')+'" ';
	    If GetControlText('FORMATION6') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION6="'+GetControlText('FORMATION6')+'" ';
	    If GetControlText('FORMATION7') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION7="'+GetControlText('FORMATION7')+'" ';
	    If GetControlText('FORMATION8') <> ''     Then WherePFO := WherePFO + ' AND PFO_FORMATION8="'+GetControlText('FORMATION8')+'" ';
	    WherePFOSuivi := '';
	    If GetControlText('PERIODESUIVI') <> ''   Then
	    Begin
	    	Nb := StrToInt(GetControlText('PERIODESUIVI'));
	    	Nb := -Nb;
	    	WherePFOSuivi := WherePFOSuivi + ' AND PFO_DATEDEBUT>="'+UsDatetime(PlusMois(Date, Nb))+'" ';
	    End;
        Try
    	    If (GetControlText('DATEFINSUIVI') <> '') And (GetControlText('DATEFINSUIVI') <> '  /  /    ')  Then WherePFOSuivi := WherePFOSuivi + ' AND PFO_DATEDEBUT<="'+UsDatetime(StrToDate(GetControlText('DATEFINSUIVI')))+'" ';
        Except
        End;
	    StSQL := StSQL + WherePFO + WherePFOSuivi;
	    OrderBy := ' ORDER BY PFO_SALARIE,PFO_DATEDEBUT,PFO_CODESTAGE';
	    If GetControlText('XX_RUPTURE1') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE1');
	    If GetControlText('XX_RUPTURE2') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE2');
	    If GetControlText('XX_RUPTURE3') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE3');
	    If GetControlText('XX_RUPTURE4') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE4');
        TobInscReal.LoadDetailFromSQL(StSQL+OrderBy);

        // Chargement des inscriptions au prévisionnel
        TobInscPrev := TOB.Create('~LesInscPrev', Nil, -1);
        StSQL := 'SELECT PFI_CODESTAGE, PST_LIBELLE, PFI_CURSUS, PCU_LIBELLE, PFI_DUREESTAGE, PFI_AUTRECOUT, PFI_JOURSTAGE, '+
        		 'PFI_ETABLISSEMENT,PFI_LIBEMPLOIFOR,PFI_TRAVAILN1,PFI_TRAVAILN2,PFI_TRAVAILN3,PFI_TRAVAILN4,PFI_FORMATION1,PFI_FORMATION2,PFI_FORMATION3,PFI_FORMATION4,PFI_FORMATION5,PFI_FORMATION6,PFI_FORMATION7,PFI_FORMATION8 '+ //ruptures
        		 'FROM INSCFORMATION LEFT JOIN CURSUS ON PCU_CURSUS=PFI_CURSUS AND PCU_RANGCURSUS="0" '+
        		 'LEFT JOIN STAGE ON PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME="0000" '+
        		 'WHERE PFI_MILLESIME="'+RendMillesimePrevisionnel()+'" AND PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--" ';
        WherePFI := ConvertPrefixe(WherePFO, 'PFO', 'PFI');
        StSQL := StSQL + WherePFI;
	    OrderBy := 'ORDER BY PFI_SALARIE,PFI_CODESTAGE';
	    If GetControlText('XX_RUPTURE1') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE1');
	    If GetControlText('XX_RUPTURE2') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE2');
	    If GetControlText('XX_RUPTURE3') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE3');
	    If GetControlText('XX_RUPTURE4') <> ''    Then OrderBy := OrderBy + ','+GetControlText('XX_RUPTURE4');
	    OrderBy := ConvertPrefixe(OrderBy, 'PFO','PFI');
        TobInscPrev.LoadDetailFromSQL(StSQL+OrderBy);

		// Création de la Tob de l'édition
        TobEtat := Tob.Create('~Les salariés pour édition',Nil,-1);

        For i := 0 to TobSalaries.Detail.Count-1 do
        begin
            TSal := CreateTobVide;

        	TSal.PutValue('SALARIE',        TobSalaries.Detail[i].GetValue('PSA_SALARIE'));
        	TSal.PutValue('PSA_TRAVAILN1',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN1'));
        	TSal.PutValue('PSA_TRAVAILN2',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN2'));
        	TSal.PutValue('PSA_TRAVAILN3',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN3'));
        	TSal.PutValue('PSA_TRAVAILN4',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN4'));
        	TSal.PutValue('LIBSALARIE',     TobSalaries.Detail[i].GetValue('PSA_LIBELLE')+' '+TobSalaries.Detail[i].GetValue('PSA_PRENOM'));
            TSal.PutValue('RESPONSABLE',    TobSalaries.Detail[i].GetValue('PSE_RESPONSFOR'));
        	TSal.PutValue('LIBRESPONSABLE', RechDom(Tablette, TobSalaries.Detail[i].GetValue('PSE_RESPONSFOR'), False));
        	TSal.PutValue('LIBEMPLOI',      RechDom('PGLIBEMPLOI', TobSalaries.Detail[i].GetValue('PSA_LIBELLEEMPLOI'), False));
        	TSal.PutValue('DATEENTREE',     StrToDate(TobSalaries.Detail[i].GetValue('PSA_DATEENTREE')));
        	TSal.PutValue('ETABLISSEMENT',  TobSalaries.Detail[i].GetValue('PSA_ETABLISSEMENT'));
        	TSal.PutValue('DATESUIVI',      Date);
        	TSal.PutValue('NBMOIS',         IntToStr(MonthsBetween(Date,TobSalaries.Detail[i].GetValue('PSA_DATEENTREE')))+' mois après la date d''entrée du salarié.');
        	TSal.PutValue('NBJOURSPREM', 'Le salarié n''a jamais suivi de formation depuis son entrée.');

			j := 0;
			T2 := Nil;
            PremDate := 0;
            NbJoursPrem := 0;
            TotNbHeuresPrev := 0;
            TotNbHeuresReal := 0;
            TotCoutsPrev    := 0;
            TotCoutsReal    := 0;
            TotJoursReal    := 0;
            TotJoursPrev    := 0;
			
        	// Prévisionnel d'abord
        	T := TobInscPrev.FindFirst(['PFI_SALARIE'],[TobSalaries.Detail[i].GetValue('PSA_SALARIE')], False);
        	While T <> Nil Do
        	Begin
        		If j > 0 Then 
        		Begin
                    T2 := CreateTobVide;
        			T2.PutValue('SALARIE',    TobSalaries.Detail[i].GetValue('PSA_SALARIE'));
		        	T2.PutValue('PSA_TRAVAILN1',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN1'));
		        	T2.PutValue('PSA_TRAVAILN2',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN2'));
		        	T2.PutValue('PSA_TRAVAILN3',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN3'));
		        	T2.PutValue('PSA_TRAVAILN4',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN4'));
		        	T2.PutValue('LIBSALARIE',     TobSalaries.Detail[i].GetValue('PSA_LIBELLE')+' '+TobSalaries.Detail[i].GetValue('PSA_PRENOM'));
		            T2.PutValue('RESPONSABLE',    TobSalaries.Detail[i].GetValue('PSE_RESPONSFOR'));
		        	T2.PutValue('LIBRESPONSABLE', RechDom(Tablette, TobSalaries.Detail[i].GetValue('PSE_RESPONSFOR'), False));
		        	T2.PutValue('LIBEMPLOI',      RechDom('PGLIBEMPLOI', TobSalaries.Detail[i].GetValue('PSA_LIBELLEEMPLOI'), False));
		        	T2.PutValue('DATEENTREE',     StrToDate(TobSalaries.Detail[i].GetValue('PSA_DATEENTREE')));
		        	T2.PutValue('ETABLISSEMENT',  TobSalaries.Detail[i].GetValue('PSA_ETABLISSEMENT'));
		        	T2.PutValue('DATESUIVI',      Date);
		        	T2.PutValue('NBMOIS',         IntToStr(MonthsBetween(Date,TobSalaries.Detail[i].GetValue('PSA_DATEENTREE')))+' mois après la date d''entrée du salarié.');
		        	T2.PutValue('NBJOURSPREM',    'Le salarié n''a jamais suivi de formation depuis son entrée.');
        		End
        		Else T2 := TSal;
        		
        		T2.PutValue('CODESTAGE',    T.GetValue('PFI_CODESTAGE'));
        		T2.PutValue('LIBSTAGE',     T.GetValue('PST_LIBELLE'));
        		T2.PutValue('CURSUS',       T.GetValue('PCU_LIBELLE'));
        		
        		T2.PutValue('NBHEURESPREV', T.GetValue('PFI_DUREESTAGE')); TotNbHeuresPrev := TotNbHeuresPrev + T.GetValue('PFI_DUREESTAGE');
        		T2.PutValue('NBHEURESREAL', 0);
        		T2.PutValue('NBHEURESECART',T.GetValue('PFI_DUREESTAGE'));
        		
        		T2.PutValue('COUTPREV',     T.GetValue('PFI_AUTRECOUT')); TotCoutsPrev := TotCoutsPrev + T.GetValue('PFI_AUTRECOUT');
        		T2.PutValue('COUTREAL',     0);
        		T2.PutValue('COUTECART',    T.GetValue('PFI_AUTRECOUT'));
        		
        		TotJoursPrev := TotJoursPrev + T.GetValue('PFI_JOURSTAGE');
        		(*T2.PutValue('NBJOURSPREV',  T.GetValue('PFI_JOURSTAGE')); 
        		T2.PutValue('NBJOURSREAL',  0);
        		T2.PutValue('NBJOURSECART', T.GetValue('PFI_JOURSTAGE'));*)
        		
        		If GetControlText('XX_RUPTURE1') <> '' Then 
        		Begin
        			T2.PutValue('RUPTURE1', T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE1'),'PFO','PFI')));
                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE1'));
                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE1',RechDom(TabletteRupt, T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE1'),'PFO','PFI')), False));
        		End;
				If GetControlText('XX_RUPTURE2') <> '' Then
				Begin
					T2.PutValue('RUPTURE2', T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE2'),'PFO','PFI')));
                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE2'));
                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE2',RechDom(TabletteRupt, T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE2'),'PFO','PFI')), False));
				End;
				If GetControlText('XX_RUPTURE3') <> '' Then
				Begin
					T2.PutValue('RUPTURE3', T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE3'),'PFO','PFI')));
                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE3'));
                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE3',RechDom(TabletteRupt, T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE3'),'PFO','PFI')), False));
				End;
				If GetControlText('XX_RUPTURE4') <> '' Then 
				Begin
					T2.PutValue('RUPTURE4', T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE4'),'PFO','PFI')));
                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE4'));
                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE4',RechDom(TabletteRupt, T.GetValue(ConvertPrefixe(GetControlText('XX_RUPTURE4'),'PFO','PFI')), False));
				End;
        		
        		T := TobInscPrev.FindNext(['PFI_SALARIE'],[TobSalaries.Detail[i].GetValue('PSA_SALARIE')], False);
        		j := j + 1;
        	End;
        	
        	// Réalisé ensuite
        	T := TobInscReal.FindFirst(['PFO_SALARIE'],[TobSalaries.Detail[i].GetValue('PSA_SALARIE')], False);
        	While T <> Nil Do
        	Begin
        		T3 := TobEtat.FindFirst(['SALARIE','CODESTAGE'],[T.GetValue('PFO_SALARIE'),T.GetValue('PFO_CODESTAGE')], False);
        		If T3 = Nil Then
        		Begin
	        		If j > 0 Then 
	        		Begin
                        T2 := CreateTobVide;
	        			T2.PutValue('SALARIE',    TobSalaries.Detail[i].GetValue('PSA_SALARIE'));
			        	T2.PutValue('PSA_TRAVAILN1',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN1'));
			        	T2.PutValue('PSA_TRAVAILN2',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN2'));
			        	T2.PutValue('PSA_TRAVAILN3',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN3'));
			        	T2.PutValue('PSA_TRAVAILN4',  TobSalaries.Detail[i].GetValue('PSA_TRAVAILN4'));
			        	T2.PutValue('LIBSALARIE',     TobSalaries.Detail[i].GetValue('PSA_LIBELLE')+' '+TobSalaries.Detail[i].GetValue('PSA_PRENOM'));
			            T2.PutValue('RESPONSABLE',    TobSalaries.Detail[i].GetValue('PSE_RESPONSFOR'));
			        	T2.PutValue('LIBRESPONSABLE', RechDom(Tablette, TobSalaries.Detail[i].GetValue('PSE_RESPONSFOR'), False));
			        	T2.PutValue('LIBEMPLOI',      RechDom('PGLIBEMPLOI', TobSalaries.Detail[i].GetValue('PSA_LIBELLEEMPLOI'), False));
			        	T2.PutValue('DATEENTREE',     StrToDate(TobSalaries.Detail[i].GetValue('PSA_DATEENTREE')));
			        	T2.PutValue('ETABLISSEMENT',  TobSalaries.Detail[i].GetValue('PSA_ETABLISSEMENT'));
			        	T2.PutValue('DATESUIVI',      Date);
			        	T2.PutValue('NBMOIS',         IntToStr(MonthsBetween(Date,TobSalaries.Detail[i].GetValue('PSA_DATEENTREE')))+' mois après la date d''entrée du salarié.');
	        		End
	        		Else T2 := TSal;
	        		
	        		T2.PutValue('CODESTAGE',    T.GetValue('PFO_CODESTAGE'));
	        		T2.PutValue('LIBSTAGE',     T.GetValue('PST_LIBELLE'));
	        		T2.PutValue('DATEDEBUT',     T.GetValue('PFO_DATEDEBUT'));
                    If (PremDate = 0) Or (PremDate > T.GetValue('PFO_DATEDEBUT')) Then PremDate := T.GetValue('PFO_DATEDEBUT');
	        		T2.PutValue('CURSUS',       T.GetValue('PCU_LIBELLE'));

	        		T2.PutValue('NBHEURESPREV', 0);
	        		T2.PutValue('NBHEURESREAL', T.GetValue('PFO_NBREHEURE'));
	        		T2.PutValue('NBHEURESECART',T.GetValue('PFO_NBREHEURE') * -1); TotNbHeuresReal := TotNbHeuresReal + T.GetValue('PFO_NBREHEURE');
	        		
	        		T2.PutValue('COUTPREV',     0);
	        		T2.PutValue('COUTREAL',     T.GetValue('PFO_AUTRECOUT'));
	        		T2.PutValue('COUTECART',    T.GetValue('PFO_AUTRECOUT') * -1); TotCoutsReal := TotCoutsReal + T.GetValue('PFO_AUTRECOUT');
	        		
	        		TotJoursReal := TotJoursReal + Round(T.GetValue('PFO_NBREHEURE') / 7);
	        		(*T2.PutValue('NBJOURSPREV',  0);
	        		T2.PutValue('NBJOURSREAL',  Round(T.GetValue('PFO_NBREHEURE') / 7)); 
	        		T2.PutValue('NBJOURSECART', Round(T.GetValue('PFO_NBREHEURE') / 7) * -1);*)
	        		
	        		If GetControlText('XX_RUPTURE1') <> '' Then T2.PutValue('RUPTURE1', T.GetValue(GetControlText('XX_RUPTURE1')));
					If GetControlText('XX_RUPTURE2') <> '' Then T2.PutValue('RUPTURE2', T.GetValue(GetControlText('XX_RUPTURE2')));
					If GetControlText('XX_RUPTURE3') <> '' Then T2.PutValue('RUPTURE3', T.GetValue(GetControlText('XX_RUPTURE3')));
					If GetControlText('XX_RUPTURE4') <> '' Then T2.PutValue('RUPTURE4', T.GetValue(GetControlText('XX_RUPTURE4')));
					
	        		If GetControlText('XX_RUPTURE1') <> '' Then
	        		Begin
	        			T2.PutValue('RUPTURE1', T.GetValue(GetControlText('XX_RUPTURE1')));
	                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE1'));
	                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE1',RechDom(TabletteRupt, T.GetValue(GetControlText('XX_RUPTURE1')), False));
	        		End;
					If GetControlText('XX_RUPTURE2') <> '' Then
					Begin
						T2.PutValue('RUPTURE2', T.GetValue(GetControlText('XX_RUPTURE2')));
	                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE2'));
	                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE2',RechDom(TabletteRupt, T.GetValue(GetControlText('XX_RUPTURE2')), False));
					End;
					If GetControlText('XX_RUPTURE3') <> '' Then
					Begin
						T2.PutValue('RUPTURE3', T.GetValue(GetControlText('XX_RUPTURE3')));
	                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE3'));
	                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE3',RechDom(TabletteRupt, T.GetValue(GetControlText('XX_RUPTURE3')), False));
					End;
					If GetControlText('XX_RUPTURE4') <> '' Then 
					Begin
						T2.PutValue('RUPTURE4', T.GetValue(GetControlText('XX_RUPTURE4')));
	                    TabletteRupt := RechTablette(GetControlText('XX_RUPTURE4'));
	                    If (TabletteRupt <> '') Then T2.PutValue('LIBRUPTURE4',RechDom(TabletteRupt, T.GetValue(GetControlText('XX_RUPTURE4')), False));
					End;
        		End
        		Else
        		Begin
        			T2 := T3;
	        		T2.PutValue('NBHEURESREAL', 		 T2.GetValue('NBHEURESREAL') + T.GetValue('PFO_NBREHEURE'));
	        		T2.PutValue('NBHEURESECART',		 T2.GetValue('NBHEURESPREV') - T2.GetValue('NBHEURESREAL'));
	        		
	        		T2.PutValue('COUTREAL', 		 	 T2.GetValue('COUTREAL') + T.GetValue('PFO_AUTRECOUT'));
	        		T2.PutValue('COUTECART',		 	 T2.GetValue('COUTPREV') - T2.GetValue('COUTREAL'));
	        		
	        		TotJoursReal := TotJoursReal + Round(T.GetValue('PFO_NBREHEURE') / 7);
	        		(*T2.PutValue('NBJOURSREAL', 		 	 T2.GetValue('NBJOURSREAL') + Round(T.GetValue('PFO_NBREHEURE') / 7));
	        		T2.PutValue('NBJOURSECART',		 	 T2.GetValue('NBJOURSPREV') - T2.GetValue('NBJOURSREAL'));*)
        		End;
        		
        		T := TobInscReal.FindNext(['PFO_SALARIE'],[TobSalaries.Detail[i].GetValue('PSA_SALARIE')], False);
        		j := j + 1;
        	End;
        	
        	If PremDate <> 0 Then
        	Begin
        		NbJoursPrem := PremDate - TSal.GetValue('DATEENTREE');
				TSal.PutValue('NBJOURSPREM', 'Le salarié a suivi sa première formation '+IntToStr(NbJoursPrem)+' jour(s) après son entrée.');
        		(*T2.PutValue('NBJOURSREAL', 		 	 T2.GetValue('NBJOURSREAL') + Round(T.GetValue('PFO_NBREHEURE') / 7));
        		T2.PutValue('NBJOURSECART',		 	 T2.GetValue('NBJOURSPREV') - T2.GetValue('NBJOURSREAL'));*)
				If T2 <> Nil Then
				Begin
					T2.PutValue('NBJOURSPREM',   'Le salarié a suivi sa première formation '+IntToStr(NbJoursPrem)+' jour(s) après son entrée.');
					//T2.PutValue('NBJOURSPREMCOURT', NbJoursPrem);
					
					(*T2.PutValue('NBJOURSPREV', 		 	 TotJoursPrev);
	        		T2.PutValue('NBJOURSREAL', 		 	 TotJoursReal);
	        		T2.PutValue('NBJOURSECART',		 	 TotJoursPrev-TotJoursReal);*)
				End;
        	End;
        	
			(*TSal.AddChampSupValeur('NBTOTHEURESPREV', TotNbHeuresPrev);
			TSal.AddChampSupValeur('NBTOTHEURESREAL', TotNbHeuresReal);
			T2.AddChampSupValeur('NBTOTHEURESECART',TotNbHeuresPrev-TotNbHeuresReal);
				
			T2.AddChampSupValeur('TOTCOUTSPREV',    TotCoutsPrev);
			T2.AddChampSupValeur('TOTCOUTSREAL',    TotCoutsReal);
			T2.AddChampSupValeur('TOTCOUTSECART',   TotCoutsPrev-TotCoutsReal);*)
			
        	If TSal <> Nil Then
        	Begin
				TSal.AddChampSupValeur('NBTOTHEURESPREV', TotNbHeuresPrev);
				TSal.AddChampSupValeur('NBTOTHEURESREAL', TotNbHeuresReal);
				TSal.AddChampSupValeur('NBTOTHEURESECART',TotNbHeuresPrev-TotNbHeuresReal);
				
				TSal.AddChampSupValeur('TOTCOUTSPREV',    TotCoutsPrev);
				TSal.AddChampSupValeur('TOTCOUTSREAL',    TotCoutsReal);
				TSal.AddChampSupValeur('TOTCOUTSECART',   TotCoutsPrev-TotCoutsReal);
				
				TSal.PutValue('NBJOURSPREMCOURT', 		  NbJoursPrem);
				
				TSal.PutValue('NBJOURSPREV', 		 	  TotJoursPrev);
        		TSal.PutValue('NBJOURSREAL', 		 	  TotJoursReal);
        		TSal.PutValue('NBJOURSECART',		 	  TotJoursPrev-TotJoursReal);
			End;
        end;

		FreeAndNil(TobInscPrev);
		FreeAndNil(TobInscReal);
		FreeAndNil(TobSalaries);
		
        StPages := '';
       {$IFDEF EAGLCLIENT}
        StPages := AglGetCriteres (Pages, FALSE);
        {$ENDIF}
        TFQRS1(Ecran).LaTob:= TobEtat;
end;

procedure TOF_PGEDTSUIVIINTEGR.SalarieElipsisClick(Sender : TObject);
begin
	ElipsisSalarieMultidos (Sender)
end;

procedure TOF_PGEDTSUIVIINTEGR.RespElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StOrder,StWhere : String;
    St : String;
{$ENDIF}
begin
	{$IFDEF EMANAGER}
	If TypeUtilisat = 'R' then StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
		' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
	else StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
		'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
	StOrder := 'PSI_LIBELLE';
	LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
	{$ELSE}
	ElipsisResponsableMultidos (Sender);
	{$ENDIF}
end;

(*Function TOF_PGEDTSUIVIINTEGR.ChampRuptureForm(TypeRupt,ValeurRupt : String) : String;
var Tablette : String;
    j : Integer;
begin
        If ValeurRupt = 'PFO_ETABLISSEMENT' then Tablette := 'TTETABLISSEMENT';
        If ValeurRupt = 'PFO_LIBEMPLOIFOR' then Tablette := 'PGLIBEMPLOI';
        For j := 1 to 4 do
        begin
                If ValeurRupt = 'PFO_TRAVAILN'+IntToStr(j) then
                begin
                Tablette := 'PGTRAVAILN'+IntToStr(j);
                end;
        end;
        Result := Tablette;
end;*)

procedure TOF_PGEDTSUIVIINTEGR.OnClose;
begin
  Inherited ;
  If (TobEtat <> nil) then FreeAndNil (TobEtat);
  SetLength(Tablettes,0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. : Rend (in)visible la checkbox du niveau de rupture
Suite ........ : précédent en fonction de la valeur saisie. Effectue également
Suite ........ : le contrôle de cohérence des ruptures.
Mots clefs ... : 
*****************************************************************}
Procedure TOF_PGEDTSUIVIINTEGR.OnChangeRuptures (Sender : TObject);
Var
  i,j : Integer;
  Valeur : String;
Begin
     // Récupération du niveau de rupture
     i := StrToInt(Copy(TControl(Sender).Name,6,1));

     // Contrôle de cohérence des ruptures
     For j:= 2 To 3 Do
     Begin
        If (THValComboBox(GetControl('VRUPT'+IntToStr(j))).Value = '') And
           (THValComboBox(GetControl('VRUPT'+IntToStr(j+1))).Value <> '') Then
          Begin
               PGIBox('Le niveau de rupture '+IntToStr(j)+' doit être renseigné',Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End
     End;

     Valeur := THValComboBox(GetControl('VRUPT'+IntToStr(i))).Value;
     For j:=1 To 4 Do
     Begin
          If (i <> j) And (Valeur = THValComboBox(GetControl('VRUPT'+IntToStr(j))).Value) And (Valeur <> '') Then
          Begin
               PGIBox('La rupture '+IntToStr(i)+' doit être différente de la '+IntToStr(j),Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End;
     End;

     GetControl('BValider').Enabled := True;
End;

Function TOF_PGEDTSUIVIINTEGR.RechTablette (NomChamp : String) : String;
Var
  Champ,Valeur : String;
  Q : TQuery;
  i : Integer;
Begin
  Valeur := '';
  If (NomChamp <> '') And (Length(NomChamp) > 5) Then
  Begin
    Champ := Copy(NomChamp,5,255);

    // Afin de limiter les accès en base, on sauvegarde les résultats précédents de recherche de tablette
    If (Length(Tablettes) > 0) Then
    Begin
          For i:=0 To Length(Tablettes)-1 Do
          Begin
                If (Tablettes[i][0] = Champ) Then Begin Valeur := Tablettes[i][1]; Break; End;
          End;
    End;

    // Recherche de la tablette concernée
    If (Valeur = '') Then
    Begin
          Q := OpenSql('SELECT DO_COMBO FROM DECOMBOS WHERE DO_NOMCHAMP like "%' + Champ + '%" ', True);
          Valeur := Q.FindField('DO_COMBO').AsString;
          Ferme(Q);
          // Sauvegarde de la valeur lue
          SetLength(Tablettes, Length(Tablettes)+1);
          Tablettes[Length(Tablettes)-1][0] := Champ;
          Tablettes[Length(Tablettes)-1][1] := Valeur;
    End;
  End;
  Result := Valeur;
End;
//FIN PT25


Initialization
registerclasses([TOF_PGEDITFRAISSAL,TOF_PGEDITSUIVIBUDFORM,TOF_PGFEUILLEPRESENCEFOR,TOF_PGEDITFACTUREFORM,TOF_PGEDITCOMPARATIFFOR,TOF_PGEDITBUDGETFOR,TOF_PGEDITCOMBUDGET,TOF_PGEDITCOMPSESSION,TOF_PGEDITFORREABUD,TOF_PGEDITPLANFORMATION,TOF_PGETATLIBREFOR,TOF_PGEDITSUIVICURSUS,TOF_PGEDITSUIVIFORM,TOF_PGEDTSUIVIINTEGR]);
end.

