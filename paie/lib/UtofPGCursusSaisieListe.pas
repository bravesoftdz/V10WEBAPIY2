{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 16/09/2002
Modifié le ... :   /  /
Description .. : Source TOF des FICHES de saisie en liste des cursus
Mots clefs ... : TOF;CURSUS
Contient les sources TOF pour les saisies en liste de cursus
*****************************************************************
PT1 | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
    |            |       |    | pour les critères code salarié uniquement
PT2 | 28/09/2007 | V_7   | FL | Emanager / Report / Adaptation cursus + accès assistant
PT3 | 24/10/2007 | V_8   | FL | Emanager / Report / Ajout inscription en masse des salariés
PT4 | 04/12/2007 | V_8   | FL | Emanager / Report / Gestion des sous-niveaux
PT5 | 15/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier 
PT6 | 02/04/2008 | V_803 | FL | Gestion des groupes de travail
PT7 | 17/04/2008 | V_804 | FL | Ne pas autoriser la modification d'éléments STD si pas droits multi + ajout de PGBundleCatalogue
    |            |       |    | + Gestion elipsis salarié et responsable
PT8 | 22/04/2008 | V_804 | FL | En multidossier, restriction des stages disponibles en fonction du prédéfini cursus
PT9 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT10| 04/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
}                                                        

unit UtofPGCursusSaisieListe;

interface
Uses StdCtrls,Controls,Classes,pgOutils,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,FE_Main,
{$ELSE}
     eQRS1, MaineAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,PgoutilsFormation,EntPaie,LookUp,ParamDat,HQry,SaisieList,uTableFiltre,HTB97 ;

Type
    TOF_PGSAISIELISTECURSUS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    TF : TTableFiltre;
    procedure NouveauCursus(Sender : TObject);
    procedure TreeViewDblClick (Sender : Tobject);
    procedure StageElipsisClick (Sender : Tobject);
  end ;

    TOF_PGSAISIELISTECURSUSPREV = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    TF : TTableFiltre;
    {$IFDEF EMANAGER}
    TypeUtilisat : String;
    {$ENDIF}
    Utilisateur : String;
    LeCursus : String; //PT3
    MillesimeEC : String; //PT3
    procedure TreeViewDblClick (Sender : Tobject);
    procedure SalarieElipsisClick (Sender : TObject);
//    procedure AfficheLesFormations(Cursus : String);
    {$IFDEF EMANAGER}
    procedure SaisieEnMasse    (Sender : TObject);
    procedure CreerStagePrev(Stage: string);
    function CalCoutBudgetFor(Stage, Millesime, Etab: string; NbInsc, Rang : Integer): Double;
    {$ENDIF}
  end ;

      TOF_PGSAISIELISTESESSIONS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    TF : TTableFiltre;
    procedure NouveauCursus(Sender : TObject);
    procedure TreeViewDblClick (Sender : Tobject);
  end ;

implementation

Uses ParamSoc,GalOutil;

{TOF_PGSAISIELISTECURSUS}
procedure TOF_PGSAISIELISTECURSUS.OnLoad ;
begin
  Inherited ;
    SetControLText('XX_WHERE', DossiersAInterroger(GetControlText('PCU_PREDEFINI'),GetControlText('NODOSSIER'),'PCU')); //PT5
end;

procedure TOF_PGSAISIELISTECURSUS.OnArgument (S : String ) ;
var BNewCursus : TToolBarButton97;
    THStage : THEdit;
begin
  Inherited ;
	BNewCursus := TToolbarButton97(GetControl('BNEWCURSUS'));
	If BNewCursus <> Nil then BNewCursus.OnClick := NouveauCursus;
	TF  :=  TFSaisieList(Ecran).LeFiltre;
	TF.LeTreeView.OnDblClick := TreeViewDblClick;
	
	//PT5
	If PGBundleCatalogue then //PT7
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False); //PT5
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PCU_PREDEFINI',False);
			SetControlText   ('PCU_PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT6
		
		THStage := THEdit(GetControl('PCC_CODESTAGE')); //PT8 Déplacement
		If THStage <> Nil then THStage.OnElipsisClick := StageElipsisClick;
	end
	else
	begin
		SetControlVisible('PCU_PREDEFINI', False);
		SetControlVisible('TPCU_PREDEFINI',False);
		SetControlVisible('NODOSSIER',     False); //PT5
		SetControlVisible('TNODOSSIER',    False);
	end;

    SetControLText('XX_WHERE', DossiersAInterroger(GetControlText('PCU_PREDEFINI'),GetControlText('NODOSSIER'),'PCU')); //PT5

	TFSaisieList(Ecran).BCherche.Click;
end ;

procedure TOF_PGSAISIELISTECURSUS.NouveauCursus(Sender : TObject);
begin
	AGLLanceFiche('PAY','CURSUS','','','ACTION=CREATION');
	TF.RefreshEntete;
end;

procedure TOF_PGSAISIELISTECURSUS.TreeViewDblClick (Sender : TObject) ;
var St,StAction,Pred : String ;
    Q : TQuery;
begin
	St := TF.TOBFiltre.GetValue('PCU_CURSUS')+';'+IntToStr(TF.TOBFiltre.GetValue('PCU_RANGCURSUS'));

    //PT7
    Q := OpenSQL('SELECT PCU_PREDEFINI FROM CURSUS WHERE PCU_CURSUS="'+TF.TOBFiltre.GetValue('PCU_CURSUS')+'" AND PCU_RANGCURSUS="'+IntToStr(TF.TOBFiltre.GetValue('PCU_RANGCURSUS'))+'"', True);
    If Not Q.EOF Then Pred := Q.FindField('PCU_PREDEFINI').AsString;
    Ferme(Q);
    If (Pred='STD') And (Not PGDroitMultiForm) Then StAction := 'CONSULTATION'
    Else StAction := 'MODIFICATION';

	AGLLanceFiche('PAY','CURSUS','',St,'ACTION='+StAction);
	TF.RefreshEntete(TF.TOBFiltre.GetValue('PCU_CURSUS'));
end ;

procedure TOF_PGSAISIELISTECURSUS.StageElipsisClick (Sender : Tobject);
var StWhere,StOrder,StListe : String;
begin
	//StWhere := StWhere + ' AND (PST_PREDEFINI="STD" OR (PST_NODOSSIER="'+V_PGI.NoDossier+'" AND PST_PREDEFINI="DOS")) AND PST_MILLESIME="0000"';
	StWhere := ' AND PST_MILLESIME="0000" ';

	// On restreint la liste des stages possibles en fonction du prédéfini/no dossier du cursus
	//If TF.TOBFiltre.GetValue('PCU_PREDEFINI') = 'STD' Then
	//Begin
		{If Not PGDroitMultiForm Then //PT8
			StWhere := StWhere + DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)
		Else
			StWhere := StWhere + DossiersAInterroger('','','PST',True,True);}
	//End
	//Else
	//Begin
	//	StWhere := StWhere + DossiersAInterroger('',TF.TOBFiltre.GetValue('PCU_NODOSSIER'),'PST',True,True)
	//End;
  	StWhere := StWhere + DossiersAInterroger(TF.TOBFiltre.GetValue('PCU_PREDEFINI'),TF.TOBFiltre.GetValue('PCU_NODOSSIER'),'PST',True,True);
	
	StListe := '(PST_LIBELLE||" "||PST_LIBELLE1) AS LIBELLE';
	If PGDroitMultiForm Then StListe := StListe + ',PST_PREDEFINI,DOS_LIBELLE'; //PT8
	StOrder := 'PST_LIBELLE,PST_LIBELLE1';

	LookupList(THEdit(Sender),'Liste des stages','STAGE LEFT JOIN DOSSIER ON PST_NODOSSIER=DOS_NODOSSIER','PST_CODESTAGE',StListe,StWhere,StOrder, True,-1);
end;

{TOF_PGSAISIELISTECURSUSPREV}
{procedure TOF_PGSAISIELISTECURSUSPREV.AfficheLesFormations(Cursus : String);
var Q : TQuery;
    TobStage : Tob;
    i : Integer;
begin
        Q := OpenSQL('SELECT PST_LIBELLE FROM STAGE LEFT JOIN CURSUSSTAGE ON PCC_CODESTAGE=PST_CODESTAGE AND PCC_MILLESIME=PST_MILLESIME '+
        'WHERE PCC_MILLESIME="0000" AND PCC_CURSUS="'+Cursus+'"',True);
        TobStage := Tob.Create('LesStages',Nil,-1);
        TobStage.LoadDetailDB('LesStages','','',Q,False);
        Ferme(Q);
        TFSaisieList(Ecran).LignesMessage := TobStage.Detail.Count + 1 ;
        TFSaisieList(Ecran).RichMessage1.Clear;
        TFSaisieList(Ecran).RichMessage1.Lines.Append('Liste des formations du cursus');
        For i := 0 to TobStage.Detail.Count - 1 do
        begin
      //          TFSaisieList(Ecran).RichMessage1.Lines.Append('- '+TobStage.Detail[i].GetValue('PST_LIBELLE')+TobStage.Detail[i].GetValue('PST_LIBELLE1'));
        end;
        TobStage.Free;
end;          }

procedure TOF_PGSAISIELISTECURSUSPREV.OnLoad ;
begin
  Inherited ;
        TFSaisieList(Ecran).BCherche.Click;
        //PT9
        // N'affiche pas les salariés confidentiels
        If TF.WhereTable <> '' Then TF.WhereTable := TF.WhereTable + ' AND ';
        If PGBundleHierarchie Then
        	TF.WhereTable := TF.WhereTable + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))' //PT10
        Else
        	TF.WhereTable := TF.WhereTable + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))'; //PT10
end;

procedure TOF_PGSAISIELISTECURSUSPREV.OnArgument (S : String ) ;
//var ArgMillesime,TypeSaisie,LeCursus,Utilisateur : String; //PT2
var TypeSaisie : String;
    Edit : THEdit;
    {$IFDEF EMANAGER}
    Bt : TToolbarbutton97; //PT3
    {$ENDIF}
begin
  Inherited ;
        TypeSaisie := ReadTokenSt(S);
        LeCursus := Trim(ReadTokenPipe(S,';'));
        //ArgMillesime := Trim(ReadTokenPipe(S,';')); //PT2
        MillesimeEC := Trim(ReadTokenPipe(S,';')); //PT2
        TF  :=  TFSaisieList(Ecran).LeFiltre;
//        TFSaisieList(Ecran).ZonesMessages := 1;
        Utilisateur := '';
//        AfficheLesFormations(LeCursus);
        SetControlVisible('BTree',False);
        {$IFNDEF EMANAGER}SetControlVisible('PANTREEVIEW',True);{$ENDIF} //PT2
        {$IFDEF EMANAGER}
        Utilisateur := ReadTokenSt(S);
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+Utilisateur+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        TFSaisieList(Ecran).ParamTreeView := False;
        TF.WhereTable := ' WHERE PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_CURSUS="'+LeCursus+'"'+
        ' AND '+AdaptByRespEmanager(TypeUtilisat,'PFI',Utilisateur,True);  //PT2
        SetControlProperty('PFI_LIBEMPLOIFOR','Plus','AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
                        'WHERE '+AdaptByRespEmanager(TypeUtilisat,'PSE',Utilisateur,False)+')'); //PT2
        {$ELSE}
        TF.WhereTable := ' WHERE PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_CURSUS="'+LeCursus+'"'; //PT2
        {$ENDIF}
        SetControlText('PCC_CURSUS',LeCursus);
        TF.LeTreeView.OnDblClick := TreeViewDblClick;
        If ExisteSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="'+LeCursus+'" AND PCC_MILLESIME="'+MillesimeEC+'"') then SetControlText('PCC_MILLESIME',MillesimeEC) //PT2
        else SetControlText('PCC_MILLESIME','0000');
        Edit := THEdit(GetControl('PFI_SALARIE'));
        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        {$IFDEF EMANAGER}
        //PT3 - Début
        Bt := TToolbarButton97(GetControl('BMULTISAL'));
        If Bt <> Nil then Bt.onClick := SaisieEnMasse;
        //PT3 - Fin
        {$ENDIF}
end ;

procedure TOF_PGSAISIELISTECURSUSPREV.SalarieElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
	If PGBundleHierarchie Then //PT7
	Begin
		ElipsisSalarieMultidos (Sender);
	End
	Else
	Begin
        StWhere := '';
        {$IFDEF EMANAGER}
        //PT4 - Début
//        If TypeUtilisat = 'R' then StWhere := 'PSE_RESPONSFOR="'+Utilisateur+'"'
//        else StWhere := 'PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+Utilisateur+'")';
		StWhere := StWhere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
		//PT4 - Fin
        {$ENDIF}
        StOrder := 'PSA_SALARIE';
//        LookupList(THEdit(Sender),'Liste des stages','SALARIES','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM','','', True,-1,St);
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
        LookupList(THEdit(Sender), 'Liste des salariés', 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
	End;
end;

procedure TOF_PGSAISIELISTECURSUSPREV.TreeViewDblClick (Sender : TObject) ;
var St : String ;
begin
        St := TF.TOBFiltre.GetValue('PCC_CODESTAGE');
        {$IFNDEF EMANAGER}   //PT2
        AGLLanceFiche('PAY','STAGE','',St,'ACTION=CONSULTATION')
        {$ELSE}
        AGLLanceFiche('PAY','EM_CATALOGUE','',St,'ACTION=CONSULTATION')
        {$ENDIF}
end ;

//PT3 - Début
{$IFDEF EMANAGER}
procedure TOF_PGSAISIELISTECURSUSPREV.SaisieEnMasse (Sender : TObject);
var Ret : String;
  Q: TQuery;
  TobLesFormations, T: Tob;
  i, NbInsc: Integer;
  CoutSal, TotalFrais,NbJourCur, CoutSalCur, TotalFraisCur, NbHCur,CtPedagCursus,CtPedag: Double;
  FraisH, FraisR, FraisT, Salaire, NbJour: Double;
  Req : String;
  Bt : TToolBarButton97;
  TauxChargeBudget : Double;
  TobCursusStage,TobSalaries,TS : TOB;
  Stage,ListeSalaries,RetSauve : String;
  Salarie : String;
  Rang : Integer;
begin
    Ret := AGLLanceFiche('PAY','EM_MULSALARIE','','','SAISIEMASSE');
    If Ret = '' then exit;

     Q := OpenSQL('SELECT PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="' + MillesimeEC + '"', True);
     if not Q.Eof then TauxChargeBudget := Q.FindField('PFE_TAUXBUDGET').AsFloat
     else TauxChargeBudget := 1;
     Ferme(Q);

     // Création du cursus au prévisionnel si ce n'est pas déjà fait
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
          FreeAndNil(TobCursusStage);
     end;

     TF.DisableTom;
     TF.StartUpdate;

     // Chargement de la liste des stages du cursus
     TobLesFormations := Tob.Create('LesFormations', nil, -1);
     TobLesFormations.LoadDetailFromSQL('SELECT PCC_CODESTAGE,PST_JOURSTAGE,PST_LIEUFORM,PST_DUREESTAGE,PST_NBANIM,' +
          'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_NATUREFORM ' + //PT17
          'FROM CURSUSSTAGE LEFT JOIN STAGE ON PCC_CODESTAGE=PST_CODESTAGE AND PCC_MILLESIME=PST_MILLESIME WHERE PCC_CURSUS="' + LeCursus +
          '" AND PCC_RANGCURSUS=0 AND PCC_MILLESIME="'+MillesimeEC+'"');

     // Chargement des données des salariés
     RetSauve := Ret;
     While RetSauve<>'' Do
     Begin
          ListeSalaries := ListeSalaries + '"' + ReadTokenSt(RetSauve) + '"';
          If RetSauve <> '' Then ListeSalaries := ListeSalaries + ',';
     End;
     TobSalaries := TOB.Create('LesSalaries', Nil, -1);
     TobSalaries.LoadDetailFromSQL('SELECT PSA_SALARIE,PSA_DADSCAT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,' +
          'PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI ' +
          'FROM SALARIES WHERE PSA_SALARIE IN ('+ListeSalaries+')');

     // Boucle sur les stagiaires sélectionnés
     While Ret <> '' Do
     Begin
          Salarie := ReadTokenSt(Ret);
          TS := TobSalaries.FindFirst(['PSA_SALARIE'],[Salarie], False);
          If TS = Nil Then Break;  // Ne doit pas passer par ici normalement!

          Q := OpenSQL('SELECT MAX (PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_ETABLISSEMENT="' + TS.GetValue('PSA_ETABLISSEMENT') + '" AND PFI_MILLESIME="' + MillesimeEC + '"', True);
          if not Q.eof then Rang := Q.FindField('RANG').AsInteger + 1
          else Rang := 1;
          Ferme(Q);

          CoutSalCur    := 0;
          TotalFraisCur := 0;
          NbHCur        := 0;
          NbJourCur     := 0;
          CtPedagCursus := 0;
          
          // Boucle sur les stages du cursus
          For i := 0 to TobLesFormations.Detail.Count - 1 Do
          Begin
               T := Tob.Create('INSCFORMATION', nil, -1);
               T.PutValue('PFI_SALARIE',Salarie);
               T.PutValue('PFI_TYPEPLANPREV',     'PLF');

               T.PutValue('PFI_MILLESIME',        MillesimeEC);
               T.PutValue('PFI_LIBELLE',          Rechdom('PGSALARIE',Salarie,False));
               T.PutValue('PFI_ETATINSCFOR',      'ATT');
               T.PutValue('PFI_RESPONSFOR',       Utilisateur);
               T.PutValue('PFI_NBINSC',           1);
               T.PutValue('PFI_RANG',             Rang);

               T.PutValue('PFI_ETABLISSEMENT',    TS.GetValue('PSA_ETABLISSEMENT'));
               T.PutValue('PFI_LIBEMPLOIFOR',     TS.GetValue('PSA_LIBELLEEMPLOI'));
               T.PutValue('PFI_TRAVAILN1',        TS.GetValue('PSA_TRAVAILN1'));
               T.PutValue('PFI_TRAVAILN2',        TS.GetValue('PSA_TRAVAILN2'));
               T.PutValue('PFI_TRAVAILN3',        TS.GetValue('PSA_TRAVAILN3'));
               T.PutValue('PFI_TRAVAILN4',        TS.GetValue('PSA_TRAVAILN4'));
               T.PutValue('PFI_CODESTAT',         TS.GetValue('PSA_CODESTAT'));
               T.PutValue('PFI_LIBREPCMB1',       TS.GetValue('PSA_LIBREPCMB1'));
               T.PutValue('PFI_LIBREPCMB2',       TS.GetValue('PSA_LIBREPCMB2'));
               T.PutValue('PFI_LIBREPCMB3',       TS.GetValue('PSA_LIBREPCMB3'));
               T.PutValue('PFI_LIBREPCMB4',       TS.GetValue('PSA_LIBREPCMB4'));
               T.PutValue('PFI_DADSCAT',          TS.GetValue('PSA_DADSCAT'));

               T.PutValue('PFI_CURSUS',           LeCursus);
               T.PutValue('PFI_CODESTAGE',        TobLesFormations.Detail[i].GetValue('PCC_CODESTAGE'));
               T.PutValue('PFI_JOURSTAGE',        TobLesFormations.Detail[i].GetValue('PST_JOURSTAGE'));
               T.PutValue('PFI_DUREESTAGE',       TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE'));
               T.PutValue('PFI_LIEUFORM',         TobLesFormations.Detail[i].GetValue('PST_LIEUFORM'));
               T.PutValue('PFI_FORMATION1',       TobLesFormations.Detail[i].GetValue('PST_FORMATION1'));
               T.PutValue('PFI_FORMATION2',       TobLesFormations.Detail[i].GetValue('PST_FORMATION2'));
               T.PutValue('PFI_FORMATION3',       TobLesFormations.Detail[i].GetValue('PST_FORMATION3'));
               T.PutValue('PFI_FORMATION4',       TobLesFormations.Detail[i].GetValue('PST_FORMATION4'));
               T.PutValue('PFI_FORMATION5',       TobLesFormations.Detail[i].GetValue('PST_FORMATION5'));
               T.PutValue('PFI_FORMATION6',       TobLesFormations.Detail[i].GetValue('PST_FORMATION6'));
               T.PutValue('PFI_FORMATION7',       TobLesFormations.Detail[i].GetValue('PST_FORMATION7'));
               T.PutValue('PFI_FORMATION8',       TobLesFormations.Detail[i].GetValue('PST_FORMATION8'));
               T.PutValue('PFI_NATUREFORM',       TobLesFormations.Detail[i].GetValue('PST_NATUREFORM'));
               T.PutValue('PFI_HTPSTRAV',         TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE'));
               T.PutValue('PFI_HTPSNONTRAV',      0);
               T.PutValue('PFI_MOTIFINSCFOR',     '2');
               T.PutValue('PFI_NIVPRIORITE',      '01');
               T.PutValue('PFI_DATEACCEPT',       IDate1900);
               T.PutValue('PFI_REFUSELE',         IDate1900);
               T.PutValue('PFI_REPORTELE',        IDate1900);
               T.PutValue('PFI_REALISE',          '-');

               If VH_Paie.PGForValoSalairePrev = 'VCR' then
                    Salaire := ForTauxHoraireReel(Salarie,0,0,'',ValPrevisionnel)
               else
                    Salaire := ForTauxHoraireCategoriel(T.Getvalue('PFI_LIBEMPLOIFOR'), T.Getvalue('PFI_MILLESIME'));
               CoutSal := Salaire * TobLesFormations.Detail[i].GetValue('PST_JOURSTAGE') * T.Getvalue('PFI_NBINSC');
               Q := OpenSQL('SELECT PSE_RESPONSFOR FROM DEPORTSAL WHERE PSE_SALARIE="' + T.Getvalue('PFI_SALARIE') + '"', True);
               if not Q.eof then
                    T.PutValue('PFI_RESPONSFOR', Q.Findfield('PSE_RESPONSFOR').AsString);
               Ferme(Q);
               T.PutValue('PFI_COUTREELSAL',      CoutSal * TauxChargeBudget);
               CoutSalCur := CoutSalCur + CoutSal;

               Req := 'SELECT PFF_FRAISHEBERG,PFF_FRAISREPAS,PFF_FRAISTRANSP FROM FORFAITFORM ' +
                      'WHERE PFF_MILLESIME="' + T.GetValue('PFI_MILLESIME') + '" AND PFF_LIEUFORM="' + T.GetValue('PFI_LIEUFORM') + '"';
               If (VH_Paie.PGForGestFraisByPop) Then
                    Req := Req + ' AND PFF_POPULATION IN (SELECT PNA_POPULATION FROM SALARIEPOPUL WHERE PNA_SALARIE="'+Salarie+'")'
               Else
                    Req := Req + ' AND PFF_ETABLISSEMENT="' + T.GetValue('PFI_ETABLISSEMENT') + '"';
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
               NbJour := T.GetValue('PFI_JOURSTAGE');
               NbInsc := T.GetValue('PFI_NBINSC');
               FraisH := FraisH * (NbJour - 1);
               if FraisH < 0 then FraisH := 0;
               If FraisH > 0 then FraisR := FraisR*((NbJour*2)-1)
               else FraisR := FraisR*NbJour;
               if FraisR < 0 then FraisR := 0;
               TotalFrais := FraisH + FraisR + FraisT;
               TotalFrais := TotalFrais * NbInsc;
               T.PutValue('PFI_FRAISFORFAIT', TotalFrais);

               TotalFraisCur := TotalFraisCur + TotalFrais;
               NbHCur := NbHcur + (TobLesFormations.Detail[i].GetValue('PST_DUREESTAGE') * NbInsc);
               NbJourCur := NbJourCur + (NbJour * NbInsc);
               CtPedag := CalCoutBudgetFor(TobLesFormations.Detail[i].GetValue('PCC_CODESTAGE'), MillesimeEC, TS.GetValue('PSA_ETABLISSEMENT'), NbInsc, Rang);
               CtPedagCursus := CtPedagCursus + CtPedag;
               T.PutValue('PFI_AUTRECOUT', CtPedag);
               T.InsertOrUpdateDB;
               if T <> nil then T.Free;
          end;

          // Création de l'inscription au cursus
          TF.Insert;
          TF.PutValue('PFI_SALARIE',         Salarie);
          TF.PutValue('PFI_CURSUS',          LeCursus);
          TF.PutValue('PFI_MILLESIME',       MillesimeEC);
          TF.PutValue('PFI_CODESTAGE',       '--CURSUS--');
          TF.PutValue('PFI_LIBELLE',         Rechdom('PGSALARIE',Salarie,False));
          TF.PutValue('PFI_ETATINSCFOR',     'ATT');
          TF.PutValue('PFI_RESPONSFOR',      Utilisateur);
          TF.PutValue('PFI_NBINSC',          1);
          TF.PutValue('PFI_TYPEPLANPREV',    'PLF');
          TF.PutValue('PFI_RANG',            Rang);
          TF.PutValue('PFI_ETABLISSEMENT',   TS.GetValue('PSA_ETABLISSEMENT'));
          TF.PutValue('PFI_LIBEMPLOIFOR',    TS.GetValue('PSA_LIBELLEEMPLOI'));
          TF.PutValue('PFI_TRAVAILN1',       TS.GetValue('PSA_TRAVAILN1'));
          TF.PutValue('PFI_TRAVAILN2',       TS.GetValue('PSA_TRAVAILN2'));
          TF.PutValue('PFI_TRAVAILN3',       TS.GetValue('PSA_TRAVAILN3'));
          TF.PutValue('PFI_TRAVAILN4',       TS.GetValue('PSA_TRAVAILN4'));
          TF.PutValue('PFI_CODESTAT',        TS.GetValue('PSA_CODESTAT'));
          TF.PutValue('PFI_LIBREPCMB1',      TS.GetValue('PSA_LIBREPCMB1'));
          TF.PutValue('PFI_LIBREPCMB2',      TS.GetValue('PSA_LIBREPCMB2'));
          TF.PutValue('PFI_LIBREPCMB3',      TS.GetValue('PSA_LIBREPCMB3'));
          TF.PutValue('PFI_LIBREPCMB4',      TS.GetValue('PSA_LIBREPCMB4'));
          TF.PutValue('PFI_DADSCAT',         TS.GetValue('PSA_DADSCAT'));
          TF.PutValue('PFI_MOTIFINSCFOR',    '2');
          TF.PutValue('PFI_NIVPRIORITE',     '01');
          TF.PutValue('PFI_COUTREELSAL',     CoutSalCur * TauxChargeBudget);
          TF.PutValue('PFI_FRAISFORFAIT',    TotalFraisCur);
          TF.PutValue('PFI_DUREESTAGE',      NbHCur);
          TF.PutValue('PFI_HTPSTRAV',        NbHCur);
          TF.PutValue('PFI_HTPSNONTRAV',     0);
          TF.PutValue('PFI_JOURSTAGE',       NbJourCur);
          TF.PutValue('PFI_AUTRECOUT',       CtPedagCursus);
          TF.PutValue('PFI_DATEACCEPT',      IDate1900);
          TF.PutValue('PFI_REFUSELE',        IDate1900);
          TF.PutValue('PFI_REPORTELE',       IDate1900);
          TF.PutValue('PFI_REALISE',         '-');
          TF.PutValue('PFI_TYPOFORMATION',   '001');
          TF.PutValue('PFI_PREDEFINI',       'DOS');
          TF.PutValue('PFI_NODOSSIER',       '000000');
          TF.Post;
     End;

     TF.EnableTOM;
     TF.EndUpdate;

     if TobLesFormations <> nil then FreeAndNil(TobLesFormations);
     if TobSalaries      <> nil then FreeAndNil(TobSalaries);

     TF.RefreshEntete;
     Bt := TToolBarButton97(GetControl('BCHERCHE'));
     if Bt <> nil then Bt.Click;
end;

function TOF_PGSAISIELISTECURSUSPREV.CalCoutBudgetFor(Stage, Millesime, Etab : String; NbInsc, Rang: Integer): Double;
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
    'AND (PFI_ETABLISSEMENT<>"' + Etab + '" OR PFI_RANG<>' + IntToStr(Rang) + ')', True);
  if not Q.Eof then NbInscStage := Q.FindField('NBINSC').AsInteger
  else NbInscStage := 0;
  Ferme(Q);
  NbInscStage := NbInscStage + NbInsc;
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
  Result := (NbInsc) * (coutFonc + CoutAnim + Coutunit);
end;

procedure TOF_PGSAISIELISTECURSUSPREV.CreerStagePrev(Stage: string);
var
  Q, QAnim: TQuery;
  TobStage, T, TobInvest, TI: Tob;
  NbHeure: Integer;
  SalaireAnim, NbAnim: Double;
begin
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
      NbHeure := T.GetValue('PST_DUREESTAGE');
      NbAnim := T.GetValue('PST_NBANIM');
      T.ChangeParent(TobStage, -1);
      T.PutValue('PST_MILLESIME', MillesimeEC);
      T.PutValue('PST_COUTSALAIR', SalaireAnim * NbHeure * NbAnim);
      T.InsertOrUpdateDB;
    end;
    TobStage.Free;
    // Duplication investissement pour la formation
    Q := OpenSQL('SELECT * FROM INVESTSESSION ' +
      'WHERE PIS_CODESTAGE="' + Stage + '" AND PIS_ORDRE=-1 AND PIS_MILLESIME="0000"', True);
    TobInvest := Tob.Create('INVESTSESSION', nil, -1);
    TobInvest.LoadDetailDB('INVESTSESSION', '', '', Q, False);
    Ferme(Q);
    TI := TobInvest.FindFirst(['PIS_CODESTAGE'], [Stage], False);
    while TI <> nil do
    begin
      TI.ChangeParent(TobInvest, -1);
      TI.PutValue('PIS_MILLESIME', MillesimeEC);
      TI.InsertOrUpdateDB(False);
      TI := TobInvest.FindNext(['PIS_CODESTAGE'], [Stage], False);
    end;
    TobInvest.Free;
  end;
end;
//PT3 - Fin
{$ENDIF}

{TOF_PGSAISIELISTESESSIONS}
procedure TOF_PGSAISIELISTESESSIONS.OnLoad ;
begin
  Inherited ; 
end;

procedure TOF_PGSAISIELISTESESSIONS.OnArgument (S : String ) ;
var BNewCursus : TToolBarButton97;
begin
  Inherited ;
        BNewCursus := TToolbarButton97(GetControl('BNEWCURSUS'));
        If BNewCursus <> Nil then BNewCursus.OnClick := NouveauCursus;
        TF  :=  TFSaisieList(Ecran).LeFiltre;
        TF.LeTreeView.OnDblClick := TreeViewDblClick;
end ;

procedure TOF_PGSAISIELISTESESSIONS.NouveauCursus(Sender : TObject);
begin
        AGLLanceFiche('PAY','CURSUS','','','SESSIONCURSUS;ACTION=CREATION');
        TF.RefreshEntete;
end;

procedure TOF_PGSAISIELISTESESSIONS.TreeViewDblClick (Sender : TObject) ;
var St : String ;
begin
        St := TF.TOBFiltre.GetValue('PCU_CURSUS')+';'+IntToStr(TF.TOBFiltre.GetValue('PCU_RANGCURSUS'));
        AGLLanceFiche('PAY','CURSUS','',St,'SESSIONCURSUS;ACTION=MODIFICATION');
end ;


Initialization
registerclasses([TOF_PGSAISIELISTECURSUS,TOF_PGSAISIELISTECURSUSPREV,TOF_PGSAISIELISTESESSIONS]);
end.

