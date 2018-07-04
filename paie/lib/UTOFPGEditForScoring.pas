{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 05/12/2003
Modifié le ... :   /  /
Description .. : Edition scoring formation
Mots clefs ... : EDIT_SCORING,FORMATION
*****************************************************************
PT1  | 16/11/2004 | V_60  | JL | Correction conversion type variant incorrect
PT2  | 26/04/2006 |       |    | Ajout TOF pour analtyse scoring
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT3  | 15/11/2006 | V_70  | JL | FQ 13682 Correction affichage libellé réponse
PT4  | 24/05/2007 | V_720 | JL | Modif jointure sur tablette réponse
PT5  | 24/05/2007 | V_720 | FL | FQ 13682 Gestion des ruptures et correction requêtes
PT6  | 22/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT7  | 02/04/2008 | V_803 | FL | Adaptation partage formation
PT8  | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue
}

unit UTOFPGEditForScoring;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, UtilEAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1, EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,  ParamSoc, HQry, UTob, HTB97, ed_tools,PGOutils2,entpaie,
  HStatus,P5DEF,PGOutilsFormation;

Const NB_RUPTURES = 3; //PT5

type TFunctionPointer = Function(Champs : array of string; Valeurs : array of variant; Multiniveau : boolean) : TOB of object;  //PT5

type
  TOF_PGFOREDITSCORING = class(TOF)
  private
    TobEtat : Tob;
    Tablettes : Array of Array[0..1] of String; //PT5
    procedure EditionQuestionnaire(Sender: TObject);
    procedure ExitEdit(Sender : TObject);
    procedure OnChangeRuptures (Sender : TObject); //PT5
    Function  RechTablette (NomChamp : String) : String;  //PT5
    Function  ChercheResultats(TobResultats, TobRuptures: TOB; NumQuestion : String; Valeur : Variant; Premier: Boolean): TOB; //PT5
    Function  ChercheEtat(TobEtat, TobValeurs : TOB; NumQuestion : Integer; Premier: Boolean): TOB; //PT5
  public
    procedure OnUpdate                     ; override;
    procedure OnClose                      ; override;
    procedure OnArgument(Arguments: string); override;
  end;

type
  TOF_PGCUBESCORING = class(TOF)
    procedure OnArgument(Arguments: string); override;
    private
    procedure ExitEdit(Sender : TObject);
  end;


implementation

Uses UtilPGI;


{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : ??/??/??
Modifié le ... : 24/05/2007  / PT5
Description .. : Création de l'état
Mots clefs ... : SCORING
*****************************************************************}
procedure TOF_PGFOREDITSCORING.OnUpdate;
var
  Q: TQuery;
  TobQuest,TobReponses,TobResultats,TobRuptures, T,TPRep,TPResult, TIndice : Tob;
  i, r, cpt, NumQuestion, NbRepTot, NbRep: Integer;
  Pourcentage: Double;
  OrderBy,Where,Select,SelectRupt,Tri,Requete,GroupBy,Tablette : String;
  AvecLibSession, First : Boolean;
  Req : String; //PT7
begin
  inherited;

     If tobEtat <> Nil then FreeAndNil(TobEtat);
     
     AvecLibSession := False;
     Select     := '';
     SelectRupt := '';
     Tri        := '';
     OrderBy    := '';
     GroupBy    := '';

     // Construction dynamique du SELECT,GROUPBY et ORDERBY en fonction des ruptures
     For i := NB_RUPTURES downto 1 do
     Begin
          SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
          If (GetControlText('VRUPT'+IntToStr(i)) <> '') Then
          Begin
               If (GetControlText('VRUPT'+IntToStr(i)) <> 'PFO_ORDRE') And (GetControlText('VRUPT'+IntToStr(i)) <> 'PFO_CODESTAGE') Then
               Begin
                    OrderBy := GetControlText('VRUPT'+IntToStr(i))+','+OrderBy;
                    GroupBy := GetControlText('VRUPT'+IntToStr(i))+','+GroupBy;
                    Select  := GetControlText('VRUPT'+IntToStr(i))+','+Select;
                    Tri     := GetControlText('VRUPT'+IntToStr(i))+';'+Tri;
               End;

               SelectRupt := GetControlText('VRUPT'+IntToStr(i))+','+SelectRupt;
               If GetControlText('VRUPT'+IntToStr(i)) = 'PFO_ORDRE' Then AvecLibSession := True;
          End;
     End;

     // Récupération des éléments composant la clause WHERE
     Where := RecupWhereCritere(TPageControl(GetControl('PAGES')));

     // Création de la TOB principale
     TobEtat := Tob.Create('Edition', nil, -1);
     
     // Liste des Thèmes / Questions / Tablettes
     Q := OpenSQL('SELECT * FROM SCORINGDEF WHERE POD_NUMSCORING > 0',True);
     TobQuest := Tob.Create('LesQuestions', nil, -1);
     TobQuest.LoadDetailDB('LesQuestions', '', '', Q, False);
     Ferme(Q);

     // Liste des réponses associées aux questions
     //PT7 - Début
     Req := 'SELECT CC_CODE, CC_LIBELLE, CC_LIBRE FROM ';
     If PGBundleInscFormation Then
     	Req := Req + GetBase(GetBasePartage(BUNDLE_FORMATION),'CHOIXCOD')
     Else
     	Req := Req + 'CHOIXCOD';
     Req := Req + ' WHERE CC_TYPE="PSC" ORDER BY CC_CODE';
     //PT7 - Fin
     Q := OpenSQL(Req, True);
     TobReponses := Tob.Create('LesReponses', nil, -1);
     TobReponses.LoadDetailDB('LesReponses', '', '', Q, False);
     Ferme(Q);

     // Liste des ruptures
     If SelectRupt <> '' Then
     Begin
          Requete := 'SELECT DISTINCT ' + Copy(SelectRupt, 1, Length(SelectRupt) -1);
          If AvecLibSession Then Requete := Requete + ',PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN';
          Requete := Requete + ' FROM FORMATIONS RIGHT JOIN SCORING ON PFO_CODESTAGE=PSC_CODESTAGE AND PFO_ORDRE=PSC_ORDRE AND PFO_SALARIE=PSC_SALARIE';
          If AvecLibSession Then Requete := Requete + ' LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE ';
          If Where <> '' Then Requete := Requete + ' ' + Where;
          If PGBundleInscFormation Then Requete := Requete + ' AND (PSS_PREDEFINI="STD" OR (PSS_PREDEFINI="DOS" AND PSS_NODOSSIER="'+V_PGI.NoDossier+'"))'; //PT7
          Q := OpenSQL(Requete, True);
          TobRuptures := Tob.Create('LesRuptures', Nil, -1);
          TobRuptures.LoadDetailDB('LesRuptures', '', '', Q, False);
          Ferme(Q);
          cpt := TobRuptures.Detail.Count - 1;
     End
     Else cpt := 0; // Aucune rupture : il faut faire le traitement au moins une fois

     // Construction de la requête des résultats
     Requete := 'SELECT ' + Select;
     For i := 0 to TobQuest.Detail.Count - 1 Do Requete := Requete + ' PSC_SCFO'+IntToStr(i+1)+',';
     Requete := Requete + ' PFO_CODESTAGE,PFO_ORDRE,PSC_SALARIE';
     Requete := Requete + ' FROM FORMATIONS JOIN SCORING ON PSC_CODESTAGE=PFO_CODESTAGE AND PSC_ORDRE=PFO_ORDRE AND PSC_SALARIE=PFO_SALARIE ';
     If (Where <> '') Then Requete := Requete + Where;
     If PGBundleInscFormation Then Requete := Requete + ' AND (PFO_PREDEFINI="STD" OR (PFO_PREDEFINI="DOS" AND PFO_NODOSSIER="'+V_PGI.NoDossier+'"))'; //PT7
     Requete := Requete + ' GROUP BY ' + GroupBy;
     For i := 0 to TobQuest.Detail.Count - 1 Do Requete := Requete + ' PSC_SCFO'+IntToStr(i+1)+',';
     Requete := Requete + 'PFO_CODESTAGE,PFO_ORDRE,PSC_SALARIE';

     Q := OpenSQL(Requete, True);
     If Q.EOF Then
     Begin
          // Libération des TOBs
          FreeAndNil(TobQuest);
          FreeAndNil(TobReponses);
          FreeAndNil(TobRuptures);
          TFQRS1(Ecran).LaTob := TobEtat;
          Exit;
     End;
     TobResultats := Tob.Create('LesResultats', Nil, -1);
     TobResultats.LoadDetailDB('LesResultats', '', '', Q, False);
     Ferme(Q);

     // Démarrage de l'écran d'attente
     InitMoveProgressForm(Nil, 'Chargement des données', 'Veuillez patienter SVP ...', cpt+1, False, True);

     // Traitement pour chaque rupture
     For r := 0 To cpt Do
     Begin
       // Mise à jour du libellé d'attente
       MoveCurProgressForm('Détail '+IntToStr(r+1)+' / '+IntToStr(cpt+1));

       // Parcours des questions
       For i := 0 to TobQuest.Detail.Count - 1 Do
       Begin
          // Récupération de la question
          NumQuestion := TobQuest.Detail[i].GetValue('POD_NUMSCORING');
          NbRepTot    := 0;
          TIndice     := Nil;
          Tablette    := Copy(TobQuest.Detail[i].GetValue('POD_TABLESCORE'),3,1);  //Récupération du n° de tablette : SC1, SC2

          // Tri des informations pour la question en cours
          TobResultats.Detail.Sort(Tri+'PSC_SCFO'+IntToStr(NumQuestion)+';PFO_CODESTAGE;PFO_ORDRE;PSC_SALARIE');

          // Parcours des réponses
          First := True;
          TPRep := TobReponses.FindFirst(['CC_LIBRE'],['SC'+Tablette], False);
          While TPRep <> Nil Do
          Begin

            // Recherche des résultats correspondant à la réponse en cours
            If (GetControlText('XX_RUPTURE1') = '') Then
                 TPResult := ChercheResultats (TobResultats, Nil, IntToStr(NumQuestion), TPRep.GetValue('CC_CODE'), True)
            Else
                 TPResult := ChercheResultats (TobResultats, TobRuptures.Detail[r], IntToStr(NumQuestion), TPRep.GetValue('CC_CODE'), True);

            If (TPResult <> Nil) And (TPResult.GetValue('PSC_SCFO'+IntToStr(NumQuestion)) <> '') Then
            Begin
               While (TPResult <> Nil) And (TPResult.GetValue('PSC_SCFO'+IntToStr(NumQuestion)) <> '') Do
               Begin
                 // Recherche d'un enregistrement identique
                 T := Nil;
                 If (GetControlText('XX_RUPTURE1') = '') Then
                   T := TobEtat.FindFirst(['POD_NUMSCORING','CODEREPONSE'],
                                          [NumQuestion,
                                          TPResult.GetValue('PSC_SCFO'+IntToStr(NumQuestion))], False)
                 Else If (GetControlText('XX_RUPTURE3') <> '') Then
                   T := TobEtat.FindFirst(['POD_NUMSCORING','CODERUPTURE1','CODERUPTURE2','CODERUPTURE3','CODEREPONSE'],
                                          [NumQuestion,
                                          TPResult.GetValue(GetControlText('XX_RUPTURE1')),
                                          TPResult.GetValue(GetControlText('XX_RUPTURE2')),
                                          TPResult.GetValue(GetControlText('XX_RUPTURE3')),
                                          TPResult.GetValue('PSC_SCFO'+IntToStr(NumQuestion))], False)
                 Else If (GetControlText('XX_RUPTURE2') <> '') Then
                   T := TobEtat.FindFirst(['POD_NUMSCORING','CODERUPTURE1','CODERUPTURE2','CODEREPONSE'],
                                          [NumQuestion,
                                          TPResult.GetValue(GetControlText('XX_RUPTURE1')),
                                          TPResult.GetValue(GetControlText('XX_RUPTURE2')),
                                          TPResult.GetValue('PSC_SCFO'+IntToStr(NumQuestion))], False)
                 Else If (GetControlText('XX_RUPTURE1') <> '') Then
                   T := TobEtat.FindFirst(['POD_NUMSCORING','CODERUPTURE1','CODEREPONSE'],
                                          [NumQuestion,
                                          TPResult.GetValue(GetControlText('XX_RUPTURE1')),
                                          TPResult.GetValue('PSC_SCFO'+IntToStr(NumQuestion))], False);

                 If T = Nil Then
                 Begin
                   // Création de la TOB
                    T := Tob.Create('FilleEdition', TobEtat, -1);
                    T.AddChampSupValeur('POD_NUMSCORING', NumQuestion, False);
                    T.AddChampSupValeur('POD_LIBELLE', TobQuest.Detail[i].GetValue('POD_LIBELLE'), False);
                    T.AddChampSupValeur('LIBTHEME', RechDom('PGFTHEMESCORING', TobQuest.Detail[i].GetValue('POD_THEMESCORE'), False), False);  //PT5
                    T.AddChampSupValeur('POD_THEMESCORE', TobQuest.Detail[i].GetValue('POD_THEMESCORE'), False);
                    T.AddChampSupValeur('CODEREPONSE',TPRep.GetValue('CC_CODE'));
                    T.AddChampSupValeur('REPONSE', RechDom('PGFSCORING'+Tablette, TPRep.GetValue('CC_CODE'), False), False);
                    T.AddChampSupValeur('NBREPONSE', 1); // Nb réponses initialisé à 1
                    T.AddChampSupValeur('NBREPTOT',  0);
                    T.AddChampSupValeur('POURCENTAGE', 0);

                    // Ruptures
                    T.AddChampSupValeur('CODERUPTURE1',''); T.AddChampSupValeur('CODERUPTURE2',''); T.AddChampSupValeur('CODERUPTURE3','');
                    T.AddChampSupValeur('RUPTURE1',''); T.AddChampSupValeur('RUPTURE2',''); T.AddChampSupValeur('RUPTURE3','');
                    If (GetControlText('XX_RUPTURE1') <> '') Then
                    Begin
                         T.PutValue('CODERUPTURE1',TPResult.GetValue(GetControlText('XX_RUPTURE1')));
                         T.PutValue('RUPTURE1',RechDom(RechTablette(GetControlText('XX_RUPTURE1')),TPResult.GetValue(GetControlText('XX_RUPTURE1')),False));
                    End;
                    If (GetControlText('XX_RUPTURE2') <> '') Then
                    Begin
                         T.PutValue('CODERUPTURE2',TPResult.GetValue(GetControlText('XX_RUPTURE2')));
                         T.PutValue('RUPTURE2',RechDom(RechTablette(GetControlText('XX_RUPTURE2')),TPResult.GetValue(GetControlText('XX_RUPTURE2')),False));
                    End;
                    If (GetControlText('XX_RUPTURE3') <> '') Then
                    Begin
                         T.PutValue('CODERUPTURE3',TPResult.GetValue(GetControlText('XX_RUPTURE3')));
                         T.PutValue('RUPTURE3',RechDom(RechTablette(GetControlText('XX_RUPTURE3')),TPResult.GetValue(GetControlText('XX_RUPTURE3')),False));
                    End;

                    // Session
                    If AvecLibSession Then
                    Begin
                         T.AddChampSupValeur('PSS_DATEDEBUT',TobRuptures.Detail[r].GetValue('PSS_DATEDEBUT'));
                         T.AddChampSupValeur('PSS_DATEFIN',TobRuptures.Detail[r].GetValue('PSS_DATEFIN'));
                         T.AddChampSupValeur('PSS_LIBELLE',TobRuptures.Detail[r].GetValue('PSS_LIBELLE'));
                    End
                    Else
                    Begin
                         T.AddChampSupValeur('PSS_DATEDEBUT','');T.AddChampSupValeur('PSS_DATEFIN','');T.AddChampSupValeur('PSS_LIBELLE','');
                    End;
                 End
                 Else
                 Begin
                    // L'enregistrement existe déjà : on incrémente simplement le compteur
                    NbRep := T.GetValue('NBREPONSE');
                    T.PutValue('NBREPONSE', NbRep+1);
                 End;

                 NbRepTot := NbRepTot + 1;
                 If First Then TIndice := T; // Sauvegarde de l'indice de la Tob dans l'état
                 First := False;

                 If (GetControlText('XX_RUPTURE1') = '') Then
                      TPResult := ChercheResultats (TobResultats, Nil, IntToStr(NumQuestion), TPRep.GetValue('CC_CODE'), False)
                 Else
                      TPResult := ChercheResultats (TobResultats, TobRuptures.Detail[r], IntToStr(NumQuestion), TPRep.GetValue('CC_CODE'), False);
               End;
            End
            Else
            // Aucune réponse n'a été donnée, on la met à 0
            Begin
                    // Création de la TOB
                    T := Tob.Create('FilleEdition', TobEtat, -1);
                    T.AddChampSupValeur('POD_NUMSCORING', NumQuestion, False);
                    T.AddChampSupValeur('POD_LIBELLE', TobQuest.Detail[i].GetValue('POD_LIBELLE'), False);
                    T.AddChampSupValeur('LIBTHEME', RechDom('PGFTHEMESCORING', TobQuest.Detail[i].GetValue('POD_THEMESCORE'), False), False);  //PT5
                    T.AddChampSupValeur('POD_THEMESCORE', TobQuest.Detail[i].GetValue('POD_THEMESCORE'), False);
                    T.AddChampSupValeur('CODEREPONSE',TPRep.GetValue('CC_CODE'));
                    T.AddChampSupValeur('REPONSE', RechDom('PGFSCORING'+Tablette, TPRep.GetValue('CC_CODE'), False), False);
                    T.AddChampSupValeur('NBREPONSE', 0);
                    T.AddChampSupValeur('NBREPTOT',  0);
                    T.AddChampSupValeur('POURCENTAGE', 0);

                    // Ruptures
                    T.AddChampSupValeur('CODERUPTURE1',''); T.AddChampSupValeur('CODERUPTURE2',''); T.AddChampSupValeur('CODERUPTURE3','');
                    T.AddChampSupValeur('RUPTURE1',''); T.AddChampSupValeur('RUPTURE2',''); T.AddChampSupValeur('RUPTURE3','');
                    If (GetControlText('XX_RUPTURE1') <> '') Then
                    Begin
                         T.PutValue('CODERUPTURE1',TobRuptures.Detail[r].GetValue(GetControlText('XX_RUPTURE1')));
                         T.PutValue('RUPTURE1',RechDom(RechTablette(GetControlText('XX_RUPTURE1')),TobRuptures.Detail[r].GetValue(GetControlText('XX_RUPTURE1')),False));
                    End;
                    If (GetControlText('XX_RUPTURE2') <> '') Then
                    Begin
                         T.PutValue('CODERUPTURE2',TobRuptures.Detail[r].GetValue(GetControlText('XX_RUPTURE2')));
                         T.PutValue('RUPTURE2',RechDom(RechTablette(GetControlText('XX_RUPTURE2')),TobRuptures.Detail[r].GetValue(GetControlText('XX_RUPTURE2')),False));
                    End;
                    If (GetControlText('XX_RUPTURE3') <> '') Then
                    Begin
                         T.PutValue('CODERUPTURE3',TobRuptures.Detail[r].GetValue(GetControlText('XX_RUPTURE3')));
                         T.PutValue('RUPTURE3',RechDom(RechTablette(GetControlText('XX_RUPTURE3')),TobRuptures.Detail[r].GetValue(GetControlText('XX_RUPTURE3')),False));
                    End;

                    // Session
                    If AvecLibSession Then
                    Begin
                         T.AddChampSupValeur('PSS_DATEDEBUT',TobRuptures.Detail[r].GetValue('PSS_DATEDEBUT'));
                         T.AddChampSupValeur('PSS_DATEFIN',TobRuptures.Detail[r].GetValue('PSS_DATEFIN'));
                         T.AddChampSupValeur('PSS_LIBELLE',TobRuptures.Detail[r].GetValue('PSS_LIBELLE'));
                    End
                    Else
                    Begin
                         T.AddChampSupValeur('PSS_DATEDEBUT','');T.AddChampSupValeur('PSS_DATEFIN','');T.AddChampSupValeur('PSS_LIBELLE','');
                    End;

                    If (First = True) Then TIndice := T; // Sauvegarde de l'indice de la Tob dans l'état
             End;

             // Réponse suivante
             TPRep := TobReponses.FindNext(['CC_LIBRE'], ['SC'+Tablette], False);
             First := False;
          End;
          If TIndice <> Nil Then
          Begin
               // Calcul des totaux et des pourcentages
               T := ChercheEtat (TobEtat, TIndice, NumQuestion, True);
               While T <> Nil Do
               Begin
                    NbRep := T.GetValue('NBREPONSE');
                    if NbRepTot <> 0 then Pourcentage := (NbRep * 100) / NbRepTot else Pourcentage := 0;
                    T.PutValue('POURCENTAGE', Pourcentage);
                    T.PutValue('NBREPTOT', NbRepTot);

                    T := ChercheEtat (TobEtat, TIndice, NumQuestion, False);
               End;
          End;
       End;
     End;

     // Libération des TOBs
     FreeAndNil(TobQuest);
     FreeAndNil(TobReponses);
     FreeAndNil(TobResultats);
     FreeAndNil(TobRuptures);

     // Tri la TOB par rapport aux ruptures
     TobEtat.Detail.Sort('RUPTURE1;CODERUPTURE1;RUPTURE2;CODERUPTURE2;RUPTURE3;CODERUPTURE3;LIBTHEME;POD_NUMSCORING;REPONSE');

     // Fermeture de l'écran d'attente
     FiniMoveProgressForm;

     TFQRS1(Ecran).LaTob := TobEtat;
end;


procedure TOF_PGFOREDITSCORING.OnClose;
begin
  inherited;
  If tobEtat <> Nil then FreeAndNil(TobEtat);
end;

procedure TOF_PGFOREDITSCORING.OnArgument(Arguments: string);
var
  arg   : string;
  Edit  : THEdit;
  i     : Integer;       //PT5
  Combo : THVaLComboBox; //PT5
  CkBox : TCheckBox;     //PT5
begin
  inherited;
	Arg := ReadTokenPipe(Arguments, ';');
	if arg = 'QUESTIONNAIRE' then TFQRS1(Ecran).BValider.OnClick := EditionQuestionnaire
	else
	begin
		TFQRS1(Ecran).CodeEtat := 'PSO';
		TFQRS1(Ecran).FCodeEtat := 'PSO';
	end;
	Edit := THEdit(GetControl('PFO_SALARIE'));
	If Edit <> Nil then Edit.OnExit := ExitEdit;
	Edit := THEdit(GetControl('PFO_SALARIE1'));
	If Edit <> Nil then Edit.OnExit := ExitEdit;

	//PT5 - Début
	// Adaptation des ruptures
	For i := 1 To NB_RUPTURES do
	begin
		Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
		If Combo <> Nil Then
		begin
			// Empêcher la suppression de la rupture principale
			Combo.Items.Add('<<Aucun>>');
			Combo.Values.Add('');

			// Rendre invisible toutes les check-boxes de saut de page par défaut
			CkBox := TCheckBox(GetControl('CSAUTRUPT'+IntToStr(i)));
			CkBox.Visible := False;

			// Sur le changement de valeur d'une combo, il faut afficher le saut de page
			Combo.OnChange := OnChangeRuptures;

			// Ajout des valeurs
			Combo.Items.Add('Formation');
			Combo.Values.Add('PFO_CODESTAGE');
			Combo.Items.Add('Session');
			Combo.Values.Add('PFO_IDSESSIONFOR');
			Combo.Items.Add('Etablissement');
			Combo.Values.Add('PFO_ETABLISSEMENT');
			Combo.Items.Add('Centre de formation');
			Combo.Values.Add('PFO_CENTREFORMGU');
			Combo.Items.Add('Type (Interne/Externe)');
			Combo.Values.Add('PFO_NATUREFORM');
		end;
	end;
	//PT5 - Fin

	//PT7 - Début
	If PGBundleInscFormation Or PGBundleCatalogue Then //PT8
	Begin
		SetControlText('XX_WHEREPREDEF', ' AND (PFO_PREDEFINI="STD" OR (PFO_PREDEFINI="DOS" AND PFO_NODOSSIER="'+V_PGI.NoDossier+'"))');
		SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True));
	End;
	//PT7 - Début
end;

procedure TOF_PGFOREDITSCORING.EditionQuestionnaire(Sender: TObject);
var
  Q: TQuery;
   TobQuest, T : Tob;
  i, NumQuestion, NbRepTot, NbRep: Integer;
  Pourcentage: Double;
  Reponse: string;
  Pages: TPageControl;
begin
  Pages := TPageControl(GetControl('PAGES'));
  Q := OpenSQL('SELECT PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE FROM FORMATIONS WHERE PFO_CODESTAGE<"000007"', True);
  //        TobFormations := Tob.Create('Les formations',Nil,-1)
  //        TobFormations.LoadDetailDB('lesformations','','',Q,False);
  Ferme(Q);
  Q := OpenSQL('SELECT POD_NUMSCORING,POD_LIBELLE,POD_THEMESCORE,POD_TABLESCORE,CC_CODE,CC_LIBELLE FROM SCORINGDEF ' +
    'LEFT JOIN CHOIXCOD ON CC_TYPE="PSC" AND POD_TABLESCORE=CC_LIBRE', True);
  TobQuest := Tob.Create('LesQuestions', nil, -1);
  TobQuest.LoadDetailDB('LesQuestions', '', '', Q, False);
  Ferme(Q);
  TobEtat := Tob.Create('Edition', nil, -1);
  for i := 0 to TobQuest.detail.Count - 1 do
  begin
    NumQuestion := TobQuest.Detail[i].GetValue('POD_NUMSCORING');
    Reponse := TobQuest.Detail[i].GetValue('CC_CODE');
    T := Tob.Create('FilleEdition', TobEtat, -1);
    T.AddChampSupValeur('POD_NUMSCORING', NumQuestion, False);
    T.AddChampSupValeur('POD_LIBELLE', TobQuest.Detail[i].GetValue('POD_LIBELLE'), False);
    T.AddChampSupValeur('LIBTHEME', RechDom('PGFTEHEMSCORING', TobQuest.Detail[i].GetValue('POD_THEMESCORE'), False), False);
    T.AddChampSupValeur('QUESTION', TobQuest.Detail[i].GetValue('CC_LIBELLE'), False);
    T.AddChampSupValeur('POD_THEMESCORE', TobQuest.Detail[i].GetValue('POD_THEMESCORE'), False);
    Q := OpenSQL('SELECT COUNT (PSC_SALARIE) NBREP FROM SCORING WHERE PSC_SCFO' + IntToStr(NumQuestion) + '<>""', True);
    if not Q.Eof then NbRepTot := Q.FindField('NBREP').AsInteger
    else NbRepTot := 0;
    Ferme(Q);
    Q := OpenSQL('SELECT COUNT (PSC_SALARIE) NBREP FROM SCORING WHERE PSC_SCFO' + IntToStr(NumQuestion) + '="' + Reponse + '"', True);
    if not Q.Eof then NbRep := Q.FindField('NBREP').AsInteger
    else NbRep := 0;
    T.AddChampSupValeur('NBREPTOT', NbRepTot, False);
    T.AddChampSupValeur('NBREPONSE', NbRep, False);
    if NbRepTot <> 0 then Pourcentage := (NbRep * 100) / NbRepTot
    else Pourcentage := 0;
    T.AddChampSupValeur('POURCENTAGE', Pourcentage, False);
    Ferme(Q);
  end;
  TobQuest.Free;
  TobEtat.Detail.Sort('POD_THEMESCORE;POD_NUMSCORING');
  LanceEtatTOB('E', 'PFO', 'PSQ', TobEtat, True, False, False, Pages, '', '', False);
  TobEtat.Free;
end;

procedure TOF_PGFOREDITSCORING.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 25/05/2007 / PT5
Modifié le ... :   /  /
Description .. : Gestion des checkboxes pour les ruptures
Mots clefs ... :
*****************************************************************}
procedure TOF_PGFOREDITSCORING.OnChangeRuptures(Sender: TObject);
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
     For j:= 1 To (NB_RUPTURES-1) Do
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
     For j:=1 To NB_RUPTURES Do
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
     TCheckBox(GetControl('CSAUTRUPT'+IntToStr(i))).Visible := Affiche;
     // Si on cache, on remet également l'état à "décoché"
     If (Affiche = False) Then TCheckBox(GetControl('CSAUTRUPT'+IntToStr(i))).Checked := Affiche;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 28/05/2007  / PT5
Modifié le ... :   /  /    
Description .. : Recherche de la tablette associée à un champ
Mots clefs ... : 
*****************************************************************}
Function TOF_PGFOREDITSCORING.RechTablette (NomChamp : String) : String;
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
Auteur  ...... : FL
Créé le ...... : 31/05/2007 / PT5
Modifié le ... :   /  /
Description .. : Recherche un résultat donné dans la TOB en
Suite ........ : fonction des ruptures.
Mots clefs ... :
*****************************************************************}
Function TOF_PGFOREDITSCORING.ChercheResultats(TobResultats, TobRuptures: TOB; NumQuestion : String; Valeur: Variant; Premier: Boolean): TOB;
Var
     TRes : TOB;
     Fonction : TFunctionPointer;
begin
     TRes := Nil;

     If Premier Then Fonction := TobResultats.FindFirst
     Else Fonction := TobResultats.FindNext;

     If (GetControlText('XX_RUPTURE1') = '') Then
          TRes := Fonction(['PSC_SCFO'+NumQuestion],[Valeur], False)
     Else If (GetControlText('XX_RUPTURE3') <> '') Then
          TRes := Fonction([GetControlText('XX_RUPTURE1'),GetControlText('XX_RUPTURE2'),GetControlText('XX_RUPTURE3'),'PSC_SCFO'+NumQuestion],
                                             [TobRuptures.GetValue(GetControlText('XX_RUPTURE1')),
                                              TobRuptures.GetValue(GetControlText('XX_RUPTURE2')),
                                              TobRuptures.GetValue(GetControlText('XX_RUPTURE3')),
                                              Valeur], False)
     Else If (GetControlText('XX_RUPTURE2') <> '') Then
          TRes := Fonction([GetControlText('XX_RUPTURE1'),GetControlText('XX_RUPTURE2'),'PSC_SCFO'+NumQuestion],
                                             [TobRuptures.GetValue(GetControlText('XX_RUPTURE1')),
                                              TobRuptures.GetValue(GetControlText('XX_RUPTURE2')),
                                              Valeur], False)
     Else If (GetControlText('XX_RUPTURE1') <> '') Then
          TRes := Fonction([GetControlText('XX_RUPTURE1'),'PSC_SCFO'+NumQuestion],
                                             [TobRuptures.GetValue(GetControlText('XX_RUPTURE1')),
                                              Valeur], False);

     Result := TRes;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 31/05/2007 / PT5
Modifié le ... :   /  /
Description .. : Recherche un résultat donné dans la TOB d'édition en
Suite ........ : fonction des ruptures.
Mots clefs ... :
*****************************************************************}
Function TOF_PGFOREDITSCORING.ChercheEtat(TobEtat, TobValeurs : TOB; NumQuestion : Integer; Premier: Boolean): TOB;
Var
     TRes : TOB;
     Fonction : TFunctionPointer;
begin
     TRes := Nil;

     If Premier Then Fonction := TobEtat.FindFirst
     Else Fonction := TobEtat.FindNext;


     If (GetControlText('XX_RUPTURE1') = '') Then
              TRes := Fonction(['POD_NUMSCORING'], [NumQuestion], False)
     Else If (GetControlText('XX_RUPTURE3') <> '') Then
              TRes := Fonction(['POD_NUMSCORING','CODERUPTURE1','CODERUPTURE2','CODERUPTURE3'], [NumQuestion,
                                     TobValeurs.GetValue('CODERUPTURE1'),
                                     TobValeurs.GetValue('CODERUPTURE2'),
                                     TobValeurs.GetValue('CODERUPTURE3')], False)
     Else If (GetControlText('XX_RUPTURE2') <> '') Then
              TRes := Fonction(['POD_NUMSCORING','CODERUPTURE1','CODERUPTURE2'], [NumQuestion,
                                     TobValeurs.GetValue('CODERUPTURE1'),
                                     TobValeurs.GetValue('CODERUPTURE2')], False)
     Else If (GetControlText('XX_RUPTURE1') <> '') Then
              TRes := Fonction(['POD_NUMSCORING','CODERUPTURE1'], [NumQuestion,
                                     TobValeurs.GetValue('CODERUPTURE1')], False);

     Result := TRes;
end;


{TOF_PGCUBESCORING}
procedure TOF_PGCUBESCORING.OnArgument(Arguments: string);
var Num : Integer;
    DD,DF : TdateTime;
    Edit : THEdit;
begin
  inherited;

	Edit := THEdit(GetControl('PFO_SALARIE'));
	If Edit <> Nil then Edit.OnExit := ExitEdit;
	Edit := THEdit(GetControl('PFO_SALARIE_'));
	If Edit <> Nil then Edit.OnExit := ExitEdit;

	For Num := 1 to VH_Paie.PGNbreStatOrg do
	begin
		if Num >4 then Break;
		VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
	end;
	VisibiliteStat (GetControl ('PPFO_CODESTAT'),GetControl ('TPPFO_CODESTAT')) ;
	For Num  :=  1 to VH_Paie.NBFormationLibre do
	begin
		if Num >8 then Break;
		VisibiliteChampFormation (IntToStr(Num),GetControl ('PFO_FORMATION'+IntToStr(Num)),GetControl ('TPFO_FORMATION'+IntToStr(Num)));
	end;
	RendMillesimeRealise(DD,DF);
	
	SetControlText('PFO_DATEDEBUT',DateToStr(DD));
	SetControlText('PFO_DATEDEBUT_',DateToStr(DF));
	
	//PT7 - Début
	If PGBundleInscFormation Or PGBundleCatalogue Then  //PT8
	Begin
		SetControlText('XX_WHEREPREDEF', DossiersAInterroger('',V_PGI.NoDossier,'PFO',True,False));
		SetControlProperty ('PSC_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True));
	End;
	//PT7 - Début
end;

procedure TOF_PGCUBESCORING.ExitEdit(Sender: TObject);
var edit : thedit;
begin
	edit:=THEdit(Sender);
	if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

initialization
  registerclasses([TOF_PGFOREDITSCORING,TOF_PGCUBESCORING]);
end.

