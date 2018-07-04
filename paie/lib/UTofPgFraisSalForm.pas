{***********UNITE*************************************************
Auteur  ...... JL
Créé le ...... : 02/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : FRAISSALFORM ()
Mots clefs ... : TOF;PGFRAISSALFORM
*****************************************************************
PT1  | 10/05/2001 | V_50  | JL | Correction erreur indice hors des limites si Nb frais ds paramsoc > nb libellé saisies dans la tablette
     |            |       |    | + correction pour affichage des colonnes numériques de la grille
---  | 20/03/2006 |       | JL | modification clé annuaire ----
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2  | 17/05/2007 | V_720 | FL | FQ 11532 Adaptation des plafonds de frais en fonction de la population
PT3  | 07/06/2007 | V_720 | FL | Gestion des multi-sessions
PT4  | 19/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT5  | 02/04/2008 | V_803 | FL | Adaptation partage formation
PT6  | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT9  | 04/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
}
Unit UTofPgFraisSalForm;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,Utob,HSysMenu,HTB97,PgOutilsFormation,PGPopulOutils,PGOutils2 ;

Type
  TOF_PGFRAISSALFORM = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    Grille, GrilleAnim : THGrid;
    LeStage,LaSession,LeMillesime,OrgCollecteur,LeLieu,MillesimeEC : String;
    TabFrais : array  [1..15] of String;
    COL_START_FRAIS,COL_START_FRAIS_ANIM : Integer; //PT5
    procedure RemplirGrille;
    procedure MajFrais(Sender : TObject);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);

  end ;

Implementation

Uses UtilPGI,vierge;

procedure TOF_PGFRAISSALFORM.OnArgument (S : String ) ;
VAR TypeSaisie : String;
    BValid : TToolBarButton97;
    Q : TQuery;
    i : Integer;
    DD,DF : TDateTime;
    Dossier : String; //PT5
begin
  Inherited ;
        If PGBundleHierarchie then
        Begin
            COL_START_FRAIS := 5;
            COL_START_FRAIS_ANIM := 4;
        end
        else
        Begin
            COL_START_FRAIS := 4;
            COL_START_FRAIS_ANIM := 3;
        End;
        
        TypeSaisie := Trim(ReadTokenPipe(S,';'));
        LeStage := Trim(ReadTokenPipe(S,';'));
        LaSession := Trim(ReadTokenPipe(S,';'));
        LeMillesime := Trim(ReadTokenPipe(S,';'));
        MillesimeEC := RendMillesimeRealise(DD,DF);
        If PGBundleInscFormation Then Dossier := Trim(ReadTokenPipe(S,';')); //PT5
        If Dossier <> '' Then Dossier := ' AND PSS_NODOSSIER="'+Dossier+'"';
        
        Q := OpenSQL('SELECT PSS_ORGCOLLECTSGU,PSS_LIEUFORM,PSS_DATEDEBUT,PSS_DATEFIN,PSS_LIBELLE,PST_LIBELLE'+
        ' FROM SESSIONSTAGE LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME'+
        ' WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE='+LaSession+' AND PSS_MILLESIME="'+LeMillesime+'"',True);  //DB2
        if not Q.eof then
        begin
                SetControlText('STAGE',Q.FindField('PST_LIBELLE').AsString);
                SetControlText('SESSION',Q.FindField('PSS_LIBELLE').AsString);
                SetControlText('DATEDEBUT',DateToStr(Q.FindField('PSS_DATEDEBUT').AsDateTime));
                SetControlText('DATEFIN',DateToStr(Q.FindField('PSS_DATEFIN').AsDateTime));
                LeLieu := Q.FindField('PSS_LIEUFORM').AsString;
                OrgCollecteur := Q.FindField('PSS_ORGCOLLECTSGU').AsString;
                SetControlText('LIEUFORM',RechDom('PGLIEUFORMATION',leLieu,False));
                SetControlText('ORGCOLLECT',RechDom('PGORGCOLLECTEUR',OrgCollecteur,False));
        end;
        Ferme(Q);
        
        Grille := THGrid(GetControl('GFRAIS'));
        If Grille <> Nil then Grille.OnCellExit := OnCellExit;
        GrilleAnim := THGrid(GetControl('GFRAISANIM'));
        
        //PT3 - Début
        Grille.ColWidths     [0]:=60; // Matricule
        Grille.ColWidths     [1]:=90; // Nom, prénom
        Grille.ColWidths     [2]:=0;  // Colonne cachée pour le code plan
        Grille.ColWidths     [3]:=80; // Libellé du type de plan
        If PGBundleHierarchie then  Grille.ColWidths [4] :=80; // Dossier
        
        GrilleAnim.ColWidths [0]:=60; // Matricule
        GrilleAnim.ColWidths [1]:=70; // Nom
        GrilleAnim.ColWidths [2]:=80; // Prénom
		If PGBundleHierarchie then GrilleAnim.ColWidths [3]:=80; // Dossier
		
        //GrilleAnim.ColWidths [3] := 40; //PT5
        //GrilleAnim.ColFormats[3] := '# ##0.00';

        //PT3 - Fin
        For i:=COL_START_FRAIS to Grille.ColCount-1 do
        begin
                Grille.ColWidths [i] := 40;
                Grille.ColFormats[i] := '# ##0.00';
        end;
        
        //PT5 - Début
        For i:=COL_START_FRAIS_ANIM to GrilleAnim.ColCount-1 do
        begin
                GrilleAnim.ColWidths [i] := 40;
                GrilleAnim.ColFormats[i] := '# ##0.00';
        end;
        //PT5 - Fin
        
        SetControlChecked('MONTANT',True);
        SetControlChecked('QUANTITE',True);
        SetControlChecked('PLAFOND',True);
        RemplirGrille;
        BValid := TToolBarButton97(GetControl('BVAL'));
        If BValid <> Nil then BValid.OnClick := MajFrais;
end ;

Procedure TOF_PGFRAISSALFORM.RemplirGrille;
var Q : TQuery;
    i,j,a : Integer;
    LeSalarie,LeFrais,LePlan,SQL : String;
    TobAnim,TobSal,TobFrais,TobLibFrais : Tob;
    Titres,TitresAnim : HTStringList; //PT3
    NbAnim,LaColonne : Integer;
    //var HMTRAD : THSystemMenu;
    Requete : String; //PT5
begin
        Grille.ColCount := (VH_Paie.PGFNbFraisLibre*2) + COL_START_FRAIS;
        GrilleAnim.ColCount := (VH_Paie.PGFNbFraisLibre*2) + COL_START_FRAIS_ANIM;
        // Titres salariés
        Titres := HTStringList.Create;
        Titres.Add('Matricule');
        Titres.Add('Nom, Prénom');
        Titres.Add('Code plan');
        Titres.Add('Type de plan');
        If PGBundleHierarchie then Titres.Add('Dossier');
        //PT3 - Début
        // Titres animateurs
        TitresAnim := HTStringList.Create;
        TitresAnim.Add('Matricule');
        TitresAnim.Add('Nom');
        TitresAnim.Add('Prénom');
        If PGBundleHierarchie then TitresAnim.Add('Dossier');
        
        //PT5 - Début
        If PGBundleHierarchie then 
        Begin
			Grille.FixedCols     := Grille.FixedCols + 1;
        	GrilleAnim.FixedCols := GrilleAnim.FixedCols + 1;
        End;
        //PT5 - Fin
        
        //PT3 - Fin
        // Titres des colonnes pour les frais
        //Q := OpenSQL('SELECT CC_LIBELLE,CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PFA"',True);
        If PGBundleInscFormation Then
        	Requete := 'SELECT CC_CODE,CC_LIBELLE FROM '+GetBase(GetBasePartage(BUNDLE_FORMATION),'CHOIXCOD')+' WHERE CC_TYPE="PFA" ORDER BY CC_CODE' //PT5
        Else
        	Requete := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PFA" ORDER BY CC_CODE';
        Q := OpenSQL(Requete, True);
        TobLibFrais := Tob.Create('Les libellés',Nil,-1);
        TobLibFrais.LoadDetailDB('tablette','','',Q,False);
        Ferme(Q);
        
        For i := 0 to VH_Paie.PGFNbFraisLibre - 1 do
        begin
                If i < TobLibFrais.Detail.Count then   // PT1
                begin
                        TabFrais[i+1] := TobLibFrais.Detail[i].GetValue('CC_CODE');
                        if TobLibFrais.Detail[i] <> Nil Then
                        Begin
                              Titres.Add('Qté');
                              TitresAnim.Add('Qté');
                        //End;
                        //if TobLibFrais.Detail[i] <> Nil Then
                        //Begin
                              Titres.Add(TobLibFrais.Detail[i].GetValue('CC_LIBELLE'));
                              TitresAnim.Add(TobLibFrais.Detail[i].GetValue('CC_LIBELLE'));
                        End;
                End
                //PT3 - Début
                Else
                Begin
                         // Si pas de libellé, on enlève la colonne
                         Grille.ColCount := Grille.ColCount - 2;
                         GrilleAnim.ColCount := GrilleAnim.ColCount - 2;
                //PT3 - Fin
                end;
        end;
        TobLibFrais.Free;

        // Mise à jour de la grille avec les titres
        Grille.Titres := Titres;
        GrilleAnim.Titres := TitresAnim;  //PT3
        Titres.Free;
        TitresAnim.Free;
        for i := COL_START_FRAIS_ANIM to GrilleAnim.ColCount - 1 do  // PT1
        begin
                GrilleAnim.ColAligns[i]:=taRightJustify;
        end;
        for i := COL_START_FRAIS to Grille.ColCount - 1 do  // PT1
        begin
                Grille.ColAligns[i]:=taRightJustify;
        end;

        // Renseignement des matricules pour les animateurs
        If PGBundleHierarchie then 
        	SQL := 'SELECT PAN_SALARIE,PSI_LIBELLE,PSI_PRENOM,PSI_NODOSSIER FROM SESSIONANIMAT '+
        	'LEFT JOIN INTERIMAIRES ON PAN_SALARIE=PSI_INTERIMAIRE '+
        	'WHERE PAN_SALARIE<>"" AND '+
        	'PAN_MILLESIME="'+LeMillesime+'" AND PAN_CODESTAGE="'+LeStage+'" AND PAN_ORDRE='+LaSession+
	        //PT6
	        // N'affiche pas les salariés confidentiels
	    	' AND (PAN_SALARIE="" OR PAN_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))'  //PT9
        else 
        	SQL := 'SELECT PAN_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SESSIONANIMAT '+
        	'LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE '+
        	'WHERE PAN_SALARIE<>"" AND '+
        	'PAN_MILLESIME="'+LeMillesime+'" AND PAN_CODESTAGE="'+LeStage+'" AND PAN_ORDRE='+LaSession+
	        //PT6
	        // N'affiche pas les salariés confidentiels
	    	' AND (PAN_SALARIE="" OR PAN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))';  //PT9
        Q := OpenSQL(SQL,True);  //DB2
        TobAnim := Tob.Create('Les formateurs',Nil,-1);
        TobAnim.LoadDetailDB('Les animamteurs','','',Q,False);
        Ferme(Q);
        
        NbAnim := TobAnim.Detail.Count;
        //PT3 - Début
        If NbAnim = 0 Then
        Begin
               GrilleAnim.RowCount := 1 + 1; // Forçage du nombre de lignes à 2 pour en avoir une vide en dessous du titre
                                             // afin de corriger un problème de visualisation avec l'AGL
               grilleAnim.CellValues[0,2] := '';
               grilleAnim.CellValues[1,2] := '';
               grilleAnim.CellValues[2,2] := '';
               If PGBundleHierarchie then grilleAnim.CellValues[3,2] := '';
               grilleAnim.Enabled := False;
        End
        Else
        //PT3 - Fin
               GrilleAnim.RowCount := 1 + NbAnim;
               
        For i := 0 to TobAnim.Detail.Count -1 do
        begin
          If PGBundleHierarchie then
          begin
            grilleAnim.CellValues[0,i+1] := TobAnim.detail[i].GetValue('PAN_SALARIE');
            grilleAnim.CellValues[1,i+1] := TobAnim.detail[i].GetValue('PSI_LIBELLE');
            grilleAnim.CellValues[2,i+1] := TobAnim.detail[i].GetValue('PSI_PRENOM');
            grilleAnim.CellValues[3,i+1] := TobAnim.detail[i].GetValue('PSI_NODOSSIER');
          end
          else
          begin
            grilleAnim.CellValues[0,i+1] := TobAnim.detail[i].GetValue('PAN_SALARIE');
            grilleAnim.CellValues[1,i+1] := TobAnim.detail[i].GetValue('PSA_LIBELLE');
            grilleAnim.CellValues[2,i+1] := TobAnim.detail[i].GetValue('PSA_PRENOM');
          end;
        end;
        TobAnim.Free;

        // Renseignement des matricules pour les salariés
        If PGBundleHierarchie then 
        	SQL := 'SELECT PFO_SALARIE,PSI_LIBELLE,PSI_PRENOM,PFO_TYPEPLANPREV,PSI_NODOSSIER FROM FORMATIONS LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFO_SALARIE WHERE PFO_MILLESIME="'+LeMillesime+'" AND'+ //PT5
        	' PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE='+LaSession+
	        //PT6
	        // N'affiche pas les salariés confidentiels
	    	' AND PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
        Else 
        	SQL := 'SELECT PFO_SALARIE,PSA_LIBELLE,PSA_PRENOM,PFO_TYPEPLANPREV FROM FORMATIONS LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE WHERE PFO_MILLESIME="'+LeMillesime+'" AND'+ //PT5
        	' PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE='+LaSession+
	        //PT6
	        // N'affiche pas les salariés confidentiels
	    	' AND PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
        Q := OpenSQL(SQL,True); //DB2
        TobSal := Tob.Create('Les Salaries',Nil,-1);
        TobSal.LoadDetailDB('table FORMATIONS','','',Q,False);
        Ferme(Q);
        
        Grille.RowCount := 1 + TobSal.Detail.Count;
        For i := 0 to TobSal.Detail.Count -1 do
        begin
          If PGBundleHierarchie then
          begin
            grille.CellValues[0,i+1] := TobSal.detail[i].GetValue('PFO_SALARIE');
            grille.CellValues[1,i+1] := TobSal.detail[i].GetValue('PSI_LIBELLE') + ' ' + TobSal.detail[i].GetValue('PSI_PRENOM');
            grille.CellValues[2,i+1] := TobSal.detail[i].GetValue('PFO_TYPEPLANPREV');
            grille.CellValues[3,i+1] := RechDom('PGTYPEPLANPREV',TobSal.detail[i].GetValue('PFO_TYPEPLANPREV'),False);
            grille.CellValues[4,i+1] := TobSal.detail[i].GetValue('PSI_NODOSSIER');
          end
          else
          begin
            grille.CellValues[0,i+1] := TobSal.detail[i].GetValue('PFO_SALARIE');
            grille.CellValues[1,i+1] := TobSal.detail[i].GetValue('PSA_LIBELLE') + ' ' + TobSal.detail[i].GetValue('PSA_PRENOM');
            grille.CellValues[2,i+1] := TobSal.detail[i].GetValue('PFO_TYPEPLANPREV');
            grille.CellValues[3,i+1] := RechDom('PGTYPEPLANPREV',TobSal.detail[i].GetValue('PFO_TYPEPLANPREV'),False);
          end;
        end;
        TobSal.Free;

        // Renseignement des valeurs pour les animateurs
        For i := 1 to GrilleAnim.RowCount -1 do
        begin
                LeSalarie := GrilleAnim.CellValues[0,i];
                Q := OpenSQL('SELECT PFS_MONTANT,PFS_QUANTITE,PFS_FRAISSALFOR FROM FRAISSALFORM WHERE PFS_SALARIE="'+LeSalarie+'" AND '+
                'PFS_CODESTAGE="'+LeStage+'" AND PFS_ORDRE='+LaSession+' AND PFS_MILLESIME="'+LeMillesime+'"',True); //DB2
                TobFrais := Tob.Create('Les frais',Nil,-1);
                TobFrais.LoadDetailDB('Les frais','','',Q,False);
                Ferme(Q);
                
                For j := 0 To TobFrais.Detail.Count -1 do
                begin
                        LeFrais := TobFrais.Detail[j].GetValue('PFS_FRAISSALFOR');
                        LaColonne := 0;
                        For a := 1 to VH_Paie.PGFNbFraisLibre do
                        begin
                                If TabFrais[a] = LeFrais then
                                begin
                                        LaColonne := (a-1)*2 + COL_START_FRAIS_ANIM + 1;
                                        continue;
                                end;
                        end;
                        If LaColonne > 0 then
                        begin
                                GrilleAnim.CellValues[LaColonne,i] := TobFrais.detail[j].GetValue('PFS_MONTANT');
                                GrilleAnim.CellValues[LaColonne-1,i] := TobFrais.detail[j].GetValue('PFS_QUANTITE');
                        end;
                end;
                TobFrais.Free;
        end;

        // Renseignement des valeurs pour les salariés
        For i := 1 to Grille.RowCount -1 do
        begin
                LeSalarie := Grille.CellValues[0,i];
                LePlan    := Grille.CellValues[2,i]; //PT3
                Q := OpenSQL('SELECT PFS_MONTANT,PFS_QUANTITE,PFS_FRAISSALFOR FROM FRAISSALFORM WHERE PFS_SALARIE="'+LeSalarie+'" AND '+
                'PFS_CODESTAGE="'+LeStage+'" AND PFS_ORDRE='+LaSession+' AND PFS_MILLESIME="'+LeMillesime+'" AND PFS_TYPEPLANPREV="'+LePlan+'"',True);  //DB2 //PT3
                TobFrais := Tob.Create('Les frais',Nil,-1);
                TobFrais.LoadDetailDB('Les frais','','',Q,False);
                Ferme(Q);
                
                For j := 0 To TobFrais.Detail.Count -1 do
                begin
                        LeFrais := TobFrais.Detail[j].GetValue('PFS_FRAISSALFOR');
                        LaColonne := 0;
                        For a := 1 to VH_Paie.PGFNbFraisLibre do
                        begin
                                If TabFrais[a] = LeFrais then
                                begin
                                        LaColonne := (a-1)*2 + COL_START_FRAIS + 1;
                                        continue;
                                end;
                        end;
                        If LaColonne > 0 then
                        begin
                                Grille.CellValues[LaColonne,i] := TobFrais.detail[j].GetValue('PFS_MONTANT');
                                Grille.CellValues[LaColonne-1,i] := TobFrais.detail[j].GetValue('PFS_QUANTITE');
                        end;
                end;
                TobFrais.Free;
        end;
        //PT3 - Début
{        For i := 0 to 2 do
        begin
                Grille.ColWidths [i]:=60;
                GrilleAnim.ColWidths [i]:=60;
        end;
        For i:=3 to Grille.ColCount-1 do
        begin
                Grille.ColWidths [i]:=50;
                GrilleAnim.ColWidths [i]:=50;
        end;
}
        //PT3 - Fin
        TFVierge(Ecran).HMTrad.ResizeGridColumns(Grille) ;
        TFVierge(Ecran).HMTrad.ResizeGridColumns(GrilleAnim) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : ??/??/????
Modifié le ... : 07/06/2007 / PT3
Description .. : Mise à jour des frais en base.
Suite ........ : En cas de multi-session, mise à jour également des
Suite ........ : montants de la session entête
Mots clefs ... : FRAIS;MULTI SESSIONS
*****************************************************************}
procedure TOF_PGFRAISSALFORM.MajFrais(Sender : TObject);
var i,j : Integer;
    LeSalarie,LeFrais,StQte,StMontant : String;
    Total,TotalPlaf,Montant,Qte,Plafond,SalaireANim,FraisAnim : Double;
    TobFrais,TobFormations,TF,TFF : Tob;
    Q : TQuery;
begin
        // Liste des frais
        Q := OpenSQL('SELECT * FROM FRAISSALFORM WHERE '+
        'PFS_CODESTAGE="'+LeStage+'" AND PFS_ORDRE='+LaSession+' AND PFS_MILLESIME="'+LeMillesime+'"',True);  //DB2
        TobFrais := Tob.Create('FRAISSALFORM',Nil,-1);
        TobFrais.LoadDetailDB('FRAISSALFORM','','',Q,False);
        Ferme(Q);

        //PT3 - Début
        // Liste des informations de formation
        TobFormations := Tob.Create('LesFormations', Nil, -1);
        TobFormations.LoadDetailFromSQL('SELECT PFO_SALARIE,PFO_TYPEPLANPREV,PFO_IDSESSIONFOR,PFO_NOSESSIONMULTI,PFO_PGTYPESESSION ' +
                                        'FROM FORMATIONS ' +
                                        'WHERE PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE="'+LaSession+'" AND PFO_MILLESIME="'+LeMillesime+'"');
        //PT3 - Fin

        // Contrôle des montants saisis
        For i := 1 To GrilleAnim.RowCount - 1 do
        begin
                For j := 1 to VH_Paie.PGFNbFraisLibre do
                begin
                        StQte := GrilleAnim.CellValues[COL_START_FRAIS_ANIM+(j-1)*2,i];
                        StMontant := GrilleAnim.CellValues[COL_START_FRAIS_ANIM+1+(j-1)*2,i];
                        If (( StQte = '') or (StQte = '0')) and ((StMontant <> '') and (StMontant <> '0')) then
                        begin
                                LeSalarie := GrilleAnim.CellValues[0,i];
                                LeFrais := TabFrais[j];
                                GrilleAnim.Col := COL_START_FRAIS_ANIM+(j-1)*2;
                                GrilleAnim.Row := i;
                                PGIBox('La quantité du frais '+RechDom('PGFRAISSALFORM',LeFrais,false)+' n''est pas renseigné pour l''animateur '+RechDom('PGSALARIE',LeSalarie,False),Ecran.Caption);
                                Exit;
                        end;
                        If ((StQte <> '') and (StQte <> '0')) and ((StMontant = '') or (StMontant = '0'))  then
                        begin
                                LeSalarie := GrilleAnim.CellValues[0,i];
                                LeFrais := TabFrais[j];
                                GrilleAnim.Col := COL_START_FRAIS_ANIM+1+(j-1)*2;
                                GrilleAnim.Row := i;
                                PGIBox('Le montant du frais '+RechDom('PGFRAISSALFORM',LeFrais,false)+' n''est pas renseigné pour l''animateur '+RechDom('PGSALARIE',LeSalarie,False),Ecran.Caption);
                                Exit;
                        end;
                end;
        end;
        For i := 1 To Grille.RowCount - 1 do
        begin
                For j := 1 to VH_Paie.PGFNbFraisLibre do
                begin
                        StQte := Grille.CellValues[COL_START_FRAIS+(j-1)*2,i];
                        StMontant := Grille.CellValues[COL_START_FRAIS+1+(j-1)*2,i];
                        If ((StQte = '') or (StQte = '0')) and ((StMontant <> '') and (StMontant <> '0')) then
                        begin
                                LeSalarie := Grille.CellValues[0,i];
                                LeFrais := TabFrais[j];
                                Grille.Col := COL_START_FRAIS+(j-1)*2;
                                Grille.Row := i;
                                PGIBox('La quantité du frais '+RechDom('PGFRAISSALFORM',LeFrais,false)+' n''est pas renseigné pour le stagiaire '+RechDom('PGSALARIE',LeSalarie,False),Ecran.Caption);
                                Exit;
                        end;
                        If ((StQte <> '0') and (StQte <> '')) and ((StMontant = '') or (StMontant = '0')) then
                        begin
                                LeSalarie := Grille.CellValues[0,i];
                                LeFrais := TabFrais[j];
                                Grille.Col := COL_START_FRAIS+1+(j-1)*2;
                                Grille.Row := i;
                                PGIBox('Le montant du frais '+RechDom('PGFRAISSALFORM',LeFrais,false)+' n''est pas renseigné pour le stagiaire '+RechDom('PGSALARIE',LeSalarie,False),Ecran.Caption);
                                Exit;
                        end;
                end;
        end;

        // Mise à jour des données pour les animateurs
        For i := 1 To GrilleAnim.RowCount - 1 do
        begin
                LeSalarie := GrilleAnim.CellValues[0,i];
                If LeSalarie = '' then continue;
                TF := Nil;
                For j := 1 to VH_Paie.PGFNbFraisLibre do
                begin
                        LeFrais := TabFrais[j];
                        If LeFrais = '' Then Break; //PT3
                        TF := TobFrais.FindFirst(['PFS_SALARIE','PFS_FRAISSALFOR'],[LeSalarie,LeFrais],False);
                        if TF <> Nil then
                        begin
                                If (GrilleAnim.CellValues[COL_START_FRAIS_ANIM+1+(j-1)*2,i] <> '') and (GrilleAnim.CellValues[COL_START_FRAIS_ANIM+1+(j-1)*2,i] <> '0') then
                                begin
                                        Qte := Arrondi(StrToFloat(GrilleAnim.CellValues[COL_START_FRAIS_ANIM+(j-1)*2,i]),2);
                                        Montant := Arrondi(StrToFloat(GrilleAnim.CellValues[COL_START_FRAIS_ANIM+1+(j-1)*2,i]),2);
                                        TF.PutValue('PFS_QUANTITE',Qte);
                                        TF.PutValue('PFS_MONTANT',Montant);
                                        Plafond := RendPlafondFraisForm(LeFrais,LeLieu,OrgCollecteur,MillesimeEC,DonnePopulation(LeSalarie)); //PT2
                                        Plafond := Arrondi(Plafond*Qte,2);
                                        If (Montant>Plafond) and (plafond>0) Then TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Plafond))
                                        Else TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Montant));
                                        If (VH_Paie.PGForGestPlafByPop) Then TF.PutValue('PFS_POPULATION', DonnePopulation(LeSalarie)); //PT2
                                        TF.UpdateDB(False);
                                end
                                Else
                                begin
                                        TF.DeleteDB(False);
                                end;
                        end
                        Else
                        begin
                                If (GrilleAnim.CellValues[COL_START_FRAIS_ANIM+1+(j-1)*2,i] <> '') and (GrilleAnim.CellValues[COL_START_FRAIS_ANIM+1+(j-1)*2,i] <> '0') then
                                begin
                                        TF := Tob.Create('FRAISSALFORM',TobFrais,-1);
                                        TF.PutValue('PFS_CODESTAGE',LeStage);
                                        TF.PutValue('PFS_ORDRE',StrToInt(LaSession));
                                        TF.PutValue('PFS_MILLESIME',LeMillesime);
                                        TF.PutValue('PFS_FRAISSALFOR',LeFrais);
                                        TF.PutValue('PFS_SALARIE',LeSalarie);
                                        Montant := Arrondi(StrToFloat(GrilleAnim.CellValues[COL_START_FRAIS_ANIM+1+(j-1)*2,i]),2);
                                        TF.PutValue('PFS_MONTANT',Montant);
                                        If ISNumeric (GrilleAnim.CellValues[COL_START_FRAIS_ANIM+(j-1)*2,i]) then Qte := Arrondi(StrToFloat(GrilleAnim.CellValues[COL_START_FRAIS_ANIM+(j-1)*2,i]),2)
                                        else Qte := 0;
                                        TF.PutValue('PFS_QUANTITE',Qte);
                                        TF.PutValue('PFS_CTPEDAG','X');
                                        TF.PutValue('PFS_CTSALARIE','-');
                                        Plafond := RendPlafondFraisForm(LeFrais,LeLieu,OrgCollecteur,MillesimeEC,DonnePopulation(LeSalarie));//PT2
                                        Plafond := Arrondi(Plafond*Qte,2);
                                        If (Montant>Plafond) and (plafond>0) Then TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Plafond))
                                        Else TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Montant));
                                        //PT3 - Début
                                        TF.PutValue('PFS_PREDEFINI', GetPredefiniPopulation(TYP_POPUL_FORM_PREV));
                                        TF.PutValue('PFS_NODOSSIER', PgrendNodossier());
                                        TF.PutValue('PFS_CODEPOP','');
                                        If (VH_Paie.PGForGestPlafByPop) Then TF.PutValue('PFS_POPULATION', DonnePopulation(LeSalarie)); //PT2

                                        TFF := TobFormations.FindFirst(['PFO_SALARIE'],[LeSalarie], False);
                                        If TFF <> Nil Then
                                        Begin
                                             TF.PutValue('PFS_PGTYPESESSION',TFF.GetValue('PFO_PGTYPESESSION'));
                                             TF.PutValue('PFS_NOSESSIONMULTI',TFF.GetValue('PFO_NOSESSIONMULTI'));
                                             TF.PutValue('PFS_IDSESSIONFOR',TFF.GetValue('PFO_IDSESSIONFOR'));
                                        End
                                        Else
                                        Begin
                                             TF.PutValue('PFS_PGTYPESESSION','AUC');
                                             TF.PutValue('PFS_NOSESSIONMULTI',-1);
                                             TF.PutValue('PFS_IDSESSIONFOR','');
                                        End;
                                        TF.PutValue('PFS_TYPEPLANPREV','');
                                        //PT3 - Fin

                                        TF.InsertDB(Nil,False);
                                end;
                        end;
                end;

                //PT3 - Début
                // Gestion des multi-sessions : Mise à jour des frais totaux de la session en-tête
                If (TF <> Nil) And (TF.GetValue('PFS_PGTYPESESSION') = TYP_SOUSSESSION) Then MajSousSessionSal(LeSalarie,LeStage,TF.GetValue('PFS_NOMULTISESSION'));
                //PT3 - Fin
        end;

        // Mise à jour des données pour les salariés
        For i := 1 To Grille.RowCount - 1 do
        begin
                LeSalarie := Grille.CellValues[0,i];
                If LeSalarie = '' then continue;
                TF := Nil;
                For j := 1 to VH_Paie.PGFNbFraisLibre do
                begin
                        LeFrais := TabFrais[j];
                        If LeFrais = '' Then Break; //PT3
                        TF := TobFrais.FindFirst(['PFS_SALARIE','PFS_FRAISSALFOR','PFS_TYPEPLANPREV'],[LeSalarie,LeFrais,Grille.CellValues[2,i]],False); //PT3
                        if TF <> Nil then
                        begin
                                If (Grille.CellValues[COL_START_FRAIS+1+(j-1)*2,i] <> '') and (Grille.CellValues[COL_START_FRAIS+1+(j-1)*2,i] <> '0') then
                                begin
                                        Qte := Arrondi(StrToFloat(Grille.CellValues[COL_START_FRAIS+(j-1)*2,i]),2);
                                        Montant := Arrondi(StrToFloat(Grille.CellValues[COL_START_FRAIS+1+(j-1)*2,i]),2);
                                        TF.PutValue('PFS_QUANTITE',Qte);
                                        TF.PutValue('PFS_MONTANT',Montant);
                                        Plafond := RendPlafondFraisForm(LeFrais,LeLieu,OrgCollecteur,MillesimeEC,DonnePopulation(LeSalarie)); //PT2
                                        Plafond := Arrondi(Plafond*Qte,2);
                                        If (Montant>Plafond) and (plafond>0) Then TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Plafond))
                                        Else TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Montant));
                                        If (VH_Paie.PGForGestPlafByPop) Then TF.PutValue('PFS_POPULATION', DonnePopulation(LeSalarie)); //PT2
                                        TF.UpdateDB(False);
                                end
                                Else
                                begin
                                        TF.DeleteDB(False);
                                end;
                        end
                        Else
                        begin
                                If (Grille.CellValues[COL_START_FRAIS+1+(j-1)*2,i] <> '') and (Grille.CellValues[COL_START_FRAIS+1+(j-1)*2,i] <> '0') then
                                begin
                                        TF := Tob.Create('FRAISSALFORM',TobFrais,-1);
                                        TF.PutValue('PFS_CODESTAGE',LeStage);
                                        TF.PutValue('PFS_ORDRE',StrToInt(LaSession));
                                        TF.PutValue('PFS_MILLESIME',LeMillesime);
                                        TF.PutValue('PFS_FRAISSALFOR',LeFrais);
                                        TF.PutValue('PFS_SALARIE',LeSalarie);
                                        Montant := Arrondi(StrToFloat(Grille.CellValues[COL_START_FRAIS+1+(j-1)*2,i]),2);
                                        TF.PutValue('PFS_MONTANT',Montant);
                                        If ISNumeric (Grille.CellValues[COL_START_FRAIS+(j-1)*2,i]) then Qte := Arrondi(StrToFloat(Grille.CellValues[COL_START_FRAIS+(j-1)*2,i]),2)
                                        else Qte := 0;
                                        TF.PutValue('PFS_QUANTITE',Qte);
                                        TF.PutValue('PFS_CTPEDAG','X');
                                        TF.PutValue('PFS_CTSALARIE','-');
                                        Plafond := RendPlafondFraisForm(LeFrais,LeLieu,OrgCollecteur,MillesimeEC,DonnePopulation(LeSalarie)); //PT2
                                        Plafond := Plafond*Qte;
                                        If (Montant>Plafond) and (plafond>0) Then TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Plafond))
                                        Else TF.PutValue('PFS_MONTANTPLAF',FloatToStr(Arrondi(Montant,2)));
                                        //PT3 - Début
                                        TF.PutValue('PFS_PREDEFINI', GetPredefiniPopulation(TYP_POPUL_FORM_PREV));
                                        TF.PutValue('PFS_NODOSSIER', PgrendNodossier());
                                        TF.PutValue('PFS_CODEPOP','');
                                        If (VH_Paie.PGForGestPlafByPop) Then TF.PutValue('PFS_POPULATION', DonnePopulation(LeSalarie)); //PT2

                                        TFF := TobFormations.FindFirst(['PFO_SALARIE'],[LeSalarie], False);
                                        If TFF <> Nil Then
                                        Begin
                                             TF.PutValue('PFS_PGTYPESESSION',TFF.GetValue('PFO_PGTYPESESSION'));
                                             TF.PutValue('PFS_NOSESSIONMULTI',TFF.GetValue('PFO_NOSESSIONMULTI'));
                                             TF.PutValue('PFS_IDSESSIONFOR',TFF.GetValue('PFO_IDSESSIONFOR'));
                                             TF.PutValue('PFS_TYPEPLANPREV',TFF.GetValue('PFO_TYPEPLANPREV'));
                                        End
                                        Else
                                        Begin
                                             TF.PutValue('PFS_PGTYPESESSION','AUC');
                                             TF.PutValue('PFS_NOSESSIONMULTI',-1);
                                             TF.PutValue('PFS_IDSESSIONFOR','');
                                             TF.PutValue('PFS_TYPEPLANPREV','');
                                        End;
                                        //PT3 - Fin

                                        TF.InsertDB(Nil,False);
                                end;
                        end;
                end;

                //PT3 - Début
                // Gestion des multi-sessions : Mise à jour des frais totaux de la session en-tête
                If (TF <> Nil) And (TF.GetValue('PFS_PGTYPESESSION') = TYP_SOUSSESSION) Then MajSousSessionSal(LeSalarie,LeStage,TF.GetValue('PFS_NOMULTISESSION'));
                //PT3 - Fin

                Q := OpenSQL('SELECT SUM(PFS_MONTANT) AS MONTANT, SUM(PFS_MONTANTPLAF) AS PLAFOND FROM FRAISSALFORM WHERE '+
                'PFS_SALARIE="'+LeSalarie+'" AND PFS_MILLESIME="'+LeMillesime+'" AND PFS_CODESTAGE="'+LeStage+'" AND PFS_ORDRE='+LaSession+'',True);  //DB2
                Total := 0;
                TotalPlaf := 0;
                if not Q.eof then
                begin
                        Total := Q.FindField('MONTANT').AsFloat;
                        TotalPlaf := Q.FindField('PLAFOND').AsFloat;
                end;
                Ferme(Q);
                // MAJ SALAIRE ANIMATEUR DANS SESSIONSTAGE
                Q := OpenSQL('SELECT SUM(PAN_SALAIREANIM) SALAIRE FROM SESSIONANIMAT WHERE '+
                'PAN_CODESTAGE="'+LeStage+'" AND PAN_ORDRE='+LaSession+' AND PAN_MILLESIME="'+LeMillesime+'"',True);
                If Not Q.Eof then SalaireAnim := Q.FindField('SALAIRE').AsFloat
                Else SalaireAnim := 0;
                Ferme(Q);
                Q := OpenSQL('SELECT SUM(PFS_MONTANT) FRAIS FROM SESSIONANIMAT '+
                'LEFT JOIN FRAISSALFORM ON PAN_SALARIE=PFS_SALARIE AND PAN_CODESTAGE=PFS_CODESTAGE AND PAN_ORDRE=PFS_ORDRE AND PAN_MILLESIME=PFS_MILLESIME '+
                'WHERE PAN_FRASPEDAG="X" AND PAN_CODESTAGE="'+LeStage+'" AND PAN_ORDRE='+LaSession+' AND PAN_MILLESIME="'+LeMillesime+'"',True);//DB2
                If Not Q.Eof then FraisAnim := Q.FindField('FRAIS').AsFloat
                Else FraisAnim := 0;
                Ferme(Q);
                ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_COUTSALAIR='+StrfPOINT(SalaireAnim + FraisANim)+' '+  //DB2
                'WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE='+LaSession+' AND PSS_MILLESIME="'+LeMillesime+'"');  //DB2
                MAJCoutsFormation(LeMillesime,LeStage,LaSession);

                ExecuteSQL('UPDATE FORMATIONS SET PFO_FRAISREEL='+StrfPOINT(Total)+',PFO_FRAISPLAF='+StrfPOINT(TotalPlaf)+''+ //PT1
                ' WHERE PFO_SALARIE="'+LeSalarie+'" AND PFO_MILLESIME="'+LeMillesime+'" AND PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE='+LaSession+'');  //DB2
                CalcCtInvestSession('FRAIS',LeStage,LeMillesime,StrToInt(LaSession));
        end;
        TobFrais.free;
        FreeAndNil(TobFormations); //PT3
end;

procedure TOF_PGFRAISSALFORM.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var i : Integer;
    Plafond,Qte : Double;
    leFrais : String;
begin
        For i := 1 to VH_Paie.PGFNbFraisLibre do
        begin
                If Acol = COL_START_FRAIS+(i-1)*2 then
                begin
                        If (IsNumeric(Grille.CellValues[ACol,ARow])) and (Grille.cellValues[ACol+1,ARow] = '') then
                        begin
                                LeFrais := TabFrais[i];
                                Plafond := RendPlafondFraisForm(LeFrais,LeLieu,OrgCollecteur,MillesimeEC,DonnePopulation(Grille.CellValues[0,ARow])); //PT2
                                Qte := StrToFLoat(Grille.CellValues[ACol,ARow]);
                                Grille.cellValues[ACol+1,ARow] := FloatToStr(Arrondi(Plafond * Qte,2));
                        end;
                end;
        end;
end;

Initialization
  registerclasses ( [ TOF_PGFRAISSALFORM ] ) ;
end.

