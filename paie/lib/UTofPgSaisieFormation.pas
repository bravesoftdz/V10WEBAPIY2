{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSAISIEFORMATION ()
Mots clefs ... : TOF;PGSAISIEFORMATION
*****************************************************************
PT1  | 24/04/2003 | V_42  | JL | Développement pour CWAS + modif lookuplist
PT2  | 20/04/2005 | V_60  | JL | FQ 12165 Gestion affichage récupération prévisionnel en fonction paramsoc
---  | 20/03/2006 |       | JL | modification clé annuaire ----
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT3  | 25/01/2006 | V_750 | JL | Suppression du treeview
PT4  | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour le lookuplist pour le critère salarié
PT5  | 11/04/2007 | V_720 | FL | FQ 14047 Ajout du bouton de récupération du prévisionnel et suppression des liaisons
     |            |       |    | vers l'onglet "Récupération Budget" (PINSCSAL)
PT6  | 02/04/2008 | V_803 | FL | Ajout d'une saisie groupée
PT7  | 02/04/2008 | V_803 | FL | Correction de la récupération du prévisionnel
PT8  | 25/04/2008 | V_804 | FL | Gestion de l'inscription multiple
PT9  | 27/05/2008 | V_804 | FL | Ajout de champs manquants lors des inscriptions en masse et récupération prévisionnel
PT10 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT11 | 08/09/2008 | V_804 | NA | FQ 15695 Mettre l'état "Validé" si récup du prévisionnel et saisie en masse
}
Unit UTofPgSaisieFormation;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,uTableFiltre, EntPaie,
     SaisieList,hTB97,LookUp,UTOB,ParamDat,HSysMenu,UTOFPGMul_InscFor,ParamSoc,PGOutilsFormation,pgoutils;

Type
  TOF_PGSAISIEFORMATION = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TF : TTableFiltre;
    TypeSaisie,LeStage,LaSession,LeMillesime,MillesimeEC : String;
    recupererInsc : Boolean;
    procedure GrilleSalariesDblClick(Sender : TObject);
    procedure BSalarieClick(Sender : TObject);
    procedure SalarieElipsisClick(Sender : TObject);
    procedure RecupPrevisionnel(Salarie,TypePlan,Typologie : String; FormationEffectuee:boolean=True);     //PT5
    procedure BRecupClick (Sender : TObject); //PT5
    procedure SaisieEnMasse (Sender : TObject); //PT6
  end ;

  TOF_PGSAISIEFORMATIONINIT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

Uses UtilPGI;

procedure TOF_PGSAISIEFORMATION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGSAISIEFORMATION.OnLoad ;
begin
  Inherited ;
	//PT10
	// N'affiche pas les salariés confidentiels
	If PGBundleHierarchie Then
		TF.WhereTable := TF.WhereTable + ' AND PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
	Else
		TF.WhereTable := TF.WhereTable + ' AND PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
end ;


procedure TOF_PGSAISIEFORMATION.OnArgument (S : String ) ;
var DateDebut,DateFin : TDateTime;
    ESalarie : THEdit;
    BSalarie, Brecup, BT : TToolBarButton97;
    GSalaries : THGrid;
    Utilisateur : String;
begin
  Inherited ;
{        If Not VH_Paie.PGForPrevisionnel then       //PT5
        begin
             SetControlVisible('PINSCSAL',False);  //PT2
             SetControlProperty('PINSCSAL','TabVisible',False);
        end
        else
        begin
             SetControlVisible('PINSCSAL',True);  //PT2
             SetControlProperty('PINSCSAL','TabVisible',True);
        end;
}
        Utilisateur := V_PGI.UserSalarie;
        TF  :=  TFSaisieList(Ecran).LeFiltre;
        TypeSaisie := Trim(ReadTokenPipe(S,';'));
        LeStage := Trim(ReadTokenPipe(S,';'));
        LaSession := Trim(ReadTokenPipe(S,';'));
        LeMillesime := Trim(ReadTokenPipe(S,';'));
        recupererInsc := False;
        If TypeSaisie = 'RECUPPREVPLA' then
        begin
             recupererInsc := True;
             TypeSaisie := 'SAISIESTAGEFORMAT';
        end;
        MillesimeEC := RendMillesimeRealise(DateDebut,DateFin);
        TF.WhereTable := 'WHERE PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE="'+LaSession+'" AND PFO_MILLESIME="'+LeMillesime+'"';
        If (TypeSaisie = 'SAISIEINSC') OR (TypeSaisie = 'CWASSAISIEFORMATION') or (TypeSaisie = 'SAISIESESSION') Then
        begin
                GSalaries := THGrid(GetControl('GSALARIES'));
                If GSalaries <> Nil Then GSalaries.OnDblClick := GrilleSalariesDblClick;
        end;
        BSalarie := TToolBarButton97(GetControl('BSALARIE'));
        If BSalarie <> Nil then BSalarie.OnClick := BSalarieClick;
        SetControlVisible('BPARAMLISTE',True);
        SetActiveTabSheet('PSESSION');
        ESalarie := THEdit(GetControl('PFO_SALARIE'));
        If ESalarie <> Nil then ESalarie.OnElipsisClick := SalarieElipsisClick;
        //PT5 - Début
        BRecup := TToolBarButton97(GetControl('BRECUPERATION'));
        If BRecup <> Nil then BRecup.OnClick := BRecupClick;
        //PT5 - Fin
        //PT6 - Début
      	Bt := TToolbarButton97(GetControl('BMULTISAL'));
       	If Bt <> Nil then
       	Begin
            //If PGBundleHierarchie Or  Then
            //Begin
            	//Bt := TToolbarButton97(GetControl('BMULTISAL'));
        		Bt.onClick := SaisieEnMasse;
        		Bt.Visible := True;
        	//End;
        End;
        //PT6 - Fin
end ;

procedure TOF_PGSAISIEFORMATION.SalarieElipsisClick(Sender : TObject);
var StFrom,StWhere : String;
begin
	//PT6
	If PGBundleHierarchie Then
		ElipsisSalarieMultidos(Sender)
	Else
	Begin
        StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
        StWhere := '';
        {$IFDEF EMANAGER}
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        {$ENDIF}
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT4
        LookupList(THEdit(Sender), 'Liste des salariés',StFrom,'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
    End;
end;

procedure TOF_PGSAISIEFORMATION.GrilleSalariesDblClick(Sender : TObject);
var Salarie : String;
    Typologie,TypePlan : String;
begin
        Salarie  :=  THGrid(Sender).CellValues [1,THGrid(Sender).Row];
        Typologie  :=  THGrid(Sender).CellValues [7,THGrid(Sender).Row];
        TypePlan  :=  THGrid(Sender).CellValues [0,THGrid(Sender).Row];
        RecupPrevisionnel(Salarie,TypePlan,Typologie);
end;


procedure TOF_PGSAISIEFORMATION.BSalarieClick(Sender : TObject);
var St : String;
begin
        TF.DisableTOM;
        St := TF.GetValue('PFO_CODESTAGE')+';'+IntToStr(TF.GetValue('PFO_ORDRE'))+';'+TF.GetValue('PFO_MILLESIME')+';'+TF.GetValue('PFO_SALARIE');
        AGLLanceFiche('PAY','FORMATIONS','',St,'');
        TF.EnableTOM;
end;

procedure TOF_PGSAISIEFORMATION.RecupPrevisionnel(Salarie,TypePlan,Typologie : String; FormationEffectuee:Boolean=True);  //PT5
var Stage,Session,Millesime : String;
    Q : TQuery;
    DatePaie : TDateTime;
    DefautHTpsT : Boolean;
    Req,Pfx     : String;
begin
        DefautHTpsT := GetParamSoc('SO_PGDIFTPSTRAV');
        If Salarie = '' Then Exit;
        Stage := LeStage;
        Session := laSession;
        Millesime := LeMillesime;
        If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" AND PFO_CODESTAGE="'+Stage+'" '+
        ' AND PFO_ORDRE='+Session+' AND PFO_MILLESIME="'+Millesime+'"') then  //DB2
        begin
                PGIBox('Récupération impossible car ce salarié est déjà inscrit',Ecran.Caption);
                Exit;
        end;
        TF.StartUpdate;
        TF.Insert;
        TF.PutValue('PFO_SALARIE',Salarie);
        If Typologie <> '' then TF.PutValue('PFO_TYPOFORMATION',Typologie)
        else TF.PutValue('PFO_TYPOFORMATION','001');
        //PT5 - Début
        // One ne recherche la date du bulletin que si la formation a été effectuée (donc pas dans le cas d'une inscription).
        // Sinon, on considère 01/01/1900.
        If (FormationEffectuee) Then
        Begin
            //PT9
        	If PGBundleHierarchie Then
        		Q := OpenSQL('SELECT MAX (PPU_DATEFIN) MAXDATE FROM '+GetBase(GetBaseSalarie(Salarie),'PAIEENCOURS')+' WHERE PPU_SALARIE="'+Salarie+'"',True)
        	Else
                Q := OpenSQL('SELECT MAX (PPU_DATEFIN) MAXDATE FROM PAIEENCOURS WHERE PPU_SALARIE="'+Salarie+'"',True);  //PT7
             If Not Q.Eof then DatePaie := Q.FindField('MAXDATE').AsDateTime
             Else DatePaie := V_PGI.DateEntree;
             DatePaie := PlusMois(DatePaie,1);
             DatePaie := FinDeMois(DatePaie);
             Ferme(Q);
             If DatePaie < V_PGI.dateEntree then DatePaie := FinDeMois(V_PGI.DateEntree);
             TF.PutValue('PFO_DATEPAIE',DatePaie);
        End
        Else
        Begin
             DatePaie := iDate1900;
             TF.PutValue('PFO_DATEPAIE',DatePaie);
        End;
        //PT5 - Fin
        TF.PutValue('PFO_TYPEPLANPREV',TypePlan);
        TF.PutValue('PFO_CODESTAGE',Stage);
        TF.PutValue('PFO_ORDRE',StrToInt(Session));
        TF.PutValue('PFO_MILLESIME',Millesime);  //PT9

        //PT7 - Début
        Req := 'SELECT PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_PRENOM,PSE_RESPONSFOR'+
        ',PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4 FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE WHERE PSA_SALARIE="'+Salarie+'"';

        Req := AdaptMultiDosForm (Req, False);
        If PGBundleHierarchie Then Pfx := 'PSI' Else Pfx := 'PSA';
        //PT7 - Fin

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


        Q := OpenSQL('SELECT PSS_DATEDEBUT,PSS_HEUREDEBUT,PSS_DATEFIN,PSS_HEUREFIN,PSS_DUREESTAGE,PSS_ORGCOLLECTSGU,PSS_DEBUTDJ,PSS_DEBUTDJ,PSS_LIEUFORM,PSS_CENTREFORMGU'+
        ',PSS_NATUREFORM,PSS_FINDJ,PSS_DEBUTDJ,PSS_FORMATION1,PSS_FORMATION2,PSS_FORMATION3,PSS_FORMATION4,PSS_FORMATION5,PSS_FORMATION6,PSS_FORMATION7,PSS_FORMATION8,'+
        'PSS_INCLUSDECL,PSS_NOSESSIONMULTI,PSS_IDSESSIONFOR,PSS_PGTYPESESSION '+
        'FROM SESSIONSTAGE LEFT JOIN STAGE ON '+
        ' PSS_MILLESIME=PST_MILLESIME AND PST_CODESTAGE=PSS_CODESTAGE'+
        ' WHERE PSS_ORDRE='+Session+' AND PSS_CODESTAGE="'+Stage+'" AND PSS_MILLESIME="'+Millesime+'"',True);  //DB2
        if not Q.eof then
        begin
                TF.PutValue('PFO_DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsDateTime);
                TF.PutValue('PFO_DATEFIN',Q.FindField('PSS_DATEFIN').AsDateTime);
                TF.PutValue('PFO_NATUREFORM',Q.FindField('PSS_NATUREFORM').AsString);
                TF.PutValue('PFO_LIEUFORM',Q.FindField('PSS_LIEUFORM').AsString);
                TF.PutValue('PFO_CENTREFORMGU',Q.FindField('PSS_CENTREFORMGU').AsString);//DB2
                TF.PutValue('PFO_DEBUTDJ',Q.FindField('PSS_DEBUTDJ').AsString);
                TF.PutValue('PFO_FINDJ',Q.FindField('PSS_FINDJ').AsString);
                TF.PutValue('PFO_ORGCOLLECT',   -1);  //PT9
                TF.PutValue('PFO_ORGCOLLECTGU',Q.FindField('PSS_ORGCOLLECTSGU').AsString);
                TF.PutValue('PFO_NBREHEURE',Q.FindField('PSS_DUREESTAGE').AsFloat);
                if (DefautHTpsT) or (TypePlan = 'PLF') then
                begin
                     TF.PutValue('PFO_HTPSTRAV', Q.FindField('PSS_DUREESTAGE').AsFloat);
                     TF.PutValue('PFO_HTPSNONTRAV', 0);
                end
                else
                begin
                     TF.PutValue('PFO_HTPSTRAV', 0);
                     TF.PutValue('PFO_HTPSNONTRAV', Q.FindField('PSS_DUREESTAGE').AsFloat);
                end;
                
                //PT9 - Début
		        TF.PutValue('PFO_HEUREDEBUT',    Q.FindField('PSS_HEUREDEBUT').AsDateTime);
		        TF.PutValue('PFO_HEUREFIN',      Q.FindField('PSS_HEUREFIN').AsDateTime);
		        TF.PutValue('PFO_FORMATION1',    Q.FindField('PSS_FORMATION1').AsString);
		        TF.PutValue('PFO_FORMATION2',    Q.FindField('PSS_FORMATION2').AsString);
		        TF.PutValue('PFO_FORMATION3',    Q.FindField('PSS_FORMATION3').AsString);
		        TF.PutValue('PFO_FORMATION4',    Q.FindField('PSS_FORMATION4').AsString);
		        TF.PutValue('PFO_FORMATION5',    Q.FindField('PSS_FORMATION5').AsString);
		        TF.PutValue('PFO_FORMATION6',    Q.FindField('PSS_FORMATION6').AsString);
		        TF.PutValue('PFO_FORMATION7',    Q.FindField('PSS_FORMATION7').AsString);
		        TF.PutValue('PFO_FORMATION8',    Q.FindField('PSS_FORMATION8').AsString);
		        TF.PutValue('PFO_INCLUSDECL',    Q.FindField('PSS_INCLUSDECL').AsString);

		        TF.PutValue('PFO_NOSESSIONMULTI', Q.FindField('PSS_NOSESSIONMULTI').AsString);
		        TF.PutValue('PFO_IDSESSIONFOR',   Q.FindField('PSS_IDSESSIONFOR').AsString);
		        TF.PutValue('PFO_PGTYPESESSION',  Q.FindField('PSS_PGTYPESESSION').AsString);

		        TF.PutValue('PFO_DATECOINVES',	IDate1900);
		        TF.PutValue('PFO_DATEACCEPT',	IDate1900);
		        TF.PutValue('PFO_REFUSELE',		IDate1900);
		        TF.PutValue('PFO_REPORTELE',	IDate1900);


            // PT11 If GetParamSocSecur('SO_PGFORMVALIDREA', False) Then
            If not GetParamSocSecur('SO_PGFORMVALIDREA', False) Then   // pt11
                    TF.PutValue('PFO_ETATINSCFOR',   'ATT')
                Else
                    TF.PutValue('PFO_ETATINSCFOR',   'VAL');

		        //PT9 - Fin

                //PT5 - Début
                // Si la formation n'a pas été effectuée, il ne faut pas cocher la case
                If (FormationEffectuee) Then
                Begin
                     If TypeSaisie = 'SAISIEFP' Then
                     begin
                         TF.PutValue('PFO_EFFECTUE','X');
                     end;
                End
                Else
                Begin
                     TF.PutValue('PFO_EFFECTUE','-');
                End;
                //PT5 - Fin
        end;
        Ferme(Q);

        //PT7 - Début
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
        //PT7 - Fin
        // Si l'update ne se passe pas bien, on rollback sinon la grille ne se rafraîchit pas
        //PT5 - Début
        If Not TF.OM.UpdateRecord Then
             TF.cancel
        Else
          Begin
               TF.DisableTOM;
               TF.Post;
               TF.EnableTOM;
          End;
        //PT5 - Fin
        TF.EndUpdate;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008 / PT6
Modifié le ... :   /  /
Description .. : Saisie groupée des inscriptions
Mots clefs ... :
*****************************************************************}
procedure TOF_PGSAISIEFORMATION.SaisieEnMasse (Sender : TObject);
var Stage,Session,Millesime     : String;
    DatePaie                    : TDateTime;
    Salarie, Pfx                : String;
    ListeSalaries,RetSauve,Ret  : String;
    TobSalaries,TS,TobFormation : TOB;
    Bt                          : TToolbarButton97;
begin
    Stage := LeStage;
    Session := laSession;
    Millesime := LeMillesime;
    
	If PGBundleHierarchie Or GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE) Then //PT8
		Pfx := 'PSI'
    Else
    	Pfx := 'PSA';
    
    Ret := AGLLanceFiche('PAY','INTERIMINSC_MUL','','','SESSION;PFO_CODESTAGE="'+Stage+'" AND PFO_MILLESIME="'+Millesime+'"');
    
    If Ret = '' then exit;
    
     // Chargement des données des salariés
     RetSauve := Ret;
     While RetSauve<>'' Do
     Begin
          ListeSalaries := ListeSalaries + '"' + ReadTokenSt(RetSauve) + '"';
          If RetSauve <> '' Then ListeSalaries := ListeSalaries + ',';
     End;
     
     // Chargement des données salariés
     TobSalaries := TOB.Create('LesSalaries', Nil, -1);
     If Pfx = 'PSI' Then //PT8
     	TobSalaries.LoadDetailFromSQL('SELECT INTERIMAIRES.*,PSE_RESPONSFOR ' +
          'FROM INTERIMAIRES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSI_INTERIMAIRE WHERE PSI_INTERIMAIRE IN ('+ListeSalaries+')')
     Else
     	TobSalaries.LoadDetailFromSQL('SELECT SALARIES.*,PSE_RESPONSFOR ' +
          'FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE WHERE PSA_SALARIE IN ('+ListeSalaries+')');

     // Chargement des informations du stage
     TobFormation := Tob.Create('LaFormation', nil, -1);
     TobFormation.LoadDetailFromSQL('SELECT * FROM SESSIONSTAGE '+
        ' WHERE PSS_ORDRE='+Session+' AND PSS_CODESTAGE="'+Stage+'" AND PSS_MILLESIME="'+Millesime+'"');

     TF.DisableTom;
     TF.StartUpdate;
     
     // Boucle sur les stagiaires sélectionnés
     While Ret <> '' Do
     Begin
          Salarie := ReadTokenSt(Ret);
          If PFx = 'PSI' Then //PT8
          	TS := TobSalaries.FindFirst(['PSI_INTERIMAIRE'],[Salarie], False)
          Else
          	TS := TobSalaries.FindFirst(['PSA_SALARIE'],[Salarie], False);

          If TS = Nil Then Break;  // Ne doit pas passer par ici normalement!
        
        
        If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" AND PFO_CODESTAGE="'+Stage+'" '+
        ' AND PFO_ORDRE='+Session+' AND PFO_MILLESIME="'+Millesime+'"') then 
        begin
                PGIBox('Ce salarié est déjà inscrit : '+RechDom('PGSALARIEINT', Salarie, False),Ecran.Caption);
                Exit;
        end;
        
        TF.Insert;
        
        TF.PutValue('PFO_SALARIE',Salarie);
        TF.PutValue('PFO_TYPOFORMATION','001');

        DatePaie := iDate1900;
        TF.PutValue('PFO_DATEPAIE',DatePaie);

        TF.PutValue('PFO_TYPEPLANPREV','PLF');
        TF.PutValue('PFO_CODESTAGE',Stage);
        TF.PutValue('PFO_ORDRE',StrToInt(Session));
        TF.PutValue('PFO_MILLESIME',Millesime); //PT9
        
        TF.PutValue('PFO_TRAVAILN1',	TS.GetValue(Pfx + '_TRAVAILN1'));
        TF.PutValue('PFO_TRAVAILN2',	TS.GetValue(Pfx + '_TRAVAILN2'));
        TF.PutValue('PFO_TRAVAILN3',	TS.GetValue(Pfx + '_TRAVAILN3'));
        TF.PutValue('PFO_TRAVAILN4',	TS.GetValue(Pfx + '_TRAVAILN4'));
        TF.PutValue('PFO_CODESTAT',		TS.GetValue(Pfx + '_CODESTAT'));
        TF.PutValue('PFO_NOMSALARIE',	TS.GetValue(Pfx + '_LIBELLE'));
        TF.PutValue('PFO_PRENOM',		TS.GetValue(Pfx + '_PRENOM'));
        TF.PutValue('PFO_ETABLISSEMENT',TS.GetValue(Pfx + '_ETABLISSEMENT'));
        TF.PutValue('PFO_LIBEMPLOIFOR',	TS.GetValue(Pfx + '_LIBELLEEMPLOI'));
        TF.PutValue('PFO_LIBREPCMB1',	TS.GetValue(Pfx + '_LIBREPCMB1'));
        TF.PutValue('PFO_LIBREPCMB2',	TS.GetValue(Pfx + '_LIBREPCMB2'));
        TF.PutValue('PFO_LIBREPCMB3',	TS.GetValue(Pfx + '_LIBREPCMB3'));
        TF.PutValue('PFO_LIBREPCMB4',	TS.GetValue(Pfx + '_LIBREPCMB4'));
        
        TF.PutValue('PFO_RESPONSFOR',	TS.GetValue('PSE_RESPONSFOR'));

        TF.PutValue('PFO_DATEDEBUT',	TobFormation.Detail[0].GetValue('PSS_DATEDEBUT'));
        TF.PutValue('PFO_DATEFIN',		TobFormation.Detail[0].GetValue('PSS_DATEFIN'));
        TF.PutValue('PFO_NATUREFORM',	TobFormation.Detail[0].GetValue('PSS_NATUREFORM'));
        TF.PutValue('PFO_LIEUFORM',		TobFormation.Detail[0].GetValue('PSS_LIEUFORM'));
        TF.PutValue('PFO_CENTREFORMGU',	TobFormation.Detail[0].GetValue('PSS_CENTREFORMGU'));
        TF.PutValue('PFO_DEBUTDJ',		TobFormation.Detail[0].GetValue('PSS_DEBUTDJ'));
        TF.PutValue('PFO_FINDJ',		TobFormation.Detail[0].GetValue('PSS_FINDJ'));
        TF.PutValue('PFO_ORGCOLLECT',   -1);
        TF.PutValue('PFO_ORGCOLLECTGU',	TobFormation.Detail[0].GetValue('PSS_ORGCOLLECTSGU'));
        TF.PutValue('PFO_NBREHEURE',	TobFormation.Detail[0].GetValue('PSS_DUREESTAGE'));
        
        TF.PutValue('PFO_HEUREDEBUT',    TobFormation.Detail[0].GetValue('PSS_HEUREDEBUT'));
        TF.PutValue('PFO_HEUREFIN',      TobFormation.Detail[0].GetValue('PSS_HEUREFIN'));
        TF.PutValue('PFO_FORMATION1',    TobFormation.Detail[0].GetValue('PSS_FORMATION1'));
        TF.PutValue('PFO_FORMATION2',    TobFormation.Detail[0].GetValue('PSS_FORMATION2'));
        TF.PutValue('PFO_FORMATION3',    TobFormation.Detail[0].GetValue('PSS_FORMATION3'));
        TF.PutValue('PFO_FORMATION4',    TobFormation.Detail[0].GetValue('PSS_FORMATION4'));
        TF.PutValue('PFO_FORMATION5',    TobFormation.Detail[0].GetValue('PSS_FORMATION5'));
        TF.PutValue('PFO_FORMATION6',    TobFormation.Detail[0].GetValue('PSS_FORMATION6'));
        TF.PutValue('PFO_FORMATION7',    TobFormation.Detail[0].GetValue('PSS_FORMATION7'));
        TF.PutValue('PFO_FORMATION8',    TobFormation.Detail[0].GetValue('PSS_FORMATION8'));
        TF.PutValue('PFO_INCLUSDECL',    TobFormation.Detail[0].GetValue('PSS_INCLUSDECL'));
        
        TF.PutValue('PFO_HTPSTRAV', 	TobFormation.Detail[0].GetValue('PSS_DUREESTAGE'));
        TF.PutValue('PFO_HTPSNONTRAV', 	0);
        
        TF.PutValue('PFO_NOSESSIONMULTI', TobFormation.Detail[0].GetValue('PSS_NOSESSIONMULTI'));
        TF.PutValue('PFO_IDSESSIONFOR',   TobFormation.Detail[0].GetValue('PSS_IDSESSIONFOR'));
        TF.PutValue('PFO_PGTYPESESSION',  TobFormation.Detail[0].GetValue('PSS_PGTYPESESSION'));
                
        TF.PutValue('PFO_DATECOINVES',	IDate1900);
        TF.PutValue('PFO_DATEACCEPT',	IDate1900);
        TF.PutValue('PFO_REFUSELE',		IDate1900);
        TF.PutValue('PFO_REPORTELE',		IDate1900);
        
        //pt11
      If not GetParamSocSecur('SO_PGFORMVALIDREA',False) Then  // pt11
            TF.PutValue('PFO_ETATINSCFOR',   'ATT')
        Else
            TF.PutValue('PFO_ETATINSCFOR',   'VAL');

                TF.PutValue('PFO_EFFECTUE',		'-');

			If PGBundleInscFormation Then
			Begin
        		TF.PutValue('PFO_NODOSSIER', GetNoDossierSalarie(Salarie));
        		TF.PutValue('PFO_PREDEFINI', 'DOS');
        	End
        	Else
        	Begin
        		TF.PutValue('PFO_NODOSSIER', '000000');
        		TF.PutValue('PFO_PREDEFINI', 'STD');
        	End;

		TF.Post;
     End;

     TF.EnableTOM;
     TF.EndUpdate;

     if TobFormation <> nil then FreeAndNil(TobFormation);
     if TobSalaries      <> nil then FreeAndNil(TobSalaries);

     TF.RefreshEntete;
     Bt := TToolBarButton97(GetControl('BCHERCHE'));
     if Bt <> nil then Bt.Click;
end;


{TOF_PGSAISIEFORMATIONINIT}

procedure TOF_PGSAISIEFORMATIONINIT.OnArgument (S : String ) ;
var TF : TTableFiltre;
begin
  Inherited ;
        TF  :=  TFSaisieList(Ecran).LeFiltre;
        TF.WhereTable := 'WHERE PFO_SALARIE=:PSA_SALARIE AND PFO_NATUREFORM="004"';
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 11/04/2007 / PT5
Modifié le ... :   /  /
Description .. : Récupération du prévisionnel pour injection des 
Suite ........ : données directement dans l'écran
Mots clefs ... : PREVISIONNEL
*****************************************************************}
procedure TOF_PGSAISIEFORMATION.BRecupClick (Sender : TObject);
var i : Integer;
begin
     AGLLanceFiche('PAY','MUL_INSCFOR','','','RECUPPREVISIONNEL;'+LeStage+';'+LeMillesime); //PT7
     If PGTobRecupPrev <> Nil then
     begin
          For i := 0 to PGTobRecupPrev.Detail.Count - 1 do
          begin
               { Appel de la récupération du prévisionnel avec False pour ne pas rechercher les salaires et impacter
                 les jours de formation }
               RecupPrevisionnel(PGTobRecupPrev.Detail[i].GetValue('SALARIE'),PGTobRecupPrev.Detail[i].GetValue('TYPEPLAN'),'', False);
          end;
          PGTobRecupPrev.Free;
          MAJCoutsFormation(LeMillesime, LeStage, LaSession, Nil);
          TF.RefreshLignes;
     end;
end;

Initialization
  registerclasses ( [ TOF_PGSAISIEFORMATION,TOF_PGSAISIEFORMATIONINIT ] ) ;
end.



