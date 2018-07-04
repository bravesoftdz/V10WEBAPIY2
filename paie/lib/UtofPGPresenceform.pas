{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPRESENCEFORM ()
Mots clefs ... : TOF;PGPRESENCEFORM
*****************************************************************
--- | 20/03/2006 |       | JL | modification clé annuaire ----
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT1 | 25/01/2007 | V_75  | JL | Suppression du Tree view
PT2 | 20/04/2007 | V_720 | FL | FQ 14054 Ajout du paramétrage de la combo cursus afin de limiter les choix
PT3 | 02/04/2008 | V_803 | FL | Corrections pour le partage formation
PT4 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
}
Unit UtofPGPresenceform;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,uTableFiltre,PGOutilsFormation,
     SaisieList,hTB97,LookUp,ParamDat,UtomSessionStage,UTOFPGMul_InscFor,ParamSoc;

Type
  TOF_PGPRESENCEFORM = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TF : TTableFiltre;
    MillesimeEC : String;
    LeStage,LeMillesime,LaSession,TypeSaisie : String;
    recupererInsc : Boolean;
    procedure BRecupClick (Sender : TObject);
    procedure RecupPrevisionnel(Salarie,TypePlan : String);
    procedure InsererDIFBulletin;
    procedure SupprimerDIF;
   end ;

Implementation

Uses UtilPGI;

procedure TOF_PGPRESENCEFORM.OnUpdate ;
begin
  Inherited ;

end ;

procedure TOF_PGPRESENCEFORM.OnLoad ;
var i : Integer;
begin
  Inherited ;

        If TypeSaisie = 'SAISIEFP' then
        begin
                If (LeStage <> '') and (laSession <> '') and (LeMillesime <> '') Then
                begin
                        SetControlText('PSS_CODESTAGE',LeStage);
                        SetControlText('PSS_MILLESIME',LeMillesime);
                        SetControlText('XX_WHERE','PSS_ORDRE='+LaSession);  //DB2
                        SetControlVisible('BTree',False);
                        SetControlVisible('PANTREEVIEW',False);
                end;
        end;
        
        //PT4
        // N'affiche pas les salariés confidentiels
        If PGBundleHierarchie Then
        	TF.WhereTable := TF.WhereTable + ' AND PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
        Else
        	TF.WhereTable := TF.WhereTable + ' AND PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
        
        TFSaisieList(Ecran).BCherche.Click;
        If RecupererInsc then
        begin
        If PGTobRecupPrev <> Nil then
             begin
                  For i := 0 to PGTobRecupPrev.Detail.Count - 1 do
                  begin
                       RecupPrevisionnel(PGTobRecupPrev.Detail[i].GetValue('SALARIE'),PGTobRecupPrev.Detail[i].GetValue('TYPEPLAN'));
                  end;
                  PGTobRecupPrev.Free;
             end;
        end;
end ;

procedure TOF_PGPRESENCEFORM.OnArgument (S : String ) ;
var DD,DF : TDateTime;
    BRecup : TToolBarButton97;
begin
        TypeSaisie := ReadTokenPipe(S,';');
        LeStage := ReadTokenPipe(S,';');
        LaSession := ReadTokenPipe(S,';');
        LeMillesime := ReadTokenPipe(S,';');
        RecupererInsc := False;
        If TypeSaisie = 'RECUPPREVREA' then
        begin
             RecupererInsc := True;
             TypeSaisie := 'SAISIEFP';
        end;
        TF  :=  TFSaisieList(Ecran).LeFiltre;
        if (Ecran <> nil) and (Ecran is TFSaisieList ) then
        MillesimeEC := RendMillesimeRealise(DD,DF);
        TF.WhereTable := 'WHERE PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE="'+LaSession+'" AND PFO_MILLESIME="'+LeMillesime+'"';
        If (PGBundleInscFormation) Then
        Begin
        	If Not PGDroitMultiForm then 
        		TF.WhereTable := TF.WhereTable + DossiersAInterroger('',V_PGI.NoDossier,'PFO',False,True)
        	Else
        		TF.WhereTable := TF.WhereTable + DossiersAInterroger('','','PFO',False,True);
       	End;
        BRecup := TToolBarButton97(GetControl('BRECUPERATION'));
        If BRecup <> Nil then BRecup.OnClick := BRecupClick;
        SetControlProperty('PFO_CURSUS','Plus',' AND PCU_CURSUS IN (SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_CODESTAGE="'+LeStage+'")'); //PT2
end ;

procedure TOF_PGPRESENCEFORM.BRecupClick (Sender : TObject);
var i : Integer;
begin
     AGLLanceFiche('PAY','MUL_INSCFOR','','','RECUPPRESENCE;'+LeStage);
     If PGTobRecupPrev <> Nil then
     begin
          For i := 0 to PGTobRecupPrev.Detail.Count - 1 do
          begin
          	   RecupPrevisionnel(PGTobRecupPrev.Detail[i].GetValue('SALARIE'),PGTobRecupPrev.Detail[i].GetValue('TYPEPLAN'));
          end;
          PGTobRecupPrev.Free;
          MAJCoutsFormation(LeMillesime, LeStage, LaSession, Nil);
          TF.RefreshLignes;
     end;
end;

procedure TOF_PGPRESENCEFORM.RecupPrevisionnel(Salarie,TypePlan : String);
var Stage,Session,Millesime : String;
    Q : TQuery;
    DatePaie : TDateTime;
    DefautHTpsT : Boolean;
    Req,Pfx     : String;
begin
        DefautHTpsT := GetParamSoc('SO_PGDIFTPSTRAV');

        //PT3 - Début
//        Stage := TF.TOBFiltre.GetValue('PSS_CODESTAGE');
//        Session := IntToStr(TF.TOBFiltre.GetValue('PSS_ORDRE'));
//        Millesime := TF.TOBFiltre.GetValue('PSS_MILLESIME');
        Stage := LeStage;
        Session :=  LaSession;
        Millesime := LeMillesime;
        //PT3 - Fin

        If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" AND PFO_CODESTAGE="'+Stage+'" '+
        ' AND PFO_ORDRE='+Session+' AND PFO_MILLESIME="'+Millesime+'"') then  //DB2
        begin
                PGIBox('Récupération impossible car ce salarié est déjà inscrit',Ecran.Caption);
                Exit;
        end;
        
        TF.StartUpdate;
        TF.Insert;
        TF.PutValue('PFO_SALARIE',Salarie);
        TF.PutValue('PFO_TYPOFORMATION','001');
        
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
        
        TF.PutValue('PFO_TYPEPLANPREV',TypePlan);
        TF.PutValue('PFO_CODESTAGE',Stage);
        TF.PutValue('PFO_ORDRE',StrToInt(Session));
        TF.PutValue('PFO_MILLESIME',Millesime); 
        
        //PT3 - Début
        Req := 'SELECT PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_PRENOM,PSE_RESPONSFOR'+
        ',PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4 FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE WHERE PSA_SALARIE="'+Salarie+'"';

		Req := AdaptMultiDosForm (Req, False);
        If PGBundleHierarchie Then Pfx := 'PSI' Else Pfx := 'PSA';
        //PT3 - Fin
                
        Q := OpenSQL(Req,True);
        if not Q.eof then
        begin
                TF.PutValue('PFO_TRAVAILN1',	Q.FindField(Pfx+'_TRAVAILN1').AsString);
                TF.PutValue('PFO_TRAVAILN2',	Q.FindField(Pfx+'_TRAVAILN2').AsString);
                TF.PutValue('PFO_TRAVAILN3',	Q.FindField(Pfx+'_TRAVAILN3').AsString);
                TF.PutValue('PFO_TRAVAILN4',	Q.FindField(Pfx+'_TRAVAILN4').AsString);
                TF.PutValue('PFO_CODESTAT',		Q.FindField(Pfx+'_CODESTAT').AsString);
                TF.PutValue('PFO_NOMSALARIE',	Q.FindField(Pfx+'_LIBELLE').AsString);
                TF.PutValue('PFO_PRENOM',		Q.FindField(Pfx+'_PRENOM').AsString);
                TF.PutValue('PFO_ETABLISSEMENT',Q.FindField(Pfx+'_ETABLISSEMENT').AsString);
                TF.PutValue('PFO_LIBEMPLOIFOR', Q.FindField(Pfx+'_LIBELLEEMPLOI').AsString);
                TF.PutValue('PFO_RESPONSFOR',	Q.FindField('PSE_RESPONSFOR').AsString);
                TF.PutValue('PFO_LIBREPCMB1',	Q.findField(Pfx+'_LIBREPCMB1').AsString);
                TF.PutValue('PFO_LIBREPCMB2',	Q.findField(Pfx+'_LIBREPCMB2').AsString);
                TF.PutValue('PFO_LIBREPCMB3',	Q.findField(Pfx+'_LIBREPCMB3').AsString);
                TF.PutValue('PFO_LIBREPCMB4',	Q.findField(Pfx+'_LIBREPCMB4').AsString);
        end;
        Ferme(Q);


        Q := OpenSQL('SELECT PSS_DATEDEBUT,PSS_HEUREDEBUT,PSS_DATEFIN,PSS_HEUREFIN,PSS_DUREESTAGE,PSS_ORGCOLLECTSGU,PSS_DEBUTDJ,PSS_DEBUTDJ,PSS_LIEUFORM,PSS_CENTREFORMGU'+
        ',PSS_NATUREFORM,PSS_FINDJ,PSS_DEBUTDJ,PSS_FORMATION1,PSS_FORMATION2,PSS_FORMATION3,PSS_FORMATION4,PSS_FORMATION5,PSS_FORMATION6,PSS_FORMATION7,PSS_FORMATION8,'+
        'PSS_INCLUSDECL,PSS_NOSESSIONMULTI,PSS_IDSESSIONFOR,PSS_PGTYPESESSION '+
        'FROM SESSIONSTAGE LEFT JOIN STAGE ON '+
        ' PSS_MILLESIME=PST_MILLESIME AND PST_CODESTAGE=PSS_CODESTAGE'+
        ' WHERE PSS_ORDRE='+LaSession+' AND PSS_CODESTAGE="'+LeStage+'" AND PSS_MILLESIME="'+LeMillesime+'"',True);  //DB2
        if not Q.eof then
        begin
                TF.PutValue('PFO_DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsDateTime);
                TF.PutValue('PFO_DATEFIN',Q.FindField('PSS_DATEFIN').AsDateTime);
                TF.PutValue('PFO_NATUREFORM',Q.FindField('PSS_NATUREFORM').AsString);
                TF.PutValue('PFO_LIEUFORM',Q.FindField('PSS_LIEUFORM').AsString);
                TF.PutValue('PFO_CENTREFORMGU',Q.FindField('PSS_CENTREFORMGU').AsString);//DB2
                TF.PutValue('PFO_DEBUTDJ',Q.FindField('PSS_DEBUTDJ').AsString);
                TF.PutValue('PFO_FINDJ',Q.FindField('PSS_FINDJ').AsString);
                TF.PutValue('PFO_ORGCOLLECT',   -1);  
                TF.PutValue('PFO_ORGCOLLECTGU',Q.FindField('PSS_ORGCOLLECTSGU').AsString);
                If TypeSaisie = 'SAISIEFP' Then
                begin
                        TF.PutValue('PFO_EFFECTUE','X');
                        TF.PutValue('PFO_NBREHEURE',Q.FindField('PSS_DUREESTAGE').AsString);
                        if (DefautHTpsT) or (TypePlan = 'PLF') then
                        begin
                             TF.PutValue('PFO_HTPSTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                             TF.PutValue('PFO_HTPSNONTRAV',0);
                        end
                        else
                        begin
                             TF.PutValue('PFO_HTPSTRAV',0);
                             TF.PutValue('PFO_HTPSNONTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
                        end;
                        
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
		
                        TF.PutValue('PFO_ETATINSCFOR',   'VAL');
                end;
        end;
        Ferme(Q);
        
        //PT3 - Début
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
        //PT3 - Fin
        
        // Si l'update ne se passe pas bien, on rollback sinon la grille ne se rafraîchit pas
        If Not TF.OM.UpdateRecord Then
             TF.cancel
        Else
        Begin
             TF.DisableTOM;
             TF.Post;
             TF.EnableTOM;
             InsererDIFBulletin;
        End;
        
        TF.EndUpdate;
end;

procedure TOF_PGPRESENCEFORM.InsererDIFBulletin;
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
begin
     SupprimerDIF;
     NbhT := TF.GetValue('PFO_HTPSTRAV');
     NbhNonT := TF.GetValue('PFO_HTPSNONTRAV');
     TypeFor := TF.GetValue('PFO_TYPEPLANPREV');
     //PT3
//     Jours := TF.TOBFiltre.GetValue('PSS_JOURSTAGE');
    Jours := 0;
    Q := OpenSQL('SELECT PSS_JOURSTAGE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE="'+LaSession+'" AND PSS_MILLESIME="'+LeMillesime+'"', True);
    If Not Q.EOF Then Jours := Q.FindField('PSS_JOURSTAGE').AsFloat;
    Ferme(Q);

     Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI## PPF_TYPEPLANPREV="'+TypeFor+'"',True);
     TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
     TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
     Ferme(Q);
     Salarie := TF.GetValue('PFO_SALARIE');
     Etab := TF.GetValue('PFO_ETABLISSEMENT');
     DateD := TF.GetValue('PFO_DATEDEBUT');
     DateF := TF.GetValue('PFO_DATEFIN');
     DJ := TF.GetValue('PFO_DEBUTDJ');
     FJ := TF.GetValue('PFO_FINDJ');
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
          If TobMvtDIF.Detail[i].GetValue('PPF_ALIMABS') = 'X' then
          begin
               TypeConge := TobMvtDIF.Detail[i].GetValue('PPF_TYPECONGE');
               AbsOk := PGDIFGenereAbs(HeuresTAlim,HeuresNonTAlim,Jours,Salarie,DJ,FJ,TypeConge,Etab,DateD,DateF);
               If AbsOk = -1 then
               begin
                    PGIBox('Impossible de créer l''absence correspondante car le salarié possède déja une absence sur cette période',Ecran.Caption);
                    Exit;
               end;

          end;
          If TobMvtDIF.Detail[i].GetValue('PPF_ALIMRUB') = 'X' then
          begin
               Rubrique := TobMvtDIF.Detail[i].GetValue('PPF_RUBRIQUE');
               Travail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURETRAV');
               HorsTravail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURENONTRAV');
               TypeAlim := TobMvtDIF.Detail[i].GetValue('PPF_ALIMENT');
               RecupereLesFormationsDIF(Salarie,TypeFor,Rubrique,TypeAlim,TF.GetValue('PFO_DATEPAIE'),Travail = 'X',HorsTravail = 'X');
          end;
     end;
     TobMvtDIF.free;
end;

procedure TOF_PGPRESENCEFORM.SupprimerDIF;
var Q : TQuery;
    TypeFor : String;
    TobMvtDIF : Tob;
    i : Integer;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim,Predef : String;
begin
     TypeFor := TF.GetValue('PFO_TYPEPLANPREV');
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
               PGDIFSuppAbs(TF.GetValue('PFO_SALARIE'),TypeConge,TF.GetValue('PFO_DATEDEBUT'),TF.GetValue('PFO_DATEFIN'));
          end;
          If TobMvtDIF.Detail[i].GetValue('PPF_ALIMRUB') = 'X' then
          begin
               Rubrique := TobMvtDIF.Detail[i].GetValue('PPF_RUBRIQUE');
               TypeAlim := TobMvtDIF.Detail[i].GetValue('PPF_ALIMENT');
               PGDIFSuppRub(TF.GetValue('PFO_SALARIE'),Rubrique,TypeAlim,TF.GetValue('PFO_DATEPAIE'));
          end;
     end;
     TobMvtDIF.free;
end;

Initialization
  registerclasses ( [ TOF_PGPRESENCEFORM ] ) ;
end.
