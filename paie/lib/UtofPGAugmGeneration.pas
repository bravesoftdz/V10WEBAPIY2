{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAUGMINTEGRATION ()
Mots clefs ... : TOF;PGAUGMINTEGRATION
*****************************************************************}
Unit UtofPGAugmGeneration ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms,
     sysutils,                 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTob,
     Ed_Tools,
     UTofPGMul_AugmGenere,
     P5Util,
     EntPaie,
     ParamSoc,
     Vierge,
     UTOF ; 

Type
  TOF_PGAUGMGENERE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    AnneeAugm,WhereSalarie : String;
    LesChamps,WhereSalaireHisto : String;
    DateGeneration : TDateTime;
    procedure CreerLigneSalarie;
    Function RendSalaires(TypeSalaire : String;TS : Tob) : Double;
    Function SalarieExclus (TS,TC : Tob;Fixe,Variable : Double) : String;
    Function Comparaison(Operateur : String;Val1,Val2 : Variant) : Boolean;
    Function WhereSalaires : String;
    Function ChampSalaires(TypeSalaire : String) : String;
  end ;

Implementation

procedure TOF_PGAUGMGENERE.OnUpdate ;
begin
  Inherited ;
            CreerLigneSalarie;
end ;

procedure TOF_PGAUGMGENERE.OnArgument (S : String ) ;
begin
  Inherited ;
  SetControlText('DATEGENERE',DateToStr(Date));
  AnneeAugm := ReadTokenPipe(S,';');
  WhereSalarie := ReadTokenPipe(S,';');
  TFVierge(Ecran).Retour := '';
  SetControlText('MOTIFAUGM','');
  DateGeneration := IDate1900;
end ;

Procedure TOF_PGAUGMGENERE.CreerLigneSalarie;
var TobSal,TobAug,TA,TCritere,TS : Tob;
    i : Integer;
    Q : TQuery;
    Fixe,Variable : Double;
    Salarie : String;
    LaDate : TDateTime;
    AvecComparaison,AvecSalaires : Boolean ;
    ChampFixe,ChampVar : String;
begin
  ChampFixe := ChampSalaires('FIXE');
  ChampVar := ChampSalaires('VARIABLE');
  WhereSalaireHisto := WhereSalaires;
  If ChampFixe <> '' then ChampFixe := '('+ChampFixe+') LEFIXE';
  If ChampVar <> '' then ChampVar := '('+ChampVar+') LEVARIABLE';
  LesChamps := '';
  If ChampFixe <> '' then LesChamps := ','+ChampFixe;
  If ChampVar <> '' then LesChamps := LesChamps+','+ChampVar;
  If GetCheckBoxState('EXCLUSION') = CbChecked then AvecComparaison := True
   else AvecComparaison := False;
   If GetCheckBoxState('SALAIRES') = CbChecked then AvecSalaires := True
   else AvecSalaires := False;
   LaDate := StrToDate(GetControlText('DATEGENERE'));
   DateGeneration := LaDate;
   If AnneeAugm = '' then
   begin
       PGIBox('Vous devez choisir une année',Ecran.Caption);
   end;
   Q := OpenSQL('SELECT * FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+ WhereSalarie,True);
   TobSal := Tob.Create('LesSalaris',Nil,-1);
   TobSal.LoadDetailDB('LesSalaries','','',Q,False);
   Ferme(Q);
   Q := OpenSQL('SELECT * FROM BUDGETPAIE WHERE PBG_ANNEE="'+AnneeAugm+'"',True);
   TobAug :=  Tob.Create('BUDGETPAIE',Nil,-1);
   TobAug.LoadDetailDB('BUDGETPAIE','','',Q,False);
   ferme(Q);
   Q := OpenSQL('SELECT * FROM AUGMPARAM WHERE PAP_ACTIF="X"',True);
   TCritere :=  Tob.Create('Lesexclusions',Nil,-1);
   TCritere.LoadDetailDB('Lesexclusions','','',Q,False);
   ferme(Q);
   if ((ListeAug.nbSelected) > 0) and (not ListeAug.AllSelected) then
   begin
        InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', ListeAug.NbSelected-1,
                    False,True);
         for i := 0 to ListeAug.NbSelected - 1 do
         begin
              ListeAug.GotoLeBookmark(i);
              {$IFDEF EAGLCLIENT}
              TFMAug.Q.TQ.Seek(ListeAug.Row - 1);
              {$ENDIF}
              Salarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
              TS := TobSal.FinDFirst(['PSA_SALARIE'],[Salarie],False);
              If TS = Nil then Continue;
              TA := TobAug.FindFirst(['PBG_SALARIE'],[TS.getValue('PSA_SALARIE')],False);
              If (TA = Nil) or (AvecSalaires) then
              begin
                   If TA = Nil then
                   begin
                        TA := Tob.Create('BUDGETPAIE',TobAug,-1);
                        TA.PutValue('PBG_SALARIE',TS.getValue('PSA_SALARIE'));
                        TA.PutValue('PBG_TYPEBUDG','AUG');
                        TA.PutValue('PBG_NUMORDRE',1);
                        TA.PutValue('PBG_ANNEE',AnneeAugm);
                   end;
                   Fixe := RendSalaires('FIXE',TS);
                   Variable := RendSalaires('VARIABLE',TS);
                   If Fixe = 1 then Fixe := 0;
                   If Variable = 1 then Variable := 0;
                   TA.PutValue('PBG_FIXEAV',Fixe);
                   TA.PutValue('PBG_VARIABLEAV',Variable);
                   TA.PutValue('PBG_FIXEAP',Fixe);
                   TA.PutValue('PBG_VARIABLEAP',Variable);
                  { if TS.getValue('PSE_CODESERVICE') <> Null then Service := TS.getValue('PSE_CODESERVICE')
                   else Service := '';
                   Q := OpenSQL('SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE '+
                   'ON PSO_CODESERVICE=PGS_CODESERVICE '+
                   'LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
                   'WHERE PHO_NIVEAUH=2 AND (PSO_CODESERVICE="'+Service+'" OR PSO_SERVICESUP="'+Service+'")',True);
                   If Not Q.Eof then ResponsValid := Q.FindField('PGS_RESPONSVAR').AsString
                   else ResponsValid := '';
                   Ferme(Q);
                   If ResponsValid = '' then
                   begin
                        Q := OpenSQL('SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE '+
                        'ON PSO_CODESERVICE=PGS_CODESERVICE '+
                        'LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
                        'WHERE PHO_NIVEAUH=1 AND (PSO_CODESERVICE="'+Service+'" OR PSO_SERVICESUP="'+Service+'")',True);
                        If Not Q.Eof then ResponsValid := Q.FindField('PGS_RESPONSVAR').AsString
                        else ResponsValid := '';
                        Ferme(Q);
                   end;         }
                   TA.PutValue('PBG_VALIDERPAR','');
                   If AvecComparaison then TA.PutValue('PBG_ETATINTAUGM',SalarieExclus (TS,TCritere,Fixe,Variable))
                   else TA.PutValue('PBG_ETATINTAUGM','001');
                   If (TA.GetValue('PBG_ETATINTAUGM') <> '011') and (TA.GetValue('PBG_ETATINTAUGM') <> '012') then TA.PutValue('PBG_MOTIFAUGM',GetControlText('MOTIFAUGM'))
                   else TA.PutValue('PBG_MOTIFAUGM','');
                   TA.PutValue('PBG_DATECREATION',LaDate);
                   TA.PutValue('PBG_ETABLISSEMENT',TS.getValue('PSA_ETABLISSEMENT'));
                   TA.PutValue('PBG_LIBELLE',TS.getValue('PSA_LIBELLE'));
                   TA.PutValue('PBG_TRAVAILN1',TS.getValue('PSA_TRAVAILN1'));
                   TA.PutValue('PBG_TRAVAILN2',TS.getValue('PSA_TRAVAILN2'));
                   TA.PutValue('PBG_TRAVAILN3',TS.getValue('PSA_TRAVAILN3'));
                   TA.PutValue('PBG_TRAVAILN4',TS.getValue('PSA_TRAVAILN4'));
                   TA.PutValue('PBG_CODESTAT',TS.getValue('PSA_CODESTAT'));
                   TA.PutValue('PBG_LIBREPCMB1',TS.getValue('PSA_LIBREPCMB1'));
                   TA.PutValue('PBG_LIBREPCMB2',TS.getValue('PSA_LIBREPCMB2'));
                   TA.PutValue('PBG_LIBREPCMB3',TS.getValue('PSA_LIBREPCMB3'));
                   TA.PutValue('PBG_LIBREPCMB4',TS.getValue('PSA_LIBREPCMB4'));
                   TA.PutValue('PBG_CONDEMPLOI',TS.getValue('PSA_CONDEMPLOI'));
                   TA.PutValue('PBG_DADSPROF',TS.getValue('PSA_DADSPROF'));
                   TA.PutValue('PBG_LIBELLEEMPLOI',TS.getValue('PSA_LIBELLEEMPLOI'));
                   TA.PutValue('PBG_SALAIREMOIS1',TS.getValue('PSA_SALAIREMOIS1'));
                   TA.PutValue('PBG_SALAIREMOIS2',TS.getValue('PSA_SALAIREMOIS2'));
                   TA.PutValue('PBG_SALAIREMOIS3',TS.getValue('PSA_SALAIREMOIS3'));
                   TA.PutValue('PBG_SALAIREMOIS4',TS.getValue('PSA_SALAIREMOIS4'));
                   TA.PutValue('PBG_SALAIREMOIS5',TS.getValue('PSA_SALAIREMOIS5'));
                   TA.PutValue('PBG_SALAIRANN1',TS.getValue('PSA_SALAIRANN1'));
                   TA.PutValue('PBG_SALAIRANN2',TS.getValue('PSA_SALAIRANN2'));
                   TA.PutValue('PBG_SALAIRANN3',TS.getValue('PSA_SALAIRANN3'));
                   TA.PutValue('PBG_SALAIRANN4',TS.getValue('PSA_SALAIRANN4'));
                   TA.PutValue('PBG_SALAIRANN5',TS.getValue('PSA_SALAIRANN5'));
                   TA.PutValue('PBG_HORAIREMOIS',TS.getValue('PSA_HORAIREMOIS'));
                   TA.PutValue('PBG_UNITEPRISEFF',TS.getValue('PSA_UNITEPRISEFF'));
                   TA.PutValue('PBG_DADSEXOBASE',TS.getValue('PSA_DADSEXOBASE'));
                   TA.PutValue('PBG_CONFIDENTIEL',TS.getValue('PSA_CONFIDENTIEL'));
                   TA.PutValue('PBG_PCTVARIABLE',0);
                   TA.PutValue('PBG_PCTFIXE',0);
                   TA.PutValue('PBG_PCTFIXERESP',0);
                   TA.PutValue('PBG_PCTVARRESP',0);
                   TA.PutValue('PBG_PCTFIXESAISIE',0);
                   TA.PutValue('PBG_PCTVARSAISIE',0);
                   TA.PutValue('PBG_INTFIXE','-');
                   TA.PutValue('PBG_INTVARIABLE','-');
                   TA.PutValue('PBG_RESPONSVAR',TS.getValue('PSE_RESPONSVAR'));
                   TA.PutValue('DATEFIXE',IDate1900);
                   TA.PutValue('DATEVAR',IDate1900);
                   TA.InsertOrUpdateDB(False);
              end;
              MoveCurProgressForm ('Salarié : '+
                             TS.GetValue('PSA_LIBELLE'));
         end;
   end;
   if (ListeAug.AllSelected = TRUE) then
   begin
         {$IFDEF EAGLCLIENT}
   if (TFMAug.bSelectAll.Down) then
    TFMAug.Fetchlestous;
  {$ENDIF}
        Q_MulAug.First;
        InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', Q_MulAug.RecordCount,
                    False,True);
        while not Q_MulAug.EOF do
        begin
              Salarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
              TS := TobSal.FinDFirst(['PSA_SALARIE'],[Salarie],False);
              If TS = Nil then continue;
              TA := TobAug.FindFirst(['PBG_SALARIE'],[Salarie],False);
              If (TA = Nil) or (AvecSalaires) then
              begin
                   If TA = Nil then
                   begin
                        TA := Tob.Create('BUDGETPAIE',TobAug,-1);
                        TA.PutValue('PBG_SALARIE',TS.getValue('PSA_SALARIE'));
                        TA.PutValue('PBG_TYPEBUDG','AUG');
                        TA.PutValue('PBG_NUMORDRE',1);
                        TA.PutValue('PBG_ANNEE',AnneeAugm);
                   end;
                   Fixe := RendSalaires('FIXE',TS);
                   Variable := RendSalaires('VARIABLE',TS);
                   If Fixe = 1 then Fixe := 0;
                   If Variable = 1 then Variable := 0;
                   TA.PutValue('PBG_FIXEAV',Fixe);
                   TA.PutValue('PBG_VARIABLEAV',Variable);
                   TA.PutValue('PBG_FIXEAP',Fixe);
                   TA.PutValue('PBG_VARIABLEAP',Variable);
                   If AvecComparaison then TA.PutValue('PBG_ETATINTAUGM',SalarieExclus (TS,TCritere,Fixe,Variable))
                   else TA.PutValue('PBG_ETATINTAUGM','001');
                   If (TA.GetValue('PBG_ETATINTAUGM') <> '011') and (TA.GetValue('PBG_ETATINTAUGM') <> '012') then TA.PutValue('PBG_MOTIFAUGM',GetControlText('MOTIFAUGM'))
                   else TA.PutValue('PBG_MOTIFAUGM','');
                   TA.PutValue('PBG_DATECREATION',LaDate);
{                   if TS.getValue('PSE_CODESERVICE') <> Null then Service := TS.getValue('PSE_CODESERVICE')
                   else Service := '';
                   Q := OpenSQL('SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE '+
                   'ON PSO_CODESERVICE=PGS_CODESERVICE '+
                   'LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
                   'WHERE PHO_NIVEAUH=2 AND (PSO_CODESERVICE="'+Service+'" OR PSO_SERVICESUP="'+Service+'")',True);
                   If Not Q.Eof then ResponsValid := Q.FindField('PGS_RESPONSVAR').AsString
                   else ResponsValid := '';
                   Ferme(Q);
                   If ResponsValid = '' then
                   begin
                        Q := OpenSQL('SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE '+
                        'ON PSO_CODESERVICE=PGS_CODESERVICE '+
                        'LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
                        'WHERE PHO_NIVEAUH=1 AND (PSO_CODESERVICE="'+Service+'" OR PSO_SERVICESUP="'+Service+'")',True);
                        If Not Q.Eof then ResponsValid := Q.FindField('PGS_RESPONSVAR').AsString
                        else ResponsValid := '';
                        Ferme(Q);
                   end;}
                   TA.PutValue('PBG_VALIDERPAR','');
                   TA.PutValue('PBG_ETABLISSEMENT',TS.getValue('PSA_ETABLISSEMENT'));
                   TA.PutValue('PBG_LIBELLE',TS.getValue('PSA_LIBELLE'));
                   TA.PutValue('PBG_TRAVAILN1',TS.getValue('PSA_TRAVAILN1'));
                   TA.PutValue('PBG_TRAVAILN2',TS.getValue('PSA_TRAVAILN2'));
                   TA.PutValue('PBG_TRAVAILN3',TS.getValue('PSA_TRAVAILN3'));
                   TA.PutValue('PBG_TRAVAILN4',TS.getValue('PSA_TRAVAILN4'));
                   TA.PutValue('PBG_CODESTAT',TS.getValue('PSA_CODESTAT'));
                   TA.PutValue('PBG_LIBREPCMB1',TS.getValue('PSA_LIBREPCMB1'));
                   TA.PutValue('PBG_LIBREPCMB2',TS.getValue('PSA_LIBREPCMB2'));
                   TA.PutValue('PBG_LIBREPCMB3',TS.getValue('PSA_LIBREPCMB3'));
                   TA.PutValue('PBG_LIBREPCMB4',TS.getValue('PSA_LIBREPCMB4'));
                   TA.PutValue('PBG_CONDEMPLOI',TS.getValue('PSA_CONDEMPLOI'));
                   TA.PutValue('PBG_DADSPROF',TS.getValue('PSA_DADSPROF'));
                   TA.PutValue('PBG_LIBELLEEMPLOI',TS.getValue('PSA_LIBELLEEMPLOI'));
                   TA.PutValue('PBG_SALAIREMOIS1',TS.getValue('PSA_SALAIREMOIS1'));
                   TA.PutValue('PBG_SALAIREMOIS2',TS.getValue('PSA_SALAIREMOIS2'));
                   TA.PutValue('PBG_SALAIREMOIS3',TS.getValue('PSA_SALAIREMOIS3'));
                   TA.PutValue('PBG_SALAIREMOIS4',TS.getValue('PSA_SALAIREMOIS4'));
                   TA.PutValue('PBG_SALAIREMOIS5',TS.getValue('PSA_SALAIREMOIS5'));
                   TA.PutValue('PBG_SALAIRANN1',TS.getValue('PSA_SALAIRANN1'));
                   TA.PutValue('PBG_SALAIRANN2',TS.getValue('PSA_SALAIRANN2'));
                   TA.PutValue('PBG_SALAIRANN3',TS.getValue('PSA_SALAIRANN3'));
                   TA.PutValue('PBG_SALAIRANN4',TS.getValue('PSA_SALAIRANN4'));
                   TA.PutValue('PBG_SALAIRANN5',TS.getValue('PSA_SALAIRANN5'));
                   TA.PutValue('PBG_HORAIREMOIS',TS.getValue('PSA_HORAIREMOIS'));
                   TA.PutValue('PBG_UNITEPRISEFF',TS.getValue('PSA_UNITEPRISEFF'));
                   TA.PutValue('PBG_RESPONSVAR',TS.getValue('PSE_RESPONSVAR'));
                   TA.PutValue('PBG_DADSEXOBASE',TS.getValue('PSA_DADSEXOBASE'));
                   TA.PutValue('PBG_CONFIDENTIEL',TS.getValue('PSA_CONFIDENTIEL'));
                   TA.PutValue('PBG_PCTVARIABLE',0);
                   TA.PutValue('PBG_PCTFIXE',0);
                   TA.InsertOrUpdateDB(False);
              end;
              MoveCurProgressForm ('Salarié : '+
                             TS.GetValue('PSA_LIBELLE'));
              Q_MulAug.Next;
         end;
   end;
   TobAug.Free;
   TobSal.Free;
   TCritere.Free;
   FiniMoveProgressForm;
   TFVierge(Ecran).Retour := 'Maj';
end;

Function TOF_PGAUGMGENERE.RendSalaires(TypeSalaire : String;TS : Tob) : Double;
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
          If Champ <> '' then
          begin
               Montant := Montant + TS.GetDouble(champ);
          end;
     end;
     Result := Montant;
end;

Function TOF_PGAUGMGENERE.SalarieExclus (TS,TC : Tob;Fixe,Variable : Double) : String;
var Anciennete : Word;
    i,a : Integer;
    LeCritere,LaDonnee,Valeur,Operateur,NouvelEtat,LeChamp,ValChamp : String;
    DE : TDateTime;
    Q : TQuery;
    TobDetail : Tob;
    VerifCritere : Boolean;
    TypeContrat : String;
    JourAbs : Double;
    DateAbs : TDateTime;
    HoraireMens : Double;
    DateHisto : TDateTime;
    PremMois, PremAnnee,NbMoisAug: WORD;
begin
     Result := '001';
     PremMois := 0;
     PremAnnee := 0;
     For i := 0 to TC.detail.Count - 1 do
     begin
          LeCritere := TC.Detail[i].GetValue('PAP_CRITEREAUGM');
          NouvelEtat := TC.Detail[i].GetValue('PAP_ETATINTAUGM');
          Q := OpenSQL('SELECT * FROM AUGMEXCLUS WHERE PAE_CRITEREAUGM="'+LeCritere+'"',True);
          TobDetail := Tob.Create('LeDetail',Nil,-1);
          TobDetail.LoadDetailDB('LeDetail','','',Q,False);
          Ferme(Q);
          VerifCritere := True;
          if TobDetail.detail.Count = 0 then VerifCritere := False;
          For a := 0 to TobDetail.detail.Count - 1 do
          begin
               Operateur := Tobdetail.Detail[a].GetValue('PAE_OPERATTEST');
               Valeur := Tobdetail.Detail[a].GetValue('PAE_VALEUREX');
               LaDonnee := TobDetail.Detail[a].GetValue('PAE_CRITEREAUGEX');
               Q := OpenSQL('SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE="PGM" AND CO_CODE="'+LaDonnee+'"',True);
               If Not Q.Eof then LeChamp := Q.FindField('CO_LIBRE').AsString
               else LeChamp := '';
               Ferme(Q);
               //Durée d'absence
               If LaDonnee = 'ABS' then
               begin
                    DateAbs := PlusMois(dateGeneration, -12);
                    Q := OpenSQL('SELECT SUM(PCN_JOURS) NBJOURS FROM ABSENCESALARIE WHERE PCN_TYPEMVT="ABS" '+
                    'AND PCN_SALARIE="'+TS.GetValue('PSA_SALARIE')+'" AND PCN_DATEFINABS>="'+UsDateTime(DateAbs)+'"',True);
                    If Not Q.Eof then JourAbs := Q.FindField('NBJOURS').AsFloat
                    Else JourAbs := 0;
                    Ferme(Q);
                    If not Comparaison(Operateur,JourAbs,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Ancienneté
               else If (LaDonnee = 'ANC') or (LaDonnee = 'ANR') then
               begin
                    DE := TS.Getvalue('PSA_DATEANCIENNETE');
                    Anciennete := AncienneteMois(DE, dateGeneration);
                    If LaDonnee = 'ANR' then Anciennete := Anciennete + TS.Getvalue('PSA_REGULANCIEN');
                    If not Comparaison(Operateur,Anciennete,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Dernière augmentation
               else If (LaDonnee = 'DAU') or (LaDonnee = 'DAG') then
               begin
                    Q := OpenSQL('SELECT Max(PHS_DATEAPPLIC) DATEHISTO FROM HISTOSALARIE WHERE '+
                    'PHS_SALARIE="'+TS.GetValue('PSA_SALARIE')+'"'+WhereSalaireHisto,True);
                    If Not Q.Eof then DateHisto := Q.FindField('DATEHISTO').AsDateTime
                    Else DateHisto := IDate1900;
                    Ferme(Q);
                    AglNombreDeMoisComplet(DateHisto, DateGeneration, PremMois, PremAnnee, NbMoisAug);
                    If not Comparaison(Operateur,NbMoisAug,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Motif de sortie
               else If LaDonnee = 'MOS' then
               begin
                     If TS.FieldExists(LeChamp) then
                     begin
                          ValChamp := TS.GetValue(LeChamp);
                          If not Comparaison(Operateur,ValChamp,Valeur) then
                          begin
                               VerifCritere := False;
                               Break;
                          end;
                     end;
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
                    HoraireMens := TS.GetValue('PSA_HORAIREMOIS');
                    If not Comparaison(Operateur,HoraireMens,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Type de contrat
               else If LaDonnee = 'TYC' then
               begin
                    Q := OpenSQL('SELECT PCI_TYPECONTRAT FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+TS.Getvalue('PSA_SALARIE')+'" '+
                    'AND (PCI_FINCONTRAT>="'+UsDateTime(DateGeneration)+'" OR PCI_FINCONTRAT IS NULL OR PCI_FINCONTRAT="'+UsDateTime(IDate1900)+'") '+
                    'ORDER BY PCI_DEBUTCONTRAT DESC',True);
                    If Not Q.Eof then TypeContrat := Q.Fields[0].AsString
                    else TypeContrat := '';
                    Ferme(Q);
                    If not Comparaison(Operateur,TypeContrat,Valeur) then
                    begin
                      VerifCritere := False;
                      Break;
                    end;
               end
               //Champs fiche salarié
                else
                begin
                     If TS.FieldExists(LeChamp) then
                     begin
                          ValChamp := TS.GetValue(LeChamp);
                          If not Comparaison(Operateur,ValChamp,Valeur) then
                          begin
                               VerifCritere := False;
                               Break;
                          end;
                     end;
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

Function TOF_PGAUGMGENERE.Comparaison(Operateur : String;Val1,Val2 : Variant) : Boolean;
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

Function TOF_PGAUGMGENERE.WhereSalaires : String;
var LesSalairesF,LesSalairesV,TempSalaire,Champ,Requete : String;
    i : Integer;
begin
     LesSalairesF := GetParamSoc('SO_PGAUGFIXE');
     LesSalairesV := GetParamSoc('SO_PGAUGVARIABLE');
     If (LesSalairesF = '') and (LesSalairesV = '')  then
     begin
          Result := '';
          Exit;
     end;
     Requete := '';
     While LesSalairesF <> '' do
     begin
          Champ := '';
          TempSalaire := ReadTokenPipe(LesSalairesF,';');
          For i := 1 to 5 do
          begin
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PHS_BSALAIREANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PHS_BSALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then
          begin
               If Requete <> '' then Requete := Requete + ' OR '+Champ+'="X"'
               else Requete := ' AND ('+Champ+'="X"';
          end;
     end;
     While LesSalairesV <> '' do
     begin
          Champ := '';
          TempSalaire := ReadTokenPipe(LesSalairesV,';');
          For i := 1 to 5 do
          begin
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PHS_BSALAIREANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PHS_BSALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then
          begin
               If Requete <> '' then Requete := Requete + ' OR '+Champ+'="X"'
               else Requete := ' AND ('+Champ+'="X"';
          end;
     end;
     If Requete <> '' then Requete := Requete+')';
     Result := Requete;
end;

Function TOF_PGAUGMGENERE.ChampSalaires(TypeSalaire : String) : String;
var LesSalaires,TempSalaire,Champ,Requete : String;
    i : Integer;
begin
     LesSalaires := '';
     If TypeSalaire = 'FIXE' then LesSalaires := GetParamSoc('SO_PGAUGFIXE');
     If TypeSalaire = 'VARIABLE' then LesSalaires := GetParamSoc('SO_PGAUGVARIABLE');
     If LesSalaires = '' then
     begin
          Result := '';
          Exit;
     end;
     Requete := '';
     While LesSalaires <> '' do
     begin
          Champ := '';
          TempSalaire := ReadTokenPipe(LesSalaires,';');
          For i := 1 to 5 do
          begin
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PHS_SALAIREANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PHS_SALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then
          begin
               If Requete <> '' then Requete := Requete + ' + '+Champ
               else Requete := Champ;
          end;
     end;
     Result := Requete;
end;

Initialization
  registerclasses ( [ TOF_PGAUGMGENERE ] ) ;
end.

