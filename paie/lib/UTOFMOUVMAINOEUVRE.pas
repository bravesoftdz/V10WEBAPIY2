{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/06/2001
Modifié le ... : 05/10/2001
Description .. : Source TOF Edition de main d'oeuvre
Mots clefs ... : PAIE;MOUVMAINOEUVRE
*****************************************************************
PT- 1 : 10/10/2001 : JL 562: Procédures de vérification des dates + calendrier
PT- 2 : 08/11/2001 : JL 562: exclusion des contrats CTT et des stagiaires dans toutes les requêtes
PT- 3 : 30/11/2001 : JL 569: La procedure VerifDateMois est maintenant appelée ds la procedure Effectif
PT- 4 : 30/11/2001 : JL 569: L'affichage des dates est gérés par une nouvelle procédure appelée sur le OnChange de la date + procedure effectif sur le OnExit (aulieu du Onchange)
PT- 5 : 30/11/2001 : JL 569: Gestion des 2 états : contrôle du fait qu'il y ait eu des entrées/sorties ou non
PT- 6 : 27/03/2002 : JL 571: Modification des requêtes pour prendre en compte les changements d'établissement
PT- 7 : 02/05/2002 : JL 571: Modification du calcul de l'effectif à la fin du mois précédent
                             et récupération établissement pour construire la tob ds l'état
PT- 8 ! 13/08/2002 : JL 582: Fiche bug n° 10182 calcul intérimaires
                        V_42 : FQ 10545 ajout contrôle établissement    + 17/09/2003 ajout PEI_FINEMPLOI=IDATE1900 au cas ou date fin non renseigné
PT- 9 : 18/03/2003 : JL V_42 : correction nombre de sortie toujours a 0 a cause portage CWAS
PT- 10 : 16/04/2003 : JL V_42 : Correction requête pour oracle (champ DADSPROF="" non pris en compte)
PT11    28/04/2004 SB V_50 FQ 10812 Intégration de la gestion des déclarants
PT12    17/07/2004 JL V_50 FQ 11110 Ajout PSA_PRISEFFECTIF dans clause
---- PH 10/08/2005 Suppression directive de compil $IFDEF AGL550B ----
PT13 : 16/11/2006 JL V_70 Modification complètre des procèdure OnUpdate et effectif pour FQ 13594 et 12702 (prise en compte contrat + etab complémentaire)
PT14 : 06/02/2007 FC V_80 Mise en place filtrage des habilitations/populations
PT15 : 23/02/2007 JL V_70 FQ 13943 Corrections requete effectif et entrée
PT16 : 12/03/2007 JL V_70 Correction affichage nationalité
}
Unit UTOFMOUVMAINOEUVRE ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,QRS1,FE_Main,
{$ELSE}
     eQRS1,UtileAGl,MaineAgl,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
     UTOB,
     ParamSoc,Pgoutils, HTB97,PGEdtEtat,
     P5Def //PT14
     ;

Type
  TOF_MOUVMAINOEUVRE = Class (TOF)

     BVALIDER: TToolbarButton97;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TobEtat :Tob;
    procedure Effectif(Sender : Tobject);
    procedure Precedent(Sender : Tobject);
    procedure VerifDate(Sender : Tobject);
    procedure DateElipsisclick(Sender: TObject);
    procedure Affichelibelle(Sender : TObject);// PT- 4
    procedure ExitEffectifDebut(Sender : TObject);  // PT- 7
    procedure EtatAvecTob(Sender:TObject);
    procedure AffichDeclarant(Sender:TObject);  { PT11 }
    end ;

Implementation

// Confection de la requete passée au générateur
procedure TOF_MOUVMAINOEUVRE.OnClose ;
begin
   Inherited ;
   If TobEtat <> Nil then FreeAndNil(TobEtat);
end;

procedure TOF_MOUVMAINOEUVRE.OnUpdate ;
var
   Datedeb,datefin,DateEff,MDateEff,DateDeb2 : TDateTime;
   Mois,SQL,E,orderby,SQLEtab,orderbyEtab : String;
   TOT1S,TOT2S,ENS,SOS,HOS,FES : String;
   TOT1,TOT2,EN,SO,HO,FE,i : integer;
   Tobcontrat,T,TobSal : Tob;
   Q : TQuery;
   MotifSortie,MotifEntree,Nationalite : String;
   EtabCompl,Etab,WhereEtab : String;
   WhereEtabPaie2,WhereEtabPaie1,WhereEtabSal : String;
   st, Etablis, Habilitation: string;  //PT14
   j: Integer;                         //PT14
begin
Inherited ;
        If TobEtat <> Nil then FreeAndNil(TobEtat);
        TobEtat := Tob.Create('Edition',Nil,-1);
        PGGlbEtabMMO := GetControlText('ETAB');             // PT- 7
        TOT1S := GetControlText('XX_EFFTOTAL');
        TOT2S := GetControlText('XX_EFFFINAL');
        HOS := GetControlText('XX_H');
        FES := GetControlText('XX_F');
        ENS := GetControlText('XX_NBENTREE');
        SOS := GetControlText('XX_NBSORTIES');
        IF (TOT1S='') OR (TOT1S='0') then TOT1 := 0
        Else TOT1 := StrToInt(TOT1S);
        IF (TOT2S='') OR (TOT2S='0') then TOT2 := 0
        Else TOT2 := StrToInt(TOT2S);
        IF (HOS='') OR (HOS='0') then HO := 0
        Else HO := StrToInt(HOS);
        IF (FES='') OR (FES='0') then FE := 0
        Else FE := StrToInt(FES);
        IF (ENS='') OR (ENS='0') then EN := 0
        Else EN := StrToInt(ENS);
        IF (SOS='') OR (SOS='0') then SO := 0
        Else SO := StrToInt(SOS);
        if TOT2 <> (HO + FE) then
        begin
                LastError := 1;
                LastErrorMsg := 'Attention l''effectif total est différent de la somme du nombre d''hommes et de femmes';
        end;
        if TOT2 <> (TOT1 + (EN - SO)) then
        begin
                LastError := 1;
                LastErrorMsg := 'Attention l''effectif total est différent de la somme entre celui du mois précédent + entrées - sorties';
        end;
        DateDeb := idate1900; DateFin := idate1900;
        DateDeb2 := IDate1900;
        Mois := GetControlText('MOIS');
        if IsValidDate(Mois) then
        begin
                DateDeb := DebutDeMois(StrToDate(Mois));
                SetControlText('XX_VARIABLEDEB',DateToStr(DateDeb));
                DateFin := FinDeMois(StrToDate(Mois));
                SetControlText('XX_VARIABLEFIN',DateToStr(DateFin));
                MDateEff := PlusMois(StrToDATE(Mois),-1);
                DateEff := FinDeMois(MDateEff);
                DateDeb2 := DebutDeMois(DateEff);
                SetControlText('XX_VARIABLEEFF',DateToStr(DateEff));
        end;
        E := GetControlText('ETAB');
        EtabCompl := GetControlText('ETABCOMPL');
        Etab := ReadTokenPipe(EtabCompl,';');
        While Etab <> '' do
        begin
             WhereEtab := WhereEtab + ' OR PCI_ETABLISSEMENT="'+Etab+'"';
             Etab := ReadTokenPipe(EtabCompl,';');
        end;
        WhereEtab := '(PCI_ETABLISSEMENT="'+E+'"' + WhereEtab+')';

        //DEB PT14
        Habilitation := '';
        if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
          Habilitation := ' AND PCI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ') ';
        //FIN PT14

        Q := OpenSQL('SELECT PCI_MOTIFSORTIE,PSA_LIBELLE,PSA_PRENOM,PSA_NATIONALITE,PSA_SEXE,PSA_DATENAISSANCE,PSA_CODEEMPLOI,'+
        'PCI_SALARIE,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,PCI_TYPECONTRAT,PCI_MOTIFSORTIE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
        WhereEtab + Habilitation + ' AND ((PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'") OR '+ {PT14}
        '(PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_FINCONTRAT<="'+UsDateTime(DateFin)+'"))',True);
        Tobcontrat := Tob.Create('LesContrats',Nil,-1);
        Tobcontrat.LoadDetailDB('SALARIES','','',Q,False);
        For i := 0 to TobContrat.Detail.Count  - 1  do
        begin
             T := Tob.Create('FilleEtat',TobEtat,-1);
             T.AddChampSupValeur('PSA_SALARIE',TobContrat.Detail[i].GetValue('PCI_SALARIE'));
             T.AddChampSupValeur('LIBELLE',TobContrat.Detail[i].GetValue('PSA_LIBELLE'));
             T.AddChampSupValeur('PRENOM',TobContrat.Detail[i].GetValue('PSA_PRENOM'));
             Nationalite := TobContrat.Detail[i].GetValue('PSA_NATIONALITE'); //PT16
             If Nationalite = 'FRA' then T.AddChampSupValeur('NATIONALITE','F')
             else
             begin
              Q := OpenSQL('SELECT PY_MEMBRECEE FROM PAYS WHERE PY_PAYS="'+Nationalite+'"',True);
              If Not Q.Eof then
              begin
                If Q.FindField('PY_MEMBRECEE').AsString = 'X' then T.AddChampSupValeur('NATIONALITE','C')
                Else T.AddChampSupValeur('NATIONALITE','A');
              end
              else T.AddChampSupValeur('NATIONALITE','A');
               Ferme(Q);
             end;

             T.AddChampSupValeur('DATENAISSANCE',DateToStr(TobContrat.Detail[i].GetValue('PSA_DATENAISSANCE')));
             T.AddChampSupValeur('SEXE',TobContrat.Detail[i].GetValue('PSA_SEXE'));
             T.AddChampSupValeur('LIBEMPLOI',RechDom('PGCODEPCSESE',TobContrat.Detail[i].GetValue('PSA_CODEEMPLOI'),False));
             T.AddChampSupValeur('CODEEMPLOI',TobContrat.Detail[i].GetValue('PSA_CODEEMPLOI'));
             T.AddChampSupValeur('DATEENTREE',DateToStr(TobContrat.Detail[i].GetValue('PCI_DEBUTCONTRAT')));

             If TobContrat.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CCD' then MotifEntree := 'RD'
             else If TobContrat.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CAA' then MotifEntree := 'RD'
             else If TobContrat.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CAC' then MotifEntree := 'RD'
             else If TobContrat.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CAP' then MotifEntree := 'RD'
             else MotifEntree := 'RI';
             T.AddChampSupValeur('MOTIFENTREE',MotifEntree);
             T.AddChampSupValeur('DATESORTIE',DateToStr(TobContrat.Detail[i].GetValue('PCI_FINCONTRAT')));
             Q := OpenSQL('SELECT PMS_MAINOEUVRE FROM MOTIFSORTIEPAY WHERE PMS_CODE="'+TobContrat.Detail[i].GetValue('PCI_MOTIFSORTIE')+'"', True);
             If Not Q.Eof then MotifSortie := Q.FindField('PMS_MAINOEUVRE').AsString;
             Ferme(Q);
             T.AddChampSupValeur('MOTIFSORTIE',MotifSortie);
        end;
        Ferme(Q);
        EtabCompl := GetControlText('ETABCOMPL');
        Etab := ReadTokenPipe(EtabCompl,';');
        While Etab <> '' do
        begin
             WhereEtabPaie1 := WhereEtabPaie1 + ' OR P1.PPU_ETABLISSEMENT="'+Etab+'"';
             WhereEtabPaie2 := WhereEtabPaie2 + ' OR P2.PPU_ETABLISSEMENT="'+Etab+'"';
             WhereEtabSal := WhereEtabSal + ' OR PSA_ETABLISSEMENT="'+Etab+'"';
             Etab := ReadTokenPipe(EtabCompl,';');
        end;
        WhereEtabSal := '(PSA_ETABLISSEMENT="'+E+'"' + WhereEtabSal+')';
        WhereEtabPaie1 := '(P1.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie1+')';
        WhereEtabPaie2 := '(P2.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie2+')';

        //DEB PT14
        Habilitation := '';
        if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
          Habilitation := ' AND (' + MonHabilitation.LeSQL + ')';
        //FIN PT14

        SQL := 'SELECT PSA_SEXE,PSA_CODEEMPLOI,PSA_SALARIE,PSA_DATESORTIE,PSA_DATEENTREE, PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE'+  //PT6
        ',PSA_NATIONALITE,PSA_MOTIFSORTIE,PSA_MOTIFENTREE FROM SALARIES left join PAIEENCOURS P1 on psa_SALARIE=P1.PPU_SALARIE'+
        ' AND P1.PPU_DATEDEBUT="'+UsDateTime(DateDeb2)+'"'+
        ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="") AND'+    //PT- 10 PT12
        ' (('+WhereEtabSal+' AND ((PSA_DATESORTIE >="'+UsDateTime(Datedeb)+'" AND PSA_DATESORTIE <="'+UsDateTime(DateFin)+'")'+
        ' OR (PSA_DATEENTREE >="'+UsDateTime(Datedeb)+'" AND  PSA_DATEENTREE <="'+UsDateTime(DateFin)+'")))'+
        ' OR (Exists (SELECT PPU_SALARIE FROM PAIEENCOURS P2 WHERE P2.PPU_SALARIE=PSA_SALARIE AND'+
        ' P2.PPU_DATEDEBUT="'+UsDateTime(Datedeb)+'" AND P1.PPU_ETABLISSEMENT<>P2.PPU_ETABLISSEMENT'+
        ' AND ('+WhereEtabPaie1+' OR '+WhereEtabPaie2+'))))'+ Habilitation + {PT14}
        ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
        WhereEtab+' AND ((PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'") OR '+
        '(PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_FINCONTRAT<="'+UsDateTime(DateFin)+'")))'+
        ' GROUP BY PSA_SEXE,PSA_CODEEMPLOI,PSA_SALARIE,PSA_DATESORTIE,PSA_DATEENTREE, PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE,PSA_NATIONALITE,PSA_MOTIFSORTIE,PSA_MOTIFENTREE';
        orderby := ' ORDER BY PSA_SALARIE' ;

        Q := OpenSQL(SQL+orderby,True);
        TobSal := Tob.Create('Lessalaries',Nil,-1);
        TobSal.LoadDetailDB('Lessalaries','','',Q,False);
        For i := 0 to TobSal.Detail.Count - 1 do
        begin
             T := Tob.Create('FilleEtat',TobEtat,-1);
             T.AddChampSupValeur('PSA_SALARIE',TobSal.Detail[i].GetValue('PSA_SALARIE'));
             T.AddChampSupValeur('LIBELLE',TobSal.Detail[i].GetValue('PSA_LIBELLE'));
             T.AddChampSupValeur('PRENOM',TobSal.Detail[i].GetValue('PSA_PRENOM'));
             Nationalite := TobSal.Detail[i].GetValue('PSA_NATIONALITE'); //PT16
             If Nationalite = 'FRA' then T.AddChampSupValeur('NATIONALITE','F')
             else
             begin
               Q := OpenSQL('SELECT PY_MEMBRECEE FROM PAYS WHERE PY_PAYS="'+TobSal.Detail[i].GetValue('PSA_NATIONALITE')+'"',True);
               If Not Q.Eof then
               begin
                  If Q.FindField('PY_MEMBRECEE').AsString = 'X' then T.AddChampSupValeur('NATIONALITE','C')
                  Else T.AddChampSupValeur('NATIONALITE','A');
               end
             Else T.AddChampSupValeur('NATIONALITE','A');
             Ferme(Q);
             end;
             T.AddChampSupValeur('DATENAISSANCE',DateToStr(TobSal.Detail[i].GetValue('PSA_DATENAISSANCE')));
             T.AddChampSupValeur('SEXE',TobSal.Detail[i].GetValue('PSA_SEXE'));
             T.AddChampSupValeur('LIBEMPLOI',RechDom('PGCODEPCSESE',TobSal.Detail[i].GetValue('PSA_CODEEMPLOI'),False));
             T.AddChampSupValeur('CODEEMPLOI',TobSal.Detail[i].GetValue('PSA_CODEEMPLOI'));
             T.AddChampSupValeur('DATEENTREE',DateToStr(TobSal.Detail[i].GetValue('PSA_DATEENTREE')));
             T.AddChampSupValeur('MOTIFENTREE','RI');
             If TobSal.Detail[i].GetValue('PSA_DATESORTIE') <> IDate1900 then T.AddChampSupValeur('DATESORTIE',DateToStr(TobSal.Detail[i].GetValue('PSA_DATESORTIE')))
             else T.AddChampSupValeur('DATESORTIE','');
             Q := OpenSQL('SELECT PMS_MAINOEUVRE FROM MOTIFSORTIEPAY WHERE PMS_CODE="'+TobSal.Detail[i].GetValue('PSA_MOTIFSORTIE')+'"', True);
             If Not Q.Eof then MotifSortie := Q.FindField('PMS_MAINOEUVRE').AsString;
             Ferme(Q);
             T.AddChampSupValeur('MOTIFSORTIE',MotifSortie);
        end;
        Ferme(Q);
         TobEtat.Detail.Sort('PSA_DATEENTREE');

        //DEB PT14
        Habilitation := '';
        if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
        begin // Champs etablissement seul comme critere de la population
          st := MonHabilitation.LesEtab;
          Etablis := ReadTokenSt(St);
          j := 0;
          while Etablis <> '' do
          begin
            j := j + 1;
            if Etablis <> '' then
            begin
              if j > 1 then Habilitation := Habilitation + ',';
              Habilitation := Habilitation + '"' + Etablis + '"';
            end;
            Etablis := ReadTokenSt(St);
          end;
          if j > 0 then
            Habilitation := ' AND ET_ETABLISSEMENT IN (' + Habilitation + ')';
        end;
        //FIN PT14

        SQLEtab := 'SELECT ET_ETABLISSEMENT,ET_SIRET,ET_APE,ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE'+
        ' ,ANN_APNOM,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE'+
        ' FROM ETABLISS '+
        ' left join ETABCOMPL on ET_ETABLISSEMENT=ETB_ETABLISSEMENT'+
        ' left join ANNUAIRE on ETB_CODEDDTEFPGU=ANN_GUIDPER'+
        ' WHERE ET_ETABLISSEMENT="'+E+'"' + Habilitation; //PT14
        orderbyEtab := ' ORDER BY ET_ETABLISSEMENT' ;

        If (ENS='') or (ENS='0') THEN          // PT- 5
        begin
                If (SOS='') or (SOS='0') Then
                begin
                        TFQRS1(Ecran).CodeEtat := 'MO2';
                        TFQRS1(Ecran).WhereSQL := SQLEtab+orderbyEtab;
                end
                Else
                begin
                        TFQRS1(Ecran).CodeEtat := 'MOT';
                        TFQRS1(Ecran).LaTob := TobEtat;
                end;
        end
        else
        Begin
                TFQRS1(Ecran).CodeEtat := 'MOT';
                TFQRS1(Ecran).LaTob := TobEtat;
        end;
end ;
// On calcule les effectifs au lancement de la fiche
procedure TOF_MOUVMAINOEUVRE.OnArgument (S : String ) ;
var ComboEtab,dateE,DateJour : THEdit;
    Check : TCheckBox ;
    QET : TQuery;
    Et,St : String ;
    Edit : THEdit;
    DateFin,DateDeb : TDateTime;
    etab,datetitreD,datetitreF,etablissement : String;
    BTob : TToolBarButton97;
    MultiCombo : THMultiValComboBox;
begin
Inherited ;

        if GetControlText('PRECEDENT')='X' then SetControlText('MOIS',DateToStr(PlusMois(Date,-1)));
        QET := OpenSQL('SELECT ET_ETABLISSEMENT,ET_LIBELLE FROM ETABLISS',True);
        Et := '';
        Etablissement := '';
        if not QET.eof then                                // PortageCWAS
        begin
                Et := QET.findfield('ET_ETABLISSEMENT').AsString;
                etablissement := QET.findfield('ET_LIBELLE').AsString;
        end;
        Ferme(QET);
        SetControlText('ETAB',Et);
        SetControlCaption('NOMETAB',etablissement);
        Check  :=  TCheckBox(GetControl('PRECEDENT'));
        DateDeb := DebutDeMois(StrToDate(GetControlText('MOIS')));
        DateFin := FinDeMois(StrToDate(GetControlText('MOIS')));
        etab := GetControlText('ETAB');
        datetitreD := DateToStr(DateDeb);
        datetitreF := DateToStr(DateFin);
        SetControlCaption('Titre','Mouvements de Main d''Oeuvre du ' + DatetitreD + ' au ' + DatetitreF) ;
        if check<>nil then Check.OnClick := precedent;
        ComboEtab := THEdit(GetControl('ETAB')) ;
        if ComboEtab<>nil then ComboEtab.OnExit := effectif; //On change remplacer
        MultiCombo:=THMultiValComboBox(getControl('ETABCOMPL'));
        If MultiCombo<>Nil then MultiCombo.OnChange:=Effectif;
        DateE := THEdit(GetControl('MOIS')) ;
        if DateE<>nil then
        begin
                DateE.Onchange := Affichelibelle; // PT- 4
                DateE.OnExit := effectif;            // PT- 4
                DateE.OnElipsisClick  :=  DateElipsisclick;
        end;
        DateJour := THEdit(GetControl('XX_VAREDITION')) ;
        If DateJour<>NIL Then DateJour.OnExit := VerifDate;
        If DateJour<>NIL Then DateJour.OnElipsisClick  :=  DateElipsisclick;
        Effectif(Nil); //PT6 Procedure effectif lancé pour effectuer les requêtes
        Edit := THEdit(GetControl('XX_EFFTOTAL'));         // PT- 7
        If Edit<>Nil Then Edit.OnExit := ExitEffectifDebut;
        Edit := THEdit(GetControl('XX_NBENTREE'));
        If Edit<>Nil Then Edit.OnExit := ExitEffectifDebut;
        Edit := THEdit(GetControl('XX_NBSORTIES'));
        If Edit<>Nil Then Edit.OnExit := ExitEffectifDebut;
        SetControlEnabled('XX_EFFFINAL',False);
        BTob:=TToolBarButton97(GetControl('BTOB'));
        If BTob<>Nil Then
        begin
                BTob.Visible := True;
                BTob.OnClick:=EtatAvecTob;
        end;

        { DEB PT11 Gestion du déclarant }
        Edit := THEdit(GetControl('DECLARANT'));
        If Edit<>Nil Then Edit.OnExit := AffichDeclarant;

        St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('ETAB')+'%") '+
        ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DMO%" )  ' ;
        SetControlProperty('DECLARANT', 'Plus' ,St );

        QET:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+
                   'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('ETAB')+'%") '+
                   'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DMO%" )  '+
                   'ORDER BY PDA_ETABLISSEMENT DESC',True);
        If Not QET.eof then
          Begin
          SetControlText('DECLARANT',QET.FindField('PDA_DECLARANTATTES').AsString);
          AffichDeclarant(nil);
          End;
        Ferme(QET);
        { FIN PT11 }
        SetControlVisible('BBlocNote',False);
        {$IFDEF EAGLCLIENT}
//        SetControlVisible('BTOB',False);
        {$ENDIF}
end ;

procedure TOF_MOUVMAINOEUVRE.ExitEffectifDebut(Sender : TObject); // PT- 7
var total : String;
begin
        Total := IntToStr(StrToInt(GetControlText('XX_EFFTOTAL'))+StrtoInt(GetControlText('XX_NBENTREE'))-StrToInt(GetControlText('XX_NBSORTIES')));
        SetControlText('XX_EFFFINAL',Total);
end;

procedure TOF_MOUVMAINOEUVRE.Affichelibelle(Sender : TObject);   // PT- 4
var DateFin,DateDeb : TDateTime;
begin
        If IsValidDate(GetControlText('MOIS')) Then
        begin
                DateDeb := DebutDeMois(StrToDate(GetControlText('MOIS')));
                DateFin := FinDeMois(StrToDate(GetControlText('MOIS')));
                SetControlCaption('Titre','Mouvements de Main d"Oeuvre du ' + DateToStr(DateDeb) + ' au ' + DateToStr(DateFin)) ;
        end;
end;

// Calcul des effectifs quand on change d'établissement
procedure TOF_MOUVMAINOEUVRE.Effectif(Sender : Tobject);
var DateEff,MDateEff,DateFin,DateDeb,DateDeb2 : TDateTime;
    etab,datelib1,datelib2,etablissement,St : String;
    QE,QSO,QEN,QET,QS,QInterim,Q : TQuery;
    TobEff,T : Tob;
    TotalEN,TotalSO,totalM,TotalF,Total : Integer;
    TotalENM,TotalSOM,TotalENF,TotalSOF : Integer;
    WhereEtab,E,WhereEtabPaie,WhereEtabSal,EtabCompl,WhereEtabPaie1,WhereEtabPaie2,WhereEtabInterim : String;
begin
        St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('ETAB')+'%") '+
        ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DMO%" )  ' ;
        SetControlProperty('DECLARANT', 'Plus' ,St );

        QET:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+
                   'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('ETAB')+'%") '+
                   'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DMO%" )  '+
                   'ORDER BY PDA_ETABLISSEMENT DESC',True);
        If Not QET.eof then
          Begin
          SetControlText('DECLARANT',QET.FindField('PDA_DECLARANTATTES').AsString);
          AffichDeclarant(nil);
          End;
        Ferme(QET);
        If not IsValidDate(GetControlText('MOIS')) Then
        begin
                PGIBox(''''+GetControlText('MOIS')+''' n''est pas une date correcte','CEGID PAIE S5');
                SetFocusControl('MOIS') ;
                exit;
        end;
        If IsValidDate(GetControlText('MOIS')) Then
        begin
                DateDeb := DebutDeMois(StrToDate(GetControlText('MOIS')));
                DateFin := FinDeMois(StrToDate(GetControlText('MOIS')));
                MDateEff := PlusMois(StrToDATE(GetControlText('MOIS')),-1);
                DateEff := FinDeMois(MDateEff);
                DateDeb2 := DebutDeMois(DateEff);
                datelib1 := DateToStr(DateEff);
                SetControlCaption('LIBEFF','Effectif au ' + Datelib1) ;
                datelib2 := DateToStr(DateFin);
                SetControlCaption('LIBFIN','Effectif au ' + Datelib2) ;

                E := GetControlText('ETAB');
                Total := 0;
                TotalM := 0;
                TotalF := 0;
                TotalEN := 0;
                TotalENM := 0;
                TotalENF := 0;
                TotalSO := 0;
                TotalSOM := 0;
                TotalSOF := 0;
                //NEW
                E := GetControlText('ETAB');
                EtabCompl := GetControlText('ETABCOMPL');
                Etab := ReadTokenPipe(EtabCompl,';');
                While Etab <> '' do
                begin
                     WhereEtab := WhereEtab + ' OR PCI_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabinterim := WhereEtabinterim + ' OR PEI_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabPaie := WhereEtabPaie + ' OR PPU_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabSal := WhereEtabSal + ' OR PSA_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabPaie1 := WhereEtabPaie1 + ' AND P1.PPU_ETABLISSEMENT<>"'+Etab+'"';
                     WhereEtabPaie2 := WhereEtabPaie2 + ' OR P2.PPU_ETABLISSEMENT="'+Etab+'"';
                     Etab := ReadTokenPipe(EtabCompl,';');
                end;
                WhereEtab := '(PCI_ETABLISSEMENT="'+E+'"' + WhereEtab+')';
                WhereEtabinterim := '(PEI_ETABLISSEMENT="'+E+'"' + WhereEtabinterim+')';
                WhereEtabSal := '(PSA_ETABLISSEMENT="'+E+'"' + WhereEtabSal+')';
                WhereEtabPaie := '(PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie+')';
                WhereEtabPaie1 := '(P1.PPU_ETABLISSEMENT<>"'+E+'"' + WhereEtabPaie1+')';
                WhereEtabPaie2 := '(P2.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie2+')';
                //Effectif
                        //Les contrats Modif requete PT15
                Q := OpenSQL('SELECT COUNT (DISTINCT PCI_SALARIE) EFFECTIF,PSA_SEXE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
                WhereEtab+' AND (PCI_DEBUTCONTRAT<"'+UsDateTime(DateDeb)+'") '+
                'AND (PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'" OR PCI_FINCONTRAT<="'+UsDateTime(IDate1900)+'"  OR PCI_FINCONTRAT IS NULL)'+
                ' AND PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="")'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalM := TotalM + T.GetValue('EFFECTIF');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalF := TotalF + T.GetValue('EFFECTIF');
                TobEff.Free;
                         //Les sans contrats Modif requete PT15
                Q := OpenSQL('SELECT COUNT(DISTINCT PSA_SALARIE) EFFECTIF,PSA_SEXE FROM SALARIES LEFT JOIN PAIEENCOURS ON PSA_SALARIE=PPU_SALARIE'+
                ' AND PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'"'+
                ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="") AND'+
                ' PSA_DATEENTREE<"'+UsDateTime(DateDeb)+'"'+
                ' AND (PSA_DATESORTIE>"'+UsDateTime(DateDeb)+'" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+UsDateTime(idate1900)+'")'+
                ' AND ('+WhereEtabPaie+' OR ('+WhereEtabSal+' AND'+
                ' PPU_DATEDEBUT IS NULL))'+
                ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
                WhereEtab+' AND (PCI_DEBUTCONTRAT<"'+UsDateTime(DateDeb)+'") '+
                'AND (PCI_FINCONTRAT>"'+UsDateTime(DateDeb)+'" OR PCI_FINCONTRAT<="'+UsDateTime(IDate1900)+'"  OR PCI_FINCONTRAT IS NULL))'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalM := TotalM + T.GetValue('EFFECTIF');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalF := TotalF + T.GetValue('EFFECTIF');
                TobEff.Free;
                Total := TotalM + TotalF;
                SetControlText('XX_EFFTOTAL',IntToStr(Total));

                //Les entrées
                        //Les contrats
                Q := OpenSQL('SELECT COUNT (PCI_SALARIE) NBENTREECT,PSA_SEXE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
                WhereEtab+' AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'"'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalENM := TotalENM + T.GetValue('NBENTREECT');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalENF := TotalENF + T.GetValue('NBENTREECT');
                TobEff.Free;
                         //Les sans contrats
                Q := OpenSQL('SELECT COUNT (DISTINCT PSA_SALARIE) NBENTREE,PSA_SEXE FROM SALARIES LEFT JOIN PAIEENCOURS P1 ON PSA_SALARIE=P1.PPU_SALARIE'+//PT6
                ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="")'+ //PT12
                ' AND (('+WhereEtabSal+' AND PSA_DATEENTREE<="'+UsDateTime(DateFin)+'" AND PSA_DATEENTREE>="'+UsDateTime(DateDeb)+'")'+
                ' OR (Exists (SELECT PPU_SALARIE FROM PAIEENCOURS P2 WHERE P2.PPU_SALARIE=PSA_SALARIE AND'+
                ' P1.PPU_DATEDEBUT="'+UsDateTime(DateDeb2)+'" AND P2.PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'"'+
                ' AND '+WhereEtabPaie1+' AND P1.PPU_ETABLISSEMENT IS NOT NULL AND '+WhereEtabPaie2+'))) '+
                ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
                WhereEtab+' AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'")'+ //PT15
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalENM := TotalENM + T.GetValue('NBENTREE');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalENF := TotalENF + T.GetValue('NBENTREE');
                TobEff.Free;
                TotalEn := TotalEnM + TotalEnF;
                SetControlText('XX_NBENTREE',IntToStr(TotalEn));
                //Les sorties
                EtabCompl := GetControlText('ETABCOMPL');
                Etab := ReadTokenPipe(EtabCompl,';');
                WhereEtabPaie2 := '';
                WhereEtabPaie1 := '';
                While Etab <> '' do
                begin
                     WhereEtabPaie1 := WhereEtabPaie1 + ' OR P1.PPU_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabPaie2 := WhereEtabPaie2 + ' AND P2.PPU_ETABLISSEMENT<>"'+Etab+'"';
                     Etab := ReadTokenPipe(EtabCompl,';');
                end;
                WhereEtabPaie1 := '(P1.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie1+')';
                WhereEtabPaie2 := '(P2.PPU_ETABLISSEMENT<>"'+E+'"' + WhereEtabPaie2+')';
                Q := OpenSQL('SELECT COUNT (PCI_SALARIE) NBSORTIECT,PSA_SEXE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
                'PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="") AND '+
                WhereEtab+' AND PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'" '+
                'AND PCI_FINCONTRAT<"'+UsDateTime(DateFin)+'"'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalSOM := TotalSOM + T.GetValue('NBSORTIECT');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalSOF := TotalSOF + T.GetValue('NBSORTIECT');
                TobEff.Free;
                Q := OpenSQL('SELECT COUNT (DISTINCT PSA_SALARIE) NBSORTIE,PSA_SEXE FROM SALARIES LEFT JOIN PAIEENCOURS P1 ON PSA_SALARIE=P1.PPU_SALARIE'+  //PT6
                ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="")'+  //PT12
                ' AND (('+WhereEtabSal+' AND PSA_DATESORTIE<="'+UsDateTime(DateFin)+'" AND PSA_DATESORTIE>="'+UsDateTime(DateDeb)+'")'+
                ' OR (EXISTS (SELECT PPU_SALARIE FROM PAIEENCOURS P2 WHERE P2.PPU_SALARIE=PSA_SALARIE AND'+
                ' P1.PPU_DATEDEBUT="'+UsDateTime(DateDeb2)+'" AND P2.PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'"'+
                ' AND '+WhereEtabPaie1+' AND P2.PPU_ETABLISSEMENT IS NOT NULL AND '+WhereEtabPaie2+')))'+
                ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
                WhereEtab+' AND (PCI_FINCONTRAT>="'+UsDateTime(datedeb)+'" AND PCI_FINCONTRAT<"'+UsDateTime(datefin)+'"))'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalSOM := TotalSOM + T.GetValue('NBSORTIE');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalSOF := TotalSOF + T.GetValue('NBSORTIE');
                TobEff.Free;
                TotalSo := TotalSoF + TotalSoM;
                SetControlText('XX_NBSORTIES',IntToStr(TotalSo));
                SetControlText('XX_EFFFINAL',IntToStr(Total + TotalEn - TotalSO));
                SetControlText('XX_H',IntToStr(TotalM + TotalEnM - TotalSOM));
                SetControlText('XX_F',IntToStr(TotalF + TotalEnF - TotalSOF));
               //FIN NEW
                //Debut PT- 8
                QInterim := OpenSQL('SELECT COUNT (DISTINCT PEI_INTERIMAIRE) AS INTERIMAIRE FROM EMPLOIINTERIM LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PEI_INTERIMAIRE'+
                ' WHERE '+WhereEtabinterim+' AND PEI_DEBUTEMPLOI<="'+UsDateTime(DateFin)+'" AND (PEI_FINEMPLOI="'+UsdateTime(IDate1900)+'" OR PEI_FINEMPLOI>="'+UsDateTime(DateFin)+'") AND PSI_TYPEINTERIM="INT"',True);
                if not QInterim.eof then SetControlText('XX_NBINTERIM',QInterim.FindField('INTERIMAIRE').AsString); // PortageCWAS
                Ferme(QInterim);
                // FIN PT- 8
        end;
end;

procedure TOF_MOUVMAINOEUVRE.Precedent(Sender : Tobject);
begin
        if sender=nil then exit;
        if GetControlText('PRECEDENT')='X' then
        begin
                SetControlEnabled('MOIS',False);
                SetControlText('MOIS',DateToStr(PlusMois(Date,-1)));
        end
        Else SetControlEnabled('MOIS',True);
        Effectif(Sender);
end;

// PT- 1
procedure TOF_MOUVMAINOEUVRE.VerifDate(Sender : TOBject);
begin
        If Not IsValidDate(THEdit(Sender).text) Then
        begin
                PGIBox(''''+THEdit(Sender).text+''' n''est pas une date correcte','CEGID PAIE S5');
                THEdit(Sender).SetFocus ;
                exit;
        end;
end;

procedure TOF_MOUVMAINOEUVRE.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_MOUVMAINOEUVRE.EtatAvecTob(Sender:TObject);
var
   M: THEdit;
   Datedeb,datefin,DateEff,MDateEff,DateDeb2:TDateTime;
   St,StEffectif:String;
begin
Datedeb := IDate1900;
datefin := IDate1900;
DateDeb2 := IDate1900;
Effectif(Nil);
M:=THEdit (GetControl ('MOIS'));
if M=nil then exit;
if IsValidDate(M.text) then
   begin
   DateDeb:=DebutDeMois(StrToDate(M.text));
   SetControlText('XX_VARIABLEDEB',DateToStr(DateDeb));
   DateFin:=FinDeMois(StrToDate(M.text));
   SetControlText('XX_VARIABLEFIN',DateToStr(DateFin));
   MDateEff:=PlusMois(StrToDATE(M.text),-1);
   DateEff:=FinDeMois(MDateEff);
   DateDeb2:=DebutDeMois(DateEff);
   SetControlText('XX_VARIABLEEFF',DateToStr(DateEff));
   End;
StEffectif:=GetControlText('XX_EFFTOTAL')+';'+
            GetControlText('XX_NBENTREE')+';'+
            GetControlText('XX_NBSORTIES')+';'+
            GetControlText('XX_EFFFINAL')+';'+
            GetControlText('XX_H')+';'+
            GetControlText('XX_F')+';'+
            GetControlText('XX_NBINTERIM')+';'+
            GetControlText('XX_VARIABLEEFF')+'';
St:=GetControlText('ETAB')+';'+DateToStr(DateDeb)+';'+DateToStr(DateDeb2)+';'+DateToStr(dateFin);
AGLLanceFiche('PAY','EDITMMOTOB','','',St+';'+StEffectif);
end;

{ DEB PT11 }
procedure TOF_MOUVMAINOEUVRE.AffichDeclarant(Sender: TObject);
begin
if GetControlText('DECLARANT')='' then exit;
SetControlText('XX_VARNOM'       ,RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
SetControlText('XX_VARTEL'       ,RechDom('PGDECLARANTTEL'   ,GetControlText('DECLARANT'),False));
SetControlText('XX_VARPOSTE'     ,RechDom('PGDECLARANTPOSTE' ,GetControlText('DECLARANT'),False));
end;
{ FIN PT11 }
Initialization
  registerclasses ( [ TOF_MOUVMAINOEUVRE ] ) ;
end.
