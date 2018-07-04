{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 28/11/2001
Modifié le ... :
Description .. :
Mots clefs ... : PAIE;MSA
*****************************************************************
PT1 22/05/2002 JL V582 : Calcul de la base CSG/RDS lors d'un changement de la valeur de l'activité
PT2 12/06/2002 JL V582 : propriété ChoixEtat = true.
PT3 25/06/2003 JL V_42 : ajout psa_salarie dans order by requete pour rupture par code au lieu libellé
PT4 05/10/2004 JL V_50 : ajout declarant
PT5 29/10/2004 JL V_60 : FQ 11251 Liste export grisé
PT6 06/04/2006 JL V_65 : FQ 13052 fonction convertprefixe dans cas edition période.
PT7 14/04/2006 JL V_65 : FQ 12147 Ajout nom du dossier
PT8 16/05/2006 JL V_65 : Calcul TCP
PT9 05/06/2006 JL V_65 : FQ 13228 Tri édition période sur élément de rémunération
PT10 04/01/2007 PH V_70 : Rajout PPU_CBRUTFISCAL dans requête SQL
PT11 25/07/2007 JL V_80 : FQ 14570 Ajout exclusion salariés sortis
                          + Remplacement LanceEtatTob par TFQRS1(Ecran).LaTob
PT12 08/08/2007 JL V_80 : FQ 13731 Ajout Nb jours dans éditions des périodes
PT13 14/11/2007 FC V_80 : FQ 14892 Evolutions cahier des charges Octobre 2007
PT14 21/01/2008 FC V_81 : FQ 15124 Dernière ligne du fichier non traitée
PT15 16/06/2008 FC V_81 : FQ 15549 Le pavé entreprise avec la CSG n'est pas édité
PT16 26/06/2008 FC V_81 : FQ 15589 Seules les dernières heures TEPA sont éditées
PT18 27/08/2008 FC V_810 : FQ 15682 Modification cahier des charges
PT19 10/09/2008 FC Ne pas faire le controle de cohérence de dates
}
Unit UTOFPGEditMSA_DeclSalaire ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBCtrls,HDB,QRS1,EdtREtat,
{$ELSE}
     eQRS1,UtileAGL,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,UTOB,ParamDat,HQry,PGEdtEtat,ParamSoc,Entpaie,HTB97,Vierge,HSysMenu ;

Type
    TOF_PGMSA_DeclSalaire = Class (TOF)
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    TobEtat : Tob;
    Etablissement:THValCombobox;
    TypeEdition : String;
    procedure CalcDateRetour;
    procedure CalcBaseCsg;
    procedure MetABlanc(Sender:TObject);
    procedure DateElipsisClick(Sender: TObject);
    procedure VerifDate(Sender:TObject);
    procedure EditionPeriodes;
    procedure EditionEvolutions;
    Procedure AjoutLigneEvolution (TobLigneEvolution : Tob;Libelle,Valeur : String);
    procedure AffichDeclarant(Sender:TObject); //PT4
    procedure ClickSortie(Sender : TObject);
//PT19    procedure VerifDateFin(Sender: TObject);
  end ;

//DEB PT13 Toute cette partie avait disparue. Pourquoi ???
Type
    TOF_PGMSARECAPENVOI = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose ; override ;
    Private
    TobEtat : Tob;
    TobAssiettes : Tob;
    GRem: THGrid;
    procedure edition(Sender : TObject);
    Procedure AjoutLigneEvolutionRecap (TobLigneEvolution : Tob;Libelle,Valeur : String);
    procedure ChargeZones();
    Procedure EditionFichierMsa;
    Function TraiteValeurNum (Valeur : String): String;
    Function TraiteValeurHeures (Valeur : String): String;
    Procedure AjoutLigneEvolutionFichier (ValeurSeg : String;Libelle,Valeur : String);
    procedure DateElipsisClick(Sender: TObject);
    procedure VerifDate(Sender:TObject);
    procedure RemplirGrilleRem;
    procedure AffichDeclarant(Sender:TObject);
    procedure MajPeriodes(Sender : TObject);
    Function RendLibAssiette(Code : String) : String;
    procedure ForceNumerique(Buf: string; var Resultat: string);
    procedure AjoutLigneHSupFichier (S : String;Ordre : integer);
    Function TraiteValeurNumDec (Valeur : String): String;
  end ;
//FIN PT13

Implementation
{
Sur le bouton Valider, on confectionne la requete
}

procedure TOF_PGMSA_DeclSalaire.OnClose ;
begin
  Inherited;
  If TobEtat <> Nil then FreeAndNil(TobEtat);
end;

procedure TOF_PGMSA_DeclSalaire.OnUpdate ;
Var Titre:THEdit;
    Defaut:THValComboBox;
    DateDebut,DateFin:TDateTime;
    SQL,Activite:String;
    QNB,Q:TQuery;
    aa,mm,jj:Word;
    THVActivite:THValComboBox;
begin
     Inherited ;
If TypeEdition = 'EVOLUTIONS' then
begin
  EditionEvolutions;
  Exit;
end;
If TypeEdition = 'PERIODES' then
begin
  EditionPeriodes;
  Exit;
end;

// Calcul des dates selon le trimestre
Q:=OpenSQL('SELECT ETB_MSASECTEUR FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+Etablissement.Value+'"',True);
If Not Q.eof then SetControlText('SECTEUR',Q.FindField('ETB_MSASECTEUR').AsString); // PortageCWAS
Ferme(Q);
THVActivite:=THValComboBox(GetControl('CACTIVITE'));
Q:=OpenSQL('SELECT CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="PMS" AND CC_CODE="'+THVACtivite.Value+'"',True);
if not Q.eof then SetControlText('CODEACT',Q.FindField('CC_ABREGE').AsString); // PortageCWAS
Ferme(Q);
Titre:=THEdit(GetControl('ETITRE'));
DateDebut:=StrToDate(GetControlText('DATEDEBUT'));
DateFin:=StrToDate(GetControlText('DATEFIN'));
DecodeDate(DateDebut,aa,mm,jj);
If mm=1 Then Titre.Text:='1ER TRIMESTRE '+IntToStr(aa);
If mm=4 Then Titre.Text:='2EME TRIMESTRE '+IntToStr(aa);
If mm=7 Then Titre.Text:='3EME TRIMESTRE '+IntToStr(aa);
If mm=10 Then Titre.Text:='4EME TRIMESTRE '+IntToStr(aa);
Defaut:=THValComboBox(GetControl('CACTIVITE'));
Activite:=defaut.Value;
QNB:=OpenSQL('SELECT COUNT (PPU_SALARIE) AS NB FROM ETABLISS'+
     ' LEFT JOIN PAIEENCOURS ON ET_ETABLISSEMENT=PPU_ETABLISSEMENT'+
     ' WHERE PPU_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PPU_DATEFIN<="'+UsDateTime(DateFin)+'"'+
     ' AND ET_ETABLISSEMENT="'+Etablissement.Value+'"',true);
If not QNB.eof then SetControlText('ENBENREG',QNB.FindField('NB').AsString); // PortageCWAS
Ferme(QNB);
SetControlText('DATEDEBUT',DateToStr(DateDebut));
SetControlText('DATEFIN',DateToStr(DateFin));
// PT10
SQL:='SELECT DISTINCT PSE_MSAINFOSCOMPL,PSE_MSALIEUTRAV,PSA_NUMEROSS,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_COEFFICIENT,ET_ETABLISSEMENT,ET_LIBELLE,ET_ACTIVITE,'+
     ' ET_SIRET,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_CBRUT,PPU_CBRUTFISCAL,(PHC_MONTANT) AS HEURESTRAV'+
     ' ,PSA_DATEENTREE,PSA_DATESORTIE FROM ETABLISS'+
     ' LEFT JOIN PAIEENCOURS ON ET_ETABLISSEMENT=PPU_ETABLISSEMENT'+
     ' LEFT JOIN SALARIES ON PSA_SALARIE=PPU_SALARIE'+
     ' LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE'+
     ' LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="21"'+
     ' WHERE PSE_MSAACTIVITE="'+THVActivite.Value+'" AND PSE_MSA="X" AND PPU_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PPU_DATEFIN<="'+UsDateTime(DateFin)+'"'+
     ' AND ET_ETABLISSEMENT="'+Etablissement.Value+'"'+
     ' GROUP BY PSE_MSAINFOSCOMPL,PSE_MSALIEUTRAV,PSA_NUMEROSS,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_COEFFICIENT,ET_ETABLISSEMENT,ET_LIBELLE,ET_ACTIVITE,PHC_MONTANT,'+
     ' ET_SIRET,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_CBRUT,PPU_CBRUTFISCAL,PSA_DATEENTREE,PSA_DATESORTIE '+
     ' ORDER BY ET_ETABLISSEMENT,PSA_LIBELLE,PSA_SALARIE,PPU_DATEDEBUT';
TFQRS1(Ecran).WhereSQL:=SQL;
end ;

procedure TOF_PGMSA_DeclSalaire.OnArgument (S : String ) ;
var QET,QAttes : TQuery;
    Et,Act,St : String;
    aa,mm,jj,Mois : word;
    Defaut,Edit : THEdit;
    DateDebut,DateFin : TDateTime;
    Activite : THValComboBox;
    Arg : String;
    Check : TCheckBox;
begin
Inherited ;
SetControlEnabled('FLISTE',False); //PT5
Arg := ReadTokenPipe(S,';');
TypeEdition := Arg;
If (Arg <> 'EVOLUTIONS') and (Arg <> 'PERIODES') then
begin
        TFQRS1(Ecran).ChoixEtat:=True;  // DEBUT PT2
        SetControlEnabled('FETAT',True);
        {$IFNDEF EAGLCLIENT}
        SetControlVisible('BPARAMETAT',True);       // FIN PT2
        {$ELSE}
        SetControlVisible('BPARAMETAT',False);
        {$ENDIF}
        QET:=OpenSQL('SELECT ET_ETABLISSEMENT,ET_LIBELLE,ETB_MSAACTIVITE,ETB_MSASECTEUR FROM ETABLISS LEFT JOIN ETABCOMPL ON ET_ETABLISSEMENT=ETB_ETABLISSEMENT',True);
        Et:='';
        Act:='';
        if not QET.eof then
        begin
                Et:=QET.findfield('ET_ETABLISSEMENT').AsString;   // PortageCWAS
                Act:=QET.FindField('ETB_MSAACTIVITE').AsString;
                SetControlText('SECTEUR',QET.FindField('ETB_MSASECTEUR').AsString);
        end;
        Ferme(QET);
         //DEBUT PT4
           St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + GetControltext('CETAB') + '%") ' +
    ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MSA%" )  ';
          SetControlProperty('DECLARANT', 'Plus', St);
        QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
        'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + GetControlText('CETAB') + '%") ' +
      'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MSA%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
        if not QAttes.eof then
        begin
                SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
                AffichDeclarant(nil);
        end;
        Ferme(QAttes);
        Edit := THEdit(GetControl('DECLARANT'));
        If Edit<>Nil Then Edit.OnExit := AffichDeclarant;
        //FIN PT4
        Etablissement:=THValCombobox(GetControl('CETAB'));
        If Etablissement<>Nil Then
        begin
                Etablissement.Value:=Et;
                Etablissement.OnChange:=MetABlanc;
        end;
        Activite:=THValComboBox(GetControl('CACTIVITE'));
        If Activite<>Nil Then
        begin
                Activite.Value:=Act;
                Activite.OnChange:=MetABlanc;   //PT1
        end;
        DecodeDate(Date,aa,mm,jj);
        Mois:=1;
        If mm<4 Then
        begin
                aa:=aa-1;
                Mois:=10;
        end;
        If (mm<7) and (mm>3) Then Mois:=1;
        If (mm<10) and (mm>6) Then Mois:=4;
        If mm>9 Then Mois:=7;
        DateDebut:=EncodeDate(aa,Mois,1);
        DateFin:=PlusMois(DateDebut,2);
        DateFin:=FinDeMois(DateFin);
        SetControlText('DATEDEBUT',DateToStr(DateDebut));
        SetControlText('DATEFIN',DateToStr(DateFin));
        Defaut:=THEdit(GetControl('DATEDEBUT'));
        If Defaut<>Nil Then
        begin
                Defaut.OnElipsisClick:=DateElipsisClick;
                Defaut.OnExit:=Verifdate;
        end;
        Defaut:=THEdit(GetControl('DATEFIN'));
        If Defaut<>Nil Then
        begin
                Defaut.OnExit:=Verifdate;
                Defaut.OnElipsisClick:=DateElipsisClick;
        end;
        CalcBaseCsg;
        CalcDateRetour;
end
else
begin
        DecodeDate(Date,aa,mm,jj);
        Mois:=1;
        If mm<4 Then
        begin
                aa:=aa-1;
                Mois:=10;
        end;
        If (mm<7) and (mm>3) Then Mois:=1;
        If (mm<10) and (mm>6) Then Mois:=4;
        If mm>9 Then Mois:=7;
        DateDebut:=EncodeDate(aa,Mois,1);
        DateFin:=PlusMois(DateDebut,2);
        DateFin:=FinDeMois(DateFin);
        SetControlText('DATEDEBUT',DateToStr(DateDebut));  //PT13
        SetControlText('DATEFIN',DateToStr(DateFin)); //PT13
        Defaut:=THEdit(GetControl('DATEDEBUT'));//PT13
        If Defaut<>Nil Then
        begin
           Defaut.OnElipsisClick:=DateElipsisClick;
//PT19           Defaut.OnExit:=Verifdatefin;
        end;
        Defaut:=THEdit(GetControl('DATEFIN')); //PT13
        If Defaut<>Nil Then
        begin
//PT19                Defaut.OnExit:=VerifdateFin;
                Defaut.OnElipsisClick:=DateElipsisClick;
        end;
        If Arg = 'EVOLUTIONS' then
        begin
//                TFQRS1(ECran).BValider.OnClick := editionEvolutions;
                Ecran.Caption := 'Editions des évolutions pour MSA';
        end;
        If Arg = 'PERIODES' then
        begin
                //TFQRS1(ECran).BValider.OnClick := EditionPeriodes;
                Ecran.Caption := 'Editions des périodes pour MSA';
        end;
        UpdateCaption(TFQRS1(Ecran)) ;
end;
If THEdit(GetControl('DOSSIER')) <> Nil then SetControlText('DOSSIER',GetParamSocSecur ('SO_LIBELLE',''));//PT7
//DEBUT PT11
  Check := TCheckBox(GetControl('CKSORTIE'));
  If Check <> Nil then Check.OnClick := ClickSortie;
//FIN PT11
end ;

Procedure TOF_PGMSA_DeclSalaire.CalcDateRetour;
Var DateRetour:TDateTime;
    aa,mm,jj:Word;
begin
If Not IsValidDate(GetControlText('DATEFIN')) Then Exit;
DateRetour:=PlusMois(StrToDate(GetControlText('DATEFIN')),1);
DecodeDate(DateRetour,aa,mm,jj);
DateRetour:=EncodeDate(aa,mm,10);
SetControlText('EDATERETOUR',DateToStr(DateRetour));
CalcBaseCsg;
end;

procedure TOF_PGMSA_DeclSalaire.CalcBaseCsg;
var dateDebut,DateFin:TDateTime;
    Q:TQuery;
    Montant:Double;
    ComboAct:THValComboBox;
begin
If Etablissement.Value='' Then exit;
ComboAct:=THValComboBox(GetControl('CACTIVITE'));
DateDebut:=StrToDate(GetControlText('DATEDEBUT'));
DateFin:=StrToDate(GetControlText('DATEFIN'));
Q:=OpenSQL(' SELECT SUM (PHB_BASECOT) AS TOTAL FROM'+
     ' HISTOBULLETIN'+
     ' LEFT JOIN COTISATION ON ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE AND PHB_NATURERUB=PCT_NATURERUB'+
     ' LEFT JOIN DEPORTSAL ON PHB_SALARIE=PSE_SALARIE'+
     ' WHERE PHB_DATEFIN>="'+UsDateTime(DateDebut)+'" AND PHB_DATEFIN<="'+UsDateTime(DateFin)+'"'+
     ' AND PCT_BASECSGCRDS="X" AND PHB_ETABLISSEMENT="'+Etablissement.Value+'" AND PSE_MSA="X" AND PSE_MSAACTIVITE="'+ComboAct.Value+'"',True);
Montant:=0;
if not Q.eof then Montant:=Q.Findfield('TOTAL').AsFloat;   // PortageCWAS
Ferme(Q);
SetControlText('CSGRDS',FloatToStr(Arrondi(Montant,0)));
//DEBUT PT8
If GetParamSocSecur ('SO_PGMSACALCTCP',False) then
begin
        if (VH_PAIE.PGProfilFnal <> '') and (VH_PAIE.PGProfilFnal <> NULL) then
        begin
                Q := OpenSQL('SELECT SUM(PHC_MONTANT) AS MONTANT ' +
                'FROM HISTOCUMSAL '+
                ' LEFT JOIN DEPORTSAL ON PHC_SALARIE=PSE_SALARIE '+
                'WHERE PHC_CUMULPAIE="35" AND PHC_DATEDEBUT>="' + UsDateTime(DateDebut) + '"  ' +
                'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" AND PHC_ETABLISSEMENT="' + Etablissement.Value + '"'+
                ' AND PSE_MSAACTIVITE="'+ComboAct.Value+'"',true);
                If Not Q.Eof then SetControlText('TCP', FloatToStr(Q.FindField('MONTANT').AsFloat))
                else SetControlText('TCP', '0');
                Ferme(Q);
        end
        else SetControlText('TCP', '0');
end
else SetControlText('TCP', '0');
//FIN PT8
end;

procedure TOF_PGMSA_DeclSalaire.MetABlanc(Sender:TObject);
var Q,QAttes : TQuery;
    Act : THValComboBox;
    St : String;
begin
If Sender=Nil Then Exit;
SetControlText('CSGRDS','0');
SetControlText('CSGRDSIMP','0');
SetControlText('CSGRDSNONIMP','0');
SetControlText('TCP','0');
SetControlText('RDSRR','0');
If THValComboBox(Sender).name='CETAB' Then    //PT1
   begin
   Act:=THValComboBox(GetControl('CACTIVITE'));
   Q:=OpenSQL('SELECT ETB_MSAACTIVITE FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+THValComboBox(Sender).Value+'"',True);
   if not Q.eof then Act.Value:=Q.FindField('ETB_MSAACTIVITE').AsString;  // PortageCWAS
   Ferme(Q);

   //DEBUT PT4
   St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + GetControltext('CETAB') + '%") ' +
    ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MSA%" )  ';
   SetControlProperty('DECLARANT', 'Plus', St);
        QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
        'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + GetControlText('CETAB') + '%") ' +
      'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MSA%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
        if not QAttes.eof then
        begin
                SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
                AffichDeclarant(nil);
        end;
        Ferme(QAttes);
        //FIN PT4
   end;
CalcBaseCsg;
end;

procedure TOF_PGMSA_DeclSalaire.DateElipsisClick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMSA_DeclSalaire.VerifDate(Sender:TObject);
var aa,jj,mm:Word;
    Jour,Mois :String;
    Datedebut,DateFin:TDateTime;
begin
IF Sender=Nil Then Exit;
If THEdit(Sender).Name='DATEDEBUT' Then
   begin
   If Not IsValidDate(THEdit(Sender).Text) Then
      begin
      PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas valide','Edition déclaration des salaires');
      SetFocusControl('DATEDEBUT');
      Exit;
      end;
   DateDebut:=StrToDate(GetControlText('DATEDEBUT'));
   DecodeDate(DateDebut,aa,mm,jj);
   Mois:=IntToStr(mm);
   Jour:=IntToStr(jj);
   If (Jour<>'1') or (Mois<>'1') Then
      If (Jour<>'1') or (Mois<>'4') Then
         If (Jour<>'1') or (Mois<>'7') Then
            If (Jour<>'1') or (Mois<>'10') Then
            begin
            PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas une date de début de trimestre','Edition déclaration des salaires');
            SetFocusControl('DATEDEBUT');
            Exit;
            end;
   DateFin:=PlusMois(DateDebut,2);
   DateFin:=FinDeMois(DateFin);
   SetControlText('DATEFIN',DateToStr(DateFin));
   CalcDateRetour;
   Exit;
   end;
If THEdit(Sender).Name='DATEFIN' Then
   begin
   If Not IsValidDate(THEdit(Sender).Text) Then
      begin
      PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas valide','Edition déclaration des salaires');
      SetFocusControl('DATEFIN');
      Exit;
      end;
   DateFin:=StrToDate(GetControlText('DATEFIN'));
   DecodeDate(DateFin,aa,mm,jj);
   Mois:=IntToStr(mm);
   Jour:=IntToStr(jj);
   If (Jour<>'31') or (Mois<>'3') Then
      If (Jour<>'30') or (Mois<>'6') Then
         If (Jour<>'30') or (Mois<>'9') Then
            If (Jour<>'31') or (Mois<>'12') Then
            begin
            PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas une date de fin de trimestre','Edition déclaration des salaires');
            SetFocusControl('DATEFIN');
            Exit;
            end;
   DateDebut:=PlusMois(DateFin,-2);
   DateDebut:=DebutDeMois(DateDebut);
   SetControlText('DATEDEBUT',DateToStr(DateDebut));
   CalcDateRetour;
   end;
end;

procedure TOF_PGMSA_DeclSalaire.EditionPeriodes;
var TobPeriodes,TE : Tob;
    Q : TQuery;
    i,x,r  : Integer;
    StWhere : String;
    Pages : TPageControl;
    St : String;
    MemSal:String;
    MemOrdre:integer;
    MemDateDeb,MemDateFin:TDateTime;
begin
        If TobEtat <> Nil then FreeAndNil(TobEtat);
        Pages := TPageControl(GetControl('PAGES'));
        StWhere := RecupWhereCritere(Pages);
        StWhere := ConvertPrefixe(StWhere,'PE2','PE3'); //PT6
        //DEBUT PT11
        If TCheckBox(GetControl('CKSORTIE')) <> Nil then
        begin
          If GetCheckBoxState('CKSORTIE') = CbChecked then
          begin
            If StWhere <> '' then StWhere := StWhere + 'AND PE3_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
            '(PSA_DATESORTIE>="'+UsDateTime(StrtoDate(GetControlText('DATEARRET')))+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL))'
            else StWhere := StWhere + 'WHERE PE3_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
            '(PSA_DATESORTIE>="'+UsDateTime(StrtoDate(GetControlText('DATEARRET')))+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL))';
          end;
        end;
        //FIN PT11
        If StWhere <> '' then
        begin
                StWhere := ' ' + StWhere + ' AND PE3_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PE3_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"'; //PT13
        end
        else StWhere := ' WHERE PE3_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PE3_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"'; //PT13
//Optimisation        Q := OpenSQL('SELECT * FROM MSAPERIODESPE31'+StWHere,True);
//        TobPeriodes.LoadDetailDB('LesPeriodes','','',Q,False);
        St := 'SELECT * FROM MSAPERIODESPE31'+StWHere + ' AND PE3_ORDRE < 100'; //PT13
        TobPeriodes := Tob.Create('LesPeriodes',Nil,-1);
        TobPeriodes.LoadDetailDBFromSQL ('LesPeriodes', St);
//        Ferme(Q);
        TobEtat := Tob.Create('Edition',Nil,-1);
        For i := 0 to TobPeriodes.Detail.Count - 1 do
        begin
                For x := 1 to 5 do
                begin
                        If TobPeriodes.Detail[i].GetValue('PE3_ELTCALCMSA'+IntToStr(x)) <> ''then
                        begin
                                TE := Tob.Create('FilleEtat',TobEtat,-1);
                                TE.AddchampSupValeur('PE3_SALARIE',TobPeriodes.Detail[i].GetValue('PE3_SALARIE'),False);
                                TE.AddchampSupValeur('PE3_ETABLISSEMENT',TobPeriodes.Detail[i].GetValue('PE3_ETABLISSEMENT'),False);
                                TE.AddchampSupValeur('PE3_NOM',TobPeriodes.Detail[i].GetValue('PE3_NOM'),False);
                                TE.AddchampSupValeur('PE3_PRENOM',TobPeriodes.Detail[i].GetValue('PE3_PRENOM'),False);
                                TE.AddchampSupValeur('PE3_NUMEROSS',TobPeriodes.Detail[i].GetValue('PE3_NUMEROSS'),False);
                                TE.AddchampSupValeur('PE3_MSAACTIVITE',TobPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE'),False);
                                TE.AddchampSupValeur('PE3_DATEDEBUT',TobPeriodes.Detail[i].GetValue('PE3_DATEDEBUT'),False);
                                TE.AddchampSupValeur('PE3_DATEFIN',TobPeriodes.Detail[i].GetValue('PE3_DATEFIN'),False);
                                TE.AddchampSupValeur('PE3_NBHEURES',TobPeriodes.Detail[i].GetValue('PE3_NBHEURES'),False);
                                TE.AddchampSupValeur('PE3_NBJOURS',TobPeriodes.Detail[i].GetValue('PE3_NBJOURS'),False);//PT12
                                TE.AddchampSupValeur('PE3_ELTCALCMSA1',TobPeriodes.Detail[i].GetValue('PE3_ELTCALCMSA'+IntToStr(x)),False);
                                TE.AddchampSupValeur('PE3_MONTANTELT1',TobPeriodes.Detail[i].GetValue('PE3_MONTANTELT'+IntToStr(x)),False);
                                TE.AddchampSupValeur('CHAMPTEXTE','',False);
                        end;
                end;
        end;
        FreeAndNil(TobPeriodes);

        //DEB PT13
        St := 'SELECT * FROM MSAPERIODESPE31'+StWHere + ' AND PE3_ORDRE >= 100 ORDER BY PE3_SALARIE,PE3_DATEDEBUT,PE3_DATEFIN,PE3_ORDRE';
        TobPeriodes := Tob.Create('LesPeriodes',Nil,-1);
        TobPeriodes.LoadDetailDBFromSQL ('LesPeriodes', St);
        r := 1;
        MemSal:='';
        MemOrdre:=0;
        MemDateDeb:=idate1900;
        MemDateFin:=idate1900;
        For i := 0 to TobPeriodes.Detail.Count - 1 do
        begin
          if {(TobPeriodes.Detail[i].GetValue('PE3_ORDRE')<>MemOrdre) //PT16
            or} (TobPeriodes.Detail[i].GetValue('PE3_SALARIE')<>MemSal)
            or (TobPeriodes.Detail[i].GetValue('PE3_DATEDEBUT')<>MemDateDeb)
            or (TobPeriodes.Detail[i].GetValue('PE3_DATEFIN')<>MemDateFin) then
//          if TobPeriodes.Detail[i].GetValue('PE3_ORDRE') = 100 then
          begin
            r := 1;
            MemOrdre:=TobPeriodes.Detail[i].GetValue('PE3_ORDRE');
            MemSal:=TobPeriodes.Detail[i].GetValue('PE3_SALARIE');
            MemDateDeb:=TobPeriodes.Detail[i].GetValue('PE3_DATEDEBUT');
            MemDateFin:=TobPeriodes.Detail[i].GetValue('PE3_DATEFIN');
          end;
          TE := Tob.Create('FilleEtat',TobEtat,-1);
          TE.AddchampSupValeur('PE3_SALARIE',TobPeriodes.Detail[i].GetValue('PE3_SALARIE'),False);
          TE.AddchampSupValeur('PE3_ETABLISSEMENT',TobPeriodes.Detail[i].GetValue('PE3_ETABLISSEMENT'),False);
          TE.AddchampSupValeur('PE3_NOM',TobPeriodes.Detail[i].GetValue('PE3_NOM'),False);
          TE.AddchampSupValeur('PE3_PRENOM',TobPeriodes.Detail[i].GetValue('PE3_PRENOM'),False);
          TE.AddchampSupValeur('PE3_NUMEROSS',TobPeriodes.Detail[i].GetValue('PE3_NUMEROSS'),False);
          TE.AddchampSupValeur('PE3_MSAACTIVITE',TobPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE'),False);
          TE.AddchampSupValeur('PE3_DATEDEBUT',TobPeriodes.Detail[i].GetValue('PE3_DATEDEBUT'),False);
          TE.AddchampSupValeur('PE3_DATEFIN',TobPeriodes.Detail[i].GetValue('PE3_DATEFIN'),False);
          TE.AddchampSupValeur('PE3_NBHEURES',StrToInt(TobPeriodes.Detail[i].GetValue('PE3_MONTANTELT3')) / 100,false);
          TE.AddchampSupValeur('PE3_NBJOURS','',False);
          TE.AddchampSupValeur('PE3_ELTCALCMSA1','H' + IntToStr(r),False);
          TE.AddchampSupValeur('PE3_MONTANTELT1',TobPeriodes.Detail[i].GetValue('PE3_MONTANTELT4'),false);
          TE.AddchampSupValeur('CHAMPTEXTE','Type ' + IntToStr(TobPeriodes.Detail[i].GetValue('PE3_MONTANTELT1')) + ' - Taux de majoration ' + IntToStr(r) + ' à ' + IntToStr(TobPeriodes.Detail[i].GetValue('PE3_MONTANTELT2')) + ' %',false);
          r := r + 1;
        end;
        FreeAndNil(TobPeriodes);
        //FIN PT13

        If GetCheckBoxState('CRUPTETAB') = CbChecked then
        begin
                If GetCheckBoxState('CALPHA') = CbChecked then TobEtat.Detail.Sort('PE3_ETABLISSEMENT;PE3_NOM;PE3_SALARIE;PE3_DATEDEBUT;PE3_ELTCALCMSA1')
                else TobEtat.Detail.Sort('PE3_ETABLISSEMENT;PE3_SALARIE;PE3_DATEDEBUT;PE3_ELTCALCMSA1');
        end
        else
        begin
                If GetCheckBoxState('CALPHA') = CbChecked then TobEtat.Detail.Sort('PE3_NOM;PE3_SALARIE;PE3_DATEDEBUT;PE3_ELTCALCMSA1')
                else TobEtat.Detail.Sort('PE3_SALARIE;PE3_DATEDEBUT;PE3_ELTCALCMSA1');
        end;
//        LanceEtatTOB('E','PMS','PPE',TobEtat,True,False,False,Pages,'','',False);
        TFQRS1(Ecran).NatureEtat := 'PMS'; //PT13
        TFQRS1(Ecran).CodeEtat := 'PPE';   //PT13
        TFQRS1(Ecran).LaTob:= TobEtat;
end;

procedure TOF_PGMSA_DeclSalaire.EditionEvolutions ;
var TobEvolutions : Tob;
    Q : TQuery;
    i : Integer;
    TypeEvolution,Donnee,StWhere : String;
    Pages : TPageControl;
    St : String;
begin
        If TobEtat <> Nil then FreeAndNil(TobEtat);
        Pages := TPageControl(GetControl('PAGES'));
        StWhere := RecupWhereCritere(Pages);
        //DEBUT PT11
        If TCheckBox(GetControl('CKSORTIE')) <> Nil then
        begin
          If GetCheckBoxState('CKSORTIE') = CbChecked then
          begin
            If StWhere <> '' then StWhere := StWhere + 'AND PE2_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
            '(PSA_DATESORTIE>="'+UsDateTime(StrtoDate(GetControlText('DATEARRET')))+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL))'
            else StWhere := StWhere + 'WHERE PE2_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
            '(PSA_DATESORTIE>="'+UsDateTime(StrtoDate(GetControlText('DATEARRET')))+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL))';
          end;
        end;
        //FIN PT11
        If StWhere <> '' then
        begin
                StWhere := ' ' + StWhere + ' AND PE2_DATEEFFET>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PE2_DATEEFFET<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"';  //PT13
        end
        else StWhere := ' WHERE PE2_DATEEFFET>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PE2_DATEEFFET<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"'; //PT13
//Optimisation        Q := OpenSQL('SELECT * FROM MSAEVOLUTIONSPE2'+StWhere,True);
//        TobEvolutions.LoadDetailDB('LesEvolutions','','',Q,False);
        St := 'SELECT * FROM MSAEVOLUTIONSPE2'+StWhere;
        TobEvolutions := Tob.Create('LesEvolutions',Nil,-1);
        TobEvolutions.LoadDetailDBFromSQL ('LesEvolutions', St);
//        Ferme(Q);
        TobEtat := Tob.Create('Edition',Nil,-1);
        For i := 0 to TobEvolutions.Detail.Count - 1 do
        begin
                TypeEvolution := TobEvolutions.Detail[i].GetValue('PE2_TYPEEVOLMSA');
                If TypeEvolution = 'PE21' then
                begin
                        If TobEvolutions.Detail[i].GetValue('PE2_DEBACTIVITE') <> IDate1900 then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Date début d''activité',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_DEBACTIVITE')));
                        If TobEvolutions.Detail[i].GetValue('PE2_FINACTIVITE') <> IDate1900 then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Date fin d''activité',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_FINACTIVITE')));

                end;
                If TypeEvolution = 'PE22' then
                begin
                        If TobEvolutions.Detail[i].GetValue('PE2_DEBSUSPCT') <> IDate1900 then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Date début suspension de contrat',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_DEBSUSPCT')));
                        If TobEvolutions.Detail[i].GetValue('PE2_FINSUSSPCT') <> IDate1900 then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Date fin suspension de contrat',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_FINSUSSPCT')));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSASUSPCT') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Motif de suspension de contrat',RechDom('PGMSASUSPCONTRAT',TobEvolutions.Detail[i].GetValue('PE2_MSASUSPCT'),False));
                end;
                If TypeEvolution = 'PE23' then
                begin
                        If TobEvolutions.Detail[i].GetValue('PE2_CPLIEUTRAV') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Code postal lieu de travail',TobEvolutions.Detail[i].GetValue('PE2_CPLIEUTRAV'));
                        If TobEvolutions.Detail[i].GetValue('PE2_PCTTPSPART') > 0 then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Pourcentage temps partiel',FloatToStr(TobEvolutions.Detail[i].GetValue('PE2_PCTTPSPART')));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Code temps partiel option temps plein',RechDom('PGMSATEMPSPARTIEL',TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Contrat de travail',RechDom('PGTYPECONTRATMSA',TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_NBHCONV') > 0 then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Nombre d''heures convention',FloatToStr(TobEvolutions.Detail[i].GetValue('PE2_NBHCONV')));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSANIVEAU') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Niveau ou coefficient hiérarchique',TobEvolutions.Detail[i].GetValue('PE2_MSANIVEAU'));
                        //DEB PT13
                        If TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') <> '' then
                        begin
                          if TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') = 'A' then
                            AjoutLigneEvolution (TobEvolutions.Detail[i],'Périodicité de la durée de travail','Annuelle');
                          if TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') = 'M' then
                            AjoutLigneEvolution (TobEvolutions.Detail[i],'Périodicité de la durée de travail','Mensuelle');
                          if TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') = 'H' then
                            AjoutLigneEvolution (TobEvolutions.Detail[i],'Périodicité de la durée de travail','Hebdomadaire');
                        end;
                        If TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE') <> '' then
                          AjoutLigneEvolution (TobEvolutions.Detail[i],'Durée du contrat de travail',TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE') / 100);
                        //FIN PT13
                end;
                If TypeEvolution = 'PE24' then
                begin
                        If TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Personnel technique',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_PERSTECH'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_PERSTECHCUMA') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Personnel technique CUMA',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_PERSTECHCUMA'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_TOPCADRE') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Top cadre (APECITA)',TobEvolutions.Detail[i].GetValue('PE2_TOPCADRE'));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSAPOLYEMP') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Poly-employeur',TobEvolutions.Detail[i].GetValue('PE2_MSAPOLYEMP'));
//PT13                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Poly-employeur',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_MSAPOLYEMP'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_FISCHORSF') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Domicilié fiscalement hors France',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_FISCHORSF'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_CLASSEELEVE') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Code classe élèves',RechDom('PGMSACODECLASSE',TobEvolutions.Detail[i].GetValue('PE2_CLASSEELEVE'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSAUG') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Unité de gestion',TobEvolutions.Detail[i].GetValue('PE2_MSAUG'));
                        If TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Code qualité',TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE'));
                        If TobEvolutions.Detail[i].GetValue('PE2_LIBELLEEMPLOI') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Emploi occupé',TobEvolutions.Detail[i].GetValue('PE2_LIBELLEEMPLOI'));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY1') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'B&énéficie loi Aubry 1',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY1'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY2') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Bénéficie loi Aubry 2',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY2'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_RETRAITEACT') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Retraité en activité',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_RETRAITEACT'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSACONVCOLL') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Convention collective de travail',TobEvolutions.Detail[i].GetValue('PE2_MSACONVCOLL'));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSACODECRCCA') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Code CRCCA',RechDom('PGSCATEGORIELMSA',TobEvolutions.Detail[i].GetValue('PE2_MSACODECRCCA'),False));
                        //DEB PT13
                        If TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Section prud''homale',RechDom('PGSECTIONPRUDMSA',TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART'),False));
                        If TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT') <> '' then
                        AjoutLigneEvolution (TobEvolutions.Detail[i],'Collège prud''homal',RechDom('PGCOLLEGEPRUDMSA',TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT'),False));
                        //FIN PT13
                end;
        end;
        FreeAndNil(TobEvolutions);
        If GetCheckBoxState('CRUPTETAB') = CbChecked then
        begin
                If GetCheckBoxState('CALPHA') = CbChecked then TobEtat.Detail.Sort('PE2_ETABLISSEMENT;PE2_NOM;PE2_SALARIE;PE2_TYPEEVOLMSA')
                else TobEtat.Detail.Sort('PE2_ETABLISSEMENT;PE2_SALARIE;PE2_TYPEEVOLMSA');
        end
        else
        begin
                If GetCheckBoxState('CALPHA') = CbChecked then TobEtat.Detail.Sort('PE2_NOM;PE2_SALARIE;PE2_TYPEEVOLMSA')
                else TobEtat.Detail.Sort('PE2_SALARIE;PE2_TYPEEVOLMSA');
        end;
//        LanceEtatTOB('E','PMS','PEV',TobEtat,True,False,False,Pages,'','',False);
//        TobEtat.Free;
        TFQRS1(Ecran).NatureEtat := 'PMS'; //PT13
        TFQRS1(Ecran).CodeEtat := 'PEV';   //PT13
        TFQRS1(Ecran).LaTob:= TobEtat;
end;

Procedure TOF_PGMSA_DeclSalaire.AjoutLigneEvolution (TobLigneEvolution : Tob;Libelle,Valeur : String);
var TE : Tob;
begin
        TE := Tob.Create('FilleEtat',TobEtat,-1);
        TE.AddchampSupValeur('PE2_ETABLISSEMENT',TobLigneEvolution.GetValue('PE2_ETABLISSEMENT'),False);
        TE.AddchampSupValeur('PE2_SALARIE',TobLigneEvolution.GetValue('PE2_SALARIE'),False);
        TE.AddchampSupValeur('PE2_NOM',TobLigneEvolution.GetValue('PE2_NOM'),False);
        TE.AddchampSupValeur('PE2_PRENOM',TobLigneEvolution.GetValue('PE2_PRENOM'),False);
        TE.AddchampSupValeur('PE2_DATEEFFET',TobLigneEvolution.GetValue('PE2_DATEEFFET'),False);
        TE.AddchampSupValeur('PE2_TYPEEVOLMSA',TobLigneEvolution.GetValue('PE2_TYPEEVOLMSA'),False);
        TE.AddchampSupValeur('LIBELLEMODIF',Libelle,False);
        TE.AddchampSupValeur('VALEURMODIF',Valeur,False);
end;

procedure TOF_PGMSA_DeclSalaire.AffichDeclarant(Sender: TObject);  //PT4
begin
if GetControlText('DECLARANT')='' then exit;
SetControlText('ENOM',RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
SetControlText('ETEL',RechDom('PGDECLARANTTEL  ' ,GetControlText('DECLARANT'),False));
SetControlText('ELIEU',RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
end;

procedure TOF_PGMSA_DeclSalaire.ClickSortie(Sender : TObject);    //PT11
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

{PT19
procedure TOF_PGMSA_DeclSalaire.VerifDateFin(Sender: TObject);
var DateD,DateF : TDateTime;
begin
  If THEdit(Sender) = Nil then Exit;
  DateD := StrToDate(GetControlText('DATEDEBUT'));   //PT13
  DateF := StrToDate(GetControlText('DATEFIN')); //PT13
  If DateF < DateD then
  begin
    PGIBox('La date de fin ne peut pas être inférieure à la date de début',Ecran.Caption);
    SetFocusControl('DATEFIN');
  end;
end;
}

{TOF_PGMSARECAPENVOI}
//DEB PT13
Procedure TOF_PGMSARECAPENVOI.AjoutLigneEvolutionRecap (TobLigneEvolution : Tob;Libelle,Valeur : String);
var TE,TA : Tob;
    Q : TQuery;
begin
        TE := Tob.Create('FilleEtat',TobEtat,-1);
        TE.AddChampSupValeur('ORDRE',1);
        TE.AddChampSupValeur('ENTREPRISE',GetControlText('ESIREN'));
        Q := OpenSQL('SELECT ET_SIRET FROM ETABLISS WHERE ET_ETABLISSEMENT="'+TobLigneEvolution.GetValue('PE2_ETABLISSEMENT')+'"',True);
        TE.AddChampSupValeur('ETABLISSEMENT',Q.FindField('ET_SIRET').AsString);
        ferme(Q);
        TE.AddChampSupValeur('MSAACTIVITE',TobLigneEvolution.GetValue('PE2_MSAACTIVITE'));
        TE.AddchampSupValeur('TYPEINFO','ZZZ');
        TE.AddchampSupValeur('SALARIE',TobLigneEvolution.GetValue('PE2_SALARIE'),False);
        TE.AddChampSupValeur('DATEDEBUT',TobLigneEvolution.GetValue('PE2_DATEEFFET'));
        TE.AddchampSupValeur('NOM',TobLigneEvolution.GetValue('PE2_NOM'),False);
        TE.AddchampSupValeur('PRENOM',TobLigneEvolution.GetValue('PE2_PRENOM'),False);
        TE.AddchampSupValeur('NUMSS',TobLigneEvolution.GetValue('PE2_NUMEROSS'),False);
        TE.AddchampSupValeur('DATENAISS',DateToStr(TobLigneEvolution.GetValue('PE2_DATENAISSANCE')),False);
        TE.AddchampSupValeur('ACTIVITE',rechDom('PGMSAACTIVITE',TobLigneEvolution.GetValue('PE2_MSAACTIVITE'),False),False);
        TE.AddchampSupValeur('UG',TobLigneEvolution.GetValue('PE2_MSAUG'),False);
        TE.AddChampSupValeur('EXO','',False);
        TE.AddchampSupValeur('COLONNE1',Rechdom('PGMSATYPEEVOLUTION',TobLigneEvolution.GetValue('PE2_TYPEEVOLMSA'),False),False);
        TE.AddchampSupValeur('COLONNE2',DateToStr(TobLigneEvolution.GetValue('PE2_DATEEFFET')),False);
        TE.AddchampSupValeur('COLONNE3',Libelle,False);
        TE.AddchampSupValeur('COLONNE4',Valeur,False);
        TA := TobAssiettes.FindFirst(['ETABLISSEMENT','ACTIVITE'],[TobLigneEvolution.GetValue('PE2_ETABLISSEMENT'),TobLigneEvolution.GetValue('PE2_MSAACTIVITE')],False);
       If TA <> Nil then
       begin
           TE.AddchampSupValeur('ELTENTREPRISE1','51 tcp');
           TE.AddchampSupValeur('ELTENTREPRISE2','53 CSG');
           TE.AddchampSupValeur('ELTENTREPRISE3','54 CSG/RDS NON IMP.');
           TE.AddchampSupValeur('ELTENTREPRISE4','56 CSG/RDS IMP.');
           TE.AddchampSupValeur('ELTENTREPRISE5','63 CSG/RDS IMP. 6,6 %');
           //DEB PT18
           TE.AddchampSupValeur('ELTENTREPRISE6','66 Contribution pat. Indemnité retraite');
           TE.AddchampSupValeur('ELTENTREPRISE7','67 Assiette CSG 7,5% et rds >=11/10/2007');
           TE.AddchampSupValeur('ELTENTREPRISE8','68 Contribtution préretraites >=11/10/2007');
           //FIN PT18
           TE.AddchampSupValeur('ELTENTREPRISE9','58 Contr. épargne salariale');
           TE.AddchampSupValeur('ELTENTREPRISE10','59 Contr. préretraites tx normal');
{PT18           TE.AddchampSupValeur('ELTENTREPRISE8','60 Contr. préretraites tx réduit');}
           TE.AddchampSupValeur('ELTENTREPRISE11','61 Contr. retraites préstations rentes');
           TE.AddchampSupValeur('ELTENTREPRISE12','62 Contr. retraites préstations primes');
           TE.AddchampSupValeur('VALELTENTREPRISE1',TA.GetValue('CSG'));
           TE.AddchampSupValeur('VALELTENTREPRISE2',TA.GetValue('TCP'));
           TE.AddchampSupValeur('VALELTENTREPRISE3',TA.GetValue('CSGNIMP'));
           TE.AddchampSupValeur('VALELTENTREPRISE4',TA.GetValue('CSGIMP'));
           TE.AddchampSupValeur('VALELTENTREPRISE5',TA.GetValue('63'));
           //DEB PT18
           TE.AddchampSupValeur('VALELTENTREPRISE6',TA.GetValue('66'));
           TE.AddchampSupValeur('VALELTENTREPRISE7',TA.GetValue('67'));
           TE.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('68'));
           //FIN PT18
           TE.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('58'));
           TE.AddchampSupValeur('VALELTENTREPRISE10',TA.GetValue('59'));
{PT18           TE.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('60')); }
           TE.AddchampSupValeur('VALELTENTREPRISE11',TA.GetValue('61'));
           TE.AddchampSupValeur('VALELTENTREPRISE12',TA.GetValue('62'));
       end
       else
       begin
           TE.AddchampSupValeur('ELTENTREPRISE1','');
           TE.AddchampSupValeur('ELTENTREPRISE2','');
           TE.AddchampSupValeur('ELTENTREPRISE3','');
           TE.AddchampSupValeur('ELTENTREPRISE4','');
           TE.AddchampSupValeur('ELTENTREPRISE5','');
           TE.AddchampSupValeur('ELTENTREPRISE6','');
           TE.AddchampSupValeur('ELTENTREPRISE7','');
           TE.AddchampSupValeur('ELTENTREPRISE8','');
           TE.AddchampSupValeur('ELTENTREPRISE9','');
           TE.AddchampSupValeur('ELTENTREPRISE10','');
           TE.AddchampSupValeur('ELTENTREPRISE11','');  //PT18
           TE.AddchampSupValeur('ELTENTREPRISE12','');  //PT18
           TE.AddchampSupValeur('VALELTENTREPRISE1','');
           TE.AddchampSupValeur('VALELTENTREPRISE2','');
           TE.AddchampSupValeur('VALELTENTREPRISE3','');
           TE.AddchampSupValeur('VALELTENTREPRISE4','');
           TE.AddchampSupValeur('VALELTENTREPRISE5','');
           TE.AddchampSupValeur('VALELTENTREPRISE6','');
           TE.AddchampSupValeur('VALELTENTREPRISE7','');
           TE.AddchampSupValeur('VALELTENTREPRISE8','');
           TE.AddchampSupValeur('VALELTENTREPRISE9','');
           TE.AddchampSupValeur('VALELTENTREPRISE10','');
           TE.AddchampSupValeur('VALELTENTREPRISE11',''); //PT18
           TE.AddchampSupValeur('VALELTENTREPRISE12',''); //PT18
       end;
end;

procedure TOF_PGMSARECAPENVOI.edition;
var T,TPeriodes,TobEvolutions,TLA : Tob;
    i,x,a,u,r : Integer;
    Q : TQuery;
    TypeEvolution,Etab,Activite,UGMsa : String;
    DD,DF : TDateTime;
    TobEtab,TobActivite,TobUG,TG,TA : Tob;
    CalculerTcp : Boolean;
    Montant : Double;
    Pages : TPageControl;
    St : String;
begin
     Inherited ;
     Pages := TPageControl(GetControl('PAGES'));
     If GetCheckBoxState('CFICHIER') = CbChecked then EditionFichierMsa
     else
     begin
          DD := StrToDate(GetControlText('DATEDEBUT'));
          DF := StrToDate(GetControlText('DATEFIN'));
          CalculerTcp := GetParamSocSecur ('SO_PGMSACALCTCP',False);
//Optimisation          Q := OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS', True);
          St := 'SELECT ET_ETABLISSEMENT FROM ETABLISS';
          TobEtab := Tob.Create('LesEtablissements', nil, -1);
          TobEtab.LoadDetailDBFromSQL ('LesEtablissements', St);
//          TobEtab.LoadDetailDB('LesEtablissements', '', '', Q, False);
//          Ferme(Q);
          TobAssiettes := Tob.Create('LesAssiettes', nil, -1);
          for i := 2 to GRem.RowCount - 1 do
          begin
               TLA := Tob.Create('UneAssiette', TobAssiettes, -1);
               TLA.AddChampSupValeur('ETABLISSEMENT', Grem.CellValues[0, i], False);
               TLA.AddChampSupValeur('ACTIVITE', Grem.CellValues[1, i], False);
               TLA.AddChampSupValeur('UG', Grem.CellValues[2, i], False);
               if isNumeric(Grem.CellValues[3, i]) then TLA.AddChampSupValeur('TCP', StrToFloat(Grem.CellValues[3, i]), False)
               else TLA.AddChampSupValeur('TCP', 0, False);
               if isNumeric(Grem.CellValues[4, i]) then TLA.AddChampSupValeur('CSG', StrToFloat(Grem.CellValues[4, i]), False)
               else TLA.AddChampSupValeur('CSG', 0, False);
               if isNumeric(Grem.CellValues[5, i]) then TLA.AddChampSupValeur('CSGNIMP', StrToFloat(Grem.CellValues[5, i]), False)
               else TLA.AddChampSupValeur('CSGNIMP', 0, False);
               if isNumeric(Grem.CellValues[6, i]) then TLA.AddChampSupValeur('CSGIMP', StrToFloat(Grem.CellValues[6, i]), False)
               else TLA.AddChampSupValeur('CSGIMP', 0, False);
               if isNumeric(Grem.CellValues[7, i]) then TLA.AddChampSupValeur('63', StrToFloat(Grem.CellValues[7, i]), False)
               else TLA.AddChampSupValeur('63', 0, False);
               //DEB PT18
               if isNumeric(Grem.CellValues[8, i]) then TLA.AddChampSupValeur('66', StrToFloat(Grem.CellValues[8, i]), False)
               else TLA.AddChampSupValeur('66', 0, False);
               if isNumeric(Grem.CellValues[9, i]) then TLA.AddChampSupValeur('67', StrToFloat(Grem.CellValues[9, i]), False)
               else TLA.AddChampSupValeur('67', 0, False);
               if isNumeric(Grem.CellValues[10, i]) then TLA.AddChampSupValeur('68', StrToFloat(Grem.CellValues[10, i]), False)
               else TLA.AddChampSupValeur('68', 0, False);
               //FIN PT18
               if isNumeric(Grem.CellValues[11, i]) then TLA.AddChampSupValeur('58', StrToFloat(Grem.CellValues[11, i]), False)
               else TLA.AddChampSupValeur('58', 0, False);
               if isNumeric(Grem.CellValues[12, i]) then TLA.AddChampSupValeur('59', StrToFloat(Grem.CellValues[12, i]), False)
               else TLA.AddChampSupValeur('59', 0, False);
{PT18               if isNumeric(Grem.CellValues[10, i]) then TLA.AddChampSupValeur('60', StrToFloat(Grem.CellValues[10, i]), False)
               else TLA.AddChampSupValeur('60', 0, False);}
               if isNumeric(Grem.CellValues[13, i]) then TLA.AddChampSupValeur('61', StrToFloat(Grem.CellValues[13, i]), False)
               else TLA.AddChampSupValeur('61', 0, False);
               if isNumeric(Grem.CellValues[14, i]) then TLA.AddChampSupValeur('62', StrToFloat(Grem.CellValues[14, i]), False)
               else TLA.AddChampSupValeur('62', 0, False);
          end;

         If TobEtat <> Nil then FreeAndNil(TobEtat);
         TobEtat := Tob.Create('FilleEtat',Nil,-1);
//Optimisation         Q := OpenSQL('SELECT * FROM MSAPERIODESPE31 WHERE PE3_DATEDEBUT>="'+UsDateTime(DD)+'" AND PE3_DATEDEBUT<="'+UsDateTime(DF)+'"',True);
         St := 'SELECT * FROM MSAPERIODESPE31 WHERE PE3_DATEDEBUT>="'+UsDateTime(DD)+'"' +
           ' AND PE3_DATEDEBUT<="'+UsDateTime(DF)+'" AND PE3_ORDRE < 100';
         TPeriodes := Tob.Create('LesPeriodes',Nil,-1);
         TPeriodes.LoadDetailDBFromSQL ('LesPeriodes', St);
//         TPeriodes.LoadDetailDB('LesPeriodes','','',Q,False);
//         Ferme(Q);
         For i := 0 to TPeriodes.Detail.Count - 1 do
         begin
              For x := 1 to 5 do
              begin
                   If TPeriodes.Detail[i].GetValue('PE3_ELTCALCMSA'+IntToStr(x)) <> ''then
                   begin
                        T := Tob.Create('FilleEtat',TobEtat,-1);
                        T.AddChampSupValeur('ORDRE',1);
                        T.AddChampSupValeur('ENTREPRISE',GetControlText('ESIREN'));
                        Q := OpenSQL('SELECT ET_SIRET FROM ETABLISS WHERE ET_ETABLISSEMENT="'+TPeriodes.Detail[i].GetValue('PE3_ETABLISSEMENT')+'"',True);
                        T.AddChampSupValeur('ETABLISSEMENT',Q.FindField('ET_SIRET').AsString);
                        ferme(Q);
                        T.AddChampSupValeur('MSAACTIVITE',TPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE'));
                        T.AddChampSupValeur('TYPEINFO','PER');
                        T.AddChampSupValeur('SALARIE',TPeriodes.Detail[i].GetValue('PE3_SALARIE'));
                        T.AddChampSupValeur('DATEDEBUT',TPeriodes.Detail[i].GetValue('PE3_DATEDEBUT'));
                        T.AddChampSupValeur('NOM',TPeriodes.Detail[i].GetValue('PE3_NOM'));
                        T.AddChampSupValeur('PRENOM',TPeriodes.Detail[i].GetValue('PE3_PRENOM'));
                        T.AddChampSupValeur('NUMSS',TPeriodes.Detail[i].GetValue('PE3_NUMEROSS'));
                        T.AddChampSupValeur('DATENAISS',DateToStr(TPeriodes.Detail[i].GetValue('PE3_DATENAISSANCE')));
                        T.AddchampSupValeur('ACTIVITE',rechDom('PGMSAACTIVITE',TPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE'),False),False);
                        T.AddChampSupValeur('UG',TPeriodes.Detail[i].GetValue('PE3_UGMSA'));
                        T.AddChampSupValeur('EXO',TPeriodes.Detail[i].GetValue('PE3_PEXOMSA'));
                        T.AddChampSupValeur('COLONNE1',DateToStr(TPeriodes.Detail[i].GetValue('PE3_DATEDEBUT')) + ' - '+  DateToStr(TPeriodes.Detail[i].GetValue('PE3_DATEFIN')));
                        If TPeriodes.Detail[i].GetValue('PE3_NBJOURS') > 0 then T.AddChampSupValeur('COLONNE2',FloatToStr(TPeriodes.Detail[i].GetValue('PE3_NBHEURES'))+ ' / '+IntToStr(TPeriodes.Detail[i].GetValue('PE3_NBJOURS')))
                        else T.AddChampSupValeur('COLONNE2',FloatToStr(TPeriodes.Detail[i].GetValue('PE3_NBHEURES')));
                        T.AddChampSupValeur('COLONNE3',TPeriodes.Detail[i].GetValue('PE3_ELTCALCMSA'+IntToStr(x)) + ' '+RechDom('PGMSATYPEELEMENTCALC',TPeriodes.Detail[i].GetValue('PE3_ELTCALCMSA'+IntToStr(x)),False));
                        T.AddChampSupValeur('COLONNE4',TPeriodes.Detail[i].GetValue('PE3_MONTANTELT'+IntToStr(x)));
                        TA := TobAssiettes.FindFirst(['ETABLISSEMENT','ACTIVITE'],[TPeriodes.Detail[i].GetValue('PE3_ETABLISSEMENT'),TPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE')],False);
                         If TA <> Nil then
                         begin
                             T.AddchampSupValeur('ELTENTREPRISE1','51 tcp');
                             T.AddchampSupValeur('ELTENTREPRISE2','53 CSG');
                             T.AddchampSupValeur('ELTENTREPRISE3','54 CSG/RDS NON IMP. 3,8 %');
                             T.AddchampSupValeur('ELTENTREPRISE4','56 CSG/RDS IMP. 6,2 %');
                             T.AddchampSupValeur('ELTENTREPRISE5','63 CSG/RDS IMP. 6,6 %');
                             //DEB PT18
                             T.AddchampSupValeur('ELTENTREPRISE6','66 Contribution pat. Indemnité retraite');
                             T.AddchampSupValeur('ELTENTREPRISE7','67 Assiette CSG 7,5% et rds >=11/10/2007');
                             T.AddchampSupValeur('ELTENTREPRISE8','68 Contribtution préretraites >=11/10/2007');
                             //FIN PT18
                             T.AddchampSupValeur('ELTENTREPRISE9','58 Contr. épargane salariale');
                             T.AddchampSupValeur('ELTENTREPRISE10','59 Contr. préretraites tx normal');
{PT18                             T.AddchampSupValeur('ELTENTREPRISE8','60 Contr. préretraites tx réduit');}
                             T.AddchampSupValeur('ELTENTREPRISE11','61 Contr. retraites préstations rentes');
                             T.AddchampSupValeur('ELTENTREPRISE12','62 Contr. retraites préstations primes');
                             T.AddchampSupValeur('VALELTENTREPRISE1',TA.GetValue('TCP'));
                             T.AddchampSupValeur('VALELTENTREPRISE2',TA.GetValue('CSG'));
                             T.AddchampSupValeur('VALELTENTREPRISE3',TA.GetValue('CSGNIMP'));
                             T.AddchampSupValeur('VALELTENTREPRISE4',TA.GetValue('CSGIMP'));
                             T.AddchampSupValeur('VALELTENTREPRISE5',TA.GetValue('63'));
                             //DEB PT18
                             T.AddchampSupValeur('VALELTENTREPRISE6',TA.GetValue('66'));
                             T.AddchampSupValeur('VALELTENTREPRISE7',TA.GetValue('67'));
                             T.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('68'));
                             //FIN PT18
                             T.AddchampSupValeur('VALELTENTREPRISE9',TA.GetValue('58'));
                             T.AddchampSupValeur('VALELTENTREPRISE10',TA.GetValue('59'));
{PT18                             T.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('60'));}
                             T.AddchampSupValeur('VALELTENTREPRISE11',TA.GetValue('61'));
                             T.AddchampSupValeur('VALELTENTREPRISE12',TA.GetValue('62'));
                         end
                         else
                         begin
                             T.AddchampSupValeur('ELTENTREPRISE1','');
                             T.AddchampSupValeur('ELTENTREPRISE2','');
                             T.AddchampSupValeur('ELTENTREPRISE3','');
                             T.AddchampSupValeur('ELTENTREPRISE4','');
                             T.AddchampSupValeur('ELTENTREPRISE5','');
                             T.AddchampSupValeur('ELTENTREPRISE6','');
                             T.AddchampSupValeur('ELTENTREPRISE7','');
                             T.AddchampSupValeur('ELTENTREPRISE8','');
                             T.AddchampSupValeur('ELTENTREPRISE9','');
                             T.AddchampSupValeur('ELTENTREPRISE10','');
                             T.AddchampSupValeur('ELTENTREPRISE11','');  //PT18
                             T.AddchampSupValeur('ELTENTREPRISE12','');  //PT18
                             T.AddchampSupValeur('VALELTENTREPRISE1','');
                             T.AddchampSupValeur('VALELTENTREPRISE2','');
                             T.AddchampSupValeur('VALELTENTREPRISE3','');
                             T.AddchampSupValeur('VALELTENTREPRISE4','');
                             T.AddchampSupValeur('VALELTENTREPRISE5','');
                             T.AddchampSupValeur('VALELTENTREPRISE6','');
                             T.AddchampSupValeur('VALELTENTREPRISE7','');
                             T.AddchampSupValeur('VALELTENTREPRISE8','');
                             T.AddchampSupValeur('VALELTENTREPRISE9','');
                             T.AddchampSupValeur('VALELTENTREPRISE10','');
                             T.AddchampSupValeur('VALELTENTREPRISE11',''); //PT18
                             T.AddchampSupValeur('VALELTENTREPRISE12',''); //PT18
                         end;
                   end;
              end;
         end;
         FreeAndNil(Tperiodes);

        //DEB PT13
        St := 'SELECT * FROM MSAPERIODESPE31 WHERE PE3_DATEDEBUT>="'+UsDateTime(DD)+'"' +
           ' AND PE3_DATEDEBUT<="'+UsDateTime(DF)+'" AND PE3_ORDRE >= 100 ' +
           ' ORDER BY PE3_SALARIE,PE3_DATEDEBUT,PE3_DATEFIN,PE3_ORDRE';
        Tperiodes := Tob.Create('LesPeriodes',Nil,-1);
        Tperiodes.LoadDetailDBFromSQL ('LesPeriodes', St);
        r := 1;
        For i := 0 to Tperiodes.Detail.Count - 1 do
        begin
          if Tperiodes.Detail[i].GetValue('PE3_ORDRE') = 100 then
            r := 1;
          T := Tob.Create('FilleEtat',TobEtat,-1);
          T.AddChampSupValeur('ORDRE',Tperiodes.Detail[i].GetValue('PE3_ORDRE'));
          T.AddChampSupValeur('ENTREPRISE',GetControlText('ESIREN'));
          Q := OpenSQL('SELECT ET_SIRET FROM ETABLISS WHERE ET_ETABLISSEMENT="'+TPeriodes.Detail[i].GetValue('PE3_ETABLISSEMENT')+'"',True);
          if not Q.Eof then
            T.AddChampSupValeur('ETABLISSEMENT',Q.FindField('ET_SIRET').AsString)
          else
            T.AddChampSupValeur('ETABLISSEMENT','');
          ferme(Q);
          T.AddChampSupValeur('MSAACTIVITE',TPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE'));
          T.AddChampSupValeur('TYPEINFO','PER');
          T.AddChampSupValeur('SALARIE',TPeriodes.Detail[i].GetValue('PE3_SALARIE'));
          T.AddChampSupValeur('DATEDEBUT',TPeriodes.Detail[i].GetValue('PE3_DATEDEBUT'));
          T.AddChampSupValeur('NOM',TPeriodes.Detail[i].GetValue('PE3_NOM'));
          T.AddChampSupValeur('PRENOM',TPeriodes.Detail[i].GetValue('PE3_PRENOM'));
          T.AddChampSupValeur('NUMSS',TPeriodes.Detail[i].GetValue('PE3_NUMEROSS'));
          T.AddChampSupValeur('DATENAISS',DateToStr(TPeriodes.Detail[i].GetValue('PE3_DATENAISSANCE')));
          T.AddchampSupValeur('ACTIVITE',rechDom('PGMSAACTIVITE',TPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE'),False),False);
          T.AddChampSupValeur('UG',TPeriodes.Detail[i].GetValue('PE3_UGMSA'));
          T.AddChampSupValeur('EXO',TPeriodes.Detail[i].GetValue('PE3_PEXOMSA'));
          if r = 1 then
            T.AddChampSupValeur('COLONNE1','Heures supplémentaires')
          else
            T.AddChampSupValeur('COLONNE1','');
          T.AddchampSupValeur('COLONNE2',StrToInt(TPeriodes.Detail[i].GetValue('PE3_MONTANTELT3')) / 100);
          T.AddChampSupValeur('COLONNE3','Type ' + IntToStr(TPeriodes.Detail[i].GetValue('PE3_MONTANTELT1')) + ' - Taux de majoration ' + IntToStr(r) + ' à ' + IntToStr(TPeriodes.Detail[i].GetValue('PE3_MONTANTELT2')) + ' %');
          T.AddChampSupValeur('COLONNE4',TPeriodes.Detail[i].GetValue('PE3_MONTANTELT4'));
          //DEB PT15
          TA := TobAssiettes.FindFirst(['ETABLISSEMENT','ACTIVITE'],[TPeriodes.Detail[i].GetValue('PE3_ETABLISSEMENT'),TPeriodes.Detail[i].GetValue('PE3_MSAACTIVITE')],False);
           If TA <> Nil then
           begin
               T.AddchampSupValeur('ELTENTREPRISE1','51 tcp');
               T.AddchampSupValeur('ELTENTREPRISE2','53 CSG');
               T.AddchampSupValeur('ELTENTREPRISE3','54 CSG/RDS NON IMP. 3,8 %');
               T.AddchampSupValeur('ELTENTREPRISE4','56 CSG/RDS IMP. 6,2 %');
               T.AddchampSupValeur('ELTENTREPRISE5','63 CSG/RDS IMP. 6,6 %');
               //DEB PT18
               T.AddchampSupValeur('ELTENTREPRISE6','66 Contribution pat. Indemnité retraite');
               T.AddchampSupValeur('ELTENTREPRISE7','67 Assiette CSG 7,5% et rds >=11/10/2007');
               T.AddchampSupValeur('ELTENTREPRISE8','68 Assiette CSG 7,5% et rds >=11/10/2007');
               //FIN PT18
               T.AddchampSupValeur('ELTENTREPRISE9','58 Contr. épargane salariale');
               T.AddchampSupValeur('ELTENTREPRISE10','59 Contr. préretraites tx normal');
{PT18               T.AddchampSupValeur('ELTENTREPRISE8','60 Contr. préretraites tx réduit');}
               T.AddchampSupValeur('ELTENTREPRISE11','61 Contr. retraites préstations rentes');
               T.AddchampSupValeur('ELTENTREPRISE12','62 Contr. retraites préstations primes');
               T.AddchampSupValeur('VALELTENTREPRISE1',TA.GetValue('TCP'));
               T.AddchampSupValeur('VALELTENTREPRISE2',TA.GetValue('CSG'));
               T.AddchampSupValeur('VALELTENTREPRISE3',TA.GetValue('CSGNIMP'));
               T.AddchampSupValeur('VALELTENTREPRISE4',TA.GetValue('CSGIMP'));
               T.AddchampSupValeur('VALELTENTREPRISE5',TA.GetValue('63'));
               //DEB PT18
               T.AddchampSupValeur('VALELTENTREPRISE6',TA.GetValue('66'));
               T.AddchampSupValeur('VALELTENTREPRISE7',TA.GetValue('67'));
               T.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('68'));
               //FIN PT18
               T.AddchampSupValeur('VALELTENTREPRISE9',TA.GetValue('58'));
               T.AddchampSupValeur('VALELTENTREPRISE10',TA.GetValue('59'));
{PT18               T.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('60'));}
               T.AddchampSupValeur('VALELTENTREPRISE11',TA.GetValue('61'));
               T.AddchampSupValeur('VALELTENTREPRISE12',TA.GetValue('62'));
           end
           else
           begin
               T.AddchampSupValeur('ELTENTREPRISE1','');
               T.AddchampSupValeur('ELTENTREPRISE2','');
               T.AddchampSupValeur('ELTENTREPRISE3','');
               T.AddchampSupValeur('ELTENTREPRISE4','');
               T.AddchampSupValeur('ELTENTREPRISE5','');
               T.AddchampSupValeur('ELTENTREPRISE6','');
               T.AddchampSupValeur('ELTENTREPRISE7','');
               T.AddchampSupValeur('ELTENTREPRISE8','');
               T.AddchampSupValeur('ELTENTREPRISE9','');
               T.AddchampSupValeur('ELTENTREPRISE10','');
               T.AddchampSupValeur('ELTENTREPRISE11','');  //PT18
               T.AddchampSupValeur('ELTENTREPRISE12','');  //PT18
               T.AddchampSupValeur('VALELTENTREPRISE1','');
               T.AddchampSupValeur('VALELTENTREPRISE2','');
               T.AddchampSupValeur('VALELTENTREPRISE3','');
               T.AddchampSupValeur('VALELTENTREPRISE4','');
               T.AddchampSupValeur('VALELTENTREPRISE5','');
               T.AddchampSupValeur('VALELTENTREPRISE6','');
               T.AddchampSupValeur('VALELTENTREPRISE7','');
               T.AddchampSupValeur('VALELTENTREPRISE8','');
               T.AddchampSupValeur('VALELTENTREPRISE9','');
               T.AddchampSupValeur('VALELTENTREPRISE10','');
               T.AddchampSupValeur('VALELTENTREPRISE11',''); //PT18
               T.AddchampSupValeur('VALELTENTREPRISE12',''); //PT18
           end;
           //FIN PT15
          r := r + 1;
        end;
        FreeAndNil(Tperiodes);
        //FIN PT13

//Optimisation         Q := OpenSQL('SELECT * FROM MSAEVOLUTIONSPE2 WHERE PE2_DATEEFFET>="'+UsDateTime(DD)+'" AND PE2_DATEEFFET<="'+UsDateTime(DF)+'"',True);
         St := 'SELECT * FROM MSAEVOLUTIONSPE2 WHERE PE2_DATEEFFET>="'+UsDateTime(DD)+'" AND PE2_DATEEFFET<="'+UsDateTime(DF)+'"';
         TobEvolutions := Tob.Create('LesEvolutions',Nil,-1);
         TobEvolutions.LoadDetailDBFromSQL ('LesEvolutions', St);
//         TobEvolutions.LoadDetailDB('LesEvolutions','','',Q,False);
//         Ferme(Q);
         For i := 0 to TobEvolutions.Detail.Count - 1 do
         begin
                    TypeEvolution := TobEvolutions.Detail[i].GetValue('PE2_TYPEEVOLMSA');
                    If TypeEvolution = 'PE21' then
                    begin
                            If TobEvolutions.Detail[i].GetValue('PE2_DEBACTIVITE') <> IDate1900 then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Date début d''activité',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_DEBACTIVITE')));
                            If TobEvolutions.Detail[i].GetValue('PE2_FINACTIVITE') <> IDate1900 then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Date fin d''activité',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_FINACTIVITE')));

                    end;
                    If TypeEvolution = 'PE22' then
                    begin
                            If TobEvolutions.Detail[i].GetValue('PE2_DEBSUSPCT') <> IDate1900 then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Date début suspension de contrat',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_DEBSUSPCT')));
                            If TobEvolutions.Detail[i].GetValue('PE2_FINSUSSPCT') <> IDate1900 then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Date fin suspension de contrat',DateToStr(TobEvolutions.Detail[i].GetValue('PE2_FINSUSSPCT')));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSASUSPCT') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Motif de suspension de contrat',RechDom('PGMSASUSPCONTRAT',TobEvolutions.Detail[i].GetValue('PE2_MSASUSPCT'),False));
                    end;
                    If TypeEvolution = 'PE23' then
                    begin
                            If TobEvolutions.Detail[i].GetValue('PE2_CPLIEUTRAV') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Code postal lieu de travail',TobEvolutions.Detail[i].GetValue('PE2_CPLIEUTRAV'));
                            If TobEvolutions.Detail[i].GetValue('PE2_PCTTPSPART') > 0 then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Pourcentage temps partiel',FloatToStr(TobEvolutions.Detail[i].GetValue('PE2_PCTTPSPART')));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Code temps partiel option temps plein',RechDom('PGMSATEMPSPARTIEL',TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Contrat de travail',RechDom('PGTYPECONTRATMSA',TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_NBHCONV') > 0 then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Nombre d''heures convention',FloatToStr(TobEvolutions.Detail[i].GetValue('PE2_NBHCONV')));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSANIVEAU') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Niveau ou coefficient hiérarchique',TobEvolutions.Detail[i].GetValue('PE2_MSANIVEAU'));
                            //DEB PT13
                            If TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') <> '' then
                            begin
                              if TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') = 'A' then
                                AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Périodicité de la durée de travail','Annuelle');
                              if TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') = 'M' then
                                AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Périodicité de la durée de travail','Mensuelle');
                              if TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') = 'H' then
                                AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Périodicité de la durée de travail','Hebdomadaire');
                            end;
                            If TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE') <> '' then
                              AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Durée du contrat de travail',TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE') / 100);
                            //FIN PT13
                    end;
                    If TypeEvolution = 'PE24' then
                    begin
                            If TobEvolutions.Detail[i].GetValue('PE2_PERSTECH') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Personnel technique',TobEvolutions.Detail[i].GetValue('PE2_PERSTECH'));
                            If TobEvolutions.Detail[i].GetValue('PE2_PERSTECHCUMA') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Personnel technique CUMA',TobEvolutions.Detail[i].GetValue('PE2_PERSTECHCUMA'));
                            If TobEvolutions.Detail[i].GetValue('PE2_TOPCADRE') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Top cadre (APECITA)',TobEvolutions.Detail[i].GetValue('PE2_TOPCADRE'));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSAPOLYEMP') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Poly-employeur',TobEvolutions.Detail[i].GetValue('PE2_MSAPOLYEMP'));
//PT13                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Poly-employeur',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_MSAPOLYEMP'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_FISCHORSF') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Domicilié fiscalement hors France',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_FISCHORSF'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_CLASSEELEVE') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Code classe élèves',RechDom('PGMSACODECLASSE',TobEvolutions.Detail[i].GetValue('PE2_CLASSEELEVE'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSAUG') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Unité de gestion',TobEvolutions.Detail[i].GetValue('PE2_MSAUG'));
                            If TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Code qualité',TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE'));
                            If TobEvolutions.Detail[i].GetValue('PE2_LIBELLEEMPLOI') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Emploi occupé',TobEvolutions.Detail[i].GetValue('PE2_LIBELLEEMPLOI'));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY1') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'B&énéficie loi Aubry 1',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY1'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY2') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Bénéficie loi Aubry 2',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY2'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_RETRAITEACT') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Retraité en activité',RechDom('PGOUINON',TobEvolutions.Detail[i].GetValue('PE2_RETRAITEACT'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSACONVCOLL') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Convention collective de travail',TobEvolutions.Detail[i].GetValue('PE2_MSACONVCOLL'));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSACODECRCCA') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Code CRCCA',TobEvolutions.Detail[i].GetValue('PE2_MSACODECRCCA')+ ' '+RechDom('PGSCATEGORIELMSA',TobEvolutions.Detail[i].GetValue('PE2_MSACODECRCCA'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Section prud''homale',TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART')+ ' '+RechDom('PGSECTIONPRUDMSA',TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART'),False));
                            If TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT') <> '' then
                            AjoutLigneEvolutionRecap (TobEvolutions.Detail[i],'Collège prud''homal',TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT')+ ' '+RechDom('PGCOLLEGEPRUDMSA',TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT'),False));
                    end;
         end;
         FreeAndNil(TobEvolutions);
         FreeAndNil(TobAssiettes);
     end;
     TobEtat.Detail.Sort('ENTREPRISE;TYPEINFO;ETABLISSEMENT;MSAACTIVITE;SALARIE;DATEDEBUT;ORDRE;COLONNE3');
     LanceEtatTOB('E','PMS','PMG',TobEtat,True,False,False,Pages,'','',False);
     FreeAndNil(TobEtab);
end;
procedure TOF_PGMSARECAPENVOI.OnClose ;
begin
     Inherited ;
     If TobEtat <> Nil then FreeAndNil(TobEtat);
end;

procedure TOF_PGMSARECAPENVOI.OnArgument (S : String ) ;
var
  Arg,St : string;
  QAttes : TQuery;
  Edit : THEdit;
  aa,mm,jj,Mois : word;
  DateDebut,DateFin : TDateTime;
  Bt : TToolBarButton97;
begin
  inherited;
  Bt := TToolBarButton97(GetControl('BMAJPER'));
  If Bt <> Nil then Bt.OnClick := MajPeriodes;
  DecodeDate(V_PGI.dateEntree,aa,mm,jj);
        Mois:=1;
        If mm<4 Then
        begin
                aa:=aa-1;
                Mois:=10;
        end;
        If (mm<7) and (mm>3) Then Mois:=1;
        If (mm<10) and (mm>6) Then Mois:=4;
        If mm>9 Then Mois:=7;
        DateDebut:=EncodeDate(aa,Mois,1);
        DateFin:=PlusMois(DateDebut,2);
        DateFin:=FinDeMois(DateFin);
        SetControlText('DATEDEBUT',DateToStr(DateDebut));
        SetControlText('DATEFIN',DateToStr(DateFin));
  Arg := S;
  St := 'PDA_ETABLISSEMENT = "" ' +
        ' AND (PDA_TYPEATTEST = "" OR PDA_TYPEATTEST LIKE "%MSA%" )  '; //PT1
  SetControlProperty('DECLARANT', 'Plus', St);
  QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
      'WHERE (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MSA%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
  if not QAttes.eof then
  begin
        SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
        AffichDeclarant(nil);
  end;
  Edit := THEdit(GetControl('DECLARANT'));
  If Edit<>Nil Then Edit.OnExit := AffichDeclarant;
  TFVierge(Ecran).BValider.OnClick := Edition;
  GRem := THGrid(GetControl('GREM'));
  if GRem <> nil then
  begin
    GRem.ColAligns[3] := taRightJustify;
    GRem.ColAligns[4] := taRightJustify;
    GRem.ColAligns[5] := taRightJustify;
    GRem.ColAligns[6] := taRightJustify;
    GRem.ColAligns[7] := taRightJustify;
    GRem.ColAligns[8] := taRightJustify;
    GRem.ColAligns[9] := taRightJustify;
    GRem.ColAligns[10] := taRightJustify;
    GRem.ColAligns[11] := taRightJustify;
    GRem.ColAligns[12] := taRightJustify;
    GRem.ColAligns[13] := taRightJustify; //PT18
    GRem.ColAligns[14] := taRightJustify; //PT18
    GRem.HideSelectedWhenInactive := true;
    GRem.ColWidths[0] := 125;
    GRem.ColWidths[1] := 125;
    GRem.ColWidths[2] := 60;
    GRem.ColWidths[3] := 80;
    GRem.ColWidths[4] := 80;
    GRem.ColWidths[5] := 80;
    GRem.ColWidths[6] := 80;
    GRem.ColWidths[7] := 80;
    GRem.ColWidths[8] := 80;
    GRem.ColWidths[9] := 80;
    GRem.ColWidths[10] := 80;
    GRem.ColWidths[11] := 80;
    GRem.ColWidths[11] := 80;  //PT18
    GRem.ColWidths[12] := 80;  //PT18
    GRem.CellValues[4, 1] := 'cas standard';
    GRem.CellValues[5, 1] := 'rempl. non imposés';
    GRem.CellValues[6, 1] := 'rempl. imposés 6,2';
    GRem.CellValues[7, 1] := 'rempl. imposés 6,6';
    //DEB PT18
    GRem.CellValues[8, 1] := '66';
    GRem.CellValues[9, 1] := '67';
    GRem.CellValues[10, 1] := '68';
    //FIN PT18
    GRem.CellValues[11, 1] := '58';
    GRem.CellValues[12, 1] := '59';
{PT18    GRem.CellValues[10, 1] := '60';}
    GRem.CellValues[13, 1] := '61';
    GRem.CellValues[14, 1] := '62';
    GRem.ColFormats[0] := 'CB=TTETABLISSEMENT';
    GRem.ColFormats[1] := 'CB=PGMSAACTIVITE';
    GRem.ColFormats[2] := '';
    GRem.ColFormats[3] := '# ##0.00';
    GRem.ColFormats[4] := '# ##0.00';
    GRem.ColFormats[5] := '# ##0.00';
    GRem.ColFormats[6] := '# ##0.00';
    GRem.ColFormats[7] := '# ##0.00';
    GRem.ColFormats[8] := '# ##0.00';
    GRem.ColFormats[9] := '# ##0.00';
    GRem.ColFormats[10] := '# ##0.00';
    GRem.ColFormats[11] := '# ##0.00';
    GRem.ColFormats[12] := '# ##0.00';
    GRem.ColFormats[13] := '# ##0.00'; //PT18
    GRem.ColFormats[14] := '# ##0.00'; //PT18
  end;
  ChargeZones();

end;

procedure TOF_PGMSARECAPENVOI.ChargeZones();
var
  Siren, Siret: string;
begin
  SetControlText('PERIODICITE', 'T');
  SetControlText('ERAISONSOC', GetParamSoc('SO_LIBELLE'));
  Siret := GetParamSoc('SO_SIRET');
  ForceNumerique(Siret, Siren);
  SetControlText('ESIREN', Copy(Siren, 1, 9));
  RemplirGrilleRem;
end;

Procedure TOF_PGMSARECAPENVOI.EditionFichierMsa;
var
DateDeb, DateFin, Dossier, Entreprise, FileN, NIC, S, Salarie : String;
Segment, Structure, ValeurSeg : String;
Q : TQuery;
T,TA : Tob;
    i,x,r : Integer;
    TypeEvolution,Test,eTAB : String;
    FLect : TextFile;
begin
If TobEtat <> Nil then FreeAndNil(TobEtat);
TobEtat := Tob.Create('FilleEtat',Nil,-1);
//Récupération du fichier
FileN:= GetControlText ('FICHIER');

if (FileN='') then
   begin
   PGIBox ('Aucun fichier sélectionné', Ecran.Caption);
   exit;
   end;
AssignFile (FLect, FileN);
Reset (FLect);
Readln (FLect,S);
TobAssiettes  := Tob.Create('LesAssiettes',Nil,-1);
//PT14 Pb : on ne traitait pas la dernière ligne
//PT14 while (not(eof(FLect))) do
if not(Eof(FLect)) then //PT14
begin
  while (S<>'') do  //PT14
  begin
        Segment:= Copy (S,34,4);
        If Segment = 'PE41' then
        begin
             Etab := Copy (S,14,14);
             T := TobAssiettes.FindFirst(['ENTREPRISE','ETABLISSEMENT','MSAACTIVITE'],[Trim(Copy (S,1,13)),Trim(Copy (S,14,14)),Trim(Copy (S,28,4))],False);
             If T <> nil then
             begin
                 T.PutValue('ELTENTREPRISE8',RendLibAssiette(Copy (S,64,2)));
                 T.PutValue('ELTENTREPRISE9',RendLibAssiette(Copy (S,76,2)));
                 T.PutValue('ELTENTREPRISE10',RendLibAssiette(Copy (S,88,2)));
                 T.PutValue('ELTENTREPRISE11',RendLibAssiette(Copy (S,100,2))); //PT18
                 T.PutValue('ELTENTREPRISE12',RendLibAssiette(Copy (S,112,2))); //PT18
                 T.PutValue('VALELTENTREPRISE8',TraiteValeurNum(Copy (S,66,10)));
                 T.PutValue('VALELTENTREPRISE9',TraiteValeurNum(Copy (S,78,10)));
                 T.PutValue('VALELTENTREPRISE10',TraiteValeurNum(Copy (S,90,10)));
                 T.PutValue('VALELTENTREPRISE11',TraiteValeurNum(Copy (S,102,10))); //PT18
                 T.PutValue('VALELTENTREPRISE12',TraiteValeurNum(Copy (S,114,10))); //PT18
             end
             else
             begin
                 T := Tob.Create('FilleeAssiette',TobAssiettes,-1);
                 T.AddchampSupValeur('ENTREPRISE',trim(Copy (S,1,13)));
                 T.AddchampSupValeur('ETABLISSEMENT',trim(Copy (S,14,14)));
                 T.AddchampSupValeur('MSAACTIVITE',Trim(Copy (S,28,4)));
                 T.AddchampSupValeur('ELTENTREPRISE1',RendLibAssiette(Copy (S,64,2)));
                 T.AddchampSupValeur('ELTENTREPRISE2',RendLibAssiette(Copy (S,76,2)));
                 T.AddchampSupValeur('ELTENTREPRISE3',RendLibAssiette(Copy (S,88,2)));
                 T.AddchampSupValeur('ELTENTREPRISE4',RendLibAssiette(Copy (S,100,2)));
                 T.AddchampSupValeur('ELTENTREPRISE5',RendLibAssiette(Copy (S,112,2)));
                 T.AddchampSupValeur('ELTENTREPRISE6',RendLibAssiette(Copy (S,124,2)));
                 T.AddchampSupValeur('ELTENTREPRISE7',RendLibAssiette(Copy (S,136,2)));
                 T.AddchampSupValeur('ELTENTREPRISE8','');
                 T.AddchampSupValeur('ELTENTREPRISE9','');
                 T.AddchampSupValeur('ELTENTREPRISE10','');
                 T.AddchampSupValeur('ELTENTREPRISE11',''); //PT18
                 T.AddchampSupValeur('ELTENTREPRISE12',''); //PT18
                 T.AddchampSupValeur('VALELTENTREPRISE1',TraiteValeurNum(Copy (S,66,10)));
                 T.AddchampSupValeur('VALELTENTREPRISE2',TraiteValeurNum(Copy (S,78,10)));
                 T.AddchampSupValeur('VALELTENTREPRISE3',TraiteValeurNum(Copy (S,90,10)));
                 T.AddchampSupValeur('VALELTENTREPRISE4',TraiteValeurNum(Copy (S,102,10)));
                 T.AddchampSupValeur('VALELTENTREPRISE5',TraiteValeurNum(Copy (S,114,10)));
                 T.AddchampSupValeur('VALELTENTREPRISE6',TraiteValeurNum(Copy (S,126,10)));
                 T.AddchampSupValeur('VALELTENTREPRISE7',TraiteValeurNum(Copy (S,138,10)));
                 T.AddchampSupValeur('VALELTENTREPRISE8','');
                 T.AddchampSupValeur('VALELTENTREPRISE9','');
                 T.AddchampSupValeur('VALELTENTREPRISE10','');
                 T.AddchampSupValeur('VALELTENTREPRISE11','');  //PT18
                 T.AddchampSupValeur('VALELTENTREPRISE12','');  //PT18
             end;
        end;
        if not(Eof(FLect)) then   //PT14
          Readln (FLect,S)
        else
          S:='';                  //PT14
  end;
end;
Reset (FLect);
Readln (FLect,S);
//PT14 while (not(eof(FLect))) do
if not(Eof(FLect)) then //PT14
begin
  while (S<>'') do  //PT14
      begin
            Entreprise:= Copy (S,1,13);
            Segment:= Copy (S,34,4);
            If Segment = 'PE11' then Entreprise := copy(S,1,13);
            If Segment = 'PE31' then
            begin
                 For x := 1 to 5 do
                 begin
                    Test :=  Copy (S,133 + ((x-1)*8),2);
                    If Test >= '01'then
                    begin
                         T := Tob.Create('FilleEtat',TobEtat,-1);
                         T.AddchampSupValeur('ORDRE',0);
                         T.AddchampSupValeur('ENTREPRISE',Copy (S,1,13));
                         T.AddchampSupValeur('ETABLISSEMENT',Copy (S,14,14));
                         T.AddchampSupValeur('MSAACTIVITE',Copy (S,28,4));
                         T.AddChampSupValeur('TYPEINFO','PER');
                         T.AddChampSupValeur('SALARIE',Copy (S,38,13));
                         T.AddChampSupValeur('DATEDEBUT',Copy (S,104,8));
                         T.AddChampSupValeur('NOM',Copy (S,51,25));
                         T.AddChampSupValeur('PRENOM',Copy (S,76,20));
                         T.AddChampSupValeur('NUMSS',Copy (S,38,13));
                         T.AddChampSupValeur('DATENAISS',Copy (S,102,2)+'/'+Copy (S,100,2)+'/'+Copy (S,96,4));
                         T.AddChampSupValeur('ACTIVITE',RechDom('PGMSAACTIVITE',Copy (S,28,4),False));
                         T.AddChampSupValeur('UG',Copy (S,32,2)); //PT13
                         T.AddChampSupValeur('EXO','');
                         T.AddChampSupValeur('COLONNE1',Copy (S,110,2)+'/'+Copy (S,108,2)+'/'+Copy (S,104,4)+ ' - '+Copy (S,118,2)+'/'+Copy (S,116,2)+'/'+Copy (S,112,4));
                         If (Copy (S,125,3) <> '   ') and (Copy (S,125,3) <> '000') then T.AddChampSupValeur('COLONNE2',TraiteValeurHeures(Copy (S,128,5))+' / '+TraiteValeurNumDec(Copy (S,125,3)))
                         else T.AddChampSupValeur('COLONNE2',TraiteValeurHeures(Copy (S,128,5)));
                         T.AddChampSupValeur('COLONNE3',Copy (S,133 + ((x-1)*8),2)+' '+RechDom('PGMSATYPEELEMENTCALC',Copy (S,133 + ((x-1)*8),2),False));
                         T.AddChampSupValeur('COLONNE4',TraiteValeurNum(Copy (S,135 + ((x-1)*8),6)));
                         TA := TobAssiettes.FindFirst(['ENTREPRISE','ETABLISSEMENT','MSAACTIVITE'],[Trim(Copy (S,1,13)),Trim(Copy (S,14,14)),Trim(Copy (S,28,4))],False);
                         If TA <> Nil then
                         begin
                             T.AddchampSupValeur('ELTENTREPRISE1',TA.GetValue('ELTENTREPRISE1'));
                             T.AddchampSupValeur('ELTENTREPRISE2',TA.GetValue('ELTENTREPRISE2'));
                             T.AddchampSupValeur('ELTENTREPRISE3',TA.GetValue('ELTENTREPRISE3'));
                             T.AddchampSupValeur('ELTENTREPRISE4',TA.GetValue('ELTENTREPRISE4'));
                             T.AddchampSupValeur('ELTENTREPRISE5',TA.GetValue('ELTENTREPRISE5'));
                             T.AddchampSupValeur('ELTENTREPRISE6',TA.GetValue('ELTENTREPRISE6'));
                             T.AddchampSupValeur('ELTENTREPRISE7',TA.GetValue('ELTENTREPRISE7'));
                             T.AddchampSupValeur('ELTENTREPRISE8',TA.GetValue('ELTENTREPRISE8'));
                             T.AddchampSupValeur('ELTENTREPRISE9',TA.GetValue('ELTENTREPRISE9'));
                             T.AddchampSupValeur('ELTENTREPRISE10',TA.GetValue('ELTENTREPRISE10'));
                             T.AddchampSupValeur('ELTENTREPRISE11',TA.GetValue('ELTENTREPRISE11')); //PT18
                             T.AddchampSupValeur('ELTENTREPRISE12',TA.GetValue('ELTENTREPRISE12')); //PT18
                             T.AddchampSupValeur('VALELTENTREPRISE1',TA.GetValue('VALELTENTREPRISE1'));
                             T.AddchampSupValeur('VALELTENTREPRISE2',TA.GetValue('VALELTENTREPRISE2'));
                             T.AddchampSupValeur('VALELTENTREPRISE3',TA.GetValue('VALELTENTREPRISE3'));
                             T.AddchampSupValeur('VALELTENTREPRISE4',TA.GetValue('VALELTENTREPRISE4'));
                             T.AddchampSupValeur('VALELTENTREPRISE5',TA.GetValue('VALELTENTREPRISE5'));
                             T.AddchampSupValeur('VALELTENTREPRISE6',TA.GetValue('VALELTENTREPRISE6'));
                             T.AddchampSupValeur('VALELTENTREPRISE7',TA.GetValue('VALELTENTREPRISE7'));
                             T.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('VALELTENTREPRISE8'));
                             T.AddchampSupValeur('VALELTENTREPRISE9',TA.GetValue('VALELTENTREPRISE9'));
                             T.AddchampSupValeur('VALELTENTREPRISE10',TA.GetValue('VALELTENTREPRISE10'));
                             T.AddchampSupValeur('VALELTENTREPRISE11',TA.GetValue('VALELTENTREPRISE11'));  //PT18
                             T.AddchampSupValeur('VALELTENTREPRISE12',TA.GetValue('VALELTENTREPRISE12'));  //PT18
                         end
                         else
                         begin
                             T.AddchampSupValeur('ELTENTREPRISE1','');
                             T.AddchampSupValeur('ELTENTREPRISE2','');
                             T.AddchampSupValeur('ELTENTREPRISE3','');
                             T.AddchampSupValeur('ELTENTREPRISE4','');
                             T.AddchampSupValeur('ELTENTREPRISE5','');
                             T.AddchampSupValeur('ELTENTREPRISE6','');
                             T.AddchampSupValeur('ELTENTREPRISE7','');
                             T.AddchampSupValeur('ELTENTREPRISE8','');
                             T.AddchampSupValeur('ELTENTREPRISE9','');
                             T.AddchampSupValeur('ELTENTREPRISE10','');
                             T.AddchampSupValeur('ELTENTREPRISE11',''); //PT18
                             T.AddchampSupValeur('ELTENTREPRISE12',''); //PT18
                             T.AddchampSupValeur('VALELTENTREPRISE1','');
                             T.AddchampSupValeur('VALELTENTREPRISE2','');
                             T.AddchampSupValeur('VALELTENTREPRISE3','');
                             T.AddchampSupValeur('VALELTENTREPRISE4','');
                             T.AddchampSupValeur('VALELTENTREPRISE5','');
                             T.AddchampSupValeur('VALELTENTREPRISE6','');
                             T.AddchampSupValeur('VALELTENTREPRISE7','');
                             T.AddchampSupValeur('VALELTENTREPRISE8','');
                             T.AddchampSupValeur('VALELTENTREPRISE9','');
                             T.AddchampSupValeur('VALELTENTREPRISE10','');
                             T.AddchampSupValeur('VALELTENTREPRISE11',''); //PT18
                             T.AddchampSupValeur('VALELTENTREPRISE12',''); //PT18
                         end;
                    end;
                 end;
            end
            else
            //DEB PT13
            If Segment = 'PE32' then
            begin
                 if copy(S,120,2) <> '  ' then
                    AjoutLigneHSupFichier(S,100);
                 if copy(S,136,2) <> '  ' then
                    AjoutLigneHSupFichier(S,101);
                 if copy(S,152,2) <> '  ' then
                    AjoutLigneHSupFichier(S,102);
                 if copy(S,168,2) <> '  ' then
                    AjoutLigneHSupFichier(S,103);
            end
            //FIN PT13
            else
            if (segment >='PE21') or (segment <='PE24') then
            begin
                If Segment = 'PE21' then
                begin
                     If IsValidDate(Copy (S,110,2)+'/'+Copy (S,108,2)+'/'+Copy (S,104,4)) then
                     AjoutLigneEvolutionFichier (s,'Date début d''activité',Copy (S,110,2)+'/'+Copy (S,108,2)+'/'+Copy (S,104,4));
                     If IsValidDate(Copy (S,118,2)+'/'+Copy (S,116,2)+'/'+Copy (S,112,4)) then
                     AjoutLigneEvolutionFichier (s,'Date fin d''activité',Copy (S,118,2)+'/'+Copy (S,116,2)+'/'+Copy (S,112,4));
                end
                else If Segment = 'PE22' then
                begin
                     If IsValidDate(Copy (S,110,2)+'/'+Copy (S,108,2)+'/'+Copy (S,104,4)) then
                     AjoutLigneEvolutionFichier (s,'Date début suspension de contrat',Copy (S,110,2)+'/'+Copy (S,108,2)+'/'+Copy (S,104,4));
                     If IsValidDate(Copy (S,118,2)+'/'+Copy (S,116,2)+'/'+Copy (S,112,4)) then
                     AjoutLigneEvolutionFichier (s,'Date fin suspension de contrat',Copy (S,118,2)+'/'+Copy (S,116,2)+'/'+Copy (S,112,4));
                     If Copy (S,120,2) >= '01' then
                     AjoutLigneEvolutionFichier (s,'Motif de suspension de contrat',Copy (S,120,2) + ' '+RechDom('PGMSASUSPCONTRAT',Copy (S,120,2),False));
                end
                else If Segment = 'PE23' then
                begin
                     If Trim(Copy (S,115,5)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Code postal lieu de travail',Copy (S,115,5));
                     If Trim(Copy (S,120,4)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Pourcentage temps partiel',TraiteValeurHeures(Copy (S,120,4)));
                     If Trim(Copy (S,124,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Code temps partiel option temps plein',Copy (S,124,1));
                     If Trim(Copy (S,125,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Contrat de travail',RechDom('PGTYPECONTRATMSA',Copy (S,125,1),False));
                     If Trim(Copy (S,126,5)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Nombre d''heures convention', TraiteValeurHeures(Copy (S,126,5)));
                     If Trim(Copy (S,131,5)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Niveau ou coefficient hiérarchique',Copy (S,131,5));
                     //DEB PT13
                     If Trim(Copy (S,136,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Périodicité de la durée de travail',Copy (S,136,1));
                     If Trim(Copy (S,137,6)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Durée du contrat de travail',TraiteValeurNumDec(Copy (S,137,6)));
                     //FIN PT13
                end
                else If Segment = 'PE24' then
                begin
                     If Trim(Copy (S,112,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Personnel technique',Copy (S,112,1));
                     If Trim(Copy (S,124,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Personnel technique CUMA',Copy (S,124,1));
                     If Trim(Copy (S,114,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Top cadre (APECITA)',Copy (S,114,1));
                     If Trim(Copy (S,116,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Poly-employeur',Copy (S,116,1));
                     If Trim(Copy (S,118,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Domicilié fiscalement hors France',Copy (S,118,1));
                     If Trim(Copy (S,119,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Code classe élèves',Copy (S,119,1) + ' '+RechDom('PGMSACODECLASSE',Copy (S,119,1),False));
                     If Trim(Copy (S,120,2)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Unité de gestion',Copy (S,120,2));
                     If Trim(Copy (S,122,3)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Code qualité',Copy (S,122,3));
                     If Trim(Copy (S,129,25)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Emploi occupé',Copy (S,129,25));
                     If Trim(Copy (S,127,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Bénéficie loi Aubry 1',Copy (S,127,1));
                     If Trim(Copy (S,128,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Bénéficie loi Aubry 2',Copy (S,128,1));
                     If Trim(Copy (S,154,1)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Retraité en activité',Copy (S,154,1));
                     If Trim(Copy (S,155,25)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Convention collective de travail',Copy (S,155,25));
                     If Trim(Copy (S,125,2)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Code CRCCA',Copy (S,125,2)+ ' '+RechDom('PGSCATEGORIELMSA',Copy (S,125,2),False));
                     If Trim(Copy (S,180,2)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Section prud''homale',Copy (S,180,2)+ ' '+RechDom('PGSECTIONPRUDMSA',Copy (S,181,1),False));
                     If Trim(Copy (S,182,2)) <> '' then
                     AjoutLigneEvolutionFichier (s,'Collège prud''homal',Copy (S,182,2)+ ' '+RechDom('PGCOLLEGEPRUDMSA',Copy (S,183,1),False));
                end;
            end;
            if not(Eof(FLect)) then   //PT14
              Readln (FLect,S)
            else
              S:='';                  //PT14
      end;
end;
 CloseFile(FLect);
 FreeAndNil(TobAssiettes);
end;

Function TOF_PGMSARECAPENVOI.TraiteValeurNum (Valeur : String): String;
var L,i,Nb0 : Integer;
begin
     L := Length(Valeur);
     Nb0 := 0;
     For i := 1 to L do
     begin
          If Valeur[i] = '0' then Nb0 := Nb0 + 1
          else Break;
     end;
     Nb0 := Nb0 + 1;
     Result := Copy(Valeur,Nb0,L);
end;

Function TOF_PGMSARECAPENVOI.TraiteValeurNumDec (Valeur : String): String;
var L,i,Nb0 : Integer;
    Entier,Decimales : String;
begin
     L := Length(Valeur);
     Nb0 := 0;
     For i := 1 to L do
     begin
          If Valeur[i] = '0' then Nb0 := Nb0 + 1
          else Break;
     end;
     Nb0 := Nb0 + 1;
     Valeur := Copy(Valeur,Nb0,L);

     L := Length(Valeur);
     Decimales := '';
     Entier := '';
     For i := L downto 1 do
     begin
          If Length(Decimales)<2 then Decimales := Valeur[i] + Decimales
          else Entier := Valeur[i] + Entier;
     end;
     Result := entier + ','+Decimales;
end;

Function TOF_PGMSARECAPENVOI.TraiteValeurHeures (Valeur : String): String;
var L,i : Integer;
    Entier,Decimales : String;
begin
     L := Length(Valeur);
     Decimales := '';
     Entier := '';
     For i := L downto 1 do
     begin
          If Length(Decimales)<2 then Decimales := Valeur[i] + Decimales
          else Entier := Valeur[i] + Entier;
     end;
     Result := entier + ','+Decimales;
end;


Procedure TOF_PGMSARECAPENVOI.AjoutLigneEvolutionFichier (ValeurSeg : String;Libelle,Valeur : String);
var TE,TA : Tob;
    LeType : String;
begin
        LeType := Copy (ValeurSeg,34,4);
        If Letype = 'PE21' then LeType := 'Dates d''activité'
        else If Letype = 'PE22' then LeType := 'Suspension du contrat de travail'
        else If Letype = 'PE23' then LeType := 'Eléments du contrat de travail'
        else If Letype = 'PE24' then LeType := 'Eléments complémentaires';
        TE := Tob.Create('FilleEtat',TobEtat,-1);
        TE.AddchampSupValeur('ORDRE',2);
        TE.AddchampSupValeur('ENTREPRISE',Copy (ValeurSeg,1,13));
        TE.AddchampSupValeur('ETABLISSEMENT',Copy (ValeurSeg,14,14));
        TE.AddchampSupValeur('MSAACTIVITE',Copy (ValeurSeg,28,4));
        TE.AddchampSupValeur('TYPEINFO','ZZZ');
        TE.AddChampSupValeur('SALARIE',Copy (ValeurSeg,38,13));
        TE.AddChampSupValeur('DATEDEBUT',Copy (ValeurSeg,110,2)+'/'+Copy (ValeurSeg,108,2)+'/'+Copy (ValeurSeg,104,4));
        TE.AddChampSupValeur('NOM',Copy (ValeurSeg,51,25));
        TE.AddChampSupValeur('PRENOM',Copy (ValeurSeg,76,20));
        TE.AddChampSupValeur('NUMSS',Copy (ValeurSeg,38,13));
        TE.AddChampSupValeur('DATENAISS',Copy (ValeurSeg,102,2)+'/'+Copy (ValeurSeg,100,2)+'/'+Copy (ValeurSeg,96,4));
        TE.AddChampSupValeur('ACTIVITE',RechDom('PGMASACTIVITE',Copy (ValeurSeg,28,4),False));
        TE.AddChampSupValeur('UG',Copy (ValeurSeg,32,2)); //PT13
        TE.AddChampSupValeur('EXO','');
        TE.AddChampSupValeur('COLONNE1',LeType);
        TE.AddChampSupValeur('COLONNE2',Copy (ValeurSeg,110,2)+'/'+Copy (ValeurSeg,108,2)+'/'+Copy (ValeurSeg,104,4));
        TE.AddchampSupValeur('COLONNE3',Libelle,False);
        TE.AddchampSupValeur('COLONNE4',Valeur,False);
        TA := TobAssiettes.FindFirst(['ENTREPRISE','ETABLISSEMENT','MSAACTIVITE'],[Trim(Copy (ValeurSeg,1,13)),Trim(Copy (ValeurSeg,14,14)),Trim(Copy (ValeurSeg,28,4))],False);
       If TA <> Nil then
       begin
           TE.AddchampSupValeur('ELTENTREPRISE1',TA.GetValue('ELTENTREPRISE1'));
           TE.AddchampSupValeur('ELTENTREPRISE2',TA.GetValue('ELTENTREPRISE2'));
           TE.AddchampSupValeur('ELTENTREPRISE3',TA.GetValue('ELTENTREPRISE3'));
           TE.AddchampSupValeur('ELTENTREPRISE4',TA.GetValue('ELTENTREPRISE4'));
           TE.AddchampSupValeur('ELTENTREPRISE5',TA.GetValue('ELTENTREPRISE5'));
           TE.AddchampSupValeur('ELTENTREPRISE6',TA.GetValue('ELTENTREPRISE6'));
           TE.AddchampSupValeur('ELTENTREPRISE7',TA.GetValue('ELTENTREPRISE7'));
           TE.AddchampSupValeur('ELTENTREPRISE8',TA.GetValue('ELTENTREPRISE8'));
           TE.AddchampSupValeur('ELTENTREPRISE9',TA.GetValue('ELTENTREPRISE9'));
           TE.AddchampSupValeur('ELTENTREPRISE10',TA.GetValue('ELTENTREPRISE10'));
           TE.AddchampSupValeur('ELTENTREPRISE11',TA.GetValue('ELTENTREPRISE11')); //PT18
           TE.AddchampSupValeur('ELTENTREPRISE12',TA.GetValue('ELTENTREPRISE12')); //PT18
           TE.AddchampSupValeur('VALELTENTREPRISE1',TA.GetValue('VALELTENTREPRISE1'));
           TE.AddchampSupValeur('VALELTENTREPRISE2',TA.GetValue('VALELTENTREPRISE2'));
           TE.AddchampSupValeur('VALELTENTREPRISE3',TA.GetValue('VALELTENTREPRISE3'));
           TE.AddchampSupValeur('VALELTENTREPRISE4',TA.GetValue('VALELTENTREPRISE4'));
           TE.AddchampSupValeur('VALELTENTREPRISE5',TA.GetValue('VALELTENTREPRISE5'));
           TE.AddchampSupValeur('VALELTENTREPRISE6',TA.GetValue('VALELTENTREPRISE6'));
           TE.AddchampSupValeur('VALELTENTREPRISE7',TA.GetValue('VALELTENTREPRISE7'));
           TE.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('VALELTENTREPRISE8'));
           TE.AddchampSupValeur('VALELTENTREPRISE9',TA.GetValue('VALELTENTREPRISE9'));
           TE.AddchampSupValeur('VALELTENTREPRISE10',TA.GetValue('VALELTENTREPRISE10'));
           TE.AddchampSupValeur('VALELTENTREPRISE11',TA.GetValue('VALELTENTREPRISE11'));  //PT18
           TE.AddchampSupValeur('VALELTENTREPRISE12',TA.GetValue('VALELTENTREPRISE12'));  //PT18
       end
       else
       begin
           TE.AddchampSupValeur('ELTENTREPRISE1','');
           TE.AddchampSupValeur('ELTENTREPRISE2','');
           TE.AddchampSupValeur('ELTENTREPRISE3','');
           TE.AddchampSupValeur('ELTENTREPRISE4','');
           TE.AddchampSupValeur('ELTENTREPRISE5','');
           TE.AddchampSupValeur('ELTENTREPRISE6','');
           TE.AddchampSupValeur('ELTENTREPRISE7','');
           TE.AddchampSupValeur('ELTENTREPRISE8','');
           TE.AddchampSupValeur('ELTENTREPRISE9','');
           TE.AddchampSupValeur('ELTENTREPRISE10','');
           TE.AddchampSupValeur('ELTENTREPRISE11','');  //PT18
           TE.AddchampSupValeur('ELTENTREPRISE12','');  //PT18
           TE.AddchampSupValeur('VALELTENTREPRISE1','');
           TE.AddchampSupValeur('VALELTENTREPRISE2','');
           TE.AddchampSupValeur('VALELTENTREPRISE3','');
           TE.AddchampSupValeur('VALELTENTREPRISE4','');
           TE.AddchampSupValeur('VALELTENTREPRISE5','');
           TE.AddchampSupValeur('VALELTENTREPRISE6','');
           TE.AddchampSupValeur('VALELTENTREPRISE7','');
           TE.AddchampSupValeur('VALELTENTREPRISE8','');
           TE.AddchampSupValeur('VALELTENTREPRISE9','');
           TE.AddchampSupValeur('VALELTENTREPRISE10','');
           TE.AddchampSupValeur('VALELTENTREPRISE11',''); //PT18
           TE.AddchampSupValeur('VALELTENTREPRISE12',''); //PT18
       end;
end;

Procedure TOF_PGMSARECAPENVOI.AjoutLigneHSupFichier (S : String;Ordre : integer);
var T,TA : Tob;
begin
       T := Tob.Create('FilleEtat',TobEtat,-1);
       T.AddchampSupValeur('ORDRE',Ordre);
       T.AddchampSupValeur('ENTREPRISE',Copy (S,1,13));
       T.AddchampSupValeur('ETABLISSEMENT',Copy (S,14,14));
       T.AddchampSupValeur('MSAACTIVITE',Copy (S,28,4));
       T.AddChampSupValeur('TYPEINFO','PER');
       T.AddChampSupValeur('SALARIE',Copy (S,38,13));
       T.AddChampSupValeur('DATEDEBUT',Copy (S,104,8));
       T.AddChampSupValeur('NOM',Copy (S,51,25));
       T.AddChampSupValeur('PRENOM',Copy (S,76,20));
       T.AddChampSupValeur('NUMSS',Copy (S,38,13));
       T.AddChampSupValeur('DATENAISS',Copy (S,102,2)+'/'+Copy (S,100,2)+'/'+Copy (S,96,4));
       T.AddChampSupValeur('ACTIVITE',RechDom('PGMSAACTIVITE',Copy (S,28,4),False));
       T.AddChampSupValeur('UG',Copy (S,32,2));    //PT13
       T.AddChampSupValeur('EXO','');
       if Ordre = 100 then
       begin
         T.AddChampSupValeur('COLONNE1','Heures supplémentaires');
         T.AddChampSupValeur('COLONNE2',TraiteValeurHeures(copy(S,125,5)));
         T.AddChampSupValeur('COLONNE3','Type ' + copy(S,120,2)+' - Taux de majoration 1 à ' + Copy (S,122,3) + ' %');
         T.AddChampSupValeur('COLONNE4',Copy (S,130,6));
       end;
       if Ordre = 101 then
       begin
         T.AddChampSupValeur('COLONNE1','');
         T.AddChampSupValeur('COLONNE2',TraiteValeurHeures(copy(S,141,5)));
         T.AddChampSupValeur('COLONNE3','Type ' + copy(S,136,2)+' - Taux de majoration 2 à ' + Copy (S,138,3) + ' %');
         T.AddChampSupValeur('COLONNE4',Copy (S,146,6));
       end;
       if Ordre = 102 then
       begin
         T.AddChampSupValeur('COLONNE1','');
         T.AddChampSupValeur('COLONNE2',TraiteValeurHeures(copy(S,157,5)));
         T.AddChampSupValeur('COLONNE3','Type ' + copy(S,152,2)+' - Taux de majoration 3 à ' + Copy (S,154,3) + ' %');
         T.AddChampSupValeur('COLONNE4',Copy (S,162,6));
       end;
       if Ordre = 103 then
       begin
         T.AddChampSupValeur('COLONNE1','');
         T.AddChampSupValeur('COLONNE2',TraiteValeurHeures(copy(S,173,5)));
         T.AddChampSupValeur('COLONNE3','Type ' + copy(S,168,2)+' - Taux de majoration 4 à ' + Copy (S,170,3) + ' %');
         T.AddChampSupValeur('COLONNE4',Copy (S,178,6));
       end;
       TA := TobAssiettes.FindFirst(['ENTREPRISE','ETABLISSEMENT','MSAACTIVITE'],[Trim(Copy (S,1,13)),Trim(Copy (S,14,14)),Trim(Copy (S,28,4))],False);
       If TA <> Nil then
       begin
           T.AddchampSupValeur('ELTENTREPRISE1',TA.GetValue('ELTENTREPRISE1'));
           T.AddchampSupValeur('ELTENTREPRISE2',TA.GetValue('ELTENTREPRISE2'));
           T.AddchampSupValeur('ELTENTREPRISE3',TA.GetValue('ELTENTREPRISE3'));
           T.AddchampSupValeur('ELTENTREPRISE4',TA.GetValue('ELTENTREPRISE4'));
           T.AddchampSupValeur('ELTENTREPRISE5',TA.GetValue('ELTENTREPRISE5'));
           T.AddchampSupValeur('ELTENTREPRISE6',TA.GetValue('ELTENTREPRISE6'));
           T.AddchampSupValeur('ELTENTREPRISE7',TA.GetValue('ELTENTREPRISE7'));
           T.AddchampSupValeur('ELTENTREPRISE8',TA.GetValue('ELTENTREPRISE8'));
           T.AddchampSupValeur('ELTENTREPRISE9',TA.GetValue('ELTENTREPRISE9'));
           T.AddchampSupValeur('ELTENTREPRISE10',TA.GetValue('ELTENTREPRISE10'));
           T.AddchampSupValeur('ELTENTREPRISE11',TA.GetValue('ELTENTREPRISE11'));  //PT18
           T.AddchampSupValeur('ELTENTREPRISE12',TA.GetValue('ELTENTREPRISE12'));  //PT18
           T.AddchampSupValeur('VALELTENTREPRISE1',TA.GetValue('VALELTENTREPRISE1'));
           T.AddchampSupValeur('VALELTENTREPRISE2',TA.GetValue('VALELTENTREPRISE2'));
           T.AddchampSupValeur('VALELTENTREPRISE3',TA.GetValue('VALELTENTREPRISE3'));
           T.AddchampSupValeur('VALELTENTREPRISE4',TA.GetValue('VALELTENTREPRISE4'));
           T.AddchampSupValeur('VALELTENTREPRISE5',TA.GetValue('VALELTENTREPRISE5'));
           T.AddchampSupValeur('VALELTENTREPRISE6',TA.GetValue('VALELTENTREPRISE6'));
           T.AddchampSupValeur('VALELTENTREPRISE7',TA.GetValue('VALELTENTREPRISE7'));
           T.AddchampSupValeur('VALELTENTREPRISE8',TA.GetValue('VALELTENTREPRISE8'));
           T.AddchampSupValeur('VALELTENTREPRISE9',TA.GetValue('VALELTENTREPRISE9'));
           T.AddchampSupValeur('VALELTENTREPRISE10',TA.GetValue('VALELTENTREPRISE10'));
           T.AddchampSupValeur('VALELTENTREPRISE11',TA.GetValue('VALELTENTREPRISE11')); //PT18
           T.AddchampSupValeur('VALELTENTREPRISE12',TA.GetValue('VALELTENTREPRISE12')); //PT18
       end
       else
       begin
           T.AddchampSupValeur('ELTENTREPRISE1','');
           T.AddchampSupValeur('ELTENTREPRISE2','');
           T.AddchampSupValeur('ELTENTREPRISE3','');
           T.AddchampSupValeur('ELTENTREPRISE4','');
           T.AddchampSupValeur('ELTENTREPRISE5','');
           T.AddchampSupValeur('ELTENTREPRISE6','');
           T.AddchampSupValeur('ELTENTREPRISE7','');
           T.AddchampSupValeur('ELTENTREPRISE8','');
           T.AddchampSupValeur('ELTENTREPRISE9','');
           T.AddchampSupValeur('ELTENTREPRISE10','');
           T.AddchampSupValeur('ELTENTREPRISE11','');  //PT18
           T.AddchampSupValeur('ELTENTREPRISE12','');  //PT18
           T.AddchampSupValeur('VALELTENTREPRISE1','');
           T.AddchampSupValeur('VALELTENTREPRISE2','');
           T.AddchampSupValeur('VALELTENTREPRISE3','');
           T.AddchampSupValeur('VALELTENTREPRISE4','');
           T.AddchampSupValeur('VALELTENTREPRISE5','');
           T.AddchampSupValeur('VALELTENTREPRISE6','');
           T.AddchampSupValeur('VALELTENTREPRISE7','');
           T.AddchampSupValeur('VALELTENTREPRISE8','');
           T.AddchampSupValeur('VALELTENTREPRISE9','');
           T.AddchampSupValeur('VALELTENTREPRISE10','');
           T.AddchampSupValeur('VALELTENTREPRISE11',''); //PT18
           T.AddchampSupValeur('VALELTENTREPRISE12',''); //PT18
       end;
end;

procedure TOF_PGMSARECAPENVOI.DateElipsisClick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMSARECAPENVOI.VerifDate(Sender:TObject);
var aa,jj,mm:Word;
    Jour,Mois:String;
    Datedebut,DateFin:TDateTime;
begin
IF Sender=Nil Then Exit;
If THEdit(Sender).Name='DATEDEBUT' Then
   begin
   If Not IsValidDate(THEdit(Sender).Text) Then
      begin
      PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas valide','Edition déclaration des salaires');
      SetFocusControl('DATEDEBUT');
      Exit;
      end;
   DateDebut:=StrToDate(GetControlText('DATEDEBUT'));
   DecodeDate(DateDebut,aa,mm,jj);
   Mois:=IntToStr(mm);
   Jour:=IntToStr(jj);
   If (Jour<>'1') or (Mois<>'1') Then
      If (Jour<>'1') or (Mois<>'4') Then
         If (Jour<>'1') or (Mois<>'7') Then
            If (Jour<>'1') or (Mois<>'10') Then
            begin
            PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas une date de début de trimestre','Edition déclaration des salaires');
            SetFocusControl('DATEDEBUT');
            Exit;
            end;
   DateFin:=PlusMois(DateDebut,2);
   DateFin:=FinDeMois(DateFin);
   SetControlText('DATEFIN',DateToStr(DateFin));
   Exit;
   end;
If THEdit(Sender).Name='DATEFIN' Then
   begin
   If Not IsValidDate(THEdit(Sender).Text) Then
      begin
      PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas valide','Edition déclaration des salaires');
      SetFocusControl('DATEFIN');
      Exit;
      end;
   DateFin:=StrToDate(GetControlText('DATEFIN'));
   DecodeDate(DateFin,aa,mm,jj);
   Mois:=IntToStr(mm);
   Jour:=IntToStr(jj);
   If (Jour<>'31') or (Mois<>'3') Then
      If (Jour<>'30') or (Mois<>'6') Then
         If (Jour<>'30') or (Mois<>'9') Then
            If (Jour<>'31') or (Mois<>'12') Then
            begin
            PGIBox('La date saisie '''+THEdit(Sender).text+'''n''est pas une date de fin de trimestre','Edition déclaration des salaires');
            SetFocusControl('DATEFIN');
            Exit;
            end;
   DateDebut:=PlusMois(DateFin,-2);
   DateDebut:=DebutDeMois(DateDebut);
   SetControlText('DATEDEBUT',DateToStr(DateDebut));
   end;
end;


procedure TOF_PGMSARECAPENVOI.RemplirGrilleRem;
var
  Q: TQuery;
  TobEtab, TobGrille, TobActivite, TG,TobUG: Tob;
  WhereEtab,Etab, Activite,UGMsa : string;
  i, a,u: Integer;
  Montant: Double;
  CalculerTcp : Boolean;
  DD,DF : TDateTime;
  HMTrad: THSystemMenu;
begin
  DD := StrTodate(GetControlText('DATEDEBUT'));
  DF := StrTodate(GetControlText('DATEFIN'));   
  CalculerTcp := GetParamSocSecur ('SO_PGMSACALCTCP',False);
  Q := OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS'+WhereEtab, True);
  TobEtab := Tob.Create('LesEtablissements', nil, -1);
  TobEtab.LoadDetailDB('LesEtablissements', '', '', Q, False);
  Ferme(Q);
  TobGrille := Tob.Create('RemplirLaGrille', nil, -1);
  for i := 0 to TobEtab.Detail.Count - 1 do
  begin
    Etab := TobEtab.Detail[i].GetValue('ET_ETABLISSEMENT');
    Q := OpenSQL('SELECT DISTINCT (PE3_MSAACTIVITE) FROM MSAPERIODESPE31 WHERE PE3_ETABLISSEMENT="' + Etab + '"' +
      ' AND PE3_DATEDEBUT>="' + UsDatetime(DD) + '" AND PE3_DATEFIN<="' + UsDatetime(DF) + '"', True);
    TobActivite := Tob.Create('LesActivites', nil, -1);
    TobActivite.LoadDetailDB('LesActivites', '', '', Q, False);
    Ferme(Q);
    for a := 0 to TobActivite.Detail.Count - 1 do
    begin
      Activite := TobActivite.detail[a].GetValue('PE3_MSAACTIVITE');
      Q := OpenSQL('SELECT DISTINCT (PE3_UGMSA) FROM MSAPERIODESPE31 WHERE PE3_MSAACTIVITE="'+Activite+'" AND PE3_ETABLISSEMENT="' + Etab + '"' +
      ' AND PE3_DATEDEBUT>="' + UsDatetime(DD) + '" AND PE3_DATEFIN<="' + UsDatetime(DF) + '"', True);
      TobUG := Tob.Create('LesUG', nil, -1);
      TobUG.LoadDetailDB('LesUG', '', '', Q, False);
      Ferme(Q);
      for u := 0 to TobUG.Detail.Count - 1 do
      begin
           UGMsa := TobUG.detail[u].GetValue('PE3_UGMSA');
            Q := OpenSQL(' SELECT SUM (PHB_BASECOT) AS TOTAL FROM' +
            ' HISTOBULLETIN' +
            ' LEFT JOIN COTISATION ON ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE AND PHB_NATURERUB=PCT_NATURERUB' +
            ' LEFT JOIN MSAPERIODESPE31 ON PHB_SALARIE=PE3_SALARIE AND PHB_DATEDEBUT=PE3_DATEDEBUT AND PHB_DATEFIN=PE3_DATEFIN AND PE3_ORDRE=0' +
            ' WHERE PHB_DATEFIN>="' + UsDateTime(DD) + '" AND PHB_DATEFIN<="' + UsDateTime(DF) + '"' +
            ' AND PCT_BASECSGCRDS="X" AND PE3_MSAACTIVITE="' + Activite + '" AND PE3_ETABLISSEMENT="' + Etab + '" AND PE3_UGMSA="' + UGMsa + '"', True);
            Montant := 0;
            if not Q.eof then Montant := Q.Findfield('TOTAL').AsFloat; // PortageCWAS
            Ferme(Q);
            TG := Tob.Create('UneFille', TobGrille, -1);
            TG.AddChampSupValeur('ETABLISSEMENT', Etab, False);
            TG.AddChampSupValeur('ACTIVITE', Activite, False);
            TG.AddChampSupValeur('UG', UGMsa, False);
            TG.AddChampSupValeur('CSG', Montant, False);
            //DEBUT PT5
            If CalculerTcp then
            begin
                  if (VH_PAIE.PGProfilFnal <> '') and (VH_PAIE.PGProfilFnal <> NULL) then
                  begin
                       Q := OpenSQL('SELECT SUM(PHC_MONTANT) AS MONTANT ' +
                       'FROM HISTOCUMSAL '+
                       ' LEFT JOIN MSAPERIODESPE31 ON PHC_SALARIE=PE3_SALARIE AND PHC_DATEDEBUT=PE3_DATEDEBUT AND PHC_DATEFIN=PE3_DATEFIN AND PE3_ORDRE=0' +
                       'WHERE PHC_CUMULPAIE="35" AND PHC_DATEDEBUT>="' + UsDateTime(DD) + '"  ' +
                       'AND PHC_DATEFIN<="' + UsDateTime(DF) + '" AND PE3_ETABLISSEMENT="' + Etab + '"',true);
                       If Not Q.Eof then TG.AddChampSupValeur('TCP', Q.FindField('MONTANT').AsFloat, False)
                       else TG.AddChampSupValeur('TCP', 0, False);
                       Ferme(Q);
                  end
                  else TG.AddChampSupValeur('TCP', 0, False);
            end
            else TG.AddChampSupValeur('TCP', 0, False);
            //FIN PT5
          end;
          FreeAndNil(TobUG);
      end;
    FreeAndNil(TobActivite);
  end;
  FreeAndNil(TobEtab);
  GRem.Rowcount := 2 + TobGrille.Detail.Count;
  for i := 0 to TobGrille.Detail.Count - 1 do
  begin
    GRem.CellValues[0, i + 2] := TobGrille.Detail[i].GetValue('ETABLISSEMENT');
    GRem.CellValues[1, i + 2] := TobGrille.Detail[i].GetValue('ACTIVITE');
    GRem.CellValues[2, i + 2] := TobGrille.Detail[i].GetValue('UG');
    GRem.CellValues[4, i + 2] := FloatToStr(Arrondi(TobGrille.Detail[i].GetValue('CSG'), 0));
    GRem.CellValues[3, i + 2] := FloatToStr(Arrondi(TobGrille.Detail[i].GetValue('TCP'), 0)); //PT5
  end;
  FreeAndNil(TobGrille);
  HMTrad.ResizeGridColumns(GRem) ;
end;

procedure TOF_PGMSARECAPENVOI.AffichDeclarant(Sender: TObject);
var St : String;
    Q : TQuery;
begin
if GetControlText('DECLARANT')='' then exit;
Q := OpenSQL('SELECT PDA_CIVILITE FROM DECLARANTATTEST WHERE PDA_DECLARANTATTES="'+GetControlText('DECLARANT')+'"',True);
If not Q.Eof then SetControlText('CIVILITE',Q.FindField('PDA_CIVILITE').AsString)
else SetControlText('CIVILITE','');
Ferme(Q);
SetControlText('SIGNATAIRE',RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
SetControlText('LIEUEDITION',RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
if St = 'AUT' then SetControlText('QUALITE', RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False))
else SetControlText('QUALITE', RechDom('PGQUALDECLARANT2', St, False));
end;

procedure TOF_PGMSARECAPENVOI.MajPeriodes(Sender : TObject);
begin
     RemplirGrilleRem;
end;

Function TOF_PGMSARECAPENVOI.RendLibAssiette(Code : String) : String;
begin
     If Code = '51' then Result := '51 tcp'
     else If Code = '53' then Result := '53 CSG'
     else If Code = '54' then Result := '54 CSG/RDS NON IMP.'
     else If Code = '56' then Result := '56 CSG/RDS IMP.'
     else If Code = '63' then Result := '63 CSG/RDS IMP. 6,6 %'
     else If Code = '66' then Result := '66 Contribution pat. Indemnité retraite'  //PT18
     else If Code = '67' then Result := '67 Assiette CSG 7,5% et rds >=11/10/2007'  //PT18
     else If Code = '68' then Result := '68 Assiette CSG 7,5% et rds >=11/10/2007'  //PT18
     else If Code = '58' then Result := '58 Contr. épargne salariale'
     else If Code = '59' then Result := '59 Contr. préretraites tx normal'
{PT18     else If Code = '60' then Result := '60 Contr. préretraites tx réduit'}
     else If Code = '61' then Result := '61 Contr. retraites préstations rentes'
     else If Code = '62' then Result := '62 Contr. retraites préstations primes';
end;

procedure TOF_PGMSARECAPENVOI.ForceNumerique(Buf: string; var Resultat: string);
var
i, j: integer;
begin
j := 1;
Resultat := StringOfChar(' ', Length(Buf));

for i := 1 to Length(Buf) do
    if ((Buf[i] >= '0') and (Buf[i] <= '9')) then
       begin
       Resultat[j] := Buf[i];
       j := j + 1;
       end;

Resultat := Trim(Resultat);
end;
//FIN PT13

Initialization
  registerclasses ( [ TOF_PGMSA_DeclSalaire,TOF_PGMSARECAPENVOI ] ) ;
end.

