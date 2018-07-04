{***********UNITE*************************************************
Auteur  ...... : Paie Pgi  RMA
Créé le ...... : 27/02/2006
Modifié le ... : 08/06/2006
Description .. : Edition des rémunérations des salariés
Mots clefs ... :
*****************************************************************}
{
 PT01  : 12/06/06 RMA V700.7 Correction FQ 13274 ne pas tenir compte des mvt dupliqués dans le comptage
              des congés payés dans l'etat Suivi des absences
 PT02  : 14/06/06 RMA V8 Nouvel état : Analyse de la classification
 PT03  : 08/01/2007 GG V8 Ajout des champs NumeroSS et StatutProf pour l'edition des
                          entrées - sorties
 PT4   : 05/10/2007 FC V_80 FQ 14828
 PT05  : 11/12/2007 RMA V_80 FQ 13957 Nouvel état : Liste des salariés présents
 PT7  : 05/08/2008 NA  V_80 FQ 15490 Modif requête pour le suivi des absences (plantage sous oracle)
}
Unit UTOFEDITREMUNERATION;

Interface

Uses StdCtrls, Classes, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
     eQRS1,
     UtileAgl,
  {$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,
  ParamSoc, HQry,
  UTOB ,UtobDebug ;

Type
  TOF_EDIT_REMUNERATION = Class (TOF)
    private
      Origine :String;
      Norme :Boolean;        //True ==> Nouvelle norme : 2 ruptures possibles (Etab + Champ libre)
      Effectif : Boolean;    //False => Table PAIEENCOURS/Fiche EDIT_REMUNERATION,True => Table SALARIES/Fiche EDIT_EFFECTIF
      PaiePeriode : Boolean; //False => Exercice sociale / True => derniere période paie
      OnSort : Boolean;      //True ==> On a fait la X pour sortir ,on ne test plus la validé de la période saisie
      TobEdit : Tob;

      Procedure Change(Sender: TObject) ;
      Procedure VerifPeriode1(Sender: TObject) ;
      Procedure VerifPeriode2(Sender: TObject) ;
      Procedure ControleChampCompl(PrefTable : String) ;
      Procedure ControleCheckRupture ;
      Procedure InformeRupture(PrefTable : String);
      Procedure Edition_absence(AjoutChamp ,Where ,Suite : String; DatDeb ,DatFin :TDateTime);
      Procedure Edition_AnalBrut(AjoutChamp ,Where ,Suite : String; DatDeb ,DatFin :TDateTime);
      Procedure Edition_Pcharges(Where ,Suite : String; DatDeb ,DatFin :TDateTime);
    public
      Procedure OnArgument (S : String ) ; override ;
      Procedure OnUpdate                 ; override ;
      Procedure OnClose                  ; override ;
   end ;

Implementation
Uses PgEditOutils,PgOutils2,P5Def,EntPaie,PgEdtEtat;

Procedure TOF_EDIT_REMUNERATION.OnArgument (S : String ) ;
var
  Check : TCheckBox;
  DebPer, FinPer, Defaut : THEdit;
  DebExer, FinExer : TDateTime;
  MoisE, AnneeE, ComboExer, DPer, FPer, ExerPerEncours : string;
  I : Integer;
  Label1 ,Label2 :THLabel;

Begin
  Inherited ;
  Origine := Trim(S);
  If Origine = '' Then Exit;
  Norme := True;
  Effectif := False;
  PaiePeriode := False;
  OnSort := False;

  If Origine = 'EDIT_REMUNERATION' Then
    Begin
      TFQRS1(Ecran).Caption := 'Montants des rémunérations';
      TFQRS1(Ecran).CodeEtat:= 'PMR';
    End;
  If Origine = 'EDIT_MASSESALARIALE' Then
    Begin
      TFQRS1(Ecran).Caption := 'Analyse de la masse salariale';
      TFQRS1(Ecran).CodeEtat:= 'PMS';
    End;
  If Origine = 'EDIT_PREPARIDR' Then
    Begin
      TFQRS1(Ecran).Caption := 'Etat préparatoire IDR';
      TFQRS1(Ecran).CodeEtat:= 'PID';
    End;
  If Origine = 'EDIT_SUIVIEFFECTIF' Then
    Begin
      TFQRS1(Ecran).Caption := 'Salariés non mouvementés';
      TFQRS1(Ecran).CodeEtat:= 'PNM';
      Effectif := True;
      PaiePeriode := True;
    End;
  If Origine = 'EDIT_ENTREESORTIE' Then
    Begin
      TFQRS1(Ecran).Caption := 'Liste des entrées / sorties';
      TFQRS1(Ecran).CodeEtat:= 'PES';
      Effectif := True;
      PaiePeriode := True;
    End;
  If Origine = 'EDIT_NATIONALITE' Then
    Begin
      TFQRS1(Ecran).Caption := 'Liste des nationalités';
      TFQRS1(Ecran).CodeEtat:= 'PLN';
      Effectif := True;
    End;
  If Origine = 'EDIT_SUIAPPRENTI' Then
    Begin
      TFQRS1(Ecran).Caption := 'Suivi des apprentis';
      TFQRS1(Ecran).CodeEtat:= 'PSA';
    End;
  If Origine = 'EDIT_ABSENCE' Then
    Begin
      TFQRS1(Ecran).Caption := 'Suivi des absences';
      TFQRS1(Ecran).CodeEtat:= 'PAR';
      Effectif := True;
    End;
  If Origine = 'EDIT_ANALBRUT' Then
    Begin
      TFQRS1(Ecran).Caption := 'Analyse des composants du brut';
      TFQRS1(Ecran).CodeEtat:= 'PAB';
      Check := TCheckBox(GetControl('CKBULLCOMPL'));
      If Check <> nil Then
      Begin
         Check.Enabled := True;
         Check.Visible := True;
      End;
    End;
  If Origine = 'EDIT_PCHARGES' Then
    Begin
      TFQRS1(Ecran).Caption := 'Masse salariale - Poids des charges';
      TFQRS1(Ecran).CodeEtat:= 'PCH';
      PaiePeriode := True;
      SetControlEnabled('PPU_SALARIE',False);
      SetControlEnabled('CKPERIODE',False);
      SetControlEnabled('CALPHA',False);
      SetControlVisible('TPPU_SALARIE',False);
      SetControlVisible('PPU_SALARIE',False);
      SetControlVisible('CKPERIODE',False);
      SetControlVisible('CALPHA',False);
    End;
  //PT02 Deb ==>
  If Origine = 'EDIT_CLASSIF' Then
    Begin
      TFQRS1(Ecran).Caption := 'Analyse de la classification';
      TFQRS1(Ecran).CodeEtat:= 'PAC';
    End;
  //PT02 Fin <==
  //PT05 Debut ==>
  If Origine = 'EDIT_PRESENT' Then
    Begin
      TFQRS1(Ecran).Caption := 'Liste des salariés présents';
      TFQRS1(Ecran).CodeEtat:= 'PRT';
      Effectif := True;
      PaiePeriode := True;
      Label1 := THLabel(GetControl('TPHC_DATEDEBUT'));
      Label2 := THLabel(GetControl('TPHC_DATEFIN'));
      If (Label1 <> nil) And (Label2 <> nil) Then
      Begin
         Label1.Caption := 'Salariés présents du';
         Label2.Caption := 'au';
      End;
    End;
  //PT05 Fin <===
  UpdateCaption(Ecran);

  Defaut := ThEdit(Getcontrol('DOSSIER'));
  If Defaut <> nil Then Defaut.text := GetParamSoc('SO_LIBELLE');

  DebPer := THEdit(GetControl('XX_VARIABLEDEB'));
  FinPer := THEdit(GetControl('XX_VARIABLEFIN'));
  DebExer:= idate1900;
  FinExer:= idate1900;
  If DebPer <> nil Then DebPer.OnExit := VerifPeriode1;
  If FinPer <> nil Then FinPer.OnExit := VerifPeriode2;

  If PaiePeriode = False Then
  Begin
     If RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) then
     Begin
       If DebPer <> nil Then DebPer.text := DateToStr(DebExer);
       If FinPer <> nil Then FinPer.text := DateToStr(FinExer);
     End;
  End
  Else
  Begin
     If RendPeriodeEnCours(ExerPerEncours, DPer, FPer) Then
     Begin
       If DebPer <> nil Then DebPer.text := DPer;
       If FinPer <> nil Then FinPer.text := FPer;
     End;
  End;

  Check := TCheckBox(GetControl('CKPERIODE'));
  If Check <> nil Then
  Begin
     If Origine = 'EDIT_SUIVIEFFECTIF' Then
     Begin
       Check.Enabled := True;
       Check.Visible := True;
       Check.Caption := 'Salariés présents dans la période';
     End;
     If Origine = 'EDIT_NATIONALITE' Then
     Begin
       Check.Enabled := True;
       Check.Visible := True;
       Check.Caption := 'Exclure nationalité française';
       Label1 := THLabel(GetControl('TPHC_DATEDEBUT'));
       Label2 := THLabel(GetControl('TPHC_DATEFIN'));
       If (Label1 <> nil) And (Label2 <> nil) Then
       Begin
          Label1.Caption := 'Salariés présents du';
          Label2.Caption := 'au';
       End;
     End;
     If Origine = 'EDIT_ABSENCE' Then
     Begin
       Check.Enabled := True;
       Check.Visible := True;
     End;
     Check.Checked := True;
     Check.OnClick := Change;
  End;
  Check := TCheckBox(GetControl('CALPHA'));
  If Check <> nil Then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CETAB'));
  If Check <> nil Then Check.OnClick := Change;
  If Origine = 'EDIT_REMUNERATION' Then ControleChampCompl ('PPU');
  If Origine = 'EDIT_MASSESALARIALE' Then ControleChampCompl ('PPU');
  If Origine = 'EDIT_PREPARIDR' Then ControleChampCompl ('PPU');
  If Origine = 'EDIT_SUIVIEFFECTIF' Then ControleChampCompl ('PSA');
  If Origine = 'EDIT_ENTREESORTIE' Then ControleChampCompl ('PSA');
  If Origine = 'EDIT_NATIONALITE' Then ControleChampCompl ('PSA');
  If Origine = 'EDIT_SUIAPPRENTI' Then ControleChampCompl ('PPU');
  If Origine = 'EDIT_ABSENCE' Then ControleChampCompl ('PSA');
  If Origine = 'EDIT_ANALBRUT' Then ControleChampCompl ('PPU');
  If Origine = 'EDIT_PCHARGES' Then ControleChampCompl ('PPU');
  If Origine = 'EDIT_CLASSIF' Then ControleChampCompl ('PPU'); //PT02
  If Origine = 'EDIT_PRESENT' Then ControleChampCompl ('PSA'); //PT05

  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));

  For I := 1 to 5 do
  Begin
    Check := TCheckBox(GetControl('CN'+IntToStr(I)));
    If Check <> nil Then Check.OnClick := Change;
  End;
  For I := 1 to 4 do
  Begin
    Check := TCheckBox(GetControl('CL'+IntToStr(I)));
    If Check <> nil Then Check.OnClick := Change;
  End;
End;

Procedure TOF_EDIT_REMUNERATION.OnUpdate ;
var
  DebPer ,FinPer ,Ch_Rupt :THEdit;
  DatDeb ,DatFin :TDateTime;
  Ch1 ,Ch2 ,Ch3 :TCheckBox;
  Pages :TPageControl;
  SQL ,Orderby ,AjoutChamp ,Where ,Suite :String;

Begin
  Inherited ;
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);
  DebPer := THEdit(GetControl('XX_VARIABLEDEB'));
  FinPer := THEdit(GetControl('XX_VARIABLEFIN'));
  DatDeb := StrToDate(DebPer.Text);
  DatFin := StrToDate(FinPer.Text);
  Ch1 := TCheckBox(GetControl('CKPERIODE'));
  Ch2 := TCheckBox(GetControl('CETAB'));
  Ch3 := TCheckBox(GetControl('CALPHA'));
  Ch_Rupt := THEdit(GetControl('XX_RUPTURE1'));

  AjoutChamp :='';
  If Trim(Ch_Rupt.Text)<>'' Then AjoutChamp := Trim(Ch_Rupt.Text) + ',' ;
  If Trim(Where)= '' Then
    Begin
      Where := 'WHERE';
      Suite := '';
    End
  Else
    Suite := ' AND';

  If Origine = 'EDIT_ABSENCE' Then //Traitement via une TOB
  Begin
    FreeAndNil(TobEdit);
    TobEdit := Tob.Create('Edit_abs',Nil,-1);
    Edition_absence(AjoutChamp ,Where ,Suite ,DatDeb ,DatFin);
    TFQRS1(Ecran).LaTob:= TobEdit;
    Exit;
  End;
  If Origine = 'EDIT_ANALBRUT' Then //Traitement via une TOB
  Begin
    FreeAndNil(TobEdit);
    TobEdit := Tob.Create('Edit_analBrut',Nil,-1);
    Edition_AnalBrut(AjoutChamp ,Where ,Suite ,DatDeb ,DatFin);
    TFQRS1(Ecran).LaTob:= TobEdit;
    Exit;
  End;
  If Origine = 'EDIT_PCHARGES' Then //Traitement via une TOB
  Begin
    FreeAndNil(TobEdit);
    TobEdit := Tob.Create('Edit_pcharges',Nil,-1);
    Edition_Pcharges(Where ,Suite ,DatDeb ,DatFin);
    TFQRS1(Ecran).LaTob:= TobEdit;
    Exit;
  End;

  If (Ch1 <> nil) And (Ch2 <> nil) And (Ch3 <> nil) And (Ch_Rupt <> nil) Then
  Begin
     If Effectif = False Then
     Begin
       If Origine = 'EDIT_REMUNERATION' Then
       Begin
         If (Ch1.Checked = False) Then
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_LIBELLE,PPU_PRENOM,PPU_CODEPOSTAL,PPU_VILLE,'+
                 'PPU_LIBELLEEMPLOI,PPU_CBRUT,PPU_CNETAPAYER,PPU_CHEURESTRAV,'+ AjoutChamp + 'PPU_CIVILITE' +
                 ' FROM PAIEENCOURS ';
         End
         Else
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,MIN(PPU_DATEDEBUT) AS PPU_DATEDEBUT,MAX(PPU_DATEFIN) AS PPU_DATEFIN,'+
                 'PPU_LIBELLE,PPU_PRENOM,PPU_CODEPOSTAL,PPU_VILLE,'+
                 'PPU_LIBELLEEMPLOI,SUM(PPU_CBRUT) AS PPU_CBRUT,SUM(PPU_CNETAPAYER) AS PPU_CNETAPAYER,'+
                 'SUM(PPU_CHEURESTRAV) AS PPU_CHEURESTRAV,'+ AjoutChamp + 'PPU_CIVILITE' +
                 ' FROM PAIEENCOURS ';
         End;
         SQL := SQL + Where + Suite + ' (PPU_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatFin)+'") ';
         If (Ch1.Checked = True) Then
            SQL := SQL + 'GROUP BY PPU_ETABLISSEMENT,'+ AjoutChamp +'PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CODEPOSTAL,PPU_VILLE,PPU_LIBELLEEMPLOI,PPU_CIVILITE ';
       End;
       If Origine = 'EDIT_MASSESALARIALE' Then
       Begin
         If (Ch1.Checked = False) Then
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_LIBELLE,PPU_PRENOM,'+
                 'PPU_CBRUTFISCAL,PPU_CCOUTPATRON,'+ AjoutChamp + 'PPU_CIVILITE' +
                 ' FROM PAIEENCOURS ';
         End
         Else
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,MIN(PPU_DATEDEBUT) AS PPU_DATEDEBUT,MAX(PPU_DATEFIN) AS PPU_DATEFIN,'+
                 'PPU_LIBELLE,PPU_PRENOM,SUM(PPU_CBRUTFISCAL) AS PPU_CBRUTFISCAL,'+
                 'SUM(PPU_CCOUTPATRON) AS PPU_CCOUTPATRON,'+ AjoutChamp + 'PPU_CIVILITE' +
                 ' FROM PAIEENCOURS ';
         End;
         SQL := SQL + Where + Suite + ' (PPU_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatFin)+'") ';
         If (Ch1.Checked = True) Then
            SQL := SQL + 'GROUP BY PPU_ETABLISSEMENT,'+ AjoutChamp +'PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CIVILITE ';
       End;
       If Origine = 'EDIT_PREPARIDR' Then
       Begin
         If (Ch1.Checked = False) Then
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_LIBELLE,PPU_PRENOM,'+
                 'PPU_CBRUT,'+ AjoutChamp + 'PPU_CIVILITE,PSA_DATENAISSANCE,PSA_SEXE,PSA_QUALIFICATION,PSA_DATEANCIENNETE' +
                 ' FROM PAIEENCOURS LEFT JOIN SALARIES on (PPU_SALARIE=PSA_SALARIE And PPU_LIBELLE=PSA_LIBELLE) ';
         End
         Else
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,MIN(PPU_DATEDEBUT) AS PPU_DATEDEBUT,MAX(PPU_DATEFIN) AS PPU_DATEFIN,'+
                 'PPU_LIBELLE,PPU_PRENOM,SUM(PPU_CBRUT) AS PPU_CBRUT,'+ AjoutChamp + 'PPU_CIVILITE' +
                 ',PSA_DATENAISSANCE,PSA_SEXE,PSA_QUALIFICATION,PSA_DATEANCIENNETE' +
                 ' FROM PAIEENCOURS LEFT JOIN SALARIES on (PPU_SALARIE=PSA_SALARIE And PPU_LIBELLE=PSA_LIBELLE) ';
         End;
         SQL := SQL + Where + Suite + ' (PPU_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatFin)+'") ';
         If (Ch1.Checked = True) Then
            SQL := SQL + 'GROUP BY PPU_ETABLISSEMENT,'+ AjoutChamp +'PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CIVILITE,PSA_DATENAISSANCE,PSA_SEXE,PSA_QUALIFICATION,PSA_DATEANCIENNETE ';
       End;
       If Origine = 'EDIT_SUIAPPRENTI' Then
       Begin
         If (Ch1.Checked = False) Then
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_LIBELLE,PPU_PRENOM,'+
                 'PPU_CBRUT,PPU_CNETAPAYER,PPU_CHEURESTRAV,PPU_CCOUTPATRON,'+ AjoutChamp + 'PPU_CIVILITE,PSA_DATEENTREE,PSA_DATESORTIE,PSA_DATENAISSANCE' +
                 ' FROM PAIEENCOURS LEFT JOIN SALARIES on (PPU_SALARIE=PSA_SALARIE And PPU_LIBELLE=PSA_LIBELLE) ';
         End
         Else
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,MIN(PPU_DATEDEBUT) AS PPU_DATEDEBUT,MAX(PPU_DATEFIN) AS PPU_DATEFIN,'+
                 'PPU_LIBELLE,PPU_PRENOM,SUM(PPU_CBRUT) AS PPU_CBRUT,SUM(PPU_CNETAPAYER) AS PPU_CNETAPAYER,SUM(PPU_CHEURESTRAV) AS PPU_CHEURESTRAV,'+
                 'SUM(PPU_CCOUTPATRON) AS PPU_CCOUTPATRON,'+ AjoutChamp + 'PPU_CIVILITE,PSA_DATEENTREE,PSA_DATESORTIE,PSA_DATENAISSANCE' +
                 ' FROM PAIEENCOURS LEFT JOIN SALARIES on (PPU_SALARIE=PSA_SALARIE And PPU_LIBELLE=PSA_LIBELLE) ';
         End;
         SQL := SQL + Where + Suite + ' (PPU_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatFin)+'" And PSA_CATDADS = "003") ';
         If (Ch1.Checked = True) Then
            SQL := SQL + 'GROUP BY PPU_ETABLISSEMENT,'+ AjoutChamp +'PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CIVILITE,PSA_DATEENTREE,PSA_DATESORTIE,PSA_DATENAISSANCE ';
       End;
       //PT02 Debut ==>
       If Origine = 'EDIT_CLASSIF' Then
       Begin
         If (Ch1.Checked = False) Then
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_LIBELLEEMPLOI,PPU_CIVILITE,' +
                 'PPU_DATEDEBUT,PPU_DATEFIN,PPU_BULCOMPL,'+ AjoutChamp +
                 'PPU_COEFFICIENT,PPU_QUALIFICATION,PPU_NIVEAU,PPU_INDICE,PHC_MONTANT' +
                 ' FROM PAIEENCOURS' +
                 ' LEFT JOIN HISTOCUMSAL ON (PPU_ETABLISSEMENT = PHC_ETABLISSEMENT And PPU_SALARIE = PHC_SALARIE) And' +
                 ' (PHC_DATEDEBUT = PPU_DATEDEBUT AND PHC_DATEFIN = PPU_DATEFIN) ';
         End
         Else
         Begin
            SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_LIBELLEEMPLOI,PPU_CIVILITE,' +
                 'MIN(PPU_DATEDEBUT) AS PPU_DATEDEBUT,MAX(PPU_DATEFIN) AS PPU_DATEFIN,PPU_BULCOMPL,'+ AjoutChamp +
                 'PPU_COEFFICIENT,PPU_QUALIFICATION,PPU_NIVEAU,PPU_INDICE,SUM(PHC_MONTANT) AS PHC_MONTANT' +
                 ' FROM PAIEENCOURS' +
                 ' LEFT JOIN HISTOCUMSAL ON (PPU_ETABLISSEMENT = PHC_ETABLISSEMENT And PPU_SALARIE = PHC_SALARIE) And' +
                 ' (PHC_DATEDEBUT = PPU_DATEDEBUT AND PHC_DATEFIN = PPU_DATEFIN) ';
         End;
         SQL := SQL + Where + Suite + ' (PPU_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatFin)+'") And PPU_BULCOMPL <> "X" AND PHC_CUMULPAIE = "03" ';
         If (Ch1.Checked = True) Then
            SQL := SQL + 'GROUP BY PPU_ETABLISSEMENT,'+ AjoutChamp +'PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CIVILITE,' +
                         'PPU_BULCOMPL,PPU_LIBELLEEMPLOI,PPU_COEFFICIENT,PPU_QUALIFICATION,PPU_NIVEAU,PPU_INDICE ';
       End;
       //PT02 Fin <==
       Orderby := 'Order by ';
       If (Ch2.Checked = True) And (Norme = True) Then
          Orderby := Orderby + 'PPU_ETABLISSEMENT,';

       If (Ch3.Checked = True) Then
          Orderby := Orderby + AjoutChamp + 'PPU_LIBELLE,PPU_PRENOM,'
         Else
          Orderby := Orderby + AjoutChamp + 'PPU_SALARIE,';

       Orderby := Orderby + 'PPU_DATEDEBUT,PPU_DATEFIN';
       SQL := SQL + Orderby ;
     End
     Else  //Effectif = true
     Begin
       If Origine = 'EDIT_SUIVIEFFECTIF' Then
       Begin
         SQL:='SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_DATEENTREE,' +
              'PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_CIVILITE,'+ AjoutChamp + 'PSA_MOTIFSUSPPAIE' +
              ' FROM SALARIES ';
         SQL := SQL + Where + Suite + ' PSA_SALARIE NOT IN (SELECT DISTINCT PPU_SALARIE From PAIEENCOURS Where ' +
          '(PPU_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatFin)+'")) ';

         If (Ch1.Checked = True) Then
            SQL := SQL + 'And ((PSA_DATEENTREE <= "'+UsDateTime(DatFin)+'") And ' +
                   '(PSA_DATESORTIE >= "'+UsDateTime(DatDeb)+'" Or PSA_DATESORTIE = "01/01/1900")) ';
       End;
       If Origine = 'EDIT_ENTREESORTIE' Then
       Begin
         //PT03
         //SQL:='SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SORTIEDEFINIT,' +
         //     'PSA_CIVILITE,PSA_MOTIFENTREE,PSA_MOTIFSORTIE,'+ AjoutChamp + 'PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT' +
         //     ' FROM SALARIES' +
         //     ' LEFT JOIN CONTRATTRAVAIL On ((PCI_SALARIE = PSA_SALARIE) And (PCI_ETABLISSEMENT = PSA_ETABLISSEMENT) ' +
         //     'And (PCI_DEBUTCONTRAT >= PSA_DATEENTREE)) ';
         SQL:='SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SORTIEDEFINIT,' +
              'PSA_CIVILITE,PSA_MOTIFENTREE,PSA_MOTIFSORTIE,PSA_NUMEROSS,PSA_DADSPROF,'+ AjoutChamp + 'PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,' +
              ' (SELECT CO_LIBELLE FROM COMMUN WHERE CO_TYPE="PDP" AND CO_CODE = PSA_DADSPROF) AS STATUTPROF'+
              ' FROM SALARIES' +
              ' LEFT JOIN CONTRATTRAVAIL On ((PCI_SALARIE = PSA_SALARIE) And (PCI_ETABLISSEMENT = PSA_ETABLISSEMENT)) ';
//              ' And (PCI_DEBUTCONTRAT >= PSA_DATEENTREE)) '; //PT4
         //Fin PT03
         SQL := SQL + Where + Suite + ' ((PSA_DATEENTREE >= "'+UsDateTime(DatDeb)+'" And PSA_DATEENTREE <= "'+UsDateTime(DatFin)+'")';
         SQL := SQL + ' Or (PSA_DATESORTIE >= "'+UsDateTime(DatDeb)+'" And PSA_DATESORTIE <= "'+UsDateTime(DatFin)+'")) ';
       End;
       If Origine = 'EDIT_NATIONALITE' Then
       Begin
         SQL:='SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_CIVILITE,' +
              'PSA_NATIONALITE,PSA_CARTESEJOUR,PSA_DATEXPIRSEJOUR,'+ AjoutChamp + 'PSA_DELIVPAR' +
              ' FROM SALARIES ';
         SQL := SQL + Where + Suite + ' ((PSA_DATEENTREE <= "'+UsDateTime(DatFin)+'")' +
           ' AND (PSA_DATESORTIE >= "'+UsDateTime(DatDeb)+'" OR PSA_DATESORTIE = "'+UsDateTime(idate1900)+'")) ';

         If (Ch1.Checked = True) Then
            SQL := SQL + 'And (PSA_NATIONALITE <> "FRA") ';
       End;
       //PT05 Debut ====>
       If Origine = 'EDIT_PRESENT' Then
       Begin
         SQL:='SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_CIVILITE,PSA_NUMEROSS,' +
              'PSA_DATEENTREE,PSA_DATESORTIE,PSA_DATENAISSANCE,PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,' +
              'PSA_CODEPOSTAL,PSA_VILLE,'+ AjoutChamp + 'PSA_LIBELLEEMPLOI' +
              ' FROM SALARIES ';
         SQL := SQL + Where + Suite + ' ((PSA_DATEENTREE <= "'+UsDateTime(DatFin)+'")' +
           ' AND (PSA_DATESORTIE >= "'+UsDateTime(DatDeb)+'" OR PSA_DATESORTIE = "'+UsDateTime(idate1900)+'")) ';
       End;
       //PT05 Fin <=====
       Orderby := 'Order by ';
       If (Ch2.Checked = True) And (Norme = True) Then
          Orderby := Orderby + 'PSA_ETABLISSEMENT,';

       If (Ch3.Checked = True) Then
          Orderby := Orderby + AjoutChamp + 'PSA_LIBELLE,PSA_PRENOM,'
         Else
          Orderby := Orderby + AjoutChamp + 'PSA_SALARIE,';

       Orderby := Orderby + 'PSA_DATEENTREE,PSA_DATESORTIE';
       SQL := SQL + Orderby ;
     End;
     TFQRS1(Ecran).WhereSQL:=SQL;
  End;
End;

//Evite le messsage : Impossible de Focaliser une fenetre désactivée ou invisible
// quand on quitte la fenetre par la croix alors qu'une date est en erreur
Procedure TOF_EDIT_REMUNERATION.OnClose ;
Var
  DateDefaut : THEdit;

Begin
  Inherited ;
  OnSort := True;
  DateDefaut := THEdit(GetControl('XX_VARIABLEDEB'));
  If DateDefaut <> nil Then
     If (not(IsValidDate(DateDefaut.text))) Then
        SetControltext('XX_VARIABLEDEB',DateToStr(Date));

  DateDefaut := THEdit(GetControl('XX_VARIABLEFIN'));
  If DateDefaut <> nil Then
     If (not(IsValidDate(DateDefaut.text))) Then
        SetControltext('XX_VARIABLEFIN',DateToStr(Date));

  If (Origine = 'EDIT_ABSENCE') Or (Origine = 'EDIT_ANALBRUT') Or (Origine = 'EDIT_PCHARGES') Then
     FreeAndNil(TobEdit);
End ;

Procedure TOF_EDIT_REMUNERATION.Change(Sender: TObject) ;
var
  Ch_RuptEtab :THEdit;
  Check: TCheckBox;

Begin
  Ch_RuptEtab := THEdit(GetControl('XX_RUPTURE2'));
  Check := TCheckBox(GetControl('CETAB'));
  If Ch_RuptEtab <> nil Then Ch_RuptEtab.Text := '';

  If Norme = True Then
    Begin
    //Nouvelle norme : 2 ruptures possibles (Etab + Champ libre)
      ControleCheckRupture ;
      If Origine = 'EDIT_REMUNERATION' Then InformeRupture('PPU');
      If Origine = 'EDIT_MASSESALARIALE' Then InformeRupture('PPU');
      If Origine = 'EDIT_PREPARIDR' Then InformeRupture('PPU');
      If Origine = 'EDIT_SUIVIEFFECTIF' Then InformeRupture('PSA');
      If Origine = 'EDIT_ENTREESORTIE' Then InformeRupture('PSA');
      If Origine = 'EDIT_NATIONALITE' Then InformeRupture('PSA');
      If Origine = 'EDIT_SUIAPPRENTI' Then InformeRupture('PPU');
      If Origine = 'EDIT_ABSENCE' Then InformeRupture('PSA');
      If Origine = 'EDIT_ANALBRUT' Then InformeRupture('PPU');
      If Origine = 'EDIT_PCHARGES' Then InformeRupture('PPU');
      If Origine = 'EDIT_CLASSIF' Then InformeRupture('PPU'); //PT02
      If Origine = 'EDIT_PRESENT' Then InformeRupture('PSA'); //PT05

      If (Ch_RuptEtab <> nil) And (Check <> nil) Then
      Begin
        If Effectif = False Then
        Begin
           If (Check.Checked=True) Then Ch_RuptEtab.Text := 'PPU_ETABLISSEMENT';
        End
        Else
           If (Check.Checked=True) Then Ch_RuptEtab.Text := 'PSA_ETABLISSEMENT';
      End;
    End
  Else
    Begin
    //Ancienne norme : 1 rupture possible
      BloqueChampLibre(Ecran);
      RecupChampRupture(Ecran);
    End;
End;

Procedure TOF_EDIT_REMUNERATION.VerifPeriode1(Sender: TObject) ;
var
  DebPer ,FinPer :THEdit;
  DatDeb ,DatFin :TDateTime;

Begin
  If OnSort = False Then
  Begin
    DebPer := THEdit(GetControl('XX_VARIABLEDEB'));
    FinPer := THEdit(GetControl('XX_VARIABLEFIN'));
    If (DebPer <> nil) And (FinPer <> nil) Then
    Begin
      DatDeb := StrToDate(DebPer.Text);
      DatFin := StrToDate(FinPer.Text);
      If DatDeb > DatFin Then
      Begin
        PGIBOX('La date saisie doit être inférieure à '+FinPer.Text+'',''+TFQRS1(Ecran).Caption+'');
        DebPer.SetFocus;
      End;
    End;
  End;
End;

Procedure TOF_EDIT_REMUNERATION.VerifPeriode2(Sender: TObject) ;
var
  DebPer ,FinPer :THEdit;
  DatDeb ,DatFin :TDateTime;

Begin
  If OnSort = False Then
  Begin
    DebPer := THEdit(GetControl('XX_VARIABLEDEB'));
    FinPer := THEdit(GetControl('XX_VARIABLEFIN'));
    If (DebPer <> nil) And (FinPer <> nil) Then
    Begin
      DatDeb := StrToDate(DebPer.Text);
      DatFin := StrToDate(FinPer.Text);
      If DatDeb > DatFin Then
      Begin
        PGIBOX('La date saisie doit être supérieure à '+DebPer.Text+'',''+TFQRS1(Ecran).Caption+'');
        FinPer.SetFocus;
      End;
    End;
  End;
End;

Procedure TOF_EDIT_REMUNERATION.ControleChampCompl (PrefTable : String) ;
var
  Ch1, Ch2, Ch3: string;
  I: Integer;
  Champ: THValComboBox;
  Check: TCheckBox;

Begin
  If Trim(PrefTable) = '' Then Exit;
  Ch1 := PrefTable + '_CODESTAT';
  Ch2 := 'T' + PrefTable + '_CODESTAT';
  Ch3 := 'R_CODESTAT';
  VisibiliteStat(GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
  VisibiliteStat(GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
  VisibiliteStat(GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
  Champ := THValComboBox(GetControl(Ch1));
  If (Champ <> nil) And (Champ.Visible = True) Then
  Begin
     Check := TCheckBox(GetControl('CN5'));
     If Check <> nil Then
     Begin
       Check.Visible := True;
       Check.Enabled := True;
     End;
  End;

  For I := 1 to 4 do
  Begin
    Ch1 := PrefTable + '_TRAVAILN' + IntToStr(I);
    Ch2 := 'T' + PrefTable + '_TRAVAILN' + IntToStr(I);
    Ch3 := 'R_TRAVAILN' + IntToStr(I);
    VisibiliteChampSalarie(IntToStr(I), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
    VisibiliteChampSalarie(IntToStr(I), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
    VisibiliteChampSalarie(IntToStr(I), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    Champ := THValComboBox(GetControl(Ch1));
    If (Champ <> nil) And (Champ.Visible = True) Then
    Begin
       Check := TCheckBox(GetControl('CN'+IntToStr(I)));
       If Check <> nil Then
       Begin
         Check.Visible := True;
         Check.Enabled := True;
       End;
    End;
  End;

  For I := 1 to 4 do
  Begin
    Ch1 := PrefTable + '_LIBREPCMB' + IntToStr(I);
    Ch2 := 'T' + PrefTable + '_LIBREPCMB' + IntToStr(I);
    Ch3 := 'R_LIBREPCMB' + IntToStr(I);
    VisibiliteChampLibreSal(IntToStr(I), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
    VisibiliteChampLibreSal(IntToStr(I), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
    VisibiliteChampLibreSal(IntToStr(I), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    Champ := THValComboBox(GetControl(Ch1));
    If (Champ <> nil) And (Champ.Visible = True) Then
    Begin
       Check := TCheckBox(GetControl('CL'+IntToStr(I)));
       If Check <> nil Then
       Begin
         Check.Visible := True;
         Check.Enabled := True;
       End;
    End;
  End;
End;

Procedure TOF_EDIT_REMUNERATION.ControleCheckRupture ;
var
  TabLieuTravail : array[1..10] of TCheckBox;
  PosCheck,PosUnCheck,i : integer;
  Ok : boolean ;

Begin
  TabLieuTravail[1]:=TCheckBox(GetControl('CN1'));
  TabLieuTravail[2]:=TCheckBox(GetControl('CN2'));
  TabLieuTravail[3]:=TCheckBox(GetControl('CN3'));
  TabLieuTravail[4]:=TCheckBox(GetControl('CN4'));
  TabLieuTravail[5]:=TCheckBox(GetControl('CN5'));
  TabLieuTravail[6]:=TCheckBox(GetControl('CL1'));
  TabLieuTravail[7]:=TCheckBox(GetControl('CL2'));
  TabLieuTravail[8]:=TCheckBox(GetControl('CL3'));
  TabLieuTravail[9]:=TCheckBox(GetControl('CL4'));
  TabLieuTravail[10]:=nil;
  PosUnCheck:=0;
  PosCheck:=0;

  For i:=1 to 9 do
     If (TabLieuTravail[i]<>nil) Then Ok:=False Else Begin Ok:=True; break; End;

  If Ok=False Then
  Begin
    //Coche une rupture
    For i:=1 to 9 do
      If (TabLieuTravail[i].checked=True) Then PosCheck:=i;
    If PosCheck > 0 Then
      For i:=1 to 9 do
         If i<>PosCheck then TabLieuTravail[i].enabled:=False;

    //Décoche une rupture ,  rend enable(True) les autres champs de rupture
    For i:=1 to 9 do
      If (TabLieuTravail[i].checked=False) and (TabLieuTravail[i].enabled=True) then PosUnCheck:=i;
    If (PosCheck=0) and (PosUnCheck>0) then
      For i:=1 to 9 do
         TabLieuTravail[i].enabled:=True;
  End;
End;

Procedure TOF_EDIT_REMUNERATION.InformeRupture(PrefTable : String);
var
  CN1,CN2,CN3,CN4,CN5,CL1,CL2,CL3,CL4:TCheckBox;
  Rupture,Champ1 : THEdit;

Begin
  CN1:=TCheckBox(GetControl('CN1'));
  CN2:=TCheckBox(GetControl('CN2'));
  CN3:=TCheckBox(GetControl('CN3'));
  CN4:=TCheckBox(GetControl('CN4'));
  CN5:=TCheckBox(GetControl('CN5'));
  CL1:=TCheckBox(GetControl('CL1'));
  CL2:=TCheckBox(GetControl('CL2'));
  CL3:=TCheckBox(GetControl('CL3'));
  CL4:=TCheckBox(GetControl('CL4'));
  Rupture:=THEdit(GetControl('XX_RUPTURE1'));
  Champ1 :=THEdit(GetControl('XX_VARIABLE1'));

  If (Champ1<>nil) and (Rupture<>nil) Then
      Champ1.text:='';Rupture.Text:='';
  If (CN1<>nil) and (CN2<>nil) and (CN3<>nil) and (CN4<>nil) and (CN5<>nil) Then
     If (Champ1<>nil) and (Rupture<>nil) Then
     Begin
       If (CN1.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat1;
          Rupture.Text:=PrefTable+'_TRAVAILN1';
       End;
       If (CN2.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat2;
          Rupture.Text:=PrefTable+'_TRAVAILN2';
       End;
       If (CN3.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat3;
          Rupture.Text:=PrefTable+'_TRAVAILN3';
       End;
       If (CN4.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat4;
          Rupture.Text:=PrefTable+'_TRAVAILN4';
       End;
       If (CN5.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibCodeStat;
          Rupture.Text:=PrefTable+'_CODESTAT';
       End;
     End;
  If (CL1<>nil) and (CL2<>nil) and (CL3<>nil) and (CL4<>nil) Then
     If (Champ1<>nil) and (Rupture<>nil) Then
     Begin
       If (CL1.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo1;
          Rupture.Text:=PrefTable+'_LIBREPCMB1';
       End;
       If (CL2.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo2;
          Rupture.Text:=PrefTable+'_LIBREPCMB2';
       End;
       If (CL3.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo3;
          Rupture.Text:=PrefTable+'_LIBREPCMB3';
       End;
       If (CL4.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo4;
          Rupture.Text:=PrefTable+'_LIBREPCMB4';
       End;
     End;
End;

Procedure TOF_EDIT_REMUNERATION.Edition_absence(AjoutChamp ,Where ,Suite : String; DatDeb ,DatFin :TDateTime);
var
  Ch1 ,Ch2 ,Ch3 :TCheckBox;
  Ch_Rupt :THEdit;
  SQL ,Orderby ,Ctrouve :String;
  i ,y :Integer;
  Q :TQuery;
  TSQL ,Tcom ,T :Tob;
  J1,J2,J3,J4,J5,H1,H2,H3,H4,H5 :Double;
  DebMin ,FinMax ,PaieMax :String;

Begin
  Ch1 := TCheckBox(GetControl('CKPERIODE'));
  Ch2 := TCheckBox(GetControl('CETAB'));
  Ch3 := TCheckBox(GetControl('CALPHA'));
  Ch_Rupt := THEdit(GetControl('XX_RUPTURE1'));

  If (Ch1 <> nil) And (Ch2 <> nil) And (Ch3 <> nil) And (Ch_Rupt <> nil) Then
  Begin
  {deb PT7
    SQL:='SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_CIVILITE,PSA_HORAIREMOIS,'+ AjoutChamp +
         'PPU_CHEURESTRAV,PCN_TYPEMVT,PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_DATEPAIEMENT,PCN_DATEFIN,PCN_CODETAPE,' +
         'PCN_JOURS,PCN_HEURES,PMA_TYPEABS' +
         ' FROM SALARIES ' +
         'LEFT JOIN ABSENCESALARIE ON ((PSA_SALARIE = PCN_SALARIE) And ' +
         '(PCN_TYPEMVT = "ABS" Or (PCN_TYPEMVT = "CPA" And PCN_TYPECONGE = "PRI"))) ' +
         'LEFT JOIN PAIEENCOURS ON (PSA_SALARIE = PPU_SALARIE And PSA_LIBELLE = PPU_LIBELLE) ' +
         'And (PCN_DATEDEBUTABS >= PPU_DATEDEBUT AND PCN_DATEFINABS <= PPU_DATEFIN) ' +
         'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PCN_TYPECONGE=PMA_MOTIFABSENCE ';}
    SQL:='SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_CIVILITE,PSA_HORAIREMOIS,'+ AjoutChamp +
         'PPU_CHEURESTRAV,PCN_TYPEMVT,PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_DATEPAIEMENT,PCN_DATEFIN,PCN_CODETAPE,' +
         'PCN_JOURS,PCN_HEURES,PMA_TYPEABS' +
         ' FROM SALARIES ' +
         'LEFT JOIN ABSENCESALARIE ON PSA_SALARIE = PCN_SALARIE ' +
         'LEFT JOIN PAIEENCOURS ON PSA_SALARIE = PPU_SALARIE And PSA_LIBELLE = PPU_LIBELLE ' +
         'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PCN_TYPECONGE=PMA_MOTIFABSENCE ';

  //  SQL := SQL + Where + Suite + ' (PCN_DATEDEBUTABS >= "'+UsDateTime(DatDeb)+'" AND PCN_DATEFINABS <= "'+UsDateTime(DatFin)+'") ';
    SQL := SQL + Where + Suite + ' (PCN_DATEDEBUTABS >= "'+UsDateTime(DatDeb)+'" AND PCN_DATEFINABS <= "'+UsDateTime(DatFin)+'") '+
           'AND (PCN_TYPEMVT = "ABS" Or (PCN_TYPEMVT = "CPA" And PCN_TYPECONGE = "PRI")) '+
           'AND (PCN_DATEDEBUTABS >= PPU_DATEDEBUT AND PCN_DATEFINABS <= PPU_DATEFIN) ';
  // fin pt7
    SQL := SQL + 'And (PCN_MVTDUPLIQUE <> "X") '; //PT01

    Orderby := 'Order by ';
    If (Ch2.Checked = True) And (Norme = True) Then
       Orderby := Orderby + 'PSA_ETABLISSEMENT,';

    If (Ch3.Checked = True) Then
       Orderby := Orderby + AjoutChamp + 'PSA_LIBELLE,PSA_PRENOM,'
     Else
       Orderby := Orderby + AjoutChamp + 'PSA_SALARIE,';

    Orderby := Orderby + 'PCN_DATEDEBUTABS,PCN_DATEFINABS';
    SQL := SQL + Orderby ;
    Q := OpenSQL(SQL,True);
    TSQL := Tob.Create('Edit1',Nil,-1);
    TSQL.LoadDetailDB('Edit1','','',Q,False);
    Ferme(Q);
    //TOBDebug(TSQL);

    SQL := 'SELECT CO_CODE,CO_LIBRE FROM COMMUN Where CO_TYPE = "PTX"';
    Q := OpenSQL(SQL,True);
    Tcom := Tob.Create('Edit2',Nil,-1);
    Tcom.LoadDetailDB('Edit2','','',Q,False);
    Ferme(Q);

    J1 := 0;
    J2 := 0;
    J3 := 0;
    J4 := 0;
    J5 := 0;
    H1 := 0;
    H2 := 0;
    H3 := 0;
    H4 := 0;
    H5 := 0;
    DebMin := '31/12/2099';
    FinMax := '01/01/1900';
    PaieMax:= '01/01/1900';

    For i := 0 To TSQL.Detail.Count -1 do
    begin
       If TSQL.Detail[i].GetValue('PCN_TYPEMVT') = 'CPA' Then Ctrouve := 'CPP'
       Else
       Begin
          Ctrouve := 'AUT';
          For y := 0 To Tcom.Detail.Count -1 do
          begin
             If Tcom.Detail[y].GetValue('CO_CODE') = TSQL.Detail[i].GetValue('PMA_TYPEABS') Then
             Begin
                Ctrouve := Tcom.Detail[y].GetValue('CO_LIBRE');
                Break;
             End;
          End;
       End;
       If Ctrouve = 'CPP' Then
       Begin
          J1 := J1 + TSQL.Detail[i].GetValue('PCN_JOURS');
          H1 := H1 + TSQL.Detail[i].GetValue('PCN_HEURES');
       End;
       If Ctrouve = 'RTT' Then
       Begin
          J2 := J2 + TSQL.Detail[i].GetValue('PCN_JOURS');
          H2 := H2 + TSQL.Detail[i].GetValue('PCN_HEURES');
       End;
       If Ctrouve = 'MAL' Then
       Begin
          J3 := J3 + TSQL.Detail[i].GetValue('PCN_JOURS');
          H3 := H3 + TSQL.Detail[i].GetValue('PCN_HEURES');
       End;
       If Ctrouve = 'MAT' Then
       Begin
          J4 := J4 + TSQL.Detail[i].GetValue('PCN_JOURS');
          H4 := H4 + TSQL.Detail[i].GetValue('PCN_HEURES');
       End;
       If Ctrouve = 'AUT' Then
       Begin
          J5 := J5 + TSQL.Detail[i].GetValue('PCN_JOURS');
          H5 := H5 + TSQL.Detail[i].GetValue('PCN_HEURES');
       End;
       If IsValidDate(TSQL.Detail[i].GetValue('PCN_DATEDEBUTABS')) Then
          If StrToDate(DebMin) > StrToDate(TSQL.Detail[i].GetValue('PCN_DATEDEBUTABS')) Then
             DebMin := TSQL.Detail[i].GetValue('PCN_DATEDEBUTABS');

       If IsValidDate(TSQL.Detail[i].GetValue('PCN_DATEFINABS')) Then
          If StrToDate(FinMax) < StrToDate(TSQL.Detail[i].GetValue('PCN_DATEFINABS')) Then
             FinMax := TSQL.Detail[i].GetValue('PCN_DATEFINABS');

       If TSQL.Detail[i].GetValue('PCN_TYPEMVT') = 'CPA' Then
       Begin
          If TSQL.Detail[i].GetValue('PCN_CODETAPE') = 'P' Then
             If IsValidDate(TSQL.Detail[i].GetValue('PCN_DATEFIN')) Then
                If StrToDate(PaieMax) < StrToDate(TSQL.Detail[i].GetValue('PCN_DATEFIN')) Then
                   PaieMax := TSQL.Detail[i].GetValue('PCN_DATEFIN');
       End
       Else
       Begin
          If IsValidDate(TSQL.Detail[i].GetValue('PCN_DATEPAIEMENT')) Then
             If StrToDate(PaieMax) < StrToDate(TSQL.Detail[i].GetValue('PCN_DATEPAIEMENT')) Then
                PaieMax := TSQL.Detail[i].GetValue('PCN_DATEPAIEMENT');
       End;

       If (Ch1.Checked = False) Or
          ((Ch1.Checked = True) And (i = TSQL.Detail.Count -1)) Or
          ((Ch1.Checked = True) And (TSQL.Detail[i].GetValue('PSA_SALARIE')<>TSQL.Detail[i+1].GetValue('PSA_SALARIE'))) Then
       Begin
          T := Tob.Create('Edit3',TobEdit,-1);
          If (Ch2.Checked = True) And (Norme = True) Then
              T.AddChampSupValeur('RUPTURE1',TSQL.Detail[i].GetValue('PSA_ETABLISSEMENT'))
           Else
              T.AddChampSupValeur('RUPTURE1','');
          If Trim(AjoutChamp) <> '' Then
              T.AddChampSupValeur('RUPTURE2',TSQL.Detail[i].GetValue(Ch_Rupt.Text))
           Else
              T.AddChampSupValeur('RUPTURE2','');
          If (Ch1.Checked = False) Then
              T.AddChampSupValeur('RUPTURE3',TSQL.Detail[i].GetValue('PSA_SALARIE'))
           Else
              T.AddChampSupValeur('RUPTURE3','');

          T.AddChampSupValeur('PSA_ETABLISSEMENT',TSQL.Detail[i].GetValue('PSA_ETABLISSEMENT'));
          T.AddChampSupValeur('PSA_SALARIE',TSQL.Detail[i].GetValue('PSA_SALARIE'));
          T.AddChampSupValeur('PSA_LIBELLE',TSQL.Detail[i].GetValue('PSA_LIBELLE'));
          T.AddChampSupValeur('PSA_PRENOM',TSQL.Detail[i].GetValue('PSA_PRENOM'));
          T.AddChampSupValeur('PSA_CIVILITE',TSQL.Detail[i].GetValue('PSA_CIVILITE'));

          If (Ch1.Checked = False) Then
              T.AddChampSupValeur('PSA_HORAIREMOIS',TSQL.Detail[i].GetValue('PSA_HORAIREMOIS'))
           Else
              T.AddChampSupValeur('PSA_HORAIREMOIS','');

          If Trim(AjoutChamp) <> '' Then
              T.AddChampSupValeur(Ch_Rupt.Text,TSQL.Detail[i].GetValue(Ch_Rupt.Text));

          If (Ch1.Checked = False) Then
              T.AddChampSupValeur('PPU_CHEURESTRAV',TSQL.Detail[i].GetValue('PPU_CHEURESTRAV'))
           Else
              T.AddChampSupValeur('PPU_CHEURESTRAV','');

          T.AddChampSupValeur('PCN_TYPEMVT',TSQL.Detail[i].GetValue('PCN_TYPEMVT'));
          T.AddChampSupValeur('PCN_TYPECONGE',TSQL.Detail[i].GetValue('PCN_TYPECONGE'));
          T.AddChampSupValeur('PCN_DATEDEBUTABS',DebMin);
          T.AddChampSupValeur('PCN_DATEFINABS',FinMax);
          T.AddChampSupValeur('PCN_DATEPAIEMENT',PaieMax);
          T.AddChampSupValeur('PCN_JOURS',TSQL.Detail[i].GetValue('PCN_JOURS'));
          T.AddChampSupValeur('PCN_HEURES',TSQL.Detail[i].GetValue('PCN_HEURES'));
          T.AddChampSupValeur('PMA_TYPEABS',TSQL.Detail[i].GetValue('PMA_TYPEABS'));
          T.AddChampSupValeur('J1',J1);
          T.AddChampSupValeur('H1',H1);
          T.AddChampSupValeur('J2',J2);
          T.AddChampSupValeur('H2',H2);
          T.AddChampSupValeur('J3',J3);
          T.AddChampSupValeur('H3',H3);
          T.AddChampSupValeur('J4',J4);
          T.AddChampSupValeur('H4',H4);
          T.AddChampSupValeur('J5',J5);
          T.AddChampSupValeur('H5',H5);
          J1 := 0;
          J2 := 0;
          J3 := 0;
          J4 := 0;
          J5 := 0;
          H1 := 0;
          H2 := 0;
          H3 := 0;
          H4 := 0;
          H5 := 0;
          DebMin := '31/12/2099';
          FinMax := '01/01/1900';
          PaieMax:= '01/01/1900';
       End;
    End;
    TSQL.Free;
    Tcom.Free;
    //TOBDebug(TobEdit);
  End;
End;

Procedure TOF_EDIT_REMUNERATION.Edition_AnalBrut(AjoutChamp ,Where ,Suite : String; DatDeb ,DatFin :TDateTime);
var
  Ch1 ,Ch2 ,Ch3 ,Ch4 :TCheckBox;
  Ch_Rupt :THEdit;
  SQL ,Orderby ,DebMin ,FinMax :String;
  i ,y ,S :Integer;
  Q :TQuery;
  TSQL ,T ,Tprofil :Tob;
  M1,M2,M3,M4,M5,M6,M7,M8 :Double;
  Cfait ,Cprofil : Boolean;

Begin
  Ch1 := TCheckBox(GetControl('CKPERIODE'));
  Ch2 := TCheckBox(GetControl('CETAB'));
  Ch3 := TCheckBox(GetControl('CALPHA'));
  Ch4 := TCheckBox(GetControl('CKBULLCOMPL'));
  Ch_Rupt := THEdit(GetControl('XX_RUPTURE1'));

  If (Ch1 <> nil) And (Ch2 <> nil) And (Ch3 <> nil) And (Ch4 <> nil) And (Ch_Rupt <> nil) Then
  Begin
    SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_LIBELLE,PPU_PRENOM,PPU_CIVILITE,PPU_BULCOMPL,'+ AjoutChamp +
         'PSA_PROFILCGE,PHB_NATURERUB,PHB_RUBRIQUE,PHB_IMPRIMABLE,PHB_MTREM,PHB_SENSBUL,PRM_THEMEREM,PCR_CUMULPAIE' +
         ' FROM PAIEENCOURS ' +
         'LEFT JOIN SALARIES ON (PPU_SALARIE = PSA_SALARIE) ' +
         'LEFT JOIN HISTOBULLETIN ON (PPU_ETABLISSEMENT = PHB_ETABLISSEMENT And PPU_SALARIE = PHB_SALARIE And ' +
         'PPU_DATEDEBUT = PHB_DATEDEBUT And PPU_DATEFIN = PHB_DATEFIN And PHB_NATURERUB = "AAA") ' +
         'LEFT JOIN REMUNERATION ON ##PRM_PREDEFINI## (PHB_RUBRIQUE = PRM_RUBRIQUE And PHB_NATURERUB = PRM_NATURERUB) ' +
         'LEFT JOIN CUMULRUBRIQUE ON ##PCR_PREDEFINI## (PHB_RUBRIQUE = PCR_RUBRIQUE And PHB_NATURERUB = PCR_NATURERUB) ';
    SQL := SQL + Where + Suite + ' (PPU_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatFin)+'") ';
    SQL := SQL + 'And PHB_MTREM <> 0 And PHB_IMPRIMABLE = "X" And PCR_CUMULPAIE = "01" ';

    If (Ch4.Checked = True) Then
       SQL := SQL + 'And PPU_BULCOMPL = "X" '
     Else
       SQL := SQL + 'And PPU_BULCOMPL <> "X" ';

    Orderby := 'Order by ';
    If (Ch2.Checked = True) And (Norme = True) Then
       Orderby := Orderby + 'PPU_ETABLISSEMENT,';

    If (Ch3.Checked = True) Then
       Orderby := Orderby + AjoutChamp + 'PPU_LIBELLE,PPU_PRENOM,'
     Else
       Orderby := Orderby + AjoutChamp + 'PPU_SALARIE,';

    Orderby := Orderby + 'PPU_DATEDEBUT,PPU_DATEFIN';
    SQL := SQL + Orderby ;
    Q := OpenSQL(SQL,True);
    TSQL := Tob.Create('Edit1',Nil,-1);
    TSQL.LoadDetailDB('Edit1','','',Q,False);
    Ferme(Q);
    //TOBDebug(TSQL);

    SQL:= 'SELECT PPM_TYPEPROFIL,PPM_PROFIL,PPM_RUBRIQUE,PPM_NATURERUB FROM PROFILRUB ' +
          'WHERE ##PPM_PREDEFINI## (PPM_TYPEPROFIL = "CGE" AND PPM_NATURERUB = "AAA" AND PPM_IMPRIMABLE = "X")';

    Q := OpenSQL(SQL,True);
    Tprofil := Tob.Create('P_CGE',Nil,-1);
    Tprofil.LoadDetailDB('P_CGE','','',Q,False);
    Ferme(Q);
    //TOBDebug(Tprofil);

    M1 := 0;    //SAL : Salaires
    M2 := 0;    //HEU : Heures Supp
    M3 := 0;    //ABS : Absences
    M4 := 0;    //COM + présent dans profil type CGE associé au salarié  : Conges
    M5 := 0;    //COM : Compléments
    M6 := 0;    //INI : Primes / Indemnités
    M7 := 0;    //AVT : Avantages
    M8 := 0;    //Autres avec PCR_CUMULPAIE = "01" Composant du BRUT
    DebMin := '31/12/2099';
    FinMax := '01/01/1900';

    For i := 0 To TSQL.Detail.Count -1 do
    begin
      Cfait := False;
      If TSQL.Detail[i].GetValue('PHB_SENSBUL') = 'P' Then S := 1 Else S := -1;

      If TSQL.Detail[i].GetValue('PRM_THEMEREM') = 'SAL' Then
      Begin
         M1 := M1 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'));
         Cfait := True;
      End;
      If TSQL.Detail[i].GetValue('PRM_THEMEREM') = 'HEU' Then
      Begin
         M2 := M2 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'));
         Cfait := True;
      End;
      If TSQL.Detail[i].GetValue('PRM_THEMEREM') = 'ABS' Then
      Begin
         M3 := M3 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'));
         Cfait := True;
      End;
      If TSQL.Detail[i].GetValue('PRM_THEMEREM') = 'COM' Then
      Begin
         Cprofil := False;
         For y := 0 To Tprofil.Detail.Count -1 do
         begin
            If (Tprofil.Detail[y].GetValue('PPM_PROFIL') = TSQL.Detail[i].GetValue('PSA_PROFILCGE')) And
               (Tprofil.Detail[y].GetValue('PPM_RUBRIQUE') = TSQL.Detail[i].GetValue('PHB_RUBRIQUE')) Then
            Begin
               Cprofil := True;
               Break;
            End;
         End;
         If (Cprofil = True) Then
             M4 := M4 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'))
          Else
             M5 := M5 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'));

         Cfait := True;
      End;
      If TSQL.Detail[i].GetValue('PRM_THEMEREM') = 'INI' Then
      Begin
         M6 := M6 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'));
         Cfait := True;
      End;
      If TSQL.Detail[i].GetValue('PRM_THEMEREM') = 'AVT' Then
      Begin
         M7 := M7 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'));
         Cfait := True;
      End;
      If (Cfait = False) Then M8 := M8 + (S * TSQL.Detail[i].GetValue('PHB_MTREM'));

      If IsValidDate(TSQL.Detail[i].GetValue('PPU_DATEDEBUT')) Then
         If StrToDate(DebMin) > StrToDate(TSQL.Detail[i].GetValue('PPU_DATEDEBUT')) Then
             DebMin := TSQL.Detail[i].GetValue('PPU_DATEDEBUT');

      If IsValidDate(TSQL.Detail[i].GetValue('PPU_DATEFIN')) Then
         If StrToDate(FinMax) < StrToDate(TSQL.Detail[i].GetValue('PPU_DATEFIN')) Then
             FinMax := TSQL.Detail[i].GetValue('PPU_DATEFIN');

      If (i = TSQL.Detail.Count -1) Or
         (TSQL.Detail[i].GetValue('PPU_SALARIE') <> TSQL.Detail[i+1].GetValue('PPU_SALARIE')) Or
         ((Ch1.Checked = False) And ((StrToDate(DebMin) <> StrToDate(TSQL.Detail[i+1].GetValue('PPU_DATEDEBUT'))) Or (StrToDate(FinMax) <> StrToDate(TSQL.Detail[i+1].GetValue('PPU_DATEFIN'))))) Then
      Begin
          T := Tob.Create('Fille',TobEdit,-1);
          If (Ch2.Checked = True) And (Norme = True) Then
              T.AddChampSupValeur('RUPTURE1',TSQL.Detail[i].GetValue('PPU_ETABLISSEMENT'))
           Else
              T.AddChampSupValeur('RUPTURE1','');
          If Trim(AjoutChamp) <> '' Then
              T.AddChampSupValeur('RUPTURE2',TSQL.Detail[i].GetValue(Ch_Rupt.Text))
           Else
              T.AddChampSupValeur('RUPTURE2','');
          If (Ch1.Checked = False) Then
              T.AddChampSupValeur('RUPTURE3',TSQL.Detail[i].GetValue('PPU_SALARIE'))
           Else
              T.AddChampSupValeur('RUPTURE3','');

          T.AddChampSupValeur('PPU_ETABLISSEMENT',TSQL.Detail[i].GetValue('PPU_ETABLISSEMENT'));
          T.AddChampSupValeur('PPU_SALARIE',TSQL.Detail[i].GetValue('PPU_SALARIE'));
          T.AddChampSupValeur('PPU_LIBELLE',TSQL.Detail[i].GetValue('PPU_LIBELLE'));
          T.AddChampSupValeur('PPU_PRENOM',TSQL.Detail[i].GetValue('PPU_PRENOM'));
          T.AddChampSupValeur('PPU_CIVILITE',TSQL.Detail[i].GetValue('PPU_CIVILITE'));

          If Trim(AjoutChamp) <> '' Then
              T.AddChampSupValeur(Ch_Rupt.Text,TSQL.Detail[i].GetValue(Ch_Rupt.Text));

          T.AddChampSupValeur('PPU_DATEDEBUT',DebMin);
          T.AddChampSupValeur('PPU_DATEFIN',FinMax);
          T.AddChampSupValeur('M01',M1);
          T.AddChampSupValeur('M02',M2);
          T.AddChampSupValeur('M03',M3);
          T.AddChampSupValeur('M04',M4);
          T.AddChampSupValeur('M05',M5);
          T.AddChampSupValeur('M06',M6);
          T.AddChampSupValeur('M07',M7);
          T.AddChampSupValeur('M08',M8);
          M1 := 0;
          M2 := 0;
          M3 := 0;
          M4 := 0;
          M5 := 0;
          M6 := 0;
          M7 := 0;
          M8 := 0;
          DebMin := '31/12/2099';
          FinMax := '01/01/1900';
      End;
    End;
    TSQL.Free;
    Tprofil.Free;
    //TOBDebug(TobEdit);
  End;
End;

Procedure TOF_EDIT_REMUNERATION.Edition_Pcharges(Where ,Suite : String; DatDeb ,DatFin :TDateTime);
var
  Ch1      :TCheckBox;
  Ch_Rupt  :THEdit;
  YY,MM,JJ : WORD;
  SQL ,Where2 ,AjoutChamp1 ,AjoutChamp2 ,Lmasse :String;
  i, y, x, z,Borne, Trouve :Integer;
  Q :TQuery;
  DatD, DatF : TDateTime;
  T : Tob;
  TA : array[1..3] of Tob;
  TB : array[1..3] of Tob;
  T1, T2, T3 :Double;

Begin
  Ch1 := TCheckBox(GetControl('CETAB'));
  Ch_Rupt := THEdit(GetControl('XX_RUPTURE1'));

  If (Ch1 <> nil) And (Ch_Rupt <> nil) Then
  Begin
    AjoutChamp1 :='';
    AjoutChamp2 :='';
    If (Ch1.Checked = True) And (Norme = True) Then AjoutChamp1 := 'PHB_ETABLISSEMENT,';
    If (Ch1.Checked = True) And (Norme = True) Then AjoutChamp2 := 'PPU_ETABLISSEMENT,';
    If Trim(Ch_Rupt.Text)<>'' Then AjoutChamp1 := AjoutChamp1 + ConvertPrefixe(Trim(Ch_Rupt.Text),'PPU','PHB') + ',' ;
    If Trim(Ch_Rupt.Text)<>'' Then AjoutChamp2 := AjoutChamp2 + Trim(Ch_Rupt.Text) + ',' ;
    Where2 := Trim(Where);
    If Where2 <> '' Then Where2 := ConvertPrefixe(Where2,'PPU','PHB');

    For i := 1 To 3 do
    Begin
       If i=1 Then
       Begin
          DatD := DatDeb;
          DatF := DatFin;
       End
        Else
       Begin
          DecodeDate(DatDeb,YY,MM,JJ);
          DatD := EncodeDate(YY-(i-1),MM,JJ);
          DecodeDate(DatFin,YY,MM,JJ);
          DatF := EncodeDate(YY-(i-1),MM,JJ);
       End;

       SQL :='SELECT PHB_ORGANISME,CC_LIBELLE,' + AjoutChamp1 + 'SUM(PHB_MTPATRONAL) AS MTPAT FROM HISTOBULLETIN ';
       SQL := SQL +'LEFT JOIN CHOIXCOD ON (CC_CODE = PHB_ORGANISME And CC_TYPE = "PTG") ';
       SQL := SQL + Where2 + Suite + ' (PHB_DATEDEBUT >= "'+UsDateTime(DatD)+'" AND PHB_DATEFIN <= "'+UsDateTime(DatF)+'") ';
       SQL := SQL + 'And PHB_MTPATRONAL <> 0 And PHB_IMPRIMABLE = "X" And PHB_NATURERUB = "COT"';
       SQL := SQL + ' GROUP BY '+ AjoutChamp1 + 'PHB_ORGANISME,CC_LIBELLE';
       SQL := SQL + ' Order by ' + AjoutChamp1 + 'PHB_ORGANISME,CC_LIBELLE';

       Q := OpenSQL(SQL,True);
       TA[i] := Tob.Create('C_employeur',Nil,-1);
       TA[i].LoadDetailDB('C_employeur','','',Q,False);
       Ferme(Q);

       SQL:='SELECT SUM(PPU_CBRUT) AS PPU_CBRUT,SUM(PPU_CHEURESTRAV) AS PPU_CHEURESTRAV,'+ AjoutChamp2 +
            'COUNT(PPU_DATEDEBUT) AS NBBULL' +
            ' FROM PAIEENCOURS ';
       SQL := SQL + Where + Suite + ' (PPU_DATEDEBUT >= "'+UsDateTime(DatD)+'" AND PPU_DATEFIN <= "'+UsDateTime(DatF)+'")';
       If Trim(AjoutChamp2) <> '' Then
       Begin
          SQL := SQL + ' GROUP BY '+ Copy(AjoutChamp2,1,LENGTH(AjoutChamp2)-1) ;
          SQL := SQL + ' Order by ' + Copy(AjoutChamp2,1,LENGTH(AjoutChamp2)-1);
       End;
       Q := OpenSQL(SQL,True);
       TB[i] := Tob.Create('M_salariale',Nil,-1);
       TB[i].LoadDetailDB('M_salariale','','',Q,False);
       Ferme(Q);
       //TOBDebug(TA[i]);
       //TOBDebug(TB[i]);
    End;
    For x := 1 To 3 do
    Begin
       For i := 0 To TA[x].Detail.Count -1 do
       begin
          Trouve := -1;
          If x > 1 Then
          Begin
             For y := 0 To TobEdit.Detail.Count -1 do
             begin
                If (((TobEdit.Detail[y].GetValue('RUPTURE1') = '') Or
                     (TobEdit.Detail[y].GetValue('RUPTURE1') = TA[x].Detail[i].GetValue('PHB_ETABLISSEMENT'))) And
                    ((TobEdit.Detail[y].GetValue('RUPTURE2') = '') Or
                     (TobEdit.Detail[y].GetValue('RUPTURE2') = TA[x].Detail[i].GetValue(ConvertPrefixe(Trim(Ch_Rupt.Text),'PPU','PHB')))) And
                    (TobEdit.Detail[y].GetValue('PHB_ORGANISME') = TA[x].Detail[i].GetValue('PHB_ORGANISME'))) Then
                Begin
                  Trouve := y;
                  Break;
                End;
             End;
          End;
          If Trouve = -1 Then
          Begin
             T := Tob.Create('Fille',TobEdit,-1);
             If (Ch1.Checked = True) And (Norme = True) Then
             Begin
                T.AddChampSupValeur('RUPTURE1',TA[x].Detail[i].GetValue('PHB_ETABLISSEMENT'));
                T.AddChampSupValeur('PPU_ETABLISSEMENT',TA[x].Detail[i].GetValue('PHB_ETABLISSEMENT'));
              End
              Else
                T.AddChampSupValeur('RUPTURE1','');
             If Trim(Ch_Rupt.Text) <> '' Then
             Begin
                T.AddChampSupValeur('RUPTURE2',TA[x].Detail[i].GetValue(ConvertPrefixe(Trim(Ch_Rupt.Text),'PPU','PHB')));
                T.AddChampSupValeur(Ch_Rupt.Text,TA[x].Detail[i].GetValue(ConvertPrefixe(Trim(Ch_Rupt.Text),'PPU','PHB')));
              End
              Else
                T.AddChampSupValeur('RUPTURE2','');
             T.AddChampSupValeur('RUPTURE3',1);
             T.AddChampSupValeur('PHB_ORGANISME',TA[x].Detail[i].GetValue('PHB_ORGANISME'));
             T.AddChampSupValeur('TITRE_LIGNE',TA[x].Detail[i].GetValue('CC_LIBELLE'));
             T.AddChampSupValeur('M01',0);
             T.AddChampSupValeur('T01',0);
             T.AddChampSupValeur('M02',0);
             T.AddChampSupValeur('T02',0);
             T.AddChampSupValeur('M03',0);
             T.AddChampSupValeur('T03',0);
             DatF := DatFin;
             T.AddChampSupValeur('DATE_FIN1',DateToStr(DatF));
             DecodeDate(DatFin,YY,MM,JJ);
             DatF := EncodeDate(YY-1,MM,JJ);
             T.AddChampSupValeur('DATE_FIN2',DateToStr(DatF));
             DatF := EncodeDate(YY-2,MM,JJ);
             T.AddChampSupValeur('DATE_FIN3',DateToStr(DatF));
             Trouve := TobEdit.Detail.Count -1;
          End;
          If Trouve <> -1 Then
          Begin
             If x = 1 Then TobEdit.Detail[Trouve].PutValue('M01',TA[x].Detail[i].GetValue('MTPAT'));
             If x = 2 Then TobEdit.Detail[Trouve].PutValue('M02',TA[x].Detail[i].GetValue('MTPAT'));
             If x = 3 Then TobEdit.Detail[Trouve].PutValue('M03',TA[x].Detail[i].GetValue('MTPAT'));
          End;
       End;
    End;
    For z := 1 to 4 do
    Begin
       If Z = 1 Then Lmasse := '001';
       If Z = 2 Then Lmasse := '002';
       If Z = 3 Then Lmasse := '003';
       If Z = 4 Then Lmasse := '004';

       For x := 1 To 3 do
       Begin
          For i := 0 To TB[x].Detail.Count -1 do
          begin
             Trouve := -1;
             If x > 1 Then
             Begin
                For y := 0 To TobEdit.Detail.Count -1 do
                begin
                   If ((TobEdit.Detail[y].GetValue('RUPTURE3') = 2) And
                       ((TobEdit.Detail[y].GetValue('RUPTURE1') = '') Or
                        (TobEdit.Detail[y].GetValue('RUPTURE1') = TB[x].Detail[i].GetValue('PPU_ETABLISSEMENT'))) And
                       ((TobEdit.Detail[y].GetValue('RUPTURE2') = '') Or
                        (TobEdit.Detail[y].GetValue('RUPTURE2') = TB[x].Detail[i].GetValue(Trim(Ch_Rupt.Text)))) And
                       (TobEdit.Detail[y].GetValue('PHB_ORGANISME') = Lmasse)) Then
                   Begin
                     Trouve := y;
                     Break;
                   End;
                End;
             End;
             If Trouve = -1 Then
             Begin
                T := Tob.Create('Fille',TobEdit,-1);
                If (Ch1.Checked = True) And (Norme = True) Then
                Begin
                   T.AddChampSupValeur('RUPTURE1',TB[x].Detail[i].GetValue('PPU_ETABLISSEMENT'));
                   T.AddChampSupValeur('PPU_ETABLISSEMENT',TB[x].Detail[i].GetValue('PPU_ETABLISSEMENT'));
                 End
                 Else
                   T.AddChampSupValeur('RUPTURE1','');
                If Trim(Ch_Rupt.Text) <> '' Then
                Begin
                   T.AddChampSupValeur('RUPTURE2',TB[x].Detail[i].GetValue(Trim(Ch_Rupt.Text)));
                   T.AddChampSupValeur(Ch_Rupt.Text,TB[x].Detail[i].GetValue(Trim(Ch_Rupt.Text)));
                 End
                 Else
                   T.AddChampSupValeur('RUPTURE2','');
                T.AddChampSupValeur('RUPTURE3',2);
                T.AddChampSupValeur('PHB_ORGANISME',Lmasse);
                If Z = 1 Then T.AddChampSupValeur('TITRE_LIGNE','MASSE SALARIALE');
                If Z = 2 Then T.AddChampSupValeur('TITRE_LIGNE','COUT POUR L''EMPLOYEUR');
                If Z = 3 Then T.AddChampSupValeur('TITRE_LIGNE','HEURES TRAVAILLEES');
                If Z = 4 Then T.AddChampSupValeur('TITRE_LIGNE','SALAIRE MOYEN');
                T.AddChampSupValeur('M01',0);
                T.AddChampSupValeur('T01',0);
                T.AddChampSupValeur('M02',0);
                T.AddChampSupValeur('T02',0);
                T.AddChampSupValeur('M03',0);
                T.AddChampSupValeur('T03',0);
                DatF := DatFin;
                T.AddChampSupValeur('DATE_FIN1',DateToStr(DatF));
                DecodeDate(DatFin,YY,MM,JJ);
                DatF := EncodeDate(YY-1,MM,JJ);
                T.AddChampSupValeur('DATE_FIN2',DateToStr(DatF));
                DatF := EncodeDate(YY-2,MM,JJ);
                T.AddChampSupValeur('DATE_FIN3',DateToStr(DatF));
                Trouve := TobEdit.Detail.Count -1;
             End;
             If Trouve <> -1 Then
             Begin
                If Z = 1 Then
                Begin
                   If x = 1 Then TobEdit.Detail[Trouve].PutValue('M01',TB[x].Detail[i].GetValue('PPU_CBRUT'));
                   If x = 2 Then TobEdit.Detail[Trouve].PutValue('M02',TB[x].Detail[i].GetValue('PPU_CBRUT'));
                   If x = 3 Then TobEdit.Detail[Trouve].PutValue('M03',TB[x].Detail[i].GetValue('PPU_CBRUT'));
                End;
                If Z = 3 Then
                Begin
                   If x = 1 Then TobEdit.Detail[Trouve].PutValue('M01',TB[x].Detail[i].GetValue('PPU_CHEURESTRAV'));
                   If x = 2 Then TobEdit.Detail[Trouve].PutValue('M02',TB[x].Detail[i].GetValue('PPU_CHEURESTRAV'));
                   If x = 3 Then TobEdit.Detail[Trouve].PutValue('M03',TB[x].Detail[i].GetValue('PPU_CHEURESTRAV'));
                End;
                If Z = 4 Then
                Begin
                   T1 := 0;
                   If TB[x].Detail[i].GetValue('NBBULL') <> 0 Then
                       T1 := TB[x].Detail[i].GetValue('PPU_CBRUT')/TB[x].Detail[i].GetValue('NBBULL');
                   If x = 1 Then
                   Begin
                      TobEdit.Detail[Trouve].PutValue('M01',T1);
                      TobEdit.Detail[Trouve].PutValue('T01',TB[x].Detail[i].GetValue('NBBULL'));
                   End;
                   If x = 2 Then
                   Begin
                      TobEdit.Detail[Trouve].PutValue('M02',T1);
                      TobEdit.Detail[Trouve].PutValue('T02',TB[x].Detail[i].GetValue('NBBULL'));
                   End;
                   If x = 3 Then
                   Begin
                      TobEdit.Detail[Trouve].PutValue('M03',T1);
                      TobEdit.Detail[Trouve].PutValue('T03',TB[x].Detail[i].GetValue('NBBULL'));
                   End;
                End;
             End;
          End;
       End;
    End;
    TobEdit.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;PHB_ORGANISME');
    //TOBDebug(TobEdit);
    T1 := 0;
    T2 := 0;
    T3 := 0;
    Borne := 0;
    For i := 0 To TobEdit.Detail.Count -2 do
    Begin
       T1 := T1 + TobEdit.Detail[i].GetValue('M01');
       T2 := T2 + TobEdit.Detail[i].GetValue('M02');
       T3 := T3 + TobEdit.Detail[i].GetValue('M03');

       If (TobEdit.Detail[i].GetValue('RUPTURE1') <> TobEdit.Detail[i+1].GetValue('RUPTURE1')) Or
          (TobEdit.Detail[i].GetValue('RUPTURE2') <> TobEdit.Detail[i+1].GetValue('RUPTURE2')) Or
          (TobEdit.Detail[i].GetValue('RUPTURE3') <> TobEdit.Detail[i+1].GetValue('RUPTURE3')) Then
       Begin
          For y := Borne To i do
          Begin
             TobEdit.Detail[y].PutValue('T01',T1);
             TobEdit.Detail[y].PutValue('T02',T2);
             TobEdit.Detail[y].PutValue('T03',T3);
          End;
          For y := i+1 To TobEdit.Detail.Count -1 do
          Begin
             If (TobEdit.Detail[y].GetValue('RUPTURE1') = TobEdit.Detail[i].GetValue('RUPTURE1')) And
                (TobEdit.Detail[y].GetValue('RUPTURE2') = TobEdit.Detail[i].GetValue('RUPTURE2')) And
                (TobEdit.Detail[y].GetValue('PHB_ORGANISME') = '002') Then
             Begin
                For x := 1 To 3 Do
                Begin
                   For z := 0 to TB[x].Detail.Count -1 do
                   Begin
                      If (((TobEdit.Detail[y].GetValue('RUPTURE1') = '') Or
                          (TobEdit.Detail[y].GetValue('RUPTURE1') = TB[x].Detail[z].GetValue('PPU_ETABLISSEMENT'))) And
                         ((TobEdit.Detail[y].GetValue('RUPTURE2') = '') Or
                          (TobEdit.Detail[y].GetValue('RUPTURE2') = TB[x].Detail[z].GetValue(Trim(Ch_Rupt.Text))))) Then
                      Begin
                         If x = 1 Then TobEdit.Detail[y].PutValue('M01',TB[x].Detail[z].GetValue('PPU_CBRUT') + T1);
                         If x = 2 Then TobEdit.Detail[y].PutValue('M02',TB[x].Detail[z].GetValue('PPU_CBRUT') + T2);
                         If x = 3 Then TobEdit.Detail[y].PutValue('M03',TB[x].Detail[z].GetValue('PPU_CBRUT') + T3);
                      End;
                   End;
                End;
             End;
          End;
          Borne := i+1;
          T1 := 0;
          T2 := 0;
          T3 := 0;
       End;
    End;
    For x := 1 To 3 do
    Begin
       TA[x].Free;
       TB[x].Free;
    End;
  End;
End;

Initialization
  registerclasses ( [ TOF_EDIT_REMUNERATION ] ) ;
End.
