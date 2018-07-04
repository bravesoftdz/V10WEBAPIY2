{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PRESENCESALARIE (PRESENCESALARIE)
Mots clefs ... : TOM;PRESENCESALARIE
*****************************************************************
PT1  10/08/2007  FLO  Recalcul automatique des compteurs lors d'une modification manuelle
PT2  30/08/2007  FLO  Ne pas générer de recalcul si suppression d'un compteur à recalculer
}
Unit UTOMCompteursCalc ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     sysutils,
     ComCtrls,
     HEnt1,
     UTob,
     UTOM;

Type
  TOM_PRESENCESALARIE = Class (TOM)
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ; //PT1
    procedure OnAfterDeleteRecord        ; override ; //PT1
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;

  private
    ChangeField : Boolean;
    Salarie     : String;    //PT1
    DateModif   : TDateTime; //PT1
    Recalculer  : Boolean;   //PT2
    procedure ChangeCompteur (Sender : TObject);
    procedure ControleDates;
    end ;

Implementation

Uses PGPresence, HMsgBox, HCtrls;

Const A_INTEGRER = 'AIN';
Const INTEGRE    = 'INP';
Const MANUEL     = 'NON';
Const A_RECALC   = 'ARE'; //PT3
Const SAISIE     = 'SAI';

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.OnDeleteRecord ;
begin
  Inherited ;

     // Si le compteur est à intégrer ou intégré, la suppression est interdite
     If (GetField('PYP_PGINDICATPRES') = INTEGRE) Or (GetField('PYP_PGINDICATPRES') = A_INTEGRER) Then
     Begin
          LastError := 1;
          LastErrorMsg := 'Suppression impossible. Ce compteur est utilisé en paie.';
     End;

     //PT1 - Début
     If (LastError = 0) Then
     Begin
          Recalculer  := Not (GetField('PYP_PGINDICATPRES') = A_RECALC); //PT3
          Salarie     := GetField('PYP_SALARIE');
          DateModif   := GetField('PYP_DATEDEBUTPRES');
     End;
     //PT1 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Création ou modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.OnUpdateRecord ;
Var Q : TQuery;
    DD : TDateTime;
    Champs : String;
begin
  Inherited ;

     // En création, il faut récupérer certaines valeurs de la table SALARIES et COMPTEURPRESENCE
     If DS.State in [dsInsert] Then
     Begin
          { Vérification de la saisie des champs. Du fait que les dates sont renseignées par défaut
            au 01/01/1900, elles sont considérées comme OK. Du coup, on est obligé de les contrôler.
            Et pour plus de lisibilité, on refait tout. }
          If (GetControlText('PYP_SALARIE') = '') Then
          Begin
               LastError := 1;
               Champs := '#13#10- '+TraduireMemoire('le salarié');
          End;

          DD := StrToDate(GetControlText ('PYP_DATEDEBUTPRES'));
          If (DD=0) Or (DD = iDate1900) Then
          Begin
               LastError := 1;
               Champs := Champs + '#13#10- '+TraduireMemoire('la date de début');
          End;

          DD := StrToDate(GetControlText ('PYP_DATEFINPRES'));
          If (DD=0) Or (DD = iDate1900) Then
          Begin
               LastError := 1;
               Champs := Champs + '#13#10- '+TraduireMemoire('la date de fin');
          End;

          If (GetControlText('PYP_COMPTEURPRES') = '') Then
          Begin
               LastError := 1;
               Champs := Champs + '#13#10- '+TraduireMemoire('le compteur associé');
          End;

          If LastError=1 Then
          Begin
               PGIBox(Format(TraduireMemoire('Vous devez renseigner : %s'),[Champs]),TraduireMemoire('Contrôle de saisie'));
               LastErrorMsg := '';
          End;

          ControleDates;
          If LastError <> 0 Then Exit;

          // Forçage du compteur car il n'est pas récupéré par l'AGL
          SetField('PYP_COMPTEURPRES', GetControlText('PYP_COMPTEURPRES'));

          // Recherche des informations du salarié
          Q := OpenSQL ('SELECT * FROM SALARIES WHERE PSA_SALARIE="'+GetControlText('PYP_SALARIE')+'"', True);
          If Not Q.EOF Then
          Begin
               SetField('PYP_ETABLISSEMENT', Q.FindField('PSA_ETABLISSEMENT').AsString);
               SetField('PYP_TRAVAILN1',     Q.FindField('PSA_TRAVAILN1').AsString);
               SetField('PYP_TRAVAILN2',     Q.FindField('PSA_TRAVAILN2').AsString);
               SetField('PYP_TRAVAILN3',     Q.FindField('PSA_TRAVAILN3').AsString);
               SetField('PYP_TRAVAILN4',     Q.FindField('PSA_TRAVAILN4').AsString);
               SetField('PYP_CODESTAT',      Q.FindField('PSA_CODESTAT').AsString);
               SetField('PYP_CONVENTION',    Q.FindField('PSA_CONVENTION').AsString);
               SetField('PYP_DADSPROF',      Q.FindField('PSA_DADSPROF').AsString);
               SetField('PYP_DADSCAT',       Q.FindField('PSA_DADSCAT').AsString);
               SetField('PYP_LIBREPCMB1',    Q.FindField('PSA_LIBREPCMB1').AsString);
               SetField('PYP_LIBREPCMB2',    Q.FindField('PSA_LIBREPCMB2').AsString);
               SetField('PYP_LIBREPCMB3',    Q.FindField('PSA_LIBREPCMB3').AsString);
               SetField('PYP_LIBREPCMB4',    Q.FindField('PSA_LIBREPCMB4').AsString);
          End;
          Ferme(Q);
          If LastError <> 0 Then Exit;

          // Recherche des informations du compteur
          Q := OpenSQL ('SELECT PYR_PERIODICITEPRE, PYR_THEMEPRE FROM COMPTEURPRESENCE WHERE PYR_COMPTEURPRES="'+GetControlText('PYP_COMPTEURPRES')+'" AND PYR_DATEVALIDITE="'+UsDateTime(StrToDate(GetControlText('PYR_DATEVALIDITE')))+'"', True);
          If Not Q.EOF Then
          Begin
               SetField('PYP_THEMEPRE',       Q.FindField('PYR_THEMEPRE').AsString);
               SetField('PYP_PERIODICITEPRE', Q.FindField('PYR_PERIODICITEPRE').AsString);
          End;
          Ferme(Q);
          If LastError <> 0 Then Exit;

          // Forçage des valeurs de l'état, de l'indicateur et du type
          SetField ('PYP_TYPECALPRES',   '001');
          SetField ('PYP_PGINDICATPRES', MANUEL);
          SetField ('PYP_ETATPRES',      SAISIE);
     End
     Else If DS.State In [dsEdit] Then
     Begin
          // En modification, on change l'état et l'indicateur (à n'effectuer que si la modification est possible) 
          If GetControlEnabled('PYP_QUANTITEPRES') Then
          Begin
               SetField('PYP_PGINDICATPRES',MANUEL);
               SetField('PYP_ETATPRES',     SAISIE);
          End;
     End;

     //PT1 - Début
     If (LastError = 0) Then
     Begin
          Recalculer  := Not (GetField('PYP_PGINDICATPRES') = A_RECALC); //PT3
          Salarie     := GetField('PYP_SALARIE');
          DateModif   := GetField('PYP_DATEDEBUTPRES');
     End;
     //PT1 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007 / PT1
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.OnAfterDeleteRecord;
begin
  inherited;
     If (LastError = 0) And Recalculer Then CompteursARecalculer(DateModif, Salarie);  //PT3
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007 / PT1
Modifié le ... :   /  /
Description .. : Suite à la création ou modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.OnAfterUpdateRecord ;
begin
  Inherited ;
     If (LastError = 0) And Recalculer Then CompteursARecalculer(DateModif, Salarie);  //PT3
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Chargement des données
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.OnLoadRecord ;
Var
     Combo  : THValComboBox;
begin
  Inherited ;

     // Si le compteur est à intégrer, ou intégré, on ne peut ni le modifier, ni le supprimer
     If (DS.State In [dsEdit,dsBrowse]) And ((GetField('PYP_PGINDICATPRES')=INTEGRE) Or (GetField('PYP_PGINDICATPRES')=A_INTEGRER)) Then
     Begin
          SetControlEnabled ('PYP_QUANTITEPRES', False);
          SetControlEnabled ('BDelete', False);
          SetControlEnabled ('BDefaire', False);
     End;

     // Au chargement de la fiche, il faut renseigner le compteur
     If (DS.State In [dsEdit,dsBrowse]) Then
     Begin
          Combo := THValComboBox(GetControl('COMPTEUR'));
          If Combo <> Nil Then
          Begin
               Combo.Items.Add(GetControlText('PYR_LIBELLE'));
               Combo.Values.Add(GetControlText('PYP_COMPTEURPRES')+';'+GetControlText('PYR_DATEVALIDITE'));
               // Sélection de l'élément
               SetControlText('COMPTEUR', GetControlText('PYP_COMPTEURPRES')+';'+GetControlText('PYR_DATEVALIDITE'));
          End;
          SetControlEnabled ('COMPTEUR', False);
     End;

     Combo := THValComboBox(GetControl('COMPTEUR'));
     If Combo <> Nil Then Combo.OnChange := ChangeCompteur;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Modification d'un champ
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.OnChangeField ( F: TField ) ;
Var
     Q      : TQuery;
     Combo  : THValComboBox;
     AncCpt,Temp : String;
begin
  Inherited ;

     If ChangeField = False Then Exit;

     If Ds.State In [DsInsert] Then
     Begin
          // Actualisation de la combo des compteurs en fonction de la date de début saisie
          If (F.FieldName = 'PYP_DATEDEBUTPRES') And (StrToDate(GetControlText('PYP_DATEDEBUTPRES')) <> iDate1900) Then
          Begin
               // Récupération de la combo
               Combo := THValComboBox(GetControl('COMPTEUR'));

               If Combo <> Nil Then
               Begin
                    ChangeField := False;

                    // Réinitialisation de la combo pour ne pas avoir de problème
                    Combo.Items.Clear; Combo.Values.Clear;
                    SetControlText('PYP_COMPTEURPRES', '');
                    SetControlText('PYR_DATEVALIDITE', '');
                    SetControlText('PYR_LIBELLE', '');

                    AncCpt := '';
                    Q := OpenSQL ('SELECT PYR_COMPTEURPRES, PYR_LIBELLE, MAX(PYR_DATEVALIDITE) AS DATEVALIDITE FROM COMPTEURPRESENCE '+
                    ' WHERE ##PYR_PREDEFINI## AND (PYR_DATEVALIDITE <= "'+UsDateTime(StrToDate(GetControlText('PYP_DATEDEBUTPRES')))+'")' +
                    ' GROUP BY PYR_COMPTEURPRES, PYR_DATEVALIDITE, PYR_LIBELLE ORDER BY PYR_COMPTEURPRES, PYR_DATEVALIDITE DESC', True);
                    While Not Q.Eof Do
                    Begin
                         If AncCpt <> Q.FindField('PYR_COMPTEURPRES').AsString Then
                         Begin
                              Temp := Q.FindField('PYR_COMPTEURPRES').AsString;
                              Temp := Temp + ';';
                              Temp := Temp + DateToStr(Q.FindField('DATEVALIDITE').AsDateTime);
                              Combo.Values.Add(Temp);
                              Combo.Items.Add(Q.FindField('PYR_LIBELLE').AsString);

                              AncCpt := Q.FindField('PYR_COMPTEURPRES').AsString;
                         End;
                         Q.Next;
                    End;
                    Ferme(Q);

                    ChangeField := True;
               End;
          End;

          // Contrôle de cohérence des dates
          If (F.FieldName = 'PYP_DATEFINPRES') Or (F.FieldName = 'PYP_DATEDEBUTPRES') Then
               ControleDates;
     End;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Chargement de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.OnArgument ( S: String ) ;
Var
     Action, DateVal, Libelle, Compteur : String;
begin
  Inherited ;

     ChangeField := True;

     // Récupération des paramètres en entrée
     Action  := ReadTokenSt(S);
     Compteur:= ReadTokenSt(S);
     DateVal := ReadTokenSt(S);
     Libelle := ReadTokenSt(S);

     SetControlText('PYP_COMPTEURPRES', Compteur);
     SetControlText('PYR_DATEVALIDITE', DateVal);
     SetControlText('PYR_LIBELLE', Libelle);

     {$IFDEF EAGLCLIENT}
          SetControlText('PYP_DATEDEBUTPRES', DateToStr(iDate1900));
          SetControlText('PYP_DATEFINPRES',   DateToStr(iDate1900));
     {$ENDIF}
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Contrôle la validité des dates saisies
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.ControleDates;
Var 
    DD,DF : TDateTime;
begin
     // Contrôle des dates de début et de fin
     DD := StrToDate(GetControlText('PYP_DATEDEBUTPRES'));
     DF := StrToDate(GetControlText('PYP_DATEFINPRES'));
     If (DD <> iDate1900) And (DF <> iDate1900) And (DD > DF) Then
     Begin
          LastError := 1;
          LastErrorMsg := TraduireMemoire('La date de début doit être inférieure ou égale à la date de fin');
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Actualisation de champs sur changement de compteur
Mots clefs ... :
*****************************************************************}
procedure TOM_PRESENCESALARIE.ChangeCompteur(Sender: TObject);
Var
     Temp : String;
begin
     // Actualisation du libellé du compteur
     SetControlText ('PYR_LIBELLE', THValComboBox(Sender).Text);

     // Mise à jour de la référence du compteur dans les champs cachés
     Temp := THValComboBox(Sender).Value;
     SetControlText('PYP_COMPTEURPRES',ReadTokenSt(Temp));
     SetControlText('PYR_DATEVALIDITE',ReadTokenSt(Temp));
end;

Initialization
  registerclasses ( [ TOM_PRESENCESALARIE ] ) ; 
end.

