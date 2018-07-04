{***********UNITE*************************************************************
Auteur  ...... : NA
Créé le ...... : 05/03/2007
Modifié le ... : 13/07/2007 FLO
Description .. : Source TOM de la TABLE : COMPTEURPRESENCE (COMPTEURPRESENCE)
Mots clefs ... : TOM;COMPTEURPRESENCE
******************************************************************************
PT1  11/07/2007  FLO         Ajout de contrôles de saisie
PT2  09/08/2007  GGU/FLO     Recalcul automatique des compteurs lors d'une modification ou d'une suppression de compteur.
PT3  24/07/2007  FLO         Permettre l'intégration en paie également pour les compteurs annuels + contrôles divers +
                             association compteurs/profils
PT4  30/08/2007  FLO         Blocage de la périodicité en cas de duplication
}
Unit UTOMCOMPTEURPRESENCE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     Dialogs,
     Graphics,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     HDB,
     FichList,
     Fe_Main,
{$ELSE}
     MaineAgl,
{$ENDIF}
     UTob,
     UTOM;

Type
  TOM_COMPTEURPRESENCE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;  //PT3
    procedure OnArgument ( S: String )   ; override ;
    procedure OnAfterUpdateRecord        ; override ; //PT2
    procedure OnAfterDeleteRecord        ; override ; //PT2

    private
    Compteur, paramcompteur, periodicitevar, libellecpteur, variablepres, originecreation: String; //PT2
    DateModif   : TDateTime; //PT2
    VerifImpact : Boolean;   //PT2
    VerifProfilsAssoc ,LectureSeule : Boolean; //PT3
    CEG, STD, DOS, Duplic, suppressprofil : boolean;
    procedure ClickPlanning  (Sender : TObject);
    procedure ClickPgColor   (Sender : TObject);
    procedure ClickPgDefaire (Sender : TObject);
    procedure ClickDupliquer (Sender : TObject);
    procedure ClickIntegrePaie (Sender : TObject); //PT3
    procedure ClickGestionAssociee (Sender : TObject); //PT3
    procedure OnChangeCompteur (Sender : TObject);
    procedure OnSelectionPeriode (Sender : TObject);
    procedure Variablepresence   (Sender : TObject);
    procedure VerifieCompteur;
    procedure VerifiePeriodicite;
    procedure GerePredefini;
    procedure MAJCouleur;
    end ;

Implementation

Uses HCtrls,
     HEnt1,
     HMsgBox,
     sysutils,
     HTB97,
     Menus,
     PGOutils,
     PGOutils2,
     Pgpresence;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/03/2007
Modifié le ... : 13/07/2007 / PT1
Description .. : Création d'un nouvel enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnNewRecord ;
begin
  Inherited ;
     SetField('PYR_PGCOLORPRE','$00804000');
     periodicitevar := '';
     if paramcompteur = '' then
     begin
        If (CEG = TRUE) then
          SetField('PYR_PREDEFINI', 'CEG')
        Else
          SetField('PYR_PREDEFINI', 'DOS');
     end
     else // initialise les données du compteur à partir des valeurs de la variable de présence associée
     begin
     setfield('PYR_PREDEFINI', readtokenst(paramcompteur));
     setfield('PYR_LIBELLE', readtokenst(paramcompteur));
     setfield('PYR_THEMEPRE', readtokenst(paramcompteur));
     setfield('PYR_PERIODICITEPRE', readtokenst(paramcompteur));
     periodicitevar := getfield('PYR_PERIODICITEPRE');
     setfield('PYR_VARIABLEPRES', readtokenst(paramcompteur));
     end;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 13/07/2007
Modifié le ... :   /  /
Description .. : Suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnDeleteRecord;
var
Q : Tquery;
begin
  inherited;
      Compteur    := GetField('PYR_COMPTEURPRES');
      DateModif   := GetField('PYR_DATEVALIDITE');
      libellecpteur := getfield('PYR_LIBELLE');
      suppressprofil := false;

     // Vérification que le compteur n'est pas utilisé. Si c'est le cas, on interdit la suppression
     If ExisteSQL ('SELECT 1 FROM PGPRESENCESAL WHERE PYP_COMPTEURPRES="'+ compteur +'" AND PYP_DATEDEBUTPRES'+
     ' >= "'+USDATETIME(Datemodif)+'"')  Then
     Begin
          LastError := 1;
          PGIBox (TraduireMemoire('Ce compteur est utilisé.'), TraduireMemoire('Suppression impossible'));
     End;

     //PT3 - Début
     // si le compteur n'a pas d'autres dates de validité et
     // Si le compteur est associé à un ou plusieurs profils, on demande confirmation à l'utilisateur
     Q:= Opensql('SELECT COUNT(PYR_COMPTEURPRES) AS NB FROM COMPTEURPRESENCE WHERE PYR_COMPTEURPRES ='+
     '"'+ compteur +'"', true);
     if not Q.EOF then
     begin
       if Q.Findfield('NB').asinteger = 1 then
       begin
          If ExisteSQL ('SELECT 1 FROM PROFILCOMPTEURPRES WHERE PPV_COMPTEURPRES="'+ compteur +'"') Then
          Begin
              if PGIAsk(TraduireMemoire('Attention : Le compteur est utilisé par un ou plusieurs profil(s). Voulez-vous continuer?')) = mrNo Then
              LastError := 1  else   Suppressprofil := true;

          End;
        end;
     End;
     Ferme(Q);

     //PT3 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/03/2007
Modifié le ... : 13/07/2007 / PT1
Description .. : Mise à jour de l'enregistrement en base
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnUpdateRecord ;
Var
     LesChamps : String;
begin
  Inherited ;
     LesChamps := '';
     LastError := 0;

     // Contrôle de la clé
     If (GetControlText('PYR_PREDEFINI') = '') Then
     Begin
          LastError := 1;
          LesChamps := LesChamps + '#13#10- Le type prédéfini';
     End
     Else
     Begin
          If GetControlText('PYR_PREDEFINI') = 'DOS' Then
               SetField('PYR_NODOSSIER', PgrendNodossier())
          Else
               SetField('PYR_NODOSSIER', '000000');
     End;

     If (Trim(GetControlText('PYR_COMPTEURPRES')) = '') Then
     Begin
          LastError := 1;
          LesChamps := LesChamps + '#13#10- Le numéro de compteur';
     End;
     If (GetControlText('PYR_DATEVALIDITE') = '') Or (StrToDate(GetControlText('PYR_DATEVALIDITE')) = iDate1900) Then
     Begin
          LastError := 1;
          LesChamps := LesChamps + '#13#10- La date de validité';
     End;

     // Contrôle du compteur
     If (GetControlText('PYR_LIBELLE') = '') Then
     Begin
          LastError := 1;
          LesChamps := LesChamps + '#13#10- Le libellé';
     End;
     If (GetControlText('PYR_THEMEPRE') = '') Then
     Begin
          LastError := 1;
          LesChamps := LesChamps + '#13#10- Le thème';
     End;
     If (GetControlText('PYR_PERIODICITEPRE') = '') Then
     Begin
          LastError := 1;
          LesChamps := LesChamps + '#13#10- La périodicité';
     End;
     If (GetControlText('PYR_VARIABLEPRES') = '') Then
     Begin
          LastError := 1;
          LesChamps := LesChamps + '#13#10- La variable de présence associée';
     End;

     // Contrôle que si le compteur doit être intégré en paie, alors la rubrique et le type de champ soient renseignés
     {$IFDEF EAGLCLIENT}
     If (GetControl('PYR_INTEGREPAIE') As THCheckBox).Checked Then
     {$ELSE}
     If (GetControl('PYR_INTEGREPAIE') As THDBCheckBox).Checked Then
     {$ENDIF}
     Begin
          If (GetControlText('PYR_RUBRIQUE') = '') Or (GetControlText('PYR_TYPECHAMPRUB') = '') Then
          Begin
               LastError := 1;
               LesChamps := LesChamps + '#13#10- La rubrique de paie et le type de champ';
          End;
     End;

     If LastError = 1 Then
     Begin
          PGIInfo(Format(TraduireMemoire('Veuillez renseigner le(s) champ(s) suivant(s) : %s'),[LesChamps]));
          Exit;
     End;

     if originecreation = 'VAR'  then  // si création à partir d'une variable de présence , vérifié que le compteur n'existe pas déjà
     begin
         If existeSql('SELECT PYR_COMPTEURPRES FROM COMPTEURPRESENCE WHERE ##PYR_PREDEFINI## AND PYR_COMPTEURPRES'+
         ' = "'+(GetControlText('PYR_COMPTEURPRES'))+'"') Then
         begin
           PGIError (TraduireMemoire('Création impossible. Ce compteur existe déjà.'));
           LastError := 1;
         end;
     end;


     // Contrôle du numéro de compteur (dans le cas où il aurait été saisi avant le type de dossier
     VerifieCompteur;

     // Vérification de la cohérence de périodicité entre celle saisie et celle de la variable
     If LastError = 0 Then VerifiePeriodicite;

     //PT3 - Début
     // Vérifie l'association aux profils
     If (VerifProfilsAssoc) And (ds.State = dsInsert)  Then
          If Not ExisteSQL ('SELECT 1 FROM PROFILCOMPTEURPRES WHERE PPV_COMPTEURPRES="'+GetField('PYR_COMPTEURPRES')+'"') Then
               If PGIAsk(TraduireMemoire('Le compteur n''est associé à aucun profil. Voulez-vous continuer?')) = MrNo Then
                    LastError := 1;
     //PT3 - Fin

     //PT2 - Début
     If (LastError = 0) Then
     Begin
          VerifImpact := (ds.State = dsEdit) or ((ds.state = dsInsert) and duplic);
          Compteur    := GetField('PYR_COMPTEURPRES');
          DateModif   := GetField('PYR_DATEVALIDITE');
     End;
     //PT2 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnAfterDeleteRecord;
var
libellemodif,libellemodif2, daterecalcul, paramrecalcul : string;
begin
  inherited;
     If (LastError = 0) And (PresenceMeterIsUsed(Compteur)) Then
     begin
         Daterecalcul := datetostr(datemodif);
         libellemodif := 'Le compteur de présence ' + compteur + ' ' + libellecpteur;
         libellemodif2 := 'a été supprimé';
         paramrecalcul := libellemodif + ';' + libellemodif2 + ';' + daterecalcul;
         AglLanceFiche('PAY','PRESRECALCUL','','',paramrecalcul);
      //    CompteursARecalculer(DateModif);
     end;
     // Suppression des références au compteur au niveau des profils
     if suppressprofil then
     ExecuteSQL('DELETE FROM PROFILCOMPTEURPRES WHERE PPV_COMPTEURPRES="'+Compteur+'"');  //PT3
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la création ou modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnAfterUpdateRecord ;
var
libellemodif,libellemodif2, daterecalcul, paramrecalcul : string;
modifvariable : boolean;
begin
  Inherited ;
   Daterecalcul := datetostr(datemodif);
   if (variablepres <> getfield('PYR_VARIABLEPRES')) then modifvariable := true else
   modifvariable := false;
     If (LastError = 0) And VerifImpact And Modifvariable
     then
      begin
          if  (PresenceMeterIsUsed(Compteur)) Then
          begin
            libellemodif := 'Le compteur de présence ' + getfield('PYR_COMPTEURPRES') + ' ' +
            getfield('PYR_LIBELLE');
            libellemodif2 := 'a été modifié';
            paramrecalcul := libellemodif + ';' + libellemodif2 + ';' + daterecalcul;
            AglLanceFiche('PAY','PRESRECALCUL','','',paramrecalcul);
            //  CompteursARecalculer(DateModif);
          end;
      end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/03/2007
Modifié le ... : 13/07/2007 / PT1
Description .. : Chargement 
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnLoadRecord ;

begin
  Inherited ;

     MAJCouleur;
     if ds.state <> dsinsert then 
     periodicitevar := readtokenst(paramcompteur);

     // Sur NewRecord, on désactive le bouton, sinon on active
     If AfterInserting Then
          SetControlEnabled('BDUPLIQUER', False)
     Else
          SetControlEnabled('BDUPLIQUER', True);

     // Vérification des droits de l'utilisateur par rapport au prédéfini
     GerePredefini;
     if ds.state <> dsinsert then
     begin
        if not lectureseule then
        begin  // compteur , prédéfini et date de validité non modifiables
           SetControlEnabled('PYR_PREDEFINI', False);
           SetControlEnabled('PYR_COMPTEURPRES', False);
           SetControlEnabled('PYR_DATEVALIDITE', False);
        end;
     end;

     If ds.State = dsInsert Then SetControlEnabled('BDelete', False);

     // Blocage de la périodicité si le compteur est utilisé
     If ds.State = dsBrowse Then
          If ExisteSQL ('SELECT 1 FROM PGPRESENCESAL WHERE PYP_COMPTEURPRES="'+GetField('PYR_COMPTEURPRES')+'" AND PYR_LIBELLE="'+GetField('PYR_LIBELLE')+'"') Then
               SetControlEnabled('PYR_PERIODICITEPRE', False);

     variablepres := Getfield('PYR_VARIABLEPRES');   // PT2
     OnSelectionPeriode(Nil); //PT3
     ClickIntegrePaie(Nil); //PT3
end ;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 27/08/2007 / PT3
Modifié le ... :   /  /
Description .. : Change field
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnChangeField ( F: TField ) ;
var S,Periodicite,Predefini :String;
begin
  Inherited ;

     If (F.FieldName = 'PYR_PREDEFINI') Or (F.FieldName = 'PYR_PERIODICITEPRE') Then
     Begin
          Predefini := GetField('PYR_PREDEFINI');
          If Predefini = 'CEG' Then
               S := ' AND PVA_PREDEFINI="CEG"'
          Else If Predefini = 'STD' Then
               S := ' AND (PVA_PREDEFINI="CEG" OR PVA_PREDEFINI="STD")'
          Else
               S := '';

          Periodicite := GetField('PYR_PERIODICITEPRE');
          If (Periodicite = periodiciteHebdomadaire) Or (Periodicite = periodiciteFinDeCycle) Then  // hebdo ou fin cycle
               S := S + ' AND (PVA_VARPERIODICITE="003" OR PVA_VARPERIODICITE="005")'
          Else If (Periodicite = periodiciteMensuelle) Or (Periodicite = periodiciteJournaliere) Then  // mensuelle et journalier
               S := S + ' AND PVA_VARPERIODICITE="'+Periodicite+'"';

          SetControlProperty('PYR_VARIABLEPRES', 'Plus', S);
     End;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 14/06/2007
Modifié le ... : 13/07/2007 / PT1
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnArgument ( S: String ) ;
var
Btn  : TToolBarButton97;
Edit : THEdit;
mode : string;
{$IFDEF EAGLCLIENT}
Ck   : THCheckBox;
Combo: THValComboBox;
{$ELSE}
Ck   : THDBCheckBox;
Combo: THDBValComboBox;
{$ENDIF}
MenuPop : TPopUpMenu;
i : Integer;
begin
  Inherited ;
     Duplic := false;
   // Lecture des arguments pour rechercher les valeurs par défaut
     mode := readtokenst(S);

     paramcompteur:= S;

     if paramcompteur <> '' then originecreation := 'VAR'  else originecreation := 'CRT';

     Btn := TToolBarButton97(GetControl('BTNPGCOLORPRE'));
     if Assigned(Btn) then Btn.OnClick := ClickPgColor;

     Btn := TToolBarButton97(GetControl('BTNDEFAIREPRE'));
     if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;

     Btn := TToolBarButton97(GetControl('B_VARIABLE'));
     if Assigned(Btn) then Btn.OnClick := Variablepresence;

     MenuPop := TPopUpMenu(GetControl('DUPLIQUEPOPUP'));
     If MenuPop <> Nil then
          For i := 0 to MenuPop.Items.Count - 1 do
               MenuPop.Items[i].OnClick := ClickDupliquer;

     {$IFDEF EAGLCLIENT}
     Ck := THCheckBox (GetControl('PYR_EDITPLANPRES'));
     {$ELSE}
     Ck := THDBCheckBox (GetControl('PYR_EDITPLANPRES'));
     {$ENDIF}
     If Assigned(Ck) Then Ck.OnClick := ClickPlanning;

     Edit := THEdit (GetControl('PYR_COMPTEURPRES'));
     If Assigned(Edit) Then Edit.OnChange := OnChangeCompteur;

     {$IFDEF EAGLCLIENT}
     Combo := THValComboBox(GetControl('PYR_PERIODICITEPRE'));
     {$ELSE}
     Combo := THDBValComboBox(GetControl('PYR_PERIODICITEPRE'));
     {$ENDIF}
     If Combo <> Nil Then Combo.OnChange := OnSelectionPeriode;

     {$IFDEF EAGLCLIENT}
     Ck := THCheckBox (GetControl('PYR_INTEGREPAIE'));
     {$ELSE}
     Ck := THDBCheckBox (GetControl('PYR_INTEGREPAIE'));
     {$ENDIF}
     If Assigned(Ck) Then Ck.OnClick := ClickIntegrePaie;

     AccesPredefini('TOUS', CEG, STD, DOS);

     //PT3 - Début
     Btn := TToolBarButton97(GetControl('BGESTIONASS'));
     if Assigned(Btn) then Btn.OnClick := ClickGestionAssociee;

     VerifProfilsAssoc := True;
     //PT3 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/03/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton de couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.ClickPgColor(Sender: TObject);
Var
  Col : String;
  Colors : TColorDialog;
  EditCol : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
     Colors:=TColorDialog.Create(Ecran);
     if Colors=nil then exit;

     if (Colors.Execute) then
     Begin
          Col:=ColorToString(Colors.Color);
          EditCol := THEdit(GetControl('PGCOLORPRE'));
          {$IFDEF EAGLCLIENT}
          Edit := THEdit(GetControl('PYR_PGCOLORPRE'));
          {$ELSE}
          Edit := THDBEdit(GetControl('PYR_PGCOLORPRE'));
          {$ENDIF}
          If Assigned(Edit) then
          Begin
               Edit.Font.Color := Colors.Color;
               if DS.State in [DsBrowse] then Ds.Edit;
               if GetField('PYR_PGCOLORPRE') <> Col then
               Begin
                    SetField('PYR_PGCOLORPRE',Col);
                    if Assigned(EditCol) then EditCol.Color := Colors.Color;
               End;
          End;
     End;
     Colors.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/03/2007
Modifié le ... :   /  /    
Description .. : Bouton supprimer choix de la couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.ClickPgDefaire(Sender: TObject);
Var
  stVal : String;
  EditCol : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
     stVal := '$00804000';
     
     EditCol := THEdit(GetControl('PGCOLORPRE'));
     {$IFDEF EAGLCLIENT}
     Edit := THEdit(GetControl('PYR_PGCOLORPRE'));
     {$ELSE}
     Edit := THDBEdit(GetControl('PYR_PGCOLORPRE'));
     {$ENDIF}

     If Assigned(Edit) then
     Begin
          Edit.Font.Color := StringToColor(StVal);
          if DS.State in [DsBrowse] then Ds.Edit;
          if GetField('PYR_PGCOLORPRE') <> StVal then
          Begin
               SetField('PYR_PGCOLORPRE',StVal);
               if Assigned(EditCol) then EditCol.Color := StringToColor(StVal);
          End;
     End;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 11/07/2007
Modifié le ... :   /  /    
Description .. : Rend accessible ou non la sélection d'une couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.ClickPlanning(Sender: TObject);
begin
     {$IFDEF EAGLCLIENT}
     If (GetControl('PYR_EDITPLANPRES') As THCheckBox).Checked Then
     {$ELSE}
     If (GetControl('PYR_EDITPLANPRES') As THDBCheckBox).Checked Then
     {$ENDIF}
     Begin
          SetControlEnabled('TPYR_PGCOLORPRE', True);
          SetControlEnabled('BTNPGCOLORPRE', True);
          SetControlEnabled('BTNDEFAIREPRE', True);
     End
     Else
     Begin
          SetControlEnabled('TPYR_PGCOLORPRE', False);
          SetControlEnabled('BTNPGCOLORPRE', False);
          SetControlEnabled('BTNDEFAIREPRE', False);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 12/07/2007
Modifié le ... :   /  /    
Description .. : Contrôle du numéro de compteur en fonction du prédéfini
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.VerifieCompteur;
var
  Compteur, Predefini : String;
begin
     Compteur  := GetControlText('PYR_COMPTEURPRES');
     Predefini := GetControlText('PYR_PREDEFINI');

     // Contrôle de la longueur
     If Length(Trim(Compteur)) <> 3 Then
     Begin
          PGIBox(TraduireMemoire('Le numéro de compteur doit comporter 3 caractères.'), Ecran.Caption);
          SetFocusControl('PYR_COMPTEURPRES');
          Lasterror := 1;
          Exit;
     End;

     If  ((Predefini = 'STD') And (Compteur[3] <> '1') And (Compteur[3] <> '3')) Then
     Begin
          PGIBox (TraduireMemoire('Le numéro du compteur doit se terminer par 1 ou 3.'), Ecran.Caption);
          SetFocusControl('PYR_COMPTEURPRES');
          Lasterror := 1;
          Exit;
     End;

     If ((Predefini = 'CEG') And (Compteur[3] <> '0') And (Compteur[3] <> '2') And (Compteur[3] <> '4') And
        (Compteur[3] <> '6') And (Compteur[3] <> '8')) Then
     Begin
          PGIBox (TraduireMemoire('Le numéro du compteur doit se terminer par 0,2,4,6 ou 8.'), Ecran.Caption);
          SetFocusControl('PYR_COMPTEURPRES');
          Lasterror := 1;
          Exit;
     End;

     If ((Predefini = 'DOS') And (Compteur[3] <> '5') And (Compteur[3] <> '7') And (Compteur[3] <> '9')) Then
     Begin
          PGIBox (TraduireMemoire('Le numéro du compteur doit se terminer par 5,7 ou 9.'), Ecran.Caption);
          SetFocusControl('PYR_COMPTEURPRES');
          Lasterror := 1;
          Exit;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 12/07/2007
Modifié le ... :   /  /
Description .. : Modification à la volée du masque de saisie du numéro de 
Suite ........ : compteur qui peut être soit de la forme AA0 ou 000
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnChangeCompteur(Sender: TObject);
var Compteur : String;
begin
     Compteur := GetControlText('PYR_COMPTEURPRES');

     If (Length(Trim(Compteur)) > 0) And (Compteur[1] <> ' ') Then
     Begin
          If UpCase(Compteur[1]) in ['A'..'Z'] Then
               THEdit(GetControl('PYR_COMPTEURPRES')).EditMask := 'LL0'
          Else
               THEdit(GetControl('PYR_COMPTEURPRES')).EditMask := '000';
     End
     Else
          THEdit(GetControl('PYR_COMPTEURPRES')).EditMask := 'AAA'
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 13/07/2007
Modifié le ... : 24/07/2007 / PT3
Description .. : A la sélection d'un période, si c'est la Mensuelle ou Annuelle qui a été
Suite ........ : choisie, on affiche le GroupBox, sinon on le cache et on
Suite ........ : vide les champs qui en dépendent
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.OnSelectionPeriode(Sender: TObject);
begin
    If (THValComboBox(GetControl('PYR_PERIODICITEPRE')).Value <> periodiciteAnnuelle) And   //PT3
       (THValComboBox(GetControl('PYR_PERIODICITEPRE')).Value <> periodiciteMensuelle) then 
    Begin
          TGroupBox(Getcontrol('INTEGRATIONPAIE')).Visible := False;
          SetField('PYR_INTEGREPAIE', '-');
          SetField('PYR_RUBRIQUE', '');
          SetField('PYR_TYPECHAMPRUB', '');
    End
    Else
          TGroupBox(Getcontrol('INTEGRATIONPAIE')).Visible := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 13/07/2007
Modifié le ... :   /  /    
Description .. : Gestion des autorisations du prédéfini et si pas de droit :
Suite ........ : - Blocage de l'écran
Suite ........ : - Restriction des choix dans la combo
Mots clefs ... : 
*****************************************************************}
Procedure TOM_COMPTEURPRESENCE.GerePredefini;
Var Predefini    : String;
    i            : Integer;
    {$IFDEF EAGLCLIENT}
    Combo        : THValComboBox;
    {$ELSE}
    Combo        : THDBValComboBox;
    {$ENDIF}

    NoDossier    : String;
    MI           : TMenuItem;
Begin
     // Blocage de l'écran si pas le droit de modifier ce type de compteur
     AccesFicheDupliquer (Ecran, GetField('PYR_PREDEFINI'), NoDossier, LectureSeule);
     SetControlEnabled('PGCOLORPRE', False); // Reforcer le champ coloré car il est mis à jour par la fonction

     // On ne peut pas dupliquer un compteur qui n'est pas modifiable
     If LectureSeule Then SetControlEnabled('BDUPLIQUER', False);

     // Restrictions des choix de prédéfini si pas tous les droits
     Predefini := GetField('PYR_PREDEFINI');
     {$IFDEF EAGLCLIENT}
     Combo := THValComboBox(GetControl('PYR_PREDEFINI'));
     {$ELSE}
     Combo := THDBValComboBox(GetControl('PYR_PREDEFINI'));
     {$ENDIF}
     If Combo = Nil Then Exit;

     // Si CEGID n'est pas autorisé, on le supprime de la combo
     If (Not CEG) And (Predefini <> 'CEG') Then
     Begin
          i := Combo.Values.IndexOf('CEG');
          If i >= 0 Then
          Begin
               Combo.Values.Delete(i);
               Combo.Items.Delete(i);
          End;
     End;

     // Si STANDARD n'est pas autorisé, on le supprime de la combo
     If (Not STD) And (Predefini <> 'STD') Then
     Begin
          i := Combo.Values.IndexOf('STD');
          If i >= 0 Then
          Begin
               Combo.Values.Delete(i);
               Combo.Items.Delete(i);
          End;
     End;

     // S'il ne reste que DOSSIER dans la combo, on la bloque
     If Combo.Values.Count = 1 Then
     Begin
          SetControlText('PYR_PREDEFINI', Combo.Values[0]);
          SetControlEnabled('PYR_PREDEFINI', False);
     End;

     If (Predefini = 'DOS') Or (Predefini = 'CEG') Then
     Begin
          For i := 0 To TPopupMenu(GetControl('DUPLIQUEPOPUP')).Items.Count - 1 Do
          Begin
               MI := TPopupMenu(GetControl('DUPLIQUEPOPUP')).Items[i];
               If (MI <> Nil) And (MI.Name = 'NIVEAU') Then MI.Enabled := False;
          End;
     End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 13/07/2007
Modifié le ... :   /  /    
Description .. : Contrôle la cohérence de périodicité entre celle saisie et 
Suite ........ : celle dépendant de la variable de présence sélectionnée
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.VerifiePeriodicite;
Var
     Periodicite, Variable, PeriodVar : String;
     Q : TQuery;
begin

     Periodicite := GetControlText('PYR_PERIODICITEPRE');
     Variable    := GetControlText('PYR_VARIABLEPRES');


     If (Periodicite <> '') And (Variable <> '') Then
     Begin
           if periodicitevar  <> '' // création du compteur à partir d'une variable qui peut ne pas être en base
           then
              periodVar := periodicitevar
           else
           begin
              Q := OpenSQL('SELECT PVA_VARPERIODICITE FROM VARIABLEPAIE WHERE PVA_VARIABLE="'+Variable+'"', True);
              If Not Q.EOF Then
              Begin
                 PeriodVar := Q.FindField('PVA_VARPERIODICITE').AsString;
              End;
              Ferme(Q);
           end;
           
           // Périodicité du compteur : Jour, Mensuel => Variable identique
           If ((Periodicite = periodiciteMensuelle) And (PeriodVar <> periodiciteMensuelle)) Or ((Periodicite = periodiciteJournaliere) And (PeriodVar <> periodiciteJournaliere)) Then
           Begin
               LastError := 1;
               PGIBox (TraduireMemoire('Les périodicités du compteur et de la variable ne concordent pas.'), Ecran.Caption);
           End;

           // Périodicité du compteur : Hebdo, Fin cycle => Variable hebdo ou fin cycle
           If ((Periodicite = periodiciteHebdomadaire) Or (Periodicite = periodiciteFinDeCycle)) And ((PeriodVar <> periodiciteHebdomadaire) And (PeriodVar <> periodiciteFinDeCycle)) Then
           Begin
               LastError := 1;
               PGIBox (TraduireMemoire('Les périodicités du compteur et de la variable ne concordent pas.'), Ecran.Caption);
           End;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 13/07/2007
Modifié le ... :   /  /
Description .. : Duplication d'un compteur. 2 modes :
Suite ........ : - par date
Suite ........ : - par niveau prédéfini
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.ClickDupliquer(Sender: TObject);
var T : TOB;
begin
     // Récupération des données actuelles
     T := TOB.Create('COMPTEURPRESENCE', Nil, -1);
     T.GetEcran(Ecran);

     // Création du nouvel enregistrement
    (GetControl('BINSERT') as TToolbarButton97).OnClick(Sender);

     // Récupération des anciennes données
     T.PutEcran(Ecran);

     // Duplication en modifiant la date : On ne modifie que la date
     If (TMenuItem(Sender).Name = 'DATE') Then
     Begin
          SetField('PYR_COMPTEURPRES', '');
          SetField('PYR_DATEVALIDITE', DateToStr(Date()));
     End
     // Duplication du niveau Prédéfini : On ne modifie que le prédéfini
     Else If (TMenuItem(Sender).Name = 'NIVEAU') Then
     Begin
          SetField('PYR_COMPTEURPRES', '');
          SetField('PYR_PREDEFINI', 'DOS');
          SetControlEnabled('PYR_PREDEFINI', False);
     End;

     MAJCouleur;

     // Blocage de la périodicité
     SetControlEnabled('PYR_PERIODICITEPRE', False); //PT4
     Duplic := true;
     T.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 27/08/2007 / PT3
Modifié le ... :   /  /
Description .. : Click sur Intégration en paie
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.ClickIntegrePaie(Sender: TObject);
var
Q : TQUERY;
st: string;
datevalid : Tdatetime;
begin
     {$IFDEF EAGLCLIENT}
     If (GetControl('PYR_INTEGREPAIE') As THCheckBox).Checked Then
     {$ELSE}
     If (GetControl('PYR_INTEGREPAIE') As THDBCheckBox).Checked Then
     {$ENDIF}
      Begin
        SetControlEnabled('PYR_RUBRIQUE',True);
        SetControlEnabled('PYR_TYPECHAMPRUB',True);
        // si le compteur est intégré en paie et la date de validité <= date dernière intégration, rubrique et champ non modifiables
        st := 'SELECT MAX(PYP_DATEFINBUL) AS DATEFINBUL FROM PRESENCESALARIE WHERE PYP_COMPTEURPRES = "'+getfield('PYR_COMPTEURPRES')+'"';
        Q:= opensql(st, true);
        if not Q.EOF then
        begin
          datevalid := StrToDate(getfield('PYR_DATEVALIDITE'));
          if datevalid <= Q.findfield('DATEFINBUL').asdatetime then
          begin
            setcontrolenabled('PYR_RUBRIQUE', false);
            setcontrolenabled('PYR_TYPECHAMPRUB', false);
            setcontrolenabled('PYR_INTEGREPAIE', false);
          end;
        end;

        ferme(Q);
      End
      Else
      Begin
          SetControlEnabled('PYR_RUBRIQUE',False);
          SetControlEnabled('PYR_TYPECHAMPRUB',False);
      End;
end;

procedure TOM_COMPTEURPRESENCE.MAJCouleur;
var
  EditColor : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
     {$IFDEF EAGLCLIENT}
     Edit := THEdit(GetControl('PYR_PGCOLORPRE'));
     {$ELSE}
     Edit := THDBEdit(GetControl('PYR_PGCOLORPRE'));
     {$ENDIF}

     If Assigned(Edit) then
     Begin
          if Edit.text<>'' then
          Begin
               Edit.Font.Color := StringToColor(Edit.Text);
               EditColor := THEdit(GetControl('PGCOLORPRE'));
               if Assigned(EditColor) then
               Begin
                    EditColor.Color := StringToColor(Edit.Text);
                    EditColor.Enabled := False;
               End;
          End;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 27/08/2007 / PT3
Modifié le ... :   /  /
Description .. : Lancement de l'écran d'association aux profils
Mots clefs ... :
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.ClickGestionAssociee(Sender: TObject);
begin
     VerifProfilsAssoc := False;

     // Validation du compteur
     (GetControl('BValider') As TToolbarButton97).Click;

     // Si le compteur est valide, on ouvre la gestion associée
     If LastError = 0 Then
          AglLanceFiche('PAY','PROFILCOMPTPRES','','','COMPTEUR;'+GetField('PYR_COMPTEURPRES')+';'+DateToStr(GetField('PYR_DATEVALIDITE'))+';'+GetField('PYR_LIBELLE'));

     VerifProfilsAssoc := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Paramétrage variable de présence
Mots clefs ... : 
*****************************************************************}
procedure TOM_COMPTEURPRESENCE.Variablepresence(Sender: TObject);
var
Q : TQuery;
nat,dos,pred,lafiche, variablepres : string;
begin
  variablepres :=  GetControlText('PYR_VARIABLEPRES');
  if (variablepres  <> '') then
  begin
     Q:= opensql('Select PVA_NATUREVAR,PVA_PREDEFINI,PVA_NODOSSIER FROM VARIABLEPAIE WHERE PVA_VARIABLE = "'+variablepres+'"',true) ;
     if not Q.EOF then
     begin
        nat := Q.Findfield('PVA_NATUREVAR').asstring;
        dos := Q.Findfield('PVA_NODOSSIER').asstring;
        Pred := Q.Findfield('PVA_PREDEFINI').asstring;

        lafiche := 'VARIABLE_' + nat;
        AGLLancefiche('PAY',lafiche,'',pred+ ';' + dos + ';' + variablepres, nat + '' +';PRE');
     end;
     ferme(Q);
  end;
end;

Initialization
  registerclasses ( [ TOM_COMPTEURPRESENCE ] ) ;
end.
