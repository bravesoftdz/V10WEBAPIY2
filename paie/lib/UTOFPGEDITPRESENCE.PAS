{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 02/07/2007
Modifié le ... :   /  /
Description .. : Source TOF des états de la gestion de présence :
Suite ........ : - PYP : Compteurs de présence
Mots clefs ... : TOF;PGEDITPRESENCE;COMPTEUR
*****************************************************************}
Unit UTOFPGEDITPRESENCE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     UTOF,
     UTOB;

Const NB_RUPTURES = 2;
Const NB_COLONNES = 6;
Const ANNU  = 1;
Const MENSU = 2;
Const HEBDO = 3;
Const JOUR  = 4;
Const CYCLE = 5;

Type
  TOF_PGEDITCOMPTEURSPRES = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    TobEtat   : TOB;
    Tablettes : Array of Array[0..1] of String;
    NbCols    : Integer;
    Function  RechTablette (NomChamp : String) : String;
    procedure OnChangeRuptures (Sender : TObject);
    procedure GenereDateFin (Sender : TObject);
    procedure ChangePeriodicite (Sender : TObject);
    Procedure ChangeTheme (Sender : TObject);
    procedure ChangeNbColonnes (Sender : TObject);
  end ;

Implementation

Uses
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1, uTobDebug, //mul,E dtREtat
     mul,
{$ELSE}
     eMul,
     eQRS1, 
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     EntPaie,
     HQry,
     P5Def;


procedure TOF_PGEDITCOMPTEURSPRES.OnUpdate ;
var OrderBy,Where,Select : String;
    i,num : Integer;
    Pages : TPageControl;
    Q : TQuery;
    TobCompteurs, T : TOB;
    TypePeriode : Integer;
    DD : TDateTime;
    Libelle,Tablette : String;
    Salarie,Compteur : String;
    Trouve : Boolean;
    TblColonnes : Array[1..NB_COLONNES] Of String;
    Requete : String;
begin
    Inherited ;
        If Assigned(TobEtat) Then FreeAndNil (TobEtat);
        TobEtat := Tob.Create('Edition',Nil,-1);

        // Construction dynamique du SELECT et ORDER BY pour les ruptures
        Select := 'PYP_SALARIE,PYP_COMPTEURPRES,PYR_LIBELLE,PYP_DATEDEBUTPRES,PYP_DATEFINPRES,PYP_QUANTITEPRES';
        OrderBy := '';
        For i := 1 To NB_RUPTURES Do
        Begin
                If (GetControlText('VRUPT'+IntToStr(i)) <> '') And (GetControlText('VRUPT'+IntToStr(i)) <> 'PYP_COMPTEURPRES') then
                Begin
                        If OrderBy <> '' then OrderBy := OrderBy + ',' + GetControlText('VRUPT'+IntToStr(i))
                        else OrderBy := GetControlText('VRUPT'+IntToStr(i));
                        Select := Select+','+GetControlText('VRUPT'+IntToStr(i));
                end;
        End;
        If OrderBy <> '' then OrderBy := OrderBy + ',';
        OrderBy := OrderBy + 'PYP_COMPTEURPRES,PYP_SALARIE';

        // Récupération de la clause Where
        Pages := TPageControl(GetControl('PAGES'));
        Where := RecupWhereCritere(Pages);

        // Génération des libellés de colonnes
        TypePeriode := StrToInt(GetControlText('PYP_PERIODICITEPRE'));
        i := 1;
        Requete := 'SELECT PYP_DATEDEBUTPRES,PYP_DATEFINPRES FROM PGPRESENCESAL '+Where+' GROUP BY PYP_COMPTEURPRES,PYR_LIBELLE,PYP_DATEDEBUTPRES,PYP_DATEFINPRES ORDER BY PYP_DATEDEBUTPRES';
        Q := OpenSQL (Requete, True);
        If Q.EOF Then
        Begin
               // Aucune donnée : pas la peine d'aller plus loin
               TFQRS1(Ecran).LaTob:= TobEtat;
               Ferme(Q);
               Exit;
        End;
        While (Not Q.EOF) And (i<=NbCols) Do
        Begin
               Libelle := '';
               Case TypePeriode Of
                    ANNU: Libelle := DateToStr(Q.FindField('PYP_DATEDEBUTPRES').AsDateTime)+' au '+DateToStr(Q.FindField('PYP_DATEFINPRES').AsDateTime);
                    MENSU:Libelle := DateToStr(Q.FindField('PYP_DATEDEBUTPRES').AsDateTime)+' au '+DateToStr(Q.FindField('PYP_DATEFINPRES').AsDateTime);
                    HEBDO:Libelle := DateToStr(Q.FindField('PYP_DATEDEBUTPRES').AsDateTime)+' au '+DateToStr(Q.FindField('PYP_DATEFINPRES').AsDateTime);
                    JOUR: Libelle := DateToStr(Q.FindField('PYP_DATEDEBUTPRES').AsDateTime);
               End;

               SetControlText('LIBELLE'+IntToStr(i),Libelle);
               TblColonnes[i] := DateToStr(Q.FindField('PYP_DATEDEBUTPRES').AsDateTime);
               i := i + 1;
               Q.Next;
        End;
        Ferme(Q);

        { Construction de la requête }

        Requete := 'SELECT ' + Select + ' FROM PGPRESENCESAL '+Where+' ORDER BY ' + OrderBy;

        // Création de la TOB
        Q := OpenSQL(Requete, True);
        TobCompteurs := Tob.Create('LesCompteurs', nil, -1);
        TobCompteurs.LoadDetailDB('LesCompteurs', '', '', Q, False);
        Ferme(Q);

        // Intégration des frais en une seule ligne dans la TOB finale
        For i := 0 To TobCompteurs.Detail.Count-1 Do
        Begin
                // Recherche du salarié
                Salarie  := TobCompteurs.Detail[i].GetValue('PYP_SALARIE');
                Compteur := TobCompteurs.Detail[i].GetValue('PYP_COMPTEURPRES');
                T := TobEtat.FindFirst(['PYP_SALARIE','PYP_COMPTEURPRES'],[Salarie,Compteur],True);
                If T = Nil Then
                Begin
                      T := Tob.Create('LaFille', TobEtat, -1);
                      T.AddChampSupValeur('PYP_SALARIE',      Salarie);
                      T.AddChampSupValeur('PYP_COMPTEURPRES', Compteur);
                      T.AddChampSupValeur('PYR_LIBELLE',      TobCompteurs.Detail[i].GetValue('PYR_LIBELLE'));

                      // Ajout des ruptures
                      T.AddChampSupValeur('RUPTURE1',TobCompteurs.Detail[i].GetValue(GetControlText('VRUPT1')));
                      Tablette := RechTablette(GetControlText('VRUPT1'));
                      If (Tablette <> '') Then T.AddChampSupValeur('LIBRUPTURE1',RechDom(Tablette, TobCompteurs.Detail[i].GetValue(GetControlText('VRUPT1')), False));

                      T.AddChampSupValeur('RUPTURE2',TobCompteurs.Detail[i].GetValue(GetControlText('VRUPT2')));
                      Tablette := RechTablette(GetControlText('VRUPT2'));
                      If (Tablette <> '') Then T.AddChampSupValeur('LIBRUPTURE2',RechDom(Tablette, TobCompteurs.Detail[i].GetValue(GetControlText('VRUPT2')), False));

                      // Création à vide de tous les champs correspondant aux compteurs
                      T.AddChampSupValeur('MTT1',0);
                      T.AddChampSupValeur('MTT2',0);
                      T.AddChampSupValeur('MTT3',0);
                      T.AddChampSupValeur('MTT4',0);
                      T.AddChampSupValeur('MTT5',0);
                      T.AddChampSupValeur('MTT6',0);
                End;

                // Mise à jour de la TOB pour la quantité
                If T <> Nil Then
                Begin
                      Num := 1;
                      Trouve := False;
                      While (Not Trouve) And (Num <= NbCols) Do
                      Begin
                             DD := TobCompteurs.Detail[i].GetValue('PYP_DATEDEBUTPRES');
                             If TblColonnes[Num] = DateToStr(DD) Then
                             Begin
                                  T.PutValue('MTT'+IntToStr(Num), TobCompteurs.Detail[i].GetValue('PYP_QUANTITEPRES'));
                                  Trouve := True;
                             End;
                             Num := Num + 1;
                      End;
                End;
        End;

        // Libération de la TOB intermédiaire
        FreeAndNil (TobCompteurs);

        // Attribution de la TOB à l'état
        TFQRS1(Ecran).LaTob:= TobEtat;
end ;

procedure TOF_PGEDITCOMPTEURSPRES.OnArgument (S : String ) ;
Var
     Edit  : THEdit;
     i,Num : Integer;
     Combo : THValComboBox;
     CkBox : TCheckBox;
     Spin  : THSpinEdit;
begin
  Inherited ;

  // Céation d'un listener sur la date de début
  Edit := THEdit(GetControl('PYP_DATEDEBUTPRES'));
  If Edit <> Nil then Edit.OnExit := GenereDateFin;

  // Céation d'un listener sur le thème
  Combo := THValComboBox(GetControl('PYP_THEMEPRE'));
  If Combo <> Nil then Combo.OnChange := ChangeTheme;

  // Idem sur le nombre d'occurrences
  Spin := THSpinEdit(GetControl('NBOCC'));
  If Spin <> Nil then Spin.OnChange := ChangeNbColonnes;

  // Idem sur la périodicité
  Combo := THValComboBox(GetControl('PYP_PERIODICITEPRE'));
  If Combo <> Nil then Combo.OnChange := ChangePeriodicite;

  // Suppression des éléments journaliers et cycliques qui n'ont pas d'intérêt pour l'édition
  For i := Combo.Values.Count-1 DownTo 0 Do
  Begin
     // 004 : Journalier, 005 : Cycle
     If (Combo.Values[i] = '004') Or (Combo.Values[i] = '005') Then
     Begin
          Combo.Values.Delete(i);
          Combo.Items.Delete(i);
     End;
  End;

  // Sélection de l'hebdo par défaut
 // SetControlText('PYP_PERIODICITEPRE', Combo.Values[2]);

  // Adaptation des ruptures
  For i := 1 To NB_RUPTURES do
  begin
          Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
          If Combo <> Nil Then
          begin
                  // Rendre invisible toutes les check-boxes de saut de page par défaut
                  CkBox := TCheckBox(GetControl('CSRUPT'+IntToStr(i)));
                  CkBox.Visible := False;

                  // Sur le changement de valeur d'une combo, il faut afficher le saut de page
                  Combo.OnChange := OnChangeRuptures;

                  // Ajout des valeurs
                  If i <> 1 Then
                  Begin
                       Combo.Items.Add ('');
                       Combo.Values.Add('');
                  End;
                  Combo.Items.Add ('Compteur');
                  Combo.Values.Add('PYP_COMPTEURPRES');
                  Combo.Items.Add ('Etablissement');
                  Combo.Values.Add('PYP_ETABLISSEMENT');
                  For Num := 1 To VH_Paie.PGNbreStatOrg Do
                  Begin
                      Case Num Of
                         1 : Combo.Items.Add(VH_Paie.PGLibelleOrgStat1);
                         2 : Combo.Items.Add(VH_Paie.PGLibelleOrgStat2);
                         3 : Combo.Items.Add(VH_Paie.PGLibelleOrgStat3);
                         4 : Combo.Items.Add(VH_Paie.PGLibelleOrgStat4);
                      End;
                      Combo.Values.Add('PYP_TRAVAILN'+IntToStr(Num));
                  End;
          end;
  end;
  SetControlText('VRUPT1', 'PYP_COMPTEURPRES');

  // Recherche des libellés des zones TRAVAIL, CODESTAT, LIBREPCMB
  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PYP_TRAVAILN' + IntToStr(Num)), GetControl('TPYP_TRAVAILN' + IntToStr(Num)));
  end;

  VisibiliteStat(GetControl('PYP_CODESTAT'), GetControl('TPYP_CODESTAT'));

  for Num := 1 to 4 do
  begin
    VisibiliteChamplibresal(Inttostr(Num), GetControl('PYP_LIBREPCMB' + IntToStr(Num)), GetControl('TPYP_LIBREPCMB' + IntToStr(Num)));
  end;

  If Assigned(TobEtat) Then FreeAndNil(TobEtat);
end ;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/07/2007
Modifié le ... :   /  /
Description .. : Gestion des checkboxes pour les ruptures
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITCOMPTEURSPRES.OnChangeRuptures(Sender: TObject);
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
     TCheckBox(GetControl('CSRUPT'+IntToStr(i))).Visible := Affiche;
     // Si on cache, on remet également l'état à "décoché"
     If (Affiche = False) Then TCheckBox(GetControl('CSRUPT'+IntToStr(i))).Checked := Affiche;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/07/2007
Modifié le ... :   /  /
Description .. : Génération de la date de fin de la période de prise en 
Suite ........ : compte des compteurs
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITCOMPTEURSPRES.GenereDateFin(Sender: TObject);
Var
     TypePeriode : Integer;
     DD,DF       : TDateTime;
begin
      If (GetControlText('PYP_PERIODICITEPRE') <> '') And (StrToDate(GetControlText('PYP_DATEDEBUTPRES')) <> iDate1900) Then
      Begin
          TypePeriode := StrToInt(GetControlText('PYP_PERIODICITEPRE'));
          DD          := StrToDate(GetControlText('PYP_DATEDEBUTPRES'));
          Case TypePeriode Of
               ANNU  : DF := PlusDate(DD, NbCols, 'A') ;              // Génération pour x années
               MENSU : DF := FinDeMois(PlusDate(DD, NbCols, 'M'));   // Génération pour x mois
               HEBDO : DF := PlusDate(DD, NbCols, 'S');               // Génération pour x semaines
               JOUR  : DF := PlusDate(DD, NbCols, 'J');               // Génération pour x jours
               Else
                    DF := DD;
          End;
          SetControlText('PYP_DATEFINPRES', DateToStr(DF));
      End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/07/2007
Modifié le ... :   /  /    
Description .. : Recherche de la tablette associée à un champ
Mots clefs ... :
*****************************************************************}
Function TOF_PGEDITCOMPTEURSPRES.RechTablette (NomChamp : String) : String;
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
Auteur  ...... : FLO
Créé le ...... : 04/07/2007
Modifié le ... :   /  /    
Description .. : Met à jour la date de fin et le nombre de colonnes par 
Suite ........ : défaut sur sélection d'une périodicité
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITCOMPTEURSPRES.ChangePeriodicite(Sender: TObject);
Var
     TypePeriode,Nb : Integer;
begin
      // Mise à jour du nombre de colonnes par défaut
      If (GetControlText('PYP_PERIODICITEPRE') <> '') Then
      Begin
          TypePeriode := StrToInt(GetControlText('PYP_PERIODICITEPRE'));
          Nb := 1;
          Case TypePeriode Of
               ANNU  : Nb := 3;     // Génération pour 3 années
               MENSU : Nb := 6;     // Génération pour un semestre
               HEBDO : Nb := 4;     // Génération pour un mois
          End;
          SetControlText('NBOCC', IntToStr(Nb)) ;
          NbCols := Nb;

          // Actualisation de la date de fin
          GenereDateFin(Sender);
      End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 04/07/2007
Modifié le ... :   /  /    
Description .. : Met à jour la date de fin avec le nombre de colonnes
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITCOMPTEURSPRES.ChangeNbColonnes(Sender: TObject);
begin
      NbCols := StrToInt(GetControlText('NBOCC'));

      // Actualisation de la date de fin
      GenereDateFin(Sender);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 04/07/2007
Modifié le ... :   /  /    
Description .. : Rafraîchissement de la combo des compteurs en fonction 
Suite ........ : du thème choisi
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGEDITCOMPTEURSPRES.ChangeTheme(Sender: TObject);
begin
      If (GetControlText('PYP_THEMEPRE') <> '') Then
      Begin
          SetControlText('PYP_COMPTEURPRES', '');
          SetControlProperty('PYP_COMPTEURPRES', 'Plus', ' AND PYR_THEMEPRE="'+GetControlText('PYP_THEMEPRE')+'"');
      End;
end;

Initialization
  registerclasses ( [ TOF_PGEDITCOMPTEURSPRES ] ) ;
end.

