{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - FLO
Créé le ...... : 25/01/2008
Modifié le ... :   /  /
Description .. : Gestion des absences des collaborateurs
               : Vignette : PG_VIG_GESABS
               : Table    : ABSENCESALARIE
Mots clefs ... :
*****************************************************************
PT1  | 01/02/2008 | FLO | La vignette est maintenant capable de gérer le détail et la validation des absences
                          + ajout de critères supplémentaires
PT2  | 29/04/2008 | FLO | Prise en compte des salariés non sortis
}
unit PGVIGGESABS;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_GESABS= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    TypeConge           : string;  
    procedure AfficheListe;
    procedure AfficheDetail; //PT1
    procedure Visualisation;
    procedure SaisieAbsence;
    procedure PgCalculJourHeures ;
    procedure PGAffectRecapitulatif;
    procedure InitialiseRecap;
    procedure InitialiseSaisie; //PT1
    procedure ValiderDemandes; //PT1
 end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  uToolsPlugin,
  uToolsOWC,
  HEnt1,ULibEditionPaie,
  PGCalendrier;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_GESABS.GetInterface (NomGrille: string): Boolean;
begin
    If FParam = 'VALIDERDEM' Then //PT1
    Begin
         AvecSelection := True;
         Result := inherited GetInterface ('FListe');
         AvecSelection := False;
    End
    Else
        Result := inherited GetInterface ('');

    If FParam = '' Then ChargeSalaries  (Self, 'SALARIE', 'RESPONSABS'); //PT1

    If FParam = 'NOUVEAU' Then
    Begin
        WarningOk := False;
        MessageErreur := '';
    End;

    // Rafraîchissement de la vignette en fonction de la fonction lancée
    If (FParam <> 'DEBUTABS') And (FParam <> 'FINABS')
     And (FParam <> 'DEBUTDJ') And (FParam <> 'FINDJ')
     And (FParam <> 'TYPECONGE') And (FParam <> 'SALARIE') then
        Visualisation;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_GESABS.GetClauseWhere;
begin
  inherited;

    If GetControlValue('CRITMOTIF') <> '' Then //PT1
        ClauseWhere := ' AND PCN_TYPECONGE="'+GetControlValue('CRITMOTIF')+'" '
    Else
        ClauseWhere := '';

    If GetControlValue('CRITETAT') <> '' Then  //PT1
        ClauseWhere := ClauseWhere+' AND PCN_VALIDRESP="'+GetControlValue('CRITETAT')+'" ';
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_GESABS.RecupDonnees;
var T,TS : TOB;
begin
  inherited;
    If FParam = 'SUPPRIMER' Then
    Begin
        TS := TOB.Create('$LesAbsences', Nil, -1);
        T  := TOB.Create('$ITEM1', TS, -1);
        T.AddChampSupValeur('PCN_TYPEMVT',      GetControlValue('TYPEMVT'));
        T.AddChampSupValeur('PCN_TYPECONGE',    GetControlValue('TYPECONGE'));
        T.AddChampSupValeur('PCN_ORDRE',        GetControlValue('ORDRE'));
        T.AddChampSupValeur('PCN_NBJOURS',      StrToFloat(GetControlValue('JOURS')));
        T.AddChampSupValeur('PCN_VALIDRESP',    GetControlValue('ETAT'));
        T.AddChampSupValeur('PCN_SALARIE',      GetControlValue('SALARIE'));
        MessageErreur := SupprimeAbsence (TS);
        FreeAndNil(TS);
    End;

    If FParam = 'VALIDERDEM' Then //PT1
        ValiderDemandes;

    If FParam = 'SAISIE' Then //PT1
    Begin
        // Modification
        If GetControlValue('ORDRE') <> '' Then
            PgMajAbsEtatValidSal(GetControlValue('SALARIE'),GetControlValue('ETAT'),GetControlValue('TYPEMVT'),V_PGI.UserSalarie,'ATT',GetControlValue('ORDRE'))
        // Création
        Else
            SaisieAbsence;
    End;

    If (FParam = 'DEBUTABS') or (FParam = 'FINABS')
     or (FParam = 'DEBUTDJ') or (FParam = 'FINDJ')
     or (FParam = 'TYPECONGE') then
    Begin
        PgCalculJourHeures;
        WarningOK := False;
    End
    Else If (FParam = 'SALARIE') Then
    Begin
        WarningOK := False;
        PGAffectRecapitulatif;
    End
    Else If (FParam = 'VISUDETAIL') Then  //PT1
    Begin
        WarningOK := False;
        AfficheDetail;
        PGAffectRecapitulatif;
    End
    Else If (FParam <> 'NOUVEAU') Then
        AfficheListe;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_GESABS.DrawGrid (Grille: string);
begin
  inherited;

    SetFormatCol(Grille, 'DATEDEBUT',       'C.0 ---');
    SetFormatCol(Grille, 'DATEFIN',         'C.0 ---');
    SetFormatCol(Grille, 'NBJOURS',         'D 0 ---');
    SetFormatCol(Grille, 'ETAT',            'C.0 ---');

    SetWidthCol (Grille, 'TYPEMVT',         -1);
    SetWidthCol (Grille, 'CODETYPE',        -1);
    SetWidthCol (Grille, 'CODEETAT',        -1);
    SetWidthCol (Grille, 'ORDRE',           -1);
    SetWidthCol (Grille, 'SALARIE',         -1);
    SetWidthCol (Grille, 'NOM',             80);
    SetWidthCol (Grille, 'PRENOM',          80);
    SetWidthCol (Grille, 'TYPE',            80);
    SetWidthCol (Grille, 'DATEDEBUT',       60);
    SetWidthCol (Grille, 'DATEFIN',         60);
    SetWidthCol (Grille, 'NBJOURS',         60);
    SetWidthCol (Grille, 'ETAT',            60);
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_GESABS.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Affichage du panel correspondant à la fonction demandée --------------------------}

procedure PG_VIG_GESABS.Visualisation;
Var Visible : Boolean;
begin

    If (FParam = 'NOUVEAU') Or (FParam = 'VISUDETAIL') Then //PT1
    Begin
        Visible := False;
        If FParam = 'NOUVEAU' Then SetControlValue('TTITRESAISIE', 'Saisie d''une absence')  //PT1
        Else If FParam = 'VISUDETAIL' Then SetControlValue('TTITRESAISIE', 'Détail de l''absence');
    End
    Else
    Begin
        Visible := True;
        SetControlVisible('P2',     Not Visible);
    End;

    SetControlVisible('FListe',     Visible);
    SetControlVisible('TPERIODE',   Visible);
    SetControlVisible('PERIODE',    Visible);
    SetControlVisible('TBMOINS',    Visible);
    SetControlVisible('TBPLUS',     Visible);
    SetControlVisible('N1',         Visible);
    SetControlVisible('BNOUVEAU',   Visible);
    SetControlVisible('BVALIDERDEM',Visible);
    SetControlVisible('BMORE',      Visible);
    SetControlVisible('BSelectAll', Visible);    
    SetControlVisible('PANRECHERCHE', (Visible And ((GetControlValue('CRITMOTIF') <> '') Or (GetControlValue('CRITETAT') <> ''))));

    SetControlVisible('P1',         Not Visible);
    SetControlVisible('BANNULER',   Not Visible);
    SetControlVisible('BVALIDER',   Not Visible);
    SetControlVisible('BRECAP',     Not Visible);

    Visible := (FParam = 'VISUDETAIL');
    SetControlVisible('BSUPPRIMER', Visible);
    SetControlVisible('ETAT',       Visible);
    SetControlVisible('TETAT',      Visible);

    If (FParam = 'NOUVEAU') Then InitialiseSaisie; //PT1
end;

{-----Affichage des données de la vignette ---------------------------------------------}

procedure PG_VIG_GESABS.AfficheListe;
var
  StSQL,sN1         : String;
  DataTob           : TOB;
  EnDateDu, EnDateAu: TDatetime;
  i                 : integer;
  T                 : TOB;
  Q                 : TQuery;
begin
    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

    Try
        StSQL := 'SELECT PCN_SALARIE,PSA_LIBELLE,PSA_PRENOM,PCN_TYPEMVT,PCN_TYPECONGE, PMA1.PMA_LIBELLE AS LIBELLE, PCN_ORDRE, PCN_DATEDEBUTABS, PCN_DATEFINABS,'+
                'PCN_JOURS, PCN_VALIDRESP, CO2.CO_LIBELLE AS ETAT '+
                'FROM PGMVTABS '+
                'LEFT OUTER JOIN MOTIFABSENCE PMA1 ON PCN_TYPECONGE=PMA1.PMA_MOTIFABSENCE AND PMA1.PMA_TYPEMOTIF = "ABS" '+
                'LEFT OUTER JOIN COMMUN CO2 ON PCN_VALIDRESP=CO2.CO_CODE AND CO2.CO_TYPE="PAE" AND CO2.CO_CODE<>"REP" AND CO2.CO_ABREGE="EV" '+
                'LEFT JOIN DEPORTSAL ON PSE_SALARIE=PCN_SALARIE '+
                'WHERE '+AdaptByTypeResp('RESPONSABS', V_PGI.UserSalarie, False, 'PSE', False) + //PT2
                 ' AND (((PCN_DATEDEBUTABS >= "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEDEBUTABS <= "'+
                 USDATETIME (EnDateAu)+'"))'+
                 ' OR ((PCN_DATEFINABS >= "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEFINABS <= "'+
                 USDATETIME (EnDateAu)+'"))'+
                 ' OR ((PCN_DATEDEBUTABS < "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEFINABS > "'+
                 USDATETIME (EnDateAu)+'"))) '+ClauseWhere;
        // Pas de requête en cache sinon les modifications ne sont pas visibles
        //DataTob := OpenSelectInCache (StSQL); //PT1
        //ConvertFieldValue(DataTob);
        DataTob := Tob.Create('~TEMP', Nil, -1);
        Q := OpenSQL(StSQL, True);
        DataTob.LoadDetailDB('$DATA','','',Q,False);
        Ferme(Q);

        For i:=0 To DataTob.Detail.Count-1 Do
        Begin
          T := TOB.Create('£REPONSE',TobDonnees,-1);
          T.AddChampSupValeur('TYPEMVT',    DataTob.Detail[i].GetValue('PCN_TYPEMVT'));
          T.AddChampSupValeur('CODETYPE',   DataTob.Detail[i].GetValue('PCN_TYPECONGE'));
          T.AddChampSupValeur('TYPE',       DataTob.Detail[i].GetValue('LIBELLE'));
          T.AddChampSupValeur('ORDRE',      DataTob.Detail[i].GetValue('PCN_ORDRE'));
          T.AddChampSupValeur('SALARIE',    DataTob.Detail[i].GetValue('PCN_SALARIE'));
          T.AddChampSupValeur('NOM',        DataTob.Detail[i].GetValue('PSA_LIBELLE'));
          T.AddChampSupValeur('PRENOM',     DataTob.Detail[i].GetValue('PSA_PRENOM'));
          T.AddChampSupValeur('DATEDEBUT',  DateToStr(DataTob.Detail[i].GetValue('PCN_DATEDEBUTABS')));
          T.AddChampSupValeur('DATEFIN',    DateToStr(DataTob.Detail[i].GetValue('PCN_DATEFINABS')));
          T.AddChampSupValeur('NBJOURS',    DataTob.Detail[i].GetValue('PCN_JOURS'));
          T.AddChampSupValeur('CODEETAT',   DataTob.Detail[i].GetValue('PCN_VALIDRESP'));
          T.AddChampSupValeur('ETAT',       DataTob.Detail[i].GetValue('ETAT'));
        End;

        If (TobDonnees.Detail.Count = 0) Then
            PutGridDetail('FListe', TobDonnees);
    finally
        FreeAndNil (DataTob);
    end;
end;

{-----Affichage du détail d'une absence ------------------------------------------------}
//PT1
procedure PG_VIG_GESABS.AfficheDetail;
var
  StSQL             : String;
  DataTob,T         : TOB;
  Salarie,TypeMvt,Ordre : String;
  Actif             : Boolean;
  Q                 : TQuery;
begin
    Try
        // Récupération dans le flux de la ligne séléctionnée dans la grille sous forme de TOB
        T := GetLinkedValue('FListe');

        // Récupération de l'élément sélectionné
        If (T <> Nil) And (T.Detail.Count > 0) Then
        Begin
            Salarie   := T.Detail[0].GetValue('SALARIE');
            TypeMvt   := T.Detail[0].GetValue('TYPEMVT');
            Ordre     := T.Detail[0].GetValue('ORDRE');

            // Sauvegarde de la clé de l'enregistrement
            SetControlValue('TYPEMVT', TypeMvt);
            SetControlValue('ORDRE',   Ordre);

            // Récupération des données supplémentaires
            StSQL := 'SELECT PCN_DEBUTDJ,PCN_FINDJ,PCN_HEURES,PCN_LIBELLE,PCN_LIBCOMPL1,PCN_LIBCOMPL2,PCN_VALIDRESP,PCN_EXPORTOK '+
                     'FROM ABSENCESALARIE WHERE PCN_TYPEMVT="'+TypeMvt+'" AND PCN_SALARIE="'+Salarie+'" AND PCN_ORDRE="'+Ordre+'"';
            DataTob := Tob.Create('~TEMP', Nil, -1);
            Q := OpenSQL(StSQL, True);
            DataTob.LoadDetailDB('$DATA','','',Q,False);
            Ferme(Q);

            If DataTob <> Nil Then
            Begin
                SetControlValue('SALARIE',      Salarie);
                SetControlValue('TYPECONGE',    T.Detail[0].GetValue('CODETYPE'));
                SetControlValue('DATEDEBUTABS', T.Detail[0].GetValue('DATEDEBUT'));
                SetControlValue('DATEFINABS',   T.Detail[0].GetValue('DATEFIN'));
                SetControlValue('DEBUTDJ',      DataTob.Detail[0].GetValue('PCN_DEBUTDJ'));
                SetControlValue('FINDJ',        DataTob.Detail[0].GetValue('PCN_FINDJ'));
                SetControlValue('JOURS',        T.Detail[0].GetValue('NBJOURS'));
                SetControlValue('HEURES',       FloatToStr(DataTob.Detail[0].GetValue('PCN_HEURES')));
                SetControlValue('LIBELLE',      DataTob.Detail[0].GetValue('PCN_LIBELLE'));
                SetControlValue('LIBCOMPL1',    DataTob.Detail[0].GetValue('PCN_LIBCOMPL1'));
                SetControlValue('LIBCOMPL2',    DataTob.Detail[0].GetValue('PCN_LIBCOMPL2'));
                SetControlValue('ETAT',         DataTob.Detail[0].GetValue('PCN_VALIDRESP'));

                Actif := False;
                SetControlEnabled('SALARIE',        Actif);
                SetControlEnabled('TYPECONGE',      Actif);
                SetControlEnabled('DATEDEBUTABS',   Actif);
                SetControlEnabled('DATEFINABS',     Actif);
                SetControlEnabled('DEBUTDJ',        Actif);
                SetControlEnabled('FINDJ',          Actif);
                SetControlEnabled('LIBCOMPL1',      Actif);
                SetControlEnabled('LIBCOMPL2',      Actif);
                SetControlEnabled('ETAT',           not Actif);

                // Blocage complet si déjà intégré en paie
                Actif := (DataTob.Detail[0].GetValue('PCN_EXPORTOK') <> 'X') And (DataTob.Detail[0].GetValue('PCN_EXPORTOK') <> 'ENC');
                SetControlVisible('BSUPPRIMER',     Actif);
                SetControlVisible('BVALIDER',       Actif);
                SetControlEnabled('ETAT',           Actif);
            End;
        End;
    Finally
        If Assigned(DataTob) Then FreeAndNil (DataTob);
    end;
end;

{-----Validation de l'absence saisie ---------------------------------------------------}

procedure PG_VIG_GESABS.SaisieAbsence;
var
   Tabs,T_Param,T_Sal         : Tob;
   Error : integer;
   NomControl,ComplMessage    : String;
Begin
     If GetControlValue('SALARIE') = '' Then
     Begin
        MessageErreur := TraduireMemoire('Vous devez renseigner : ') + TraduireMemoire('le salarié');
        Exit; 
     End;

     { Chargement de l'absence saisie }
     T_Param := Tob.Create('Absence saisie',nil,-1);
     T_Param.AddChampSupValeur('PCN_SALARIE',       GetControlValue('SALARIE'));
     T_Param.AddChampSupValeur('PCN_TYPECONGE',     GetControlValue('TYPECONGE'));
     T_Param.AddChampSupValeur('PCN_DATEDEBUTABS',  StrToDate(GetControlValue('DATEDEBUTABS')));
     T_Param.AddChampSupValeur('PCN_DATEFINABS',    StrToDate(GetControlValue('DATEFINABS')));
     T_Param.AddChampSupValeur('PCN_DATEVALIDITE',  StrToDate(GetControlValue('DATEFINABS')));
     T_Param.AddChampSupValeur('PCN_DEBUTDJ' ,      GetControlValue('DEBUTDJ'));
     T_Param.AddChampSupValeur('PCN_FINDJ'   ,      GetControlValue('FINDJ'));
     T_Param.AddChampSupValeur('PCN_JOURS'   ,      GetControlValue('JOURS'));
     T_Param.AddChampSupValeur('PCN_HEURES'  ,      GetControlValue('HEURES'));
     T_Param.AddChampSupValeur('PCN_LIBELLE' ,      GetControlValue('LIBELLE'));
     T_Param.AddChampSupValeur('PCN_LIBCOMPL1' ,    GetControlValue('LIBCOMPL1'));
     T_Param.AddChampSupValeur('PCN_LIBCOMPL2' ,    GetControlValue('LIBCOMPL2'));

     { Chargement des infos salariés }
     T_Sal := ChargeTobSalarie(GetControlValue('SALARIE'));

     { Contrôle la conformité de la saisie d'absence }
     PgControleTobAbsFirst(T_Param,T_Sal,Error,NomControl,ComplMessage);
     if Error <> 0 then
       Begin                             { Gestion des erreurs }
       MessageErreur := TraduireMemoire(RecupMessageErrorAbsence(Error)+ComplMessage);
       End
    else
       Begin
       ComplMessage := '';
       Tabs := TOB.Create('ABSENCESALARIE',nil,-1);
       { Création de la tob absence au format de la Table }
       InitialiseTobAbsMotif(Tabs,T_Param,T_Sal,'VAL','-','RES','VIG');
       { Contrôle la possibilité de l'enregistrement de l'absence }
       PgControleTobAbsSecond(TAbs,T_Sal,'E','CREATION','RESP','SAL',Error,NomControl,ComplMessage);
       if Error <> 0 then                { Gestion des erreurs }
         MessageErreur := TraduireMemoire(RecupMessageErrorAbsence(Error)+ComplMessage)
       else
         Begin
         { Créatioin de l'absence }
         TAbs.InsertOrUpdateDB;
         { Calcul du récapitualtif et maj }
         PGExeCalculRecapAbsEnCours(GetControlValue('SALARIE'));
         PgAffectRecapitulatif;
         End;
       End;

     IF Assigned(T_Param)       then FreeAndNil(T_Param);
     IF Assigned(TAbs)          then FreeAndNil(Tabs);
     IF Assigned(T_Sal)         then FreeAndNil(T_Sal);

end;

{-----Calcul du nombre de jours et d'heures de l'absence en fonction de la saisie ------}

procedure  PG_VIG_GESABS.PgCalculJourHeures ();
var
  nb_j,nb_h : double;
  Etab         : string;
  Tob_motifAbs        : Tob;
  DJ, FJ : Integer;
  debabs ,finabs : TDateTime;
begin
  Tob_motifAbs := Tob.create('Les motifs',nil,-1);
  Tob_motifAbs.LoadDetailDB('MOTIFABSENCE','','',nil,False);
  Etab := RechDom('PGSALARIEETAB',GetControlValue('SALARIE'),False);

  debabs := StrToDate(GetControlValue('DATEDEBUTABS')) ;
  finabs := StrToDate(GetControlValue('DATEFINABS'));

  AffecteNodemj(GetControlValue('DEBUTDJ'), DJ);
  AffecteNodemj(GetControlValue('FINDJ'), FJ);
  TypeConge :=  GetControlValue('TYPECONGE');

  PgCalendrier.CalculNbJourAbsence(debabs,finabs,GetControlValue('SALARIE'),Etab,TypeConge,Tob_motifAbs,nb_j,nb_h,DJ,FJ);
  SetControlValue ('JOURS',FloatToStr(nb_j));
  SetControlValue ('HEURES',FloatToStr(nb_h));
  SetControlValue('LIBELLE',RendLibAbsence(Typeconge,GetControlValue('DEBUTDJ'),GetControlValue('FINDJ'),debabs,finabs));
  FreeAndNil(Tob_motifAbs);
end;

{-----Affichage du récapitulatif des compteurs -----------------------------------------}

procedure PG_VIG_GESABS.PgAffectRecapitulatif;
Var
  T_Recap : Tob;
begin
 { Chargement du recap et affectation des valeurs }
 T_Recap := ChargeTob_Recapitulatif(GetControlValue('SALARIE'));
 if Assigned(T_Recap) then
    Begin
    SetControlValue('LBRESTCP' , FloatToStr(T_Recap.GetValue('PRS_RESTN1')+T_Recap.GetValue('PRS_RESTN')));
    SetControlValue('LBRESTRTT', FloatToStr(T_Recap.GetValue('PRS_CUMRTTREST')));
    SetControlValue('LBATTCP'  , FloatToStr(T_Recap.GetValue('PRS_PRIATTENTE')));
    SetControlValue('LBATTRTT' , FloatToStr(T_Recap.GetValue('PRS_RTTATTENTE')));
    SetControlValue('LBVALCP'  , FloatToStr(T_Recap.GetValue('PRS_PRIVALIDE')));
    SetControlValue('LBVALRTT' , FloatToStr(T_Recap.GetValue('PRS_RTTVALIDE')));
    SetControlValue('LBACQCP'  , FloatToStr(T_Recap.GetValue('PRS_ACQUISCPSUIV')));
    SetControlValue('LBACQRTT' , FloatToStr(T_Recap.GetValue('PRS_ACQUISRTTSUIV')));
    End;
end;

{-----Initialisation du récapitulatif des compteurs ------------------------------------}

procedure PG_VIG_GESABS.InitialiseRecap;
begin
    SetControlValue('LBRESTCP' , '0');
    SetControlValue('LBRESTRTT', '0');
    SetControlValue('LBATTCP'  , '0');
    SetControlValue('LBATTRTT' , '0');
    SetControlValue('LBVALCP'  , '0');
    SetControlValue('LBVALRTT' , '0');
    SetControlValue('LBACQCP'  , '0');
    SetControlValue('LBACQRTT' , '0');
end;

{-----Initialisation du panel de saisie d'une absence ----------------------------------}
//PT1
procedure PG_VIG_GESABS.InitialiseSaisie;
Var
    DebAbs, FinAbs  : TDateTime;
    Tob_motifAbs    : Tob;
    Etab            : String;
    Nb_j,Nb_h       : Double;
Begin
  Try
    SetControlValue('TYPEMVT',      '');
    SetControlValue('ORDRE',        '');
    SetControlValue('SALARIE',      '');
    SetControlValue('TYPECONGE',    'PRI');
    SetControlValue('DATEDEBUTABS', DateToStr(Date));
    SetControlValue('DATEFINABS',   DateToStr(Date));
    SetControlValue('DEBUTDJ',      'MAT');
    SetControlValue('FINDJ',        'PAM');
    SetControlValue('LIBCOMPL1',    '');
    SetControlValue('LIBCOMPL2',    '');

    DebAbs := Date;
    FinAbs := Date;

    { Calcul du nombre de jour d'absence }
    Tob_motifAbs := Tob.create('Les motifs',nil,-1);
    Tob_motifAbs.LoadDetailDB('MOTIFABSENCE','','',nil,False);
    Etab := RechDom('PGSALARIEETAB',V_PGI.UserSalarie,False);
    PGCalendrier.CalculNbJourAbsence(DebAbs,FinAbs,V_Pgi.UserSalarie,Etab,'PRI',Tob_motifAbs,nb_j,nb_h,0,1);

    SetControlValue ('JOURS',       FloatToStr(nb_j));
    SetControlValue ('HEURES',      FloatToStr(nb_h));
    SetControlValue ('LIBELLE',     RendLibAbsence('PRI','MAT','PAM',debabs,finabs));

    SetControlEnabled('SALARIE',        True);
    SetControlEnabled('TYPECONGE',      True);
    SetControlEnabled('DATEDEBUTABS',   True);
    SetControlEnabled('DATEFINABS',     True);
    SetControlEnabled('DEBUTDJ',        True);
    SetControlEnabled('FINDJ',          True);
    SetControlEnabled('LIBCOMPL1',      True);
    SetControlEnabled('LIBCOMPL2',      True);

    InitialiseRecap;
  Finally
    FreeAndNil(Tob_motifAbs);
  End;
End;

{-----Validation des demandes sélectionnées --------------------------------------------}
//PT1
procedure PG_VIG_GESABS.ValiderDemandes ;
Var i : Integer;
    T : TOB;
Begin
    If TobSelection.Detail.Count = 0 Then
    Begin
        MessageErreur := TraduireMemoire('Veuillez sélectionner au moins une ligne.');
        Exit;
    End
    Else
    Begin
        For i := 0 to TobSelection.Detail.Count - 1 Do
        Begin
            T := TobSelection.Detail[i];
            If T.GetValue('CODEETAT') <> 'VAL' Then
                PgMajAbsEtatValidSal(T.GetValue('SALARIE'),'VAL',T.GetValue('TYPEMVT'),V_PGI.UserSalarie,T.GetValue('CODEETAT'),T.GetValue('ORDRE'));
        End;
    End;
End;

end.


