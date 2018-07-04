{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... : 20/06/2006
Description .. : utilitaires vignettes
Suite ........ : - DatesPeriode : rend date début et date fin de période +
Suite ........ : libellé de la période.
Suite ........ : - GlisserPeriode : si action TBMOIS ou TBPLUS renvoi
Suite ........ :  date de référence  recalculée
Mots clefs ... :
*****************************************************************
PT1  | 18/01/2008 | FLO | Ajout des inscriptions au prévisionnel et réalisé,
     |            |     | chargement de la liste des emplois, et ajout d'un paramètre
     |            |     | pour le chargement des salariés afin de prendre en compte l'option non nominative
PT2  | 23/01/2008 | FLO | Ajout de la fonction ChargeValeursCombo pour charger les données d'une tablette dynamiquement
PT3  | 24/01/2008 | FLO | Ajout de la fonction de calcul des totaux pour n'importe qu'elle grille
PT4  | 28/01/2008 | FLO | Ajout des fonctions de gestion des absences + prise en compte des assistants + chargement des salariés et
     |            |     | libellés emplois pour tous types de responsables
PT5  | 01/02/2008 | FLO | Ajout d'une fonction pour retourner les libellés des booléens + Modif gestion des n° de dossier
PT6  | 18/03/2008 | FLO | Prise en compte des traitements multidossier
PT7  | 18/03/2008 | FLO | Inscription à une session toujours effectuée par défaut
PT8  | 18/03/2008 | FLO | Ajout du préfixe pour AdaptByTypeResp + Ajout de la sauvegarde du type utilisateur
PT9  | 25/03/2008 | FLO | Ajout du paramètre "Responsable" pour les inscriptions dans le cas où il s'agit
     |            |     | d'une inscription par l'adjoint ou le secrétaire.
PT10 | 03/04/2008 | FLO | Adaptation de la récupération du type utilisateur : peut-être de plusieurs niveaux en même temps
PT11 | 03/04/2008 | FLO | Création d'un objet pour enregistrer le type utilisateur. Pourquoi? Sais pas... Comme ça d'un coup ça ne marche plus avec un String...
PT12 | 03/04/2008 | FLO | Mise en copie du responsable formation général
PT13 | 07/04/2008 | FLO | Correction de l'adaptation par responsable pour qu'un adjoint ne voit que les salariés du service dont il fait partie
PT14 | 15/04/2008 | FLO | Récupération du mail d'abord depuis la table UTILISAT, puis ensuite depuis DEPORTSAL
PT15 | 29/04/2008 | FLO | Ajout d'une fonction incluant la clause where sur les salariés sortis
PT16 | 30/04/2008 | FLO | Correction du GetBase en mode entreprise
PT17 | 05/05/2008 | FLO | La table CHOIXCOD doit être préfixée de la base commune lorsque le partage formation est actif
PT18 | 11/06/2008 | FLO | Pas d'envoi de mail si pas de salarié
}
unit PgVIGUTIL;

interface

uses
  classes,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  UTob,
  HCtrls,
  UToolsPlugin;

Const UNDEFINED = '#UNDEF#'; //PT1
Const OUI = 'Oui';  //PT5
Const NON = 'Non';
Const ERR_MSG = 'Une exception s''est produite durant le traitement.';
Const TYPE_UTILISATEUR = 'TYPE_UTILISATEUR_';

    function PgRecupInfoUserAndResp (UserSal : string) : Tob;
    function PgValidMailIndicUser(UserResp : string) : tob;
    function PgRecupBodyMailUser(UserResp : string; TAbs : Tob) : tob;

    { Fonctions communes }
    function  DatesPeriode (DateJ: TDateTime; var DateDu, DateAu : TDateTime; Periode : string; var SN : string):  Boolean;
    function  GlisserPeriode (Param : string; var DateJ: TDateTime; Periode : string):  Boolean;
    procedure AfficheTotal (CurrentObject : TVignettePlugin; TobDonnees : TOB; NumColTitre : Integer = 0); //PT3
    function  LibBool (Valeur : String) : String; //PT5
	Function  EnvoieMail (Salarie, Titre : String; Corps : HTStrings) : Boolean;
	Function  GetBase (Dossier,LaTable : String) : String; //PT6
    Procedure SauveTypeUtilisateur (Hierarchie : String); //PT8
	Function  RecupTypeUtilisateur (Hierarchie : String) : String; //PT8

    { Fonctions communes de paie }
    function PgRecupAbsUser (UserSal : String) : TOB ;
    function PgRecupRecapUser () : TOB ;
    function PgRecupJourHeureAbs (TAbs : Tob ) : TOB ;
    function PgValidAbsUser (TAbs : Tob ) : TOB ;
    Function SupprimeAbsence (TobSelection : TOB; Salarie : String = '') : String; //PT4
    Function SalariesNonSortis (Prefixe : String = 'PSA') : String; //PT15

    { Fonctions communes de chargement }
    procedure ChargeSalaries (CurrentObject : TVignettePlugin; Combo, TypeResponsable : String; InclureIndefini : Boolean = False; Responsable : String = ''; ServiceSup : Boolean = False);   //PT1 //PT4
    procedure ChargeEmplois (CurrentObject : TVignettePlugin; Combo, TypeResponsable : String; Responsable : String = ''; ServiceSup : Boolean = False); //PT1 //PT4
    procedure ChargeValeursCombo (CurrentObject : TVignettePlugin; Combo, Tablette : String; Requete : String = ''; TousPossible : Boolean = False; TousLibelle : String = '<<TOUS>>'); //PT2
    Function  AdaptByTypeResp (TypeResponsable, Responsable : String; ServiceSup : Boolean; Prefixe : String='PSE'; NonSortis : Boolean = True) : String; //PT4 //PT8

    { Fonctions communes Formation }
    procedure ChargeStages   (CurrentObject : TVignettePlugin; Combo : String; ClauseWhere : String = '');
    Function  InscritSalarieSession (Stage, Session, MillesimeEC, Salarie, TypePlan : String; TypeResponsable : String) : String;  //PT1 //PT9
    Function  InscritSalariePrevisionnel (Stage, MillesimeEC, Salarie, NbInsc, TypePlan, LibEmploi : String; HeuresTT, HeuresHT : Double; DemandeSalarie : Boolean; TypeResponsable : String) : String;  //PT1 //PT9

    //PT11
    Type TTypeUser = class (TObject)
    public
        Liste : String;
    End;

implementation
 uses
  Hent1,
  SysUtils,
  PgCalendrier,
  ParamSoc,
  PGOutilsFormation,
  variants,
  ULibEditionPaie,
  StrUtils,
  eHttp,
  Registry,Windows,esession;

function DatesPeriode (DateJ: TDateTime; var DateDu, DateAu : TDateTime; Periode : string; var SN : string):  Boolean;
var
  JourSemaine                       : integer;
  Year, Month, Day                  : Word;
  iNumSemaine                       : integer;
  SDate                             : string;

begin
  JourSemaine :=  DayofWeek(DateJ);
  DecodeDate(DateJ, Year, Month, Day);
  if (Periode = '001') or (Periode = 'J')then
  // Jour
  begin
    DateDu := DateJ;
    Dateau := DateJ;
    sN := DateTimetoStr(DateJ);
  end
  else
  if (Periode = '002') or (Periode = 'N')then
  // Semaine
  begin
    DateDu := DateJ - (JourSemaine - 2);
    DateAu :=  DateDu + 6;
    iNumSemaine := NumSemaine(DateJ,IsLeapYear(Year));
    sN := 'Sem.' + InttoStr(iNumSemaine) + ' ' +IntToStr(Year);
  end
  else
  if (Periode = '003') or (Periode = 'M') then
  // Mois
  begin
    DateDu := DEBUTDEMOIS(DateJ);
    DateAu := FINDEMOIS(DateJ);
    case Month of
      1 :sN := 'Janv.'+ ' ' + IntToStr(Year);
      2 :sN := 'Fév.'+ ' ' + IntToStr(Year);
      3 :sN := 'Mars'+ ' ' + IntToStr(Year);
      4 :sN := 'Avril'+ ' ' + IntToStr(Year);
      5 :sN := 'Mai'+ ' ' + IntToStr(Year);
      6 :sN := 'Juin'+ ' ' + IntToStr(Year);
      7 :sN := 'Juil.'+ ' ' + IntToStr(Year);
      8 :sN := 'Août'+ ' ' + IntToStr(Year);
      9 :sN := 'Sept.'+ ' ' + IntToStr(Year);
      10 :sN := 'Oct.'+ ' ' + IntToStr(Year);
      11 :sN := 'Nov.'+ ' ' + IntToStr(Year);
      12 :sN := 'Déc.'+ ' ' + IntToStr(Year);
    end;
  end
  else
  if (Periode = '004') or (Periode = 'T') then
  // Trimestre
  begin
    Day := 1;
    Case Month of
      1..3 :Month := 1;
      4..6 :Month := 4;
      7..9 :Month := 7;
      10..12 :Month := 10;
    end;
    DateDu := EncodeDate(Year, Month, Day);
    case Month of
      1..3 :Month := 3;
      4..6 :Month := 6;
      7..9 :Month := 9;
      10..12 :Month := 12;
    end;

    case Month of
      1, 3, 5, 7, 8, 10, 12 :Day := 31;
      4, 6, 9, 11           :Day := 30;
    else
      if IsLeapYear(Year) then
      begin
        Day := 29;
      end
      else
      begin
        Day := 28;
      end;
    end;
    DateAu := EncodeDate(Year, Month, Day);
    case Month of
      1..3 :sN := '1°Trim. '+ IntToStr(Year);
      4..6 :sN := '2°Trim. '+ IntToStr(Year);
      7..9 :sN := '3°Trim. '+ IntToStr(Year);
      10..12 :sN := '4°Trim. '+ IntToStr(Year);
    end;
  end
  else
  if (Periode = '005') or (Periode = 'Q') then
  // Quadrimestre
  begin
    Day := 1;
    case Month of
      1..4 :Month := 1;
      5..8 :Month := 5;
      9..12 :Month := 9;
    end;
    DateDu := EncodeDate(Year, Month, Day);
    case Month of
      1..4 :Month := 4;
      5..8 :Month := 8;
      9..12 :Month := 12;
    end;
    case Month of
      1, 3, 5, 7, 8, 10, 12 :Day := 31;
      4, 6, 9, 11           :Day := 30;
    else
      if IsLeapYear(Year) then
      begin
        Day := 29;
      end
      else
      begin
        Day := 28;
      end;
    end;
    DateAu := EncodeDate(Year, Month, Day);
    case Month of
      1..4 :sN := '1°Quad. '+ InttoStr(Year);
      5..8 :sN := '2°Quad. '+ InttoStr(Year);
      9..12 :sN := '3°Quad. '+ InttoStr(Year);
    end;
   end
   else
   if (Periode = '006') or (Periode = 'A')then
   // Année
   begin
     SDate:= '01/01/'+InttoStr(Year);
     DateDu := StrToDate (SDate);
     SDate:= '31/12/'+InttoStr(Year);
     DateAu := StrToDate (SDate);
     sN := InttoStr(Year);
   end;

    Result := TRUE;
end;
function GlisserPeriode (Param : string; var DateJ: TDateTime; Periode : string):  Boolean;
begin
    if Param = 'plus' then
    begin
         if Periode = '001' then  DateJ := PlusDate (DateJ, +1, 'J');
         if Periode = '002' then  DateJ := PlusDate (DateJ, +7, 'J');
         if Periode = '003' then  DateJ := PlusDate (DateJ, +1, 'M');
         if Periode = '004' then  DateJ := PlusDate (DateJ, +3, 'M');
         if Periode = '005' then  DateJ := PlusDate (DateJ, +4, 'M');
         if Periode = '006' then  DateJ := PlusDate (DateJ, +1, 'A');
    end
    else
    if Param = 'moins' then
    begin
         if Periode = '001' then  DateJ := PlusDate (DateJ, -1, 'J');
         if Periode = '002' then  DateJ := PlusDate (DateJ, -7, 'J');
         if Periode = '003' then  DateJ := PlusDate (DateJ, -1, 'M');
         if Periode = '004' then  DateJ := PlusDate (DateJ, -3, 'M');
         if Periode = '005' then  DateJ := PlusDate (DateJ, -4, 'M');
         if Periode = '006' then  DateJ := PlusDate (DateJ, -1, 'A');
    end;

    Result := TRUE;
end;


{function PGRendMailSalarie () : String;
var
  Q                     : TQuery;
  st                    : string;
begin
  st := 'SELECT US_AUXILIAIRE FROM UTILSAT WHERE US_UTILISATEUR="'+V_PGI.UserSalarie+'"';
  Q := OpenSql(st, true);
  FERME (Q);

end;}

function PgRecupAbsUser (UserSal : String) : TOB ;
var
  Q                     : TQuery;
  st                    : string;
begin
  Result := nil;

  st := 'SELECT PCN_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE, '+
        'PSA_DATESORTIE,PCN_DATEVALIDITE,PCN_TYPECONGE,PCN_SENSABS, '+
        'PCN_JOURS,PCN_HEURES,PCN_LIBELLE,PCN_CODETAPE, '+
        'PCN_MVTDUPLIQUE,PCN_TYPEMVT,PCN_DATEDEBUTABS,PCN_DATEFINABS, '+
        'PCN_ORDRE,PCN_ETABLISSEMENT,PCN_SAISIEDEPORTEE, '+
        'PCN_VALIDSALARIE,PCN_VALIDRESP,PCN_VALIDABSENCE, '+
        'PCN_EXPORTOK,PCN_ETATPOSTPAIE  FROM ABSENCESALARIE '+
        'LEFT JOIN SALARIES ON PSA_SALARIE=PCN_SALARIE '+
        'WHERE (PCN_TYPECONGE="PRI" OR PCN_TYPEMVT="ABS") '+
        'AND PCN_MVTDUPLIQUE="-" AND PCN_ETATPOSTPAIE <> "NAN" '+
        'AND PCN_SALARIE = "'+UserSal+'" '+
//        '(PSA_DATESORTIE<="01/01/1900" OR PSA_DATESORTIE is null '+
//        'OR PSA_DATESORTIE>="'+USDateTime(GetParamSocSecur('SO_PGECABDATEINTEGRATION',idate1900)+'")
        'ORDER BY PCN_DATEDEBUTABS DESC';
  Q := OpenSql(st, true);
  if not Q.Eof then
     Begin
     Result := TOB.Create('Les absences',nil,-1);
     Result.LoadDetailDB ('Les absences', '', '', Q, False);
     End;
  Ferme (Q);
end;



function PgRecupRecapUser () : TOB ;
begin
ddwriteln('V_PGI.UserSalarie='+V_PGI.UserSalarie);
     Result := ChargeTob_Recapitulatif(V_PGI.UserSalarie);
end;

function PgRecupJourHeureAbs (TAbs : Tob) : TOB ;
var
  nb_j,nb_h          : double;
  Etab,TypeConge     : string;
  TAbsUser,Tob_motifAbs       : Tob;
  DJ, FJ             : Integer;
  debabs ,finabs     : TDateTime;
begin
ddwriteln('PgRecupJourHeureAbs');
  Tob_motifAbs := Tob.create('Les motifs',nil,-1);
  Tob_motifAbs.LoadDetailDB('MOTIFABSENCE','','',nil,False);
  Etab := RechDom('PGSALARIEETAB',TAbs.GetValue('SALARIE'),False);
  debabs := TAbs.GetValue('DATEDEBUTABS');//EPZ StrToDate(TAbs.GetValue('DATEDEBUTABS')) ;
  finabs := TAbs.GetValue('DATEFINABS'); //EPZ StrToDate(TAbs.GetValue('DATEFINABS'));

  AffecteNodemj(TAbs.GetValue('DEBUTDJ'), DJ);
  AffecteNodemj(TAbs.GetValue('FINDJ'), FJ);
  TypeConge :=  TAbs.GetValue('TYPECONGE');

  PgCalendrier.CalculNbJourAbsence(debabs,finabs,TAbs.GetValue('SALARIE'),Etab,TypeConge,Tob_motifAbs,nb_j,nb_h,DJ,FJ);

  TAbsUser := Tob.Create('AbsUser',nil,-1);
  TAbsUser.AddChampSupValeur('JOURS',nb_j);
  TAbsUser.AddChampSupValeur('HEURES',nb_h);
ddwriteln('PgRecupJourHeureAbs-nb_j'+FloatToStr(nb_j));
ddwriteln('PgRecupJourHeureAbs-nb_h'+FloatToStr(nb_h));
  TAbsUser.AddChampSupValeur('LIBELLE',RendLibAbsence(Typeconge,TAbs.GetValue('DEBUTDJ'),TAbs.GetValue('FINDJ'),debabs,finabs));

  Result := TAbsUser;
  FreeAndNil(Tob_motifAbs);

ddwriteln('PgRecupJourHeureAbs-Fin');
end;

function PgValidAbsUser (TAbs : Tob ) : TOB ;
var
   T_Param,T_Sal         : Tob;
   Error : integer;
   MessageErreur,NomControl,ComplMessage    : String;
Begin
  result := nil;
  MessageErreur := '';
ddwriteln('PgValidAbsUser Salarie = '+TAbs.GetValue('SALARIE'));
  { Chargement de l'absence saisie }
  T_Param := Tob.Create('Absence saisi',nil,-1);
  T_Param.AddChampSupValeur('PCN_SALARIE'       ,TAbs.GetValue('SALARIE'));
  T_Param.AddChampSupValeur('PCN_TYPECONGE'     ,TAbs.GetValue('TYPECONGE'));
//EPZ :
  T_Param.AddChampSupValeur('PCN_TYPEMVT'     ,'');

  T_Param.AddChampSupValeur('PCN_DATEDEBUTABS'  ,TAbs.GetValue('DATEDEBUTABS'));
  T_Param.AddChampSupValeur('PCN_DATEFINABS'    ,TAbs.GetValue('DATEFINABS'));
  T_Param.AddChampSupValeur('PCN_DATEVALIDITE'  ,TAbs.GetValue('DATEFINABS'));
  T_Param.AddChampSupValeur('PCN_DEBUTDJ'       ,TAbs.GetValue('DEBUTDJ'));
  T_Param.AddChampSupValeur('PCN_FINDJ'         ,TAbs.GetValue('FINDJ'));
  T_Param.AddChampSupValeur('PCN_JOURS'         ,TAbs.GetValue('JOURS'));
  T_Param.AddChampSupValeur('PCN_HEURES'        ,TAbs.GetValue('HEURES'));
  T_Param.AddChampSupValeur('PCN_LIBELLE'       ,TAbs.GetValue('LIBELLE'));
  T_Param.AddChampSupValeur('PCN_LIBCOMPL1'     ,TAbs.GetValue('LIBCOMPL1'));
  T_Param.AddChampSupValeur('PCN_LIBCOMPL2'     ,TAbs.GetValue('LIBCOMPL2'));
ddwriteln('ChargeTobSalarie');
  { Chargement des infos salariés }
  T_Sal := ChargeTobSalarie(TAbs.GetValue('SALARIE'));
ddwriteln('ChargeTobSalarie done');

  { Contrôle la conformité de la saisie d'absence }
ddwriteln('PgControleTobAbsFirst');
  PgControleTobAbsFirst(T_Param,T_Sal,Error,NomControl,ComplMessage);
ddwriteln('PgControleTobAbsFirst done');
  if Error <> 0 then
       Begin                             { Gestion des erreurs }
       MessageErreur := TraduireMemoire(RecupMessageErrorAbsence(Error)+ComplMessage);
       End
  else
       Begin
       ComplMessage := '';
       Result := TOB.Create('ABSENCESALARIE',nil,-1);
       Result.AddChampSupValeur('ERROR',MessageErreur);//EPZ
       { Création de la tob absence au format de la Table }
ddwriteln('InitialiseTobAbsMotif');
       InitialiseTobAbsMotif(Result,T_Param,T_Sal,'ATT','-','SAL','VIG');
ddwriteln('InitialiseTobAbsMotif done');
       { Contrôle la possibilité de l'enregistrement de l'absence }
ddwriteln('PgControleTobAbsSecond');
       PgControleTobAbsSecond(Result,T_Sal,'E','CREATION','SAL','SAL',Error,NomControl,ComplMessage);
ddwriteln('PgControleTobAbsSecond done');
       if Error <> 0 then                { Gestion des erreurs }
         MessageErreur := TraduireMemoire(RecupMessageErrorAbsence(Error)+ComplMessage)
       else
         Begin
         { Création de l'absence }
ddwriteln('InsertOrUpdateDB');
         Result.InsertOrUpdateDB;
ddwriteln('InsertOrUpdateDB done');
         { Calcul du récapitualtif et maj }
ddwriteln('PGExeCalculRecapAbsEnCours');
         PGExeCalculRecapAbsEnCours(TAbs.GetValue('SALARIE'));
ddwriteln('PGExeCalculRecapAbsEnCours done');
         // ****A faire plus tard PgAffectRecapitulatif;
         End;
       End;
ddwriteln('PgValidAbsUser MessageErreur='+MessageErreur);

  if MessageErreur <> '' then
    Begin
    IF Assigned(Result) then FreeAndNil(Result);
    Result := TOB.Create('Error',nil,-1);
    Result.AddChampSupValeur('ERROR',MessageErreur);
    End;

  IF Assigned(T_Param)       then FreeAndNil(T_Param);
  IF Assigned(T_Sal)         then FreeAndNil(T_Sal);

End;

function PgRecupInfoUserAndResp (UserSal : String ) : Tob;
Var
Resp,NomResp,MailResp,MailSal : String;
OkMailEnvoie : Boolean;
MailDate     : TDateTime;
Begin
  if UserSal='' then
  begin
    UserSal := V_PGI.USerSalarie;
  end;
ddwriteln('UserSal='+UserSal);
  ChargeInfoRespSalMail(UserSal,'SAL','SAL',Resp,NomResp,MailResp,MailSal,OkMailEnvoie,MailDate);
  Result := Tob.Create('Les infos mails',nil,-1);
  Result.AddChampSupValeur('RESP'      ,Resp);
  Result.AddChampSupValeur('NOMRESP'   ,NomResp);
  Result.AddChampSupValeur('MAILRESP'  ,MailResp);
  Result.AddChampSupValeur('SALARIE'   ,UserSal);
  Result.AddChampSupValeur('NOMSAL'    ,RechDom('PGSALARIE',UserSal,False));
  Result.AddChampSupValeur('MAILSAL'   ,MailSal);
  Result.AddChampSupValeur('MAILENVOYE',OkMailEnvoie);
  Result.AddChampSupValeur('MAILDATE'  ,OkMailEnvoie);
End;

function PgValidMailIndicUser(UserResp : string) : tob;
Begin
PgMajIndicMailSalENV(UserResp);
Result := nil;
End;


function PgRecupBodyMailUser(UserResp : string; TAbs : Tob) : tob;
Var
  SL : TStringList;
  S  : String;
  i : integer;
Begin
Result := nil;
SL := PgRecupBodyMailCollaborateur(TAbs,UserResp);
if Assigned(SL) then
   Begin
   if SL.Count > 0 then
     Begin
     Result := TOB.Create('Le mail',nil,-1);
     For i:=0 to SL.Count-1 do
       S:= SL.Strings[i] + ' ';
     Result.AddChampSupValeur('TEXTE',S);
     End;
   End;
End;

{-----Chargement de la combo-box des stages --------------------------------------------}
procedure ChargeStages (CurrentObject : TVignettePlugin; Combo : String; ClauseWhere : String = '');
var
    TobStage,T,DataTob : Tob;
    StSQL              : String;
    i                  : Integer;
begin
    Try
        // Chargement des données
        StSQL := 'SELECT PST_CODESTAGE,PST_LIBELLE,PST_LIBELLE1 FROM STAGE WHERE '+
                 'PST_MILLESIME="0000" AND PST_ACTIF="X" AND PST_NATUREFORM<>"004" ';

        If PGBundleCatalogue Then //PT5 //PT6
        Begin
            //StSQL := StSQL + ' AND NOT (PST_PREDEFINI="DOS" AND PST_NODOSSIER<>"'+V_PGI.NoDossier+'") ';
            If PGDroitMultiForm Then
                StSQL := StSQL + DossiersAInterroger('','','PST',True,True)
            Else
                StSQL := StSQL + DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True);
        End;

        StSQL := StSQL + ClauseWhere;
        StSQL := StSQL + ' ORDER BY PST_LIBELLE';

        DataTob := TOB.Create('~TEMP', nil, -1);

        Try
        	DataTob := CurrentObject.OpenSelectInCache (StSQL);
        Except
        End;
        CurrentObject.ConvertFieldValue(DataTob);

        // Alimentation de la combo
        TobStage := Tob.Create('LesStages',Nil,-1);
        For i := 0 To (DataTob.Detail.Count -1) Do
        begin
            T := Tob.Create('£DETAIL', TobStage, -1);
            T.AddChampSupValeur('VALUE', DataTob.Detail[i].GetValue('PST_CODESTAGE'));
            T.AddChampSupValeur('ITEM', DataTob.Detail[i].GetValue('PST_LIBELLE') + ' ' + DataTob.Detail[i].GetValue('PST_LIBELLE1'));
        end;
        CurrentObject.SetComboDetail(Combo, TobStage);
    Finally
        If Assigned (DataTob)  Then FreeAndNil(DataTob);
        If Assigned (TobStage) Then FreeAndNil(TobStage);
    End;
end;

{-----Adaptation d'une requête au type de responsable ----------------------------------}
{ Permet de construire la clause where à partir de DEPORTSAL afin d'obtenir la liste des éléments
  restreints à l'équipe du responsable. Fonctionne également avec le SECRETAIRE et l'ASSISTANT.
  Si ServiceSup est à True, on récupère les données des services sous-jacents. }
//PT4
Function AdaptByTypeResp (TypeResponsable, Responsable : String; ServiceSup : Boolean; Prefixe : String='PSE'; NonSortis : Boolean = True) : String; //PT8
var Where,Suffixe : String;
    Copie, TypeRespTemp  : String; //PT10
begin
	//PT10
    Copie        := TypeResponsable;
	TypeRespTemp := ReadTokenSt(Copie);

	While TypeRespTemp <> '' Do //PT10
	Begin
		If Where <> '' Then Where := Where + ' OR '; //PT10

	    If (Pos('ADJOINT', TypeRespTemp) > 0) Or (Pos('SECRETAIRE', TypeRespTemp) > 0) Then
	    Begin
	        Suffixe := RightStr(TypeRespTemp, 3);

	        If Not ServiceSup Then
	            Where := Where + '('+Prefixe+'_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_'+TypeRespTemp+'="'+Responsable+'")))' //PT13
	        Else
	            Where := Where + '('+Prefixe+'_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_CODESERVICE IN '+ //PT13
	                     '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
	                     'WHERE PGS_'+TypeRespTemp+'="'+Responsable+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_'+TypeRespTemp+'="'+Responsable+'")))';
	    End
	    Else
	    Begin
	        If Not ServiceSup Then
	            Where := Where + Prefixe+'_'+TypeRespTemp+'="'+Responsable+'"'
	        Else
	            Where := Where + '('+Prefixe+'_'+TypeRespTemp+'="'+Responsable+'" OR '+Prefixe+'_'+TypeRespTemp+' IN '+
	                     '(SELECT PGS_'+TypeRespTemp+' FROM SERVICES WHERE PGS_CODESERVICE IN '+
	                     '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
	                     ' PGS_'+TypeRespTemp+'="'+Responsable+'")))';
	    End;

        TypeRespTemp := ReadTokenSt(Copie);
	End;

	If Where <> '' Then Where := '(' + Where + ')'; //PT10

	// Salariés non sortis uniquement //PT15
	If (Where <> '') And (NonSortis) Then
	Begin
		If Prefixe = 'PSA' Then
			Where := Where + ' AND ' + SalariesNonSortis
		Else
			// On prend aussi les salariés sans nom pour les inscriptions au prévisionnel non nominatives
			Where := Where + ' AND (' + Prefixe+'_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+SalariesNonSortis+') OR '+Prefixe+'_SALARIE = "") ';
	End;

    Result := Where;
End;

{-----Chargement de la combo-box des salariés rattachés au responsable -----------------}
{ Le type de responsable doit être le suffixe de la colonne DEPORTSAL : RESPONSFOR, RESPONSABS, etc.
  Si le type passé en paramètre commence par "ASSIST" ou "SECRETAIRE", la fonction renvoie les salariés rattachés au
  responsable correspondant à l'assistant/secrétaire. }
procedure ChargeSalaries (CurrentObject : TVignettePlugin; Combo, TypeResponsable : String; InclureIndefini : Boolean = False; Responsable : String = ''; ServiceSup : Boolean = False);   //PT1
var TobSalaries,TS,DataTob : Tob;
    i                      : Integer;
    StSQL				   : String;
begin
    Try
        If Responsable = '' Then Responsable := V_PGI.UserSalarie;

        DataTob := TOB.Create('~TEMP', nil, -1);
        StSQL   := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE '+
                   'WHERE '+AdaptByTypeResp(TypeResponsable, Responsable, ServiceSup)+' ORDER BY PSA_LIBELLE,PSA_PRENOM'; //PT4 //PT15
		StSQL := AdaptMultiDosForm (StSQL); //PT6

		Try
        	DataTob := CurrentObject.OpenSelectInCache (StSQL);
        Except
    	End;
        CurrentObject.ConvertFieldValue(DataTob);

        TobSalaries := Tob.Create('LesSalaries',Nil,-1);
        For i := 0 To (DataTob.Detail.Count -1) Do
        begin
            TS := Tob.Create('~Salarie',TobSalaries,-1);
            TS.AddChampSupValeur('VALUE',DataTob.Detail[i].GetValue('PSA_SALARIE'));
            TS.AddChampSupValeur('ITEM', DataTob.Detail[i].GetValue('PSA_LIBELLE') + ' '+DataTob.Detail[i].GetValue('PSA_PRENOM'));
        end;

        //PT1 - Début
        If InclureIndefini Then
        Begin
            TS := Tob.Create('~Salarie',TobSalaries, 0);
            TS.AddChampSupValeur('VALUE',UNDEFINED);
            TS.AddChampSupValeur('ITEM', '<<Non nominatif>>');
        End;
        //PT1 - Fin

        CurrentObject.SetComboDetail(Combo, TobSalaries);
    Finally
        If Assigned (DataTob)  Then FreeAndNil(DataTob);
        If Assigned (TobSalaries) Then FreeAndNil(TobSalaries);
    End;
end;

{-----Chargement de la combo-box des emplois des salariés rattachés au responsable -----------------}
{ Le type de responsable doit être le suffixe de la colonne DEPORTSAL : RESPONSFOR, RESPONSABS, etc.
  Si le type passé en paramètre commence par "ASSIST" ou "SECRETAIRE", la fonction renvoie les salariés rattachés au
  responsable correspondant à l'assistant/secrétaire. }
//PT1
procedure ChargeEmplois (CurrentObject : TVignettePlugin; Combo, TypeResponsable : String; Responsable : String = ''; ServiceSup : Boolean = False);
var TobEmplois,TS,DataTob : Tob;
    i                     : Integer;
    StSQL				  : String;
    Q                     : Tquery;
begin
    Try
        If Responsable = '' Then Responsable := V_PGI.UserSalarie;

        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT DISTINCT PSA_LIBELLEEMPLOI,CC_LIBELLE FROM SALARIES '+
                    'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
                    'LEFT JOIN ';
        //PT17
        If PGBundleHierarchie Then
            StSQL := StSQL + GetBase(GetBasePartage(BUNDLE_HIERARCHIE),'CHOIXCOD')
        Else
            StSQL := StSQL + 'CHOIXCOD';
        StSQL := StSQL + ' ON CC_CODE=PSA_LIBELLEEMPLOI '+
                    'WHERE CC_TYPE="PLE" AND '+AdaptByTypeResp(TypeResponsable, Responsable, ServiceSup)+ //PT4
                    ' ORDER BY PSA_LIBELLEEMPLOI';
        StSQL := AdaptMultiDosForm (StSQL); //PT6

        Try
            //PT17 Lancement de la requête normalement car l'ouverture en cache ne comprend pas les redirections
            // éventuelles dues au partage de données
        	//DataTob := CurrentObject.OpenSelectInCache (StSQL);
            //CurrentObject.ConvertFieldValue(DataTob);
            Q := OpenSQL(StSQL, True);
            DataTob.LoadDetailDB('$DATA','','',Q,False);
            Ferme(Q);
		Except
		End;

        TobEmplois := Tob.Create('LesEmplois',Nil,-1);
        For i := 0 To (DataTob.Detail.Count -1) Do
        begin
            TS := Tob.Create('~Emploi',TobEmplois,-1);
            TS.AddChampSupValeur('VALUE',DataTob.Detail[i].GetValue('PSA_LIBELLEEMPLOI'));
            TS.AddChampSupValeur('ITEM', DataTob.Detail[i].GetValue('CC_LIBELLE'));
        end;

        CurrentObject.SetComboDetail(Combo, TobEmplois);
    Finally
        If Assigned (DataTob)  Then FreeAndNil(DataTob);
        If Assigned (TobEmplois) Then FreeAndNil(TobEmplois);
    End;
end;

{-----Validation d'une inscription à une session ---------------------------------------------}
//PT1
Function InscritSalarieSession (Stage, Session, MillesimeEC, Salarie, TypePlan : String; TypeResponsable : String) : String; //PT9
var
   TForm,T_Sal              : Tob;
   Error                    : Integer;
   ComplMessage             : String;
   Q                        : TQuery;
   DD,DF                    : TDateTime;
   TauxChargeC, TauxChargeNC: Double;
   TauxHoraire,Montant      : Double;
   Prefixe,LeDossier        : String;
   Responsable 				: String;
Begin
    Result := '';

    // Contrôle de cohérence
    Try
	    if ExisteSQL('SELECT 1 FROM FORMATIONS WHERE PFO_SALARIE="' + Salarie + '" AND ' +
	                 'PFO_CODESTAGE="' + Stage + '" AND PFO_MILLESIME="' + MillesimeEC + '" AND PFO_ORDRE="'+Session+'"') Then
	    begin
	        Result := TraduireMemoire('Il existe déjà une inscription pour le salarié à cette formation.');
	        Exit;
	    end;
	Except
    	Result := TraduireMemoire(ERR_MSG) + '#UTIL_SESS_1';
    	Exit;
    End;

    { Chargement des infos salariés }
    If PGBundleHierarchie Then //PT6
    Begin
    	//T_Sal := TOB.Create('SAL_ARIE', Nil, -1);
    	//T_Sal.LoadDetailDBFromSQL('INTERT_Sal := ChargeTobInterimaire(Salarie);IMAIRES','SELECT * FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+Salarie+'"');
    	T_Sal := ChargeTobInterimaire(Salarie);
    End
    Else
    	T_Sal := ChargeTobSalarie(Salarie);

    { Contrôle la conformité de la saisie d'absence }
    Error := 0;
    if Error <> 0 then
    Begin                             { Gestion des erreurs }
        Result := TraduireMemoire(RecupMessageErrorAbsence(Error)+ComplMessage);
    End
    else
    Begin
        ComplMessage := '';

        TForm := TOB.Create('FORMATIONS',Nil,-1);
        TForm.PutValue('PFO_SALARIE',       Salarie);
        TForm.PutValue('PFO_MILLESIME',     MillesimeEC);
        TForm.PutValue('PFO_TYPEPLANPREV',  TypePlan);
        TForm.PutValue('PFO_CODESTAGE',     Stage);
        TForm.PutValue('PFO_ORDRE',         Session);

        TForm.PutValue('PFO_TYPOFORMATION',	'001');
        TForm.PutValue('PFO_ORGCOLLECT',    -1);
        TForm.PutValue('PFO_DATECOINVES',	IDate1900);
        TForm.PutValue('PFO_DATEACCEPT',	IDate1900);
        TForm.PutValue('PFO_REFUSELE',		IDate1900);
        TForm.PutValue('PFO_REPORTELE',		IDate1900);
        //PT7 - Début
        TForm.PutValue('PFO_EFFECTUE',		'X');
//        If GetParamSocSecur('SO_PGFORMVALIDREA', False) Then
//            TForm.PutValue('PFO_ETATINSCFOR',   'ATT')
//        Else
            TForm.PutValue('PFO_ETATINSCFOR',   'VAL');
        //PT7 - Fin

        //Recup session
        Try
	        Q := OpenSQL('SELECT * FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Stage+'"' +
	        			 'AND PSS_ORDRE='+Session+' '+
	        			 'AND PSS_MILLESIME="'+MillesimeEC+'"', True);
	        If not Q.Eof then
	        begin
	            TForm.PutValue('PFO_DATEDEBUT',     Q.FindField('PSS_DATEDEBUT').AsDateTime);
	            TForm.PutValue('PFO_DATEFIN',       Q.FindField('PSS_DATEFIN').AsDateTime);
	            TForm.PutValue('PFO_NATUREFORM',    Q.FindField('PSS_NATUREFORM').AsString);
	            TForm.PutValue('PFO_LIEUFORM',      Q.FindField('PSS_LIEUFORM').AsString);
	            TForm.PutValue('PFO_CENTREFORMGU',  Q.FindField('PSS_CENTREFORMGU').AsString);
	            TForm.PutValue('PFO_DEBUTDJ',       Q.FindField('PSS_DEBUTDJ').AsString);
	            TForm.PutValue('PFO_FINDJ',         Q.FindField('PSS_FINDJ').AsString);
	            TForm.PutValue('PFO_ORGCOLLECTGU',  Q.FindField('PSS_ORGCOLLECTSGU').AsString);
	            TForm.PutValue('PFO_HEUREDEBUT',    Q.FindField('PSS_HEUREDEBUT').AsDateTime);
	            TForm.PutValue('PFO_HEUREFIN',      Q.FindField('PSS_HEUREFIN').AsDateTime);
	            TForm.PutValue('PFO_FORMATION1',    Q.FindField('PSS_FORMATION1').AsString);
	            TForm.PutValue('PFO_FORMATION2',    Q.FindField('PSS_FORMATION2').AsString);
	            TForm.PutValue('PFO_FORMATION3',    Q.FindField('PSS_FORMATION3').AsString);
	            TForm.PutValue('PFO_FORMATION4',    Q.FindField('PSS_FORMATION4').AsString);
	            TForm.PutValue('PFO_FORMATION5',    Q.FindField('PSS_FORMATION5').AsString);
	            TForm.PutValue('PFO_FORMATION6',    Q.FindField('PSS_FORMATION6').AsString);
	            TForm.PutValue('PFO_FORMATION7',    Q.FindField('PSS_FORMATION7').AsString);
	            TForm.PutValue('PFO_FORMATION8',    Q.FindField('PSS_FORMATION8').AsString);
	            TForm.PutValue('PFO_INCLUSDECL',    Q.FindField('PSS_INCLUSDECL').AsString);
	            TForm.PutValue('PFO_NBREHEURE',     Q.FindField('PSS_DUREESTAGE').AsFloat);
	            if (GetParamSocSecur('SO_PGDIFTPSTRAV',False)) or (TypePlan = 'PLF') Then
	            begin
	                 TForm.PutValue('PFO_HTPSTRAV', Q.FindField('PSS_DUREESTAGE').AsFloat);
	                 TForm.PutValue('PFO_HTPSNONTRAV',0);
	            end
	            else
	            begin
	                 TForm.PutValue('PFO_HTPSTRAV',0);
	                 TForm.PutValue('PFO_HTPSNONTRAV',Q.FindField('PSS_DUREESTAGE').AsFloat);
	            end;
	            TForm.PutValue('PFO_NOSESSIONMULTI', Q.FindField('PSS_NOSESSIONMULTI').AsString);
	            TForm.PutValue('PFO_IDSESSIONFOR',   Q.FindField('PSS_IDSESSIONFOR').AsString);
	            TForm.PutValue('PFO_PGTYPESESSION',  Q.FindField('PSS_PGTYPESESSION').AsString);
	        end;
	        Ferme(Q);
	    Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_SESS_2';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;

        //Recup salarié
        If PGBundleHierarchie Then Prefixe := 'PSI' Else Prefixe := 'PSA'; //PT6
        TForm.PutValue('PFO_TRAVAILN1',     T_Sal.Getvalue(Prefixe+'_TRAVAILN1'));
        TForm.PutValue('PFO_TRAVAILN2',     T_Sal.Getvalue(Prefixe+'_TRAVAILN2'));
        TForm.PutValue('PFO_TRAVAILN3',     T_Sal.Getvalue(Prefixe+'_TRAVAILN3'));
        TForm.PutValue('PFO_TRAVAILN4',     T_Sal.Getvalue(Prefixe+'_TRAVAILN4'));
        TForm.PutValue('PFO_CODESTAT',      T_Sal.Getvalue(Prefixe+'_CODESTAT'));
        TForm.PutValue('PFO_NOMSALARIE',    T_Sal.Getvalue(Prefixe+'_LIBELLE'));
        TForm.PutValue('PFO_PRENOM',        T_Sal.Getvalue(Prefixe+'_PRENOM'));
        TForm.PutValue('PFO_ETABLISSEMENT', T_Sal.Getvalue(Prefixe+'_ETABLISSEMENT'));
        TForm.PutValue('PFO_LIBEMPLOIFOR',  T_Sal.Getvalue(Prefixe+'_LIBELLEEMPLOI'));
        TForm.PutValue('PFO_LIBREPCMB1',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB1'));
        TForm.PutValue('PFO_LIBREPCMB2',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB2'));
        TForm.PutValue('PFO_LIBREPCMB3',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB3'));
        TForm.PutValue('PFO_LIBREPCMB4',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB4'));
		//PT6 - Début
        //TForm.PutValue('PFO_NODOSSIER',     V_PGI.NoDossier);
        TForm.PutValue('PFO_NODOSSIER',     GetNoDossierSalarie(Salarie));
        If V_PGI.NoDossier <> '000000' Then  //PT5
            TForm.PutValue('PFO_PREDEFINI',     'DOS')
        Else
            TForm.PutValue('PFO_PREDEFINI',     'STD');
        //PT6 - Fin

        //PT9 - Début
        If TypeResponsable = 'RESPONSFOR' Then
        	Responsable := V_PGI.UserSalarie
        Else
        Begin
        	Try
        		Q := OpenSQL('SELECT PSE_RESPONSFOR FROM DEPORTSAL WHERE PSE_SALARIE="'+Salarie+'"', True);
        		If Not Q.EOF Then Responsable := Q.FindField('PSE_RESPONSFOR').AsString;
        		Ferme(Q);
            Except
            	Result := TraduireMemoire(ERR_MSG) + '#UTIL_SESS_22';
	        	If Assigned(TForm) then FreeAndNil(TForm);
	        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
	        	Exit;
            End;
        End;
        TForm.PutValue('PFO_RESPONSFOR', Responsable);
        //PT9 - Fin

        { Calcul des coûts salarié }
        MillesimeEC  := RendMillesimeRealise (DD,DF);
        TauxChargeC  := 1;
        TauxChargeNC := 1;

        Try
	        Q := OpenSQL('SELECT PFE_TAUXCHARGENC,PFE_TAUXCHARGEC FROM EXERFORMATION'
	                   +' WHERE PFE_MILLESIME="'+MillesimeEC+'"',True);
	        If not Q.EOF Then
	        Begin
	            TauxChargeC  := Q.FindField('PFE_TAUXCHARGEC').AsFloat;
	            TauxChargeNC := Q.FindField('PFE_TAUXCHARGENC').AsFloat;
	        End;
	        Ferme(Q);
	    Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_SESS_3';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;

        If PGBundleHierarchie Then //PT6
            LeDossier := GetNoDossierSalarie(Salarie)
        Else
            LeDossier := '';
	    if GetParamSoc('SO_PGFORVALOSALAIRE') = 'VCR' then
	    	TauxHoraire := ForTauxHoraireReel(Salarie,0,0,'',ValReel,TForm.GetValue('PFO_DATEFIN'), LeDossier)  //PT6
	    Else
	    	TauxHoraire := ForTauxHoraireCategoriel(TForm.GetValue('PFO_LIBEMPLOIFOR'),MillesimeEC);

	    Montant := TauxHoraire*TForm.GetValue('PFO_HTPSTRAV');

	    Try
		    Q := OpenSQL('SELECT PSA_DADSCAT FROM '+GetBase(GetBaseSalarie(Salarie),'SALARIES')+' WHERE PSA_SALARIE="'+Salarie+'"',True); //PT6
		    If Not Q.EOF Then
		    Begin
			    If (Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02') then
			    	Montant := Arrondi(Montant * TauxChargeC,2)
			    else
			    	Montant := Arrondi(Montant * TauxChargeNC,2);
			    Ferme(Q);
			End;
	    Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_SESS_4';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;
	    TForm.PutValue('PFO_COUTREELSAL',Montant);

	    { Calcul du coût de la session }
	    Try
			CalcCtInvestSession('SALAIRE',Stage,MillesimeEC,StrToInt(Session));
		Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_SESS_5';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;

        TForm.SetAllModifie(True);

        Try
        	TForm.InsertOrUpdateDB;
	    Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_SESS_6';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;
    End;

    If Assigned(TForm)   Then FreeAndNil(TForm);
    If Assigned(T_Sal)   Then FreeAndNil(T_Sal);
end;

{-----Validation d'une inscription au prévisionnel -------------------------------------------}
//PT1
Function InscritSalariePrevisionnel (Stage, MillesimeEC, Salarie, NbInsc, TypePlan, LibEmploi : String; HeuresTT, HeuresHT : Double; DemandeSalarie : Boolean; TypeResponsable : String) : String; //PT9
var
   TForm,T_Sal                     : Tob;
   Rang,Error                      : Integer;
   Q                               : TQuery;
   TauxChargeBudget,Salaire,NbJours: Double;
   FraisH,FraisR,FraisT,TotalFrais : Double;
   Req,Etab,ComplMessage,StSQL     : String;
   TitreMail,Prefixe,LeDossier     : String;
   TexteMail					   : HTStrings;
   Responsable 					   : String;
Begin
    Result := '';

	If PGBundleHierarchie Then Prefixe := 'PSI' Else Prefixe := 'PSA'; //PT6

    // Contrôle de cohérence
    If Salarie <> '' Then
    Begin
    	Try
	        if ExisteSQL('SELECT 1 FROM INSCFORMATION WHERE PFI_SALARIE="' + Salarie + '" AND ' +
	                 'PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + MillesimeEC + '"') Then
	        begin
	            If DemandeSalarie Then
	                Result := TraduireMemoire('Vous êtes déjà inscrit à cette formation.')
	            Else
	                Result := TraduireMemoire('Il existe déjà une inscription pour le salarié à cette formation.');
	            Exit;
	        end;
	    Except
	    	Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_1';
	    	Exit;
		End;

        { Chargement des infos salariés }
	    If PGBundleHierarchie Then //PT6
	    Begin
	    	T_Sal := ChargeTobInterimaire(Salarie);
	    End
	    Else
	    	T_Sal := ChargeTobSalarie(Salarie);
    End
    Else
    Begin
        // Cas d'une inscription non nominative
        Etab := '';
        StSQL := 'SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="'+V_PGI.UserSalarie+'"';
        StSQL := AdaptMultiDosForm (StSQL); //PT6
        Try
	        Q := OpenSQL(StSQL, True);
	        If not Q.Eof then Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
	    Finally
    		Ferme(Q);
        End;
    End;

    // Si le stage n'est pas encore au prévisionnel, on le crée
    CreerStageAuPrevisionnel(Stage, MillesimeEC);

    { Contrôle la conformité de la saisie d'absence }
    Error := 0;
    if Error <> 0 then
    Begin                             { Gestion des erreurs }
        Result := TraduireMemoire(RecupMessageErrorAbsence(Error)+ComplMessage);
    End
    else
    Begin
        ComplMessage := '';

        TForm := TOB.Create('INSCFORMATION',Nil,-1);
        TForm.PutValue('PFI_SALARIE',      Salarie);
        TForm.PutValue('PFI_MILLESIME',    MillesimeEC);
        TForm.PutValue('PFI_TYPEPLANPREV', TypePlan);
        TForm.PutValue('PFI_CODESTAGE',    Stage);
        If Salarie <> '' Then
            Req := 'PFI_ETABLISSEMENT="' +T_Sal.GetValue(Prefixe+'_ETABLISSEMENT') + '"' //PT6
        Else
            Req := 'PFI_ETABLISSEMENT="' + Etab + '"';

        Try
        	Q := OpenSQL('SELECT MAX (PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_MILLESIME="' + MillesimeEC + '" AND '+Req, True);
        	if not Q.eof then Rang := Q.FindField('RANG').AsInteger + 1
        	else Rang := 1;
        	Ferme(Q);
        Except
			Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_2';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
	    End;
        TForm.PutValue('PFI_RANG',Rang);

        //Recup session
        Try
	        Q := OpenSQL('SELECT * FROM STAGE WHERE PST_CODESTAGE="'+Stage+'" AND PST_MILLESIME='+MillesimeEC+'',True);
	        If not Q.Eof then
	        begin
	            TForm.PutValue('PFI_MILLESIME',Q.FindField('PST_MILLESIME').AsString);
	            TForm.PutValue('PFI_NATUREFORM',Q.FindField('PST_NATUREFORM').AsString);
	            TForm.PutValue('PFI_LIEUFORM',Q.FindField('PST_LIEUFORM').AsString);
	            TForm.PutValue('PFI_CENTREFORMGU',Q.FindField('PST_CENTREFORMGU').AsString);//DB2
	            TForm.PutValue('PFI_ORGCOLLECTGU',Q.FindField('PST_ORGCOLLECTSGU').AsString);
	            TForm.PutValue('PFI_FORMATION1',Q.FindField('PST_FORMATION1').AsString);
	            TForm.PutValue('PFI_FORMATION2',Q.FindField('PST_FORMATION2').AsString);
	            TForm.PutValue('PFI_FORMATION3',Q.FindField('PST_FORMATION3').AsString);
	            TForm.PutValue('PFI_FORMATION4',Q.FindField('PST_FORMATION4').AsString);
	            TForm.PutValue('PFI_FORMATION5',Q.FindField('PST_FORMATION5').AsString);
	            TForm.PutValue('PFI_FORMATION6',Q.FindField('PST_FORMATION6').AsString);
	            TForm.PutValue('PFI_FORMATION7',Q.FindField('PST_FORMATION7').AsString);
	            TForm.PutValue('PFI_FORMATION8',Q.FindField('PST_FORMATION8').AsString);
	            TForm.PutValue('PFI_INCLUSDECL',Q.FindField('PST_INCLUSDECL').AsString);
	            TForm.PutValue('PFI_JOURSTAGE',Q.FindField('PST_JOURSTAGE').AsString);
	            TForm.PutValue('PFI_NBANIM',Q.FindField('PST_NBANIM').AsString);
	        end;
	        Ferme(Q);
	    Except
			Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_3';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
		End;

        TForm.PutValue('PFI_HTPSTRAV',      HeuresTT);
        TForm.PutValue('PFI_HTPSNONTRAV',   HeuresHT);
        TForm.PutValue('PFI_DUREESTAGE',    HeuresTT + HeuresHT);
        // Etat de l'inscription : suivant paramsoc pour le manager, forcée à ATT si demande salarié
        If Not DemandeSalarie Then
        Begin
            If GetParamSocSecur('SO_PGFORMVALIDPREV', False) Then
                TForm.PutValue('PFI_ETATINSCFOR',   'ATT')
            Else
                TForm.PutValue('PFI_ETATINSCFOR',   'VAL');
        End
        Else
                TForm.PutValue('PFI_ETATINSCFOR',   'ATT');
        TForm.PutValue('PFI_DATEDIF',       V_PGI.DateEntree);
        TForm.PutValue('PFI_NIVPRIORITE',   '01');
        If DemandeSalarie Then
            TForm.PutValue('PFI_MOTIFINSCFOR',  '001')
        Else
            TForm.PutValue('PFI_MOTIFINSCFOR',  '002');
        TForm.PutValue('PFI_TYPOFORMATION', '001');
        TForm.PutValue('PFI_DATEACCEPT',    IDate1900);
        TForm.PutValue('PFI_REFUSELE',      IDate1900);
        TForm.PutValue('PFI_REPORTELE',     IDate1900);
        TForm.PutValue('PFI_REALISE',       '-');
		//PT6 - Début
        //TForm.PutValue('PFI_NODOSSIER',     V_PGI.NoDossier);
        TForm.PutValue('PFI_NODOSSIER',     GetNoDossierSalarie(Salarie));
        If V_PGI.NoDossier <> '000000' Then //PT5
            TForm.PutValue('PFI_PREDEFINI',     'DOS')
        Else
            TForm.PutValue('PFI_PREDEFINI',     'STD');
        //PT6 - Fin

        //Recup salarie
        If Salarie <> '' Then
        Begin
            TForm.PutValue('PFI_LIBELLE',       T_Sal.Getvalue(Prefixe+'_LIBELLE')+' '+T_Sal.Getvalue(Prefixe+'_PRENOM'));
            TForm.PutValue('PFI_TRAVAILN1',     T_Sal.Getvalue(Prefixe+'_TRAVAILN1'));
            TForm.PutValue('PFI_TRAVAILN2',     T_Sal.Getvalue(Prefixe+'_TRAVAILN2'));
            TForm.PutValue('PFI_TRAVAILN3',     T_Sal.Getvalue(Prefixe+'_TRAVAILN3'));
            TForm.PutValue('PFI_TRAVAILN4',     T_Sal.Getvalue(Prefixe+'_TRAVAILN4'));
            TForm.PutValue('PFI_CODESTAT',      T_Sal.Getvalue(Prefixe+'_CODESTAT'));
            TForm.PutValue('PFI_DADSCAT',       T_Sal.Getvalue(Prefixe+'_DADSCAT'));
            TForm.PutValue('PFI_ETABLISSEMENT', T_Sal.Getvalue(Prefixe+'_ETABLISSEMENT'));
            TForm.PutValue('PFI_LIBEMPLOIFOR',  T_Sal.Getvalue(Prefixe+'_LIBELLEEMPLOI'));
            TForm.PutValue('PFI_LIBREPCMB1',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB1'));
            TForm.PutValue('PFI_LIBREPCMB2',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB2'));
            TForm.PutValue('PFI_LIBREPCMB3',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB3'));
            TForm.PutValue('PFI_LIBREPCMB4',    T_Sal.Getvalue(Prefixe+'_LIBREPCMB4'));
        End
        Else
        Begin
            If (NbInsc = '1') Then
                TForm.PutValue('PFI_LIBELLE',       NbInsc+' salarié')
            Else
                TForm.PutValue('PFI_LIBELLE',       NbInsc+' salariés');
            TForm.PutValue('PFI_ETABLISSEMENT', Etab);
            TForm.PutValue('PFI_LIBEMPLOIFOR',  LibEmploi);
        End;

        If (Not DemandeSalarie) And ((TypeResponsable = 'RESPONSFOR') Or (Salarie='')) Then //PT9
            Responsable := V_PGI.UserSalarie
        Else
        Begin
        	Try
            	Q := OpenSQL('SELECT PSE_RESPONSFOR FROM DEPORTSAL WHERE PSE_SALARIE="'+Salarie+'"', True);
            	If Not Q.EOF Then Responsable := Q.FindField('PSE_RESPONSFOR').AsString;
            	Ferme(Q);
            Except
            	Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_4';
	        	If Assigned(TForm) then FreeAndNil(TForm);
	        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
	        	Exit;
            End;
        End;
        TForm.PutValue('PFI_RESPONSFOR', Responsable);

        TForm.PutValue('PFI_NBINSC',        NbInsc);

        { Coût salarié }
        Try
        	Q := OpenSQL('SELECT PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="' + MillesimeEC + '"', True);
        	if not Q.Eof then TauxChargeBudget := Q.FindField('PFE_TAUXBUDGET').AsFloat
        	else TauxChargeBudget := 1;
        	Ferme(Q);
        Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_5';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;

        If (Salarie <> '') Then
        Begin
            If PGBundleHierarchie Then //PT6
                LeDossier := GetNoDossierSalarie(Salarie)
            Else
                LeDossier := '';
            If (GetParamSoc('SO_PGFORVALOSALAIRE') = 'VCR') then
                Salaire := ForTauxHoraireReel(Salarie,0,0,'',False,0,LeDossier)
            else
                Salaire := ForTauxHoraireCategoriel(TForm.GetValue('PFI_LIBEMPLOIFOR'), MillesimeEC);
        End
        Else
        Begin
            Salaire := ForTauxHoraireCategoriel(LibEmploi, MillesimeEC);
        End;
        TForm.PutValue('PFI_COUTREELSAL', Salaire * (HeuresTT+HeuresHT) * TauxChargeBudget);

        { Frais forfaitaires }
        Req := 'SELECT PFF_FRAISHEBERG,PFF_FRAISREPAS,PFF_FRAISTRANSP FROM FORFAITFORM ' +
              'WHERE PFF_MILLESIME="' + MillesimeEC + '" AND PFF_LIEUFORM="' + TForm.Getvalue('PFI_LIEUFORM') + '"';
        If GetParamSocSecur('SO_PGFORPOPFRAIS', False) Then
        Begin
            If Salarie <> '' Then
            Begin
                Req := Req + ' AND PFF_POPULATION IN (SELECT PNA_POPULATION FROM SALARIEPOPUL WHERE PNA_SALARIE="'+Salarie+'")';
            End
            Else
            Begin
                Req := Req + ' AND PFF_POPULATION="'+RecherchePopulation(TForm)+'"';
            End;
        End
        Else
            Req := Req + ' AND PFF_ETABLISSEMENT="' + TForm.Getvalue('PFI_ETABLISSEMENT') + '"';

		Try
        	Q := OpenSQL(Req, True);
	        FraisH := 0;
	        FraisR := 0;
	        FraisT := 0;
	        if not Q.eof then
	        begin
	          FraisH := Q.FindField('PFF_FRAISHEBERG').AsFloat;
	          FraisR := Q.FindField('PFF_FRAISREPAS').AsFloat;
	          FraisT := Q.FindField('PFF_FRAISTRANSP').AsFloat;
	        end;
	        Ferme(Q);
	    Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_6';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;

        NbJours := TForm.GetValue('PFI_JOURSTAGE');
        If Salarie = '' then NbJours := NbJours / StrToInt(NbInsc);
        FraisH := FraisH * (NbJours - 1);
        if FraisH < 0 then FraisH := 0;
        If FraisH = 0 Then FraisR := FraisR * NbJours
        Else FraisR := FraisR * ((NbJours * 2) - 1);
        if FraisR < 0 then FraisR := 0;
        TotalFrais := FraisH + FraisR + FraisT;
        If Salarie = '' Then
            TForm.PutValue('PFI_FRAISFORFAIT', TotalFrais * StrToInt(NbInsc))
        Else
            TForm.PutValue('PFI_FRAISFORFAIT', TotalFrais);

        TForm.PutValue('PFI_AUTRECOUT', CalCoutPedagogique(Stage, MillesimeEC, TForm.GetValue('PFI_ETABLISSEMENT'), '1', StrToInt(NbInsc)));

        TForm.SetAllModifie(True);

        Try
        	TForm.InsertOrUpdateDB;
        Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_7';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
    	End;

        {Mise à jour des coûts pédagogiques pour tous les salariés inscrits}
        Try
        	MajCoutsInscPrev (Stage,MillesimeEC);
        Except
        	Result := TraduireMemoire(ERR_MSG) + '#UTIL_INSC_8';
        	If Assigned(TForm) then FreeAndNil(TForm);
        	If Assigned(T_Sal) then FreeAndNil(T_Sal);
        	Exit;
        End;
    End;

    // Une fois l'inscription effectuée, on envoie un mail au responsable
    If (Error = 0) And DemandeSalarie Then
    Begin
    	PrepareMailFormation (Salarie, TForm.GetValue('PFI_RESPONSFOR'), 'SAISIE'+TypePlan, TForm, TitreMail, TexteMail);
    	EnvoieMail (TForm.GetValue('PFI_RESPONSFOR'), TitreMail, TexteMail);
    	If Assigned(TexteMail) Then TexteMail.Free;
    End;

    IF Assigned(TForm) then FreeAndNil(TForm);
    IF Assigned(T_Sal) then FreeAndNil(T_Sal);
End;

{-----Envoie un mail au salarié --------------------------------------------------------------}

Function EnvoieMail (Salarie, Titre : String; Corps : HTStrings) : Boolean;
Var
  sServer : string;
  sFrom   : string;
  sTitre  : string;
  Destinataire : String;
  DestCopie    : String;

    // Sous-fonction pour récupérer le mail du destinataire
    Function GetDestinataire (CodeSalarie : String) : String;
    Var Q : TQuery;
    Begin
    	Result := '';

    	//PT14
    	// On recherche d'abord l'adresse email du salarié au niveau des utilisateurs.
    	Q := OpenSQL('SELECT US_EMAIL FROM UTILISAT WHERE US_AUXILIAIRE="'+CodeSalarie+'"', True);
    	If Not Q.EOF Then Result := Q.FindField('US_EMAIL').AsString;
    	Ferme(Q);

    	// Ensuite si non trouvé, on regarde dans DEPORTSAL
    	If Result = '' Then
    	Begin
			Q := OpenSQL('SELECT PSE_EMAILPROF FROM DEPORTSAL WHERE PSE_SALARIE="'+CodeSalarie+'"',True);
			If Not Q.Eof then Result := Q.FindField('PSE_EMAILPROF').AsString;
			Ferme(Q);
		End;
    End;
Begin
  sServer := '';
  sFrom   := '';
  Result := False;
  If Salarie = '' Then Exit; //PT18

  Try
  with TRegistry.Create do begin
    RootKey := HKEY_LOCAL_MACHINE;
    if OpenKey('SOFTWARE\CEGID_RM\PgiService\Journal', False) then begin
      if ValueExists('smtpServer') then sServer := ReadString('smtpServer');
      if ValueExists('smtpFrom')   then sFrom   := ReadString('smtpFrom');
      CloseKey;
    end;
    Free;
  end;

  Destinataire := GetDestinataire(Salarie);

  if sServer = '' then Exit;
  if sFrom   = '' then Exit;
  if Destinataire = '' then Exit;

  sTitre := Titre;

	//PT?12
  If Pos('Validation',sTitre) > 0 Then // Pas de mise en copie quand simple demande salarié
      If GetParamSocSecur('SO_PGFORMAILCOPIE', False) Then DestCopie := GetParamSocSecur('SO_PGFORMAILADR',''); //PT12
  If DestCopie <> '' Then DestCopie := ';'+DestCopie;

  if SendMailSmtp(sServer, sFrom, Destinataire+DestCopie, sTitre, Corps.Text) then begin
    Result := True;
  end;
  Except
    Result := False;
  End;
End;

//PT2
{-----Charge les données de la tablette adéquate dans la combo -------------------------------}
{ Si Requete est non vide, les données sont chargées par rapport à celle-ci. Sinon, c'est le
  nom de la tablette qui prime.
  Attention : Seules les tablettes basées sur CHOIXCOD et COMMUN sont gérées en automatique. }
procedure ChargeValeursCombo (CurrentObject : TVignettePlugin; Combo, Tablette : String; Requete : String = ''; TousPossible : Boolean = False; TousLibelle : String='<<TOUS>>');
var TobLecture,T         : TOB;
    StSQL                : String;
    Q                    : TQuery;
    Code,Lib,Table,Where : String;
Begin
    If Tablette <> '' Then
    Begin
        Q := OpenSQL('SELECT DO_CODE,DO_CHAMPLIB,DO_PREFIXE,DO_WHERE FROM DECOMBOS WHERE DO_COMBO="'+Tablette+'"', True);
        If Not Q.EOF Then
        Begin
            Code  := Q.FindField('DO_CODE').AsString;
            Lib   := Q.FindField('DO_CHAMPLIB').AsString;
            Table := Q.FindField('DO_PREFIXE').AsString;
            If Table = 'CC' Then Table := 'CHOIXCOD'
            Else If Table = 'CO' Then Table := 'COMMUN';
            Where := Q.FindField('DO_WHERE').AsString;
            Where := StringReplace(Where, '&#@', '', [rfReplaceAll,rfIgnoreCase]);
        End;
        Ferme(Q);
    End;

    If Requete <> '' Then
        StSQL := Requete
    Else
        StSQL := 'SELECT '+Code+' AS VALUE,'+Lib+' AS ITEM FROM '+Table+' WHERE '+Where;

    TobLecture := TOB.Create('LaTablette', Nil, -1);
    TobLecture := CurrentObject.OpenSelectInCache (StSQL);
    CurrentObject.ConvertFieldValue(TobLecture);

    If TousPossible Then
    Begin
        T := TOB.Create('~DATA', TobLecture, 0);
        T.AddChampSupValeur('ITEM', TousLibelle);
        T.AddChampSupValeur('VALUE', '');
    End;

    CurrentObject.SetComboDetail(Combo, TobLecture);
    FreeAndNil(TobLecture);
End;

//PT3
{-----Affiche le total des colonnes d'une grille ---------------------------------------}
{ Attention : il faut que les composants PANTOTAL et GRTOTAL existent!
  La fonction gère 10 colonnes de cumul au maximum.
  NumColTitre permet de spécifier au besoin le numéro de la colonne qui recevra "Totaux:"
  dans le cas où la 1ère colonne (n° 0) est cachée.
}
procedure AfficheTotal (CurrentObject : TVignettePlugin; TobDonnees : TOB; NumColTitre : Integer = 0);
Var TobData,T,TT : TOB;
    ListeASommer : Array[1..10] of String;
    i,nbChamps : Integer;
    NumChamp   : Integer;
begin
    If TobDonnees.Detail.Count > 0 Then
    Begin
        Try
            NbChamps := 0;

            TobData := TOB.Create('£LesTotaux', Nil, -1);
            TT := TOB.Create('£Total',TobData, -1);

            // Liste des champs à sommer et création de la liste des champs
            // pour que ça corresponde aux colonnes de la grille de base
            T := TobDonnees.Detail[0];
            For i := 0 To T.NombreChampSup-1 Do
            Begin
                NumChamp := i + 1000;

                TT.AddChampSup(T.GetNomChamp(NumChamp),False);

                If (VarType(T.GetValeur(NumChamp)) = varInteger) Or (VarType(T.GetValeur(NumChamp)) = varDouble) Or
                   (VarType(T.GetValeur(NumChamp)) = varByte) Then
                Begin
                    NbChamps := NbChamps + 1;
                    ListeASommer[NbChamps] := T.GetNomChamp(NumChamp);
                End
                Else
                    If i = NumColTitre Then TT.PutValue(T.GetNomChamp(NumChamp), 'Totaux');
            End;

            // Totaux
            For i := 1 To NbChamps Do
            Begin
                TT.PutValue(ListeASommer[i], TobDonnees.SommeSimple(ListeASommer[i],[''],['']));
            End;

            CurrentObject.PutGridDetail('GRTOTAL', TobData);
        Finally
            FreeAndNil(TobData);
        End;
    End
    Else
    Begin
        CurrentObject.SetControlVisible('PANTOTAL', False);  //Cache le panneau si pas de données à afficher
    End;
end;

//PT4
{-----Suppression d'une demande d'absence ----------------------------------------------}

Function SupprimeAbsence (TobSelection : TOB; Salarie : String) : String;
Var
    TypeConge, TypeMvt,Ordre        : String;
    CodeTape,Sens,CodeEtat          : String;
    APayer, MBase,NbJours           : Double;
    Periode, Error, i               : Integer;
    Q                               : TQuery;
    Pris, Acquis, Restants, Base    : Double;
    TobMotifAbs                     : TOB;
    Retour                          : String;
Begin
    APayer  := 0;
    Periode := 0;
    Result  := '';

    If TobSelection.Detail.Count = 0 Then
    Begin
        Result := TraduireMemoire('Veuillez sélectionner une absence.');
        Exit;
    End;

    Try
    // Chargement des motifs d'absence
    TobMotifAbs := TOB.Create('$LesMotifs', Nil, -1);
    TobMotifAbs.LoadDetailDB('MOTIFABSENCE', '', 'PMA_MOTIFABSENCE', Nil, False);

    // Récupération de la session sélectionnée et de sa clé
    For i := 0 To TobSelection.Detail.Count-1 Do
    Begin
        TypeMvt   := TobSelection.Detail[i].GetValue('PCN_TYPEMVT');  //PT5
        TypeConge := TobSelection.Detail[i].GetValue('PCN_TYPECONGE');
        Ordre     := TobSelection.Detail[i].GetValue('PCN_ORDRE');
        NbJours   := TobSelection.Detail[i].GetValue('PCN_NBJOURS');
        CodeEtat  := TobSelection.Detail[i].GetValue('PCN_VALIDRESP');
        If Salarie = '' Then
        Salarie   := TobSelection.Detail[i].GetValue('PCN_SALARIE');

        // Contrôle de l'absence
        If CodeEtat <> 'ATT' Then
        Begin
            Result := TraduireMemoire('Vous ne pouvez annuler que les demandes en attente.');
            Break;
        End;

        // Récupération d'informations complémentaires
        Q := OpenSQL('SELECT PCN_CODETAPE,PCN_APAYES,PCN_SENSABS,PCN_PERIODECP FROM ABSENCESALARIE WHERE PCN_TYPEMVT="'+TypeMvt+'" AND PCN_SALARIE="'+Salarie+'" AND PCN_ORDRE="'+Ordre+'" AND PCN_RESSOURCE=""', True);
        If Not Q.EOF Then
        Begin
            CodeTape    := Q.FindField('PCN_CODETAPE').AsString;
            APayer      := Q.FindField('PCN_APAYES').AsFloat;
            Sens        := Q.FindField('PCN_SENSABS').AsString;
            Periode     := Q.FindField('PCN_PERIODECP').AsInteger;
        End
        Else
        Begin
            Result := Format(TraduireMemoire('Une erreur est survenue durant la suppression du mouvement du %s.'),[DateToStr(TobSelection.Detail[i].GetValue('PCN_DATEDEBUT'))]);
            Ferme(Q);
            Break;
        End;
        Ferme(Q);

        // Contrôles plus poussés
        If (CodeTape <> '...') and ((TypeConge = 'PRI') or (TypeMvt = 'CPA') or (TypeConge = 'AJU') or (TypeConge = 'AJP')) then
        Begin
            Result := TraduireMemoire('Vous ne pouvez pas supprimer ce mouvement.');
            Break;
        End;

        If (APayer > 0) and (Sens = '+') and ((TypeConge = 'AJU') or (TypeConge = 'AJP')) Then
        Begin
            Result := TraduireMemoire('Vous ne pouvez pas supprimer ce mouvement.#13#10' +
                                             'Une prise de congés est imputée sur ce dernier.');
            Break;
        End;

        If (TypeConge <> 'PRI') and (Periode > 0) then
        Begin
            AffichelibelleAcqPri(IntToStr(Periode - 1), Salarie, Idate1900, 0, Pris, Acquis, Restants, Base, MBase, false, False);
            If (Pris > 0) then
            Begin
                Result := TraduireMemoire('Vous ne pouvez pas supprimer ce mouvement pour cette période.#13#10' +
                                                 'Il existe des congés payés entamés sur la période suivante.');
                Break;
            End;
        End;

        // Suppression de l'absence
        ExecuteSQL ('DELETE FROM ABSENCESALARIE WHERE PCN_TYPEMVT="'+TypeMvt+'" AND PCN_SALARIE="'+Salarie+'" AND PCN_ORDRE="'+Ordre+'" AND PCN_RESSOURCE=""');

        // Mise à jour du récapitulatif salarié
        PgSupprRecapitulatif(TobMotifAbs, Salarie, TypeConge, NbJours, Error, Retour);

        if Error <> 0 then
        Begin
            Result := RecupMessageErrorAbsence(Error);
            Break;
        End;
    End;
    Finally
        FreeAndNil(TobMotifAbs);
    End;
end;

//PT5
{-----Retourne le libellé d'un booléen en base sous forme Oui / Non ----------------------------}

Function LibBool (Valeur : String) : String;
Begin
      If Valeur = 'X' Then
          Result := OUI
      Else
          Result := NON;
End;

//PT6
{-----Préfixe un nom de table par le nom de la base --------------------------------------------}

Function GetBase (Dossier, LaTable : String) : String;
Var Q : TQuery;
Begin
    If Dossier <> '' Then
    Begin
   		Q := OpenSQL('SELECT DOS_NOMBASE FROM DOSSIER WHERE DOS_NODOSSIER="'+Dossier+'"', True); //PT16
   		If Not Q.EOF Then Dossier := Q.FindField('DOS_NOMBASE').AsString;
   		Ferme(Q);
    	Result := Dossier+'.DBO.'+LaTable;
    End
    Else
        Result := LaTable;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/03/2008 / PT8
Modifié le ... :   /  /
Description .. : Récupère le type d'utilisateur connecté et le sauvegarde
Suite ........ : dans la session sous la clé "TYPE_UTILISATEUR_XXX"
Mots clefs ... :
*****************************************************************}
procedure SauveTypeUtilisateur (Hierarchie : String);
Var TypeUtilisat : String;
    LaSession    : TISession;
    TU           : TTypeUser; //PT11
Begin
	// Récupération du type d'utilisateur
	Try
	   TypeUtilisat := '';
       If ExisteSQL('SELECT 1 FROM SERVICES WHERE PGS_RESPONS'+Hierarchie+'="'+V_PGI.UserSalarie+'"') then TypeUtilisat := TypeUtilisat + 'RESPONS'+Hierarchie+';'; //PT10
       If ExisteSQL('SELECT 1 FROM SERVICES WHERE PGS_ADJOINT'+Hierarchie+'="'+V_PGI.UserSalarie+'"') then TypeUtilisat := TypeUtilisat + 'ADJOINT'+Hierarchie+';';  //PT10
       If ExisteSQL('SELECT 1 FROM SERVICES WHERE PGS_SECRETAIRE'+Hierarchie+'="'+V_PGI.UserSalarie+'"') then TypeUtilisat := TypeUtilisat + 'SECRETAIRE'+Hierarchie+';'; //PT10

       If TypeUtilisat = '' Then TypeUtilisat := 'RESPONS'+Hierarchie; // Responsable par défaut : de toute façon, il ne verra personne //PT10
    Except
    	TypeUtilisat := 'RESPONS'+Hierarchie; //PT10
    End;

    // Sauvegarde dans la session
    LaSession := LookupCurrentSession;

    //PT11
    TU := TTypeUser.Create();
    TU.Liste := TypeUtilisat;

    LaSession.UserObjects.AddObject(TYPE_UTILISATEUR+Hierarchie, TU);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/03/2008 / PT8
Modifié le ... :   /  /
Description .. : Retourne le type d'utilisateur sous la forme type : RESPONSFOR
Mots clefs ... :
*****************************************************************}
Function RecupTypeUtilisateur (Hierarchie : String) : String;
Var LaSession    : TISession;
    TU           : TTypeUser; //PT11
Begin
    LaSession := LookupCurrentSession;
    TU := TTypeUser(LaSession.UserObjects.Objects[LaSession.UserObjects.IndexOf(TYPE_UTILISATEUR+Hierarchie)]); //PT10  + Hierarchie;
    Result := TU.Liste; //PT11
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 29/04/2008 / PT15
Modifié le ... :   /  /
Description .. : Retourne la partie SQL pour n'avoir que les salariés encore dans la
Suite ........ : société au jour de la connexion.
Mots clefs ... :
*****************************************************************}
Function SalariesNonSortis (Prefixe : String = 'PSA') : String;
Begin
	Result := ' ('+Prefixe+'_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR '+
			  Prefixe+'_DATESORTIE IS NULL OR '+Prefixe+'_DATESORTIE>="'+UsDateTime(Date)+'") ';
End;

end.

