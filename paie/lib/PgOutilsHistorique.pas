{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 13/12/2005
Modifié le ... :   /  /
Description .. : Stockage des procédures communes
Mots clefs ... : PAIE;OUTILS;Historique
*****************************************************************}
{
PT1  : 25/05/2007 GGU    Mise à jour de la population du salarié lors de la mise à jour
                         de l'historique
PT2    25/09/2007 FC V_80 FQ 14807 Analyse des éléments dynamiques et des éléments nationaux
PT3    18/04/2008 GGU V_81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
PT4    29/04/2008 GGU FQ 013;15284 l'historique du nouveau salarié n'est pas initialisé correctement. Le champ PGTYPEINFOLS devrait être initialisé avec "SAL"
PT5    03/07/2008 PH FQ 013;15599 Phase initialisation à partir TOMSalarie PHD_TRAITEMENTOK = '-'
PT7    05/08/2008 PH FQ 15720 Mauvais nom de champ situation de famille table salarie en historique par avance
}
unit PGOutilsHistorique;

interface

uses SysUtils,
     HEnt1,
     UTOB,
     HCtrls,
     Controls,
     Ed_Tools,
     Hqry,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HMsgBox;

procedure PGSalariesHisto;
procedure PGRecupAncienHisto(Salarie : String);
procedure CreerHisto(LeSalarie : String;LaDate : TDateTime;TypeHisto,LeChamp,LaValeur,AncValeur,Tablette,LeType : String; TRAITEMENTFIN : String = 'X');
procedure PGMajDonneesAfficheZL(Partiel : boolean = False;DateValidite : TDateTime = 0;bModeValeur : boolean =True);
Procedure PGMajHistorique(Salarie : String;DD,DF : TDateTime);
Procedure PGRecupValeurHisto(var TobSal : Tob;DF : TDateTime);
  


implementation

Uses
  PgPopulOutils, Classes,
  PGTablesDyna //PT3
  ;

procedure PGSalariesHisto;
var TSal : Tob;
    i : Integer;
    Q : TQuery;
    Salarie : String;
begin
 If existeSQL('SELECT * FROM PGHISTODETAIL WHERE PHD_PGTYPEHISTO="001" OR PHD_PGTYPEHISTO="002"') then exit;
  If not V_PGI.Sav then exit;
  Q := OpenSQL('SELECT PSA_SALARIE FROM SALARIES',True);
  TSal := Tob.Create('lessalries',Nil,-1);
  TSal.LoadDetailDB('lessalries','','',Q,False);
  Ferme(Q);
  For i := 0 to TSal.Detail.Count - 1 do
  begin
       Salarie := TSal.Detail[i].GetValue('PSA_SALARIE');
       PGRecupAncienHisto(Salarie);
  end;
  FreeAndNil(TSal);
end;

{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 13/12/2005
Modifié le ... :   /  /
Description .. : Recopie ancienne table historique
Suite ........ : dans nouvelle tablette PGHISTODETAIL
Mots clefs ... : PAIE,HISTORIQUE
*****************************************************************}
procedure PGRecupAncienHisto(Salarie : string);
var TobAncValeur,TobAncH : Tob;
    Q : TQuery;
    i : Integer;
    DateEven : TDateTime;
    TyHisto : String;
begin
  Q := OpenSQL('SELECT * FROM HISTOSALARIE WHERE PHS_SALARIE="'+Salarie+'" ORDER BY PHS_DATEEVENEMENT',True);
  TobAncH := Tob.Create('AncHistorique',Nil,-1);
  TobAncH.LoadDetailDB('AncHistorique','','',Q,False);
  Ferme(Q);
  TobAncValeur := Tob.create('LesAnciennesValeurs',Nil,-1);
  TobAncValeur.AddChampSupValeur('PHS_CODEEMPLOI','');
  TobAncValeur.AddChampSupValeur('PHS_LIBELLEEMPLOI','');
  TobAncValeur.AddChampSupValeur('PHS_QUALIFICATION','');
  TobAncValeur.AddChampSupValeur('PHS_COEFFICIENT','');
  TobAncValeur.AddChampSupValeur('PHS_INDICE','');
  TobAncValeur.AddChampSupValeur('PHS_NIVEAU','');
  TobAncValeur.AddChampSupValeur('PHS_CODESTAT','');
  TobAncValeur.AddChampSupValeur('PHS_TRAVAILN1','');
  TobAncValeur.AddChampSupValeur('PHS_TRAVAILN2','');
  TobAncValeur.AddChampSupValeur('PHS_TRAVAILN3','');
  TobAncValeur.AddChampSupValeur('PHS_TRAVAILN4','');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREMOIS1','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREANN1','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREMOIS2','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREANN2','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREMOIS3','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREANN3','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREMOIS4','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREANN4','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREMOIS5','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIREANN5','0');
  TobAncValeur.AddChampSupValeur('PHS_DADSPROF','0');
  TobAncValeur.AddChampSupValeur('PHS_DADSCAT','0');
  TobAncValeur.AddChampSupValeur('PHS_HORAIREMOIS','0');
  TobAncValeur.AddChampSupValeur('PHS_TAUXHORAIRE','0');
  TobAncValeur.AddChampSupValeur('PHS_SALAIRETHEO','0');
  TobAncValeur.AddChampSupValeur('PHS_BETABLISSEMENT','0');
  For i := 0 to TobAncH.Detail.Count - 1 do
  begin
    If Copy(TobAncH.Detail[i].GetValue('PHS_COMMENTAIRE'),1,4) = 'Init' then TyHisto := '001'
    else TyHisto := '002';
    DateEven := TobAncH.Detail[i].GetValue('PHS_DATEEVENEMENT'); //PHS_DATEAPPLIC
    If TobAncH.Detail[i].GetValue('PHS_BCODEEMPLOI') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'48',TobAncH.Detail[i].GetValue('PHS_CODEEMPLOI'),TobAncValeur.GetValue('PHS_CODEEMPLOI'),'PGCODEEMPLOI','T');
         TobAncValeur.PutValue('PHS_CODEEMPLOI',TobAncH.Detail[i].GetValue('PHS_CODEEMPLOI'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BLIBELLEEMPLOI') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'49',TobAncH.Detail[i].GetValue('PHS_LIBELLEEMPLOI'),TobAncValeur.GetValue('PHS_LIBELLEEMPLOI'),'PGLIBEMPLOI','T');
         TobAncValeur.PutValue('PHS_LIBELLEEMPLOI',TobAncH.Detail[i].GetValue('PHS_LIBELLEEMPLOI'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BQUALIFICATION') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'46',TobAncH.Detail[i].GetValue('PHS_QUALIFICATION'),TobAncValeur.GetValue('PHS_QUALIFICATION'),'PGLIBQUALIFICATION','T');
         TobAncValeur.PutValue('PHS_QUALIFICATION',TobAncH.Detail[i].GetValue('PHS_QUALIFICATION'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BCOEFFICIENT') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'45',TobAncH.Detail[i].GetValue('PHS_COEFFICIENT'),TobAncValeur.GetValue('PHS_COEFFICIENT'),'PGLIBCOEFFICIENT','T');
         TobAncValeur.PutValue('PHS_COEFFICIENT',TobAncH.Detail[i].GetValue('PHS_COEFFICIENT'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BINDICE') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'220',TobAncH.Detail[i].GetValue('PHS_INDICE'),TobAncValeur.GetValue('PHS_INDICE'),'PGLIBINDICE','T');
         TobAncValeur.PutValue('PHS_INDICE',TobAncH.Detail[i].GetValue('PHS_INDICE'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BNIVEAU') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'221',TobAncH.Detail[i].GetValue('PHS_NIVEAU'),TobAncValeur.GetValue('PHS_NIVEAU'),'PGLIBNIVEAU','T');
         TobAncValeur.PutValue('PHS_NIVEAU',TobAncH.Detail[i].GetValue('PHS_NIVEAU'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BCODESTAT') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'161',TobAncH.Detail[i].GetValue('PHS_CODESTAT'),TobAncValeur.GetValue('PHS_CODESTAT'),'PGCODESTAT','T');
         TobAncValeur.PutValue('PHS_CODESTAT',TobAncH.Detail[i].GetValue('PHS_CODESTAT'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BTRAVAILN1') = 'X' then
    begin
         CreerHisto(Salarie,DateEven,TyHisto,'157',TobAncH.Detail[i].GetValue('PHS_TRAVAILN1'),TobAncValeur.GetValue('PHS_TRAVAILN1'),'PGTRAVAILN1','T');
         TobAncValeur.PutValue('PHS_TRAVAILN1',TobAncH.Detail[i].GetValue('PHS_TRAVAILN1'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BTRAVAILN2') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'158',TobAncH.Detail[i].GetValue('PHS_TRAVAILN2'),TobAncValeur.GetValue('PHS_TRAVAILN2'),'PGTRAVAILN2','T');
       TobAncValeur.PutValue('PHS_TRAVAILN2',TobAncH.Detail[i].GetValue('PHS_TRAVAILN2'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BTRAVAILN3') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'159',TobAncH.Detail[i].GetValue('PHS_TRAVAILN3'),TobAncValeur.GetValue('PHS_TRAVAILN3'),'PGTRAVAILN3','T');
       TobAncValeur.PutValue('PHS_TRAVAILN3',TobAncH.Detail[i].GetValue('PHS_TRAVAILN3'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BTRAVAILN4') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'160',TobAncH.Detail[i].GetValue('PHS_TRAVAILN4'),TobAncValeur.GetValue('PHS_TRAVAILN4'),'PGTRAVAILN4','T');
       TobAncValeur.PutValue('PHS_TRAVAILN4',TobAncH.Detail[i].GetValue('PHS_TRAVAILN4'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREMOIS1') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'110',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS1'),TobAncValeur.GetString('PHS_SALAIREMOIS1'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREMOIS1',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS1'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREANN1') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'115',TobAncH.Detail[i].GetString('PHS_SALAIREANN1'),TobAncValeur.GetString('PHS_SALAIREANN1'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREANN1',TobAncH.Detail[i].GetString('PHS_SALAIREANN1'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREMOIS2') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'111',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS2'),TobAncValeur.GetString('PHS_SALAIREMOIS2'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREMOIS2',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS2'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREANN2') = 'X' then
    begin
        CreerHisto(Salarie,DateEven,TyHisto,'116',TobAncH.Detail[i].GetString('PHS_SALAIREANN2'),TobAncValeur.GetString('PHS_SALAIREANN2'),'','F');
        TobAncValeur.PutValue('PHS_SALAIREANN2',TobAncH.Detail[i].GetString('PHS_SALAIREANN2'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREMOIS3') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'112',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS3'),TobAncValeur.GetString('PHS_SALAIREMOIS3'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREMOIS3',TobAncH.Detail[i].GetValue('PHS_SALAIREMOIS3'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREANN3') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'117',TobAncH.Detail[i].GetString('PHS_SALAIREANN3'),TobAncValeur.GetString('PHS_SALAIREANN3'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREANN3',TobAncH.Detail[i].GetValue('PHS_SALAIREANN3'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREMOIS4') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'113',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS4'),TobAncValeur.GetString('PHS_SALAIREMOIS4'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREMOIS4',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS4'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREANN4') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'118',TobAncH.Detail[i].GetString('PHS_SALAIREANN4'),TobAncValeur.GetString('PHS_SALAIREANN4'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREANN4',TobAncH.Detail[i].GetString('PHS_SALAIREANN4'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREMOIS5') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'114',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS5'),TobAncValeur.GetString('PHS_SALAIREMOIS5'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREMOIS5',TobAncH.Detail[i].GetString('PHS_SALAIREMOIS5'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIREANN5') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'119',TobAncH.Detail[i].GetString('PHS_SALAIREANN5'),TobAncValeur.GetString('PHS_SALAIREANN5'),'','F');
       TobAncValeur.PutValue('PHS_SALAIREANN5',TobAncH.Detail[i].GetString('PHS_SALAIREANN5'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BDADSPROF') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'400',TobAncH.Detail[i].GetValue('PHS_DADSPROF'),TobAncValeur.GetValue('PHS_DADSPROF'),'PGSPROFESSIONNEL','T');
       TobAncValeur.PutValue('PHS_DADSPROF',TobAncH.Detail[i].GetValue('PHS_DADSPROF'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BDADSCAT') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'401',TobAncH.Detail[i].GetValue('PHS_DADSCAT'),TobAncValeur.GetValue('PHS_DADSCAT'),'PGSCATEGORIEL','T');
       TobAncValeur.PutValue('PHS_DADSCAT',TobAncH.Detail[i].GetValue('PHS_DADSCAT'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BHORAIREMOIS') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'64',TobAncH.Detail[i].GetString('PHS_HORAIREMOIS'),TobAncValeur.GetString('PHS_HORAIREMOIS'),'','F');
       TobAncValeur.PutValue('PHS_HORAIREMOIS',TobAncH.Detail[i].GetString('PHS_HORAIREMOIS'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BTAUXHORAIRE') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'67',TobAncH.Detail[i].GetString('PHS_TAUXHORAIRE'),TobAncValeur.GetString('PHS_TAUXHORAIRE'),'','F');
       TobAncValeur.PutValue('PHS_TAUXHORAIRE',TobAncH.Detail[i].GetString('PHS_TAUXHORAIRE'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BSALAIRETHEO') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'66',TobAncH.Detail[i].GetString('PHS_SALAIRETHEO'),TobAncValeur.GetString('PHS_SALAIRETHEO'),'','F');
       TobAncValeur.PutValue('PHS_SALAIRETHEO',TobAncH.Detail[i].GetString('PHS_SALAIRETHEO'));
    end;
    If TobAncH.Detail[i].GetValue('PHS_BETABLISSEMENT') = 'X' then
    begin
       CreerHisto(Salarie,DateEven,TyHisto,'8',TobAncH.Detail[i].GetValue('PHS_ETABLISSEMENT'),TobAncValeur.GetValue('PHS_ETABLISSEMENT'),'TTETABLISSEMENT','T  ');
       TobAncValeur.PutValue('PHS_ETABLISSEMENT',TobAncH.Detail[i].GetValue('PHS_ETABLISSEMENT'));
    end;
{   If TobAncH.Detail[i].GetValue('PHS_BDTLIBRE1') = 'X' then CreerHisto(Salarie,DateEven,'PHS_DTLIBRE1',TobAncH.Detail[i].GetValue('PHS_DTLIBRE1'));
    If TobAncH.Detail[i].GetValue('PHS_BDTLIBRE2') = 'X' then CreerHisto(Salarie,DateEven,'PHS_DTLIBRE2',TobAncH.Detail[i].GetValue('PHS_DTLIBRE2'));
    If TobAncH.Detail[i].GetValue('PHS_BDTLIBRE3') = 'X' then CreerHisto(Salarie,DateEven,'PHS_DTLIBRE3',TobAncH.Detail[i].GetValue('PHS_DTLIBRE3'));
    If TobAncH.Detail[i].GetValue('PHS_BDTLIBRE4') = 'X' then CreerHisto(Salarie,DateEven,'PHS_DTLIBRE4',TobAncH.Detail[i].GetValue('PHS_DTLIBRE4'));
    If TobAncH.Detail[i].GetValue('PHS_BBOOLLIBRE1') = 'X' then CreerHisto(Salarie,DateEven,'PHS_BOOLLIBRE1',TobAncH.Detail[i].GetValue('PHS_BOOLLIBRE1'));
    If TobAncH.Detail[i].GetValue('PHS_BBOOLLIBRE2') = 'X' then CreerHisto(Salarie,DateEven,'PHS_BOOLLIBRE2',TobAncH.Detail[i].GetValue('PHS_BOOLLIBRE2'));
    If TobAncH.Detail[i].GetValue('PHS_BBOOLLIBRE3') = 'X' then CreerHisto(Salarie,DateEven,'PHS_BOOLLIBRE3',TobAncH.Detail[i].GetValue('PHS_BOOLLIBRE3'));
    If TobAncH.Detail[i].GetValue('PHS_BBOOLLIBRE4') = 'X' then CreerHisto(Salarie,DateEven,'PHS_BOOLLIBRE4',TobAncH.Detail[i].GetValue('PHS_BOOLLIBRE4'));
    If TobAncH.Detail[i].GetValue('PHS_BCBLIBRE1') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CBLIBRE1',TobAncH.Detail[i].GetValue('PHS_CBLIBRE1'));
    If TobAncH.Detail[i].GetValue('PHS_BCBLIBRE2') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CBLIBRE2',TobAncH.Detail[i].GetValue('PHS_CBLIBRE2'));
    If TobAncH.Detail[i].GetValue('PHS_BCBLIBRE3') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CBLIBRE3',TobAncH.Detail[i].GetValue('PHS_CBLIBRE3'));
    If TobAncH.Detail[i].GetValue('PHS_BCBLIBRE4') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CBLIBRE4',TobAncH.Detail[i].GetValue('PHS_CBLIBRE4'));
    If TobAncH.Detail[i].GetValue('PHS_BPROFIL') = 'X' then CreerHisto(Salarie,DateEven,'PHS_PROFIL',TobAncH.Detail[i].GetValue('PHS_PROFIL'));
    If TobAncH.Detail[i].GetValue('PHS_BPERIODEBUL') = 'X' then CreerHisto(Salarie,DateEven,'PHS_PERIODBUL',TobAncH.Detail[i].GetValue('PHS_PERIODBUL'));
    If TobAncH.Detail[i].GetValue('PHS_BPGHHORMOIS') = 'X' then CreerHisto(Salarie,DateEven,'PHS_PGHHORAIREMOIS',TobAncH.Detail[i].GetValue('PHS_PGHHORAIREMOIS'));
    If TobAncH.Detail[i].GetValue('PHS_BPGHHORHEBDO') = 'X' then CreerHisto(Salarie,DateEven,'PHS_PGHHORHEBDO',TobAncH.Detail[i].GetValue('PHS_PGHHORHEBDO'));
    If TobAncH.Detail[i].GetValue('PHS_BPGHHORANNUEL') = 'X' then CreerHisto(Salarie,DateEven,'PHS_PGHHORANNUEL',TobAncH.Detail[i].GetValue('PHS_PGHHORANNUEL'));
    If TobAncH.Detail[i].GetValue('PHS_BPGHTAUXHOR') = 'X' then CreerHisto(Salarie,DateEven,'PHS_PGHTAUXHORAIRE',TobAncH.Detail[i].GetValue('PHS_PGHTAUXHORAIRE'));
    If TobAncH.Detail[i].GetValue('PHS_BTAUXPARTIEL') = 'X' then CreerHisto(Salarie,DateEven,'PHS_TAUXPARTIEL',TobAncH.Detail[i].GetValue('PHS_TAUXPARTIEL'));
    If TobAncH.Detail[i].GetValue('PHS_BPROFILREM') = 'X' then CreerHisto(Salarie,DateEven,'PHS_PROFILREM',TobAncH.Detail[i].GetValue('PHS_PROFILREM'));
    If TobAncH.Detail[i].GetValue('PHS_BCONDEMPLOI') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CONDEMPLOI',TobAncH.Detail[i].GetValue('PHS_CONDEMPLOI'));
    If TobAncH.Detail[i].GetValue('PHS_BTTAUXPARTIEL') = 'X' then CreerHisto(Salarie,DateEven,'PHS_TTAUXPARTIEL',TobAncH.Detail[i].GetValue('PHS_TTAUXPARTIEL'));
    If TobAncH.Detail[i].GetValue('PHS_BCHARLIBRE1') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CHARLIBRE1',TobAncH.Detail[i].GetValue('PHS_CHARLIBRE1'));
    If TobAncH.Detail[i].GetValue('PHS_BCHARLIBRE2') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CHARLIBRE2',TobAncH.Detail[i].GetValue('PHS_CHARLIBRE2'));
    If TobAncH.Detail[i].GetValue('PHS_BCHARLIBRE3') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CHARLIBRE3',TobAncH.Detail[i].GetValue('PHS_CHARLIBRE3'));
    If TobAncH.Detail[i].GetValue('PHS_BCHARLIBRE4') = 'X' then CreerHisto(Salarie,DateEven,'PHS_CHARLIBRE4',TobAncH.Detail[i].GetValue('PHS_CHARLIBRE4'));}
  end;
  FreeAndNil(TobAncValeur);
  FreeAndNil(TobAncH);
end;

procedure CreerHisto(LeSalarie : String;LaDate : TDateTime;TypeHisto,LeChamp,LaValeur,AncValeur,Tablette,LeType : String; TRAITEMENTFIN : String = 'X'); // PT5
var TH : Tob;
begin
     TH := Tob.Create('PGHISTODETAIL',Nil,-1);
     TH.PutValue('PHD_SALARIE',LeSalarie);
     TH.PutValue('PHD_ETABLISSEMENT','');
     TH.PutValue('PHD_ORDRE',-1);
     TH.PutValue('PHD_GUIDHISTO',AglGetGuid());
     TH.PutValue('PHD_PGINFOSMODIF',LeChamp);
     TH.PutValue('PHD_PGTYPEHISTO',TypeHisto);
     TH.PutValue('PHD_ANCVALEUR',AncValeur);
     TH.PutValue('PHD_NEWVALEUR',LaValeur);
     TH.PutValue('PHD_TYPEVALEUR',LeType);
     TH.PutValue('PHD_TABLETTE',Tablette);
     TH.PutValue('PHD_DATEAPPLIC',LaDate);
     TH.PutValue('PHD_TRAITEMENTOK',TRAITEMENTFIN); // PT5
     TH.PutValue('PHD_DATEFINVALID',IDate1900);
     TH.PutValue('PHD_TYPEBUDG','');
     TH.PutValue('PHD_NUMAUG',0);
     TH.PutValue('PHD_ANNEE','');
     TH.PutValue('PHD_PGTYPEINFOLS','SAL');  //PT4
     TH.InsertDB(Nil);
     FreeAndNil(TH);
end;

procedure PGMajDonneesAfficheZL(Partiel : boolean = False;DateValidite : TDateTime = 0;bModeValeur : boolean =True);  //PT2
Var Q,Qt,Qc : TQuery;
    i,x,j : Integer;
    TobSal,T,TobParam : Tob;
    Champ,TypeChamp,Valeur : String;
    StSQL,StSal : String;
    CodTabl,NewValeur,ConvSal,Salarie,St : String;
    LocaleTob : Tob;
begin
  StSql := '';
  StSal := '';
  if DateValidite = 0 then DateValidite := V_PGI.DateEntree; //PT2
  If Partiel then  StSal := ' WHERE PSA_DATECREATION>=(SELECT MAX(PTZ_DATEMODIF) FROM PGTEMPZONELIBRE)';
   InitMoveProgressForm(nil, 'Maj des données', 'Veuillez patienter SVP ...', 30, FALSE, TRUE);
   If Not Partiel then
   begin
      ExecuteSQL('DELETE FROM PGTEMPZONELIBRE');
      Q := OpenSQL('SELECT PSA_SALARIE FROM SALARIES'+StSal,True);
      TobSal := Tob.Create('Les Salaries',Nil,-1);
      TobSal.LoadDetailDB('Les Salaries','','',Q,False);
      Ferme(Q);
      For i := 0 to TobSal.Detail.Count - 1 do
      begin
        T := Tob.Create('PGTEMPZONELIBRE',Nil,-1);
        T.AddChampSupValeur('PTZ_SALARIE',TobSal.Detail[i].GetValue('PSA_SALARIE'),False);
        T.InsertDB(Nil,False);
        FreeAndNil(T);
      end;
      FreeAndNil(TobSal);
   end;
   Q := OpenSQL('SELECT * FROM PGPARAMAFFICHEZL',True);
   TobParam := Tob.Create('MaTob',Nil,-1);
   TobParam.LoadDetailDB('Table','','',Q,False);
   Ferme(Q);
   For i := 0 to TobParam.Detail.Count - 1 do
   begin
     For x := 1 to 30 do
     begin
     Valeur := TobParam.Detail[i].GetValue('PAZ_CHAMPDISPO'+IntToStr(x));
     Champ := Copy(Valeur,4,Length(Valeur));
     TypeChamp := Copy(Valeur,1,3);
       If Partiel then StSql := ' WHERE PTZ_SALARIE IN(SELECT PHD_SALARIE FROM PGHISTODETAIL WHERE PHD_SALARIE=PTZ_SALARIE'+
    ' AND PHD_PGINFOSMODIF="'+Champ+'" AND PHD_DATEAPPLIC<="'+UsDateTime(DateValidite)+'" AND PHD_DATEMODIF>PTZ_DATEMODIF)';  //PT2
      If Champ <> '' then
      begin
        If TypeChamp = 'ZLS' then
        begin
            //DEB PT2
            if bModeValeur then
              ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'=(SELECT ##TOP 1## PHD_NEWVALEUR ' +
              ' FROM PGHISTODETAIL WHERE PHD_SALARIE=PTZ_SALARIE'+
              ' AND PHD_PGINFOSMODIF="'+Champ+'" AND PHD_DATEAPPLIC<="'+UsDateTime(DateValidite)+'"' +
              ' ORDER BY PHD_DATEAPPLIC DESC)'+StSql)
            else
            begin
              ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'=(SELECT ##TOP 1## PHD_NEWVALEUR ' +
              ' FROM PGHISTODETAIL WHERE PHD_SALARIE=PTZ_SALARIE'+
              ' AND PHD_PGINFOSMODIF="'+Champ+'" AND PHD_DATEAPPLIC<="'+UsDateTime(DateValidite)+'"' +
              ' ORDER BY PHD_DATEAPPLIC DESC)'+StSql);

              St := 'SELECT PHD_SALARIE,PHD_NEWVALEUR,PHD_CODTABL FROM PGHISTODETAIL ' +
              ' WHERE PHD_PGINFOSMODIF="'+Champ+'" AND PHD_DATEAPPLIC<="'+UsDateTime(DateValidite)+'"' +
              ' AND PHD_NEWVALEUR <> "" AND PHD_CODTABL <> ""' +
              ' AND PHD_SALARIE IN (SELECT DISTINCT PTZ_SALARIE FROM PGTEMPZONELIBRE)' +
              ' ORDER BY PHD_DATEAPPLIC DESC';
              LocaleTob := Tob.Create('HistoDetail',Nil,-1);
              LocaleTob.LoadDetailDbFromSQL('PGHISTODETAIL',st) ;
              Salarie := '';
              for j := 0 to LocaleTob.Detail.Count - 1 do
              begin
                if Salarie <> LocaleTob.Detail[j].GetValue('PHD_SALARIE') then
                begin
                  CodTabl := LocaleTob.Detail[j].GetValue('PHD_CODTABL');
                  NewValeur := LocaleTob.Detail[j].GetValue('PHD_NEWVALEUR');
                  Salarie := LocaleTob.Detail[j].GetValue('PHD_SALARIE');
                  //PT3
                  ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'="' +
                    GetLibelleZLTableDyna(DateValidite, CodTabl, Salarie, NewValeur) + '"'+
                    ' WHERE PTZ_SALARIE = "'+ Salarie + '"');
{//PT3
                  ConvSal := '';
                  Qc := OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"',True);
                  if not Qc.Eof then
                    ConvSal := Qc.FindField('PSA_CONVENTION').AsString;
                  Ferme(Qc);

                  if ExisteSQL('SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
                    ' WHERE ##PTD_PREDEFINI## AND PTD_CODTABL="' + CodTabl + '"' +
                    ' AND PTD_VALCRIT1="' + NewValeur + '"' +
                    ' AND PTD_DTVALID=(SELECT MAX(PTD_DTVALID) FROM TABLEDIMDET WHERE PTD_CODTABL="' + CodTabl + '"' +
                    ' AND PTD_DTVALID<="' + USDATETIME(DateValidite) + '"' +
                    ' AND PTD_NIVSAIS = "CON" AND PTD_VALNIV IN ("000","' + ConvSal + '"))') then
                  begin
                    Qt := OpenSQL('SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
                    ' WHERE ##PTD_PREDEFINI## AND PTD_CODTABL="' + CodTabl + '"' +
                    ' AND PTD_VALCRIT1="' + NewValeur + '"' +
                    ' AND PTD_DTVALID=(SELECT MAX(PTD_DTVALID) FROM TABLEDIMDET WHERE PTD_CODTABL="' + CodTabl + '"' +
                    ' AND PTD_DTVALID<="' + USDATETIME(DateValidite) + '"' +
                    ' AND PTD_NIVSAIS = "CON" AND PTD_VALNIV IN ("000","' + ConvSal + '"))' +
                    ' ORDER BY PTD_DTVALID DESC,PTD_NIVSAIS',True);
                    if not Qt.Eof then
                      ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'="' +
                        Qt.FindField('PTD_LIBELLECODE').AsString + '"'+
                        ' WHERE PTZ_SALARIE = "'+ Salarie + '"');
                    Ferme(Qt);
                  end
                  else
                  begin
                    Qt := OpenSQL('SELECT PTD_LIBELLECODE FROM TABLEDIMDET ' +
                    ' WHERE ##PTD_PREDEFINI## AND PTD_CODTABL="' + CodTabl + '"' +
                    ' AND PTD_VALCRIT1="' + NewValeur + '"' +
                    ' AND PTD_DTVALID=(SELECT MAX(PTD_DTVALID) FROM TABLEDIMDET WHERE PTD_CODTABL="' + CodTabl + '"' +
                    ' AND PTD_DTVALID<="' + USDATETIME(DateValidite) + '"' +
                    ' AND PTD_NIVSAIS = "GEN")' +
                    ' ORDER BY PTD_DTVALID DESC,PTD_NIVSAIS',True);
                    if not Qt.Eof then
                      ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'="' +
                        Qt.FindField('PTD_LIBELLECODE').AsString + '"'+
                        ' WHERE PTZ_SALARIE = "'+ Salarie + '"');
                  end;
                  Ferme(Qt);
}

                end;
              end;
              FreeAndNil(LocaleTob);
            end;
            //FIN PT2
        end
        else
            //Element nat
            ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'=(SELECT ##TOP 1## PED_MONTANTEURO FROM ELTNATIONDOS'+
            ' WHERE PED_CODEELT="'+Champ+'" AND PED_VALEURNIVEAU=PTZ_SALARIE AND PED_TYPENIVEAU = "SAL" AND PED_DATEVALIDITE<="'+UsDateTime(DateValidite)+'"'+  //PT2
            ' ORDER BY PED_DATEVALIDITE DESC)');
            MoveCurProgressForm;
      end;
     end;
   end;
   FreeAndNil(TobParam);
    FiniMoveProgressForm;
end;

{
procedure PGMajDonneesAfficheZL(Partiel : boolean = False;DateValidite : TDateTime = 0);  //PT2
Var Q,Qt,Qc : TQuery;
    i,x,j : Integer;
    TobSal,T,TobParam : Tob;
    Champ,TypeChamp,Valeur : String;
    StSQL,StSal : String;
    CodTabl,NewValeur,ConvSal,Salarie : String;
    LocaleTob : Tob;
begin
  StSql := '';
  StSal := '';
  if DateValidite = 0 then DateValidite := V_PGI.DateEntree; //PT2
  If Partiel then  StSal := ' WHERE PSA_DATECREATION>=(SELECT MAX(PTZ_DATEMODIF) FROM PGTEMPZONELIBRE)';
   InitMoveProgressForm(nil, 'Maj des données', 'Veuillez patienter SVP ...', 30, FALSE, TRUE);
   If Not Partiel then
   begin
      ExecuteSQL('DELETE FROM PGTEMPZONELIBRE');
      Q := OpenSQL('SELECT PSA_SALARIE FROM SALARIES'+StSal,True);
      TobSal := Tob.Create('Les Salaries',Nil,-1);
      TobSal.LoadDetailDB('Les Salaries','','',Q,False);
      Ferme(Q);
      For i := 0 to TobSal.Detail.Count - 1 do
      begin
        T := Tob.Create('PGTEMPZONELIBRE',Nil,-1);
        T.AddChampSupValeur('PTZ_SALARIE',TobSal.Detail[i].GetValue('PSA_SALARIE'),False);
        T.InsertDB(Nil,False);
      end;
      TobSal.Free;
   end;
   Q := OpenSQL('SELECT * FROM PGPARAMAFFICHEZL',True);
   TobParam := Tob.Create('MaTob',Nil,-1);
   TobParam.LoadDetailDB('Table','','',Q,False);
   Ferme(Q);
   For i := 0 to TobParam.Detail.Count - 1 do
   begin
     For x := 1 to 30 do
     begin
     Valeur := TobParam.Detail[i].GetValue('PAZ_CHAMPDISPO'+IntToStr(x));
     Champ := Copy(Valeur,4,Length(Valeur));
     TypeChamp := Copy(Valeur,1,3);
       If Partiel then StSql := ' WHERE PTZ_SALARIE IN(SELECT PHD_SALARIE FROM PGHISTODETAIL WHERE PHD_SALARIE=PTZ_SALARIE'+
    ' AND PHD_PGINFOSMODIF="'+Champ+'" AND PHD_DATEAPPLIC<="'+UsDateTime(DateValidite)+'" AND PHD_DATEMODIF>PTZ_DATEMODIF)';  //PT2
      If Champ <> '' then
      begin
        If TypeChamp = 'ZLS' then
            ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'=(SELECT ##TOP 1## PHD_NEWVALEUR FROM PGHISTODETAIL WHERE PHD_SALARIE=PTZ_SALARIE'+
            ' AND PHD_PGINFOSMODIF="'+Champ+'" AND PHD_DATEAPPLIC<="'+UsDateTime(DateValidite)+'" ORDER BY PHD_DATEAPPLIC DESC)'+StSql)  //PT2
        else
            //Element nat
            ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'=(SELECT ##TOP 1## PED_MONTANTEURO FROM ELTNATIONDOS'+
            ' WHERE PED_CODEELT="'+Champ+'" AND PED_VALEURNIVEAU=PTZ_SALARIE AND PED_TYPENIVEAU = "SAL" AND PED_DATEVALIDITE<="'+UsDateTime(DateValidite)+'"'+  //PT2
            ' ORDER BY PED_DATEVALIDITE DESC)');
            MoveCurProgressForm;
      end;
     end;
   end;
   TobParam.Free;
    FiniMoveProgressForm;
end;
}
{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 27/03/2007
Modifié le ... :   /  /    
Description .. : Procedure de maj de l'historique salarié
Mots clefs ... : 
*****************************************************************}
Procedure PGMajHistorique(Salarie : String;DD,DF : TDateTime);
var Q : TQuery;
  TobHisto : Tob;
  i : Integer;
  Champ,TypeChamp,Valeur : String;
  //Début PT1
  UpdateIdemPop : TUpdateIdemPop;
  TmpStringList : TStringList;
  //Fin PT1
begin
  Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE PHD_SALARIE="'+Salarie+'" AND '+
  'PHD_TRAITEMENTOK="-" AND PHD_PGTYPEINFOLS="SAL" AND PHD_DATEAPPLIC>="'+UsDateTime(DD)+'" AND PHD_DATEAPPLIC<="'+UsDateTime(DF)+'" ORDER BY PHD_DATEAPPLIC',True);
  TobHisto := Tob.Create('parametrage salarie',Nil,-1);
  TobHisto.LoadDetailDB('PGHISTODETAIL', '', '', Q, FALSE, False);
  Ferme(Q);
  For i := 0 to TobHisto.Detail.Count -1 do
  begin
    Champ := TobHisto.Detail[i].GetValue('PHD_PGINFOSMODIF');
    If Champ = 'PSA_SITUATIONFAMI' then Champ := 'PSA_SITUATIONFAMIL'; // PT7
    TypeChamp := TobHisto.Detail[i].GetValue('PHD_TYPEVALEUR');
    Valeur := TobHisto.Detail[i].GetValue('PHD_NEWVALEUR');
    If (TypeChamp='I') then ExecuteSQL('UPDATE SALARIES SET '+Champ+'='+Valeur+' WHERE PSA_SALARIE="'+Salarie+'"')
    else If (TypeChamp='F') then ExecuteSQL('UPDATE SALARIES SET '+Champ+'='+STRFPOINT(StrToFloat(Valeur))+' WHERE PSA_SALARIE="'+Salarie+'"')
    else If (TypeChamp='D') then ExecuteSQL('UPDATE SALARIES SET '+Champ+'="'+UsDateTime(StrTodate(Valeur))+'" WHERE PSA_SALARIE="'+Salarie+'"')
    else ExecuteSQL('UPDATE SALARIES SET '+Champ+'="'+Valeur+'" WHERE PSA_SALARIE="'+Salarie+'"');
    TobHisto.Detail[i].PutValue('PHD_TRAITEMENTOK','X');
    TobHisto.Detail[i].UpdateDB(False);
  end;
  //Début PT1 : Mise à jour de la population du salarié lors de la mise à jour de l'historique
  if TobHisto.Detail.Count > 0 then
  begin
    UpdateIdemPop := TUpdateIdemPop.create;
    TmpStringList := UpdateIdemPop.MajSALARIEPOPULSalarie(Salarie,DD,[Champ]);
    FreeAndNil(TmpStringList);
    FreeAndNil(UpdateIdemPop);
  end;
  //Fin PT1
  FreeAndNil(TobHisto);
end;

{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 27/03/2007
Modifié le ... :   /  /    
Description .. : Récupération des valeurs de la fiche salarié pour une 
Suite ........ : donnée
Mots clefs ... : 
*****************************************************************}
Procedure PGRecupValeurHisto(var TobSal : Tob;DF : TDateTime);
var Salarie : String;
    TobChampModif : Tob;
    i  : Integer;
    Champ,Valeur,TypeValeur : String;
    Q : TQuery;
begin
  Salarie := TobSal.GetValue('PSA_SALARIE');
  Q := OpenSQL('SELECT  DISTINCT (PHD_PGINFOSMODIF) FROM PGHISTODETAIL WHERE PHD_DATEAPPLIC>"'+UsDateTime(DF)+'" AND PHD_SALARIE="'+Salarie+'"',True);
  TobChampModif := Tob.Create('Les modifs',Nil,-1);
  TobChampModif.LoadDetailDB('Les modifs','','',Q,False);
  Ferme(Q);
  For i := 0 to TobChampModif.Detail.Count - 1 do
  begin
    Champ := TobChampModif.Detail[i].GetValue('PHD_PGINFOSMODIF');
    If not TobSal.FieldExists(Champ) then Continue;
    Valeur := '';
    TypeValeur := '';
    Q := OpenSQL('SELECT ##TOP 1## PHD_NEWVALEUR,PHD_TYPEVALEUR FROM PGHISTODETAIL WHERE PHD_PGINFOSMODIF="'+Champ+'" '+
    'AND PHD_DATEAPPLIC<="'+UsDateTime(DF)+'" AND PHD_SALARIE="'+Salarie+'" ORDER BY PHD_DATEAPPLIC DESC',True);
    If Not Q.Eof then
    begin
      Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
      TypeValeur := Q.FindField('PHD_TYPEVALEUR').AsString;
      If (TypeValeur='I') then TobSal.PutValue(Champ,StrToInt(Valeur))
      else If (TypeValeur='F') then TobSal.PutValue(Champ,StrToFloat(Valeur))
      else If (TypeValeur='D') then TobSal.PutValue(Champ,StrToDate(Valeur))
      else TobSal.PutValue(Champ,Valeur);
    end;
    Ferme(Q);
  end;
  FreeAndNil(TobChampModif);
end;

end.
