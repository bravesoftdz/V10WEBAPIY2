unit PGVIGSAISABSSAL;
{***********UNITE*************************************************
Auteur  ...... : Paie -MF
Créé le ...... : 21/06/2006
Modifié le ... :
Description .. : Vignette de saisie des absences par le salarié
               : Vignette : PG_VIG_SAISABSSAL
               : Tablette : PGPERIODEVIGNETTE
               : Table    : MOTIFABSENCE
               : Vue      : PGMVTABSRESP
Mots clefs ... :
*****************************************************************}

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  uToolsOWC,
  PGVignettePaie,
  PGVIGUTIL,HMsgBox,
  HCtrls;

type
  PG_VIG_SAISABSSAL = class (TGAVignettePaie)
  private
    TypeConge           : string;
  protected
    procedure RecupDonnees; override;
    function  GetInterface (NomGrille: string = ''): Boolean; override;
    procedure GetClauseWhere; override;
    procedure DrawGrid (Grille: string); override;
    function  SetInterface : Boolean ; override;
  public
    procedure PGValiderAbs;
    procedure PgCalculJourHeures ;
    procedure PGAffectRecapitulatif;
  end;

implementation
uses
  HEnt1,
  SysUtils,
  PgCalendrier;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_SAISABSSAL.GetInterface (NomGrille: string): Boolean;
VAr
debabs, FinAbs : TDateTime;
Tob_motifAbs : Tob;
Etab : string;
Nb_j,nb_h : double;
begin
  WarningOk := False;

  Result := inherited GetInterface ('');
  MessageErreur := '';

  if (GetControlValue('SALARIE') = '') AND(GetControlValue('TYPECONGE') = '')
    AND (GetControlValue('DEBUTDJ') = '' ) AND (GetControlValue('FINDJ') = '') then
   Begin
   SetControlValue ('SALARIE',V_PGI.UserSalarie);
   SetControlValue ('TYPECONGE','PRI');
   SetControlValue ('DEBUTDJ','MAT');
   SetControlValue ('FINDJ','PAM');
   Tob_motifAbs := Tob.create('Les motifs',nil,-1);
   Tob_motifAbs.LoadDetailDB('MOTIFABSENCE','','',nil,False);
   Etab := RechDom('PGSALARIEETAB',V_pgi.UserSalarie,False);

   debabs := StrToDate(GetControlValue('DATEDEBUTABS')) ;
   finabs := StrToDate(GetControlValue('DATEFINABS'));
   { Calcul du nombre de jour d'absence }
   PgCalendrier.CalculNbJourAbsence(debabs,finabs,V_Pgi.UserSalarie,Etab,'PRI',Tob_motifAbs,nb_j,nb_h,0,1);
   SetControlValue ('JOURS',FloatToStr(nb_j));
   SetControlValue ('HEURES',FloatToStr(nb_h));
   SetControlValue('LIBELLE',RendLibAbsence('PRI','MAT','PAM',debabs,finabs));
   FreeAndNil(Tob_motifAbs);
   End;

  if (GetControlValue('LBRESTCP') = '0') or (GetControlValue('LBRESTRTT') = '0')
     or (GetControlValue('LBATTCP') = '0') or (GetControlValue('LBATTRTT') = '0')
     or (GetControlValue('LBVALCP') = '0') or (GetControlValue('LBATTRTT') = '0') then
     Begin
     PGAffectRecapitulatif;
     End;
   FreeAndNil(Tob_motifAbs);

end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_SAISABSSAL.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_SAISABSSAL.RecupDonnees;
begin
  inherited;
    if Paramfich = 'Valider' then
       PGValiderAbs ()
    else
        if  (Paramfich = 'debabs') or (Paramfich = 'finabs')
        or (Paramfich = 'DebutDj') or (Paramfich = 'FinDj')
        or (Paramfich = 'TypeConge') then  PgCalculJourHeures ;

end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_SAISABSSAL.DrawGrid (Grille: string);
begin
  inherited;

end;

function PG_VIG_SAISABSSAL.SetInterface: Boolean;
begin
 inherited SetInterface;

  result:=true;
end;

procedure PG_VIG_SAISABSSAL.PGValiderAbs;
var
   Tabs,T_Param,T_Sal         : Tob;
   Error : integer;
   NomControl,ComplMessage    : String;
Begin
     { Chargement de l'absence saisie }
     T_Param := Tob.Create('Absence saisi',nil,-1);
     T_Param.AddChampSupValeur('PCN_SALARIE',GetControlValue('SALARIE'));
     T_Param.AddChampSupValeur('PCN_TYPECONGE',GetControlValue('TYPECONGE'));
     T_Param.AddChampSupValeur('PCN_DATEDEBUTABS',StrToDate(GetControlValue('DATEDEBUTABS')));
     T_Param.AddChampSupValeur('PCN_DATEFINABS',StrToDate(GetControlValue('DATEFINABS')));
     T_Param.AddChampSupValeur('PCN_DATEVALIDITE',StrToDate(GetControlValue('DATEFINABS')));
     T_Param.AddChampSupValeur('PCN_DEBUTDJ' ,GetControlValue('DEBUTDJ'));
     T_Param.AddChampSupValeur('PCN_FINDJ'   ,GetControlValue('FINDJ'));
     T_Param.AddChampSupValeur('PCN_JOURS'   ,GetControlValue('JOURS'));
     T_Param.AddChampSupValeur('PCN_HEURES'  ,GetControlValue('HEURES'));
     T_Param.AddChampSupValeur('PCN_LIBELLE' ,GetControlValue('LIBELLE'));
     T_Param.AddChampSupValeur('PCN_LIBCOMPL1' ,GetControlValue('LIBCOMPL1'));
     T_Param.AddChampSupValeur('PCN_LIBCOMPL2' ,GetControlValue('LIBCOMPL2'));

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
       InitialiseTobAbsMotif(Tabs,T_Param,T_Sal,'ATT','-','SAL','VIG');
       { Contrôle la possibilité de l'enregistrement de l'absence }
       PgControleTobAbsSecond(TAbs,T_Sal,'E','CREATION','SAL','SAL',Error,NomControl,ComplMessage);
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


procedure  PG_VIG_SAISABSSAL.PgCalculJourHeures ();
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


procedure PG_VIG_SAISABSSAL.PgAffectRecapitulatif;
Var
  T_Recap : Tob;
begin
 { Chargement du recap et affectation des valeurs }
 T_Recap := ChargeTob_Recapitulatif(V_PGI.UserSalarie);
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

end.

