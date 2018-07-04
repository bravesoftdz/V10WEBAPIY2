unit PGVIGSAISABS;
{***********UNITE*************************************************
Auteur  ...... : Paie -MF
Créé le ...... : 21/06/2006
Modifié le ... :
Description .. : Vignette de saisie d'une absence d'un collaborateur
               : Vignette : PG_VIG_SAISABS
               : Tablette : PGPERIODEVIGNETTE
               : Table    : MOTIFABSENCE
               : Vue      : PGMVTABSRESP
Mots clefs ... :
*****************************************************************
PT1  | 29/01/2008 | FLO | Vignette servant dorénavant pour le manager
                          Absence validée par défaut
                          Plus de récapitulatif
}

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
  PG_VIG_SAISABS = class (TGAVignettePaie)
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
  end;

implementation
uses
  HEnt1,
  SysUtils,
  PgCalendrier;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_SAISABS.GetInterface (NomGrille: string): Boolean;
VAr
debabs, FinAbs : TDateTime;
Tob_motifAbs : Tob;
Etab : string;
Nb_j,nb_h : double;
begin
  WarningOk := False;

  Result := inherited GetInterface ('');
  MessageErreur := '';

//  if (GetControlValue('SALARIE') = '') AND(GetControlValue('TYPECONGE') = '') //PT1
//    AND (GetControlValue('DEBUTDJ') = '' ) AND (GetControlValue('FINDJ') = '') then
   If FParam = '' Then
   Begin
//   SetControlValue ('SALARIE',V_PGI.UserSalarie);
    ChargeSalaries (Self, 'SALARIE', 'RESPONSABS'); //PT1
   SetControlValue ('SALARIE', '');
   SetControlValue ('TYPECONGE','RTT');
   SetControlValue ('DEBUTDJ','MAT');
   SetControlValue ('FINDJ','PAM');
   debabs := Date;
   finabs := Date;
   Etab := RechDom('PGSALARIEETAB',V_pgi.UserSalarie,False);

   Tob_motifAbs := Tob.create('Les motifs',nil,-1);
   Tob_motifAbs.LoadDetailDB('MOTIFABSENCE','','',nil,False);

   { Calcul du nombre de jour d'absence }
   PgCalendrier.CalculNbJourAbsence(debabs,finabs,V_Pgi.UserSalarie,Etab,'PRI',Tob_motifAbs,nb_j,nb_h,0,1);
   SetControlValue ('JOURS',FloatToStr(nb_j));
   SetControlValue ('HEURES',FloatToStr(nb_h));
   SetControlValue('LIBELLE',RendLibAbsence('RTT','MAT','PAM',debabs,finabs));
   FreeAndNil(Tob_motifAbs);
   End;

   FreeAndNil(Tob_motifAbs);

end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_SAISABS.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_SAISABS.RecupDonnees;
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

procedure PG_VIG_SAISABS.DrawGrid (Grille: string);
begin
  inherited;

end;

function PG_VIG_SAISABS.SetInterface: Boolean;
begin
 inherited SetInterface;

  result:=true;
end;

procedure PG_VIG_SAISABS.PGValiderAbs;
var
   Tabs,T_Param,T_Sal         : Tob;
   Error : integer;
   NomControl,ComplMessage    : String;
Begin
     If GetControlValue('SALARIE') = '' Then   //PT1
     Begin
        MessageErreur := TraduireMemoire('Vous devez renseigner : ') + TraduireMemoire('le salarié');
        Exit; 
     End;
     
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
       InitialiseTobAbsMotif(Tabs,T_Param,T_Sal,'VAL','-','SAL','VIG');
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
         End;
       End;

     IF Assigned(T_Param)       then FreeAndNil(T_Param);
     IF Assigned(TAbs)          then FreeAndNil(Tabs);
     IF Assigned(T_Sal)         then FreeAndNil(T_Sal);

end;


procedure  PG_VIG_SAISABS.PgCalculJourHeures ();
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

end.

