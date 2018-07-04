{***********UNITE*************************************************
Auteur  ...... : Paie - FLO
Créé le ...... : 29/01/2008
Modifié le ... :
Description .. : Planning des absences du salarié
               : Vignette : PG_VIG_CALABSSAL
               : Table    : ABSENCESALARIE
Mots clefs ... :
*****************************************************************
}

unit PGVIGCALABSSAL;

interface

uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_CALABSSAL = class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    Procedure ChargeCalendrier (Periode : TDateTime; var TobCalendrier : TOB);
  end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  HEnt1,
  PGCalendrier,
  PGOutilsFormation,
  DateUtils;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_CALABSSAL.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;

    // Initialisation de la vignette au niveau des périodes pour que ce soit toujours un mois
    If FParam = '' Then
    Begin
        Periode := 'M';
        MajLibPeriode();
    End;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_CALABSSAL.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_CALABSSAL.RecupDonnees;
var tob_sem, tob_std, tob_jf,tob_abs, TobCalendrier:tob;
  Etab,Calend,StandCalend,EtabStandCalend,NbJrTravEtab,Repos1,Repos2,JourHeure : string;
  DateEntree,DateSortie : TDateTime;
  SN1              : String;
  EnDateDu, EnDateAu     : TDatetime;
begin
  inherited;

//  If FParam <> 'CELLULE' Then
//  Begin
      // Récupération de la période sélectionnée
      DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, SN1);

      // Charge le calendrier général du mois
      ChargeCalendrier (EnDateDu, TobCalendrier);

      // Charge le calendrier du salarié
      ChargeTobCalendrier(EnDateDu,EnDateAu,V_PGI.UserSalarie,True,False,False,tob_sem,tob_std,tob_jf,tob_abs,
      Etab,Calend,StandCalend,EtabStandCalend,NbJrTravEtab,Repos1,Repos2,JourHeure,DateEntree,DateSortie);

      GetControlProperty('GRSALARIE', 'COL');

      // Constitution de la vignette
      TobDonnees.Dupliquer(TobCalendrier, True, True);

      FreeAndNil(TobCalendrier);
//  End;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_CALABSSAL.DrawGrid (Grille: string);
begin
  inherited;
    SetVisibleCol(Grille, 'NUMSEMAINE', False);
    SetFormatCol (Grille, '1', 'C.0 ---');
    SetFormatCol (Grille, '2', 'C.0 ---');
    SetFormatCol (Grille, '3', 'C.0 ---');
    SetFormatCol (Grille, '4', 'C.0 ---');
    SetFormatCol (Grille, '5', 'C.0 ---');
    SetFormatCol (Grille, '6', 'C.0 ---');
    SetFormatCol (Grille, '7', 'C.0 ---');
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_CALABSSAL.SetInterface: Boolean;
begin
 inherited SetInterface;
    result:=true;
end;

{-----Chargement initial du calendrier -------------------------------------------------}

Procedure PG_VIG_CALABSSAL.ChargeCalendrier (Periode : TDateTime; var TobCalendrier : TOB);
var Fin : TDateTime;
    FYear, FMonth, FDay : Word;
    NbJours : Integer;
    TS, T : TOB;
    i,NumSemaine,DayOf : Integer;

    Procedure CreeSemaine (NumSemaine : Integer);
    Var j : Integer;
    Begin
            TS := TOB.Create('$LaSemaine',TobCalendrier,-1);
            TS.AddChampSupValeur('NUMSEMAINE',IntToStr(NumSemaine));
            For j := 1 To 7 Do TS.AddChampSup(IntToStr(j),False);
    End;
Begin
    TobCalendrier := TOB.Create('$LesJours', Nil, -1);
    Fin   := FinDeMois(Periode);

    DecodeDate(Fin, FYear, FMonth, FDay);

    NbJours := DaysPerMonth(FYear, FMonth);

    // Renseignement des n° de jours
    NumSemaine := 1;
    For i := 1 To NbJours Do
    Begin
        If i = 1 Then CreeSemaine(NumSemaine);

        FDay := i;
        DayOf := DayOfTheWeek(EncodeDate(FYear,FMonth,FDay));

        If (DayOf = dayMonday) And (i > 1) Then
        Begin
            NumSemaine := NumSemaine + 1;
            CreeSemaine(NumSemaine);
        End;

        T := TobCalendrier.FindFirst(['NUMSEMAINE'],[NumSemaine],False);
        If T <> Nil Then
        Begin
            T.PutValue(IntToStr(DayOf),IntToStr(i));
        End;
    End;
End;

end.


