{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : liste des handicapés présents pour la période
               : Vignette : PG_VIG_HANDICAP
               : Tablette : PGPERIODEVIGNETTE
               : Table    : HANDICAPE, SALARIES
Mots clefs ... :
*****************************************************************}
unit PGVIGHANDICAP;

interface
uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  PGVIGUTIL,
  uToolsOWC,
  uToolsPlugin,
  HCtrls;
type
  PG_VIG_HANDICAP = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeHandicap(AParam : string);

 protected
    procedure RecupDonnees; override;
    function GetInterface (NomGrille: string = ''): Boolean; override;
    procedure GetClauseWhere; override;
    procedure DrawGrid (Grille: string); override;
    function SetInterface : Boolean ; override;
 public
 end;


 implementation
 uses
  SysUtils,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_HANDICAP.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_HANDICAP.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_HANDICAP.RecupDonnees;
begin
  inherited;
    AfficheListeHandicap('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_HANDICAP.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'DATEENTREE', 'C.0 ---');
    SetFormatCol(Grille, 'DATESORTIE', 'C.0 ---');

end;

function PG_VIG_HANDICAP.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
end;


procedure PG_VIG_HANDICAP.AfficheListeHandicap(AParam : string);
var
  StSQL                  : String;
  DataTob , MaTob, TobL  : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : integer;
begin

  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

  Try
    DataTob := TOB.Create('~TEMP', nil, -1);

    StSQL := 'SELECT PSA_LIBELLE,PSA_DATEENTREE,PSA_DATESORTIE, '+
             'PSA_PRENOM, PSA_DATEENTREEPREC,PSA_DATESORTIEPREC '+
             'FROM HANDICAPE '+
             'LEFT JOIN SALARIES '+
             'ON PGH_SALARIE = PSA_SALARIE WHERE '+
             '(PSA_DATEENTREE <= "'+USDATETIME (EnDateAu)+'" AND '+
             '(PSA_DATESORTIE >= "'+USDATETIME (EnDateDu)+'" OR '+
             'PSA_DATESORTIE = "'+USDATETIME(IDate1900)+'")) OR '+
             '(PSA_DATEENTREEPREC <= "'+USDATETIME (EnDateAu)+'" AND '+
             '(PSA_DATESORTIEPREC >= "'+USDATETIME (EnDateDu)+'" OR '+
             'PSA_DATESORTIE = "'+USDATETIME(IDate1900)+'"))'+
             'ORDER BY PSA_DATEENTREE';


   MaTob := OpenSelectInCache (StSQL);
    if  (Matob.Detail.Count <> 0) then
    begin
        TobDonnees := Tob.Create ('LES HANDICAPES', nil, -1);
        for i := 0 to Matob.Detail.Count - 1 do
          begin
            TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
            TobL.AddChampSupValeur ('NOM', MaTob.detail [i].Getstring ('PSA_LIBELLE'));
            TobL.AddChampSupValeur ('PRENOM', MaTob.detail [i].Getstring ('PSA_PRENOM'));
            if  ((MaTob.detail [i] .Getvalue('PSA_DATEENTREE') < EnDateAu) and
                 ((MaTob.detail [i] .Getvalue('PSA_DATESORTIE') >= EnDateDu) or
                  (MaTob.detail [i] .Getvalue('PSA_DATESORTIE') = IDate1900))) then
            begin
              TobL.AddChampSupValeur ('DATEENTREE', MaTob.detail [i] .Getvalue ('PSA_DATEENTREE'),CSTDate);
              TobL.AddChampSupValeur ('DATESORTIE', MaTob.detail [i] .Getvalue ('PSA_DATESORTIE'),CSTDate);
            end
            else
            begin
              TobL.AddChampSupValeur ('DATEENTREE', MaTob.detail [i] .Getvalue ('PSA_DATEENTREEPREC'),CSTDate);
              TobL.AddChampSupValeur ('DATESORTIE', MaTob.detail [i] .Getvalue ('PSA_DATESORTIEPREC'),CSTDate);
            end;
         end;

    end;
//@@    TobDonnees.savetofile('c:\handicap.txt',false,true,true);
    if (TobDonnees.detail.count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
    FreeAndNil (MaTob);
  end;

end;

end.
