{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - JL
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : liste des DIF par responsable
               : Vignette : PG_VIG_LISTEDDIF
               : Tablette : PGPERIODEVIGNETTE
               : Table    : FORMATIONS
Mots clefs ... :
*****************************************************************}
unit PGVIGFORMPREV;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_FORMPREV= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure AfficheListeInterim(AParam : string);
  end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_FORMPREV.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_FORMPREV.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_FORMPREV.RecupDonnees;
begin
  inherited;
    AfficheListeInterim('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_FORMPREV.DrawGrid (Grille: string);
begin
  inherited;
   SetFormatCol(Grille, 'DATEDEMANDE', 'C.0 ---');
   SetFormatCol(Grille, 'NBHEURES', 'D.0 ---');
   SetFormatCol(Grille, 'HEURESTT', 'D.0 ---');
   SetFormatCol(Grille, 'HEURESHT', 'D.0 ---');
end;

{-----Affectation des controls/valeurs dans la TobResponse -----------------------------}

function PG_VIG_FORMPREV.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Affiche les données dans la liste ------------------------------------------------}

procedure PG_VIG_FORMPREV.AfficheListeInterim(AParam : string);
var
  StSQL,Where     : String;
  DataTob         : TOB;
  i               : integer;
  T               : TOB;
begin
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);
        If GetControlValue('MILLESIME')<> '' then
            Where :=  'AND PFI_MILLESIME="'+GetControlValue('MILLESIME')+'" '
        else
            Where := '';

        StSQL := 'SELECT PSA_LIBELLE,PSA_PRENOM,PFI_SALARIE,PFI_LIBELLE,PFI_DUREESTAGE,PFI_DATEDIF,PFI_ETATINSCFOR,PST_LIBELLE,'+
                 'PFI_HTPSTRAV,PFI_HTPSNONTRAV FROM INSCFORMATION '+
                 'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+
                 'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE '+
                 'WHERE PFI_REALISE="-" AND PFI_TYPEPLANPREV<>"DIF" AND PFI_CODESTAGE<>"--CURSUS--" '+
                 'AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" '+Where+
                 'ORDER BY PSA_LIBELLE';

        DataTob := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTob);

        for i:=0 to DataTob.detail.count-1 do
        begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            If DataTob.Detail[i].GetValue('PFI_SALARIE') <> '' Then
            Begin
                T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
                T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
            End
            Else
            Begin
                T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PFI_LIBELLE'));
                T.AddChampSupValeur('PRENOM','');
            End;
            T.AddChampSupValeur('FORMATION',DataTob.Detail[i].GetValue('PST_LIBELLE'));
            T.AddChampSupValeur('ETAT',RechDom('PGETATVALID',DataTob.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            T.AddChampSupValeur('NBHEURES',DataTob.Detail[i].GetValue('PFI_DUREESTAGE'));
            T.AddChampSupValeur('HEURESTT',DataTob.Detail[i].GetValue('PFI_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHT',DataTob.Detail[i].GetValue('PFI_HTPSNONTRAV'));
        end;

        if (TobDonnees.detail.count = 0) then
            PutGridDetail('FListe', TobDonnees);
    finally
        FreeAndNil (DataTob);
    end;
end;

end.


