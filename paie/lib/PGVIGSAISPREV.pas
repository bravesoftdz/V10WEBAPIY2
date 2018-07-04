unit PGVIGSAISPREV;
{***********UNITE*************************************************
Auteur  ...... : Paie -JL
Créé le ...... : 13/02/2007
Modifié le ... :
Description .. : Saisie DIF
               : Vignette : PG_VIG_SAISPREV
               : Tablette : PGPERIODEVIGNETTE
               : Table    : INSCFORMATION
Mots clefs ... :
*****************************************************************
PT1  | 18/01/2008 | FLO | Externalisation de l'inscription et possibilité d'inscrire en non nominatif
PT2  | 18/03/2008 | FLO | Prise en compte du type d'utilisateur connecté
PT3  | 15/04/2008 | FLO | Correction de l'affectation du temps de travail
PT4  | 30/04/2008 | FLO | Correction du chargement des salariés pour rajouter le non nominatif qui avait disparu
PT5  | 15/05/2008 | FLO | Rafraîchissement des vignettes associés
}

interface

uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_SAISPREV = class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure PGValiderInsc;
    procedure AfficheHeures ;
    procedure InitialiseSaisie;
    Procedure ChargeDonnees;
    procedure ChangeSalarie ;  //PT1
  end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  HEnt1,
  ParamSoc,
  PGCalendrier,
  PGOutilsFormation, uAncetreVignettePlugIn, uToolsPlugin;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_SAISPREV.GetInterface (NomGrille: string): Boolean;
begin
    WarningOk := False;
    Result := inherited GetInterface ('');
    MessageErreur := '';
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_SAISPREV.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_SAISPREV.RecupDonnees;
begin
  inherited;
    If ParamFich = '' then
    Begin
        ChargeDonnees;
        InitialiseSaisie;
    End
    Else If ParamFich = 'VALIDER' then
        PGValiderInsc ()
    Else if (ParamFich = 'STAGE') then
        AfficheHeures
    Else If (ParamFich = 'SALARIE') Then    //PT1
        ChangeSalarie;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_SAISPREV.DrawGrid (Grille: string);
begin
  inherited;
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_SAISPREV.SetInterface: Boolean;
begin
 inherited SetInterface;
    result:=true;
end;

{-----Initialisation de la vignette ----------------------------------------------------}

procedure PG_VIG_SAISPREV.InitialiseSaisie;
begin
     SetControlValue('TYPEPLAN','DIF');
     SetControlValue('SALARIE','');
     SetControlValue('STAGE','');
     SetControlValue('LIBEMPLOI','');
     SetControlValue('NBINSC',  '0');
     SetControlValue('HEURESTT','0');
     SetControlValue('HEURESHT','0');
     SetControlVisible('PANNONAME', False);  //PT1
end;

{-----Chargement des salariés sous la responsabilité de l'utilisateur ------------------}

procedure PG_VIG_SAISPREV.ChargeDonnees;
begin
    ChargeSalaries(Self, 'SALARIE', RecupTypeUtilisateur('FOR'), True, '', False); //PT2 //PT4
    ChargeStages  (Self, 'STAGE');
    ChargeEmplois (Self, 'LIBEMPLOI', RecupTypeUtilisateur('FOR'));      //PT1 //PT2
end;

{-----Validation d'une inscription au prévisionnel -------------------------------------}

procedure PG_VIG_SAISPREV.PGValiderInsc;
var
   Salarie : String;
   TRefresh : TOB;
Begin
    // Contrôles de saisie
    If (GetControlValue('SALARIE') = '') Or (GetControlValue('STAGE') = '') Or
       ((GetControlValue('SALARIE') = UNDEFINED) And (GetControlValue('LIBEMPLOI')='')) Then
    Begin
        MessageErreur := TraduireMemoire('Toutes les zones sont obligatoires.');
        Exit;
    End;

    Salarie := GetControlValue('SALARIE');
    If Salarie = UNDEFINED Then Salarie := '';

    //PT1
    MessageErreur := InscritSalariePrevisionnel (GetControlValue('STAGE'), RendMillesimePrevisionnel(),
                                                 Salarie, GetControlValue('NBINSC'),
                                                 GetControlValue('TYPEPLAN'), GetControlValue('LIBEMPLOI'),
                                                 StrtoFloat(GetControlValue('HEURESTT')), StrtoFloat(GetControlValue('HEURESHT')), False, RecupTypeUtilisateur('FOR'));
	//PT5
	If MessageErreur = '' Then
	Begin
        // Rafraîchissement des vignettes Liste du prévisionnel, Demandes de DIF et Compteurs DIF
  		TRefresh := TOB.Create('T_REFRESH', TobResponse, -1);
  		If GetControlValue('TYPEPLAN') = 'DIF' Then
  			TRefresh.AddChampSupValeur('NAMES','PG_VIG_LISTEDDIF;PG_VIG_COMPTDIF')
  		Else
  			TRefresh.AddChampSupValeur('NAMES','PG_VIG_LISTEPREV');
	End
	Else Exit;

    InitialiseSaisie;
end;

{-----Renseignement du nombre d'heures de la formation ---------------------------------}

procedure  PG_VIG_SAISPREV.AfficheHeures ();
var
   Stage, Champ     : String;
   Q                : TQuery;
   Heures           : Double;
Begin
    Stage := GetControlValue('STAGE');

    Try
        Q := OpenSQL('SELECT PST_DUREESTAGE FROM STAGE WHERE PST_CODESTAGE="'+Stage+'" AND PST_MILLESIME="0000"', True);
        If Not Q.EOF Then
        Begin
            Heures := Q.FindField('PST_DUREESTAGE').AsFloat;
            If (GetControlValue('TYPEPLAN') = 'DIF') And (GetParamSoc('SO_PGDIFTPSTRAV') = False) Then //PT3
                Champ := 'HEURESHT'
            Else
                Champ := 'HEURESTT';
            SetControlValue(Champ, FloatToStr(Heures));
        End
        Else
        Begin
            SetControlValue('HEURESTT', '0');
            SetControlValue('HEURESHT', '0');
        End;
        Ferme (Q);
    Except
    End;
end;

{-----Chargement des salariés sous la responsabilité de l'utilisateur ------------------}
//PT1 - Début
procedure PG_VIG_SAISPREV.ChangeSalarie;
begin
    If GetControlValue('SALARIE') = UNDEFINED Then
    Begin
        SetControlVisible('PANNONAME', True);
    End
    Else
    Begin
        SetControlVisible('PANNONAME', False);
        SetControlValue('NBINSC', '1');
        SetControlValue('LIBEMPLOI', '');
    End;
end;
//PT1 - Fin

end.


