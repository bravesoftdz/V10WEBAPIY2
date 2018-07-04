{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSALARIERUB ()
Mots clefs ... : TOF;PGSALARIERUB
*****************************************************************}
Unit UtofPGSALARIERUB ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,       //Agllancefiche
{$else}
     eMul,
     MainEAgl,      //AglLancefiche
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Grids,         
     HTB97,         //TToolBarButton97
     UTOF ;

Const
  clNature            : integer = 0;
  clRubrique          : integer = 1;
  clLibelle           : integer = 2;
  clProfil            : integer = 3;


Type
  TOF_PGSALARIERUB = Class (TOF)
    private
    CodeSalarie       : String;
    TobSalarie        : Tob;
    TobProfil         : Tob;
    TobRubrique       : Tob;
    TobSpeciaux       : Tob;

    function GeneTobProfil : Boolean;
    procedure TraiteProfil ( CodeProfil : String );
    procedure AfficheProfil;

    public
    FListe            : THGrid;
    BValide           : TToolbarButton97;
    BFerme            : TToolbarButton97;
    BInsert           : TToolbarButton97;
    BDelete           : TToolbarButton97;
    BDefaire          : TToolbarButton97;
    BImprime          : TToolbarButton97;



    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

  procedure PgSalarieRub_LanceFiche ( Argument : String = '');

Implementation

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/11/2006
Modifié le ... :   /  /
Description .. : procedure d'appel de la fiche de gestion des profils
Mots clefs ... :
*****************************************************************}
procedure PgSalarieRub_LanceFiche ( Argument : String = '');
begin
  if Trim(Argument) <> '' then
    AGLLanceFiche('PAY', 'SALARIERUB', '', '', Argument);
end;

procedure TOF_PGSALARIERUB.OnArgument (S : String ) ;
var
  Critere             : String;
  Valeur              : String;
  StrFrm              : String;

begin
  Inherited ;
  //les controles
  FListe              := THGrid(GetControl('GRIDRUBRIQUE'));
  BValide             := TToolbarButton97(GetControl('BVALIDER'));
  BFerme              := TToolbarButton97(GetControl('BFERME'));
  BImprime            := TToolbarButton97(GetControl('BIMPRIMER'));
  BInsert             := TToolbarButton97(GetControl('BINSERT'));
  BDelete             := TToolbarButton97(GetControl('BDELETE'));
  BDefaire            := TToolbarButton97(GetControl('BDEFAIRE'));

  // les propriétés

  BValide.Visible               := False;
  BValide.ModalResult           := 0;
  BFerme.Visible                := True;
  BFerme.ModalResult            := 0;
  BImprime.Visible              := False;
  BInsert.Visible               := False;
  BDelete.Visible               := False;
  BDefaire.Visible              := False;

  FListe.ColCount               := 4;
  FListe.RowCount               := 2;
  FListe.FixedRows              := 1;
  FListe.DefaultRowHeight       := 18;
  FListe.ColWidths[clNature]    := 100;
  FListe.ColWidths[clRubrique]  := 90;
  FListe.ColWidths[clLibelle]   := 150;
  FListe.ColWidths[clProfil]    := 150;
  FListe.TwoColors              := True;
  FListe.ScrollBars             := ssBoth;
  FListe.Options                := FListe.Options + [goRowSelect] - [goEditing] + [goTabs] - [goRangeSelect];
  FListe.Cells[clNature, 0]     := TraduireMemoire('Nature');
  FListe.Cells[clRubrique, 0]   := TraduireMemoire('Rubrique');
  FListe.Cells[clLibelle, 0]    := TraduireMemoire('Libellé');
  FListe.Cells[clProfil, 0]     := TraduireMemoire('Profil');
  FListe.ColFormats[clNature]   := 'CB=PGNATURERUBRIQUE||';
  FListe.ColAligns[clRubrique]  := taCenter;
  FListe.SortEnabled            := True;


  //initialisation
  Critere := READTOKENST(S);
  While Critere <> '' do
  begin
    Valeur := Copy (Critere, pos('=', Critere)+1, length(Critere));
    Critere := UpperCase( Copy (Critere, 1, pos('=', Critere)-1));
    if Critere = 'SALARIE' then
      CodeSalarie     := Valeur;

    Critere := READTOKENST(S);
  end;

  TobSalarie          := Tob.Create('SALARIES', nil, -1);
  TobSalarie.SelectDB('"'+CodeSalarie+'"', nil);

  StrFrm              := 'Liste des rubriques pour : ' + CodeSalarie + ' '
                      + Trim(Trim(TobSalarie.GetValue('PSA_LIBELLE')) + ' ' + Trim(TobSalarie.GetValue('PSA_PRENOM')));

  TForm(Ecran).Caption := StrFrm;
  UpdateCaption(TForm(Ecran));

  if GeneTobProfil then
    AfficheProfil;
end ;

procedure TOF_PGSALARIERUB.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGSALARIERUB.OnClose ;
begin
  Inherited ;

  TobSalarie.Free;
  TobProfil.Free;
  TobRubrique.Free;
  TobSpeciaux.Free;
end ;

procedure TOF_PGSALARIERUB.OnCancel () ;
begin
  Inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/11/2006
Modifié le ... :   /  /    
Description .. : Gère les différents profils du salarié
Mots clefs ... : 
*****************************************************************}
function TOF_PGSALARIERUB.GeneTobProfil: Boolean;
Var
  StrSql              : String;
  ListeProfil         : TStringList;
  i                   : integer;

begin
  // Création de la liste des champs de profil dans la fiche salarié
  ListeProfil         := TStringList.Create;
  ListeProfil.Add('PSA_ACTIVITE');        //Activité
  ListeProfil.Add('PSA_PROFILREM');       //rémunération
  ListeProfil.Add('PSA_PROFIL');          //profil ou modèle bulletin
  ListeProfil.Add('PSA_PERIODBUL');       //périodicité plafond
  ListeProfil.Add('PSA_PROFILRBS');       //réduction loi Fillon
  ListeProfil.Add('PSA_REDRTT2');         //Minoration loi Fillon
  ListeProfil.Add('PSA_REDRTT1');         //réduction RTT loi Aubry 2
  ListeProfil.Add('PSA_REDREPAS');        //réduction repas
  ListeProfil.Add('PSA_PROFILAFP');       //abattement frais professionnel
  ListeProfil.Add('PSA_PROFILAPP');       //gestion des appoints
  ListeProfil.Add('PSA_PROFILRET');       //retraite
  ListeProfil.Add('PSA_PROFILMUT');       //cotisation mutuelle
  ListeProfil.Add('PSA_PROFILPRE');       //cotisation prévoyance
  ListeProfil.Add('PSA_PROFILTSS');       //taxe sur salaire
  ListeProfil.Add('PSA_PROFILFNAL');      //FNAL +9 salariés
  ListeProfil.Add('PSA_PROFILTRANS');     //transport +9 salariés


  TobProfil           := Tob.Create('Les profils', nil, -1);
  TobRubrique         := Tob.Create('PROFILPAIE', nil, -1);
  TobSpeciaux         := Tob.Create('PROFILSPECIAUX', nil, -1);
  // chargement des profils disponibles
  StrSql              := 'SELECT * FROM PROFILPAIE WHERE ##PPI_PREDEFINI##';
  TobRubrique.LoadDetailFromSQL(StrSql);
  // chargement des profils spéciaux disponibles pour un salarié
  StrSql              := 'SELECT * FROM PROFILSPECIAUX WHERE PPS_CODE = "' + CodeSalarie + '"';
  TobSpeciaux.LoadDetailFromSQL(StrSql);

  // on scrute la liste des profils du salarié
  for i := 0 to ListeProfil.Count -1 do
  begin
    if TobSalarie.GetValue(ListeProfil.Strings[i]) <> '' then
      TraiteProfil(TobSalarie.GetValue(ListeProfil.Strings[i]));
  end;

  // on scrute la liste des profils spéciaux
  For i := 0 to TobSpeciaux.Detail.Count -1 do
  begin
    TraiteProfil(TobSpeciaux.Detail[i].GetValue('PPS_PROFIL'));
  end;

  // on retourne vrai ou faux selon s'il y a au moins une rubrique
  if TobProfil.Detail.Count -1 > 0 then
    Result              := True
  else
    Result              := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/11/2006
Modifié le ... :   /  /    
Description .. : récupère les données dans PROFILRUB selon le profil
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGSALARIERUB.TraiteProfil(CodeProfil: String);
var
  Tobc                : Tob;
  StrSql              : String;
  i                   : integer;
  Tobr                : Tob;
  LeProfil            : String;

begin
  // on charge les rubriques selon le code profil
  StrSql              := 'SELECT PPM_NATURERUB, PPM_RUBRIQUE, PPM_LIBELLE FROM PROFILRUB WHERE PPM_PROFIL="'
                         + CodeProfil+'" AND ##PPM_PREDEFINI##';
  Tobr                := Tob.Create('Les rubriques', nil, -1);
  Tobr.LoadDetailFromSQL(StrSql);
  if Tobr <> nil then
  begin
    Tobc              := TobRubrique.FindFirst(['PPI_PROFIL'], [CodeProfil], False);
    LeProfil          := Tobc.GetValue('PPI_LIBELLE');

    // on parcours les rubriques du profil et on ajoute à la tob
    for i := 0 to Tobr.Detail.Count -1 do
    begin
      Tobc := Tob.Create('une rubrique', TobProfil, -1);

      Tobc.AddChampSupValeur('NATURE', Tobr.Detail[i].GetValue('PPM_NATURERUB'));
      Tobc.AddChampSupValeur('RUBRIQUE', Tobr.Detail[i].GetValue('PPM_RUBRIQUE'));
      Tobc.AddChampSupValeur('LIBELLE', Tobr.Detail[i].GetValue('PPM_LIBELLE'));
      Tobc.AddChampSupValeur('PROFIL', LeProfil);
    end;
  end;

  Tobr.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/11/2006
Modifié le ... :   /  /    
Description .. : Affiche les rubriques
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGSALARIERUB.AfficheProfil;
begin
  // tri par défaut
  TobProfil.Detail.Sort('NATURE;RUBRIQUE');
  TobProfil.PutGridDetail(FListe, False, False, 'NATURE;RUBRIQUE;LIBELLE;PROFIL');
end;

Initialization
  registerclasses ( [ TOF_PGSALARIERUB ] ) ;
end.
