{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 04/07/2007
Modifié le ... :   /  /
Description .. : Planning des compteurs de présence
Mots clefs ... : PLANNING;PRESENCE;COMPTEURS
*****************************************************************}
{
PT1 25/10/2007 GGU V_80 On évite les doublons dans la légende et on affiche une 3e et une 4e ligne.
}
unit PGPlanningPresenceCompteurs;

interface

uses
     StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     Windows, Messages, Graphics,
     HCtrls, UTOB,
     hplanning, PGPlanningTemplate, PgPlanningOutils;

Const
  SALARIE = 01;
  GLOBAL  = 02;
  SALARIE_KEY_LENGTH = 10;

type

   //PT1
  TLegende = class
  private
  public
    GBRecap:  THGroupBox;
    ALibelle: Array of THLabel;
    ACouleur: Array of THLabel;
    ATitre  : Array of String;
    NbrParLigne : Integer;
    Constructor Create(NbParLigne : Integer = 5);
    Destructor Destroy; Override;
    Function Add(Lib, Couleur : THLabel; Titre : String) : Integer;
    Function Exist(Titre : String) : Boolean;
    Procedure GetLTWH(var vLeft, vTop, vWidth, vHeight : Integer; Index : Integer; Labl : Boolean);
  end;

  TPlanningPresenceCompteurs = class(TFormPlanning)
  private
    StWhTobEtats,StWhTobRessource : String;
    DebutPlanning, FinPlanning : TDateTime;
    Rupture: TNiveauRupture;
    Legende : TLegende; //PT1
    // Détails
//PT1
{    GBRecap:  THGroupBox;
    LblC1: THLabel;LblL1: THLabel;
    LblC2: THLabel;LblL2: THLabel;
    LblC3: THLabel;LblL3: THLabel;
    LblC4: THLabel;LblL4: THLabel;
    LblC5: THLabel;LblL5: THLabel;
    LblC6: THLabel;LblL6: THLabel;
    LblC7: THLabel;LblL7: THLabel;
    LblC8: THLabel;LblL8: THLabel;
    LblC9: THLabel;LblL9: THLabel;
    LblC10: THLabel;LblL10: THLabel;  }
    Procedure SpecifPlanningProperties(Planning : THPlanning; Sender : TObject);
    procedure LoadEtats(Deb, Fin: TDateTime);
    procedure LoadItems(Deb, Fin: TDateTime);
    procedure LoadRessources(Deb, Fin: TDateTime);
    procedure GereOptionPopup(Item: TOB; ZCode: Integer; var ReDraw: Boolean; IdRessource : String; CurrentDate : TDateTime; ZoneSel : TZoneSelected);
    Procedure AfficheLegende;
  end;

  var Tablettes : Array of Array[0..1] of String;
  
  Function RechTablette (NomChamp : String) : String;
  procedure PGPlanningPresenceCompteursOpen(DateDeb,DateFin : TDateTime ;WhereTobEtats, WhereTobRessources : string; NiveauRupt: TNiveauRupture; ListDesIntervales : TPlanningIntervalList);
  function ConvertPrefixe(StWhere, DePref, APref: string): string;

implementation

uses
{$IFDEF EAGLCLIENT}
  UtileAGL,eMul,MaineAgl,HStatus,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
  DateUtils, HMsgBox, EntPaie, HEnt1;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/07/2007
Modifié le ... :   /  /    
Description .. : Fonction de remplissage d'une chaîne
Mots clefs ... :
*****************************************************************}
Function BourreChaine ( St : String ; LgCompte : integer ) : string ;
var ll,i : Integer ;
    Bourre  : Char ;
begin
  Bourre:='0';
  Result:=St ; ll:=Length(Result) ;
  If ll<LgCompte then for i:=ll+1 to LgCompte do Result:=Result+Bourre;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : Ouverture du planning
Mots clefs ... : 
*****************************************************************}
procedure PGPlanningPresenceCompteursOpen(DateDeb,DateFin : TDateTime ;WhereTobEtats, WhereTobRessources : string; NiveauRupt: TNiveauRupture; ListDesIntervales : TPlanningIntervalList);
var
  Planning : TPlanningPresenceCompteurs;
Begin
     try
          // Création du planning
          Planning := TPlanningPresenceCompteurs.Create(Application);

          // Sécurité
          If Planning = Nil Then Exit;

          with Planning do
          begin
              { Affectation des évènements }
               OnCreatePlanning       := SpecifPlanningProperties;
               OnPopup                := GereOptionPopup;
               LoadItemsProcedure     := LoadItems;
               LoadEtatsProcedure     := LoadEtats;
               LoadRessourcesProcedure:= LoadRessources;


               { Date de début et de fin de tous les plannings et liste des différents plannings}
               AllPlanIntervalDebut   := DateDeb; DebutPlanning := DateDeb;
               AllPlanIntervalFin     := DateFin; FinPlanning   := DateFin;
               ListIntervales         := ListDesIntervales;

               { Clauses Where à appliquer }
               StWhTobEtats           := WhereTobEtats;
               StWhTobRessource       := WhereTobRessources;

               { Récupération des ruptures }
               Rupture := NiveauRupt;

               Planning.Caption       := 'Planning des compteurs de présence';

               { Récapitulatif }
               //PT1
               Legende := TLegende.Create;
               CreateDetailsGroupBox( Legende.GBRecap , 1, 1, 906, 80, 'Légende');
//PT1
{               CreateDetailsGroupBox(GBRecap, 1, 1, 906, 80, 'Légende');

               CreateDetailsLabel(LblC1,  GBRecap, 10,  20, 35,  20, '');
               CreateDetailsLabel(LblL1,  GBRecap, 55,  23, 165, 13, '', False);
               CreateDetailsLabel(LblC3,  GBRecap, 235, 20, 35,  20, '');
               CreateDetailsLabel(LblL3,  GBRecap, 280, 23, 165, 13, '', False);
               CreateDetailsLabel(LblC5,  GBRecap, 460, 20, 35,  20, '');
               CreateDetailsLabel(LblL5,  GBRecap, 505, 23, 165, 13, '', False);
               CreateDetailsLabel(LblC7,  GBRecap, 685, 20, 35,  20, '');
               CreateDetailsLabel(LblL7,  GBRecap, 730, 23, 165, 13, '', False);
               CreateDetailsLabel(LblC9,  GBRecap, 910, 20, 35,  20, '');
               CreateDetailsLabel(LblL9,  GBRecap, 955, 23, 165, 13, '', False);

               CreateDetailsLabel(LblC2,  GBRecap, 10,  45, 35,  20, '');
               CreateDetailsLabel(LblL2,  GBRecap, 55,  48, 165, 13, '', False);
               CreateDetailsLabel(LblC4,  GBRecap, 235, 45, 35,  20, '');
               CreateDetailsLabel(LblL4,  GBRecap, 280, 48, 165, 13, '', False);
               CreateDetailsLabel(LblC6,  GBRecap, 460, 45, 35,  20, '');
               CreateDetailsLabel(LblL6,  GBRecap, 505, 48, 165, 13, '', False);
               CreateDetailsLabel(LblC8,  GBRecap, 685, 45, 35,  20, '');
               CreateDetailsLabel(LblL8,  GBRecap, 730, 48, 165, 13, '', False);
               CreateDetailsLabel(LblC10, GBRecap, 910, 45, 35,  20, '');
               CreateDetailsLabel(LblL10, GBRecap, 955, 48, 165, 13, '', False);
}

               { Affichage du planning }
               ShowPlanning;
          End;
     Finally
          if Assigned(Planning.Legende) then FreeAndNil(Planning.Legende);
          if Assigned(Planning) then FreeAndNil(Planning);
          // Suppression des tablettes
          SetLength(Tablettes, 0);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : Chargement des ressources : Salariés
Mots clefs ... : 
*****************************************************************}
procedure TPlanningPresenceCompteurs.LoadRessources(Deb, Fin : TDateTime);
Var
  StSQLLoadRessources, stChampsRupture : String;
  i,j : Integer;
  T : TOB;
  NumSalarie,Key : String;
  Idem : Integer;
begin
     // Récupération des ruptures
     stChampsRupture := ListeChampRupt(',',Rupture);

     // Requête de base
     StSQLLoadRessources := 'SELECT DISTINCT '+stChampsRupture+' PYP_SALARIE,PSA_LIBELLE || " " || PSA_PRENOM AS NOM,PYR_LIBELLE ' +
                            'FROM PGPRESENCESAL ' +
                            'WHERE PYP_DATEDEBUTPRES >= "'+UsDateTime(Deb)+'" AND PYP_DATEFINPRES<="'+UsDateTime(Fin)+'" '+
                            'AND PYR_LIBELLE IN (SELECT PYR_LIBELLE FROM COMPTEURPRESENCE '+
                            'WHERE ##PYR_PREDEFINI## AND PYR_DATEVALIDITE <= "'+UsDateTime(Fin)+'" AND PYR_EDITPLANPRES="X"';
     If StWhTobEtats <> '' Then StSQLLoadRessources := StSQLLoadRessources + ' AND ' + ConvertPrefixe(StWhTobEtats, 'PYP', 'PYR');
     StSQLLoadRessources := StSQLLoadRessources + ')';

     // Critères appliqués aux ressources : Salariés
     If StWhTobRessource <> '' Then StSQLLoadRessources := StSQLLoadRessources + ' AND ' + StWhTobRessource;

     // Order by
     StSQLLoadRessources := StSQLLoadRessources + ' ORDER BY '+stChampsRupture+' PYP_SALARIE';

     // Chargement des données
     TobRessources.LoadDetailDBFromSQL('LesRessources', StSQLLoadRessources);

     // Création d'une clé composée du salarié et du compteur pour avoir une ligne par couple
     // et mise en forme des ressources pour ne pas avoir de doublon de données à l'affichage
     For i := 0 To TobRessources.Detail.Count - 1 Do
     Begin
          T := TobRessources.Detail[i];

          NumSalarie := T.GetValue('PYP_SALARIE');
          T.AddChampSupValeur('LIB_PYP_SALARIE', NumSalarie);    // On double le code salarié pour garder la référence,
                                                                 // LIB étant utilisé à l'affichage et pouvant être vidé pour la présentation

          // Construction de la clé comprenant les niveaux de ruptures, le salarié et le compteur
          Key := BourreChaine(NumSalarie, SALARIE_KEY_LENGTH) + T.GetValue('PYR_LIBELLE');
          For j := 1 To Rupture.NiveauRupt Do
               Key := Key + T.GetValue(Rupture.ChampsRupt[j]);
          T.AddChampSupValeur('KEY', Key);

          // Mise en forme
          Idem := 0;
          For j := 1 To Rupture.NiveauRupt Do
          Begin
               If (i > 0) And (Idem = (j - 1)) And (T.GetValue(Rupture.ChampsRupt[j]) = TobRessources.Detail[i-1].GetValue(Rupture.ChampsRupt[j])) Then
               Begin
                    T.AddChampSupValeur('LIB_'+Rupture.ChampsRupt[j], '');
                    Idem := Idem + 1;
               End
               Else
                    T.AddChampSupValeur('LIB_'+Rupture.ChampsRupt[j], RechDom(RechTablette(Rupture.ChampsRupt[j]),T.GetValue(Rupture.ChampsRupt[j]), False));
          End;

          If (i > 0) And (Idem = Rupture.NiveauRupt) And (T.GetValue('PYP_SALARIE') = TobRessources.Detail[i-1].GetValue('PYP_SALARIE')) Then
          Begin
               T.PutValue('LIB_PYP_SALARIE', '');
               T.PutValue('NOM', '');
          End;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 04/07/2007
Modifié le ... :   /  /
Description .. : Chargement des états : types de compteurs
Mots clefs ... : 
*****************************************************************}
procedure TPlanningPresenceCompteurs.LoadEtats(Deb, Fin : TDateTime);
Var
  StSQLLoadEtats : String;
  i : Integer;
  Color : TColor;
begin
     // Requête de base
     StSQLLoadEtats := 'SELECT PYR_COMPTEURPRES AS COMPTEUR, PYR_DATEVALIDITE, PYR_LIBELLE, PYR_PERIODICITEPRE, PYR_PGCOLORPRE, PYR_THEMEPRE '+
                       'FROM COMPTEURPRESENCE WHERE ##PYR_PREDEFINI## AND PYR_DATEVALIDITE <= "'+UsDateTime(Fin)+'" AND PYR_EDITPLANPRES="X"';

     // Critères appliqués aux états : Compteurs
     If StWhTobEtats <> '' Then
          StSQLLoadEtats := StSQLLoadEtats + ' AND ' + ConvertPrefixe(StWhTobEtats, 'PYP', 'PYR');

     // Order by
     StSQLLoadEtats := StSQLLoadEtats + ' ORDER BY PYR_COMPTEURPRES, PYR_LIBELLE';

     // Chargement des données
     TobEtats.LoadDetailDBFromSQL('LesCompteurs',StSQLLoadEtats);

     // Mise en forme pour l'abrégé (libellé affiché dans le planning)
     If (TobEtats.Detail <> Nil) And (TobEtats.Detail.Count > 0) Then
     Begin
          TobEtats.Detail[0].AddChampSupValeur('FONTSTYLE',FontStyle, true);
          TobEtats.Detail[0].AddChampSupValeur('FONTSIZE', FontSize,  true);
          TobEtats.Detail[0].AddChampSupValeur('FONTNAME', 'Arial',   true);
          For i:=0 To TobEtats.Detail.Count-1 Do
          Begin
               // Pour la couleur, on calcule le complément par rapport au fond afin d'avoir toujours qqc de lisible
               Color := StringToColor(TobEtats.Detail[i].GetValue('PYR_PGCOLORPRE'));
               If Color > $00888888 Then Color := $00000000 Else Color := $00FFFFFF;
               TobEtats.Detail[i].AddChampSupValeur('FONTCOLOR',ColorToString(Color),False);
          End;
     End;

     AfficheLegende;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : Chargement des items (données) : compteurs
Mots clefs ... : 
*****************************************************************}
procedure TPlanningPresenceCompteurs.LoadItems(Deb, Fin : TDateTime);
Var
  StSQLLoadDonnees : String;
  stChampsRupture  : String;
  i,j              : Integer;
  T                : TOB;
  Libelle,Key      : String;
begin

     stChampsRupture := ListeChampRupt(',',Rupture);

     // Requête de base
     StSQLLoadDonnees := 'SELECT '+stChampsRupture+' PYP_SALARIE, PYP_DATEDEBUTPRES, PYP_DATEFINPRES, PYP_COMPTEURPRES AS COMPTEUR, PYP_THEMEPRE, PYP_QUANTITEPRES, PYR_LIBELLE '+
                         'FROM PGPRESENCESAL WHERE PYP_DATEDEBUTPRES >= "'+UsDateTime(Deb)+'" AND PYP_DATEFINPRES <= "'+UsDateTime(Fin)+'"';

     // Critères appliqués aux ressources
     If StWhTobRessource <> '' Then StSQLLoadDonnees := StSQLLoadDonnees + ' AND ' + StWhTobRessource;

     { Cas particulier pour les compteurs : comme il faut que les références entre les compteurs de la TOB Etat
       et la TOB Donnees soient les mêmes, on ne sélectionne que les données dont le compteur appartient à la liste
       des états }
     StSQLLoadDonnees := StSQLLoadDonnees + ' AND PYP_COMPTEURPRES IN (SELECT PYR_COMPTEURPRES '+
                       'FROM COMPTEURPRESENCE WHERE ##PYR_PREDEFINI## AND PYR_DATEVALIDITE <= "'+UsDateTime(Fin)+'" AND PYR_EDITPLANPRES="X"';
     If StWhTobEtats <> '' Then StSQLLoadDonnees := StSQLLoadDonnees + ' AND ' + ConvertPrefixe(StWhTobEtats, 'PYP', 'PYR');
     StSQLLoadDonnees := StSQLLoadDonnees + ')';

     // Chargement des données
     TobItems.LoadDetailDBFromSQL('LesItems', StSQLLoadDonnees);

     For i := 0 To TobItems.Detail.Count - 1 Do
     Begin
          T := TobItems.Detail[i];

          Key := BourreChaine(T.GetValue('PYP_SALARIE'), SALARIE_KEY_LENGTH) + T.GetValue('PYR_LIBELLE');
          For j := 1 To Rupture.NiveauRupt Do
               Key := Key + T.GetValue(Rupture.ChampsRupt[j]);

          // Création d'une clé composée du salarié et du compteur pour avoir une ligne par couple
          T.AddChampSupValeur('KEY', Key);

          // Personnalisation du champ à afficher dans la bulle d'aide
          Libelle :=            'Thème      : ' + RechDom('PGTHEMECOMPTEURPRES',T.GetValue('PYP_THEMEPRE'),False);
          Libelle := Libelle + '#13#10Compteur : ' + T.GetValue('PYR_LIBELLE');
          T.AddChampSupValeur('HINT', Libelle);

          // Personnalisation du libellé affiché
          Libelle := '  ' + FloatToStr(T.GetValue('PYP_QUANTITEPRES'));
          T.AddChampSupValeur('CAPTION', Libelle);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 04/07/2007
Modifié le ... :   /  /    
Description .. : Spécifie les données "clées" à utiliser pour l'affichage des 
Suite ........ : données dans le planning à partir des TOBs
Mots clefs ... : 
*****************************************************************}
procedure TPlanningPresenceCompteurs.SpecifPlanningProperties(Planning: THPlanning; Sender: TObject);
Var i : Integer;
begin
  with Planning do
  begin
    { Ressources }
    ResChampID     := 'KEY';

    { Items }
    ChampLineID    := 'KEY';
    ChampdateDebut := 'PYP_DATEFINPRES';
    ChampDateFin   := 'PYP_DATEFINPRES';
    ChampEtat      := 'COMPTEUR';            // Référence identique à celle de l'état |
    ChampLibelle   := 'CAPTION';                                                   // v
    ChampHint      := 'HINT';

    { Etats }                                                                      // ^
    EtatChampCode    := 'COMPTEUR';          // Référence en rapport avec les données |
    EtatChampLibelle := 'PYR_LIBELLE';
    EtatChampBackGroundColor := 'PYR_PGCOLORPRE';

    { PopupMenu }
    // A créer dans l'ordre inverse de la présentation
    AddOptionPopup(GLOBAL,'Total global des compteurs');
    AddOptionPopup(SALARIE,'Total des compteurs par salarié');

    EnableOptionPopup(SALARIE,True);
    EnableOptionPopup(GLOBAL,True);

    { Alignement des colonnes fixes }
    TokenSizeColFixed   := '';
    TokenAlignColFixed  := '';
    For i := 1 To Rupture.NiveauRupt Do
    Begin
          TokenSizeColFixed   := TokenSizeColFixed  + '150;';
          TokenAlignColFixed  := TokenAlignColFixed + 'L;';
    End;
    TokenSizeColFixed   := TokenSizeColFixed  + '65;170;135';  // 3 colonnes : N°, Nom, Prénom salarié
    TokenAlignColFixed  := TokenAlignColFixed + 'C;L;L';

    { Liste des champs d'entete }
    TokenFieldColEntete := '';
    TokenFieldColFixed  := ListeChampRupt(';',Rupture) + 'LIB_PYP_SALARIE;NOM;PYR_LIBELLE';

    FormeGraphique := pgRectangle;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/07/2007
Modifié le ... :   /  /    
Description .. : Affichage du récapitulatif comprenant la légende utilisée
Mots clefs ... : 
*****************************************************************}
procedure TPlanningPresenceCompteurs.AfficheLegende;
Var
  i : Integer;
  LabelTemp, ColorTemp : THLabel;
  StLib : String;
  vLeft, vTop, vWidth, vHeight : Integer;
begin
     If (TobEtats <> Nil) And (TobEtats.Detail <> Nil) Then
     Begin
          For i := 0 To TobEtats.Detail.Count - 1 Do
          Begin
               //PT1
               StLib := TobEtats.Detail[i].GetValue('PYR_LIBELLE');
               if not Legende.Exist(StLib) then
               begin
                 Legende.Add(LabelTemp, ColorTemp, StLib);
                 Legende.GetLTWH(vLeft, vTop, vWidth, vHeight, Length(Legende.ATitre) -1, True);
                 CreateDetailsLabel(LabelTemp, Legende.GBRecap, vLeft,  vTop, vWidth, vHeight, '');
                 Legende.GetLTWH(vLeft, vTop, vWidth, vHeight, Length(Legende.ATitre) -1, False);
                 CreateDetailsLabel(ColorTemp, Legende.GBRecap, vLeft,  vTop, vWidth, vHeight, '');
              //   Legende.Add(LabelTemp, ColorTemp, StLib);
                 if (PRECAP.Height < vTop + vHeight + 10) and (not(PRECAP.Height > 250 )) then PRECAP.Height := vTop + vHeight + 10;
               end;

//PT1
{               Case i Of
               0 : Begin LabelTemp := LblL1; ColorTemp := LblC1; End;
               1 : Begin LabelTemp := LblL2; ColorTemp := LblC2; End;
               2 : Begin LabelTemp := LblL3; ColorTemp := LblC3; End;
               3 : Begin LabelTemp := LblL4; ColorTemp := LblC4; End;
               4 : Begin LabelTemp := LblL5; ColorTemp := LblC5; End;
               5 : Begin LabelTemp := LblL6; ColorTemp := LblC6; End;
               6 : Begin LabelTemp := LblL7; ColorTemp := LblC7; End;
               7 : Begin LabelTemp := LblL8; ColorTemp := LblC8; End;
               8 : Begin LabelTemp := LblL9; ColorTemp := LblC9; End;
               9 : Begin LabelTemp := LblL10; ColorTemp := LblC10; End;
               Else Begin ColorTemp := Nil; LabelTemp := Nil; End;
               End;
}
               // Couleur
               If ColorTemp <> Nil Then
               Begin
                    ColorTemp.ParentColor := False;
                    ColorTemp.Transparent := False;
                    ColorTemp.AutoSize    := True;
                    ColorTemp.Caption     := 'xxxxxxx';
                    ColorTemp.Font.Color  := StringToColor(TobEtats.Detail[i].GetValue('PYR_PGCOLORPRE'));
                    ColorTemp.Color       := StringToColor(TobEtats.Detail[i].GetValue('PYR_PGCOLORPRE'));
               End;

               // Libellé
               If LabelTemp <> Nil Then
               Begin
                    LabelTemp.Caption     := TobEtats.Detail[i].GetValue('PYR_LIBELLE');
               End;
          End;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : Gestion du clic-droit
Mots clefs ... : 
*****************************************************************}
procedure TPlanningPresenceCompteurs.GereOptionPopup(Item: TOB; ZCode: Integer; var ReDraw: Boolean; IdRessource : String; CurrentDate : TDateTime; ZoneSel : TZoneSelected );
Var NumSalarie : String;
begin
    If (IdRessource <> '') and (CurrentDate <> iDate1900) Then
    Begin
     Case ZCode Of
          SALARIE :  { Compteurs du salarié }
            Begin
               NumSalarie := Copy(IdRessource,1,SALARIE_KEY_LENGTH);
               AGLLanceFiche('PAY', 'TOTALCOMPTEURPRES','','','ACTION=CONSULTATION;'+NumSalarie+';'+DateToStr(DebutPlanning)+';'+DateToStr(FinPlanning)+';'+StWhTobEtats)
            End;

          GLOBAL :  { Compteurs globaux }
            Begin
               AGLLanceFiche('PAY', 'TOTALCOMPTEURPRES','','','ACTION=CONSULTATION;'+';'+DateToStr(DebutPlanning)+';'+DateToStr(FinPlanning)+';'+StWhTobEtats+';'+StWhTobRessource)
            End;
     End;
   End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : Conversion des préfixes de colonnes
Mots clefs ... : 
*****************************************************************}
function ConvertPrefixe(StWhere, DePref, APref: string): string;
var
  pospref: integer;
begin
  if StWhere <> '' then
    while Pos(DePref, StWhere) > 0 do
    begin
      pospref := Pos(DePref, StWhere);
      StWhere[(pospref + 1)] := APref[2];
      StWhere[(pospref + 2)] := APref[3];
    end;
  result := Trim(StWhere);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 10/07/2007
Modifié le ... :   /  /    
Description .. : Recherche de la tablette associée à un champ
Mots clefs ... : 
*****************************************************************}
Function RechTablette (NomChamp : String) : String;
Var
  Champ,Valeur : String;
  Q : TQuery;
  i : Integer;
Begin
  Valeur := '';
  If (NomChamp <> '') And (Length(NomChamp) > 5) Then
  Begin
    Champ := Copy(NomChamp,5,255);

    // Afin de limiter les accès en base, on sauvegarde les résultats précédents de recherche de tablette
    If (Length(Tablettes) > 0) Then
    Begin
          For i:=0 To Length(Tablettes)-1 Do
          Begin
                If (Tablettes[i][0] = Champ) Then Begin Valeur := Tablettes[i][1]; Break; End;
          End;
    End;

    // Recherche de la tablette concernée
    If (Valeur = '') Then
    Begin
          Try
               Q := OpenSql('SELECT DO_COMBO FROM DECOMBOS WHERE DO_NOMCHAMP like "%' + Champ + '%" ', True);
               Valeur := Q.FindField('DO_COMBO').AsString;
          Finally
               Ferme(Q);
          End;
          // Sauvegarde de la valeur lue
          SetLength(Tablettes, Length(Tablettes)+1);
          Tablettes[Length(Tablettes)-1][0] := Champ;
          Tablettes[Length(Tablettes)-1][1] := Valeur;
    End;
  End;
  Result := Valeur;
End;


{ TLegende }

//PT1
Function TLegende.Add(Lib, Couleur : THLabel; Titre : String) : Integer;
begin
  SetLength(ALibelle, Length(ALibelle) +1);
  SetLength(ACouleur, Length(ACouleur) +1);
  SetLength(ATitre, Length(ATitre) +1);
  ALibelle[Length(ALibelle) -1] := Lib;
  ACouleur[Length(ACouleur) -1] := Couleur;
  ATitre[Length(ATitre) -1] := Titre;
  result := Length(ALibelle) -1;
end;

//PT1
constructor TLegende.Create(NbParLigne: Integer);
begin
  NbrParLigne := NbParLigne;
end;

destructor TLegende.Destroy;
begin
//  if Assigned(GBRecap) then FreeAndNil();  GBRecap:  THGroupBox;
//  if Assigned() then FreeAndNil();  ALibelle: Array of THLabel;
//  if Assigned() then FreeAndNil();  ACouleur: Array of THLabel;
//  if Assigned() then FreeAndNil();  ATitre  : Array of String;

  inherited;
end;

function TLegende.Exist(Titre: String): Boolean;
var
  Index : Integer;
begin
  result := False;
  for Index := 0 to Length(ATitre) -1 do
  begin
    if ATitre[Index] = Titre then
    begin
      result := True;
      exit;
    end;
  end;     
end;

//PT1
procedure TLegende.GetLTWH(var vLeft, vTop, vWidth, vHeight: Integer;
  Index: Integer; Labl : Boolean);
begin
  vLeft   :=  10 + (Index mod NbrParLigne )* 225;
  vTop    :=  15 + 20 * Trunc(Index / NbrParLigne);
  vWidth  :=  35;
  vHeight :=  20;
  if Labl then
  begin
    vLeft :=  vLeft + 42;
    vTop  :=  vTop;//  + 6;
    vWidth  :=  165;
    vHeight :=  13;
  end;
end;

end.

