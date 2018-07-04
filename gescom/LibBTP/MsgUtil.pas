unit MsgUtil;

interface

uses Classes,
		 HEnt1,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
		 sysutils,
     Dicobtp,
     Ent1,
     Forms,
     Dialogs,
     Graphics,
     UTOM,
     UTOF,
     UTob,
     HCtrls,
     HTB97,
     HMsgBox;

     //Zone servant pour l'affichage des Grilles
Type
     AffGrille = RECORD
        ColGAppel  : String;
        LibGAppel	 : String;
        TableGapp	 : String;
        TriGapp		 : String;
        FormatGapp : String;
        LarGappel	 : String;
        AlignAppel : String;
        Alignement : String;
        SepMillier : String;
        NBDecimal  : String;
        Obligatoire: String;
        LibComplet : String;
        Visible    : String;
        BlancSiNul : String;
        Cumul      : String;
        Saisissable: String;
        TitreAppel : String;
        OKImage    : String;
        TabloMEF   : Array of Array of String;
     End;

     Procedure AfficheCouleur(Zone, Champs: THEdit);
     Procedure AfficheErreur(FormName, Erreur, Titre : String);
		 Procedure MajFontStyle (var Tobetats: Tob) ;
		 Procedure SelColorNew(Zone, Champs: THEdit; F : TForm);
     Procedure SelFonteNew(Zone : THLabel; Champs: THEdit; F : TForm);
     Procedure ChargeListeAssociee(StListe : String; StTitre : String; TTool : TToolWindow97; var GestGrille : AffGrille);
     Procedure DessineGrille(TTool : TToolWindow97; Grille : THGrid; GestGrille : AffGrille);
     Procedure RechRessourceAssociee(Var StUser, StUserName : String);

     Function DateToFloat (TDate: TDateTime) : Double;
     Function HRStrToInt(const StStr: string) : integer;
     Function LectLibCol(Prefixe, TypeLib, CodeLib, ZoneLib : String) : String;
     function MajItem(Liste: string; var GestGrille : AffGrille; ZoneARecup : String): string;
     Function RecTypeCommun(TypTablette : String; Libelle : String) : String;
		 Function StrDelete (chaine,sousChaine: string): string;

implementation

uses Grids;

const

	// libellés des messages pour la fiche Famille Ressource
  RecalcPMAP: array[1..5] of string  = (
  	  {1}  'Aucune donnée à traiter !'
      {2}, 'La Mise à jour de la Table est impossible !'
      {3}, 'Le code dépôt est obligatoire !'
      {4}, 'Attention le Code Article n''est pas renseigné ce traitement peut prendre énormément de temps !'
      {5}, 'Attention le Code Article est inexistant !'
           );


	// libellés des messages pour la fiche Famille Ressource
  FamResMsg: array[1..8] of string  = (
      {1}  'Rythme de gestion non renseigné !'
      {2}, 'Le nombre de qualités est limité à 10'
      {3}, 'Cette famille est utilisé par un type de ressource !'
      {4}, 'Nouvelle création impossible, Gestion de 3 familles de ressources au maximum !'
      {5}, 'La période de début est incorrecte.'
      {6}, 'La période de fin est incorrecte.'
      {7}, 'La période de fin est supérieure à la période de début.'
		  {8}, 'Le code famille est obligatoire.'
           );

  // Libellés des messages pour la Fiche Paramètre Planning
  ParPlann : array[1..19] of string = (
      {1}  'La famille de la ressource n''est pas renseignée!'
      {2}, 'Le mode de gestion n''est pas renseigné!'
      {3}, 'Le cadencement du planning n''est pas renseigné!'
      {4}, 'Choisissez un format de date.'
      {5}, 'Le libellé de la première colonne de tri n''est pas renseigné !'
      {6}, 'Le libellé de la deuxième colonne de tri n''est pas renseigné !'
      {7}, 'Le libellé de la troisième colonne de tri n''est pas renseigné !'
      {8}, 'La ressource doit être obligatoirement une colonne du planning !'
      {9}, 'L''heure de début est obligatoire dans le cas d''un planning horaire'
     {10}, 'L''heure de fin est obligatoire dans le cas d''un planning horaire'
     {11}, 'Lex axes de contingentements doivent être différents'
     {12}, 'Il faut au moins une colonne de Tri'
     {13}, 'Il faut au moins une colonne de Fixe'
     {14}, 'Le libellé de la première colonne Fixe n''est pas renseigné !'
     {15}, 'Le libellé de la deuxième colonne Fixe n''est pas renseigné !'
     {16}, 'Le libellé de la troisième colonne Fixe tri n''est pas renseigné !'
     {17}, 'Le libellé de la quatrième colonne Fixe n''est pas renseigné !'
     {18}, 'Le libellé de la cinquième colonne Fixe n''est pas renseigné !'
     {19}, 'Le libellé de la sixième colonne Fixe n''est pas renseigné !'

           );

	// Libellés des messages de l'écran des Types d'Action
	BTEtat	 : array[1..7] of string 	= (
      {1}  ''
      {2}, 'Le mode de gestion n''est pas renseigné !'
      {3}, ''
      {4}, 'Impossible de supprimer, type d''action présent dans un évènement !'
      {5}, ''
      {6}, ''
      {7}, ''
           );

  // Libellés des messages de la fiche evénementiels
  BTEvenement: array[1..7] of string = (
      {1}  'La désignation de l''état doit être renseignée'
      {2}, 'La date du fin de l''événement est antérieure à la date de début! '
      {3}, ''
      {4}, ''
      {5}, ''
      {6}, ''
      {7}, ''
    			 );

  MAjPlan : array[1..2] of string = (
      {1}  'La fonction %1 ne dispose d''aucune ressource'
      {2}, ''
           );
  // Libellés des messages du Planning
  BTPlanning: array[1..16] of string = (
      {1}  'Affectation Impossible, numéro d''évènement non déterminé !'
      {2}, 'Ressource déjà présente dans la grille des affectations !'
      {3}, 'La Date d''affectation est inférieure à la date d''appel !'
      {4}, 'Une affectation existe déjà pour cette ressource à cette date et à cette heure'
      {5}, 'le type d''action est obligatoire en saisie d''un évènement'
      {6}, 'Le surbooking n''est pas autorisé sur ce type d''action'
      {7}, 'La date de début et la date de fin sont égales !'
      {8}, 'La date de début est supérieure à la date de fin !'
      {9}, 'Modification impossible sur cet évènement !'
			{10},'Le type d''action n''existe pas dans la table Type Action !'
 			{11},'La création automatique à partir de l''appel ne c''est pas faite !'
			{12},'le type d''action par défaut n''est pas renseigné au niveau des paramètres'
 			{13},'le type d''action renseigné dans les paramètres n''existe pas'
      {14},'Impossible de modifier un appel réalisé ou terminé !'
      {15},'La tache associée à cet appel ne dispose d''aucun type d''action'
      {16},'La mise à jour de l''Appel a échouée. veuillez vérifier vos données'
    			 );

  // Libellés des messages du Planning
  ParamFeuilleVente: array[1..4] of string = (
      {1}  'Impossible d''enlever ce champ il est définit dans les zones de tri !'
      {2}, 'Impossible d''enlever ce champ il est obligatoire pour l''affichage !'
      {3}, 'Impossible de placer ce champ dans les zones de tri c''est une formule !'
      {4}, 'Vous devez renseigner une formule !'
    			 );

  //Libellés des Messages du module mobilité (sic)
  MobileMsg : Array [1..5] of String = (
      {1}   'Appel Inexistant',
      {2}   'Tiers Inexistant',
      {3}   'Ce client est fermé',
      {4}   'Contrat inexistant',
      {5}   'Contact inexistant'
          );

Procedure AfficheErreur(FormName, Erreur, Titre : String);
var NumError	: Integer;
    Msg     	: String;
Begin

  NumError := StrToInt(Erreur);
  Titre := TraduireMemoire(Titre);

  if NumError <> 0 then
     Begin
     If FormName = 'BTRECALCULPMAP' then
          Msg := TraduireMemoire(RecalcPMAP[NumError])
     Else If FormName = 'BTPARAMFEUILLEVENTE' then
          Msg := TraduireMemoire(ParamFeuilleVente[NumError])
     Else if FormName = 'BTFAMRES' then
    	  Msg := TraduireMemoire(FamResMsg[NumError])
     Else if FormName = 'BTPARAMPLAN' then
          Msg := TraduireMemoire(ParPlann[NumError])
     Else if FormName = 'BTPLANNINGEX' then
    	  Msg := TraduireMemoire(MajPlan[NumError])
     Else if FormName = 'BTETAT' then
    	  Msg := TraduireMemoire(BTEtat[NumError])
     Else if FormName = 'BTEVENEMENT' then
          Msg := TraduireMemoire(BTEvenement[NumError])
     Else if FormName = 'BTPLANNING' then
        Msg := TraduireMemoire(BTPlanning[NumError])
    Else if FormName = 'BTMESINTERV' then
        Msg := TraduireMemoire(MobileMsg[NumError]);
     end
  Else
	   Msg := TraduireMemoire(Erreur);

  if Msg <> '' then PGIBox(Msg, Titre);

end;

Procedure SelColorNew(Zone, Champs: THEdit; F : TForm);
Var T				: TColorDialog;
    Couleur	: Variant;
Begin

  Couleur := Champs.Text;

  T := TColorDialog.Create(F);

  T.color := Zone.Color;
  T.CustomColors.Clear;
  T.CustomColors.Add('ColorA=' + InttoHex(Zone.Color, 6));

  if T.Execute then
	   begin
     Couleur := IntTostr(T.Color);
     T.Free;
     end;

  if Couleur = '' then Couleur := StringToColor('ClWhite');

  Champs.Text := Couleur;
  Zone.Color := StringToColor(Couleur);
  Zone.Visible := True;

end;

procedure AfficheCouleur(Zone, Champs: THEdit);
var Couleur	: string;
begin

  Couleur := Champs.Text ;

  if (zone <> nil) and (champs.text <> '') then
	   begin
     Zone.Color := StringToColor(Champs.text);
     Zone.Visible := True;
     end
  else
     if (Zone <> nil) and (Champs.Text = '') then
        Zone.Visible := False;

end;

Procedure SelFonteNew(Zone : THLabel; Champs: THEdit; F : TForm);
Var T			: TFontDialog;
    Fonte	: Variant;
Begin

  Fonte := Champs.text ;

  T := TFontDialog.Create(F);

  T.Font.Size  := Zone.Font.Size;
	T.Font.Style := Zone.Font.Style;
  T.Font.Color := Zone.Font.Color;
  T.Font.Name  := Fonte;

  if T.Execute then Fonte := T.Font.Name;

  if Fonte = '' then Fonte := 'MS Sans Serif';

  Champs.Text := Fonte;

  Zone.Font.name := Fonte;
  Zone.Font.Size := T.Font.Size;
  Zone.Font.Style := T.Font.Style;
  Zone.Font.Color := T.Font.Color;

  T.Free;

end;

procedure MajFontStyle (var Tobetats: Tob) ;
var Boucle: integer;
    Etat	: string;
begin

  if (Tobetats <> nil) then
  	 begin
     for boucle := 0 to Tobetats.Detail.count - 1 do
    		 begin
	       Etat := TobEtats.detail[boucle].getvalue('BTA_FONTESTYLE') ;
         if Etat = 'fsBold' then
            TobEtats.detail[boucle].Putvalue('BTA_FONTESTYLE', 'G')
         else if Etat = 'fsItalic' then
            TobEtats.detail[boucle].Putvalue('BTA_FONTESTYLE', 'I') ;
         end;
     end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Régis H
Créé le ...... : 18/04/2005
Modifié le ... :   /  /
Description .. : fonction qui sécurise le STRTOINT
Mots clefs ... : STRTOINT
*****************************************************************}

function HRStrToInt (const StStr: string) : integer;
begin

  if (IsNumeric (StStr) ) then
    result := strtoint (StStr)
  else
    result := 0;

end;

Function LectLibCol(Prefixe, TypeLib, CodeLib, ZoneLib : String) : String;
Var Req	  : String;
		QQ	  : TQuery;
    Table : String;
Begin

  Result := '';

  if Prefixe = 'CC' then
     Table := 'CHOIXCOD'
  else if Prefixe = 'CO' then
     Table := 'COMMUN'
  else
     Table := 'CHOIXEXT';


  Req := 'SELECT ' + ZoneLib + ' FROM ' + Table;
  Req := Req + ' WHERE ' + Prefixe + '_TYPE="' + TypeLib ;
  Req := Req + '" AND ' + Prefixe +'_CODE="' + CodeLib + '"';
  QQ  := OpenSQL(Req, true,-1, '', True);

  if ZoneLib = '' then ZoneLib := Prefixe + '_LIBRE';

  if not QQ.EOF then Result := QQ.FindField(ZoneLib).asString;

  Ferme(QQ);

end;


//Suppression d'un caractère dans une chaine String
function StrDelete (chaine,sousChaine: string): string;
var indice : integer;
begin

  Repeat
  indice := pos (souschaine,chaine);
  if Indice = 1 then
     chaine := copy (chaine,2,255)
  else if Indice > 1 then
     Chaine := copy (chaine,1,Indice -1)+ copy(chaine,Indice+1,255);
  until indice <= 0;

  Result := Chaine;

end;

Function RecTypeCommun(TypTablette : String; Libelle : String) : String;
Var Req	: String;
		QQ	: TQuery;
Begin

  Result := '';

  Req := 'SELECT CO_CODE FROM COMMUN ';
  Req := Req + 'WHERE CO_TYPE="' + TypTablette + '" AND CO_LIBELLE="' + Libelle + '"';
  QQ  := OpenSQL(Req, true,-1, '', True);

  if not QQ.EOF then Result := QQ.FindField('CO_CODE').asString;

  Ferme(QQ);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Régis HARANG
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : fonction qui retourne la partie entiere de la date, cad sans
Suite ........ : l'heure
Mots clefs ... :
*****************************************************************}
function DateToFloat (TDate: TDateTime) : Double;
var
  tmp: double;
begin
  tmp := Tdate - 0.49999;
  result := arrondi (Tmp, 0) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 27/11/2007
Modifié le ... :   /  /
Description .. : Fonction qui permet de découper le mémo des liste
Suite ........ : pour récupérer les zone, la taille, le formatage...
Mots clefs ... :
*****************************************************************}
Procedure ChargeListeAssociee(StListe : String; StTitre : String; TTool : TToolWindow97; var GestGrille : AffGrille);
Var StTmp	      : String;
Begin

  StTmp := MajItem(StListe, GestGrille, 'ZZ');

  if not assigned(TTool) then exit;

  TTool.caption := StTitre;
  TTool.Hint := StTitre;

end;

function MajItem(Liste: string; var GestGrille : AffGrille; ZoneARecup : String): string;
var
  Table: string;
  Crit: string; //paramètre de jointure
  Tri: string;
  Champ: string;
  Titre: Hstring;
  Larg: string;
  Align: string;
  Params: string;
  LeTitre: Hstring;
  NomCol: Hstring;
  Perso: string;
  OkTri: Boolean;
  OkNumCol: Boolean;
begin

  ChargeHListe(Liste, Table, Crit, tri, Champ, Titre, Larg, Align, Params, LeTitre, NomCol, Perso, OkTri, OkNumCol);

  if ZoneARecup = 'C' then
     Result := Champ
  Else if ZoneARecup = 'AL' then
     Result := Align
  Else if ZoneARecup = 'ZZ' then
     Begin
     //Chargement des variables du Record AffGrille
     GestGrille.ColGAppel  := Champ;
     GestGrille.LibGAppel  := Titre;
     GestGrille.TableGapp  := Table;
     GestGrille.TriGapp	   := tri;
     GestGrille.FormatGapp := '';
     GestGrille.LarGappel  := Larg;
     GestGrille.AlignAppel := Align;
     GestGrille.Alignement := '';
     GestGrille.SepMillier := '';
     GestGrille.NBDecimal  := '';
     GestGrille.Obligatoire:= '';
     GestGrille.LibComplet := '';
     GestGrille.Visible    := '';
     GestGrille.BlancSiNul := '';
     GestGrille.Cumul      := '';
     GestGrille.Saisissable:= '';
     GestGrille.OKImage    := '';
     // Erreur ne marche que pour une seule zone !!!!!!
     While Align <> '' do
        begin
        Champ := ReadTokenSt(Align);
        GestGrille.Alignement := GestGrille.Alignement  + Copy(Champ, 1, 1) + ';';
        GestGrille.SepMillier := GestGrille.SepMillier  + Copy(Champ, 2, 1) + ';';
        GestGrille.NBDecimal  := GestGrille.NBDecimal   + Copy(Champ, 3, 1) + ';';
        GestGrille.Obligatoire:= GestGrille.Obligatoire + Copy(Champ, 4, 1) + ';';
        GestGrille.LibComplet := GestGrille.LibComplet  + Copy(Champ, 5, 1) + ';';
        GestGrille.Visible    := GestGrille.Visible     + Copy(Champ, 6, 1) + ';';
        GestGrille.BlancSiNul := GestGrille.BlancSiNul  + Copy(Champ, 7, 1) + ';';
        GestGrille.OkImage    := GestGrille.OkImage     + Copy(Champ, 8, 1) + ';';
        GestGrille.Cumul      := GestGrille.Cumul       + Copy(Champ, 8, 1) + ';';
        GestGrille.Saisissable:= GestGrille.Saisissable + Copy(Champ, 9, 1) + ';';
        end;
     //
     GestGrille.TitreAppel := LeTitre;
     Result := Table + '-' +  Crit + '-' +  tri + '-' +  Champ + '-' +  Titre + '-' +  Larg + '-' + '-' + Params + '-' +  LeTitre;
     end;

  while Champ <> '' do
     begin
     Champ := ReadTokenst(Align);
     if Pos('$', Champ) <> 0 then
        GestGrille.FormatGapp := GestGrille.FormatGapp + '$;'
     else
        GestGrille.FormatGapp := GestGrille.FormatGapp + ';'
     end;

end;
{***********A.G.L.***********************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 27/11/2007
Modifié le ... :   /  /
Description .. : Fonction qui permet de dessiner une grille à
Suite ........ : partir d'une liste paramétrable...
Mots clefs ... :
*****************************************************************}
Procedure DessineGrille(TTool : TToolWindow97; Grille : THGrid; GestGrille : AffGrille);
var StTmp       : String;
    FF          : String;
    FFQTE       : String;
    FFLQTE      : string;
    Alignement  : String;
    NomZone     : String;
    //
    Nbcol       : Integer;
    I			      : integer;
    //
    TailleCol	  : Double;
    //
    OkLib       : Boolean;
    OkVisu      : Boolean;
    OkImg       : Boolean;
    OkNulle     : Boolean;
    OkCumul     : Boolean;
    Separateur  : Boolean;
    Obligatoire : Boolean;
    //
    Dec         : integer;
    //
Begin

	FFQTE := '###';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQTE := FFQTE+'0.';
    for I := 1 to V_PGI.OkDecQ do
    begin
      FFQTE := FFQTE + '0';
    end;
  end;

  nbcol := 0;

  //calcul du Nombre de Colonne (voir indicateur (???))
  StTmp := GestGrille.ColGAppel;
  while  StTmp <> '' do
     begin
     NbCol := NbCol + 1;
     ReadTokenst(StTmp);
     end;

  if nbcol = 0 then exit;

  SetLength(GestGrille.TabloMEF,4,NbCol);

  //Dessin de la grille
	With Grille do
       Begin
 		   RowCount := 2;
  		 ColCount := NbCol + 1;
			 FixedCols := 1;
			 FixedRows := 1;
			 DefaultRowHeight := 20;
    ColWidths[0] := 20;
       Cells[0,0] := '';

    if Assigned(TTool) then TTool.Width := 0;

       if ColCount <> 1 then
    begin
      if Assigned(TTool) then TTool.Width := 581;
 	     For i := 1 to ColCount - 1 Do
           Begin
        //chargement des entête de colonnes
	         Cells[I, 0] 	:= ReadTokenst(GestGrille.LibGAppel);

        NomZone := ReadTokenSt(GestGrille.COLGAppel);
        GestGrille.TabloMEF[0, I-1] := NomZone;

        //Chargement de la taille des colonne
           TailleCol   	:= StrToFloat(ReadTokenst(GestGrille.LarGAppel));

        //chargement des paramètres de présentation de la liste
        Alignement   	:= ReadTokenst(GestGrille.Alignement);
        GestGrille.TabloMEF[1, I-1] := Alignement;
        //
        Dec           := StrToInt(ReadTokenst(GestGrille.NBDecimal));
        //
        if ReadTokenst(GestGrille.SepMillier) = '' then
          Separateur  := False
        else
          Separateur  := True;

        if ReadTokenst(GestGrille.Obligatoire) = '' then
          Obligatoire := False
        else
          Obligatoire := True;
        //
        if ReadTokenst(GestGrille.LibComplet) = '' then
        begin
          OkLib := False;
          GestGrille.TabloMEF[2, I-1] := '-';
        end
        else
        begin
          OkLib := True;
          GestGrille.TabloMEF[2, I-1] := 'X';
        end;
        //
        if ReadTokenst(GestGrille.Visible) = 'X' then
          OkVisu := False
        else
          OkVisu := True;
        //
        if ReadTokenst(GestGrille.BlancSiNul) = '-' then
          OkNulle := False
        else
          OkNulle := True;
        //
        if ReadTokenst(GestGrille.Cumul) = '-' then
          OkCumul := False
        else
          OkCumul := True;

        //Test si la colonne est visible ou non
        if OkVisu then
          TailleCol   := TailleCol * Canvas.TextWidth('W')
        else
          TailleCol   := -1;
		       ColWidths[I] := StrToInt(FloatToStr(TailleCol));

        //Positionnement du texte dans la cellule
        if      copy(Alignement,1,1)='G' then ColAligns[I] := taLeftJustify
        else if copy(Alignement,1,1)='D' then ColAligns[I] := taRightJustify
        else if copy(Alignement,1,1)='C' then ColAligns[I] := taCenter;

        //positionnement représentation graphique de la zone dans la cellule
        if ReadTokenst(GestGrille.OKImage) = '-' then
        begin
          OkImg := False;
          GestGrille.TabloMEF[3,I-1] := '-';
        end
        else
        begin
          OkImg := True;
          GestGrille.TabloMEF[3,I-1] := 'X';
					 end;
        //

        //
        if (OkLib) or (okImg) then
        begin
          //ColFormats[I] := 'CB=' + Get_Join(NomZone);
          if OkImg then
          begin
            ColDrawingModes[I]:= 'IMAGE';
       end;
        end
        else if (Dec<>0) or (Separateur) then
        begin
          if OkNulle then
          begin
            FFLQte := FFQTE+';-'+FFQTE+'; ;';
            ColFormats[I] := FFLQTE ;
          end else
          begin
            ColFormats[I] := FF ;
          end;
        end;
        //redimensionnement du container si TToolWindow
        if Assigned(TTool) then
        begin
          if TTool.width < (Screen.Width-TailleCol) then TTool.Width := TTool.Width + StrToInt(FloatToStr(TailleCol));
        end;

      end;
    end;
  end;

  //Affichage et recentrage de la ttoolwindow TTWappel après chargement
  if Assigned(TTool) then
  begin
  TTool.Height := Screen.Height div 2;
  if TTool.width >=  Screen.Width then
     Begin
     Taillecol := TTool.Width;
     TTool.Width :=  round(Taillecol / 2);
     end;
  TTool.Left:=(Screen.Width-TTool.width) div 2;
  TTool.Top:=(Screen.Height-TTool.height) div 2;
  end;

end;

Procedure RechRessourceAssociee(Var StUser, StUserName : String);
Var StSQl : String;
    QQ    : TQuery;
begin

  StSQL := 'SELECT ARS_RESSOURCE, ARS_LIBELLE FROM RESSOURCE ';
  StSQL := StSQL + 'WHERE ARS_TYPERESSOURCE="SAL" AND ARS_UTILASSOCIE="' + StUser + '" AND ARS_FERME <> "X"';
  QQ := OpenSQL(StSQL, False);

  If not QQ.eof then
  begin
    StUser      := QQ.Findfield('ARS_RESSOURCE').AsString;
    StUserName  := QQ.Findfield('ARS_LIBELLE').AsString;
  end;


end;

end.
