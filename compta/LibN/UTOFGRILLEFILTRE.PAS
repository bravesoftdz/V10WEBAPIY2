{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 03/03/2003
Modifié le ... : 17/08/2004
Description .. : Source TOF de la FICHE : CPGRILLEFILTRE
Suite ........ : 
Suite ........ : Classe abstraite permettant la gestion des fiches de type
Suite ........ :  saisie en grille filtré avec une partie filtre
Suite ........ : Exemple :
Suite ........ :   CPCODEPOSTAL, ....
Suite ........ : 
Suite ........ : JP 20/08/03 : FormKeyDown est mise en public et virtual
Suite ........ :               InsererLigne est mise en public
Suite ........ : JP 17/08/03 : Séparation des NotifyEvents et des
Suite ........ : traitements qui s'y appliquent
Mots clefs ... : TOF;CPGRILLEFILTRE
*****************************************************************}
Unit UTOFGRILLEFILTRE ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
     Forms,
     SysUtils,
     ComCtrls,
{$IFDEF VER150}
     variants,
{$ENDIF}
     Messages,
{$IFDEF EAGLCLIENT}
     maineagl,
     UtileAGL,
{$ELSE}
     db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_main,
     EdtREtat,
{$ENDIF}
     HCtrls,
     HPanel,
     HEnt1,
     HMsgBox,
     Hqry,
     UTOB,
     UTOF ;

Type
  TOF_GRILLEFILTRE = Class (TOF)

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    // Paramètres
    FNomTable       : String ;    // FROM
    FChampTri       : String ;    // ORDER BY
    FListeChamps    : String ;    // COL de la grille
    FCodeEtat       : String ;    // Pour lanceEtat
    FNatureEtat     : String ;    // Pour lanceEtat
    FTitreEcran     : String ;    // Ecran.Caption
    FListeColsUniq  : String ;    // Accès grille
    // Gestion interne
    TOBLignes       : TOB;
    TOBCourante     : TOB;
    TOBOrig         : TOB ;
    TOBEnreg        : TOB ;
    StTypeAction    : String ;
    {$IFDEF EAGLCLIENT}
      MaxRow        : Integer ;
    {$ENDIF EAGLCLIENT}
    LigneModifiee   : Boolean ;
    // évènements
    procedure FListeKeyPress( Sender : TObject ; var Key : Char ) ;
    procedure FListeRowEnter(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean) ;
    procedure FListeRowExit(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean) ;
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean) ;
    procedure Imprimer(Sender : TObject) ;
    procedure ParcoursListe (Sender : TObject) ;
    procedure AffecterEvtFiltre ;
    procedure LancerRecherche(Sender : TObject) ;
{$IFDEF EAGLCLIENT}
    procedure BNextQClick(Sender: TObject);
    procedure BPrevQClick(Sender: TObject);
{$ENDIF}
    // Traitements
    function  ValiderLigne(row : integer; avecConfirm : Boolean = False): Boolean;
    // Traitements Graphique
    procedure BoutonsNavMaj ;
    procedure RaffraichirLigne( row : Integer ) ;
    // Utils
    function  GenererCondition: String ;
    function  EstRemplit( vInLigne : Integer ) : Boolean ;
    function  EstNewEnreg : Boolean ;
    function  ExistDB( vTOBEnreg, vTOBOrig : TOB )  : Boolean ;

    {JP 17/08/04 : Traitements}
    procedure InsererLigne(AvecFocus : Boolean = True);
    procedure AnnulerModif;
    procedure SupprimerLigne;
  public
    // utilisable dans la fille
    FListe          : THGrid ;
    procedure RefreshFListe ;
    procedure RemplitGrille ;
    // Accès aux données de la TOB
    function  GetField (vStChamps : String ; row : integer = -1) : Variant;
    procedure SetField (vStChamps : String ; vStValeur : Variant ; row : integer = -1);

    {JP 20/08/03 : Pour pouvoir la surcharger}
    procedure FormKeyDown (Sender : TObject; var Key: Word; Shift: TShiftState) ; virtual;
    procedure FListeOnExit(Sender : TObject); virtual;{JP 18/05/06 : FQ 17546}

    {JP 17/08/04 : Évènements proprement dits}
    procedure BInsertClick (Sender : TObject);
    procedure BDeleteClick (Sender : TObject);
    procedure BDefaireClick(Sender : TObject);
    {JP 18/08/04 : On s'assure que les critères de recherche sont valides : par défaut renvoie True}
    function IsCritereValide : Boolean; virtual;
    // Paramétrage dans la fille
    procedure OnLoadRecord ( vTOBLignes : TOB )                ; virtual ; abstract ;
    function  OnUpdateRecord ( laTOB : TOB )       : Integer   ; virtual ; abstract ;
    function  OnAfterUpdateRecord ( laTOB : TOB )  : Integer   ; virtual ; abstract ;
    function  OnDeleteRecord ( laTOB : TOB )       : Integer   ; virtual ; abstract ;
    procedure ParametrageFiche                                 ; virtual ; abstract ;
    procedure ParametrageGrille ( vGrille : THGrid )           ; virtual ; abstract ;
    procedure OnCreateTOB ( vTOB : TOB )                       ; virtual ; abstract ;
    function 	GenererConditionPlus                 : String    ; virtual ; abstract ;
    // Propriétés
    property    NomTable        : String          read FNomTable        write FNomTable ;
    property    ChampTri        : String          read FChampTri        write FChampTri ;
    property    ListeChamps     : String          read FListeChamps     write FListeChamps ;
    property    CodeEtat        : String          read FCodeEtat        write FCodeEtat ;
    property    NatureEtat      : String          read FNatureEtat      write FNatureEtat ;
    property    TitreEcran      : String          read FTitreEcran      write FTitreEcran ;
    property    ListeColsUniq   : String          read FListeColsUniq   write FListeColsUniq ;
  end ;

Implementation

procedure TOF_GRILLEFILTRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GRILLEFILTRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GRILLEFILTRE.OnUpdate ;
begin
  Inherited ;
	ValiderLigne(FListe.row);
end ;

procedure TOF_GRILLEFILTRE.OnLoad ;
begin
  Inherited ;
  RemplitGrille ;
  LigneModifiee := False ;
end ;

procedure TOF_GRILLEFILTRE.OnArgument (S : String ) ;
begin
  Inherited ;

  StTypeAction := 'ACTION=MODIFICATION' ;
  // Traitements des arguments
  if Trim( S ) <> '' then
    StTypeAction := ReadTokenSt ( S ) ;

  // Tob contenant les enregistrements
  TOBLignes := TOB.Create('Tob virtuelle', nil, -1) ;

  // Récupération des controles importants
	FListe  := THGrid(GetControl('Grille', True));

  // Réaffectation des évènements
  TButton(GetControl('BFirst',True)).onClick       := ParcoursListe;
  TButton(GetControl('BPrev',True)).onClick        := ParcoursListe;
  TButton(GetControl('BNext',True)).onClick        := ParcoursListe;
  TButton(GetControl('BLast',True)).onClick        := ParcoursListe;
  TButton(GetControl('BImprimer',True)).onClick    := Imprimer;
  TButton(GetControl('BInsert',True)).onClick      := BInsertClick;
  TButton(GetControl('BDelete',True)).onClick      := BDeleteClick;
  TButton(GetControl('bDefaire',True)).onClick     := BDefaireClick;
  FListe.OnRowEnter                                := FListeRowEnter;
  FListe.OnRowExit                                 := FListeRowExit;
  FListe.OnCellEnter                               := GrilleCellEnter ;
  FListe.OnKeyPress                                := FListeKeyPress ;
  FListe.OnExit                                    := FListeOnExit;{JP 18/05/06 : FQ 17546}
  Ecran.OnKeyDown                                  := FormKeyDown;
{$IFDEF EAGLCLIENT}
  TButton(GetControl('BPrevQ',True)).onClick       := BPrevQClick ;
  TButton(GetControl('BNextQ',True)).onClick       := BNextQClick ;
  SetControlVisible('BPrevQ', True) ;
  SetControlVisible('BNextQ', True) ;
{$ENDIF}

  // paramètrage dans la fille :
  ParametrageFiche ;
  ParametrageGrille ( FListe ) ;

  AffecterEvtFiltre ;

  // Titre de l'écran
  Ecran.Caption := TitreEcran ;
  UpdateCaption( Ecran ) ;

  // Bouton imprimer
  SetControlVisible ('BImprimer', ( Trim(NatureEtat)<>'' ) and ( Trim(CodeEtat)<>'' ) ) ;

  // Affectation des TOB utilisés pour l'enregistrement
  TOBOrig   := TOB.Create( nomTable , nil , -1 ) ;
  TOBEnreg  := TOB.Create( nomTable , nil , -1 ) ;

end ;

procedure TOF_GRILLEFILTRE.OnClose ;
begin
  Inherited ;
  // Validation des modifications si besoin
	if not ValiderLigne(FListe.Row,True) then
  	Begin
    LastError := 1;
    Exit;
    end;
  // Libération mémoire des TOBs
  if TOBLignes <> nil then
    TOBLignes.Free;
  if TOBEnreg <> nil then
    TOBEnreg.Free ;
  if TOBOrig <> nil then
    TOBOrig.Free ;
end ;

procedure TOF_GRILLEFILTRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_GRILLEFILTRE.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : Recharge les données de la TOB dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.AnnulerModif;
var lInligne : Integer;
begin
	// Test indice...
  lInligne := FListe.row;
  if (lInligne < 1) or (lInligne > TOBLignes.Detail.count) then Exit;
  raffraichirLigne( lInligne ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : MAJ affichage des boutons de navigation dans la grille
Mots clefs ... : 
*****************************************************************}
procedure TOF_GRILLEFILTRE.BoutonsNavMaj;
begin
	// MAJ des boutons de navigations
  SetControlEnabled('BFirst',FListe.Row > 1);
  SetControlEnabled('BPrev', FListe.Row > 1);
  SetControlEnabled('BNext', FListe.Row < TOBLignes.Detail.Count );
  SetControlEnabled('BLast', FListe.Row < TOBLignes.Detail.Count );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : Lance l'impression si paramètré
Mots clefs ... : 
*****************************************************************}
procedure TOF_GRILLEFILTRE.Imprimer(Sender: TObject);
begin
  if ( Trim(NatureEtat)='' ) or ( Trim(CodeEtat)='' ) then Exit ;
	LanceEtat(  'E',
              NatureEtat,
              CodeEtat,
              True,
              False,
              False,
              Nil,
              GenererCondition,
              Ecran.Caption ,
              False,
              0 );
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : Gestion du parcours de la grille via les boutons de 
Suite ........ : naviguation
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.ParcoursListe(Sender: TObject);
var
	ctrlName : String; // nom du controle
  oldRow, newRow : Integer;
begin
	ctrlName := TButton(Sender).name;
  oldRow := FListe.Row;
  newRow := FListe.Row;
  if ctrlName = 'BFIRST' 				// Première ligne
  	then	newRow := 1;
  if ctrlName = 'BPREV' then		// Ligne précédente
   	if FListe.Row > 1 then newRow := FListe.Row - 1;
  if ctrlName = 'BNEXT' then	  // Ligne suivante
    	if FListe.Row < TOBLignes.Detail.Count then	newRow := FListe.Row + 1;
  if ctrlName = 'BLAST'  				// Dernière ligne
  	then newRow := TOBLignes.Detail.Count;
	// Déplacement
  if newRow <> oldRow then FListe.gotoRow(newRow);
end;

procedure TOF_GRILLEFILTRE.RaffraichirLigne( row: Integer );
begin
  if TOBCourante = nil then Exit ;
  // Rechargement des données de la TOB dans la ligne de la Grid
  TOBCourante.PutLigneGrid(FListe, row, FALSE, FALSE, ListeChamps );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : Valide et enregistre les données de la ligne d'index 'row'
Suite ........ : si avecConfirm = true alors demande confirmation
Suite ........ : Retourne True si le traitement peut continuer, False sinon
Mots clefs ... :
*****************************************************************}
function TOF_GRILLEFILTRE.ValiderLigne(row: integer; avecConfirm: Boolean): Boolean;
var Rep      : Integer ;
begin
	Result := True;
  if TOBCourante = nil then Exit ;
  if not EstRemplit( row ) then Exit ;
  // données modifiées ?
  if LigneModifiee then
  	begin
    // TOBs pour maj base
    TOBOrig.Dupliquer( TOBCourante, False, True ) ;
    TOBEnreg.Dupliquer( TOBCourante, False, True ) ;
    // pour rendre utilisable les méthodes accès base
    TOBOrig.VirtuelleToReelle( NomTable ) ;
    TOBEnreg.VirtuelleToReelle( NomTable ) ;
    // reprise données grilles
    TOBEnreg.GetLigneGrid( FListe , row , ';' + ListeChamps ) ;
		// Y-a-t-il une demande de confirmation à faire ?
	  if avecConfirm
    	then Rep := PGIAskCancel('Voulez-vous enregistrer les modifications ?',Ecran.Caption)
      else Rep := mrYes;
    // Traitement :
    Case Rep of
			// --> Validation des modifications
      mrYes : Try
                // Vérification
                if OnUpdateRecord( TOBEnreg ) <> 0
                  then raise EAbort.Create('') ;
                // Test unicité pour insertion
                if ExistDB( TOBEnreg , TOBOrig ) then
                  begin
                  PGIBox('L''enregistrement que vous avez saisi existe déjà. Vous devez le modifier.', Ecran.Caption ) ;
                  raise EAbort.Create('') ;
                  end ;
                // Suppression ancien enreg
                if not EstNewEnreg then
                  TOBOrig.DeleteDB ;
                // Insertion
                TOBEnreg.InsertDB(nil) ;
                // MAJ TOBCourante
                TOBCourante.GetLigneGrid( FListe , row , ';' + ListeChamps ) ;
                // MAJ indicateur nouvel enregistrement
                if EstNewEnreg then
                  TOBCourante.PutValue( 'NEWENREG' , False ) ;
                // Annul modif car pb AGL
                TOBCourante.modifie := False ;
                LigneModifiee := False ;
                // Traitement post Enreg
                if OnAfterUpdateRecord( TOBCourante ) <> 0
                  then raise EAbort.Create('') ;
              Except
                on E:Exception do
                  begin
                  Result := False ;
                  if E.Message <> '' then
                    PGIBox( 'Erreur à l''enregistrement.' + #10#13 + E.Message , Ecran.Caption ) ;
                  end ;
              end ; // Fin mrYes
			// --> Annulation des modifications
      mrNo  :	begin
              // MAJ affichage dans la grid
              raffraichirLigne( row );
              LigneModifiee := False ;
              Result := True;
              end;  // Fin mrNo
			// --> Annulation de la validation
      mrCancel : Result := False ;
    	end ; // Fin Case
    end; // fin if
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : Quand on arrive dans une nouvelle ligne, MAJ de la
Suite ........ : TOBCourante et des bontons de naviguation
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.FListeRowEnter(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  // Mise à jour de la TOB courante
  if TOBLignes.Detail.Count >= ou then
    TOBCourante := TOBLignes.Detail[ou-1];
  LigneModifiee := False ;
	// MAJ des boutons de navigations
  BoutonsNavMaj;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : Quand on quitte un ligne, MAJ auto des modifs
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.FListeRowExit(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	// On quitte une ligne, validation des modifs...
  if Not EstRemplit(ou) then Exit ;
	Cancel := not ValiderLigne(ou) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... : 11/03/2003
Description .. : Gestion des touches de raccourcis
Suite ........ :
Suite ........ :   Ctrl-I             :  Impression
Suite ........ :   Ctrl-N ou Inser    :  Insertion
Suite ........ :   Ctrl-suppr         :  Suppression
Suite ........ :   F10                :  validation
Suite ........ :   Entrée dans la Grille : comme Tabulation
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var	OkG, Vide, AvecCtrl : boolean;
begin
	if Not FListe.SynEnabled then
  	begin
    Key := 0;
    Exit;
    end;

  OkG      := ( Screen.ActiveControl = FListe ) ;
  Vide     := ( Shift = [] ) ;
  AvecCtrl := ( Shift = [ssctrl] ) ;

  Case Key of
    // Validation cellule = tabulation
    VK_RETURN : if ((OkG) and (Vide)) then KEY := VK_TAB ;
    // Validation ligne
    VK_F10    : if Vide then
                  begin
                  Key := 0 ;
                 	ValiderLigne(FListe.row);
                  end ;
    // suppresion
    VK_DELETE : if ((OkG) and AvecCtrl) then
                  begin
                  Key := 0 ;
                  SupprimerLigne;{JP 17/08/04: maintenant il s'agit d'une méthode et non plus d'une évènement}
                  end ;
    // nouvelle ligne
    VK_INSERT : if Vide then
                  begin
                  Key := 0 ;
                  InsererLigne;{JP 17/08/04: maintenant il s'agit d'une méthode et non plus d'une évènement}
                  end ;
      {^N} 78 : if AvecCtrl then
                  begin
                  Key := 0 ;
                  InsererLigne;{JP 17/08/04: maintenant il s'agit d'une méthode et non plus d'une évènement}
                  end ;
    // Impression
      {^I} 73 : if AvecCtrl then
                  begin
                  Key := 0 ;
                  Imprimer(nil) ;
                  end ;
  end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : Construit la condition de la requête en parcourant les
Suite ........ : composant de la fiche avec nom de champ.
Suite ........ : Types de composants gérés :
Suite ........ :   THEdit, THValComboBox, TCheckBox
Suite ........ :
Suite ........ : A PREVOIR : utilisation des paramètres "type de données"
Suite ........ : et "type de comparaison" des composant +
Suite ........ : THMULTIVALCOMBOBOX
Mots clefs ... :
*****************************************************************}
function TOF_GRILLEFILTRE.GenererCondition: String;
var lHEdit  : THEdit ;
    lHCombo : THValComboBox ;
    lChkBox : TCheckBox ;
    lInCpt  : Integer ;
    lStPlus : String ;
    TOBTest : TOB ;
begin
  Result := '' ;
  TOBTest := TOB.Create(nomTable,nil,-1);
  // Parcours des composants du PFiltre
  for lInCpt := 0 to Ecran.ComponentCount - 1 do
    begin
    // Cas du THEdit
    if Ecran.Components[lInCpt] is THEdit then
      begin
      lHEdit := THEdit( Ecran.Components[lInCpt] ) ;
      if Trim(lHEdit.Text)<>'' then
        begin
        if lHEdit.Name = 'XX_WHERE' then
          Result := Result + ' AND ' + lHEdit.Text ;
        if TOBTest.FieldExists(lHEdit.Name) then
          Result := Result + ' AND ' + lHEdit.Name + ' LIKE "' + LHEdit.Text + '%"' ;
        end ;
      end ;
    // Cas du THValComboBox
    if Ecran.Components[lInCpt] is THValComboBox then
      begin
      lHCombo := THValComboBox( Ecran.Components[lInCpt] ) ;
      if TOBTest.FieldExists(lHCombo.Name) then
        if Trim(lHCombo.Value)<>'' then
          Result := Result + ' AND ' + lHCombo.Name + '="' + lHCombo.Value + '"' ;
      end ;
    // Cas du TCheckBox
    if Ecran.Components[lInCpt] is TCheckBox then
      begin
      lChkBox := TCheckBox( Ecran.Components[lInCpt] ) ;
      if TOBTest.FieldExists(lChkBox.Name) then
        if lChkBox.checked
          then Result := Result + ' AND ' + lChkBox.Name + '="X"'
          else Result := Result + ' AND ' + lChkBox.Name + '="-"' ;
      end ;
    end ; // fin for

  // Ajout de la condition utilisateur
  lStPlus := GenererConditionPlus ;
  if Trim(lStPlus)<>'' then
    Result := Result + ' AND ( ' + lStPlus + ' ) ' ;

  // Remplacement du premier AND par un WHERE
  if Result <> '' then
    Result := FindEtReplace(Result, 'AND' , '', False );

  TOBTest.Free ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : Insère une ligne dans la Grille avec une nouvelle tob 
Suite ........ : associée sur la table "nomTable".
Suite ........ : 
Suite ........ : Il est posible d'initialiser les champs de la tob dans la otf 
Suite ........ : fille par l'intermédiaire de la méthode abstraite OnCreateTOB
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.InsererLigne(AvecFocus : Boolean = True);
var newLigne : TOB ;
begin
  // Ajout d'une ligne dans la Grille
  if TOBLignes.Detail.Count >= (FListe.RowCount - 1) then
    if EstRemplit ( FListe.RowCount - 1 )
      then FListe.RowCount := FListe.RowCount + 1
      else Exit ;
  // Nouvelle tob
  newLigne := TOB.Create( NomTable, TOBLignes, -1 ) ;
  newLigne.InitValeurs ;
  // Gestion nouvel enregistrement
  newLigne.AddChampSup('NEWENREG', False) ;
  newLigne.PutValue('NEWENREG', True) ;
  // Evt fille
  OnCreateTOB ( newLigne ) ;
  // Positionnement + evt
  FListe.gotoRow( FListe.RowCount - 1 ) ;
  TOBCourante := newLigne ;
  // MAJ ligne de la grille au cas ou
  RaffraichirLigne( FListe.RowCount - 1 );
  {JP 17/08/04 : On ne donne le Focus que si nécessaire}
  if AvecFocus and Fliste.CanFocus then begin
    // Positionnement
    FListe.col := 1 ;
    FListe.SetFocus ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : Supprime l'enregistrement courant en base et dans la Grille 
Suite ........ : après demande de confirmation
Mots clefs ... : 
*****************************************************************}
procedure TOF_GRILLEFILTRE.SupprimerLigne;
var TobSuppr : TOB ;
begin
  if TOBCourante = nil then Exit ;

  // Message avertissement
  if PGIAskCancel('Voulez-vous supprimez l''enregistrement courant ?', Ecran.Caption) <> mrYes
    then Exit ;

  if OnDeleteRecord( TOBCourante ) <> 0 then Exit;

  // Suppression de TOBLignes
  TOBSuppr := TOB.Create( 'Suppression TOB', nil , -1 ) ;
  TOBCourante.ChangeParent(TOBSuppr, -1) ;

  // Suppression Base si pas un nouvel enregistrement
  if not EstNewEnreg then
    Try
      if not TOBCourante.DeleteDB
        then raise EAbort.Create('');
      Except
        // Error, on rollback
        on E:Exception do
          begin
          TOBCourante.ChangeParent( TOBLignes , FListe.row ) ;
          if E.Message <> '' then
            PGIBox( 'Erreur à la suppression de l''enregistrement.' + #10#13 + E.Message , Ecran.Caption ) ;
          end ;
      end ;

  // Suppression Mémoire
  TOBSuppr.Free ;

  // MAJ Grille
  RefreshFListe ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... : 11/03/2003
Description .. : Raffraichit la Grille avec les données de TOBLignes.
Suite ........ : Si aucun enregistrement, place la première ligne en mode 
Suite ........ : insertion
Mots clefs ... : 
*****************************************************************}
procedure TOF_GRILLEFILTRE.RefreshFListe ;
var lInCol, lInRow : Integer ;
    lBoParam       : Boolean ;
begin
  // sauvegarde Position
  lInCol := FListe.Col ;
  lInRow := FListe.Row ;

  // MAJ Affichage Grille
  FListe.rowCount := 2 ;
  TOBLignes.PutGridDetail( FListe, False, False, ListeChamps, False ) ;

  // Restauration Position
  FListe.Col := lInCol ;
  if lInRow >= FListe.RowCount
    then lInRow := FListe.RowCount - 1 ;
  FListe.Row := lInRow ;
  lBoParam := False;
  FListeRowEnter(nil, lInRow, lBoParam, False);

  // Cas spécial : Aucun enregistrement -> on se positionne en mode insert !
  if TOBLignes.Detail.Count = 0 then {JP 17/08/04: maintenant il s'agit d'une méthode et non plus d'une évènement}
    begin
    InsererLigne(False) ;
    end ;        
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... : 11/03/2003
Description .. : Retourne vrai si au moins une cellule de la ligne est non 
Suite ........ : vide
Mots clefs ... : 
*****************************************************************}
function TOF_GRILLEFILTRE.EstRemplit(vInLigne: Integer): Boolean;
var lInCol    : Integer ;
    lInDebut  : Integer ;
begin
  Result := False ;
  if FListe.DBIndicator
    then lInDebut := 1
    else lInDebut := 0 ;
  for lInCol := lInDebut to FListe.RowCount - 1 do
     if Trim( FListe.Cells[lIncol , vInLigne] ) <> '' then
       begin
       Result := True ;
       Exit ;
       end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Interdit l'accès aux cellules qui appartiennent à la clé
Suite ........ : primaire et qui sont renseignés
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if pos ( IntToStr( FListe.col ) , ListeColsUniq ) <= 0  then Exit ;
  Cancel := not ( EstNewEnreg ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... : 11/03/2003
Description .. : Affecte les évènements de raffraichissement des données
Suite ........ : s'il on modifie un des composants de recherche.
Suite ........ : Valable uniquement pour les composants qui ont des noms 
Suite ........ : de champs.
Suite ........ : 
Suite ........ : A PREVOIR : rendre paramètrable le type d'évènement sur 
Suite ........ : lequel se positionne le raffraichissement des données 
Suite ........ : (actuellement OnChange, mais pourquoi pas OnExit, ....)
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.AffecterEvtFiltre;
var lInCpt  : Integer ;
    TOBTest : TOB ;
begin
  TOBTest := TOB.Create(nomTable,nil,-1);
  // Parcours des composants de la fiche
  for lInCpt := 0 to Ecran.ComponentCount - 1 do
    begin
    if not TOBTest.FieldExists(Ecran.Components[lInCpt].Name) then continue ;
    // Cas du THEdit
    if Ecran.Components[lInCpt] is THEdit then
      THEdit( Ecran.Components[lInCpt] ).OnChange := LancerRecherche ;
    // Cas du THValComboBox
    if Ecran.Components[lInCpt] is THValComboBox then
      THValComboBox( Ecran.Components[lInCpt] ).OnChange := LancerRecherche ;
    // Cas du TCheckBox
    if Ecran.Components[lInCpt] is TCheckBox then
      TCheckBox( Ecran.Components[lInCpt] ).OnClick := LancerRecherche ;
    end ; // fin for
  TOBTest.Free ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... : 11/03/2003
Description .. : Provoque le raffraichissement des données (relecture base
Suite ........ : + maj grille) avec maj de la ligne courante avant si besoin
Mots clefs ... : 
*****************************************************************}
procedure TOF_GRILLEFILTRE.LancerRecherche(Sender: TObject);
var
	lStSQL    : String;
	lStWhere  : String;
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  lQuery    : TQuery ;
  {$ENDIF}
begin
  {JP 18/08/04 : On s'assure que les critères de recherche sont valides : par défaut renvoie True}
  if not IsCritereValide then Exit;
	// Validation des données si nécessaire
//	if Sender <> nil then
  if not ValiderLigne(FListe.Row,True) then Exit;

  // Construction requête
	lStSQL    := 'SELECT * FROM ' + NomTable ;
  lStWhere  := GenererCondition ;
  if lStWhere <> ''
    then lStSQL := lStSQL + ' WHERE ' + lStWhere ;
  if Trim(ChampTri) <> ''
    then lStSQL := lStSQL + ' ORDER BY ' + ChampTri ;

  // chargement des données
  {$IFDEF EAGLCLIENT}
    MaxRow := (FListe.Height div FListe.DefaultRowHeight) ;
    TOBLignes := OpenSql( lStSQL, True, MaxRow );
    // MAJ boutons fetch
    SetControlEnabled('BNextQ' , Not TOBLignes.GetValue('EOF') ) ;
    SetControlEnabled('BPrevQ', FListe.TopRow > 1 ) ;
  {$ELSE}
    lQuery := OpenSql( lStSQL, True );
    TOBLignes.ClearDetail;
    TOBLignes.LoadDetailDB(NomTable,'','',lQuery,False);
    Ferme(lQuery);
  {$ENDIF}

  // Evènement virtuel - chargement des données de la table en tob.
  OnLoadRecord ( TOBLignes ) ;

  // Positionnement initiale
  FListe.Col := 2;
  FListe.Row := 1;

  // Affichage
  RefreshFListe ;

end;

{$IFDEF EAGLCLIENT}
{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : Gestion du chargement par page des données en eAGL
Mots clefs ... : 
*****************************************************************}
procedure TOF_GRILLEFILTRE.BNextQClick(Sender: TObject);
var lInAdd  : Integer ;
    lInCpt  : Integer ;
    lInRow  : Integer ;
begin

	// Validation des données si nécessaire
	if Sender <> nil then
    if not ValiderLigne(FListe.Row,True) then Exit;

  lInAdd := AppendSQL( TOBLignes , MaxRow ) ;
  if lInAdd > 0 then
    begin
    for lInCpt := 0 to (lInAdd - 1) do
      begin
      FListe.RowCount := FListe.RowCount + 1 ;
      // indexs enreg dans Q
      lInRow := TOBLignes.Detail.Count - lInAdd + lInCpt ;
      // Affichage Grille
      TOBLignes.Detail[ lInRow ].PutLigneGrid( FListe , FListe.RowCount - 1 , false , false, ListeChamps ) ;
      // MAJ TOB
      end ;
    end ;

  SetControlEnabled('BNextQ' , Not TOBLignes.GetValue('EOF') ) ;
  SetControlEnabled('BPrevQ', FListe.TopRow > 1 ) ;
end;
{$ENDIF}

{$IFDEF EAGLCLIENT}
procedure TOF_GRILLEFILTRE.BPrevQClick(Sender: TObject);
begin
  SendNotifyMessage( TWinControl(FListe).Handle, WM_KeyDown, 33, 0 ) ;
end;

{$ENDIF}

procedure TOF_GRILLEFILTRE.RemplitGrille;
begin
  LancerRecherche(nil) ;
end;

procedure TOF_GRILLEFILTRE.FListeKeyPress(Sender: TObject; var Key: Char);
begin
  if Key <> #9 then
    LigneModifiee := True ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 13/03/2003
Modifié le ... :   /  /    
Description .. : Retourne vrai si la ligne courante est un nouvel 
Suite ........ : enregistrement
Mots clefs ... : 
*****************************************************************}
function TOF_GRILLEFILTRE.EstNewEnreg: Boolean;
begin
  if TOBCourante = nil then
    begin
    Result := False ;
    Exit ;
    end ;
  if not TOBCourante.FieldExists('NEWENREG') then
    begin
    Result := False ;
    Exit ;
    end ;
  Result :=  TOBCourante.GetValue('NEWENREG') ;
end;

function TOF_GRILLEFILTRE.ExistDB( vTOBEnreg, vTOBOrig : TOB ) : Boolean ;
begin
  if EstNewEnreg then
    result := vTOBEnreg.ExistDB
  else
    begin
    vTOBEnreg.ChargeCle1 ;
    vTOBOrig.ChargeCle1 ;
    Result := ExisteSQL('SELECT * FROM ' + nomTable +
                                ' WHERE ' + vTOBEnreg.Cle1 +
                                ' AND NOT (' + vTOBOrig.Cle1 + ')' ) ;
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : COMPTA
Créé le ...... : 22/07/2004
Modifié le ... :   /  /    
Description .. : Retourne la valeur du champ "vStChamps" contenu dans la
Suite ........ : TOB courante 
Mots clefs ... :
*****************************************************************}
function TOF_GRILLEFILTRE.GetField(vStChamps : String ; row : integer = -1) : Variant;
begin
    if (row < 0) then row := FListe.row-1;

    if ((TOBLignes <> nil) and (row < TOBLignes.Detail.count)) then result := TOBLignes.Detail[row].GetValue(vStChamps)
    else result := null;
end;

{***********A.G.L.***********************************************
Auteur  ...... : COMPTA
Créé le ...... : 22/07/2004
Modifié le ... :   /  /
Description .. : Affecte la valeur "vstValeur" au champs "vStChamps" de la
Suite ........ : Tob courante
Mots clefs ... :
*****************************************************************}
procedure TOF_GRILLEFILTRE.SetField(vStChamps : String ; vStValeur : Variant ; row : integer = -1);
begin
    if (row < 0) then row := FListe.row-1;

    if ((TOBLignes <> nil) and (row < TOBLignes.Detail.count)) then TOBLignes.Detail[row].PutValue(vStChamps,vStValeur);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_GRILLEFILTRE.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  InsererLigne;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_GRILLEFILTRE.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  SupprimerLigne;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_GRILLEFILTRE.BDefaireClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AnnulerModif;
end;

{JP 18/08/04 : On s'assure que les critères de recherche sont valides : par défaut renvoie True
{---------------------------------------------------------------------------------------}
function TOF_GRILLEFILTRE.IsCritereValide : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_GRILLEFILTRE.FListeOnExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
{JP 18/05/06 : FQ 17546}
end;

Initialization
  registerclasses ( [ TOF_GRILLEFILTRE ] ) ;
end.
