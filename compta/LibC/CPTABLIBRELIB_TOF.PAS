{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 12/03/2002
Modifié le ... : 29/05/2002
Description .. : Source TOF de la FICHE : CPTABLELIBRELIB ()
Mots clefs ... : TOF;CPTABLELIBRELIB
*****************************************************************}
Unit CPTABLIBRELIB_TOF ;

Interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox, SaisUtil, UiUtil,
      Graphics, Ent1, HPanel, Hqry, HSysMenu,
      UtilPGI, // EstTablePartagee / openselect...
      Windows, grids,
{$IFDEF EAGLCLIENT}
      eFiche,maineagl,eFichList,HPdfPrev,UtileAGL,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
      Fiche,FE_main,FichList, EdtREtat, HDB,
{$ENDIF}
      UTOF, UTOB, LibChpLi;

procedure CPLanceFiche_ParamTablesLibres;

Type
  TOF_CPTABLIBRELIB = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure DoRowEnter(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean);
    procedure DoRowExit(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean);
    procedure DoKeyPress(Sender: TObject; var Key: Char);
    procedure DoDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    TOBLignes : TOB;
    TOBCourante : TOB;
    Tl : THValComboBox;
    FListe : THGrid;
    FDossier : string ;
    procedure AnnulerModif(Sender : TObject);
    procedure BoutonsNavMaj;
    procedure ChargeCombos;
    function  GenererCondition: String;
    procedure Imprimer(Sender : TObject);
    procedure InverseSelection(row : Integer);
    procedure ModifierChamps(Sender : TObject);
    procedure ParcoursListe (Sender : TObject);
    procedure raffraichirLigne(row : Integer; cols : String = 'CC_CODE;CC_LIBELLE;CC_ABREGE');
    procedure SelectionEntite (Sender : TObject);
    function  ValiderLigne(row : integer; avecConfirm : Boolean = False): Boolean;
    procedure MajParamLib ;    
  end ;

Implementation

Const ColonneCoche : Integer = 3;		// Index de la colonne avec saisie case à cocher
Const EntiteQte : Integer = 9;			// Entités sélectionnables
Const EntiteListe : Array[0..8] of String =
                     ('G;Comptes généraux;',
                      'T;Comptes auxiliaires;',
                      'B;Comptes budgétaires;',
                      'D;Sections budgétaires;',
                      'S;Sections analytiques;',
                      'E;Ecritures comptables;',
                      'A;Ecritures analytiques;',
                      'U;Ecritures budgétaires;',
                      'I;Immobilisations;'
                      );
Const MessageListe : Array[0..2] of String =		// Message d'erreurs
					('Voulez-vous enregistrer les modifications ?',
           'Vous devez renseigner un libellé !',
					 'L''enregistrement est inaccessible !'
           );

//==============================================================================

procedure CPLanceFiche_ParamTablesLibres;
begin
    AglLanceFiche('CP','CPTABLIBRELIB','','','') ;
end;

//==============================================================================

procedure TOF_CPTABLIBRELIB.OnUpdate ;
begin
  Inherited ;
	ValiderLigne(FListe.row);
end ;

procedure TOF_CPTABLIBRELIB.OnLoad ;
begin
  Inherited ;
  TL.ItemIndex := 0; // Position sur 1ère Entité
  SelectionEntite(nil);
  BoutonsNavMaj;
end ;

procedure TOF_CPTABLIBRELIB.OnArgument (S : String ) ;
begin
  Inherited ;

  // Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then FDossier := TableToBase( 'NATCPTE' )
    else FDossier := '' ;

  // Récupération des controles importants
  FListe := THGrid(GetControl('FListe'));
  Tl     := THValComboBox(GetControl('Tl'));
  // Paramétrage de la grille
  FListe.ColEditables[1] := False; // 1ère col non editables
  FListe.ColLengths[2] := 35; // Lib 35 car maxi.
  FListe.ColLengths[ColonneCoche] := 1;
  FListe.ColTypes[ColonneCoche] := 'B'; // 3ème col boolean
  FListe.ColFormats[ColonneCoche] := IntToStr(Integer(csCoche)); // affichage coche
  // Tob contenant les enregistrements
  TOBLignes := TOB.Create('Tables libres',nil,-1);
  // Remplissage de la combo box
  ChargeCombos;
  // Réaffectation des évènements
  TButton(GetControl('BFirst')).onClick := ParcoursListe;
  TButton(GetControl('BPrev')).onClick := ParcoursListe;
  TButton(GetControl('BNext')).onClick := ParcoursListe;
  TButton(GetControl('BLast')).onClick := ParcoursListe;
  TButton(GetControl('BImprimer')).onClick := Imprimer;
  TButton(GetControl('bDefaire')).onClick := AnnulerModif;
{$IFDEF CCS3}
  SetControlVisible('BChamps', False);
{$ENDIF}
  TButton(GetControl('BChamps')).onClick := ModifierChamps;
  FListe.OnRowEnter := DoRowEnter;
  FListe.OnRowExit := DoRowExit;
  FListe.OnKeyPress := DoKeyPress;
  FListe.OnDblClick := DoDblClick;
  Ecran.OnKeyDown := FormKeyDown;
  Tl.onChange := SelectionEntite;
end ;

procedure TOF_CPTABLIBRELIB.OnClose ;
begin
  Inherited ;
  // Validation des modifications si besoin
  if not ValiderLigne(FListe.Row,True) then
    Begin
    LastError := 1;
    Exit;
    end;
  // Libération mémoire des TOBs
  TOBLignes.Free;
{$IFDEF EAGLCLIENT}    // ajout me fiche 19337  pour les éditions  balance, GL
             AvertirCacheServer('CHOIXCOD') ;
             AvertirCacheServer('NATCPTE') ;
{$ELSE}
             ChargeTableLibre;
{$ENDIF}
   AvertirTable('TTTABLESLIBRESGEN');
   AvertirTable('TTTABLESLIBRESAUX');
   AvertirTable('TTTABLESLIBRESBUDG');
   AvertirTable('TTTABLESLIBRESBUDS');
   AvertirTable('TTTABLESLIBRESECR');
   AvertirTable('TTTABLESLIBRESGEN');
   AvertirTable('TTTABLESLIBRESSEC');

end ;

procedure TOF_CPTABLIBRELIB.ParcoursListe(Sender: TObject);
// Gestion des boutons de parcours de la grille
var
  ctrlName : String; // nom du controle
  oldRow, newRow : Integer;
begin
  ctrlName := Uppercase(TButton(Sender).name);
  oldRow := FListe.Row;
  newRow := FListe.Row;

  // première ligne
  if ctrlName = 'BFIRST' then	newRow := 1
  // Ligne précédente
  else if ctrlName = 'BPREV' then
    begin
    if FListe.Row > 1 then
      newRow := FListe.Row - 1;
    end
  // Ligne suivante
  else  if ctrlName = 'BNEXT' then
    begin
    if FListe.Row < TOBLignes.Detail.Count then
      newRow := FListe.Row + 1;
    end
  // Dernière ligne
  else if ctrlName = 'BLAST' then
    newRow := TOBLignes.Detail.Count;

  // Déplacement :
  if newRow <> oldRow then
    FListe.gotoRow(newRow);

end;

procedure TOF_CPTABLIBRELIB.ChargeCombos;
// Modifier aussi TOF_CPCHOIXTALI.ChargeComboEntite
Var i : Integer ;
    St : String ;
    it,va : String ;
Begin
  Tl.Values.Clear ;
  Tl.Items.Clear ;
  for i:=0 to EntiteQte-1 do
    BEGIN
    St:=EntiteListe[i] ;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St)) ;
    if ((EstSerie(S5)) or (EstSerie(S3))) then
      BEGIN
      // Pas de table libres sur écritures analytiques et écritures budgétaires en S3 et S5
      if (va='A') or (va='U') then Continue;
      // Pas de tables libres sur comptes budgétaires et sections budgétaires en S3
      if EstSerie(S3) then
        begin
        if ((va='B') or (va='D')) then Continue ;
        // Pas de tables libres sur les sections analytiques si pas de séria analytiques
        //if EstComptaSansAna and (va='S') then Continue;
        end;
      END ;
    Tl.Values.Add(va) ;
    Tl.Items.Add(it) ;
    END ;
End ;

procedure TOF_CPTABLIBRELIB.SelectionEntite(Sender: TObject);
// Mise à jour de la liste des tables libres en fonction de l'entité sélectionnée
var  req        : String;
     Q          : TQuery;
     param      : Boolean;
begin
  // Validation des données si nécessaire
  if Sender <> nil then
    if not ValiderLigne(FListe.Row,True) then Exit;
  Req := 'SELECT * FROM CHOIXCOD WHERE ' + GenererCondition + ' ORDER BY CC_CODE';
  Q := OpenSelect( req, FDossier );
  TOBLignes.ClearDetail;
  TOBLignes.LoadDetailDB('CHOIXCOD','','',Q,False);
  FListe.rowCount := 2;
  TOBLignes.PutGridDetail(FListe,False,False,'CC_CODE;CC_LIBELLE;CC_ABREGE',False);
  Ferme(Q);
  FListe.Col := 2;
  FListe.Row := 1;
  param := False;
  DoRowEnter(nil, 1, param, False);
end;

procedure TOF_CPTABLIBRELIB.BoutonsNavMaj;
begin
  // MAJ des boutons de navigations
  SetControlEnabled('BFirst',FListe.Row > 1);
  SetControlEnabled('BPrev',FListe.Row > 1);
  SetControlEnabled('BNext',FListe.Row < TOBLignes.Detail.Count);
  SetControlEnabled('BLast',FListe.Row < TOBLignes.Detail.Count);
end;

procedure TOF_CPTABLIBRELIB.DoRowEnter(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
// Arrivée dans une nouvelle ligne, Gestion de l'acces à la perso. des champs
var
  visible : Boolean;
begin
  TOBCourante := TOBLignes.Detail[ou-1]; 	// Mise à jour de la TOB courante
  visible := Uppercase(TOBCourante.GetValue('CC_ABREGE')) = 'X';
  SetControlEnabled('BCHAMPS', visible); 	// Perso des champs uniquement si table visible
  BoutonsNavMaj; 													// MAJ des boutons de navigations
end;

procedure TOF_CPTABLIBRELIB.DoRowExit(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  ValiderLigne(ou); 	// On quitte une ligne, validation des modifs...
end;

procedure TOF_CPTABLIBRELIB.AnnulerModif(Sender : TObject);
var
  ligne : Integer;
begin
  ligne := FListe.row;
  if (ligne < 1) or (ligne > TOBLignes.Detail.count) then Exit;
  raffraichirLigne(ligne);
end;

procedure TOF_CPTABLIBRELIB.Imprimer(Sender : TObject);
var titre : String;
begin
  titre := 'Tables libres des ' + Tl.items[Tl.ItemIndex];
  LanceEtat('E','CPE','LTL',True, False, False, Nil, GenererCondition, titre, False, 0);
end;

procedure TOF_CPTABLIBRELIB.ModifierChamps(Sender : Tobject);
var visible : boolean;
begin
  // Peut-modifier les champs de la ligne sélectionnée ?
  visible := Uppercase(TOBCourante.GetValue('CC_ABREGE')) = 'X';
  if not visible then
    Begin
    SetControlEnabled('BCHAMPS', visible);
    Exit;
    End;
  // Ouverture de la fiche de paramétrage des champs
  AGLModifLibelleChampLibre(TOBCourante.GetValue('CC_CODE'),TOBCourante.GetValue('CC_LIBELLE')) ;
end;

function TOF_CPTABLIBRELIB.ValiderLigne(row : integer; avecConfirm : Boolean = False): Boolean;
// Valide et enregistre les données de la ligne d'index 'row'
// si avecConfirm = true alors demande confirmation
// Retourne True si le traitement peut continuer, False sinon
var Rep : Integer;
begin
  Result := True;
  TOBCourante.GetLigneGrid(FListe,row, ';CC_CODE;CC_LIBELLE;CC_ABREGE');   // Mise à jour TOB
  if TOBCourante.IsOneModifie	then   // données modifiées ?
    BEGIN
    // Y-a-t-il une demande de confirmation à faire ?
    if avecConfirm
      then Rep := PGIAskCancel(MessageListe[0],Ecran.Caption)
      else Rep := mrYes;

    Case Rep of
      // Validation des modifications
      mrYes :
        BEGIN
          if TOBCourante.GetValue('CC_LIBELLE') = '' then // Le libellé est obligatoire
            BEGIN
            PGIInfo(MessageListe[0],Ecran.Caption);
            Exit;
            END;
          if TOBCourante.GetValue('CC_ABREGE') <> 'X' // MAJ du champ visible
          then TOBCourante.PutValue('CC_ABREGE','-');
          UpdateTobMS( TOBCourante, FDossier ); // Enregistrement
          MajParamLib;
          TOBCourante.modifie := False;
        END;

      // Rechargement
      mrNo  :
      	BEGIN
        // On recharge les données de la TOB
        if not TOBCourante.LoadDB then
          BEGIN         	// Pb : données inaccessibles !
          PGIBox(MessageListe[2],Ecran.Caption);
          Result := False;
          END
        else
          BEGIN
          // MAJ affichage dans la grid
          raffraichirLigne(row);
          Result := True;
          END
      	END;

      // Annulation
      mrCancel : Result := False ;
      End ;

    END;
end;

procedure TOF_CPTABLIBRELIB.MajParamLib;
var
  Code,TypeTable : String ;
  n : Integer ;
begin
  Code := TOBCourante.GetValue('CC_CODE');
  n := StrToInt(Copy(Code,Length(Code),1)) ;
  TypeTable := Tl.Value ;
  if Tl.Value='A' then TypeTable:='Y' else
  if Tl.Value='U' then TypeTable:='BE' ;
  ExecuteSQL('UPDATE ' + GetTableDossier( FDossier, 'PARAMLIB' )
                       + ' SET PL_LIBELLE="' + TOBCourante.GetValue('CC_LIBELLE') + '", '
                           + ' PL_VISIBLE="' + TOBCourante.GetValue('CC_ABREGE') + '" '
                       + ' WHERE PL_TABLE="' + TypeTable + '" '
                         + ' AND PL_CHAMP="' + TypeTable + '_TABLE' + IntToStr(n) + '"' );
end;

procedure TOF_CPTABLIBRELIB.InverseSelection(row : Integer);
begin
  // Mise à jour de la Grille
  if FListe.Cells[ColonneCoche,row] = '-'
    then FListe.Cells[ColonneCoche,row] := 'X'
    else FListe.Cells[ColonneCoche,row] := '-';
end;

function TOF_CPTABLIBRELIB.GenererCondition: String;
begin
  // Par défaut
  Result := 'CC_TYPE="NAT" And CC_CODE Like "' + Tl.Value + '%"';
  // Limitations S5, S3
  if (TL.Value = 'E') then
    BEGIN
    if (EstSerie(S3)) then
       BEGIN
       Result := 'CC_TYPE="NAT" And CC_CODE Like "'+Tl.Value+'%" AND CC_CODE<="'+Tl.Value+'00" ' ;
       END else
    if (EstSerie(S5)) then
       BEGIN
       Result := 'CC_TYPE="NAT" And CC_CODE Like "'+Tl.Value+'%" AND CC_CODE<="'+Tl.Value+'01" ' ;
       END else
       BEGIN
       Result := 'CC_TYPE="NAT" And CC_CODE Like "'+Tl.Value+'%" ' ;
       END ;
    END else
    BEGIN
    if EstSerie(S3) then
       BEGIN
       Result := 'CC_TYPE="NAT" And CC_CODE Like "'+Tl.Value+'%" AND CC_CODE<="'+Tl.Value+'02" ' ;
       END else
       BEGIN
       Result := 'CC_TYPE="NAT" And CC_CODE Like "'+Tl.Value+'%" ' ;
       END ;
    END ;
end;

procedure TOF_CPTABLIBRELIB.raffraichirLigne(row: Integer; cols : String);
begin
  // Rechargement des données de la TOB dans la ligne de la Grid
  TOBCourante.PutLigneGrid(FListe, row, FALSE, FALSE, cols);
end;

procedure TOF_CPTABLIBRELIB.DoKeyPress(Sender: TObject; var Key: Char);
begin
  // Traitement pour la colonne 'Visible' uniquement
  if FListe.Col <> ColonneCoche then exit ;
  if Key = ' ' then InverseSelection(FListe.row) ;
  if (Key<>'-') and (Key<>'X') and (Key<>'x') then Key:=#0 ;
  if Key='x' then Key:='X' ;
end;

procedure TOF_CPTABLIBRELIB.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
   OkG , Vide : boolean;
begin
  inherited;

  if Not FListe.SynEnabled then
    BEGIN
    Key:=0;
     Exit;
    END;
  OkG:=(Screen.ActiveControl=FListe);
  Vide:=(Shift=[]);
  // Validation
  if key = VK_RETURN then
  if ((OkG) and (Vide)) then KEY:=VK_TAB;
end;

procedure TOF_CPTABLIBRELIB.DoDblClick(Sender: TObject);
begin
  InverseSelection(FListe.row);
end;

Initialization
  registerclasses ( [ TOF_CPTABLIBRELIB ] ) ;
end.
