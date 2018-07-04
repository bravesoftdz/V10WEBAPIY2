{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 19/03/2002
Modifié le ... : 25/04/2002
Description .. : Source TOF de la FICHE : CPCHPLIBRELIB ()
Mots clefs ... : TOF;CPCHPLIBRELIB
*****************************************************************}
Unit CPCHPLIBRELIB_TOF ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, windows,
     ComCtrls, HCtrls, HEnt1, HMsgBox, extCtrls,
{$IFDEF EAGLCLIENT}
     eFiche,maineagl,eFichList,HPdfPrev,UtileAGL,
{$ELSE}
     db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fiche,FE_main,FichList,EdtREtat,
{$ENDIF}
     utilPGI, // EstTablePartagee / openselect...
     UTOF,
     UTOB ;

procedure CPLanceFiche_ParamChampsLibres(Args : String);

Type
  TOF_CPCHPLIBRELIB = Class (TOF)
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
    TypeTable : String ;
    TOBLignes : TOB;			// TOb de la liste des champs
    TOBCourante : TOB;		// TOB correspondant à la ligne sélectionnée
    FListe : THGrid;
    RTypeMvt : TRadioGroup;
    FDossier : string ;
    procedure AnnulerModif(Sender : TObject);
    procedure BoutonsNavMaj;
    procedure ChoixTypeMvt(Sender : TObject);
    function  GenererCondition: String;
    procedure Imprimer(Sender : TObject);
    procedure InverseSelection(row : Integer);
    procedure ParcoursListe (Sender : TObject);
    procedure RaffraichirLigne(row : Integer; cols : String = 'PL_LIBELLE;PL_VISIBLE');
    function  ValiderLigne(row : integer; avecConfirm : Boolean = False): Boolean;
    procedure BValiderClick(Sender : TObject); // FQ 15440
  end ;

Implementation
uses HTB97;

// Index de la colonne avec saisie case à cocher
Const ColonneCoche : Integer = 2;

// Message d'erreurs
Const MessageListe : Array[0..2] of String =
	 ('Voulez-vous enregistrer les modifications ?',
          'Vous devez renseigner un libellé !',
	  'L''enregistrement est inaccessible !'
         );

//==============================================================================
{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 23/09/2005
Modifié le ... :   /  /
Description .. : Il n'y avait pas d'update sur le bouton valider
Suite ........ : L'enregistrement ne se faisait que sur le onclose
Suite ........ :
Suite ........ : FQ 16538
Mots clefs ... : VALIDER UPDATE
*****************************************************************}
procedure CPLanceFiche_ParamChampsLibres(Args : String);
begin
  AGLLanceFiche('CP','CPCHPLIBRELIB','','',Args);
end;

//==============================================================================


procedure TOF_CPCHPLIBRELIB.OnUpdate ;
begin
  Inherited ;
  ValiderLigne(FListe.row);
end ;

procedure TOF_CPCHPLIBRELIB.OnLoad ;
begin
  Inherited ;
  // Se positionner sur la 1ere entite
  RTypeMvt.ItemIndex := 0 ;
  ChoixTypeMvt(nil) ;
end ;

procedure TOF_CPCHPLIBRELIB.OnArgument (S : String ) ;
var
  sTable, sLibTable : String;
begin
  Inherited ;

  // Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then FDossier := TableToBase( 'NATCPTE' )
    else FDossier := '' ;

  // Mise en place des arguments d'appel
  sTable := readTokenSt(S);
  sLibTable := readTokenSt(S);
  if sTable<>'' then
     begin
     Ecran.Caption := sLibTable ;
     TypeTable := sTable ;
     SetControlVisible('PTop',False);
     end
  else
     TypeTable := '' ;
  // Récupération des controles importants
  FListe := THGrid(GetControl('FListe'));
  RTypeMvt := TRadioGroup(GetControl('RTMVT'));
  // Paramétrage de la grille
  FListe.ColLengths[2] := 35;
  FListe.ColLengths[ColonneCoche] := 1;
  FListe.ColTypes[ColonneCoche] := 'B'; // 3ème col boolean
  FListe.ColFormats[ColonneCoche] := IntToStr(Integer(csCoche)); // affichage coche
  // Tob contenant les enregistrements
  TOBLignes := TOB.Create('Champs libres',nil,-1);
  // Réaffectation des évènements
  TButton(GetControl('BFirst')).onClick := ParcoursListe;
  TButton(GetControl('BPrev')).onClick := ParcoursListe;
  TButton(GetControl('BNext')).onClick := ParcoursListe;
  TButton(GetControl('BLast')).onClick := ParcoursListe;
  TButton(GetControl('BImprimer')).onClick := Imprimer;
  TButton(GetControl('bDefaire')).onClick := AnnulerModif;
  FListe.OnRowEnter := DoRowEnter;
  FListe.OnRowExit := DoRowExit;
  FListe.OnKeyPress := DoKeyPress;
  FListe.OnDblClick := DoDblClick;
  Ecran.OnKeyDown := FormKeyDown;
  RTypeMvt.onClick := choixTypeMvt;
  TToolBarButton97(GetControl('BValider')).OnClick := BValiderClick;
end ;

procedure TOF_CPCHPLIBRELIB.OnClose ;
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
end ;

procedure TOF_CPCHPLIBRELIB.choixTypeMvt(Sender: TObject);
var req   : String;
    Q     : TQuery;
    param : Boolean;
begin
  // MAJ du type de table
  if GetControlVisible('PTop') then
    case (RTypeMvt.ItemIndex) of
      0 : TypeTable := 'E' ;
      1 : TypeTable := 'Y' ;
      2 : TypeTable := 'BE' ;
      end ;

  // Contruction de la requête
  Req := 'SELECT * FROM PARAMLIB WHERE ' + GenererCondition + ' ORDER BY PL_CHAMP';
  // Execution de la requête
  Q := OpenSelect( req, FDossier );
  // MAJ de la TOB
  TOBLignes.ClearDetail;
  TOBLignes.LoadDetailDB('PARAMLIB','','',Q,False);
  Ferme(Q);
  // Affichage
  FListe.rowCount := 2;
  TOBLignes.PutGridDetail(FListe,False,False,'PL_LIBELLE;PL_VISIBLE',False);
  // Initialisation de la Grille
  FListe.Col := 1;
  FListe.Row := 1;
  param := False;
  DoRowEnter(nil, 1, param, False);
end;

procedure TOF_CPCHPLIBRELIB.InverseSelection(row: Integer);
begin
  // Mise à jour de la Grille
  if FListe.Cells[ColonneCoche,row] = '-'
    then FListe.Cells[ColonneCoche,row] := 'X'
    else FListe.Cells[ColonneCoche,row] := '-';
end;

function TOF_CPCHPLIBRELIB.GenererCondition: String;
var Q : TQuery ;
    n : integer ;
    sTable,sWhere,sChamp : string ;
begin
  sWhere := '' ;
  case RTypeMvt.ItemIndex of
     0 : sTable := 'E' ;
     1 : sTable := 'A' ;
     2 : sTable := 'U' ;
     end ;
  Q := OpenSelect('SELECT CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="NAT" AND CC_CODE LIKE "'+ sTable + '%" ORDER BY CC_CODE', FDossier) ;
  n := 0 ;
  while not Q.EOF do
     begin
     if Q.FindField('CC_ABREGE').AsString <> 'X' then
        begin
        sChamp := TypeTable + '_TABLE' + IntToStr(n) ;
        sWhere := sWhere + ' AND PL_CHAMP<>"' + sChamp + '"' ;
        end ;
     Q.Next ;
     Inc(n) ;
     end ;
  Ferme(Q) ;
  Result := 'PL_TABLE="' + TypeTable + '" ' + sWhere ;
end;

procedure TOF_CPCHPLIBRELIB.DoRowEnter(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  // Mise à jour de la TOB courante
  TOBCourante := TOBLignes.Detail[ou-1];
  // MAJ des boutons de navigations
  BoutonsNavMaj;
end;

procedure TOF_CPCHPLIBRELIB.DoRowExit(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  // On quitte une ligne, validation des modifs...
  ValiderLigne(ou);
end;

procedure TOF_CPCHPLIBRELIB.ParcoursListe(Sender: TObject);
var
  ctrlName : String; // nom du controle
  oldRow, newRow : Integer;
begin
  ctrlName := TButton(Sender).name;
  oldRow := FListe.Row;
  newRow := FListe.Row;
  // Première ligne
  if ctrlName = 'BFIRST'
    then newRow := 1
  // Ligne précédente
  else if ctrlName = 'BPREV' then
    begin
    if FListe.Row > 1 then
      newRow := FListe.Row - 1;
    end
  // Ligne suivante
  else if ctrlName = 'BNEXT' then
    begin
    if FListe.Row < TOBLignes.Detail.Count then
      newRow := FListe.Row + 1;
    end
  // Dernière ligne
  else if ctrlName = 'BLAST'
    then newRow := TOBLignes.Detail.Count ;

  // Déplacement
  if newRow <> oldRow then
    FListe.gotoRow(newRow);
end;

procedure TOF_CPCHPLIBRELIB.AnnulerModif(Sender: TObject);
var
  ligne : Integer;
begin
  // Test indice...
  ligne := FListe.row;
  if (ligne < 1) or (ligne > TOBLignes.Detail.count) then Exit;
  raffraichirLigne(ligne);
end;

procedure TOF_CPCHPLIBRELIB.BoutonsNavMaj;
begin
  // MAJ des boutons de navigations
  SetControlEnabled('BFirst',FListe.Row > 1);
  SetControlEnabled('BPrev',FListe.Row > 1);
  SetControlEnabled('BNext',FListe.Row < TOBLignes.Detail.Count);
  SetControlEnabled('BLast',FListe.Row < TOBLignes.Detail.Count);
end;

procedure TOF_CPCHPLIBRELIB.Imprimer(Sender: TObject);
var titre : String;
begin
  titre := 'Zones libres des mouvements ' + RTypeMvt.Items[RTypeMvt.ItemIndex];
  LanceEtat('E','CPE','LCL',True, False, False, Nil, GenererCondition, titre, False, 0);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/06/2007
Modifié le ... :   /  /    
Description .. : Validation des modifactions
Suite ........ : FQ 15440
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCHPLIBRELIB.BValiderClick(Sender : TObject);
begin
   if ValiderLigne(FListe.Row,true) then
      inherited;
end;

function TOF_CPCHPLIBRELIB.ValiderLigne( row : integer ; avecConfirm : Boolean ) : Boolean ;
// Valide et enregistre les données de la ligne d'index 'row'
// si avecConfirm = true alors demande confirmation
// Retourne True si le traitement peut continuer, False sinon
var Rep : Integer;
begin
  Result := True;
  // Mise à jour TOB
  TOBCourante.GetLigneGrid(FListe,row, ';PL_LIBELLE;PL_VISIBLE');
  // données modifiées ?
  if TOBCourante.IsOneModifie	then
    BEGIN
    // Y-a-t-il une demande de confirmation à faire ?
    if avecConfirm
      then Rep := PGIAskCancel(MessageListe[0],Ecran.Caption)
      else Rep := mrYes;

    Case Rep of
      // Validation des modifications
      mrYes :
        BEGIN
          // Le libellé est obligatoire
          if TOBCourante.GetValue('PL_LIBELLE') = '' then
            BEGIN
            PGIInfo(MessageListe[0],Ecran.Caption);
            Exit;
            END;
          // MAJ du champ visible
          if TOBCourante.GetValue('PL_VISIBLE') <> 'X'
            then TOBCourante.PutValue('PL_VISIBLE','-');
          // Enregistrement
          UpdateTobMS( TOBCourante, FDossier ) ;
	  TOBCourante.modifie := False;
        END;

      // Rechargement
      mrNo  :
      	BEGIN
        // On recharge les données de la TOB
        if not TOBCourante.LoadDB then
          // Pb : données inaccessibles !
          BEGIN
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

procedure TOF_CPCHPLIBRELIB.raffraichirLigne(row : Integer; cols : String);
begin
  // Rechargement des données de la TOB dans la ligne de la Grid
  TOBCourante.PutLigneGrid(FListe, row, FALSE, FALSE, cols);
end;

procedure TOF_CPCHPLIBRELIB.DoKeyPress(Sender: TObject; var Key: Char);
begin
  // Traitement pour la colonne 'Visible' uniquement
  if FListe.Col <> ColonneCoche then exit ;
  if Key = ' ' then InverseSelection(FListe.row) ;
  if (Key<>'-') and (Key<>'X') and (Key<>'x') then Key:=#0 ;
  if Key='x' then Key:='X' ;
end;

procedure TOF_CPCHPLIBRELIB.DoDblClick(Sender: TObject);
begin
  InverseSelection(FListe.row);
end;

procedure TOF_CPCHPLIBRELIB.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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

Initialization
  registerclasses ( [ TOF_CPCHPLIBRELIB ] ) ;
end.
