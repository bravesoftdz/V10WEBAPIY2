{***********UNITE*************************************************
Auteur  ...... : Sylvie DE ALMEIDA
Créé le ...... : 12/02/2007
Modifié le ... : 12/02/2007
Description .. : Source TOF de la FICHE : YYMODELECATREG ()
Mots clefs ... : TOF;YYMODELECATREG
*****************************************************************}
Unit YYMODELECATREG_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     UtilPGI,
{$IFNDEF EAGLCLIENT}
     db,dbTables,Fiche,FE_main,FichList, EdtREtat, HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eFiche,maineagl,eFichList,HPdfPrev,UtileAGL,
     eMul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     UTOB,
     Windows,
     HTB97,
     utobdebug;
{$IFDEF TAXES}
procedure YYLanceFiche_ModeleCatRegime;

Type
  TOF_YYMODELECATREG = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure DoRowEnter(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean);
    procedure DoRowExit(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean);
    procedure DoKeyPress(Sender: TObject; var Key: Char);
    procedure DoDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BChampsClick(Sender : TObject);
  private
    TOBLignes : TOB;
    TOBCourante : TOB;
    CM : THValComboBox;
    CM1 : THValComboBox;
    FListe : THGrid;
    LigneModifiee   : Boolean ;
    MODELE : THEdit;
    CATEGORIE : THEdit;
    CB_GERE : TCheckBox;
    bPasModele : Boolean;
    BValider: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BChamps: TToolbarButton97;
    procedure AnnulerModif(Sender : TObject);
    procedure BoutonsNavMaj;
    procedure ChargeCombos;
    procedure InverseSelection(row : Integer);
    procedure ParcoursListe (Sender : TObject);
    procedure raffraichirLigne(row : Integer; cols : String = 'CC_CODE;CC_LIBELLE;YCR_REGGERE');
    procedure MajGrille (Sender : TObject);
    function  ValiderLigne(row : integer; avecConfirm : Boolean = False): Boolean;
    procedure MajDonnees ;
    procedure SelectionCategorie(Sender : TObject);
  end ;
{$ENDIF}
Implementation

uses Grids;
{$IFDEF TAXES}
Const ColonneCoche : Integer = 3;// Index de la colonne avec saisie case à cocher
Const MessageListe : Array[0..2] of String =// Message d'erreurs
					('Voulez-vous enregistrer les modifications ?',
           'Merci de renseigner un libellé.',
					 'L''enregistrement est inaccessible.'
           );

//==============================================================================

procedure YYLanceFiche_ModeleCatRegime;
begin
    AglLanceFiche('YY','YYMODELECATREG','','','') ;
end;
//==============================================================================

procedure TOF_YYMODELECATREG.OnUpdate ;
begin
  Inherited ;
	ValiderLigne(FListe.row);
  CM.Enabled := True;
  CM1.Enabled := True;
  CB_GERE.Enabled := True;
end ;

procedure TOF_YYMODELECATREG.OnLoad ;
begin
  Inherited ;
  if bPasModele then
    exit;
  CM.ItemIndex := 0;  // Position sur 1er modèle
  CM1.ItemIndex := 0; // Position sur 1ère catégorie du modèle
  MAJGrille(nil);
  BoutonsNavMaj;
  LigneModifiee := False ;
  CM.Enabled := True;
  CM1.Enabled := True;
  CB_GERE.Enabled := True;
end ;

procedure TOF_YYMODELECATREG.OnArgument (S : String ) ;
begin
  Inherited ;
  bPasModele := false;
  // Récupération des contrôles importants
  FListe  := THGrid(GetControl('FListe'));
  CM      := THValComboBox(GetControl('CM'));
  CM1     := THValComboBox(GetControl('CM1'));
  CB_GERE := TCheckBox(GetControl('CB_GERE'));
  CB_GERE.onClick := MajGrille;
  BValider := TToolbarButton97 (GetControl('BValider'));
  BFirst := TToolbarButton97 (GetControl('BFirst'));
  BPrev := TToolbarButton97 (GetControl('BPrev'));
  BNext := TToolbarButton97 (GetControl('BNext'));
  BLast := TToolbarButton97 (GetControl('BLast')); 
  // Paramétrage de la grille
  FListe.ColEditables[1] := False; // 1ère col non éditable
  FListe.ColLengths[2] := 35; // Lib 35 car maxi.
  FListe.ColEditables[2] := False; // 2ième col non éditable
  FListe.ColWidths[3] := 60;
  FListe.ColLengths[ColonneCoche] := 1;
  FListe.ColTypes[ColonneCoche] := 'B'; // 3ème col boolean
  FListe.ColFormats[ColonneCoche] := IntToStr(Integer(csCoche)); // affichage coche
  // Tob contenant les enregistrements
  TOBLignes := TOB.Create('Définition des régimes gérés',nil,-1);
  // Remplissage des listes déroulantes
  ChargeCombos;
  // Réaffectation des événements
  TButton(GetControl('BFirst')).onClick := ParcoursListe;
  TButton(GetControl('BPrev')).onClick := ParcoursListe;
  TButton(GetControl('BNext')).onClick := ParcoursListe;
  TButton(GetControl('BLast')).onClick := ParcoursListe;
  TButton(GetControl('bDefaire')).onClick := AnnulerModif;
  TButton(GetControl('BChamps')).onClick :=BChampsClick ;
  FListe.OnRowEnter := DoRowEnter;
  FListe.OnRowExit := DoRowExit;
  FListe.OnKeyPress := DoKeyPress;
  FListe.OnDblClick := DoDblClick;
  Ecran.OnKeyDown := FormKeyDown;
  CM.onChange := SelectionCategorie;
  CM1.OnChange := MajGrille;
end ;

procedure TOF_YYMODELECATREG.SelectionCategorie;
Var
    it,va : String ;
    strSQL, St, St1 : string;
    OQuery : TQuery ;
Begin
  //Combo : catégorie
  CM1.Values.Clear ;
  CM1.Items.Clear ;

  strSQL := 'SELECT "TX1", YMODELETAXE.YMT_LIBCAT1 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT1 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT1 <> '' AND YMODELETAXE.YMT_CODE = "'+CM.values[CM.itemindex]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX2", YMODELETAXE.YMT_LIBCAT2 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT2 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT2 <> '' AND YMODELETAXE.YMT_CODE  = "'+CM.values[CM.itemindex]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX3", YMODELETAXE.YMT_LIBCAT3 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT3 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT3 <> '' AND YMODELETAXE.YMT_CODE = "'+CM.values[CM.itemindex]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX4", YMODELETAXE.YMT_LIBCAT4 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT4 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT4 <> '' AND YMODELETAXE.YMT_CODE = "'+CM.values[CM.itemindex]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX5", YMODELETAXE.YMT_LIBCAT5 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT5 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT5 <> '' AND YMODELETAXE.YMT_CODE = "'+CM.values[CM.itemindex]+'" AND YMODELETAXE.YMT_FERME <> "X"';

  OQuery := OpenSql (strSQL, TRUE);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    CM1.Values.Add(va) ;
    CM1.Items.Add(it) ;
    OQuery.next;
  end
  finally
    Ferme(OQuery);
  end;
  CM1.ItemIndex := 0; // Position sur 1ère catégorie du modèle
  ValiderLigne(FListe.row, True); // SDA le 16/02/2007
  MajGrille(nil);
end;

procedure TOF_YYMODELECATREG.OnClose ;
begin
  Inherited ;
  if bPasModele then exit; 
  // Validation des modifications si besoin
  if not ValiderLigne(FListe.Row,True) then
    Begin
    LastError := 1;
    Exit;
    end;
  // Libération mémoire des TOBs
  FreeAndNil(TOBLignes);
end ;

procedure TOF_YYMODELECATREG.ParcoursListe(Sender: TObject);
// Gestion des boutons de parcours de la grille
var
  ctrlName : String; // nom du contrôle
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

procedure TOF_YYMODELECATREG.ChargeCombos;
Var
    it,va : String ;
    strSQL, St, St1 : string;
    OQuery : TQuery ;

Begin
  //Combo : modèles de taxe
  CM.Values.Clear ;
  CM.Items.Clear ;
  strSQL := 'SELECT YMODELETAXE.YMT_CODE, YMODELETAXE.YMT_LIBELLE FROM YMODELETAXE WHERE YMODELETAXE.YMT_FERME <> "X"  ORDER BY YMODELETAXE.YMT_LIBELLE';
  OQuery := OpenSql (strSQL, TRUE);
  try
  if OQuery.Eof then
  begin
    bPasModele := True;
    bFirst.Enabled := false;
    bPrev.Enabled := false;
    bNext.Enabled := false;
    bLast.Enabled := false;
    exit;
  end;
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    CM.Values.Add(va) ;
    CM.Items.Add(it) ;
    OQuery.next;
  end;
  finally
    Ferme (OQuery);
  end;

  //Combo : catégorie
  CM1.Values.Clear ;
  CM1.Items.Clear ;

  strSQL := 'SELECT "TX1", YMODELETAXE.YMT_LIBCAT1 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT1 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT1<> '' AND YMODELETAXE.YMT_CODE = "'+CM.values[0]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX2", YMODELETAXE.YMT_LIBCAT2 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT2 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT2<> '' AND YMODELETAXE.YMT_CODE  = "'+CM.values[0]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX3", YMODELETAXE.YMT_LIBCAT3 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT3 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT3<> '' AND YMODELETAXE.YMT_CODE  = "'+CM.values[0]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX4", YMODELETAXE.YMT_LIBCAT4 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT4 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT4<> '' AND YMODELETAXE.YMT_CODE = "'+CM.values[0]+'" AND YMODELETAXE.YMT_FERME <> "X"';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX5", YMODELETAXE.YMT_LIBCAT5 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT5 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT5<> '' AND YMODELETAXE.YMT_CODE = "'+CM.values[0]+'" AND YMODELETAXE.YMT_FERME <> "X"';

  OQuery := OpenSql (strSQL, TRUE);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    CM1.Values.Add(va) ;
    CM1.Items.Add(it) ;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;
End ;

procedure TOF_YYMODELECATREG.MajGrille(Sender: TObject);
// Chargement de la grille de données
var
     param      : Boolean;
begin

   //if Sender <> nil then
   // if not ValiderLigne(FListe.Row,True) then Exit;
   FListe.VidePile(False) ;
   TOBLignes.ClearDetail;
   if CB_GERE.State = cbGrayed then
    TOBLignes.LoadDetailDBFromSql ('', 'SELECT CHOIXCOD.CC_CODE, CHOIXCOD.CC_LIBELLE, YMODELECATREG.YCR_REGGERE FROM CHOIXCOD LEFT OUTER JOIN YMODELECATREG ON CHOIXCOD.CC_CODE=YMODELECATREG.YCR_CODEREG AND YMODELECATREG.YCR_CODEMODELE = "' + CM.values[CM.itemindex] +'" and YMODELECATREG.YCR_CODECAT = "' + CM1.values[CM1.itemindex] +'" WHERE CHOIXCOD.CC_TYPE = "RTV" ORDER BY CHOIXCOD.CC_CODE')
   else
    if CB_GERE.State = cbChecked then
      TOBLignes.LoadDetailDBFromSql ('', 'SELECT CHOIXCOD.CC_CODE, CHOIXCOD.CC_LIBELLE, YMODELECATREG.YCR_REGGERE FROM CHOIXCOD LEFT OUTER JOIN YMODELECATREG ON CHOIXCOD.CC_CODE=YMODELECATREG.YCR_CODEREG AND YMODELECATREG.YCR_CODEMODELE = "' + CM.values[CM.itemindex] +'" and YMODELECATREG.YCR_CODECAT = "' + CM1.values[CM1.itemindex] +'" WHERE CHOIXCOD.CC_TYPE = "RTV" AND YMODELECATREG.YCR_REGGERE = "X"  ORDER BY CHOIXCOD.CC_CODE')
    else
      TOBLignes.LoadDetailDBFromSql ('', 'SELECT CHOIXCOD.CC_CODE, CHOIXCOD.CC_LIBELLE, YMODELECATREG.YCR_REGGERE FROM CHOIXCOD LEFT OUTER JOIN YMODELECATREG ON CHOIXCOD.CC_CODE=YMODELECATREG.YCR_CODEREG AND YMODELECATREG.YCR_CODEMODELE = "' + CM.values[CM.itemindex] +'" and YMODELECATREG.YCR_CODECAT = "' + CM1.values[CM1.itemindex] +'" WHERE CHOIXCOD.CC_TYPE = "RTV" AND YMODELECATREG.YCR_REGGERE IS NULL OR YMODELECATREG.YCR_REGGERE = "-" ORDER BY CHOIXCOD.CC_CODE');

   if TOBLignes.Detail.Count > 0 then
   begin
    FListe.rowCount := 2;
    TOBLignes.PutGridDetail(FListe,False,False,'CC_CODE;CC_LIBELLE;YCR_REGGERE',False);
    FListe.Col := 2;
    FListe.Row := 1;
    param := False;
    DoRowEnter(nil, 1, param, False);
   end
   else
   begin
    LigneModifiee := false;
    TOBLignes.PutGridDetail(FListe,False,False,'CC_CODE;CC_LIBELLE;YCR_REGGERE',False);
    FListe.Col := 2;
    FListe.Row := 1;
    param := False;
   end;

end;

procedure TOF_YYMODELECATREG.BoutonsNavMaj;
begin
  // MAJ des boutons de navigations
  SetControlEnabled('BFirst',FListe.Row > 1);
  SetControlEnabled('BPrev',FListe.Row > 1);
  SetControlEnabled('BNext',FListe.Row < TOBLignes.Detail.Count);
  SetControlEnabled('BLast',FListe.Row < TOBLignes.Detail.Count);
end;

procedure TOF_YYMODELECATREG.DoRowEnter(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  Gere : Boolean;
begin
  TOBCourante := TOBLignes.Detail[ou-1]; 	// Mise à jour de la TOB courante
  Gere := Uppercase(TOBCourante.GetValue('YCR_REGGERE')) = 'X';
  BoutonsNavMaj;
end;

procedure TOF_YYMODELECATREG.DoRowExit(Sender: TObject; ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  ValiderLigne(ou); 	// On quitte une ligne, validation des modifs...
end;

procedure TOF_YYMODELECATREG.AnnulerModif(Sender : TObject);
var
  ligne : Integer;
begin
  ligne := FListe.row;
  if (ligne < 1) or (ligne > TOBLignes.Detail.count) then Exit;
  raffraichirLigne(ligne);
end;


function TOF_YYMODELECATREG.ValiderLigne(row : integer; avecConfirm : Boolean = False): Boolean;
// Valide et enregistre les données de la ligne d'index 'row'
// si avecConfirm = true alors demande confirmation
// Retourne True si le traitement peut continuer, False sinon
var Rep : Integer;
begin
  Result := True;
  TOBCourante.GetLigneGrid(FListe,row, ';CC_CODE;CC_LIBELLE;YCR_REGGERE');   // Mise à jour TOB
  // données modifiées ?
  if LigneModifiee then
    BEGIN
    // Y-a-t-il une demande de confirmation à faire ?
    if avecConfirm
      then Rep := PGIAskCancel(MessageListe[0],Ecran.Caption)
      else Rep := mrYes;

    Case Rep of
      // Validation des modifications
      mrYes :
        BEGIN
          if TOBCourante.GetValue('YCR_REGGERE') <> 'X' // MAJ du champ REGGERE
          then TOBCourante.PutValue('YCR_REGGERE','-');
          MajDonnees;
          TOBCourante.modifie := False;
          LigneModifiee := False ;
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
          LigneModifiee := False ;
          Result := True;
          END
      	END;

      // Annulation
      mrCancel :
      BEGIN
         Result := False ;
      END
      End ;

    END;
end;

procedure TOF_YYMODELECATREG.MajDonnees;
var
  CodeReg, CodeCat, CodeMod, Gere : String ;

begin
  CodeMod := CM.values[CM.itemindex];
  //CodeMod := MODELE.EditText;
  //CodeMod := TOBCourante.GetValue('YCR_CODEMODELE');
  CodeCat := CM1.values[CM1.itemindex];
  //CodeCat := TOBCourante.GetValue('YCR_CODECAT');
  //CodeCat := CATEGORIE.EditText;
  CodeReg := TOBCourante.GetValue('CC_CODE');
  Gere := TOBCourante.GetValue('YCR_REGGERE');


  if not ExisteSQL ('SELECT YMODELECATREG.* FROM YMODELECATREG WHERE YMODELECATREG.YCR_CODEMODELE = "'+CodeMod+'" AND YMODELECATREG.YCR_CODECAT = "'+CodeCat+'" AND YMODELECATREG.YCR_CODEREG = "'+CodeReg+'"') then
  begin
  //Insertion enregistrement
   ExecuteSQL('INSERT INTO YMODELECATREG (YCR_CODEMODELE, YCR_CODECAT, YCR_CODEREG, YCR_REGGERE) VALUES ( "'
                        + CodeMod + '","'
                        + CodeCat + '","'
                        + CodeReg + '","'
                        + Gere + '") ' );
  end
  else
  begin
  //Modification enregistrement
   ExecuteSQL('UPDATE YMODELECATREG SET YMODELECATREG.YCR_REGGERE="' + Gere + '" WHERE YMODELECATREG.YCR_CODEMODELE = "'+CodeMod+'" AND YMODELECATREG.YCR_CODECAT = "'+CodeCat+'" AND YMODELECATREG.YCR_CODEREG = "'+CodeReg+'"');

  end;

end;

procedure TOF_YYMODELECATREG.InverseSelection(row : Integer);
begin
  // Mise à jour de la Grille
  if (FListe.Cells[ColonneCoche,row] = '-') or (FListe.Cells[ColonneCoche,row] = '')
    then FListe.Cells[ColonneCoche,row] := 'X'
    else FListe.Cells[ColonneCoche,row] := '-';
end;


procedure TOF_YYMODELECATREG.raffraichirLigne(row: Integer; cols : String);
begin
  // Rechargement des données de la TOB dans la ligne de la Grid
  TOBCourante.PutLigneGrid(FListe, row, FALSE, FALSE, cols);
end;

procedure TOF_YYMODELECATREG.DoKeyPress(Sender: TObject; var Key: Char);
begin
  // Traitement pour la colonne 'régime géré' uniquement
  if FListe.Col <> ColonneCoche then exit ;
  if Key = ' ' then InverseSelection(FListe.row) ;
  if (Key<>'-') and (Key<>'X') and (Key<>'x') then Key:=#0 ;
  if Key='x' then Key:='X' ;
  LigneModifiee := True;
  CM.Enabled := False;
  CM1.Enabled := False;
  CB_GERE.Enabled := False;
end;

procedure TOF_YYMODELECATREG.BChampsClick(Sender : TObject);
var i : integer ;
begin
  for i:=1 to FListe.RowCount do
  begin
    InverseSelection(i);
    LigneModifiee := True;
    CM.Enabled := False;
    CM1.Enabled := False;
    CB_GERE.Enabled := False;
  end ;
end;

procedure TOF_YYMODELECATREG.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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

procedure TOF_YYMODELECATREG.DoDblClick(Sender: TObject);
begin
    InverseSelection(FListe.row);
    LigneModifiee := True;
    CM.Enabled := False;
    CM1.Enabled := False;
    CB_GERE.Enabled := False;

end;


Initialization
  registerclasses ( [ TOF_YYMODELECATREG ] ) ;
{$ENDIF} 
end.

