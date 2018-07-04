{***********UNITE*************************************************
Auteur  ...... : Sylvie DE ALMEIDA
Créé le ...... : 05/03/2007
Modifié le ... : 05/03/2007
Description .. : Source TOF de la FICHE : YYTYPESTAXES ()
Mots clefs ... : TOF;YYTYPESTAXES
*****************************************************************}
Unit YYTYPESTAXES_TOF ;

Interface

Uses StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     Fe_Main,  //AGLLanceFiche
     FIche,
{$else}
     MaineAGL, //AGLLanceFiche
     eFiche,
     eMul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     Grids,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     HTB97,
     uTob,
     uTobDebug, 
     UTOF,
     Windows,
     Messages;

var
    MESS : array [1..5] of string = (
        {1}'Indice vide.',
        {2}'Impossible de supprimer ce type de taxe car il est utilisé dans le paramétrage des taxes.',
        {3}'Le code type est obligatoire. Merci de le renseigner.',
        {4}'Merci de renseigner le libellé.',
        {5}'Merci de sélectionner un qualifiant;');
{$IFDEF TAXES}
procedure YYLanceFiche_TypesTaxes;

const COLONNES = 'TYP_CODE;TYP_LIBELLE;TYP_QUALIFIANT;TYP_INDICE';  //'HMT_CODE;HMT_CODECPT;HMT_CODEPAYS;HMT_NUMTVA;HMT_CODEMODIF;HMT_MONTANT;HMT_TRIMESTRE;HMT_ANNEE';
const POS_CODE=0;
const POS_LIBELLE=1;
const POS_INDICE=3;
const POS_QUALIFIANT=2;
//const POS_PICTO=4;

Type
  TOF_YYTYPESTAXES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    public
    private
       FListe                  : THGrid ;
       BInsert                 : TToolbarButton97;
       BDelete                 : TToolbarButton97; 
       BValider                : TToolbarButton97; 
       NumDeclaration          : integer;
       LastError2              : integer;
       Modifiable              : boolean;
       Enregistrements         : TOB;
       AKeyDown                : TKeyEvent ;
       DelLastLine             : boolean ;  
       OnSave                  : boolean ;
       procedure GetCtrl;
       procedure Affiche;
       procedure FormatCol;
       function  ChargeEnregistrements : boolean;
       procedure ChargeEvenements;
       function  GetGridDetail : boolean;
       procedure FormKeyDown ( Sender : TObject ; var Key : Word ; Shift : TShiftState ) ;
       procedure FListeCellEnter(Sender: TObject; var ACol, ARow: Longint;var Cancel: Boolean) ;
       procedure FListeElipsisClick( Sender : TObject ) ;
       procedure FListeKeyPress( Sender : TObject ; var Key : Char);
       procedure FListeWheelUp(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : boolean);
       procedure FListeWheelDown(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : boolean);
       procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
       procedure FListeSetEditText( Sender : TObject ; ACol, ARow : Longint; const Value : string) ;
       procedure FListeCellExit( Sender : TObject ; var ACol, ARow : Longint ; var Cancel : Boolean);
       procedure FListeRowEnter(Sender : TObject ; Ou : Integer ; var Cancel : Boolean; Chg : Boolean);
       procedure FListeRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
       procedure FListeDblClick(Sender: TObject);
       procedure FListeDeleteRow( Row : integer ) ;
       procedure CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false );
       procedure NextRow ;
       procedure BDelLigneClick( Sender : TObject);
       procedure BValiderClick ( Sender : TObject);
       procedure BNewClick ( Sender : TObject);
       function  RowIsVide(Row : integer) : boolean;
       function  VerifDonnees(ACol,Arow : LongInt) : integer;
       function  DeleteModif(strCode : String):boolean;
       procedure InsertModif(strCode : string; strLibelle : string);

    end ;
{$ENDIF}
Implementation
{$IFDEF TAXES}
procedure YYLanceFiche_TypesTaxes;
begin
    AGLLanceFiche('YY','YYTYPTAXES','','','');
end;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de gestion de la Fiche
-----------------------------------------------------------------------}

procedure TOF_YYTYPESTAXES.OnArgument (S : String ) ;

begin
  // Initialisation variables :
  LastError2 := 0;
  DelLastLine := false;
  OnSave := false ;
  Inherited ;
  // On prend le controle de la fiche
  GetCtrl;
  // On charge les évènements
  ChargeEvenements;
  FListe.SetFocus;
end ;

procedure TOF_YYTYPESTAXES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYTYPESTAXES.OnLoad ;
begin
  Inherited ;
  // On gère l'affichage
  Affiche;
  FormatCol ;

  if (ChargeEnregistrements) then
     // On charge la TOB dans la grille
     Enregistrements.PutGridDetail(FListe,false,true,COLONNES,true);
end ;

procedure TOF_YYTYPESTAXES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYTYPESTAXES.OnNew ;
begin
  Inherited ;
  FListe.ColEditables[0] := true;
end ;

procedure TOF_YYTYPESTAXES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYTYPESTAXES.OnClose ;
begin
  Inherited ;
  //if assigned(Enregistrements) then
  //begin
  //  FreeAndNil(Enregistrements);
  //end ;
  // On gere le code LastError ici car il n'est pas prit en compte dans le BValider
  if LastError2 <> 0 then
  begin
       LastError := LastError2;
       PGIError(MESS[LastError2]);
       //LastErrorMsg := MESS[LastError2];
       LastError2 := 0;
       Exit;
  end;
  BValiderClick (nil);
  if FListe.Cells[0,1] <> '' then FListe.VidePile(False) ;
  Inherited ;
  if assigned(Enregistrements) then
  begin
    FreeAndNil(Enregistrements);
  end

end ;

procedure TOF_YYTYPESTAXES.OnCancel () ;
begin
  Inherited ;
end ;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de gestion des boutons
-----------------------------------------------------------------------}
procedure TOF_YYTYPESTAXES.BValiderClick;

begin

  Enregistrements.ClearDetail;
  OnSave := true;
  // On récupere les enregistrements
  if not(GetGridDetail) then exit;
  // On sauvegarde les Enregistrements en Base de données
  Enregistrements.UpdateDB;

end;

procedure TOF_YYTYPESTAXES.BNewClick;
begin
  // On crée une nouvelle ligne :
  FListe.ColEditables[0] := true;
  CreateRow(FListe.Row,true);
end;

procedure TOF_YYTYPESTAXES.BDelLigneClick( Sender : TObject);
var
 bDel : boolean;
begin
  // On supprime l'enregistrement de la base
  bDel := DeleteModif(FListe.Cells[POS_CODE,FListe.Row]);
  if bDel then
    FListeDeleteRow( FListe.Row ) ;
end;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de traitement de la Liste
-----------------------------------------------------------------------}
procedure TOF_YYTYPESTAXES.FormKeyDown ( Sender : TObject ; var Key : Word ; Shift : TShiftState ) ;
var
 Vide : boolean ;
 bDel : boolean;
 begin
 Vide := (Shift = []) ;
 AKeyDown(Sender,Key,Shift) ;
 case Key of
  VK_TAB     : if Vide then
               begin
                //if (Fliste.Col = POS_QUALIFIANT) then
                  //Fliste.SetFocus;
                  //Key :=0;
                if ( FListe.Row = FListe.RowCount - 1 ) and ( FListe.Col = FListe.Colcount - 1 ) and (Fliste.Col <> POS_QUALIFIANT) then
                  begin
                   key := 0 ;
                   NextRow ;
                   FListe.ColEditables[0] := true;
                   FListe.Col := POS_CODE; FListe.SetFocus ;
                  end;
               end;
  VK_DOWN   :  if ( FListe.Row = FListe.RowCount - 1 ) and ( Fliste.Col <> Fliste.ColCombo) then
                begin
                  Key := 0 ;
                  NextRow ;
                  if (Fliste.Col = POS_CODE) then
                   FListe.ColEditables[0] := true;
                end
                else
                begin
                   if ( Fliste.Col = Fliste.ColCombo) then
                      //FListe.setfocus;
                      Key :=0;
                end;

  VK_UP     : if ( FListe.Row = FListe.RowCount - 1 )  then
               begin
                if RowIsVide(FListe.Row) then
                begin
                  DelLastLine := true;
                  bDel := DeleteModif(FListe.Cells[POS_CODE,FListe.Row]);
                  if bDel then
                  begin
                    FListeDeleteRow(FListe.Row);
                    Key := VK_DOWN ;
                  end;
                end;
               end;
 VK_INSERT : if (Vide) then
               begin
                Key := 0 ;
                CreateRow(FListe.Row,true) ;
               end;
 VK_RETURN  : if (Vide) then
               begin
                if ( FListe.Col = FListe.ColCount -1 ) and ( FListe.Row = FListe.RowCount -1 ) then
                 Key := 0 ;
             end;
 VK_DELETE : if  Shift=[ssShift] then
              begin
               Key := 0 ;
               BDelLigneClick(nil) ;
              end
             else
              // Sauf si on est sur les cellules CODE on supprime le contenue de la cellule
              if Modifiable and (FListe.Col <> POS_CODE) then
               FListe.Cells[FListe.Col,FListe.Row] := '';
    end ;
end;

procedure TOF_YYTYPESTAXES.FListeDeleteRow( Row : integer ) ;
begin
  if not(Modifiable) then Exit;
  if ( (Fliste.Row <= 0) or (FListe.RowCount <= 1) ) then exit ;
  if Row = - 1 then Row := FListe.Row ;
  if FListe.RowCount = 1 then
  begin
     FListe.VidePile(False) ;
     CreateRow ;
  end
  else
  begin
     if ( Row <> 1 ) and ( FListe.Row = Row ) then // si on n'est pas sur la derniere cellule on remonte d'une ligne
        FListe.Row := FListe.Row - 1;
     FListe.Objects[0,Row] := nil ;
     FListe.DeleteRow(Row) ; // on supprime la ligne
  end;

end;

procedure TOF_YYTYPESTAXES.NextRow ;
var
 lBoCancel : boolean ;
 lBoChg    : boolean ;
begin
 lBoChg    := true ;
 lBoCancel := false ;
 FListeRowExit( nil , FListe.Row , lBoCancel, lBoChg );
 if lBoCancel then exit ;
 CreateRow ;
end;


procedure TOF_YYTYPESTAXES.CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false );
 procedure _Init ;
  begin
   FListe.ColEditables[0] := true;
   if FListe.CanFocus then FListe.SetFocus ;
   FListe.ShowEditor ;
  end;
 procedure _InsererLigne ;
  begin
   if not(Modifiable) then Exit;
   FListe.ColEditables[0] := true;
   FListe.InsertRow(ARow) ;
   FListe.Row      := ARow ;
   FListe.Refresh ;
  end;
 procedure _AjouterLigne ;
  begin
   if not(Modifiable) then Exit;
   FListe.ColEditables[0] := true;
   FListe.RowCount := FListe.RowCount + 1 ;
   FListe.Row      := FListe.RowCount - 1 ;          // on se place sur cette nouvelle ligne
   ARow            := FListe.Row ;

   Fliste.Refresh;
  end;
begin
  if RowIsVide(FListe.Row) then Exit; // On ne crée pas une nouvelle ligne si la precedente est vide.
  FListe.ColEditables[0] := true;
  if ARow = - 1 then ARow := FListe.Row ;
     if vBoInsert then
        _InsererLigne
     else
     begin
        if ( FListe.Row = FListe.RowCount - 1 )  then
           _AjouterLigne
        else
           if ( FListe.Row > 1 ) and ( FListe.Row < FListe.RowCount  ) then
              _InsererLigne ;
     end;
  _Init ;
end;

procedure TOF_YYTYPESTAXES.FListeDblClick(Sender: TObject);
begin
end;

procedure TOF_YYTYPESTAXES.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
end;

procedure TOF_YYTYPESTAXES.FListeCellEnter(Sender: TObject; var ACol, ARow: Longint;var Cancel: Boolean) ;
begin
end;

procedure TOF_YYTYPESTAXES.FListeElipsisClick( Sender : TObject ) ;
begin
end;

procedure TOF_YYTYPESTAXES.FListeWheelUp;
var
  bDel : boolean;
begin
  if ( FListe.Row = FListe.RowCount - 1 )  then
  begin
     if RowIsVide(FListe.Row) then
     begin
        DelLastLine := true;
        bDel := DeleteModif(FListe.Cells[POS_CODE,FListe.Row]);
        if bDel then
          FListeDeleteRow(FListe.Row);
     end;
  end;
end;

procedure TOF_YYTYPESTAXES.FListeWheelDown;
begin
  if ( FListe.Row = FListe.RowCount - 1 )  then
     NextRow ;
end;

procedure TOF_YYTYPESTAXES.FListeKeyPress( Sender : TObject ; var Key : Char);
begin
end;

procedure TOF_YYTYPESTAXES.FListeSetEditText( Sender : TObject ; ACol, ARow : Longint; const Value : string) ;
begin
end;

procedure TOF_YYTYPESTAXES.FListeCellExit( Sender : TObject ; var ACol, ARow : Longint ; var Cancel : Boolean);
var
  retour : integer;
begin
  if not(Modifiable) then Exit;
  // Lorsque que l'on supprime la deniere ligne on remonte d'une ligne de trop
  if DelLastLine then
  begin
     DelLastLine := false;
     FListe.Row := FListe.RowCount - 1;
  end;
  retour := VerifDonnees(ACol,ARow);
  if (FListe.Cells[POS_CODE, ARow] <> '') then FListe.ColEditables[0] := False;
  if (Fliste.Col = POS_CODE) and (FListe.Cells[POS_CODE, ARow] = '') then FListe.ColEditables[0] := True;
  if (retour <> 3) and (Fliste.Col <> POS_CODE) then FListe.ColEditables[0] := false;
  if retour <> 0 then
     PGIError(MESS[retour]);
end;

procedure TOF_YYTYPESTAXES.FListeRowEnter(Sender : TObject ; Ou : Integer ; var Cancel : Boolean; Chg : Boolean);
begin
end;

procedure TOF_YYTYPESTAXES.FListeRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  {for i:=0 to Fliste.ColCount-1 do
  begin
    for j:=1 to Fliste.RowCount -1 do
    begin
    retour := VerifDonnees(i,j);
    if retour <> 0 then
    begin
      if retour = 3 then FListe.ColEditables[0] := true;
      if retour <> 3 then FListe.ColEditables[0] := false;
      FListe.Col:=i ; FListe.Row:=j ; FListe.SetFocus;
      PGIError(TraduireMemoire(MESS[retour]), TraduireMemoire(Ecran.caption));
      Cancel := true;
      exit;
    end
    else
      Fliste.ColEditables[0] := false;
    end;

  end; }
end;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de traitement
-----------------------------------------------------------------------}

procedure TOF_YYTYPESTAXES.GetCtrl () ;
begin
 FListe         := THGrid(GetControl('FLISTE',true)) ;
 BInsert        := TToolbarButton97(GetControl('BInsert',true)) ;
 BDelete        := TToolbarButton97(GetControl('BDelete',true));
 BValider       := TToolbarButton97(GetControl('BValider',true)); 
 //BFerme         := TToolbarButton97(GetControl('BFerme',true));
end;

procedure TOF_YYTYPESTAXES.ChargeEvenements ;
begin
  AKeyDown                                                  := Ecran.OnKeyDown ;
  Ecran.OnKeyDown                                           := FormKeyDown ;
  FListe.OnCellEnter                                        := FListeCellEnter ;
  FListe.OnElipsisClick                                     := FListeElipsisClick ;
  FListe.OnKeyPress                                         := FListeKeyPress ;
  FListe.OnCellExit                                         := FListeCellExit ;
  FListe.OnRowEnter                                         := FListeRowEnter ;
  FListe.OnRowExit                                          := FListeRowExit ;
  FListe.OnDblClick                                         := FListeDblClick ;
  FListe.OnKeyDown                                          := FListeKeyDown ;
  FListe.OnMouseWheelDown                                   := FListeWheelDown;
  FListe.OnMouseWheelUp                                     := FListeWheelUp;
  BValider.OnClick                                          := BValiderClick ;
  //BFerme.OnClick                                            := BValiderClick ;
  BDelete.OnClick                                           := BDelLigneClick ;
  BInsert.OnClick                                           := BNewClick;
end;


procedure TOF_YYTYPESTAXES.Affiche;
var
   i : integer;
begin
   Modifiable := true;
   for i := 0 to FListe.ColCount - 1 do
       FListe.ColEditables[i] := true;
   FListe.Enabled := true;
   BInsert.Visible := true;
   BDelete.Visible := true;

end;

procedure TOF_YYTYPESTAXES.FormatCol;
begin
  // Code type de taxe
  // On interdit la modification du code type de taxe
  FListe.ColEditables[POS_CODE] := false;
  FListe.ColLengths[POS_CODE] := 3;

  //Libellé
  FListe.ColLengths[POS_LIBELLE] := 70;

  // Indice
  FListe.ColLengths[POS_INDICE] := 1;
  FListe.ColFormats[POS_INDICE] := '0';

  // Qualifiant du type de taxe
  // On gère la liste déroulante du qualifiant
  FListe.ColFormats[POS_QUALIFIANT] := 'CB=YYTYPESQUALTAXES||';
  FListe.ColLengths[POS_QUALIFIANT] := 3;
 
end;


function TOF_YYTYPESTAXES.ChargeEnregistrements : boolean;
var
  SQL : string;
  Q : TQuery;
begin
   if not(Assigned(Enregistrements)) then
      Enregistrements := TOB.Create('YTYPETAUX',nil,-1);
   SQL := 'SELECT YTYPETAUX.* ' +
          'FROM YTYPETAUX '; 
   Q := OpenSQL(SQL,true);
   try
   if not(Q.eof) then
   begin
      Enregistrements.LoadDetailDB('YTYPETAUX','','',Q,false);
      result := true;
   end
   else
   begin
      FListe.ColEditables[0] := true;
      result := false;
   end;
   finally
    Ferme(Q);
   end;
end;


function TOF_YYTYPESTAXES.GetGridDetail : boolean;
var
   T      : TOB;
   n,i, z, y  : integer;
begin
  result := true;
  z := Fliste.Col;
  y := Fliste.Row;
  {if (z = POS_QUALIFIANT) then
  begin
    Fliste.Col := POS_INDICE;
    Fliste.SetFocus;
    Fliste.Refresh;
    Fliste.Col := POS_QUALIFIANT;
    Fliste.SetFocus;

  end;}

  n := 1;
  while n < Fliste.RowCount do
  begin
     if RowIsVide(n) or // Pas de données renseignées
     (Trim(FListe.Cells[POS_CODE, n]) = '') then // Pas d'identifiant renseigné
     begin
        // On supprime la ligne
        if (Trim(FListe.Cells[POS_CODE, n]) <> '') then
           DeleteModif(FListe.Cells[POS_CODE, n]);
        FListeDeleteRow( n );
        if (n = 1) and (FListe.RowCount = 2) and (FListe.Cells[POS_CODE, n] = '') then
           break
        else
           continue;
     end;
     // Verification des données
     for i := 0 to FListe.ColCount - 1 do
     begin
        lastError2 := VerifDonnees(i,n);
        if lastError2 <> 0 then
        begin
          PGIError(MESS[LastError2]);
          LastError2 :=0;
          break;
        end
     end;
     Inc(n);
  end;
  // Cas donnée incorrecte
  if LastError2 <> 0 then
  begin
    result := false;
    LastError2 := 0;
    exit;
  end;
  if (Fliste.RowCount = 2) and RowIsVide(1) then
     Exit;// Il n'y a aucun enregistrement à sauvegarder.

  // Tout est OK on recupere les données.
  for n := 1 to Fliste.RowCount - 1 do
  begin
     T := TOB.Create('YTYPETAUX', enregistrements, -1);
     T.PutValue('TYP_CODE', FListe.Cells[POS_CODE, n]);
     T.PutValue('TYP_LIBELLE',FListe.Cells[POS_LIBELLE, n]);
     T.PutValue('TYP_QUALIFIANT',FListe.CellValues [POS_QUALIFIANT, n]);
     T.PutValue('TYP_INDICE', FListe.Cells[POS_INDICE, n]);
     if not ExisteSQL ('SELECT YTYPETAUX.* FROM YTYPETAUX WHERE TYP_CODE="'+FListe.Cells[POS_CODE, n]+'"') then
      InsertModif (FListe.Cells[POS_CODE, n], FListe.Cells[POS_LIBELLE, n]);
  end;

end;

function TOF_YYTYPESTAXES.RowIsVide(Row : integer) : boolean;
var
  i : integer ;
begin
   result := true;
   //for i := 2 to FListe.ColCount - 1 do
   for i := 1 to FListe.ColCount - 1 do
       if FListe.Cells[i,Row] <> '' then result := false;
end;

function TOF_YYTYPESTAXES.VerifDonnees(ACol,ARow : LongInt) : integer;
begin
  result := 0; 

  case ACol of
     POS_CODE :
        if Trim(FListe.Cells[ACol,ARow]) = '' then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 3;
        end;
     POS_LIBELLE :
        if Trim(FListe.Cells[ACol,ARow]) = '' then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 4;
        end;
     POS_INDICE :
        if Trim(FListe.Cells[ACol,ARow]) = '' then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 1;
        end
        else if FListe.Cells[ACol,ARow] < '1' then FListe.Cells[ACol,ARow] := '1'
        else if FListe.Cells[ACol,ARow] > '4' then FListe.Cells[ACol,ARow] := '4';
     POS_QUALIFIANT :
        if Trim(FListe.Cells[ACol,ARow]) = '' then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.CellValues [ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 5;
        end;
  end;
end;

function TOF_YYTYPESTAXES.DeleteModif(strCode : string) : boolean;
begin
  result := false;
  //pas de suppression possible si le code type est utilisé dans paramétrage taxes
  if not ExisteSQL ('SELECT YMODELECATTYP.* FROM YMODELECATTYP WHERE YMODELECATTYP.YCY_CODETYP="'+strCode+'" AND YMODELECATTYP.YCY_TYPGERE = "X"') then
  begin
    ExecuteSQL('DELETE FROM YTYPETAUX WHERE TYP_CODE="' + strCode + '"');
    result := true;
  end
  else
    PGIError(TraduireMemoire(MESS[2]), TraduireMemoire(Ecran.caption));
end;

procedure TOF_YYTYPESTAXES.InsertModif(strCode : string; strLibelle : string);
begin
   ExecuteSQL('INSERT INTO YTYPETAUX (TYP_CODE,TYP_LIBELLE) VALUES ("' + strCode + '","' + strLibelle + '")');
end;

Initialization
  registerclasses ( [ TOF_YYTYPESTAXES ] ) ;

{$ENDIF}
end.
