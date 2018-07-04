{$A+,B-,C-,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q+,R+,S-,T-,U-,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : HISTOMODIFTVA ()
Mots clefs ... : TOF;HISTOMODIFTVA
*****************************************************************}
Unit CPHISTOMODIFTVA_TOF ;

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
     Windows ;

procedure LanceModificationTVA(strParam : string);

const COLONNES = 'HMT_CODE;HMT_CODECPT;HMT_CODEPAYS;HMT_NUMTVA;HMT_CODEMODIF;HMT_MONTANT;HMT_TRIMESTRE;HMT_ANNEE';  //'HMT_CODE;HMT_CODECPT;HMT_CODEPAYS;HMT_NUMTVA;HMT_CODEMODIF;HMT_MONTANT;HMT_TRIMESTRE;HMT_ANNEE';
const POS_CODE=0;
const POS_CODECPT=1;
const POS_CODEPAYS=2;
const POS_NUMTVA=3;
const POS_CODEMODIF=4;
const POS_MONTANT=5;
const POS_TRIMESTRE=6;
const POS_ANNEE=7;

Type
  TOF_HISTOMODIFTVA = Class (TOF)
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
       BFerme                  : TToolbarButton97;
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
       function ChargeEnregistrements : boolean;
       procedure ChargeEvenements;
       function GetGridDetail : boolean;
       procedure FormKeyDown ( Sender : TObject ; var Key : Word ; Shift : TShiftState ) ;
       procedure FListeKeyPress( Sender : TObject ; var Key : Char);
       procedure FListeWheelUp(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : boolean);
       procedure FListeWheelDown(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : boolean);
       procedure FListeCellExit( Sender : TObject ; var ACol, ARow : Longint ; var Cancel : Boolean);
       procedure FListeDeleteRow( Row : integer ) ;
       procedure CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false );
       procedure NextRow ;
       procedure RenseignerLigne(ARow : integer);
       procedure BDelLigneClick( Sender : TObject);
       procedure BValiderClick ( Sender : TObject);
       procedure BNewClick ( Sender : TObject);
       function RowIsVide(Row : integer) : boolean;
       function VerifDonnees(ACol,Arow : LongInt) : integer;
    end ;

Implementation
uses uLibExportTVA,
     Vierge, //SDA le 21/01/2008
     Messages; //SDA le 21/01/2008

procedure LanceModificationTVA(strParam : string); // SDA le 21/01/2008
begin
    AGLLanceFiche('CP','CPHISTOMODIFTVA','','',strParam);
end;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de gestion de la Fiche
-----------------------------------------------------------------------}

procedure TOF_HISTOMODIFTVA.OnArgument (S : String ) ;
var
  x: integer;
  critere: string;
  Arg, Val: string;
begin

  // Initialisation variables :
  LastError2 := 0;
  DelLastLine := false;
  OnSave := false ;
  Inherited ;
  // Gestion des arguments
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        Arg := copy(Critere, 1, x - 1);
        Val := copy(Critere, x + 1, length(Critere));
        if Arg = 'CODECPT' then NumDeclaration := StrToInt(Val);
      end;
    end;
  until Critere = '';
  // On prend le controle de la fiche
  GetCtrl;
  // On charge les evenements
  ChargeEvenements;
end ;

procedure TOF_HISTOMODIFTVA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_HISTOMODIFTVA.OnLoad ;
begin
  Inherited ;

  // On gère l'affichage.
  Affiche;
  FormatCol ;

  // On charge la TOB
  if not(ChargeEnregistrements) then
  begin
     if Modifiable then
        // Il n'y a pas encore d'enregistrement on en renseigne un par defaut
        RenseignerLigne(1);
  end
  else
     // On charge la TOB dans la grille
     //SDA le 13/03/2007 Enregistrements.PutGridDetail(FListe,true,true,COLONNES,true);
     Enregistrements.PutGridDetail(FListe,false,true,COLONNES,true);
     //Fin SDA le 13/03/2007

end ;

procedure TOF_HISTOMODIFTVA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_HISTOMODIFTVA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_HISTOMODIFTVA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_HISTOMODIFTVA.OnClose ;
begin
  //SDA le 21/01/2008 Inherited ;
  BValiderClick (self);
  // On gere le code LastError ici car il n'est pas prit en compte dans le BValider
  if LastError2 <> 0 then
  begin
       LastError := LastError2;
       LastErrorMsg := MESS[LastError2];
       LastError2 := 0;
       Exit;
  end;

  if FListe.Cells[0,1] <> '' then FListe.VidePile(False) ;
  Inherited ;
  if assigned(Enregistrements) then
  begin
    FreeAndNil(Enregistrements);
  end;

end ;

procedure TOF_HISTOMODIFTVA.OnCancel () ;
begin
  Inherited ;
end ;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de gestion des boutons
-----------------------------------------------------------------------}


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Bouton Valider
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOMODIFTVA.BValiderClick;
begin
  if not(Modifiable) then Exit;
  Enregistrements.ClearDetail;
  OnSave := true;
  // On récupere les enregistrements
  if not(GetGridDetail) then exit;
  // On sauvegarde les Enregistrements
  if Enregistrements <> nil then //SDA le 21/01/2008
    Enregistrements.UpdateDB;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Bouton Nouveau
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOMODIFTVA.BNewClick;
begin
  if not(Modifiable) then Exit;
  // On crée une nouvelle ligne :
  CreateRow(FListe.Row,true);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Bouton Supprimer
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOMODIFTVA.BDelLigneClick( Sender : TObject);
begin
  if not(Modifiable) then Exit;
  // On supprime l'enregistrement de la base
  DeleteModif(StrToInt(FListe.Cells[POS_CODE,FListe.Row]));
  FListeDeleteRow( FListe.Row ) ;
end;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de traitement de la Liste
-----------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /
Description .. : Procédure permettant de gérer la navigation sur la Liste
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOMODIFTVA.FormKeyDown ( Sender : TObject ; var Key : Word ; Shift : TShiftState ) ;
var
  Vide : boolean ;
begin
  Vide := (Shift = []) ;
  AKeyDown(Sender,Key,Shift) ;
  case Key of
  VK_ESCAPE : if Vide then
                begin
                 Key := 0 ;
                 Ecran.Close ;
                end ;
  VK_TAB     : if Vide then
               begin
                if ( FListe.Row = FListe.RowCount - 1 ) and ( FListe.Col = FListe.Colcount - 1 ) then
                  begin
                   key := 0 ;
                   NextRow ;
                   FListe.Col := POS_CODEPAYS; FListe.SetFocus ;
                  end;
               end;
  VK_DOWN   : if ( FListe.Row = FListe.RowCount - 1 )  then
               begin
                Key := 0 ;
                NextRow ;
               end;
  VK_UP     : if ( FListe.Row = FListe.RowCount - 1 )  then
               begin
                if RowIsVide(FListe.Row) then
                begin
                  DelLastLine := true;
                  DeleteModif(StrToInt(FListe.Cells[POS_CODE,FListe.Row]));
                  FListeDeleteRow(FListe.Row);
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
                  //SDA le 13/03/2007
                 //NextRow ;
                 //else
                 // Key := VK_TAB ;
                 //end
                 //Fin SDA le 13/03/2007
             end ;
 VK_DELETE : if  Shift=[ssShift] then
              begin
               Key := 0 ;
               BDelLigneClick(nil) ;
              end
             else
              // Sauf si on est sur les cellules CODE on supprime le contenue de la cellule
              if Modifiable and (FListe.Col <> POS_CODE) and (FListe.Col <> POS_CODECPT) then
               FListe.Cells[FListe.Col,FListe.Row] := '';
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Procédure de suppression d'une ligne
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOMODIFTVA.FListeDeleteRow( Row : integer ) ;
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
  if not(onSave) then
  begin
     if FListe.Cells[POS_CODE,Row - 1] = '' then
        RenseignerLigne(Row - 1) // Si la nouvelle ligne courante est vide on la renseigne
     else if (FListe.RowCount = 2) and (FListe.Cells[POS_CODE,1] = '') then
        RenseignerLigne(Row); // Si la premiere ligne est vide (nouvelle declaration) on renseigne la premiere ligne
  end;
  FListe.SetFocus ;
end;

procedure TOF_HISTOMODIFTVA.NextRow ;
begin
  if not(Modifiable) then Exit;
  CreateRow ;
end;
procedure TOF_HISTOMODIFTVA.RenseignerLigne(ARow : integer);
  var
   Id : integer;
  begin
   Id := GetCodeMax('HISTOMODIFTVA')+1;
   // On gere le code de la modification
   FListe.Cells[POS_CODE,ARow] := IntToStr(Id);
   // On force le code déclaration pour que la modification soit bien liée à la bonne déclaration.
   FListe.Cells[POS_CODECPT,ARow] := IntToStr(NumDeclaration);
   // On met a jour le BDD
   InsertModif(Id,NumDeclaration);
end;

procedure TOF_HISTOMODIFTVA.CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false );
 procedure _Init ;
  begin
   //if FListe.CanFocus then
   //begin
   //   FListe.Col := POS_CODEPAYS;
   //   FListe.SetFocus ;
   //end;
   if FListe.CanFocus then FListe.SetFocus ;
   FListe.ShowEditor ;
  end;
 procedure _InsererLigne ;
  begin
   if not(Modifiable) then Exit;
   FListe.InsertRow(ARow) ;
   FListe.Row      := ARow ;
   RenseignerLigne(ARow);
   FListe.Refresh ;
  end;
 procedure _AjouterLigne ;
  begin
   if not(Modifiable) then Exit;
   FListe.RowCount := FListe.RowCount + 1 ;
   FListe.Row      := FListe.RowCount - 1 ;          // on se place sur cette nouvelle ligne
   ARow            := FListe.Row ;
   RenseignerLigne(ARow);
  end;
begin
  if RowIsVide(FListe.Row) then Exit; // On ne crée pas une nouvelle ligne si la precedente est vide.
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

procedure TOF_HISTOMODIFTVA.FListeWheelUp;
begin
  if ( FListe.Row = FListe.RowCount - 1 )  then
  begin
     if RowIsVide(FListe.Row) then
     begin
        DelLastLine := true;
        DeleteModif(StrToInt(FListe.Cells[POS_CODE,FListe.Row]));
        FListeDeleteRow(FListe.Row);
     end;
  end;
end;

procedure TOF_HISTOMODIFTVA.FListeWheelDown;
begin
  if ( FListe.Row = FListe.RowCount - 1 )  then
     NextRow ;
end;

procedure TOF_HISTOMODIFTVA.FListeKeyPress( Sender : TObject ; var Key : Char);
begin
  if not(Modifiable) then Exit;
  // On autorise que des chiffres ou le backspace pour la colonne NumTVA
  if FListe.Col = POS_NUMTVA then
     if not (Key in ['0'..'9', #8]) then Key := #0;
end;

procedure TOF_HISTOMODIFTVA.FListeCellExit( Sender : TObject ; var ACol, ARow : Longint ; var Cancel : Boolean);
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
  if retour <> 0 then
     PGIError(MESS[retour]);
end;

{-----------------------------------------------------------------------
  FONCTIONS/PROCEDURE de traitement
-----------------------------------------------------------------------}

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /
Description .. : Procédure permettant de récuperer le controle de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOMODIFTVA.GetCtrl () ;
begin
 FListe         := THGrid(GetControl('FLISTE',true)) ;
 Fliste.VidePile(False); //SDA le 21/01/2008
 BInsert        := TToolbarButton97(GetControl('BInsert',true)) ;
 BDelete        := TToolbarButton97(GetControl('BDelete',true));
 BValider       := TToolbarButton97(GetControl('BValider',true)); 
 BFerme         := TToolbarButton97(GetControl('BFerme',true));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Procedure d'assignation des evenements
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOMODIFTVA.ChargeEvenements ;
begin
  AKeyDown                                                  := Ecran.OnKeyDown ;
  Ecran.OnKeyDown                                           := FormKeyDown ;
  FListe.OnKeyPress                                         := FListeKeyPress ;
  FListe.OnCellExit                                         := FListeCellExit ;
  FListe.OnMouseWheelDown                                   := FListeWheelDown;
  FListe.OnMouseWheelUp                                     := FListeWheelUp;
  BValider.OnClick                                          := BValiderClick ;
  BFerme.OnClick                                            := BValiderClick ;
  BDelete.OnClick                                           := BDelLigneClick ;
  BInsert.OnClick                                           := BNewClick;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/01/2007
Modifié le ... :   /  /
Description .. : Procedure permettant de gérer l'affichage de la fiche en 
Suite ........ : fonction de la déclaration. Si la déclaration à été validée on 
Suite ........ : ne doit plus rien modifer aux corrections.
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOMODIFTVA.Affiche;
var
   i : integer;
begin
   Modifiable := IsModifEditable(NumDeclaration);
   if Modifiable then
   begin
     // La déclaration n'est pas encore validée.
     for i := 0 to FListe.ColCount - 1 do
         FListe.ColEditables[i] := true;
     FListe.Enabled := true;
     BInsert.Visible := true;
     BDelete.Visible := true;
   end
   else
   begin
     // La déclaration a été validée.
     //SDA le 21/01/2008 TFFiche(ecran).TypeAction := taConsult;
     TFVierge(ecran).TypeAction := taConsult; //SDa le 21/01/2008
     for i := 0 to FListe.ColCount - 1 do
         FListe.ColEditables[i] := false;
     BInsert.Visible := false;
     BDelete.Visible := false;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de formater les colonnes
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOMODIFTVA.FormatCol;
begin
  // Code Modification
  // On interdit la modification du code modif
  FListe.ColEditables[POS_CODE] := false;
  FListe.ColAligns[POS_CODE] := taRightJustify;

  // Code Déclaration
  // On desactive l'affichage du Code Déclaration
  FListe.ColLengths[POS_CODECPT] := -1;
  FListe.ColWidths[POS_CODECPT] := -1;

  // Code Pays
  // On gere l'ellipsis du code pays
  FListe.ColFormats[POS_CODEPAYS] := 'CB=TTPAYSISO2||';
  FListe.ColLengths[POS_CODEPAYS] := 2;

  // Numero TVA
  FListe.ColLengths[POS_NUMTVA] := 10;
  FListe.ColFormats[POS_NUMTVA] := '9999999999';

  // Code Motif
  FListe.ColLengths[POS_CODEMODIF] := 1;
  FListe.ColFormats[POS_CODEMODIF] := 'a';

  // Montant 
  FListe.ColTypes[POS_MONTANT] := 'R';
  FListe.ColLengths[POS_MONTANT] := 13;

  // Trimestre
  FListe.ColLengths[POS_TRIMESTRE] := 1;
  FListe.ColFormats[POS_TRIMESTRE] := '0';

  // Année
  FListe.ColLengths[POS_ANNEE] := 4;
  FListe.ColFormats[POS_ANNEE] := '0000';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 25/01/2007
Modifié le ... : 29/01/2007
Description .. : Fonction permettant de charger dans une TOB les 
Suite ........ : modification correspondant à la déclaration. 
Suite ........ : Retourne True s'il existe des enregistrements
Suite ........ : False sinon
Mots clefs ... : 
*****************************************************************}
function TOF_HISTOMODIFTVA.ChargeEnregistrements : boolean;
var
  SQL : string;
  Q : TQuery;
begin
   if not(Assigned(Enregistrements)) then
      Enregistrements := TOB.Create('HISTOMODIFTVA',nil,-1);
   try // SDA le 21/01/2008
   SQL := 'SELECT * ' +
          'FROM HISTOMODIFTVA ' +
          'WHERE HMT_CODECPT="' + IntToStr(NumDeclaration) + '"';
   Q := OpenSQL(SQL,true);
   if not(Q.eof) then
   begin
      Enregistrements.LoadDetailDB('HISTOMODIFTVA','','',Q,false);
      result := true;
   end
   else
      result := false;
   finally
    Ferme(Q);
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Fonction permettant de récuperer les valeurs contenues
Suite ........ : dans la grille retourne true si tout c'est bien passé
Suite ........ : false sinon.
Mots clefs ... : 
*****************************************************************}
function TOF_HISTOMODIFTVA.GetGridDetail : boolean;
var
   T      : TOB;
   n,i    : integer;
begin
  result := true;
  n := 1;
  while n < Fliste.RowCount do
  begin
     if RowIsVide(n) or // Pas de données renseignées
     (Trim(FListe.Cells[POS_CODE, n]) = '') or // Pas d'identifiant renseigné
     (Trim(FListe.Cells[POS_CODECPT, n]) = '') then // Pas de déclaration associée
     begin
        // On supprime la ligne
        if (Trim(FListe.Cells[POS_CODE, n]) <> '') then
           DeleteModif(StrToInt(FListe.Cells[POS_CODE, n]));
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
           break;
     end;
     Inc(n);
  end;
  // Cas donnée incorrecte
  if LastError2 <> 0 then
  begin
    result := false;
    exit;
  end;
  if (Fliste.RowCount = 2) and RowIsVide(1) then
     Exit;// Il n'y a aucun enregistrement à sauvegarder.

  // Tout est OK on recupere les données.
  for n := 1 to Fliste.RowCount - 1 do
  begin
     T := TOB.Create('HISTOMODIFTVA', enregistrements, -1);
     T.PutValue('HMT_CODE', FListe.Cells[POS_CODE, n]);
     T.PutValue('HMT_CODECPT',FListe.Cells[POS_CODECPT, n]);
     T.PutValue('HMT_CODEPAYS',FListe.CellValues [POS_CODEPAYS, n]);
     T.PutValue('HMT_NUMTVA', FListe.Cells[POS_NUMTVA, n]);
     T.PutValue('HMT_CODEMODIF', FListe.Cells[POS_CODEMODIF, n]);
     if Trim(FListe.Cells[POS_MONTANT, n]) <> '' then
        T.PutValue('HMT_MONTANT', FListe.Cells[POS_MONTANT, n])
     else
        T.PutValue('HMT_MONTANT', '0'); // On force à Zéro car c'est un double qui peut etre vide.
     T.PutValue('HMT_TRIMESTRE', FListe.Cells[POS_TRIMESTRE, n]);
     T.PutValue('HMT_ANNEE', FListe.Cells[POS_ANNEE, n])
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Fonction qui retourne vrai si tous les champs de la ligne sont 
Suite ........ : vide (execpté les champs CODE et CODECPT qui sont
Suite ........ : toujours renseignés)
Mots clefs ... : 
*****************************************************************}
function TOF_HISTOMODIFTVA.RowIsVide(Row : integer) : boolean;
var
  i : integer ;
begin
   result := true;
   for i := 2 to FListe.ColCount - 1 do
       if FListe.Cells[i,Row] <> '' then result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 30/01/2007
Modifié le ... : 30/01/2007
Description .. : Fonction qui permet de verifier pour une cellule si la valeur
Suite ........ : est correcte. Elle renvoie 0 si c'est OK ou alors un numero
Suite ........ : d'erreur correspond au tableau de Messages MESS 
Suite ........ : (uLibExportTVA)
Mots clefs ... : 
*****************************************************************}
function TOF_HISTOMODIFTVA.VerifDonnees(ACol,ARow : LongInt) : integer;
var
  montant : double;
begin
  result := 0; 

  case ACol of
     POS_MONTANT :
     begin
        if (Trim(FListe.Cells[ACol,ARow]) = '') then
        begin
           FListe.Cells[ACol,ARow] := '0';
        end;
        // On supprime le caractère séparateur qui n'est pas un espace mais le caractère A0
        montant := StrToFloat(DeleteChar(FListe.Cells[ACol,ARow],#160));
        if (montant > 9999999999) then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 2;
        end
        else if (montant < -9999999999) then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 3;
        end
        else
          FListe.Cells[ACol,ARow] := FloatToStr(ArrondirApresVirgule(montant,2));
     end;
     POS_NUMTVA :
        if Trim(FListe.Cells[ACol,ARow]) = '' then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 5;
        end;
     POS_TRIMESTRE :
        if Trim(FListe.Cells[ACol,ARow]) = '' then
        begin
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 6;
        end
        else if FListe.Cells[ACol,ARow] < '1' then FListe.Cells[ACol,ARow] := '1'
        else if FListe.Cells[ACol,ARow] > '4' then FListe.Cells[ACol,ARow] := '4';
     POS_ANNEE :
        if (Trim(FListe.Cells[ACol,ARow]) = '') then
        begin 
           FListe.Cells[ACol,ARow] := '';
           FListe.Row := ARow;
           FListe.Col := ACol;
           result := 7;
        end
        //SDA le 13/03/2007
        //else if FListe.Cells[ACol,ARow] < IntToStr(GetMinAnneeExercice) then FListe.Cells[ACol,ARow] := IntToStr(GetMinAnneeExercice)
        //else if FListe.Cells[ACol,ARow] > IntToStr(GetMaxAnneeExercice) then FListe.Cells[ACol,ARow] := IntToStr(GetMaxAnneeExercice);
        else if strToInt(FListe.Cells[ACol,ARow]) < GetMinAnneeExercice then FListe.Cells[ACol,ARow] := IntToStr(GetMinAnneeExercice)
        else if strToInt(FListe.Cells[ACol,ARow]) > GetMaxAnneeExercice then FListe.Cells[ACol,ARow] := IntToStr(GetMaxAnneeExercice);
        //Fin SDA le 13/03/2007
  end;
end;

Initialization
  registerclasses ( [ TOF_HISTOMODIFTVA ] ) ;
end.
