unit ULibWindows;

interface

uses
 hCtrls,     // pour le THGrid
 {$IFDEF EAGLCLIENT}
 
 {$ELSE}
 HDB,
 DB,         // pour le TFielType
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
 {$ENDIF}
 SysUtils,   // pour DecimalSeparator , FileExists
 {$IFDEF VER150}
 variants,
 {$ENDIF}
 Windows,
 Graphics,
 Classes,    // TComponent
 ComCtrls,   // TTabSheet
 Filtre,     // VideCrit
 Hent1,      // StrfMontant
 HMsgBox,    // PgiBox
 Forms,
 Ed_Tools,   // pour le vide list
 extctrls,
 UTOB,       // pour le TFieldType
 Grids,      // TGridDrawState
 Menus,      // TPopUoMenu
 Paramsoc,   //GetParamSocSecur
 Controls;   // TControl



const

 RC_GAUCHE_DROITE = 1 ;
 RC_DROITE_GAUCHE = -1 ;

// Anciennement TFieldType
type

 PFieldsCumul = ^FieldsCumul;
 FieldsCumul = record
  inNum       : integer;
  InColGrid   : integer;
  StFieldName : string ;
  RdTotal     : double;
  StMask      : string ;
 end;

 TMulPanelCumul = Class
  private
   FTobListe           : TOB ;
   FPanelCumul         : TPanel ;
   FPanelCumulSelect   : TPanel ;

   FFieldsCumul        : TList ;
   FFieldsCumulSelect  : TList ;

   FGrid               : THGrid ;
   FStListeChamps      : string;
   FNewSources         : string;
   FNewLiaison         : string;
   FNewTris            : string;
   FNewParams          : string;
   FNewTitres          : HString;
   FNewLargeurs        : string;
   FNewJustifs         : string;
   FNewLibelle         : HString;
   FNewNumCols         : HString;
   FNewPerso           : string;
   FNewOkTri           : boolean;
   FNewOkNumCol        : boolean;
{$IFDEF EAGLCLIENT}
   FboFetchLesTous     : boolean ;
   FStSelectSUM        : string;  // Select SUM( )......
{$ENDIF}
   FStWhereSql         : string;  // = '' si pas d'enregistrement, sinon contient
                                  // la requête qui a été faite
   procedure Update2_3Cumul ;
   {$IFDEF EAGLCLIENT}
   procedure UpdateECumul ;
   {$ENDIF}
   
  public
   constructor Create ;
   destructor  Destroy ; override ;
   procedure   InitializeField ;
   procedure   InitializeControl ;
   procedure   UpdateCumul ;
   procedure   UpdateCumulSelect;

   procedure   UpdateLayoutCumul ;
   procedure   UpdateLayoutCumulSelect ;
   function    GetCumulChamps ( vStChamps : string ) : double ;

   property PanelCumul  : TPanel read FPanelCumul  write FPanelCumul ;
   property PanelCumulSelect : TPanel read FPanelCumulSelect write FPanelCumulSelect;

   property Grid        : THGrid read FGrid            write FGrid ;
   property TobListe    : TOB    read FTobListe        write FTobListe ;
   property StWhereSql  : string read FStwhereSql      write FStWhereSql ;
   {$IFDEF EAGLCLIENT}
   property FetchLesTous: boolean read FboFetchLesTous write FboFetchLesTous ;
   {$ENDIF}
 end;

 TNextControlForm = Class(TForm)
  public
   function FindNextControl ( CurControl: TWinControl; GoForward, CheckTabStop, CheckParent : Boolean ) : TWinControl ;
  end ;


// Grise la case Acol ARow de G
procedure SetGridGrise(ACol, ARow : Integer; G : THGrid);

procedure CSetGridSep(ACol, ARow : Integer ; G : THGrid ; Canvas : TCanvas ; haut : boolean ) ;

procedure SetGridGriseDBGrid(GCan : TCanvas ; const R: TRect ; GColor : TColor);

// Passe à la colonne suivante de la grille G
procedure PasseColSuivante( var Acol, ARow : integer; G : THGrid );

// Equivalent en language C de ( A > B ? C : D)
function IIF( Valeur: boolean; Premier, Dernier: Variant ) : Variant;

// Teste la touche saisie en fonction du type que l on saisie
function TestKeyPressed(Thetype : TFieldType; Key : Char) : Char;

// Retourne une chaine avec la notion Db ou CR
function AfficheDBCR(vValeur : Double; vStMask : string = ''): string;

// Retourne le chemin d' accès du répertoire temporaire spécifié dans Windows
function GetWindowsTempPath : String ;

// Cherche le ficher lNomFichier et le supprime si il existe
function FileExistsDelete( vNomFichier : String ) : Boolean;

function DecodeKeyNum ( var vBoShift : boolean ; var Key : Word ) : integer;

// reourne le sens de deplacement dans une grille 1 de gauche à droite  -1 de droite à gauche
function CGetGridSens(G : THGrid ; ACol, ARow : Integer) : Integer;

// desactive l'affichage de la grille
procedure HGBeginUpdate( G : THGrid );

// active l'affichage de la grille
procedure HGEndUpdate( G : THGrid );

// Vide les composants d'un onglet passé en paramètre
procedure VideControlOnglet( vOnglet : TTabSheet );

// Verifie si le composant existe dans la fiche
procedure NotifyErrorComponent( Sender : TComponent; Name: string );

// Desactive tous les controls contenu dans TheControl
procedure CDisableControl( TheControl : TCustomControl );

// active tous les controls contenu dans TheControl
procedure CEnableControl( TheControl : TCustomControl );

// Recherche l'Edit actif et ouvre le lookup associé
procedure CVK_F5OnEdit ( vEcran : TForm );

procedure CDessineTriangle( G : THGrid ; vColor : TColor = clRed) ;

// Traduit le Value d'un THMultiValComboBox et retourne la traduction SQL ainsi
// que la traduction du Libelle
procedure TraductionTHMultiValComboBox( vControl : THMultiValComboBox ; var vStSql : string ; var vStLibelle : string ; vStNomChamp : string ; vVideSiTous : Boolean = True; vStAjouteValue : string = '');

// onction de completion des ocntrole d'un mul ou d'un grule
procedure CVerifCritere ( Ecran : TForm ) ;

// Charge un THValComboBox avec les champs de la liste paramètrables
procedure CChargeZ_C( vControl : THValComboBox; vStListeChamps : string ; vStNewTitres : string );

// Supprime le couple (value / item) de la liste du THValComboBox à partir d'une value
procedure CVireLigne( vControl : THValComboBox; vStCode : String ) ;

// Pour aller sur le bon controle avec des gestion des enabled / disabled
procedure GereNextControl( ff : TForm ; Ctrldepartie : String ) ; //XMG 04/06/03

// Gestion des THMultIValComboBox en ReadOnly
procedure CSelectionTextControl( Sender : TObject ) ;

procedure CUpdateGridListe ( FListe : THGrid ; FStListeParam : string ) ;

function TestJoker( ATester : string) : Boolean;
function  ConvertitCaractereJokers(ZoneDe, ZoneA : THEdit; Chp : string) : string;
function ClauseAvecJoker(ATraiter, Chp : string): string;

{JP 05/10/05 : les valeurs forcées sur les combo d'un PageControl de critères sont écrasées
               par celles mémorisées dans la tob du filtre après le FormShow. Cette méthode
               permet de mémoriser dans le filtre la valeur que l'on veut forcer.
               Mem    : TFQRS1(Ecran).ListeFiltre.Current
               Combo  : Nom de la Combo ('NATURE', 'E_ETABLISSEMENT' ...)
               Valeur : Valeur que l'on veut forcer dans la combo, indépendamment de celle
                        mémorisée dans le filtre}
procedure ForceValeurSurFiltre(var Mem : TOB; Combo, Valeur : string);
{JP 05/10/05 : Renvoie la valeur d'un composant qui est stockée dans un filtre
               Mem   : TFQRS1/TFMul(Ecran).ListeFiltre.Current
               Zone : Nom de la Combo ('NATURE', 'E_ETABLISSEMENT' ...)}
function  GetValeurDuFiltre(Mem : TOB; Zone : string) : variant;

function TraduitOperateur(vOperateurValue, vValue: variant; vBoTypeChaine : Boolean) : string;

{JP 08/03/06 : Compare deux chaînes dont l'une (ATester) peut contenir des caractères Joker.
               Renvoie True si les deux chaînes sont identiques ou si les caractères Joker
               de ATester sont en cohérence avec Les caractères de AComparer}
function CompareAvecJoker(ATester, AComparer : string) : Boolean;

function FullMajuscule(S: string): string;
{JP 05/06/07 : lorsque l'on veut cacher une colonne dans une grille, il faut mettre ColWidths à -1 ou 0 selon l'OS,
               car avec -1 sous Vista et XP, la colonne précédente est un peu "mangée"}
function GetWidthColFromOs : Integer;

function Rotate90( vBitMap : TBitMap ) : TBitMap;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 05/09/2001
Modifié le ... : 05/09/2001
Description .. : équivalent en language C de ( A > B ? C : D)
Mots clefs ... : IF
*****************************************************************}
function IIF(Valeur: boolean; Premier, Dernier: Variant): Variant;
begin
  if Valeur then
    Result := Premier
  else
    Result := Dernier;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/10/2001
Modifié le ... :   /  /
Description .. : Test la touche saisie et verifie sa validité en fonction
                 du type que l' on veut tester ( Integer, Double, ect.... )
Mots clefs ... : Key ; KeyPress
*****************************************************************}
function TestKeyPressed(Thetype : TFieldType ; Key : Char) : Char;
begin

  Result := #0;

  case TheType of
   ftDateTime : if (Key in ['0'..'9','/',#8]) then
                 Result := Key;
   ftInteger  : if (Key in ['0'..'9',#8]) then
                 Result := Key;
   ftFloat    : if Key in ['.',','] then
                 Result := DecimalSeparator
                  else
                   if (Key in ['0'..'9',#8]) then
                    Result := Key;
  end; // case

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Passe a la colonne suivante
Mots clefs ... :
*****************************************************************}
procedure PasseColSuivante( var Acol, ARow : integer; G : THGrid );
var
 lInSens : integer ;
 lBoFin  : boolean ;
 lInCol : integer;
begin

 lBoFin := false ;
 lInCol := G.FixedCols ;

 if G.Col = 1 then
  lInsens := 0  // on est dans la premiere colonne -> on ne bouge pas
   else
    if ( G.Row = ARow) then
     begin // on est dans la même ligne

      if ACol = ( G.ColCount - 2 ) then
       begin
        lInSens := 0 ;// on est en fin de tableau
        lBoFin  := true ;
       end
        else
         if (G.Col > ACol) then
          lInSens := 1 + G.Col - ACol // deplacement de gauche à droite
           else
            lInSens := - 1 - ACol + G.Col; // deplacement de droite à gauche
     end
      else
       begin
        if (G.Col = ACol) then
         lInSens := 1
          else
           lInSens := 0;
       end;

  ACol   := ACol + lInSens ;
  if ACol >= G.ColCount then  // on sort en se deplacant vers le bas
   ACol := ACol - 2 ;

  if lBoFin then
   begin
    if ( ARow + 1 ) < G.RowCount then
     ARow := ARow + 1 ;
    ACol := lInCol ;
   end
    else
     ARow   := G.Row ;

end;

procedure CSetGridSep( ACol, ARow : Integer ; G : THGrid ; Canvas : TCanvas ; haut : boolean) ;
var
 R          : TRect ;
 lOldBrush  : TBrush ;
 lOldPen    : TPen ;
begin

 lOldBrush           := TBrush.Create ;
 lOldPen             := TPen.Create ;
 // sauvegarde des valeurs courantes
 lOldBrush.assign(Canvas.Brush) ;
 lOldPen.assign(Canvas.Pen) ;

 try

 // G.Canvas.Brush.Color := clRed ;
  Canvas.Brush.Style := bsSolid ;
  Canvas.Pen.Color   := clRed ;
  Canvas.Pen.Mode    := pmCopy ;
  Canvas.Pen.Style   := psSolid ;
  Canvas.Pen.Width   := 1 ;
  R                  := G.CellRect(ACol, ARow) ;
  if haut then
   begin
    Canvas.MoveTo(R.Left, R.Top) ;
    Canvas.LineTo(R.Right+1, R.Top) ;
   end
    else
     begin
      Canvas.MoveTo(R.Left, R.Bottom-1) ;
      Canvas.LineTo(R.Right+1, R.Bottom-1) ;
     end ;

 finally
  // réaffectation des valeurs du canevas
  Canvas.Brush.Assign(lOldBrush);
  Canvas.Pen.Assign(lOldPen);
  if assigned(lOldBrush)      then lOldBrush.Free;
  if assigned(lOldPen)        then lOldPen.Free;
 end; // try

end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 21/06/2001
Modifié le ... : 20/07/2001
Description .. : Grise la case debit si credit est renseigné et inversement
Mots clefs ... :
*****************************************************************}
procedure SetGridGrise(ACol, ARow : Integer; G : THGrid);
var
 R          : TRect ;
 lOldBrush   : TBrush ;
 lOldPen     : TPen ;
begin

 lOldBrush           := TBrush.Create ;
 lOldPen             := TPen.Create ;
 // sauvegarde des valeurs courantes
 lOldBrush.assign(G.Canvas.Brush) ;
 lOldPen.assign(G.Canvas.Pen) ;

 try

  G.Canvas.Brush.Color := G.FixedColor ;
  G.Canvas.Brush.Style := bsBDiagonal ;
  G.Canvas.Pen.Color   := G.FixedColor ;
  G.Canvas.Pen.Mode    := pmCopy ;
  G.Canvas.Pen.Style   := psClear ;
  G.Canvas.Pen.Width   := 1 ;
  R                    := G.CellRect(ACol,ARow);
  G.Canvas.Rectangle(R.Left, R.Top, R.Right+1, R.Bottom+1) ;

 finally
  // réaffectation des valeurs du canevas
  G.Canvas.Brush.Assign(lOldBrush);
  G.Canvas.Pen.Assign(lOldPen);
  if assigned(lOldBrush)      then lOldBrush.Free;
  if assigned(lOldPen)        then lOldPen.Free;
 end; // try

end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 21/06/2001
Modifié le ... : 20/07/2001
Description .. : Grise la case debit si credit est renseigné et inversement
Mots clefs ... :
*****************************************************************}
procedure CDessineTriangle( G : THGrid ; vColor : TColor = clRed) ;
var
 R           : TRect ;
 T1,T2,T3    : TPoint ;
 lOldBrush   : TBrush ;
 lOldPen     : TPen ;
begin

 lOldBrush           := TBrush.Create ;
 lOldPen             := TPen.Create ;
 // sauvegarde des valeurs courantes
 lOldBrush.assign(G.Canvas.Brush) ;
 lOldPen.assign(G.Canvas.Pen) ;

 try

   R:=G.CellRect(G.Col,G.Row) ;
   G.Canvas.Brush.Color   := vColor ;
   G.Canvas.Brush.Style   := bsSolid ;
   G.Canvas.Pen.Color     := vColor ;
   G.Canvas.Pen.Mode      := pmCopy ;
   G.Canvas.Pen.Width     := 1 ;
   T1.X                   := R.Right - 5 ;
   T1.Y                   := R.Top + 1 ;
   T2.X                   := T1.X + 4 ;
   T2.Y                   := T1.Y ;
   T3.X                   := T2.X ;
   T3.Y                   := T2.Y + 4 ;
   G.Canvas.Polygon([T1,T2,T3]) ;

 finally
  // réaffectation des valeurs du canevas
  G.Canvas.Brush.Assign(lOldBrush);
  G.Canvas.Pen.Assign(lOldPen);
  if assigned(lOldBrush)      then lOldBrush.Free;
  if assigned(lOldPen)        then lOldPen.Free;
 end; // try

end ;

procedure SetGridGriseDBGrid(GCan : TCanvas ; const R: TRect ; GColor : TColor);
var
 lOldBrush   : TBrush ;
 lOldPen     : TPen ;
begin
 lOldBrush           := TBrush.Create ;
 lOldPen             := TPen.Create ;
 // sauvegarde des valeurs courantes
 lOldBrush.assign(GCan.Brush) ;
 lOldPen.assign(GCan.Pen) ;

 try
  GCan.FillRect(R);
  GCan.Brush.Color := GColor ;
  GCan.Brush.Style := bsBDiagonal ;
  GCan.Pen.Color   := GColor ;
  GCan.Pen.Mode    := pmCopy ;
  GCan.Pen.Style   := psClear ;
  GCan.Pen.Width   := 1 ;

  GCan.Rectangle(R.Left, R.Top, R.Right+1, R.Bottom+1) ;

 finally
  // réaffectation des valeurs du canevas
  GCan.Brush.Assign(lOldBrush);
  GCan.Pen.Assign(lOldPen);
  if assigned(lOldBrush)      then lOldBrush.Free;
  if assigned(lOldPen)        then lOldPen.Free;
 end; // try
end ;

{***********A.G.L.***********************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 05/09/2001
Modifié le ... : 05/09/2001
Description .. : Affiche un montant en valeur absolue avec la notion DB ou CR
Mots clefs ... : DB CR VALEUR ABSOLUE
*****************************************************************}
function AfficheDBCR(vValeur : Double; vStMask : string = ''): string;
var lStMask : string;
    i : integer;
begin
  if vStMask <> '' then
    lStMask := vStMask
  else
  begin
    lStMask := '#,##0.';
    for i := 0 to V_Pgi.OkDecV - 1 do
      lStMask := lStMask + '0';
  end;

  if vValeur > 0 then
    Result := FormatFloat(lStMask, vValeur) + ' D'
  else
    Result := FormatFloat(lStMask, Abs(vValeur)) + ' C';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2002
Modifié le ... : 14/02/2002
Description .. : Retourne le chemin d' accès du répertoire temporaire spécifié dans Windows
Mots clefs ... : TEMP , WINDOWS
*****************************************************************}
function GetWindowsTempPath : String;
var lBuffer: Array[0..1023] of Char;
    lChemin : String;
begin
  GetTempPath(1023, lBuffer);
  SetString(lChemin, lBuffer, StrLen(lBuffer));
  Result := lChemin;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2002
Modifié le ... : 22/02/2002
Description .. : Supprime le fichier vNomFichier si il existe
Suite ........ : Retourne True si le fichier n' existe pas, ou si on l'a supprimé
Suite ........ : Retourne False si erreur de suppression
Mots clefs ... : FILEEXISTS DELETEFILE
*****************************************************************}
function FileExistsDelete( vNomFichier : String ) : Boolean;
begin
  Result := True;
  if FileExists( vNomFichier ) then
    if not DeleteFile( PChar(vNomFichier) ) then Result := False;
end;


function DecodeKeyNum ( var vBoShift : boolean ; var Key : Word ) : integer;
begin

 result := -1;

 if not vBoShift then exit;

 case Key of
  VK_INSERT   : result := 0 ;
  VK_END      : result := 1;
  VK_DOWN     : result := 2 ;
  VK_NEXT     : result := 3 ;
  VK_LEFT     : result := 4 ;
  VK_CLEAR    : result := 5 ;
  VK_RIGHT    : result := 6 ;
  VK_HOME     : result := 7 ;
  VK_UP       : result := 8 ;
  VK_PRIOR    : result := 9 ;
 end; // case

 if result <> - 1 then begin Key := 0 ; vBoShift:=false ; end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Lauretn GENDREAU
Créé le ...... : 12/07/2002
Modifié le ... :   /  /
Description .. : retourne le sens de deplacement dans une grille
Suite ........ :  1 : de gauche à droite
Suite ........ : -1 : de droite à gauche
Mots clefs ... :
*****************************************************************}
function CGetGridSens(G : THGrid ; ACol, ARow : Integer) : Integer;
begin
// Sens de déplacement dans le Grid
if ( G.Row = ARow ) then
 begin
  if ( G.Col >= ACol ) then
   Result := 1
    else
     Result := -1;
 end
  else
   if ( G.Row > ARow ) then
    Result := 1
     else
      Result:=-1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. : desactive l'affichage de la grille
Mots clefs ... :
*****************************************************************}
procedure HGBeginUpdate( G : THGrid );
begin
 G.SynEnabled:=false ; G.BeginUpdate ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. : Active l'affichage de la grille
Mots clefs ... :
*****************************************************************}
procedure HGEndUpdate( G : THGrid );
begin
 G.SynEnabled:=true ; G.EndUpdate ;
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2002
Modifié le ... :   /  /    
Description .. : Vide les composants d'un onglet passé en paramètre
Mots clefs ... :
*****************************************************************}
procedure VideControlOnglet( vOnglet : TTabSheet );
var i : integer;
begin
  for i := 0 to vOnglet.ControlCount - 1 do
  begin
    VideCrit( vOnglet.Controls[i] );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 26/09/2002
Modifié le ... :   /  /
Description .. : Verifie si le composant existe dans la fiche
Mots clefs ... :
*****************************************************************}
procedure NotifyErrorComponent(Sender : TComponent; Name: string);
var lStMessage : string;
begin
  if not assigned(Sender) then
  begin
    lStMessage := TraduireMemoire('Le contrôle ') + Name + TraduireMemoire(' est manquant !' ) + #13#10 +
                  TraduireMemoire(' Vérifiez la version de votre base');
    PgiBox( lStMessage,TraduireMemoire('Erreur création de la fiche') ) ;

    // pour stopper les traitements suivants
    raise EAbort.Create('Erreur à la création');
  end; //if
end;

(*type
 TCustomControlColorPublic = class(TCustomControl)
 public
  property color ;
 end;*)

procedure CGereEnabledControl( TheControl : TCustomControl ; vBoEnabled : boolean );
var i : integer ;
begin
 TheControl.Enabled := vBoEnabled ;
 //TCustomControlColorPublic(TheControl).Color := IIF( vBoEnabled, ClWindow, ClBtnFace);
 for i:=0 to TheControl.ControlCount-1 do
 begin
   TheControl.Controls[i].Enabled := vBoEnabled ;
   //TCustomControlColorPublic(TheControl.Controls[i]).Color := IIF( vBoEnabled, ClWindow, ClBtnFace);
 end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 11/10/2002
Modifié le ... :   /  /
Description .. : Desactive tous les controls contenu dans TheControl
Mots clefs ... : 
*****************************************************************}
procedure CEnableControl( TheControl : TCustomControl );
begin
 CGereEnabledControl(TheControl,true) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 11/10/2002
Modifié le ... :   /  /    
Description .. : Active tous les controls contenu dans TheControl
Mots clefs ... : 
*****************************************************************}
procedure CDisableControl( TheControl : TCustomControl );
begin
 CGereEnabledControl(TheControl,false) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/01/2003
Modifié le ... : 15/06/2007
Description .. : Recherche l'Edit actif et ouvre le lookup associé
Suite ........ : - LG - 16/05/2007 - reecriture du code pour enlever une 
Suite ........ : fuite memoire
Mots clefs ... : 
*****************************************************************}
procedure CVK_F5OnEdit ( vEcran : TForm );
var
 i : integer ;
begin
 for i:=0 to vEcran.ComponentCount-1 do
  begin
   if (vEcran.Components[i] is THEdit) and (THEdit(vEcran.Components[i]).Focused) then
     THEdit(vEcran.Components[i]).ElipsisClick(nil) ;
  end; // for
end;


{ TMulPanelCumul }

constructor TMulPanelCumul.Create;
begin
  FFieldsCumul       := TList.Create;
  FFieldsCumulSelect := TList.Create;
end;

destructor TMulPanelCumul.Destroy;
var
 i  : integer ;
 lP : PFieldsCumul ;
begin
  for i := 0 to FFieldsCumul.Count - 1 do
  begin
   lP := FFieldsCumul.Items[i] ;
   if lP <> nil then
    Dispose(lP) ;
   FFieldsCumul.Items[i] := nil ;
  end ;
  FreeAndNil(FFieldsCumul) ;

  for i := 0 to FFieldsCumulSelect.Count - 1 do
  begin
   lP := FFieldsCumulSelect.Items[i] ;
   if lP <> nil then
    Dispose(lP) ;
   FFieldsCumulSelect.Items[i] := nil ;
  end ;
  FreeAndNil(FFieldsCumulSelect) ;

  inherited;
  
end;

procedure TMulPanelCumul.InitializeField;
var
 lStListe          : string ;
 lStAlignement     : string ;
 lStAlignements    : string ;
 lStFormat         : string ;
 lRdDec            : integer ;
 lBoSep            : boolean ;
 lBoObli           : boolean ;
 lBoLib            : boolean ;
 lBoVisu           : boolean ;
 lBoNulle          : boolean ;
 lBoCumul          : boolean ;
 lPField           : PFieldsCumul ;
 lPFieldSelect     : PFieldsCumul ;
 lRdIndex          : integer ;
 lStFieldName      : string ;
 lStFields         : string ;

 procedure VideListeRecord(L: TList);
 var i: integer;
 lP : PFieldsCumul ;
 begin

  if (FFieldsCumul = nil) or (FFieldsCumul.Count <= 0) then Exit;

   for i := 0 to FFieldsCumul.Count - 1 do
    begin
     lP := FFieldsCumul.Items[i] ;
     if lP <> nil then
      Dispose(lP) ;
     FFieldsCumul.Items[i] := nil ;
    end ;

   FFieldsCumul.Clear ;

   if (FFieldsCumulSelect = nil) or (FFieldsCumulSelect.Count <= 0) then Exit;

   for i := 0 to FFieldsCumulSelect.Count - 1 do
    begin
     lP := FFieldsCumulSelect.Items[i] ;
     if lP <> nil then
      Dispose(lP) ;
     FFieldsCumulSelect.Items[i] := nil ;
    end;
   FFieldsCumulSelect.Clear ;

 end;

begin

  if FGrid = nil then exit ;

  // GCO - 27/09/2006 - Ajout ParamSoc pour PB SIC pour empêcher les cumuls
  if GetParamSocSecur('SO_CPCONSSANSCUMUL', False) then
  begin
    if Assigned(PanelCumul) then PanelCumul.Visible := False;
    if Assigned(PanelCumulSelect) then PanelCumulSelect.Visible := False;
    Exit;
  end;

  lStListe := FGrid.ListeParam ;
  lRdIndex := 0 ;

  VideListeRecord(FFieldsCumul) ;
  VideListeRecord(FFieldsCumulSelect) ;

  ChargeHListe(lStListe,FNewSources,FNewLiaison,FNewTris,FStListeChamps,FNewTitres,FNewLargeurs,FNewJustifs,FNewParams, FNewLibelle, FNewNumCols,FNewPerso,FNewOkTri,FNewOkNumCol);
  lStAlignements := FNewJustifs ;
  lStFields      := FStListeChamps ;
{$IFDEF EAGLCLIENT}
  FStSelectSUM   := 'SELECT ';
{$ENDIF}
  while lStAlignements <> '' do
  begin
    lStAlignement  := ReadTokenSt ( lStAlignements ) ;
    lStFieldName   := ReadTokenSt ( lStFields  ) ;
    TransAlign(lStAlignement,lStFormat,lRdDec,lBoSep,lBoObli,lBoLib,lBoVisu,lBoNulle,lBoCumul) ;
{$IFDEF EAGLCLIENT}
    if lBoCumul  and ( (Pos('_', lStFieldName) > 0) or FboFetchLesTous ) then
{$ELSE}
    if lBoCumul then
{$ENDIF}
    begin
      new(lPField) ;
      FillChar(lPField^,sizeOf(lPField^),#0) ;
      lPField^.inNum       := ChampToNum(lStFieldName) ;
      lPField^.InColGrid   := lRdIndex+1 ;
      lPField^.RdTotal     := 0 ;
      lPField^.StFieldName := lStFieldName ;
      lPField^.StMask      := lStFormat;
      FFieldsCumul.Add(lPField) ;

      new(lPFieldSelect) ;
      FillChar(lPFieldSelect^,sizeOf(lPFieldSelect^),#0) ;
      lPFieldSelect^.inNum       := ChampToNum(lStFieldName) ;
      lPFieldSelect^.InColGrid   := lRdIndex+1 ;
      lPFieldSelect^.RdTotal     := 0 ;
      lPFieldSelect^.StFieldName := lStFieldName ;
      lPFieldSelect^.StMask      := lStFormat;
      FFieldsCumulSelect.Add(lPFieldSelect) ;

{$IFDEF EAGLCLIENT}
      FStSelectSum := FStSelectSum + 'SUM(' + lPField^.StFieldName + ') ' + lPField^.StFieldName + ', ';
{$ENDIF}
    end; // if
    Inc(lRdIndex) ;
  end; // while

{$IFDEF EAGLCLIENT}
  FStSelectSUM := FStSelectSUM + ' COUNT(*) TOTAL FROM ' + FNewSources;
{$ENDIF}

  if (FFieldsCumul.Count = 0) then
  begin // Panel invisible si pas de cumul
    if Assigned(PanelCumul) then PanelCumul.Visible := False;
    if Assigned(PanelCumulSelect) then PanelCumulSelect.Visible := False;
  end;
end;

procedure TMulPanelCumul.InitializeControl;
var i : integer ;
begin
  // création des NumEdit pour affichage des cumuls
  if (FFieldsCumul.Count>0) and (FPanelCumul<>nil) then
  begin
    while FPanelCumul.ControlCount > 0 do FPanelCumul.Controls[0].Free ;

    if Assigned(FPanelCumulSelect) then
    while FPanelCumulSelect.ControlCount > 0 do FPanelCumulSelect.Controls[0].Free ;

    for i := 0 to FFieldsCumul.Count - 1 do
    begin
      with THNumEdit.Create(FPanelCumul.Owner) do
      begin
       parent               := FPanelCumul ;
       parentColor          := true;
       Font.Style           := [fsBold];
       masks.PositiveMask   := PFieldsCumul( FFieldsCumul[i] )^.StMask;
       value                := 0 ;
       Top                  := 0 ;
       Ctl3D                := true ;
       ReadOnly             := true ;
       Enabled              := false ;
       Tag                  := PFieldsCumul( FFieldsCumul[i] )^.InColGrid ;
      end; // with

      if Assigned(FPanelCumulSelect) then
        with THNumEdit.Create(FPanelCumulSelect.Owner) do
        begin
          parent               := FPanelCumulSelect ;
          parentColor          := true;
          Font.Style           := [fsBold];
          masks.PositiveMask   := PFieldsCumul( FFieldsCumulSelect[i] )^.StMask;
          value                := 0 ;
          Top                  := 0 ;
          Ctl3D                := true ;
          ReadOnly             := true ;
          Enabled              := false ;
          Tag                  := PFieldsCumul( FFieldsCumulSelect[i] )^.InColGrid ;
        end; // with

    end;

  end; // if

// UpdateFormules;
// UpdateCumul;

end;

{$IFDEF EAGLCLIENT}
procedure TMulPanelCumul.UpdateECumul;
var j         : integer;
    P         : PFieldsCumul;
    lQuery    : TQuery;
    lStSql    : string;
    lNumTotal : integer;
begin

  // Mode CWAS, les cumuls sont réalisés par un SELECT SUM(
  lNumTotal := 0;
  if FStWhereSql <> '' then
  begin
    lStSql := FStSelectSUM + ' ' + FStWhereSql ;
    lQuery := nil;
    try
      try
        lQuery := OpenSql( lStSql, True,-1,'',true);
        if not lQuery.Eof then
        begin
          for j := 0 to FFieldsCumul.count - 1 do
          begin
            P := PFieldsCumul( FFieldsCumul[j] );
            P^.Rdtotal := lQuery.Findfield( P^.StFieldName ).AsFloat;
          end;
          lNumTotal := lQuery.FindField('TOTAL').AsInteger;
        end;
      except
        on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Procédure : TMulPanelCumul.UpdateCumul');
      end;
    finally
      Ferme( lQuery );
    end;
  end;
  FPanelCumul.Caption := ' ' + TraduireMemoire('Totaux') + ' (' + IntToStr(lNumTotal)  + ' ' + TraduireMemoire('Lignes') + ')' ;
  { CA - 22/06/2005 : on mémorise le nombre de lignes total dans le Tag du Panel Cumul
                      pour utilisation dans FetchLesTous de uTofViergeMul }
  FPanelCumul.Tag := lNumTotal;

end ;
{$ENDIF}


procedure TMulPanelCumul.Update2_3Cumul;
var
 j  : integer;
 P  : PFieldsCumul;
 i  : integer;
begin

  // Mode 2/3, on parcours la TOB pour faire les cumuls
  p := nil ;
  try
  for i := 0 to FTOBListe.Detail.count - 1 do
  begin
    for j := 0 to FFieldsCumul.count - 1 do
    begin
      P := PFieldsCumul( FFieldsCumul[j] );
      P^.Rdtotal := P^.Rdtotal + VarAsType(TOBListe.Detail[i].GetValue(P^.StFieldName), vardouble);
    end;
  end;
  except
   on E : Exception do
    if P <> nil then PGIError('Erreur de cumul sur le champs : ' + P^.StFieldName + #10#13 + E.message + #10#13 + 'Procédure : TMulPanelCumul.UpdateCumul' ) ;
  end ;
  FPanelCumul.Caption := ' ' + TraduireMemoire('Totaux') + ' (' + IntTostr(FTobListe.Detail.Count) + ' ' + TraduireMemoire('Lignes') + ')' ;

end ;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :    /  /
Modifié le ... : 27/09/2004
Description .. : GCO - 27/08/2004 - FQ 14561
Mots clefs ... :
*****************************************************************}
procedure TMulPanelCumul.UpdateCumul;
var j         : integer;
begin
  if ( FPanelCumul = nil ) or ( FFieldsCumul.Count = 0 ) or (FTobListe = nil) or (FTobListe.Detail = nil) then exit;
  for j := 0 to FFieldsCumul.count - 1 do PFieldsCumul( FFieldsCumul[j] )^.RdTotal := 0 ;

{$IFDEF EAGLCLIENT}
 if FboFetchLesTous then
  Update2_3Cumul 
  else
   UpdateECumul;
{$ELSE}
  Update2_3Cumul ;
{$ENDIF}
  UpdateLayoutCumul;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TMulPanelCumul.UpdateLayoutCumul;
var
  i, j    : integer;
  lEdit   : THNumedit;
  R       : TRect;
begin
  if (FPanelCumul = nil) or (FTobListe = nil) then Exit;

  if (FFieldsCumul.count>0) then
  begin
    FPanelCumul.Visible := True;

    for j := 0 to FPanelCumul.ControlCount-1 do
    begin
      if FPanelCumul.Controls[j] is THNumEdit then
      begin
        lEdit := THNumEdit(FPanelCumul.Controls[j]);
        for i := 0 to FFieldsCumul.count - 1 do
        begin
          if PFieldsCumul( FFieldsCumul[i] )^.InColGrid = lEdit.tag then
          begin
            lEdit.Value    := PFieldsCumul( FFieldsCumul[i] )^.RdTotal;
            R              := FGrid.CellRect(lEdit.tag,0) ;
            lEdit.Left     := R.Left ;
            lEdit.Width    := R.Right - R.Left ;
            lEdit.Height   := FPanelCumul.Height ;
            break;
          end;
        end;
      end;
    end;
  end
  else
    FPanelCumul.Visible := false ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMulPanelCumul.UpdateCumulSelect;
var i,j : integer;
    P   : PFieldsCumul;
    lNbSelected : integer;
begin
  if ( FPanelCumul = nil ) or ( FFieldsCumulSelect.Count = 0 ) or (FTobListe = nil) or (FTobListe.Detail = nil) then exit;

  lNbSelected := 0;

  for j := 0 to FFieldsCumul.count - 1 do PFieldsCumul( FFieldsCumulSelect[j] )^.RdTotal := 0 ;

  // Mode 2/3, on parcours la TOB pour faire les cumuls
  p := nil ;
  try
    for i := 0 to FTOBListe.Detail.count - 1 do
    begin
      if FTobListe.Detail[i].GetString('SELECTED') = 'X' then
      begin
        Inc(lNbSelected);
        for j := 0 to FFieldsCumulSelect.count - 1 do
        begin
          P := PFieldsCumul( FFieldsCumulSelect[j] );
          P^.Rdtotal := P^.Rdtotal + VarAsType(FTOBListe.Detail[i].GetValue(P^.StFieldName), vardouble);
         end;
      end;
    end;
  except
    on E : Exception do
    if P <> nil then PGIError('Erreur de cumul sur le champs : ' + P^.StFieldName + #10#13 + E.message + #10#13 + 'Procédure : TMulPanelCumul.UpdateCumul' ) ;
  end ;
  FPanelCumulSelect.Caption := ' ' + TraduireMemoire('Total multi sélection') + ' (' + IntToStr(lNBSelected)  + ' ' + TraduireMemoire('Lignes') + ')' ;
  UpdateLayoutCumulSelect;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMulPanelCumul.UpdateLayoutCumulSelect;
var i, j : integer;
    lEdit   : THNumedit;
    R       : TRect;
begin
  if (FPanelCumulSelect = nil) or (FTobListe = nil) then Exit;

  FPanelCumulSelect.Visible := false ;
  if (FFieldsCumulSelect.count > 0) then
  begin
    FPanelCumulSelect.Visible := True;

    for j := 0 to FPanelCumulSelect.ControlCount-1 do
    begin
      if FPanelCumulSelect.Controls[j] is THNumEdit then
      begin
        lEdit := THNumEdit(FPanelCumulSelect.Controls[j]);
        for i := 0 to FFieldsCumulSelect.count - 1 do
        begin
          if PFieldsCumul( FFieldsCumulSelect[i] )^.InColGrid = lEdit.tag then
          begin
            lEdit.Value    := PFieldsCumul( FFieldsCumulSelect[i] )^.RdTotal;
            R              := FGrid.CellRect(lEdit.tag,0) ;
            lEdit.Left     := R.Left ;
            lEdit.Width    := R.Right - R.Left ;
            lEdit.Height   := FPanelCumul.Height ;
            break;
          end;
        end;
      end;
    end;
  end;
end;

function TMulPanelCumul.GetCumulChamps ( vStChamps : string ) : double ;
var
 i         : integer;
 P         : PFieldsCumul;
begin

 result    := 0 ;

 for i := 0 to FFieldsCumul.count - 1 do
  begin
   P := PFieldsCumul( FFieldsCumul[i] );
   if P^.StFieldName = vStChamps then
    begin
     result := P^.RdTotal ;
     exit ;
    end ;
  end ;

 if V_PGI.SAV then
  PGIError('Le champs ' + vStChamps + ' est introuvable dans la liste') ;

end ;

///////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2003
Modifié le ... : 21/01/2005
Description .. : Traduit le Value d' un Control THMultiValComboBox et 
Suite ........ : retourne
Suite ........ : dans vStSql la traduction SQL et dans vStLibelle la 
Suite ........ : traduction
Suite ........ : du Libelle
Suite ........ : GCO - 10/02/2004
Suite ........ : -> Interprétation de la MultiValComBox lorsque << Tous >> 
Suite ........ : est
Suite ........ : sélectionnée afin de prendre uniquement les valeurs 
Suite ........ : existantes
Suite ........ : Exemple : Problème E_QUALIFPIECE : N;S;U alors qu'il 
Suite ........ : existe
Suite ........ : V : Vrac; L : Encours de Lettrage ; Z : Centralisation des GL
Suite ........ : GCO - 11/02/2004
Suite ........ : -> Ajout du paramètre vVideSiTous par défaut à True, pour
Suite ........ : renvoyer une chaine vide dans vStSql ou si False,renvoie 
Suite ........ : toutes
Suite ........ : les valeurs listées du THMultiValComboBox
Suite ........ : GCO - 07/07/2004
Suite ........ : -> Ajout de la gestion des écritures typés Z ou V dans le cas
Suite ........ : unique pour VStNomChamp = E_QUALIFPIECE et 
Suite ........ : vVideSiTous = True
Suite ........ : SBO - 21/01/2004 
Suite ........ : On ne prend pas en compte les écritures de cloture 
Suite ........ : quand <<Tous>> sélectionné (FQ15057)
Suite ........ : GCO - 27/11/2006 - FQ 19121
Mots clefs ... : 
*****************************************************************}
procedure TraductionTHMultiValComboBox( vControl : THMultiValComboBox ; var vStSql : string ; var vStLibelle : string ; vStNomChamp : string ; vVideSiTous : Boolean = True; vStAjouteValue : string = '');
var lStValue         : string;
    lStValeurDecoupe : string;
    i : integer;
begin
  vStSql     := '';
  vStLibelle := '';

  if not vControl.Tous then
  begin
    lStValue   := vControl.Value;
    if lStValue <> '' then
    begin
      vStSql := '(';
      while lStValue <> '' do
      begin
        lStValeurDecoupe := ReadTokenSt(lStValue);
        vStSql := vStSql + vStNomChamp + '="' + lStValeurDecoupe + '" OR ';
        vStLibelle := vStLibelle + RechDom( vControl.DataType, lStValeurDecoupe, False) + ', ';
      end;

      if vStAjouteValue <> '' then
      begin
        vStSql := vStSql + vStNomChamp + '=' + vStAjouteValue + ' OR ';
      end;

      // Suppression du '" OR '
      vStSql := Copy(vStSql, 0, Length(vStSql) - 3) + ')';
      // Suppression du '", '
      vStLibelle := Copy(vStLibelle, 0, Length(vStLibelle) - 2);
    end;
  end
  else
  begin
    vStLibelle := TraduireMemoire('<<Tous>>');
    if vVideSiTous then
    begin
      // GCO -> Gestion V 
      // SBO 21/01/2004 On ne prend pas en compte les écritures de cloture (FQ15057)
      // GCO - 27/11/2006 - FQ 19121
      if vStNomChamp = 'E_QUALIFPIECE' then
        vStSql := '(E_QUALIFPIECE <> "Z" AND E_QUALIFPIECE <> "L" AND E_QUALIFPIECE <> "V" AND E_QUALIFPIECE <> "C")'
      else
        vStSql := '';
    end
    else
    begin
      vStSql := '(';
      for i := 0 to vControl.Values.Count - 1 do
      begin
        lStValeurDecoupe := vControl.Values[i] ;
        vStSql := vStSql + vStNomChamp + ' = "' + lStValeurDecoupe + '" OR ';
      end;

      if vStAjouteValue <> '' then
      begin
        vStSql := vStSql + vStNomChamp + '=' + vStAjouteValue + ' OR ';
      end;

      vStSql := Copy(vStSql, 0, Length(vStSql) - 3) + ')';   // Suppression du '" OR '
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{ TMulFiltre }

{
constructor TMulFiltre.Create(vStName : string ; vPopF : TPopUpMenu ; vFiltres : THValComboBox ; vPageControl : TPageControl ) ;
begin
 FStName            := vStName ;
 FPopF              := vPopF ;
 FFiltres           := vFiltres ;
 FPageControl       := vPageControl ;
 FFiltres.OnChange  := FFiltresChange ;
 SetPopF(FPopF) ;
end;

procedure TMulFiltre.SetPopF(vPopF : TPopUpMenu);
begin
 if vPopF.Items.Count <> 5 then exit ;
 FPopF.Items[0].OnClick := OnClickPopCreate ;
 FPopF.Items[1].OnClick := OnClickPopSave ;
 FPopF.Items[2].OnClick := OnClickPopDelete ;
 FPopF.Items[3].OnClick := OnClickPopRename ;
 FPopF.Items[4].OnClick := OnClickPopFind ;
 FPopF.OnPopup          := POPFPopup ;
end;

procedure TMulFiltre.OnClickPopCreate(Sender : TObject);
begin
 NewFiltre(FStName, FFiltres, FPageControl) ;
 OnClickPopSave(Self) ;
end;

procedure TMulFiltre.OnClickPopDelete(Sender : TObject);
begin
 DeleteFiltre(FStName, FFiltres) ;
end;

procedure TMulFiltre.OnClickPopFind(Sender : TObject);
begin
 if FFiltres <> nil then FFiltres.ItemIndex := -1 ;
 VideFiltre(FFiltres,FPageControl) ;
 if assigned(FOnPopFind) then FOnPopFind(self) ;
end;

procedure TMulFiltre.OnClickPopRename(Sender : TObject);
begin
 RenameFiltre(FStName, FFiltres) ;
 OnClickPopSave(Self) ;
end;

procedure TMulFiltre.OnClickPopSave(Sender : TObject);
begin
 SaveFiltre(FStName, FFiltres, FPageControl) ;
end;

procedure TMulFiltre.POPFPopup(Sender: TObject);
begin
 UpdatePopFiltre(FPopF.Items[1],FPopF.Items[2],FPopF.Items[3],FFiltres) ;
end;

procedure TMulFiltre.Charge ;
begin
 ChargeFiltre( FStName, FFiltres, FPageControl ) ;
end;

procedure TMulFiltre.FFiltresChange(Sender: TObject);
begin
 LoadFiltre( FStName, FFiltres, FPageControl) ;
 if assigned(FOnPopChange) then FOnPopChange(self) ;
end;   }

////////////////////////////////////////////////////////////////////////////////
procedure CVerifCritere( Ecran : TForm ) ;
var
 i : integer ;
begin
// exit;
 if not V_PGI.EgaliseFourchette then exit ;
 for i:=0 to Ecran.ComponentCount-1 do
  begin
   if (Ecran.Components[i] is THCritMaskEdit) and (TControl(Ecran.Components[i]).Parent is TTabSheet) and
      (TControl(Ecran.Components[i]).Visible) and (TControl(Ecran.Components[i]).Enabled) then
      begin
       V_PGI.EgaliseOnEnter(THCritMaskEdit(Ecran.Components[i])) ;
      end ;
   end ; // for
end;

////////////////////////////////////////////////////////////////////////////////
procedure CChargeZ_C( vControl : THValComboBox; vStListeChamps : string ; vStNewTitres : string );
var lStListeChamps : string;
    lStNewTitres   : string;
    lStReadToken   : string;
begin
  vControl.Items.Clear;
  vControl.Values.Clear;

  lStListeChamps := vStListeChamps;
  lStNewTitres   := vStNewTitres;

  while lStListeChamps <> '' do
  begin
    lStReadToken := ReadTokenSt(lStListeChamps);

    if Pos( '_', Copy( lStReadToken, 1, 4 )) > 0 then
    begin
      vControl.Items.Add(ReadTokenSt(lStNewTitres));
      vControl.Values.Add(lStReadToken);
    end
    else
      ReadTokenSt(lStNewTitres);
  end;
end;

////////////////////////////////////////////////////////////////////////////////


{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 15/07/2003
Modifié le ... : 15/07/2003
Description .. : Supprime le couple (value / item) de la liste du
Suite ........ : THValComboBox à partir de la value passé en paramètre
Mots clefs ... :
*****************************************************************}
procedure CVireLigne( vControl : THValComboBox; vStCode : String ) ;
Var lInIndex : Integer ;
begin
  // Recherche de la value dans la string liste
  lInIndex := vControl.Values.IndexOf(vStCode) ;
  // Si inéxistant, rien à faire, on sort
  if lInIndex < 0 then
    Exit ;
  // Sinon on vire les valeurs correspondantes à l'index trouvé
  vControl.Values.Delete(lInIndex) ;
  vControl.Items.Delete(lInIndex) ;
end ;

////////////////////////////////////////////////////////////////////
Function AbsoluteGoForward ( Ctrl1,Ctrl2 : TWinControl ) : Boolean ;
Var prnt1, prnt2   : TWinControl ;
    oPrnt1, oPrnt2 : TWinControl ;
    form           : TCustomForm ;
    ok             : Boolean ;
Begin

  Result := TRUE ;

  if (Ctrl1=nil) or (Ctrl2=nil) then exit ;

  Form := GetParentform(Ctrl1) ;
  if Form <> GetParentForm(Ctrl2) then exit ;

  Prnt1  := Ctrl1.parent ;
  Prnt2  := nil ;
  oPrnt1 := nil ;
  oPrnt2 := nil ;

  while (Prnt1<>Prnt2) do
    begin
    prnt2 := Ctrl2 ;
    ok    := FALSE ;
    while (Not Ok) and (prnt2<>Form) do
      begin
      oPrnt2 := Prnt2 ;
      Prnt2  := prnt2.parent ;
      ok     := ( Prnt2 = Prnt1 ) ;
      end ;
    if not ok then
      begin
      oPrnt1 := prnt1 ;
      prnt1  := Prnt1.parent ;
      end ;
    end ;

  //XMG 22/04/03 début
  if (prnt1<>ctrl1.parent) and (assigned(oprnt1)) then Prnt1 := oprnt1 ;
  if (prnt2<>ctrl2.parent) and (assigned(oprnt2)) then prnt2 := oprnt2 ;
  //XMG 22/04/03 fin

  if prnt1 = prnt2 then
     Begin
     Prnt1 := Ctrl1 ;
     prnt2 := Ctrl2 ;
     End ;

  Result := ( Prnt1.tabOrder <= Prnt2.TabOrder ) or ( Prnt2.TabOrder < 0 ) ;
End ;

/////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : XMG
Créé le ...... : 07/05/2003
Modifié le ... :
Description .. : Cette procedure permet de forcer le déplacement entre
Suite ........ : deux controles fenetre en prennant en compte l'ordre de
Suite ........ : tabulation de ces controles.
Suite ........ : (même si ils sont sur un parent different)
Suite ........ :
Suite ........ : Celui est util quand l'"enabled" d'un control change par
Suite ........ : rapport au contenu de le controle inmediatement precedent.
Suite ........ :
Suite ........ : P.Ex.: Le control "soumis TPF" n'est accesible que si le
Suite ........ : regime du tiers es "FRA"nce. L'ordre de tabulation de ces
Suite ........ : controles est:
Suite ........ : 1) Le regime
Suite ........ : 2) soumis TPF
Suite ........ :
Suite ........ : Lors de l'evenement on exit du Regime on change
Suite ........ : l'accesibilité (enabled) de SoumisTPF, cela produit la
Suite ........ : "perte" du focus quand on "disable" soumisTPF ou que le
Suite ........ : control SoumisTPF soit "sauté" quand il est "enablé".
Suite ........ :
Suite ........ : On resoult le souci en appelant GereNextControl après de
Suite ........ : méttre à jour la propriété Enabled du deuxième controle
Suite ........ : lors de l'onexit du premier.
Suite ........ :
Suite ........ : Parms[0].- Pointeur à la forme
Suite ........ : Parms[1].- Nom du controle de départ
Mots clefs ... :
*****************************************************************}
procedure GereNextControl( ff : TForm ; Ctrldepartie : String ) ;
var CtrlOri     : TWinControl ;
    CurrentCtrl : TWinControl ;
    OldCtrl     : TWinControl ;
begin
  CurrentCtrl := Screen.ActiveControl ;
  OldCtrl     := CurrentCtrl ;
  CtrlOri     := TWincontrol(FF.FindComponent(CtrlDePartie)) ; //XMG 04/06/03
  CurrentCtrl := TNextControlForm(FF).FindNextControl( CtrlOri, AbsoluteGoForward(CtrlOri,CurrentCtrl), TRUE, FALSE ) ;
  if (CurrentCtrl<>OldCtrl) and (assigned(CurrentCtrl)) and (CurrentCtrl.CanFocus)
     then CurrentCtrl.SetFocus ;
end ;


function TNextControlForm.FindNextControl(CurControl: TWinControl; GoForward, CheckTabStop, CheckParent: Boolean) : TWinControl;
begin
  result := Inherited FindNextControl( CurControl, GoForward, CheckTabStop, CheckParent ) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 12/10/2004
Modifié le ... : 12/10/2004
Description .. : Place le texte d'un contrôle en sélection inversé pour
Suite ........ : marquer le focus.
Suite ........ : Attention seule les THMultiValComBox  sont gérés pour le 
Suite ........ : moment (permet de gérer correctement le focus sur les 
Suite ........ : THMultiValComBox en ReadOnly )
Mots clefs ... : 
*****************************************************************}
procedure CSelectionTextControl( Sender : TObject ) ;
var lInTaille : Integer ;               // Taille du text affiché
    lControl  : THMultiValComboBox ;    // Control à gérer
begin
  if Sender = nil then Exit ;
  if not (Sender is THMultiValComboBox ) then Exit ;

  lControl  := THMultiValComboBox( Sender ) ;
  lInTaille := Length( lControl.Text ) ;
  if lInTaille > 0 then
    begin
    lControl.SelStart  := 0 ;
    lControl.SelLength := lInTaille ;
    end ;

end ;


procedure CUpdateGridListe ( FListe : THGrid ; FStListeParam : string ) ;
Var lStSource    : String ;
    lStLiens     : String ;
    lStTri       : String ;
    lStChamps    : String ;
    lStTitres    : HString ;
    lStLargeurs  : String ;
    lStJustifs   : String ;
    lStParams    : String ;
    lStLibelle   : HString ;
    lStNumCols   : HString ;
    lStPerso     : String ;
    lBoOkTri     : boolean;
    lBoOkNumCol  : boolean;
    i            : Integer ;
    lStNomChamp  : String ;
    lStType      : String ;
    lStNomF      : String ;
    lInDec       : integer ;
    lBoSep       : boolean;
    lBoObli      : boolean;
    lBoLib       : boolean;
    lBoVisu      : boolean;
    lBoNulle     : boolean;
    lBoCumul     : boolean;
    lAlign       : TAlignment;
    lStFormat    : String ;
    lInStartCol  : Integer ;
    lStGridColSize : string ;
begin

  // Récup du paramétrage
  ChargeHListe(   FStListeParam,
                  lStSource,
                  lStLiens,
                  lStTri,
                  lStChamps,
                  lStTitres,
                  lStLargeurs,
                  lStJustifs,
                  lStParams,
                  lStLibelle,
                  lStNumCols,
                  lStPerso,
                  lBoOkTri,
                  lBoOkNumCol ) ;

  FStListeParam  := lStChamps ;
  lStGridColSize := lStLargeurs ;

  // Démarrage à la 2ème colonne si DBIndicateur activé sur la grille
  lInStartCol := Ord ( FListe.DbIndicator ) ;

  // Parcours des lignes pour formatage spécifique des colonnes
  for i := lInStartCol to FListe.ColCount - 1 do
  begin
    lStNomChamp := ReadTokenST( lStChamps ) ;
    lAlign      := TransAlign(  ReadtokenSt( lStJustifs ),
                                lStFormat,
                                lInDec,
                                lBoSep,
                                lBoObli,
                                lBoLib,
                                lBoVisu,
                                lBoNulle,
                                lBoCumul) ;

    // Nom des colonnes
    FListe.ColNames[ i ] := lStNomChamp ;

    // Alignements
    FListe.ColAligns[i]  := lALign ;

    // Titres
    FListe.Cells[i,0]    := ReadTokenST( lStTitres ) ;

    // largeur des colonnes
    FListe.ColWidths[i] := ReadtokenI( lStLargeurs ) * 8;
    // GCO - 13/12/2005 - FQ 16936
    if not lBoVisu then
      FListe.ColWidths[i]  := -1;

    // Formats des colonnes...
    if Copy(lStNomChamp,1,2) = 'BO' then
      lStType := 'BOOLEAN'
    else
      lStType := ChampToType( lStNomChamp ) ;

    if ( lStType = 'BLOB' ) then
    begin
      FListe.ColTypes[i]   := 'M';
      FListe.ColFormats[i] := IntToStr( Byte( msBook ) );
    end
    else
      if ( lStType = 'BOOLEAN' ) then
      begin
        FListe.ColTypes[i]       := 'B';
        FListe.ColFormats[i]     := IntToStr( Byte( csCheckBox ) );
        FListe.ColEditables[ i ] := False ;
      end
      else
        if ( lStType = 'DOUBLE' ) or ( lStType = '??' ) then
        begin
          FListe.ColTypes[i] := 'R';
          FListe.ColFormats[i] := StrfMask(lInDec, '', True);
          if lBoNulle then
            FListe.ColFormats[i] := FListe.ColFormats[i] + ';;' + FListe.ColFormats[i];
        end
        else
          if ( lStType = 'DATE' ) then
          begin
            FListe.ColTypes[i]   := 'D';
            FListe.ColFormats[i] := '';  //sinon effet de bord // ??
          end
          else
            if Copy( Trim( lStNomChamp ), 1, 1 ) = '@' then // Champ formule ( NON MODIFIABLE ! )
            begin
              lStNomF := copy( lStNomChamp, pos('@',lStNomChamp) + 1, length(lStNomChamp) - pos('@',lStNomChamp) );
              FListe.ColFormats[i] := 'FORMULE=' + lStNomF + 'FORMAT=' + lStFormat;
              FListe.ColEditables[ i ] := False ;
            end
            else
            begin
              FListe.ColTypes[i]   := #0;
              FListe.ColFormats[i] := '';
            end;

   end ; // for

end;

////////////////////////////////////////////////////////////////////////////////
function TestJoker( ATester : string) : Boolean;
begin
  Result := (Pos('%', ATester) > 0) or (Pos('[', ATester) > 0) or (Pos('_', ATester) > 0) or
            (Pos(']', ATester) > 0) or (Pos('?', ATester) > 0) or (Pos('*', ATester) > 0);
end;

////////////////////////////////////////////////////////////////////////////////
//Constitution de la partie de la clause SQL WHERE d'une requête en fonction d'un contrôle
//contenant des caractères joker
function ClauseAvecJoker(ATraiter, Chp : string): string;
begin
  {On remplace "?" et "*" par leur équivalent en SQL}
  if (Pos('*', ATraiter) > 0) then ATraiter := FindEtReplace(ATraiter, '*', '%', True);
  if (Pos('?', ATraiter) > 0) then ATraiter := FindEtReplace(ATraiter, '?', '_', True);

  {Si on travailler sur une fourchette de caractères, cette fourchette doit être marquée
   par "-" et non par "," ou ";"}
  if (Pos('[', ATraiter) > 0) and (Pos(']', ATraiter) > 0) then begin
    if (Pos(',', ATraiter) > 0) then ATraiter := FindEtReplace(ATraiter, ',', '-', True);
    if (Pos(';', ATraiter) > 0) then ATraiter := FindEtReplace(ATraiter, ';', '-', True);
  end;

  {Il faut qu'il y ait un caractère "%" dans la chaîne}
  if ('%' <> ATraiter[Length(ATraiter)]) then ATraiter := ATraiter + '%';

  Result := Chp + ' LIKE "' + ATraiter + '" ';
end;

////////////////////////////////////////////////////////////////////////////////
function ConvertitCaractereJokers(ZoneDe, ZoneA : THEdit; Chp : string) : string;
begin
  if not Assigned(ZoneDe) or not Assigned(ZoneA) then Exit;
  if (Chp = '') or (Pos('_', Chp) = 0) then Exit;
  if TestJoker(ZoneDe.Text) then
    Result := ClauseAvecJoker(ZoneDe.Text, Chp) + ' '
  else if (ZoneDe.Text = '') and (ZoneA.Text = '') then
    Result := ' '
  else if (ZoneDe.Text <> '') and (ZoneA.Text <> '') then
    Result := Chp + ' BETWEEN "' + ZoneDe.Text + '" AND "' + ZoneA.Text + '" '
  else if (ZoneDe.Text <> '') then
    Result := Chp + ' >= "' + ZoneDe.Text + '" '
  else if (ZoneDe.Text <> '') then
    Result := Chp + ' <= "' + ZoneA.Text + '" '
end;

////////////////////////////////////////////////////////////////////////////////

{JP 05/10/05 : les valeurs forcées sur les combo d'un PageControl de critères sont écrasées
               par celles mémorisées dans la tob du filtre après le FormShow. Cette méthode
               permet de mémoriser dans le filtre la valeur que l'on veut forcer.

               Mem    : TFQRS1(Ecran).ListeFiltre.Current
               Combo  : Nom de la Combo ('NATURE', 'E_ETABLISSEMENT' ...)
               Valeur : Valeur que l'on veut forcer dans la combo, indépendamment de celle
                        mémorisée dans le filtre

 Cette fonction a été créée dans le cadre des restrictions utilisateurs sur les établissements.
 Un utilisateur peut en utilisant un filtre public voir les données d'un autre établissement !!
{---------------------------------------------------------------------------------------}
procedure ForceValeurSurFiltre(var Mem : TOB; Combo, Valeur : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  s : string;
begin
  if not Assigned(Mem) then Exit;

  for n := 0 to Mem.Detail.Count - 1 do begin
    s := Mem.Detail[n].GetValue('N');
    if s = Combo then begin
      Mem.Detail[n].PutValue('V', Valeur);
      Break;
    end;
  end;
end;

{JP 05/10/05 : Renvoie la valeur d'un composant qui est stockée dans un filtre
               Mem  : TFQRS1/TFMul(Ecran).ListeFiltre.Current
               Zone : Nom du composant ('NATURE', 'E_ETABLISSEMENT' ...)
{---------------------------------------------------------------------------------------}
function GetValeurDuFiltre(Mem : TOB; Zone : string) : variant;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  s : string;
begin
  Result := '';

  if not Assigned(Mem) then Exit;

  for n := 0 to Mem.Detail.Count - 1 do begin
    s := Mem.Detail[n].GetValue('N');
    if s = Zone then begin
      Result := Mem.Detail[n].GetValue('V');
      Break;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/02/2006
Modifié le ... : 14/06/2006
Suite ........ : GCO - 14/06/2006 - Ajout des comparateurs suivants 'est dans'
Suite ........ : 'n est pas dans' , 'est entre' , 'n est pas entre'.  
Description .. :
Mots clefs ... :
*****************************************************************}
function TraduitOperateur(vOperateurValue, vValue: variant; vBoTypeChaine : Boolean) : string;
var lStTemp : string;
    lStValue : string;
    lInPos : integer;
    lDbValue : Double;
begin
  Result := '';

  lStValue := VarToStr(vValue);

  if vOperateurValue <> '' then
  begin
    if (vOperateurValue = '<')  or (vOperateurValue = '<=') or
       (vOperateurValue = '<>') or (vOperateurValue = '=') or
       (vOperateurValue = '>')  or (vOperateurValue = '>=') then
    begin
      if vBoTypeChaine then
        Result := vOperateurValue + '"' + vValue + '"'
      else
        Result := vOperateurValue + VariantToSql(valeur(vValue));
    end
    else
    begin
      if vOperateurValue = 'C' then
        Result := 'LIKE "' + vValue + '%"'
      else
      if vOperateurValue = 'D' then
        Result := 'NOT LIKE "' + vValue + '%"'
      else
      if (vOperateurValue = 'E') or (vOperateurValue = 'G') then
      begin
        if vOperateurValue = 'E' then
          lStTemp := ' BETWEEN '
        else
          lStTemp := ' NOT BETWEEN ';

        lInPos := Pos('ET', UpperCase(lStValue));
        if lInPos > 0 then
        begin
          lDbValue := Valeur(Copy(lStValue,0, lInPos-1));
          lStTemp := lStTemp + VariantToSql( lDbValue );
          lDbValue := Valeur(Copy(lStValue,lInPos+2, length(lStValue)));
          lStTemp := lSTTemp + ' AND ' + VariantToSql( lDbValue );
          Result := lStTemp;
        end
        else
        begin
          lInPos := Pos(';', UpperCase(lStValue));
          if lInPos > 0 then
          begin
            lDbValue := Valeur(Copy(lStValue,0, lInPos-1));
            lStTemp := lStTemp + VariantToSql( lDbValue );
            lDbValue := Valeur(Copy(lStValue,lInPos+1, length(lStValue)));
            lStTemp := lSTTemp + ' AND ' + VariantToSql( lDbValue );
            Result := lStTemp;
          end;
        end;
      end
      else
      if (vOperateurValue = 'I') or (vOperateurValue = 'J') then
      begin
        if vOperateurValue = 'I' then
          lStTemp := ' IN ('
        else
          lStTemp := ' NOT IN (';

        repeat
          lDbValue := Valeur(ReadTokenSt(lStValue));
          lStTemp := lStTemp + VariantToSql( lDbValue ) + ',';
        until ( lStValue = '');

        lStTemp := Copy(lStTemp, 0, length(lStTemp)-1) + ')';
        Result := lStTemp;
      end
      else
      if vOperateurValue = 'L' then
        Result := ' LIKE "%' + vValue + '%"'
      else
      if vOperateurValue = 'M' then
        Result := ' NOT LIKE "%' + vValue + '%"'
      else
      if vOperateurValue = 'V' then
      begin
        Result := ' = "" ';
      end
      else
      if vOperateurValue = 'VV' then
      begin
        Result := ' <> ""';
      end;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////

{JP 08/03/06 : Compare deux chaînes dont l'une (ATester) peut contenir des caractères Joker.
               Renvoie True si les deux chaînes sont identiques ou si les caractères Joker
               de ATester sont en cohérence avec Les caractères de AComparer
{---------------------------------------------------------------------------------------}
function CompareAvecJoker(ATester, AComparer : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n, p : Integer;

    {-----------------------------------------------------------------------}
    function _IsJoker(c : Char) : Boolean;
    {-----------------------------------------------------------------------}
    begin
      Result := c in ['%', '[', '_', ']', '?', '*'];
    end;

    {Deux possibilités :
     1/ La chaîne entre crochets est longue de 3 caractèes et de la forme X-Y :
        dans ce cas le caractère équivalent de AComparer doit être compris dans
        la fourchette X-Y
     2/ Dans tous les autres cas, le caractère équivalent de AComparer  peut
        prendre la valeur de n'importe quel caractère compris dans les crochets
    {-----------------------------------------------------------------------}
    function _OkFourchette : Boolean;
    {-----------------------------------------------------------------------}
    var
      j, k : Integer;
      i    : Integer;
      s    : string;
    begin
      j := n + 1;
      k := Pos(']', ATester);
      s := Copy(ATester, j, k - j);

      if (Pos(',', s) > 0) then s := FindEtReplace(s, ',', '-', True);
      if (Pos(';', s) > 0) then s := FindEtReplace(s, ';', '-', True);

      {Cas 1}
      if (Length(s) = 3) and (Pos ('-', s) = 2) then
        Result := (AComparer[p] >= s[1]) and (AComparer[p] <= s[3])
      {Cas 2}
      else begin
        i := 1;
        Result := False;
        while (i <= Length(s)) and not Result do begin
          Result := AComparer[p] = s[i];
          Inc(i);
        end;
      end;

      Inc(p);
      n := k;
    end;

begin
//  AComparer := AnsiUpperCase(AComparer);
//  ATester   := AnsiUpperCase(ATester);
  Result := True;

  if TestJoker(ATester) then begin
    p := 1;
    n := 1;
    {05/06/06 : gestion du cas où AComparer = ''}
    if AComparer = '' then begin
      {On ne renvéra True que si ATester n'est constitué que de '*' ou de '%'}
      while (n <= Length(ATester)) do begin
        if ATester[n] in ['*', '%'] then
          Inc(n)
        else begin
          Result := False;
          Break;
        end;
      end;
      Exit;
    end;

    while (n <= Length(ATester)) and Result do begin
      if _IsJoker(ATester[n]) then begin
        {Le caractère correspondant dans AComparer peut prendre n'importe quelle valeur ...}
        if ATester[n] in ['_', '?'] then
          {... on passe au caractère suivant}
          Inc(p)

        else if ATester[n] in ['*', '%'] then begin
          {Si '*' ou '%' est le dernier caractère de ATester, peu importe le reste de la
           chaîne AComparer, Result vaut True, par contre ...}
          if n < Length(ATester) then begin
            {... On regarde quel est le caractère suivant pour voir si on le trouve dans AComparer}
            Inc(n);
            while p <= Length(AComparer) do begin
              if AComparer[p] = ATester[n] then begin
                Inc(p);
                Break;
              end
              else
                Inc(p);
            end;
            {Si on a balayé AComparer sans trouver ATester[n], Result = False}
            Result := not ((p > Length(AComparer)) and (AComparer[p - 1] <> ATester[n]));
          end;
        end

        else if ATester[n] = '[' then
          Result := _OkFourchette;
      end
      else begin
        if p <= Length(AComparer) then Result := ATester[n] = AComparer[p]
                                  else Result := False;
        Inc(p);
      end;
      Inc(n);
    end;
  end
  else
    Result := AComparer = ATester;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function FullMajuscule(S: string): string;
var lIndex : integer;
    lCar : Char;
begin
  Result := '';
  lIndex := 0;

  while lIndex < Length(s) do
  begin
    lCar := S[lIndex + 1];

    if (lCar >= 'A') and (lCar <= 'Z') then
      Result := Result + lCar
    else
    if lCar in ['à','á','â','ã','ä','å','À','Á','Â','Ã','Ä','Å'] then
      Result := Result + 'A'
    else
    if lCar in ['ç','¢','Ç'] then
      Result := Result + 'C'
    else
    if lCar in ['è','é','ê','ë','È','É','Ê','Ë'] then
      Result := Result + 'E'
    else
    if lCar in ['ì','í','î','ï','Ì','Í','Î','Ï'] then
      Result := Result + 'I'
    else
    if lCar in ['ñ'] then
      Result := Result + 'N'
    else
    if lCar in ['ù','ú','û','ü','Ù','Ú','Û','Ü'] then
      Result := Result + 'U'
    else
    if lCar in ['ý','ÿ','Ý'] then
      Result := Result + 'Y'
    else
    if lCar in ['Ð'] then
      Result := Result + 'D'
    else
    if lCar in ['ò','ó','ô','õ','ö','Ò','Ó','Ô','Õ','Ö'] then
      Result := Result + 'O'
    else
      Result := Result + UpperCase(lCar);

    Inc(lIndex);
  end;
end;

{JP 05/06/07 : lorsque l'on veut cacher une colonne dans une grille, il faut mettre ColWidths à -1 ou 0 selon l'OS,
               car avec -1 sous Vista et XP, la colonne précédente est un peu "mangée".
               Cette fonction ne traite pas le cas d'un OS XP et d'un look NT ...
{---------------------------------------------------------------------------------------}
function GetWidthColFromOS : Integer;
{---------------------------------------------------------------------------------------}
begin
  if (V_PGI.WinVersion = '2K') or
     (V_PGI.WinVersion = 'NT') or
     (V_PGI.WinVersion = 'ME') or
     (V_PGI.WinVersion = '98') or
     (V_PGI.WinVersion = '95') then Result := - 1
                               else Result := 0;

end;
(*
procedure CreationRec( vTOB : TOB; vStListeChamps : string );
var
 lIndex        : integer ;
 lStChamps     : string ;
 lStSauvChamps : string ;
begin

 lIndex        := 0 ;
 lStSauvChamps := vStListeChamps ;

 repeat
  lStChamps := ReadTokenSt(vStListeChamps);
  inc(lIndex) ;
 until (vStListeChamps = '') or (lStChamps = '') ;

 Setlength( vRec , lIndex ) ;

 lIndex := 0 ;

 repeat
  lStChamps          := ReadTokenSt(lStSauvChamps);
  vRec[lIndex].Index := vTOB.GetNumChamp(lStChamps) ;
  vRec[lIndex].Nom   := lStChamps ;
  inc(lIndex) ;
 until ( lStSauvChamps = '') or (lStChamps = '') ;

end ;   *)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function Rotate90( vBitMap : TBitMap ) : TBitMap;
type
  TManoRGB  = packed record
                 rgb    : TRGBTriple;
                 dummy  : byte;
              end;
  TRGBArray = ARRAY[0..0] OF TRGBTriple;
  pRGBArray = ^TRGBArray;
var
  aStream : TMemorystream;          // zone mémoire
  header  : TBITMAPINFO;            // header bitmap
  dc      : hDC;                    // ressource pour GetDIBits
  P       : ^TManoRGB;              // pointeur vers 4 octets
  RowOut  :  pRGBArray;             // pointeur vers 3 octets
  x,y,h,w : integer;
  Bitmap  : TBitmap;
begin

  Bitmap := TBitmap.Create;

  // Bitmap.canvas.Draw(0,0,bmp1);
  {JP : problème avec le Draw !!}
  //bmp1 := TSmoothImage( FPanelActif.Controls[0] ).Picture ;

  Bitmap.Width  := vBitMap.Width;
  Bitmap.Height := vBitMap.Height;
  Bitmap.Canvas.StretchDraw(Rect(0, 0, vBitMap.width-1, vBitMap.height-1), vBitMap);
  {JP : problème avec le Draw !!}

  Bitmap.Pixelformat := pf24bit;
  aStream := TMemoryStream.Create;                 // réservation mémoire
  aStream.SetSize(Bitmap.Height*Bitmap.Width * 4); // chaque pixel = 4 octets
  with header.bmiHeader do begin                   // bitmap mémoire
    biSize := SizeOf(TBITMAPINFOHEADER);
    biWidth := Bitmap.Width;
    biHeight := Bitmap.Height;
    biPlanes := 1;
    biBitCount := 32;              // 32 bits par pixel = 4 octets
    biCompression := 0;
    biSizeimage := aStream.Size;
    biXPelsPerMeter :=1;
    biYPelsPerMeter :=1;
    biClrUsed :=0;
    biClrImportant :=0;
  end;
  dc := GetDC(0);                  // folklore des ressources GDI windows
  P  := aStream.Memory;
  // copie du bitmap dans le flux (stream). Passe de 3 à 4 octets
  GetDIBits(dc,Bitmap.Handle,0,Bitmap.Height,P,header,dib_RGB_Colors);
  ReleaseDC(0,dc);                 // folklore des ressources GDI windows
  w := bitmap.Height;
  h := bitmap.Width;
  bitmap.Width  := w;              // permute largeur / hauteur
  bitmap.height := h;
  for y := 0 to (h-1) do     // boucle pilotée par y et x coodonnées sortie
  begin
    rowOut := Bitmap.ScanLine[y]; // sortie 3 octets = 24 bits par pixels
    P  := aStream.Memory;
    inc(p,y);                    // p = adresse ligne du stream
    for x := 0 to (w-1) do
    begin
      rowout[x] := p^.rgb;        // copie 3 octets sur les 4 du stream
      inc(p,h);                   // parcours de la ligne du stream
    end;
  end;
  aStream.Free;

  //TSmoothImage( FPanelActif.Controls[0] ).Picture := Bitmap ;

//  TImage(GetControl('IMAGE')).picture.bitmap.assign(bitmap);
//  TImage(GetControl('IMAGE')).Refresh;

  Result := BitMap;

  FreeAndNil(bitmap);
end;

////////////////////////////////////////////////////////////////////////////////

end.

