//---------------------------------------------------------------------------
//--- Auteur : CATALA David
//--- Objet  : Tableau de bord / Analyse activitée : Fonctions traitements
//---------------------------------------------------------------------------
unit DpTableauBordLibrairie;

interface

uses
 Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
 StdCtrls, Grids,
{$IFDEF EAGLCLIENT}
 UTob,
{$ELSE}
 HDB,
 {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
 ExtCtrls, Hctrls, Hent1,Hmsgbox;

//--------------------------------------------------
//--- Déclaration des constantes
//--------------------------------------------------
const
 GAUCHE  =  1;
 CENTRE  =  2;
 DROITE  =  3;

//--------------------------------------------------
//--- Déclaration des procédures
//--------------------------------------------------
 procedure GriserGrille (ACol, ARow : Integer; G : THGrid);
 procedure NoirsirGrille (ACol,ARow:Integer;G:THGrid);
 procedure EcrireTexte (TypeAlign,ACol,ARow : Integer; G : THGrid);

//--------------------------------------------------
//--- Déclaration des fonctions
//--------------------------------------------------
 function FormaterMontant (Montant : Double) : String;
 function FormaterNombre (Chaine : String) : String;

 function IsDBZero (NumDossier : String) : Boolean;
 function IsCodeTiersExiste (NumDossier : String) : Boolean;
 function DonnerLibelleDossier (NumDossier : String) : String;
 function DonnerAdresseEmail (ChNumDossier : String) : String;

 function DonnerDossierFromGuid (GuidPer : String) : String;
 function DonnerGuidFromDossier (NumDossier : String) : String;
 function DonnerGuidFromTiers (CodeTiers : String) : String;
 function DonnerTiersFromGuid (GuidPer : String) : String;
 function DonnerAuxiFromGuid (GuidPer : String) : String;
 function DonnerGenFromAuxi (CodeAuxi : String) : String;

 function RechercherPosChaine (Chaine,SousChaine : String; NumOccurence : Integer) : Integer;
 function RemplacerChaine (var ChaineSource : String; SousChaineAChercher,SousChaineRemplacerPar : String) : Boolean;

implementation

//-------------------------------------------------
//--- Nom   : GriserGrille
//--- Objet : Grise une cellule de la Stringlist
//-------------------------------------------------
procedure GriserGrille (ACol,ARow:Integer;G:THGrid);
var
 Rectangle   : TRect ;
 lOldBrush   : TBrush ;
 lOldPen     : TPen ;
begin
 //--- Sauvegarde des valeurs courantes
 lOldBrush:=TBrush.Create ;
 lOldPen  :=TPen.Create ;
 lOldBrush.assign(G.Canvas.Brush) ;
 lOldPen.assign(G.Canvas.Pen) ;

 try
  G.Canvas.Brush.Color := G.FixedColor ;
  G.Canvas.Brush.Style := bsBDiagonal ;
  G.Canvas.Pen.Color   := G.FixedColor ;
  G.Canvas.Pen.Mode    := pmCopy ;
  G.Canvas.Pen.Style   := psClear ;
  G.Canvas.Pen.Width   := 1 ;
  Rectangle            := G.CellRect(ACol,ARow);
  G.Canvas.Rectangle(Rectangle.Left,Rectangle.Top,Rectangle.Right+1,Rectangle.Bottom+1) ;
 finally
  //--- Réaffectation des valeurs du canvas
  G.Canvas.Brush.Assign(lOldBrush);
  G.Canvas.Pen.Assign(lOldPen);
  if assigned(lOldBrush) then lOldBrush.Free;
  if assigned(lOldPen) then lOldPen.Free;
 end;
end ;

//---------------------------------------------------
//--- Nom   : NoisirGrille
//--- Objet : Noirsie une cellule de la Stringlist
//---------------------------------------------------
procedure NoirsirGrille (ACol,ARow:Integer;G:THGrid);
var
 Rectangle   : TRect ;
 lOldBrush   : TBrush ;
 lOldPen     : TPen ;
begin
 //--- Sauvegarde des valeurs courantes
 lOldBrush:=TBrush.Create ;
 lOldPen  :=TPen.Create ;
 lOldBrush.assign(G.Canvas.Brush) ;
 lOldPen.assign(G.Canvas.Pen) ;

 try
  G.Canvas.Brush.Color := G.FixedColor ;
  G.Canvas.Brush.Style := BsSolid;
  G.Canvas.Pen.Color   := G.FixedColor ;
  G.Canvas.Pen.Mode    := pmCopy ;
  G.Canvas.Pen.Style   := psSolid;
  G.Canvas.Pen.Width   := 1 ;
  Rectangle            := G.CellRect(ACol,ARow);
  G.Canvas.Rectangle(Rectangle.Left,Rectangle.Top,Rectangle.Right+1,Rectangle.Bottom+1) ;
 finally
  //--- Réaffectation des valeurs du canvas
  G.Canvas.Brush.Assign(lOldBrush);
  G.Canvas.Pen.Assign(lOldPen);
  if assigned(lOldBrush) then lOldBrush.Free;
  if assigned(lOldPen) then lOldPen.Free;
 end;
end ;

//------------------------------------------------------------
//--- Nom   : EcrireTexte
//--- Objet : Aligne le texte dans une cellule de la grille
//------------------------------------------------------------
procedure EcrireTexte (TypeAlign,ACol,ARow:Integer; G:THGrid);
var lOldBrush   : TBrush;
    lOldPen     : TPen;
    Rectangle   : TRect;
    Texte       : array[0..255] of Char;
begin
 //--- Sauvegarde des valeurs courantes
 lOldBrush:=TBrush.Create ;
 lOldPen  :=TPen.Create ;
 lOldBrush.assign(G.Canvas.Brush) ;
 lOldPen.assign(G.Canvas.Pen) ;

 try
  Rectangle:= G.CellRect(ACol,ARow);
  StrPCopy(Texte,G.Cells[ACol,ARow]);

  if (G.ColWidths [ACol]<G.Canvas.textwidth (Texte)+6) then
   G.ColWidths [ACol]:=(G.Canvas.textwidth (Texte)+6);

  case TypeAlign of
   Droite : G.Canvas.TextRect (Rectangle,Rectangle.Right-G.Canvas.TextWidth (Texte)-3,Rectangle.Top+2,Texte);
   Centre : G.Canvas.TextRect (Rectangle,Rectangle.Left+((Rectangle.Right-Rectangle.Left-G.canvas.TextWidth(Texte)) div 2),Rectangle.Top+2,Texte);
   Gauche : G.Canvas.TextRect (Rectangle,Rectangle.Left+2,Rectangle.Top+2,Texte);
  end;


 finally
  //--- Réaffectation des valeurs du canvas
  G.Canvas.Brush.Assign(lOldBrush);
  G.Canvas.Pen.Assign(lOldPen);
  if assigned(lOldBrush) then lOldBrush.Free;
  if assigned(lOldPen) then lOldPen.Free;
 end;
end;

//------------------------------------------------------------
//--- Nom   : FormaterMontant
//--- Objet : Sépare les milliers d'un montant par un espace
//------------------------------------------------------------
function FormaterMontant (Montant : Double) : String;
var ChResultat : String;
begin
 ChResultat:='';

 if (Montant<0) then
  ChResultat:=StrFMontant (Abs (Montant),15,2,'',True)+' C'
 else
  ChResultat:=StrFMontant (Abs (Montant),15,2,'',True)+' D';

 FormaterMontant:=ChResultat;
end;

//------------------------------------------------------------
//--- Nom   : FormaterNombre
//--- Objet : Sépare les milliers d'un nombre
//------------------------------------------------------------
function FormaterNombre (Chaine : String) : String;
var ChResultat : String;
    Longueur   : Integer;
    Indice : Integer;
begin
 Longueur:=length (Chaine);
 ChResultat:='';

 for Indice:=1 to Longueur do
  begin
   ChResultat:=ChResultat+Chaine [Indice];
   if (((Longueur-Indice) mod 3)=0) then
    if (Longueur<>Indice) then
     ChResultat:=ChResultat+' ';
  end;

 FormaterNombre:=ChResultat;
end;

//----------------------------------------------------------------------
//--- Nom   : IsCodeTiersExiste
//--- Objet : Détermine si un code Tiers existe pour le dossier donné
//----------------------------------------------------------------------
function IsCodeTiersExiste (NumDossier : String) : Boolean;
begin
 IsCodeTiersExiste:=(DonnerTiersFromGuid (DonnerGuidFromDossier (NumDossier))<>'');
end;

//--------------------------------------------------
//--- Nom   : IsDbZero
//--- Objet : Détermine si c'est la BASE DB000000
//--------------------------------------------------
function IsDBZero (NumDossier : String) : Boolean;
begin
 IsDBZero:=(NumDossier='000000');
end;

//--------------------------------------------------
//--- Nom   : DonnerLibelleDossier
//--- Objet :
//--------------------------------------------------
function DonnerLibelleDossier (NumDossier : String) : String;
var ChSql      : String;
    RSql       : Tquery;
    ChResultat : String;
begin
 ChResultat:='';
 ChSql:='SELECT DOS_LIBELLE FROM DOSSIER WHERE DOS_NODOSSIER="'+NumDossier+'"';
 RSql:=OpenSql (ChSql,True);
 if (Not RSql.Eof) then
  ChResultat:=RSql.FindField ('DOS_LIBELLE').AsString;
 Ferme (RSql);

 DonnerLibelleDossier:=ChResultat;
end;

//--------------------------------------------------
//--- Nom   : DonnerDossierFromGuid
//--- Objet : Donne Numéro d'un dossier
//--------------------------------------------------
function DonnerDossierFromGuid (GuidPer : String) : String;
var ChSql      : String;
    RSql       : TQuery;
begin
 Result:='';
 ChSql:='SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_GUIDPER="'+GuidPer+'"';
 RSql:=OpenSql (ChSql,True);
 if (Not RSql.Eof) then
  Result:=RSql.FindField ('DOS_NODOSSIER').AsString;
 Ferme (RSql);
end;

//--------------------------------------------------
//--- Nom   : DonnerGuidFromDossier
//--- Objet : Donne le Guid d'un dossier
//--------------------------------------------------
function DonnerGuidFromDossier (NumDossier : String) : String;
var ChSql      : String;
    RSql       : Tquery;
begin
 Result:='';
 ChSql:='Select DOS_GUIDPER From DOSSIER where DOS_NODOSSIER="'+NumDossier+'"';
 RSql:=OpenSql (ChSql,True);
 if (Not RSql.Eof) then
  Result:=RSql.FindField ('DOS_GUIDPER').AsString;
 Ferme (RSql);
end;

//----------------------------------------------------
//--- Nom   : DonnerGuidFromTiers
//--- Objet :
//----------------------------------------------------
function DonnerGuidFromTiers (CodeTiers : String) : String;
var ChSql : String;
    RSql  : TQuery;
begin
 Result:='';
 ChSql:='SELECT ANN_GUIDPER FROM ANNUAIRE WHERE ANN_TIERS="'+CodeTiers+'"';
 RSql:=OpenSql (ChSql,True);
 if (Not RSql.Eof) then
  Result:=RSql.FindField ('ANN_GUIDPER').AsString;
 Ferme (RSql);
end;

//----------------------------------------------------
//--- Nom   : DonnerTiersFromGuid
//--- Objet : Détermine si c'est un dossier cabinet
//----------------------------------------------------
function DonnerTiersFromGuid (GuidPer : String) : String;
var ChSql      : String;
    RSql       : Tquery;
begin
 Result:='';
 if (GuidPer<>'') then
  begin
   ChSql:='Select ANN_TIERS From ANNUAIRE where ANN_GUIDPER="'+GuidPer+'"';
   RSql:=OpenSql (ChSql,True);
   if (Not RSql.Eof) then
    Result:=RSql.FindField ('ANN_TIERS').AsString;
   Ferme (RSql);
  end;
end;

//---------------------------------------
//--- Nom   : DonnerAuxiFromGuid
//--- Objet : Donne le code Auxiliaire
//---------------------------------------
function DonnerAuxiFromGuid (GuidPer : String) : String;
var ChSql      : String;
    RSql       : Tquery;
    CodeTiers  : String;
begin
 Result:='';
 CodeTiers := DonnerTiersFromGuid(GuidPer);

 ChSql:='Select T_AUXILIAIRE From TIERS where T_TIERS="'+CodeTiers+'"';
 RSql:=OpenSql (ChSql,True);
 if (Not RSql.Eof) then
  Result:=RSql.FindField ('T_AUXILIAIRE').AsString;
 Ferme (RSql);
end;

//-------------------------------------
//--- Nom   : DonnerGenFromAuxi
//--- Objet : Donne le code Generale
//-------------------------------------
function DonnerGenFromAuxi (CodeAuxi : String) : String;
var ChSql      : String;
    RSql       : Tquery;
begin
 Result:='';
 ChSql:='Select T_COLLECTIF From TIERS where T_AUXILIAIRE="'+CodeAuxi+'"';
 RSql:=OpenSql (ChSql,True);
 if (Not RSql.Eof) then
  Result:=RSql.FindField ('T_COLLECTIF').AsString;
 Ferme (RSql);
end;

//------------------------------------
//--- Nom   : DonnerAdresseEmail
//--- Objet : Donne l'adresse Email
//------------------------------------
function DonnerAdresseEmail (ChNumDossier : String) : String;
var ChSql : String;
    Rsql  : TQuery;
begin
 Result:='';
 if (ChNumDossier<>'') then
  begin
   ChSql:='SELECT ANN_EMAIL FROM ANNUAIRE WHERE ANN_GUIDPER="'+DonnerGuidFromDossier (ChNumDossier)+'"';
   RSql:=OpenSql (ChSql,True);
   if not RSql.Eof then
    Result:=RSql.FindField ('ANN_EMAIL').AsString;
   Ferme (Rsql)
  end;
end;

{------------------------------------------------------------
 --- Nom       : FONCTION RechercherPosChaine
 --- Objet     : Recherche une sous chaine dans une chaine
 ------------------------------------------------------------}
function RechercherPosChaine (Chaine,SousChaine : String; NumOccurence : Integer) : Integer;
var Position, Indice : Integer;
    SourceChaine : String;
begin
 SourceChaine:=Chaine;
 Position:=0;

 for Indice:=1 to NumOccurence do
  begin
   Position:=Position+Pos (SousChaine,Chaine);
   Chaine:=copy (SourceChaine,Position+1,Length (SourceChaine)-Position);
  end;

 RechercherPosChaine:=Position;
end;

{----------------------------------------------------
 --- Nom       : FONCTION RemplacerChaine
 --- Objet     : Remplacerune chaine par une autre
 ----------------------------------------------------}
function RemplacerChaine (var ChaineSource : String; SousChaineAChercher,SousChaineRemplacerPar : String) : Boolean;
Var ChaineDest : String;
    TrouverSousChaine : Boolean;
    Longueur,Position : Integer;
begin
 Longueur:=Length (SousChaineAChercher);
 Position:=pos (SousChaineAChercher,ChaineSource);

 if (Position<>0) then
  begin
   ChaineDest:=copy (ChaineSource,1,Position-1)+SousChaineRemplacerPar+copy (ChaineSource,Position+Longueur,length (ChaineSource)-Position-Longueur);
   ChaineSource:=ChaineDest;
   TrouverSousChaine:=True;
  end
 else
  TrouverSousChaine:=False;

 RemplacerChaine:=TrouverSousChaine;
end;

end.

