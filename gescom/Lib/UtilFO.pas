unit UtilFO ;

interface

Uses HEnt1, HCtrls, UTOB, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     SysUtils, Classes, Graphics, Grids, Forms, Saisutil, EntGC, AGLInit,Messages,
     Windows, LCD_Lab, HPanel,
{$IFDEF MODE}
     EcheanceFO,
{$ENDIF}
     FactUtil, FactComm, FactTOB,uEntCommun ;

Type TAffichages    = (afNone, afTexte, afQuantite, afMontant, afClient ) ;
     TSetAffichages = Set of TAffichages ;

procedure ResizerAfficheur ( Aff : TLCDLabel ; Panel : THPanel ) ;
Procedure InverseAfficheur ( Aff : TLCDLabel ) ;
Procedure AffichageLCD ( LCD : TLCDLabel ; Texte : String ; Qte,Mnt : Double ; Ou : TSetAffichages ; Attente : Boolean ; DEV : RDEVISE ) ;
Function  GereEcheancesMODE ( TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG : TOB ; Action : TActionFiche ; DEV : RDevise ; Ouv : boolean ) : Boolean ;

implementation

procedure ResizerAfficheur ( Aff : TLCDLabel; Panel : THPanel ) ;
BEGIN
if Not (ctxFO in V_PGI.PGIContexte) then Exit ;
Aff.Width:=Panel.Width ;
Aff.Left:=(Panel.Width-Aff.Width) div 2 ;
Aff.CalcCharSize;
Aff.Invalidate ;
END ;

Procedure InverseAfficheur ( Aff : TLCDLabel ) ;
BEGIN
{****
if Inversion then
   BEGIN
   Aff.Top := 1 ;
   Aff.UpSideDown := TRUE ;
   Aff.Top := PAffichage.Height - FAffTXT.Height ;
   Aff.UpSideDown := TRUE ;
   END else
   BEGIN
   Aff.Top := 1 ;
   Aff.UpSideDown := FALSE ;
   Aff.Top := FAffTXT.Height ;
   Aff.UpSideDown := FALSE ;
   END ;
****}
END ;

Procedure AffichageLCD ( LCD : TLCDLabel ; Texte : String ; Qte,Mnt : Double ; Ou : TSetAffichages ; Attente : Boolean ; DEV : RDEVISE ) ;
Var Libelle, Valeur : string ;
    Taille, Pos, Nb : Integer ;
BEGIN
if Not (ctxFO in V_PGI.PGIContexte) then Exit ;
if LCD=Nil then Exit ;
if Not LCD.Visible then exit ;
if LCD.Tag <> 0 then Delay(2500) ; //If tag<>0, le dernier message doit être affiché plus temps
LCD.Tag := ord(Attente) ;
LCD.Caption := '' ;
Valeur := '' ;
Libelle := '';
if afTexte in Ou then Libelle := Uppercase(traduirememoire(Texte)) ;
if afQuantite in Ou then
   BEGIN
   Taille := LCD.NoOfChars ;
   if Length(Libelle) > Taille then
      BEGIN
      // on tronque le texte si on ne dispose pas d'assez de place pour le montant
      Pos := Taille ;
      Nb := Length (Libelle) - Taille ;
      Delete(Libelle, Pos, Nb) ;
      END else
      BEGIN
      // on complete par des espaces pour se placer sur la 2ème ligne
      while Length(Libelle) < Taille do Libelle := Libelle + ' ' ;
      END ;
      Libelle := Libelle + StrfMontant(Qte, LCD.noOfChars, V_PGI.OkDecQ, '', True) + ' X '
   END ;
if afMontant in Ou then Valeur := StrfMontant(Mnt, LCD.noOfChars, DEV.Decimale, DEV.Symbole, True) ;
if Valeur <> '' then
   BEGIN
   Taille := LCD.NoOfChars * LCD.TextLines ;
   // on tronque le texte si on ne dispose pas d'assez de place pour le montant
   if Length(Valeur) > (Taille - Length(Libelle) + 1) then
      BEGIN
      Pos := Taille - Length(Valeur) ;
      Nb := Length (Libelle) - Length(Valeur) + 1 ;
      Delete(Libelle, Pos, Nb) ;
      END ;
   // cadrage du montant à gauche
   while Length(Valeur) < (Taille - length(Libelle)) do Valeur := ' ' + Valeur ;
   END ;
LCD.Caption := Libelle + Valeur ;
END ;

Function GereEcheancesMODE ( TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG : TOB ; Action : TActionFiche ; DEV : RDevise ; Ouv : boolean ) : Boolean ;
var CleDoc : R_CLEDOC ;
GereEche : String ;
BEGIN
CleDoc:=TOB2CleDoc(TOBPiece) ;
GereEche:=GetInfoParPiece(CleDoc.NaturePiece,'GPP_GEREECHEANCE') ;
{$IFDEF MODE}
//CleDoc:=TOB2CleDoc(TOBPiece) ;
if CleDoc.NaturePiece='FFO' then
   BEGIN
   Result:=GereEcheancesFO(Nil,Cledoc,TOBPiece,TOBEches,TOBAcomptes,Action) ;
   Exit ;
   END ;
{$ENDIF}
Result:=GereEcheancesGC(TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,nil,nil,Action,DEV,Ouv) ;
END ;

end.
