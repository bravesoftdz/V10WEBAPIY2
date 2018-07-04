{***********UNITE*************************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 04/05/2000
Modifié le ... : 04/05/2000
Description .. : Rapprochement automatique récursif qui permet de trouver
Suite ........ : des solutions "approchantes" en fonction d'un seuil passé
Suite ........ : à la fonction
Mots clefs ... : RAPPROCHEMENT;COMBINATOIRE;APPROCHE;
*****************************************************************}
unit Approxim;

interface

uses Windows, Messages, SysUtils, Classes, HEnt1 ;

Type REC_AUTOX = RECORD
                 Nival,NbD,Decim,Temps : integer ;
                 Unique                : boolean ;
                 END ;

Const MaxDroiteX = 1000 ;

type T_DX = Array[0..MaxDroiteX-1] of Double ;
     T_IX = Array[0..MaxDroiteX-1] of Smallint ;

function LettrageApproche ( ALettrer,Delta : double ; Var ListeDroite : T_DX ; Var ListeSortie : T_IX ;
                            Infos : REC_AUTOX ) : integer ;

implementation

Var Encours : T_IX ;
    TotalNeg,TotalPos : T_DX ;
    NbAppels,Debray,Arrete : Comp ;
    TT1 : TDateTime ;

Var Vale : double ;
    Niveau,NbDec,Tempo : integer ;
    NbDroite     : Longint ;
    MtantRappro,DeltaRappro  : Double ;

Const NbDebray : Longint = 3125 ;

Type X_TRI= Array[0..MaxDroiteX-1] of
               RECORD
               PosI : integer ;
               X    : Double ;
               END ;

Var Montant   : X_TRI ;
    SolUnique : boolean ;

{============================================================================}
PROCEDURE POSITIONNEDEBRAYE ;
Var NbS  : Longint ;
    HH,MM,SS,CC : Word ;
BEGIN
DecodeTime(Time-TT1,HH,MM,SS,CC) ;
NbS:=3600*HH+60*MM+SS ;
if NbS>=Tempo then BEGIN Arrete:=1 ; Exit ; END ;
if NbS>0 then
   BEGIN
   Arrete:=Round(NbDebray/NbS*120) ;
   Debray:=Round(NbDebray/NbS*Tempo) ;
   Arrete:=Debray ;
   END else if NbDebray<2000000000 then
   BEGIN
   NbDebray:=2*NbDebray ;
   END else Arrete:=1 ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 04/05/2000
Modifié le ... : 04/05/2000
Description .. : Algo récursif fondamental pour le rappro combinatoire
Mots clefs ... : RAPPROCHEMENT;COMBINATOIRE;RECURSIVITE;
*****************************************************************}
FUNCTION ALGORAPPROCHE ( Reste : double ; Pos,L : integer ) : integer ;
Var ii,Position,Res : integer ;
BEGIN
AlgoRapproche:=0 ;
{gestion du "débrayage"}
NbAppels:=NbAppels+1 ;
if NbAppels=NbDebray then PositionneDebraye ;
if ((Arrete>0) and (NbAppels>Arrete)) then BEGIN AlgoRapproche:=-2 ; Exit ; END ;
if ((Debray>0) and (NbAppels>Debray)) then BEGIN AlgoRapproche:=-2 ; Exit ; END ;
{ Algorithme récursif }
Position:=Pos ;
For ii:=Position to NbDroite-1 do
    BEGIN
    Vale:=Reste+Montant[ii].X ; if EnCours[ii]>0 then Continue ;
    if Abs(Vale-MtantRappro)<=DeltaRappro then BEGIN EnCours[ii]:=1 ; AlgoRapproche:=1 ; Exit ; END ;
    if (Vale+TotalNeg[ii]>MtantRappro-DeltaRappro) or (Vale+TotalPos[ii]<MtantRappro-DeltaRappro) then Continue ;
    if l>=Niveau then Continue ;
    EnCours[ii]:=1 ; Res:=AlgoRapproche(Vale,ii+1,l+1) ;
    if Res<>0 then BEGIN AlgoRapproche:=Res ; Exit ; END ;
    EnCours[ii]:=0 ;
    END ;
END ;

FUNCTION REMPLITNEGPOSAPPROCHE : integer ;
Var i,Niv     : Longint ;
    X,STN,STP : Double ;
BEGIN
REMPLITNEGPOSAPPROCHE:=0 ; STN:=0 ; STP:=0 ; Niv:=0 ;
for i:=NbDroite-1 downto 0 do
    BEGIN
    if Arrondi(Montant[i].X-MtantRappro,NbDec)=0 then BEGIN REMPLITNEGPOSAPPROCHE:=1000+Montant[i].PosI ; Exit ; END ;
    TotalNeg[i]:=STN ; TotalPos[i]:=STP ;
    If Montant[i].X>0 then STP:=STP+Montant[i].X else STN:=STN+Montant[i].X ;
    END ;
if SolUnique then Exit ;
{Peut on espérer une solution ?}
if ((STP>0) and (STP+DeltaRappro<MtantRappro)) then Exit ;
if ((STN<0) and (STN-DeltaRappro>MtantRappro)) then Exit ;
{Somme Totale = Montant}
if Not SolUnique then
   BEGIN
   if Arrondi(STP+STN-MTantRappro,NbDec)=0 then BEGIN REMPLITNEGPOSAPPROCHE:=2 ; Exit ; END ;
   if Arrondi(STP-MTantRappro,NbDec)=0 then BEGIN REMPLITNEGPOSAPPROCHE:=3 ; Exit ; END ;
   if Arrondi(STN-MTantRappro,NbDec)=0 then BEGIN REMPLITNEGPOSAPPROCHE:=4 ; Exit ; END ;
   END ;
{si que des positifs}
if ((STP>0) and (STN=0)) then
   BEGIN
   if MtantRappro<=0 then Exit ; if Montant[NbDroite-1].X>MtantRappro then Exit ;
   i:=NbDroite-1 ; X:=0 ;
   While ((X<=MtantRappro) and (i>=0)) do BEGIN X:=X+Montant[i].X ; Dec(i) ; Inc(Niv) ; END ;
   Dec(Niv,1) ; if ((Niv>=0) and (Niv<Niveau)) then Niveau:=Niv ;
   END ;
{si que des négatifs}
if ((STP=0) and (STN<0)) then
   BEGIN
   if MtantRappro>0 then Exit ; if Montant[0].X<MtantRappro then Exit ;
   i:=0 ; X:=0 ;
   While ((X>=MtantRappro) and (i<=NbDroite-1)) do BEGIN X:=X+Montant[i].X ; Inc(i) ; Inc(Niv) ; END ;
   Dec(Niv,1) ; if ((Niv>=0) and (Niv<Niveau)) then Niveau:=Niv ;
   END ;
REMPLITNEGPOSAPPROCHE:=1 ;
END ;

{============================================================================}
PROCEDURE DEMARRETIME ;
BEGIN
TT1:=Time ; NbAppels:=0 ; Debray:=0 ; Arrete:=0 ;
END ;

{============================================================================}
PROCEDURE TriMontants ( Var LD : T_DX ) ;
Var i,j : Longint ;
    X   : real ;
    Pos : Longint ;
    MM  : X_TRI ;
BEGIN
FillChar(Montant,Sizeof(Montant),#0) ; FillChar(MM,Sizeof(MM),#0) ;
for i:=0 to NbDroite-1 do BEGIN Montant[i].PosI:=i ; Montant[i].X:=LD[i] ; END ;
for i:=NbDroite-1 downto 1 do
    BEGIN
    for j:=0 to i-1 do
        BEGIN
        if Arrondi(Montant[j].X-Montant[i].X,NbDec)>0 then
           BEGIN
           X:=Montant[j].X ; Pos:=Montant[j].PosI ;
           Montant[j].X:=Montant[i].X ; Montant[j].PosI:=Montant[i].PosI ;
           Montant[i].X:=X ; Montant[i].PosI:=Pos ;
           END ;
        END ;
    END ;
for i:=0 to NbDroite-1 do MM[i]:=Montant[NbDroite-1-i] ;
Montant:=MM ;
END ;

function LettrageApproche ( ALettrer,Delta : double ; Var ListeDroite : T_DX ; Var ListeSortie : T_IX ;
                            Infos : REC_AUTOX ) : integer ;
Var Res,i : integer ;
BEGIN
Result:=0 ; NbDroite:=Infos.NbD ; Niveau:=Infos.Nival ;
MtantRappro:=ALettrer ; DeltaRappro:=Delta ;
NbDec:=Infos.Decim ; Tempo:=Infos.Temps ; SolUnique:=Infos.Unique ;
FillChar(TotalNeg,Sizeof(TotalNeg),#0) ; FillChar(TotalPos,Sizeof(TotalPos),#0) ; FillChar(Encours,Sizeof(Encours),#0)   ;
TriMontants(ListeDroite) ; Res:=RemplitNegPosApproche ;
Case Res of
   0 : Exit ; {pas de solution possible}
   1 : BEGIN  {possible solution combinatoire}
       DemarreTime ; Res:=AlgoRapproche(0,0,0) ;
       if Res>0 then
          BEGIN
          for i:=0 to NbDroite-1 do if Encours[i]>0 then ListeSortie[Montant[i].PosI]:=1 ;
          Result:=1 ; Exit ;
          END else if Res=-2 then
          BEGIN
          Result:=2 ; Exit ;
          END ;
       END ;
   2 : BEGIN {solution triviale (Somme liste droite = montant à lettrer)}
       for i:=0 to NbDroite-1 do ListeSortie[i]:=1 ; Result:=1 ; Exit ;
       END ;
   3 : BEGIN {solution triviale (Somme liste positif = montant à lettrer)}
       for i:=0 to NbDroite-1 do if Montant[i].X>0 then ListeSortie[Montant[i].PosI]:=1 ; Result:=1 ; Exit ;
       END ;
   4 : BEGIN {solution triviale (Somme liste neg = montant à lettrer)}
       for i:=0 to NbDroite-1 do if Montant[i].X<0 then ListeSortie[Montant[i].PosI]:=1 ; Result:=1 ; Exit ;
       END ;
  else BEGIN {un élément de liste droite couvre totalement le à lettrer}
       if Res>=1000 then BEGIN ListeSortie[Res-1000]:=1 ; Result:=1 ; Exit ; END ;
       END ;
   END ;
END ;

{============================================================================}
{============================================================================}
END .
