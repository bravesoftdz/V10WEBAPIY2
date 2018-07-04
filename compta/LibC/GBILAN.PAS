{$I-,R-,A-,H-}
UNIT GBILAN ;

interface

USES DIALOGS,SYSUTILS,DB,Classes,{$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},Ent1,HEnt1,GDecladf,Gaccess, Graphics ;

{
1 st:='Comptes'   ;
2 st:='Total'     ;
3 st:='Formule'   ;
4 st:='Constante' ;
5 st:='Texte'     ;
6 Titre colonne
}

Type TCol = Record
            Tipe    : Byte ;
            Texte   : String ;
            Justif  : String50 ;
            Largeur : SmallInt ;
            Formule : String50 ;
            NomCell : String10 ;
            NomRub  : String17 ;
            end ;

Const MaxColH   = 8 ;
      MaxLigneH = 100 ;
      Decal     = 1 ;
      KTitre    = 1 ;

Type PTCOL = ^TCol ;

Type TEtatFi = Record
               Code : String17 ;
               Libelle : String ;
               Signe : Byte ;
               NbL,NbC : SmallInt ;
               Famille : String3 ;
               TitreCol : Array[1..MaxColH] Of String35 ;
               LargeurCol : Array[1..MaxColH] Of Smallint ;
               Ligne   : Array[1..MaxLigneH,1..MaxColH] Of PTCol ;
               end ;


Type PTEtatFi = ^TEtatFi ;
Type TabEtatFi = Array[1..19] Of PTEtatFi ;

{ inutile ...
Type TabEF = Array[1..19] Of PTEFT ;
}

Type TRB = Record
           RB_FAMILLES   : String200 ;
           RB_RUBRIQUE   : String17 ;
           RB_LIBELLE    : String35 ;
           RB_SIGNERUB   : String3 ;
           RB_TYPERUB    : String3 ;
           RB_COMPTE1    : String200 ;
           RB_EXCLUSION1 : String200 ;
           RB_COMPTE2    : String200 ;
           RB_EXCLUSION2 : String200 ;
           RB_AXE        : String3 ;
           end ;
Type PTRB = ^TRB ;
Type TabRB = Array[1..MaxLigneH,1..MaxColH] Of PTRB ;

PROCEDURE CHERCHEBILAN ( Num : Byte  ; EtatFi : PTEtatFi ; Signe : Byte) ;
PROCEDURE TRANSFERTETATFI(EtatFi : PTEtatFi; Var TARB : TabRB) ;

implementation

{=============================================================================}
FUNCTION QUELCALC(b:byte) : String3 ;
BEGIN
Result:='SM' ;
Case b Of 1 : Result:='SD' ; 2 : Result:='SC' ; 3 : Result:='SM' ; 4 : Result:='TD' ;
          5 : Result:='TC' ; END ;
END ;

{=============================================================================}
FUNCTION QUELSIGNE(b : Byte) : String3;
BEGIN
Result:='POS' ; If b=2 Then Result:='NEG' ;
END ;

{=============================================================================}
FUNCTION CODERUBRIQUE(Col : TCol ; l,c : Integer) : String ;
BEGIN
Result:=Col.NomRub+':'+IntToStr(c) ;
END ;

{=============================================================================}
PROCEDURE TRANSFERTETATFI(EtatFi : PTEtatFi; Var TARB : TabRB) ;
Var L,C : Integer ;
    Code,Code1 : String ;
    ColEF : PTCol ;
BEGIN
If EtatFi=NIL Then Exit ;
Fillchar(TARB,SizeOf(TARB),#0) ;
For l:=1 To EtatFi^.Nbl Do
  BEGIN
  Code:='' ;
  If EtatFi^.Ligne[l,1]<>NIL Then Code:=EtatFi^.Ligne[l,1]^.Texte ;
  Code1:='' ;
  If EtatFi^.Ligne[l,2]<>NIL Then Code1:=EtatFi^.Ligne[l,2]^.Texte ;
  For c:=3 To EtatFi^.NbC Do
    BEGIN
    ColEF:=EtatFi^.Ligne[l,c] ;
    If (ColEF<>NIL) And (ColEF^.Tipe In [1]) Then
      BEGIN
      New(TARB[l,c]) ; FillChar(TARB[l,c]^,SizeOf(TARB[l,c]^),#0) ;
      With TARB[l,c]^ Do
        BEGIN
        RB_FAMILLES   :=EtatFi^.Famille+';';
        RB_RUBRIQUE   :=IntToStr(StrToInt(EtatFi^.Famille))+'/'+CodeRubrique(ColEF^,l,c) ;
        RB_LIBELLE    :=Code1+':'+EtatFi^.TitreCol[c]+' '+Code ;
        RB_SIGNERUB   :=QuelSigne(EtatFi^.Signe) ;
        RB_TYPERUB    :='GEN' ;
        RB_COMPTE1    :=ColEF^.Texte ;
        RB_EXCLUSION1 :='';
        RB_COMPTE2    :='';
        RB_EXCLUSION2 :='';
        RB_AXE        :='';
        END ;
      END ;
    END ;
  END ;
END ;

{=============================================================================}
PROCEDURE CHARGELISTEBILAN ( Num : Longint ) ;

Var Fichier : File of Fiche_Liste ;
    Fiche : Fiche_Liste ;

BEGIN
Assign(Fichier,Socpath+'LISTEBIL') ; Reset(Fichier) ;
Seek(Fichier,Num) ; Read(Fichier,Fiche) ; Close(Fichier) ;
VSAA^.CZ:=Fiche.CZ ; VSAA^.PZ:=Fiche.PZ ; VSAA^.LZ:=Fiche.LZ ;
VSAA^.FZ:=Fiche.FZ ; VSAA^.TZ:=Fiche.TZ ; VSAA^.JZ:=Fiche.JZ ;
VSAA^.TitreUnun:=ASCII2ANSI(Trim(Fiche.TitreUnun)) ; VSAA^.TitreListe:=ASCII2ANSI(Trim(Fiche.Titre)) ;
VSAA^.SorteGeneral:=Fiche.Sorte ;
VSAA^.TotZone:=1 ;
While ((VSAA^.TotZone<=ListeMaxZone) and (VSAA^.LZ[VSAA^.TotZone]>0)) do VSAA^.TotZone:=VSAA^.TotZone+1 ;
VSAA^.totzone:=VSAA^.totzone-1 ;
{
NumListe:=Num ;
OldListe ; VerifFormat(OkParam) ; PrepareListe ;
}
END ;

{=============================================================================}
FUNCTION TROUVENOMCELL(Code : String ; Lig,Col : Integer ; EtatFi : PTEtatFi) : String ;
Var i : Integer ;
BEGIN
Result:='A1' ;
For i:=1 To Lig-1 Do
  If Trim(EtatFi^.Ligne[i,1]^.NomRub)=Trim(Code) Then Result:=CHR(64+Col)+IntToStr(i) ;
END ;

{=============================================================================}
Function QuelSens(Plus,TypeF : Byte ) : String40 ;
begin
Result:='' ;
Case TypeF Of
  BilanCpte : Case Plus Of
                1 : Result:='SD' ; 2 : Result:='SC' ; 3 : Result:='SM' ;
                4 : Result:='TD' ; 5 : Result:='TC' ;
                END ;
  BilanRubr : Case Plus Of
                1 : Result:='+' ; 2 : Result:='-' ;
                END ;
  END ;
end ;

{=============================================================================}
Function AnalyseCpt(Cpt1,Cpt2 : String ; Plus : Byte) : String200 ;
Var LeSens : String ;
begin
Result:='' ;
If Cpt1=CPt2 Then Result:=Cpt1 Else Result:=Cpt1+':'+Cpt2 ;
(*
   BEGIN
   i:=1 ; While (i<=VSAA^.ParaSoc.FormatGen.Lg) And (Cpt1[i]=Cpt2[i]) Do Inc(i) ;
   Dec(i) ;
   If i>0 Then Result:=Copy(Cpt1,1,i) Else Result:=Cpt1+':'+Cpt2 ;
   END ;
*)
LeSens:=QuelSens(Plus,BilanCpte) ;
If LeSens<>'' Then Result:=Result+'('+LeSens+')' ;
end ;

{=============================================================================}
Function AnalyseTot(EtatFi : PTEtatFi ; Cpt1,Cpt2 : String ; Plus : Byte ; Lig,Col : Integer) : String200 ;
begin
Result:='Somme('+TrouveNomCell(Cpt1,Lig,Col,EtatFi)+':'+TrouveNomCell(Cpt2,Lig,Col,EtatFi)+')' ;
If Plus=2 Then Result:='-'+Result Else Result:='+'+Result ;
end ;

{=============================================================================}
FUNCTION RECUPFOURCHETTE ( PosFiche : Longint ; TypeF,Lig,Colonne : Byte ; EtatFi : PTEtatFi) : String ;

Var Cle15,Test15 : String40 ;
    Fiche15 : enregCPTA ;
    Pos15,kk : longint ;
    OkOk,PremFois : Boolean ;
    St : String ;
BEGIN
kk:=0 ; St:='' ; PremFois:=TRUE ;
Cle15:=CreCle15(TypeF,PosFiche,0) ; Test15:=Cle15 ;
Repeat
  Inc(kk) ;
  if kk=1 then SearchKey(VSAA^.Idxff[15],Pos15,Cle15) Else NextKey(VSAA^.Idxff[15],Pos15,Cle15) ;
  OkOk:=((VSAA^.Ok) and (Copy(Cle15,1,Length(Test15))<=Test15)) ;
  if OkOk then
     BEGIN
     GetRec(VSAA^.Datff[15],Pos15,Fiche15) ;
     if (Fiche15.ColonneVJ=Colonne) and (Fiche15.TypeEnrVJ=TypeF) and (Fiche15.PosCodeJalVJ=PosFiche) then
        BEGIN
        Case TypeF Of
          BilanCpte : BEGIN
                      If PremFois Then St:=AnalyseCpt(Trim(Fiche15.Compte1VJ),Trim(Fiche15.Compte2VJ),Fiche15.SensVJ)+';'
                                  Else St:=St+AnalyseCpt(Trim(Fiche15.Compte1VJ),Trim(Fiche15.Compte2VJ),Fiche15.SensVJ)+';'
                      END ;
          BilanRubr : BEGIN
                      If PremFois Then
                        St:='='+AnalyseTot(EtatFi,Trim(Fiche15.Compte1VJ),Trim(Fiche15.Compte2VJ),Fiche15.SigneVJ,Lig,Colonne)
                                  Else
                        St:=St+AnalyseTot(EtatFi,Trim(Fiche15.Compte1VJ),Trim(Fiche15.Compte2VJ),Fiche15.SigneVJ,Lig,Colonne) ;
                      END ;
          END ;
        PremFois:=FALSE ;
        END ;
     END ;
Until Not OkOk ;
Result:=St ;
END ;

{=============================================================================}
PROCEDURE INITCOL ( Var Col : PTCol ; NumCol,NumL : SmallInt ; LeNomRub : String) ;
BEGIN
If Col=NIL Then New(Col) ;
With Col^ Do
  BEGIN
  Tipe    :=5 ;
  Texte   :='' ;
  Justif  :=VSAA^.JZ[NumCol] ;
  Largeur :=VSAA^.LZ[NumCol] ;
  Formule :=VSAA^.FZ[NumCol] ;
  NomCell :=CHR(64+NumCol)+IntToStr(NumL+Decal) ;
  NomRub  :=LeNomRub ;
  END ;
END ;

{=============================================================================}
PROCEDURE REAJUSTE(Var St,Blanc : String) ;
Var i,k : Byte ;
    St1 : String35 ;
BEGIN
Blanc:='' ; k:=0 ;
Repeat
  i:=Pos('N',St) ;
  If i>0 Then
     BEGIN
     St[i]:=' ' ; St1:='' ;
     For k:=1 To 2 Do
       BEGIN
       Inc(i) ;
       If i<=Length(St) Then
         BEGIN
         If St[i] In ['0'..'9'] Then BEGIN St1:=St1+St[i] ; St[i]:=' ' ; END ;
         END ;
       END ;
     k:=StrToInt(St1) ;
     END ;
Until i<=0 ;
St1:='' ; For i:=1 To Length(St) Do If St[i]<>' ' Then St1:=St1+ST[i] ;
If k>0 Then
   BEGIN
   Blanc:=Format_String(' ',k) ;
   END ;
St:=St1 ;
END ;

{=============================================================================}
PROCEDURE INITTITRE( k : Integer ; EtatFi : PTEtatFi);
Var Col : SmallInt ;
BEGIN
For Col:=1 To EtatFi^.NbC Do
  BEGIN
  InitCol(EtatFi^.Ligne[k,Col],Col,k,'Titre col '+IntToStr(Col)) ;
  EtatFi^.Ligne[k,Col]^.NomCell :=CHR(64+Col)+IntToStr(k) ;
  EtatFi^.Ligne[k,Col]^.tipe:=6 ;
  If k=kTitre Then EtatFi^.Ligne[k,Col]^.Texte:=EtatFi^.TitreCol[Col] ;
  EtatFi^.Ligne[k,Col]^.Formule:='' ;
  EtatFi^.Ligne[k,Col]^.Justif:='C' ;
  END ;
END ;

{=============================================================================}
PROCEDURE CHERCHEBILAN ( Num : Byte  ; EtatFi : PTEtatFi ; Signe : Byte) ;

Var Cle0,Test0 : String40 ;
    kk,PosBilan : Longint ;
    OkOk : Boolean ;
    Lig,Col : Integer ;
    Fiche0 : TBilan ;
    NbLi,i,k : SmallInt ;
    Justif,Blanc,LeNomRub : String ;
BEGIN
If EtatFi=NIL Then Exit ;
{ Fait Avant ! Fillchar(EtatFi^,SizeOf(EtatFi^),#0) ; }
ChargeListeBilan(Num) ;
{EtatFi^.Code:=VSAA^.TitreUnUn ;} EtatFi^.Libelle:=VSAA^.TitreUnUn ;
If EtatFi^.Libelle='' Then EtatFi^.Libelle:=EtatFi^.Code ;
EtatFi^.NbC:=VSAA^.TotZone ;
For i:=1 To VSAA^.TotZone Do EtatFi^.TitreCol[i]:=ASCII2ANSI(VSAA^.TZ[i]) ;
For i:=1 To VSAA^.TotZone Do EtatFi^.LargeurCol[i]:=VSAA^.LZ[i]*7 ;
EtatFi^.Signe:=Signe ;
EtatFi^.Famille:=FormatFloat('000',Num) ;
OuvreBilan(2) ;
Ouvre(15,2,FALSE,spCPTA) ;
Cle0:=CreCleBilan(Num,'') ; Test0:=Cle0 ; kk:=0 ; Lig:=1+Decal ;
NbLi:=Decal ;
For k:=1 To Decal Do InitTitre(k,EtatFi) ;
Repeat
  Inc(kk) ;
  if kk=1 then SearchKey(VSAA^.IdxfTva,PosBilan,Cle0) else NextKey(VSAA^.IdxFTva,PosBilan,Cle0) ;
  OkOk:=((VSAA^.Ok) and (Copy(Cle0,1,Length(Test0))=Test0)) ;
  if OkOk then
     BEGIN
     GetRec(VSAA^.DatFTva,PosBilan,Fiche0) ;
     Inc(NbLi) ;
     For Col:=1 To EtatFi^.NbC Do
       BEGIN
       Blanc:='' ;
       If Col=1 Then LeNomRub:='' Else LeNomRub:=EtatFi^.Ligne[Lig,1]^.Texte ;
       InitCol(EtatFi^.Ligne[Lig,Col],Col,Lig,LeNomRub) ;
       Justif:=Trim(Fiche0.JZ[Col]) ; If Justif='' Then Justif:=Trim(VSAA^.JZ[Col]) ;
       If Col<=2 Then
          BEGIN
          EtatFi^.Ligne[Lig,Col]^.Tipe:=5 ;
          If Col=1 Then
             BEGIN
             EtatFi^.Ligne[Lig,Col]^.Texte:=ASCII2ANSI(Fiche0.Code) ;
             EtatFi^.Ligne[Lig,Col]^.NomRub:=EtatFi^.Ligne[Lig,Col]^.Texte ;
             END Else EtatFi^.Ligne[Lig,Col]^.Texte:=ASCII2ANSI(Fiche0.Intitule) ;
          EtatFi^.Ligne[Lig,Col]^.Justif:=Justif ;
          END Else
          BEGIN
          EtatFi^.Ligne[Lig,Col]^.Tipe:=Fiche0.Colonnes[Col-2].Tipe ;
          Case Fiche0.Colonnes[Col-2].Tipe Of
            1 : BEGIN
                EtatFi^.Ligne[Lig,Col]^.Texte:=RecupFourchette(PosBilan,BilanCpte,Lig+Decal,Col,EtatFi) ;
                If Justif='' Then Justif:='D/2' ;
                END ;
            2 : BEGIN
                EtatFi^.Ligne[Lig,Col]^.Texte:=RecupFourchette(PosBilan,BilanRubr,Lig+Decal,Col,EtatFi) ;
                If Justif='' Then Justif:='D/2' ;
                END ;
            Else
                BEGIN
                EtatFi^.Ligne[Lig,Col]^.Texte:=Fiche0.Colonnes[Col-2].Valeur ;
                END ;
            END ;
          END ;
       Reajuste(Justif,Blanc) ;
       EtatFi^.Ligne[Lig,Col]^.Justif:=Justif ;

       If Blanc<>'' Then If EtatFi^.Ligne[Lig,Col]^.Tipe=5 Then
          EtatFi^.Ligne[Lig,Col]^.Texte:=Blanc+EtatFi^.Ligne[Lig,Col]^.Texte ;
       END ;
     Inc(Lig) ;
     END ;
Until Not OkOk ;
Ferme(15,2,FALSE) ;
FermeBilan(2) ;
EtatFi^.NbL:=NbLi ;
END ;


end.
{$A+,H+}
