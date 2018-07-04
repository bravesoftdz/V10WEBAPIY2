{$A-,H-}

unit ImpCptaU;

interface
uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Forms, Dialogs, StdCtrls, Spin, DB, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, Hctrls, FileCtrl,GACCESS,
     ExtCtrls,Ent1, HStatus, ComCtrls, HEnt1, GDeclaDf, SaisUtil, Buttons, HCompte,
     hmsgbox, Cpteutil, GBilan ;

Type TParLegalV8 = RECORD
                   LValide : Char ;
                   LDate : Word ;
                   LHeure : Word ;
                   LNumUser : Byte ;
                   LPage : SmallInt ;
                   LNoEdition : Byte ;
                   LFait : Char ;
                   END ;

Type TPerLegalV8 = Array[1..24] of TParLegalV8 ; // 24 périodes

FUNCTION TROUVEINDEXRUPT ( Sorte,NoRupt : SmallInt  ) : String ;
PROCEDURE LITSTRUCTURETRANCHEV8(i : SmallInt ; Var Deb,Lg : SmallInt) ;
function DecodeGuide(i : Byte ; St : shortstring ; Var Arret : String) : shortstring ;
PROCEDURE LITLEGALPERIODEV8 ( Lequel : SmallInt ; Var Legal : TPerLegalV8 ;
                              NoExo : SmallInt ; OkOuvre : boolean ) ;

implementation
Const ReelMaxRupt = 7 ;
      LigneMaxRupt = 13 ;


Type RecTabA = Array[1..ReelMaxRupt] of real48 ;

Type TabLig = Array[1..2,1..20] of SmallInt ;

Type TypeLigneAff = RECORD
                    NoTable,NbrChar,OldNbrChar,Posdeb,PosFin : Byte ;
                    Result : string50  ;
                    END ;

Type TabLigneAff = Array [1..LigneMaxRupt] of TypeLigneAff ;

Type RecligneAff = RECORD
                   NoTable : byte ;
                   Posdeb  : byte ;
                   PosFin  : byte ;
                   END ;

Type RecCodeAff = RECORD
                  Long : byte ;
                  Tab : Array[1..LigneMaxRupt] of RecLigneAff ;
                  END ;

Type TNomTable = ARRAY[1..19] of string10 ;

{============================================================================}
PROCEDURE INITLIGNE ( Var LC : TypeLigneAff ) ;

BEGIN
With LC do
  BEGIN
  NbrChar:=0 ; OldNbrChar:=0 ; PosDeb := 0 ; PosFin := 0 ;
  Result:='' ; NoTable:=0 ;
  END ;
END ;

{============================================================================}
PROCEDURE DECODE ( Var long,Sorte : SmallInt ; Tab : RecTabA ; Var LC : TabligneAff ) ;

Var XRecord : RecCodeAff ;
    i : SmallInt ;
    OffSet : longint ;
    st1 : string13 ;
    NomTable : TNomTable ;

BEGIN
OuvreChoixCod ;
Move(Tab,XRecord,40) ; Long:=XRecord.long ;
If Sorte=2 then OffSet:=2710 else OffSet:=10850 ;
Seek(VSAA^.FichierChCh,OffSet) ;
For i:= 1 to 19 do
    BEGIN
    Read(VSAA^.FichierChCh,St1) ;
    NomTable[i]:=Format_String(St1,10) ;
    END ;
For i:=1 to LigneMaxRupt do
    BEGIN
    InitLigne(lc[i])  ;
    LC[i].NoTable:=XRecord.Tab[i].NoTable ;
    LC[i].PosDeb:=XRecord.Tab[i].Posdeb ;
    LC[i].PosFin:=XRecord.Tab[i].PosFin ;
    With LC[i] do
      if (PosDeb=0) then LC[i].NbrChar:=0
                    else LC[i].NbrChar:=XRecord.Tab[i].PosFin-XRecord.Tab[i].Posdeb+1 ;
    LC[i].OldNbrChar:=LC[i].NbrChar ;
    LC[i].Result:= '' ;
    if LC[i].NoTable<>0 then
       BEGIN
       LC[i].Result:=NomTable[LC[i].NoTable-160*Ord(Sorte=1)] ;
       END ;
    END ;
FermeChoixCod ;
END ;

{============================================================================}
PROCEDURE LITFICHIERTAN ( Var Long,Sorte : SmallInt ; Var Tab : RecTabA ; Var LC : TabLigneAff ) ;

Var i : SmallInt ;
    OffSet : Longint ;

BEGIN
If Sorte=2 then OffSet:=110 else OffSet:=150 ;
OuvreNum ; Seek(VSAA^.FichierNum,OffSet) ;
For i:=1 to ReelMaxRupt do Read(VSAA^.FichierNum,Tab[i]) ;
FermeNum ;
Decode(Long,Sorte,Tab,LC) ;
END ;

{============================================================================}
PROCEDURE LITSTRUCTURETRANCHEV8(i : SmallInt ; Var Deb,Lg : SmallInt) ;
Var Long,Sorte : SmallInt ;
    Tab : RecTabA ;
    LC : TabLigneAff ;
BEGIN
Sorte:=2 ; Deb:=0 ; Lg:=0 ;
LitFichierTan(Long,Sorte,Tab,LC) ;
Deb:=LC[i].PosDeb ; Lg:=LC[i].NbrChar ;
END ;

{============================================================================}
PROCEDURE RECUPLECLASSEMENT ( Sorte,NoTable : SmallInt ; Var NbTranches : SmallInt ;
                              Var TabChoix : TabLig ) ;

Var Cht : SmallInt ;
    i : byte ;
    St : Array[1..7] of String13 ;
    OffSet : Longint ;

BEGIN
Case Sorte of 1 : Cht:=12710 ; 2 : Cht:=12860 ; else exit ; END ;
OffSet:=Cht+8*(NoTable-2) ;
NbTranches:=0 ; FillChar(TabChoix,SizeOf(TabChoix),#0) ;
OuvreChoixCod ; Seek(VSAA^.FichierChCh,OffSet) ;
For i:=1 to 7 do Read(VSAA^.FichierChCh,St[i]) ;
Move(St[1][1],TabChoix,SizeOf(TabChoix)) ;
Read(VSAA^.FichierChCh,St[1]) ; NbTranches:=StrToInt(St[1]) ;
FermeChoixCod ;
END ;


{============================================================================}
FUNCTION CHOIXDANSTABLE ( Choix : SmallInt ; Table : TabLig ) : Boolean ;

Var i : SmallInt ;
    Good : Boolean ;

BEGIN
i:=0 ;
Repeat
  Inc(i) ; Good:=(Choix=Table[1,i]) ;
Until ((Good) or (i>=20)) ;
ChoixDansTable:=Good ;
END ;

{============================================================================}
FUNCTION OKLC ( Var LCL : TypeLigneAff ; Var TabChoix : TabLig ) : Boolean ;

Var Good : Boolean ;

BEGIN
Good:=(Trim(LCL.Result)<>'') and (LCL.NoTable<>0) ;
If Good then Good:=Not ChoixDansTable(LCL.NoTable,TabChoix) ;
OkLC:=Good ;
END ;

{============================================================================}
FUNCTION INT2CHOIX (typ,mark,flag : SmallInt) : string13 ;
var pos,Offset,PosSeek : Longint ;
    st : string13 ;
BEGIN
PosSeek:=FilePos(VSAA^.FichierChCh) ;
(*
OffSet:=desifich(flag mod 500) ;
*)
OffSet:=Flag ;
st:='' ; if flag>4 then typ:=1 ; Pos:=mark-1+OffSet ;
Seek(VSAA^.FichierChCh,pos) ; Read(VSAA^.FichierChCh,st) ;
Int2choix:=st ;
Seek(VSAA^.FichierChCh,posSeek) ;
END ;


{============================================================================}
PROCEDURE INITTRANCHES ( Var LC : TabLigneAff ; Var NbTranches,Sorte : SmallInt ;
                         Var TabDispo,TabChoix : TabLig ; OkPremier : Boolean ) ;

Var Ligne,NbT,Cht : SmallInt ;
    St : String13 ;
    OffSetL : Longint ;

BEGIN
If Sorte=2 then BEGIN OffSetL:=10750 ; Cht:=2710  ; END
           else BEGIN OffSetL:=10900 ; Cht:=10850 ; END ;
NbT:=0 ; FillChar(TabDispo,SizeOf(TabDispo),#0) ;
OuvreChoixCod ;
Seek(VSAA^.FichierCHCH,OffSetL) ; St:='' ; For Ligne:=1 to 50 do Write(VSAA^.FichierCHCH,St) ;
For Ligne:=1 to LigneMaxRupt do If OkLC(LC[Ligne],TabChoix) then
    BEGIN
    St:=Int2Choix(1,LC[Ligne].NoTable-160*Ord(Sorte=1),Cht) ;
    Inc(NbT) ;

    Seek(VSAA^.FichierCHCH,OffSetL+NbT-1) ; Write(VSAA^.FichierCHCH,St) ;
    TabDispo[1,NbT]:=LC[Ligne].NoTable ; TabDispo[2,NbT]:=Ligne ;
    END ;
FermeChoixCod ;
If OkPremier then NbTranches:=NbT ;
END ;


{=============================================================================}
FUNCTION ORDONNE ( Sorte,NbTranches : SmallInt ; Var LC : TabLigneAff ;
                   Var TabChoix : TabLig ) : Boolean ;

Var Good,Fin : Boolean ;
    TabDispo,Tab2 : TabLig ;
    NbT2,i : SmallInt ;
BEGIN
If Lc[1].NoTable=0 Then i:=1 ;
FillChar(TabDispo,SizeOf(TabDispo),#0) ; FillChar(Tab2,SizeOf(Tab2),#0) ; NbT2:=0 ;
InitTranches(LC,NbT2,Sorte,TabDispo,Tab2,TRUE) ; i:=1 ; Good:=TRUE ; Fin:=FALSE ;
While (Not Fin) do
  BEGIN
  If (TabDispo[1,i]>0) and (TabChoix[1,i]>0) then Good:=(TabDispo[1,i]=TabChoix[1,i]) ;
  Fin:=(Not Good) or (i>=20) or (TabDispo[1,i]<=0) or (TabChoix[1,i]<=0) ; Inc(i) ;
  END ;
Ordonne:=Good ;
END ;

{============================================================================}
FUNCTION TROUVEINDEXRUPT ( Sorte,NoRupt : SmallInt  ) : String ;

Var NbTranches,Long : SmallInt ;
    TabChoix : TabLig ;
    Tab : RecTabA ;
    LC : TabLigneAff ;
    St : String ;
    i : SmallInt ;
BEGIN
St:='' ;
if ((Sorte<>0) and (NoRupt>1)) then
   BEGIN
   RecupLeClassement(Sorte,NoRupt,NbTranches,TabChoix) ;
   if NbTranches>0 then
      BEGIN
      LitFichierTan(Long,Sorte,Tab,LC) ;
      if Not Ordonne(Sorte,NbTranches,LC,TabChoix) then ;
      i:=1 ;
      While TabChoix[1,i]<>0 Do BEGIN St:=St+FormatFloat('000',TabChoix[1,i])+';' ; Inc(i) ; END ;
      END ;
   END ;
Result:=St ;
END ;

{============================================================================}
PROCEDURE TROUVEVARIABLE ( Var Stt : shortstring ; Var Cpar,Opar,Num : SmallInt ;
                           s : String1) ;
Var Lg : Byte ;
    St,St1 : shortstring ;
  Label 0 ;
BEGIN
St:=Stt ; St1:='' ;
Num:=-1 ; CPAR:=0 ; OPAR:=pos(s,st) ; If OPAR=0 Then Goto 0 ;
Lg:=1 ;
If (Length(St)>=Opar+2) And (St[OPAR+2] in ['0'..'9']) Then
   BEGIN
   Lg:=2 ;
   If (Length(St)>=Opar+3) And (St[OPAR+3] in ['0'..'9']) Then Lg:=3 ;
   END ;
If OPAR>=Length(St) then Goto 0 ; St1:=Copy(St,OPAR+1,Lg) ;
Num:=StrToInt(St1) ; CPAR:=OPAR+Lg ;
0:Stt:=St ;
END ;

{============================================================================}
PROCEDURE VARIABLEFORMULE ( Var Stt : shortstring ; St1 : String1 ) ;

Var Num,OPAR,CPAR,i : SmallInt ;
    St2 : String10 ;
    St : shortstring ;
    Tempo : Byte ;
  Label 0,1 ;

BEGIN
Tempo:=0 ; St:=Stt ;
Repeat
  TrouveVariable(St,Cpar,Opar,Num,St1) ; If CPAR=0 then Goto 0 ;
  if ((Num>=1) and (Num<=99)) then
     BEGIN
     Case St1[1] Of
       'C' : BEGIN
             st2:='1' ;
             St:=Copy(St,1,OPAR-2)+St2+Copy(St,CPAR+1,length(St)-CPAR) ;

             (*
             No:=113 ; Cle:=CreCle16(1,No,Cle,'','',0) ;
             If TrouveLaFiche1(16,Cle,GTest^,LaPos,TRUE) Then
                BEGIN
                St2:=StrF(GTest^.TauxTvaTA,0,4{OkDecV}) ;
                St:=Copy(St,1,OPAR-2)+St2+Copy(St,CPAR+1,length(St)-CPAR) ;
                END else
                BEGIN
                AlertGuideFaux(3371) ; Exit ;
                END ;
             *)
             END ;
       'T' : BEGIN
             St2:='[tVA]' ;
             St:=Copy(St,1,OPAR-2)+St2+Copy(St,CPAR+1,length(St)-CPAR) ;
             END ;
       Else Goto 0 ;
       END ;
     END ;
  Inc(Tempo) ;
Until (Tempo=100) ;
0:
If St1[1]='T' Then For i:=1 To length(st) Do If St[i]='t' Then St[i]:='T' ;
Stt:=St ;
END ;

{============================================================================}
PROCEDURE VARIABLEFORMULE1 ( Var Stt : shortstring ) ;
var Num,OPAR,CPAR : SmallInt ;
    St : shortstring ;
    St2 : shortstring ;
    Tempo : Byte ;
  Label 0 ;
BEGIN
Tempo:=0 ; St:=Stt ;
repeat
  TrouveVariable(St,Cpar,Opar,Num,'$') ; If CPAR=0 Then Goto 0 ;
  if ((Num>=1) and (Num<=999)) then
     BEGIN
     St2:='[E_DEBIT:L'+IntToStr(Num)+']+[E_CREDIT:L'+IntToStr(Num)+']' ;
     St:=Copy(St,1,OPAR-1)+St2+Copy(St,CPAR+1,length(St)-CPAR) ;
     END ;
  Inc(Tempo) ;
Until Tempo=100 ;
0:Stt:=St ;
END ;

{============================================================================}
PROCEDURE VARIABLEFORMULE2 (Var Stt : shortstring) ;
var OPAR : SmallInt ;
    Lg : Byte ;
    Tempo : Byte ;
    St : shortstring ;
  Label 0 ;
BEGIN
St:=Stt ;
Tempo:=0 ;
repeat
  OPAR:=Pos('%',St) ; If OPAR=0 Then Goto 0 ;
  Lg:=Length(St) ;
  St:=Copy(St,1,Opar-1)+'/100'+Copy(St,Opar+1,Lg-Opar+1) ;
  Inc(Tempo) ;
Until Tempo=100 ;
0:Stt:=St ;
END ;

{============================================================================}
PROCEDURE NUMSTRING (Var Stt : shortstring) ;
var st1 : shortstring ;
    St : shortstring ;
    Ope : Array [1..100] of SmallInt ;
    Find : Boolean ;
    i,imax,plus : SmallInt ;
    Tempo : Byte ;
  Label 0 ;
BEGIN
Tempo:=0 ; St:=Stt ;
repeat
   imax:=1 ; Find:=FALSE ;
   For i:=1 to Length(St) do
      BEGIN
      IF ((St[i] in ['0'..'9','.',','])=false) and (i>1) Then
         BEGIN Find:=TRUE ; Ope[imax]:=i ; imax:=imax+1 ; END ;
      END ;
   If ((Find=FALSE) or (imax=1)) Then Goto 0 ;
   OPE[imax]:=Length(st)+1 ; i:=1 ;
   if OPE[i+1]=OPE[i]+1 then Plus:=2 else Plus:=1 ;
   St1:=Copy(St,1,OPE[i+plus]-1) ;
   {
   CalcForm(St1,OPE[i]) ;
   }
   St:=St1+Copy(St,OPE[i+plus],length(St)-OPE[i+plus]+1) ;
   Inc(Tempo) ;
Until Tempo=100 ;
0:Stt:=St ;
END ;

(*
{============================================================================}
PROCEDURE NUMFORMULE ( Var St : String) ;
var OPAR,CPAR : integer ;
    Find : Boolean ;
    St0 : String ;
    Tempo : byte ;
BEGIN
Tempo:=0 ;
repeat
  CPAR:=pos(')',st) ; If CPAR=0 Then Exit ;
  OPAR:=CPAR ; Find:=FALSE ;
  While ((OPAR>0) and (Find=FALSE)) do
     If St[Opar]='(' Then Find:=TRUE else OPAR:=OPAR-1 ;
  St0:=Copy(St,OPAR+1,(CPAR-OPAR)-1) ;
  NUMString(St0) ;
  St:=Copy(St,1,OPAR-1)+St0+Copy(St,CPAR+1,length(St)-CPAR) ;
  Inc(Tempo) ;
Until Tempo=100 ;
END ;
*)

{============================================================================}
PROCEDURE GFORMULE1 ( Var Stt : shortstring ) ;
Var st : shortstring ;
BEGIN
St:=Stt ;
If Trim(St)<>'' then
   BEGIN
   VariableFormule(St,'C') ;
   VariableFormule(St,'T') ;
   VariableFormule1(St) ;
   VariableFormule2(St) ;
   {
   NumFormule(St) ;
   NumString(St)  ;
   }
   END ;
Stt:=St ;
END ;

{=============================================================================}
procedure Remplace(Var St : String ; C1,C2 : Char) ;
Var  l : byte ;
begin
For l:=1 To Length(St) Do If St[l]=C1 Then St[l]:=C2 ; St:=Trim(St) ;
end ;

{============================================================================}
PROCEDURE NUMFORMULE ( Var Stt : String ) ;

Var OPAR,CPAR : Smallint ;
    Find : Boolean ;
    St0,St1,SDeb,Sfin,SLg : String ;
    i,Deb,Fin,l,l2,FormatDate,Lg : SmallInt ;
    Tempo : Smallint ;
    St : String ;
  Label 0 ;
BEGIN
Tempo:=0 ; St:=Stt ;
repeat
  CPAR:=pos(']',st) ; If CPAR=0 Then Goto 0 ;
  SLg:='' ;
  OPAR:=CPAR ; Find:=FALSE ;
  While ((OPAR>0) and (Find=FALSE)) do
     If St[Opar]='[' Then Find:=TRUE else OPAR:=OPAR-1 ;
  St0:=Copy(St,OPAR+1,(CPAR-OPAR)-1) ; St1:='' ;
  i:=StrToInt(Copy(St0,1,2)) ;
  If St0[5] in ['0'..'9'] Then l:=2 Else l:=1 ;
  SDeb:=Copy(St0,4,l) ;
  deb:=StrToInt(SDeb) ;
  If St0[4+l+2] in ['0'..'9'] Then l2:=2 Else l2:=1 ;
  SFin:=Copy(St0,4+l+1,l2) ;
  Fin:=StrToInt(SFin) ; FormatDate:=1 ;
  If i=4 Then FormatDate:=StrToInt(Copy(St0,l+4+1+l2+1,1)) ;
  SLg:='"'+Trim(SDeb)+','+Trim(SFin)+'"' ;
  Case i Of
    1 : St1:='<'+SLg+'E_NUMEROPIECE>' ; {StrN(OldEntete.NoPiece,Mini(Fin-Deb+1,5)) ;}
    2 : St1:='<'+SLg+'E_JOURNAL>' ; {OldEntete.Jal ;}
    3 : St1:='<'+SLg+'J_INTITULE>' ; {FicheListe[3]^.IntituleJ ;}
    4 : BEGIN
        Case FormatDate Of
          1 : St1:='"dd-mm-yy"' ;
          2 : St1:='"dd-mm-yyyy"' ;
          3 : St1:='"dd"' ;
          4 : St1:='"mm"' ;
          5 : St1:='"mmmm yyyy"' ;
          6 : St1:='"yy"' ;
          7 : St1:='"yyyy"' ;
          END ;
        St1:='<'+St1+'E_DATECOMPTABLE>' ;
        END ;
    5 : St1:='<'+SLg+'E_REFINTERNE>' ;
    6 : St1:='<'+SLg+'E_GENERAL>' ;
    7 : St1:='<'+SLg+'G_LIBELLE>' ;
    8 : St1:='<'+SLg+'E_AUXILIAIRE>' ;
    9 : St1:='<'+SLg+'G_LIBELLE>' ;
    10 : St1:='<'+SLg+'<INTITULE>' ;
    END ;
  (*
  if Fin>length(St1) then Fin:=length(St1) ;
  St0:=Trim(Copy(St1,Deb,Fin-Deb+1)) ;
  *)
  St0:=St1 ;
  if OPAR>1 then St1:=Copy(St,1,OPAR-1) else St1:='' ;
  Lg:=Length(St) ;
  {
  St:=St1+St0+Copy(St,CPAR+1,length(St)-CPAR) ;
  }
  St:=St1+St0+Copy(St,CPAR+1,Lg-CPAR) ;
  Inc(Tempo) ;
Until (Tempo=100) ;
0:Remplace(St,'<','[') ; Remplace(St,'>',']') ;
stt:=st ;
END ;

{=============================================================================}
procedure Enleve(Var St : String ; C : Char) ;
Var  l : byte ;
begin
For l:=1 To Length(St) Do If St[l]=C Then St[l]:=' ' ; St:=Trim(St) ;
end ;

{=============================================================================}
function DecodeGuide(i : Byte ; St : shortstring ; Var Arret : String) : shortstring ;
Var St1,St2 : shortstring ;
    st200 : shortstring ;
    l : Byte ;
begin
Result:='' ; St:=Trim(St) ;
Case i Of
  1 : BEGIN
      If Copy(St,1,2)='CA' Then
         BEGIN
         If St[3]='*' Then St1:='1' Else St1:=Copy(St,3,Length(St)-2) ;
         Result:='[AUTO'+St1+']' ;
         END Else If St[1]='/' Then Result:='' Else BEGIN Enleve(St,'*') ; Result:=St ; END ;
      Arret[i]:='X' ;
      END ;
  2 : BEGIN
      If St='/' Then Result:='' Else
         BEGIN
         If St[1]='/' Then Enleve(St,'/') ;
         Enleve(St,'*') ; Result:=St ;
         END ;
      Arret[i]:='X' ;
      END ;
  3 : BEGIN
      If Pos('!',St)>0 Then Arret[i]:='X' Else Arret[i]:='-' ;
      If Copy(St,1,2)='£R' Then St:='' Else
         If St[1]='£' Then
            BEGIN
            St2:='' ; For l:=2 TO Length(St) Do If St[l] In ['0'..'9'] Then St2:=St2+St[l] ;
            St:='[E_REFINTERNE:L'+St2+']' ;
            END ;
      Enleve(St,'!') ; Result:=St ;
      END ;
  4 : BEGIN
      If Pos('!',St)>0 Then Arret[i]:='X' Else Arret[i]:='-' ;
      NumFormule(St) ;
      If Copy(St,1,2)='£T' Then St:='' Else
         If St[1]='£' Then
            BEGIN
            St2:='' ; For l:=2 TO Length(St) Do If St[l] In ['0'..'9'] Then St2:=St2+St[l] ;
            St:='[E_LIBELLE:L'+St2+']' ;
            END ;
      Enleve(St,'!') ; Result:=St ;
      END ;
  5,6 : BEGIN
        If Pos('!',St)>0 Then Arret[i]:='X' Else Arret[i]:='-' ;
        If St<>'' Then Arret[i]:='X' ;
        If St='SLDE' Then St:='[SOLDE]' Else If St[1]='=' Then
           BEGIN
           st:=copy(St,2,Length(St)-1) ; St200:=St ;
           GFormule1(St200) ; St:=St200 ;
           END ;
        Enleve(St,'!') ; Result:=St ;
        END ;
  END ;
end ;


{============================================================================}
PROCEDURE LITLEGALPERIODEV8 ( Lequel : SmallInt ; Var Legal : TPerLegalV8 ;
                              NoExo : SmallInt ; OkOuvre : boolean ) ;

Const OffsetLecture = 3000 ;
      OffsetDecalEd = 4800 ;
      OffsetDecalAn = 240 ;
      OffSetDecalMois = 10 ;
      ReelMax = 40 ;
Type RecTab = Array[1..ReelMax] of real48 ;

Var i : SmallInt ;
    Tab : RecTab ;

BEGIN
FillChar(Legal,SizeOf(Legal),#0) ;
if OkOuvre then OuvreNum ;
Seek(VSAA^.FichierNum,OffsetLecture+(Lequel*OffsetDecalEd)+OffsetDecalAn*(NoExo-1)) ;
For i:=1 to ReelMax do Read(VSAA^.FichierNum,Tab[i]) ;
Move(Tab,Legal,SizeOf(Legal)) ;
if OkOuvre then FermeNum ;
END ;



end.
{$A+,H+}
