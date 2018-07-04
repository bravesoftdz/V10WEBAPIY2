unit QRRupt;

interface

uses classes,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    {$IFDEF VER150}
     variants,
    {$ENDIF}
     HCtrls,DB,Dialogs, SysUtils, HEnt1, Graphics, Forms, tCalcCum,Ent1  ,UentCommun ;

(*=======================================================================================
======================== POUR FAIRE DES RUPTURES COMPTABLES =============================
Dans la forme de l'état :

1) déclarer un L : TStringList dans la partie private (NE PAS L'INITIALISER !!)
2) Creer une Bande de type rbSubDetail Nommé BDL avec des QRLabel ex Cod,Lib,Tot1,Tot2
3) Creer un QRDetailLink Nommé QRDL avec comme Propriété
     Master      = QR : le QuickReport
     DetailBand  = BDL
     PrintBefore = TRUE
     IsRupture   = TRUE

4) Sur le click du bouton print faire

    ChargeRupt(L,'RUG','001',{'06x','75x'} ou {'',''}) ; (fourchettes de code ruptures)
    ..... code d'impression ex : QR.Preview ; ....
    VideRupt(L) ;

5) Sur le BeforePrint de la Bande Detail faire :

   AddRupt(L,TGeneG_GENERAL.AsString,[1,TGeneG_TOTALDEBIT.AsFloat,0,0,0,0]) ;

   (Rq : pour n'imprimer que les ruptures : PrintBand:=FALSE ;)

6) Sur le OnNeedData du QRDL mettre :

    Var Cod,Lib : String ;
    Tot : Array [0..12] of Double ;
    begin
    MoreData:=PrintRupt(L,TGeneG_GENERAL.AsString,Cod,Lib,PourTotalGen,QR.EnRupture,Tot) ;
    RG.Caption:=Cod ; RL.Caption:=Lib ;
    RT1.Caption:=FloatToStr(Tot[0]) ; RT2.Caption:=FormatFloat('#,##0.00',Tot[1]) ;
    END ;

========================================================================================
======================== POUR FAIRE DES GROUPES / RUPTURES =============================
Dans la forme de l'état :

1) déclarer un L : TStringList dans la partie private (NE PAS L'INITIALISER !!)
2) Creer une Bande de type rbSubDetail Nommé BDL avec des QRLabel ex Cod,Lib,Tot1,Tot2
3) Creer un QRDetailLink Nommé QRDL avec comme Propriété
     Master      = QR : le QuickReport ou QRD le subdetail
     DetailBand  = BDL
     PrintBefore = FALSE

4) Sur le click du bouton print faire

    ChargeGroup(L,['MOIS','NUM']) ; mettre autant de string que de groupes
    ..... code d'impression ex : QR.Preview ; ....
    VideGroup(L) ;

5) Sur le BeforePrint de la Bande Detail faire :

   AddGroup(L,[EE_MOIS,EE_NUMEROPIECE],[1,EE_DEBIT.AsFloat,EE_CREDIT.AsFloat]) ;

   (Rq : pour n'imprimer que les ruptures : PrintBand:=FALSE ;)

6) Sur le OnNeedData du QRDL mettre :

   Var Cod,Lib : String ;
       Tot : Array[0..12] of Double ;
   begin
   MoreData:=PrintGroup(L,E,[EE_MOIS,EE_NUMEROPIECE],Cod,Lib,Tot) ; E = Query de groupe
   CodR.Caption:=Cod ; LibR.Caption:=Lib ;
   CC.Caption:=FormatFloat('#,##0',Tot[0]) ;
   DR.Caption:=FormatFloat('#,##0.00',Tot[1]) ;
   CR.Caption:=FormatFloat('#,##0.00',Tot[2]) ;

========================================================================================*)

Const LgSpecif1 : Integer = 3 ;

type  TOR = Class
       Filled    : Boolean ;
       Printed   : Boolean ;
       Principal : Boolean ;
       Libelle   : String ;
       NbChar    : Integer ;
       Tot       : Array[0..77] of Double ; //25 à 77
       END ;

Type TRuptInf = Record
                CodRupt,LibRupt : String ;
                End ;

procedure ChargeRupt (Var L : TStringList ; Code,Plan,Code1,Code2 : String) ;
procedure ChargeRuptDevise (Var L : TStringList) ;
procedure AddRupt (L : TStringList ; St : String ; tot : Array of double) ;
procedure AddRuptCorres (L : TStringList ; St : String ; Tot : Array of double) ;
Function  PrintRupt (L : TStringList ; St : String ; Var Cod,Lib : string ; Var PourTotalGen : Boolean ; EnRupture : Boolean ; Var Tot : array of Double) : boolean ;
Function PrintRupt2 (L : TStringList ; St : String ; Var Cod,Lib : string ; Var PourTotalGen : Boolean ; EnRupture : Boolean ; Var Tot : array of Double ;
                     F : Array of TField ; Q : TQuery = NIL; STTS : String = '') : boolean ;
(*
Function PrintRupt2 (L : TStringList ; St : String ; Var Cod,Lib : string ; Var PourTotalGen : Boolean ; EnRupture : Boolean ; Var Tot : array of Double ;
                     E : TQuery ; F : Array of TField) : boolean ;
*)
procedure VideRupt (Var L : TStringList ) ;
procedure ReinitRupt (L : TStringList) ;
procedure NiveauRupt (Var L : TStringList) ;
function  PlusDeUneRuptureSurCetteRacine(L : TStringList ; St : String) : Boolean ;

procedure ChargeRuptCorresp (Var L : TStringList ; Lequel,Code1,Code2 : String; CodLib : Boolean) ;//Lequel=GE1 ou A52 etc.. Code1 rt Code2 bornes

procedure ChargeGroup (Var L : TStringList ; Lib : Array of String) ;
procedure AddGroupLibre (L : TStringList ; Q : TQuery ; fb : TfichierBase ; Ftri : String ; Tot : Array of double) ;
procedure AddGroup (L : TStringList ; F : Array of TField ; Tot : Array of double) ;
Function PrintGroup (L : TStringList ; E : TQuery ; F : Array of TField ; Var Cod,Lib : string ; Var Tot : array of Double ; Var NumR : Integer ; STTS : String = '') : boolean ;
Function PrintGroupLibre (L : TStringList ; Q : TQuery ; fb : TfichierBase ; Ftri : String ;
                          Var Cod,Lib,Lib1 : string ; Var Tot : array of Double ;
                          Var NumR : Integer ; Var Col : TColor ; Var LibRuptInf : Array Of TRuptInf  ; STTS : String = '') : boolean ;
procedure VideGroup (Var L : TStringList ) ;

procedure ChargeRecap (Var L : TStringList ) ;
//procedure AddRecap (L : TStringList ; Cod,Lib : String ; Tot : Array of double) ;
procedure AddRecap (L : TStringList ; F : Array Of String ; FL : Array Of String ; Tot : Array of double) ;
Function  PrintRecap (L : TStringList ; Var Cod,Lib : string ; Var Tot : array of Double) : boolean ;
procedure VideRecap (Var L : TStringList ) ;

{ Pour Edtions qui gérent les Ruptures avec des comptes associés }
Function  DansRupt (L : TStringList ; St : String) : Boolean ;
Function  DansLibre (F : Array of TField ; L1, L2, Trie : String) : Boolean ;
Function  DansRuptLibre(QL : TQuery ; fb : TfichierBase ; L1, L2, Ftri : String) : Boolean ;
Function  DansRuptCorresp(L : TStringList ; St : String) : Boolean ;
Procedure DansQuelleRupt(L : TStringList ; St : String ; Var LaRupture : TRuptInf ) ;
Procedure DansQuelleRuptDev(L : TStringList ; St : String ; Var LaDevise : tLDev ) ;
Function  IsRuptDetail(L : TStringList ; St : String) : Boolean ;

Function DansChoixCodeLibre(Co : String ; Q : TQuery ; fb : TfichierBase ; L1, L2, Trie : String) : Boolean ;

Function SimulePrintGroup (L : TStringList ; E : TQuery ; F : Array of TField ; Var Cod,Lib : string ; Var Tot : array of Double ; Var NumR : Integer) : boolean ;

procedure ChargeRegroup (Var L : TStringList ; Qui : String) ;
Function TypeRegroup (L : TStringList ; St : String) : Byte ;
Procedure AddRegroup (L : TStringList ; St : String ; Tot : Array of double) ;
Procedure PrintRegroup (L : TStringList ; St : String ; Var Tot1 : TabTot) ;

implementation

procedure ChargeRuptDevise (Var L : TStringList) ;//Lequel=GE1 ou A52 etc.. Code1 rt Code2 bornes
Var LOR : TOR ;
    Q : TQuery ;
    St : String ;

BEGIN
L:=TStringList.Create ; L.Sorted:=TRUE ; L.Clear ;
St:='Select D_DEVISE, D_LIBELLE, D_SYMBOLE, D_DECIMALE from DEVISE Where D_FERME<>"X" ' ;
St:=St+' Order by D_DEVISE' ;
Q:=OpenSQL(St,TRUE) ;
While Not Q.EOF do
   BEGIN
   LOR:=TOR.Create ;
   LOR.Libelle:=Q.Fields[0].AsString+';'+Q.Fields[1].AsString+';'+Q.Fields[2].AsString+';'+IntToStr(Q.Fields[3].AsInteger)+';' ;
   LOR.Principal:=TRUE ;
   LOR.NbChar:=Length(Q.Fields[0].AsString)+1 ;
   FillChar(LOR.Tot,SizeOf(LOR.Tot),0) ;
   L.AddObject(Q.Fields[0].AsString+'x',LOR) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;


{============================== GESTION DES RUPTURES COMPTABLES ======================}
procedure ChargeRuptCorresp (Var L : TStringList ; Lequel,Code1,Code2 : String ; CodLib : Boolean ) ;//Lequel=GE1 ou A52 etc.. Code1 rt Code2 bornes
Var LOR : TOR ;
    Q : TQuery ;
    St : String ;

BEGIN
L:=TStringList.Create ; L.Sorted:=TRUE ; L.Clear ;
St:='Select CR_CORRESP, CR_LIBELLE from CORRESP Where CR_TYPE="'+Lequel+'" ' ;
If (Code1<>'') And (Code2<>'') Then St:=St+'And CR_CORRESP>="'+Code1+'" And CR_CORRESP<="'+Code2+'" ' ;
St:=St+' Order by CR_CORRESP' ;
Q:=OpenSQL(St,TRUE) ;
While Not Q.EOF do
   BEGIN
   LOR:=TOR.Create ;
   If CodLib then LOR.Libelle:=Q.Fields[0].AsString+' '+Q.Fields[1].AsString Else LOR.Libelle:=Q.Fields[1].AsString ;
   LOR.Principal:=TRUE ;
   LOR.NbChar:=Length(Q.Fields[0].AsString)+1 ;
   FillChar(LOR.Tot,SizeOf(LOR.Tot),0) ;
   L.AddObject(Q.Fields[0].AsString+'x',LOR) ;
   Q.Next
   END ;
Ferme(Q) ;
END ;

procedure ChargeRupt (Var L : TStringList ; Code,Plan,Code1,Code2 : String) ;
Var LOR : TOR ;
    Q : TQuery ;
    St : String ;
BEGIN
L:=TStringList.Create ; L.Sorted:=TRUE ;
St:='Select RU_CLASSE, RU_LIBELLECLASSE from Rupture where RU_NATURERUPT="'+Code+'" and RU_PLANRUPT="'+Plan+'"' ;
If (Code1<>'') And (Code2<>'') Then St:=St+'And RU_CLASSE>="'+Code1+'" And RU_CLASSE<="'+Code2+'" ' ;
Q:=OpenSQL(St,TRUE) ;
L.Clear ;
While Not Q.EOF do
   BEGIN
   LOR:=TOR.Create ;
   LOR.Libelle:=Q.Fields[1].AsString ;
   LOR.Principal:=FALSE ;
   FillChar(LOR.Tot,SizeOf(LOR.Tot),0) ;
   L.AddObject(Q.Fields[0].AsString,LOR) ;
   Q.Next
   END ;
Ferme(Q) ;
END ;

procedure ReinitRupt (L : TStringList) ;
Var LOR : TOR ;
    i : integer ;
BEGIN
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   FillChar(LOR.Tot,SizeOf(LOR.Tot),0) ;
   LOR.Printed:=FALSE ;
   LOR.Filled:=FALSE ;
   END ;
END ;



Procedure NiveauRupt (Var L : TStringList) ;
Var LOR         : TOR ;
    Min,LaPos,i,j : integer ;
    PremCar     : String ;
    ii : integer ;
BEGIN
i:=0 ; ii:=0 ;
While (i<=L.Count-1) do
  BEGIN
  inc(ii) ;
  LaPos:=i ;
  PremCar:=Copy(L[i],1,1) ;
  Min:=Length(L[i]) ;
  While (i<=L.Count-1) and (Copy(L[i],1,1)=PremCar) do
    BEGIN
    if Length(L[i])<Min then Min:=Length(L[i]) ;
    Inc(i) ;
    END ;
  for j:=LaPos to i-1 do
      BEGIN
      LOR:=TOR(L.Objects[j]) ;
      LOR.Principal:=(Length(L[j])=Min) ;
      END ;
  If ii>2000 Then Exit ;
  END ;
END ;

Function IsRuptDetail(L : TStringList ; St : String) : Boolean ;
Var //LOR : TOR ;
    i,Lg : integer ;
BEGIN
Result:=TRUE ; If St='' Then Exit ;
Lg:=Length(St) ;
For i:=0 to L.Count-1 do
   BEGIN
   if Copy(St,1,Length(st)-1)=Copy(L[i],1,Length(st)-1) then
      BEGIN
//      LOR:=TOR(L.Objects[i]) ;
      If Length(L[i])-1>Lg Then Result:=FALSE ;
      END ;
   END ;
END ;

Function PlusDeUneRuptureSurCetteRacine(L : TStringList ; St : String) : Boolean ;
Var i,NbRupt : Integer ;
    St1 : String ;

BEGIN
Result:=FALSE ; NbRupt:=0 ;
For i:=0 To L.Count-1 Do
  BEGIN
  St1:=Copy(L.Strings[i],1,Length(St)) ; If St1=St Then Inc(NbRupt) ;
  If NbRupt>=2 Then BEGIN Result:=TRUE ; Break ; END ;
  END ;
END ;

procedure AddRupt (L : TStringList ; St : String ; Tot : Array of double) ;
Var LOR : TOR ;
    i,j : integer ;
BEGIN
For i:=0 to L.Count-1 do
   BEGIN
   if Copy(St,1,Length(L[i])-1)=Copy(L[i],1,Length(L[i])-1) then
      BEGIN
      LOR:=TOR(L.Objects[i]) ;
      for j:=0 to High(Tot) do LOR.Tot[j]:=LOR.Tot[j]+Tot[j] ;
      LOR.Filled:=TRUE ;
      END ;
   END ;
END ;

Procedure DansQuelleRupt(L : TStringList ; St : String ; Var LaRupture : TRuptInf ) ;
Var LOR : TOR ;
    i,j : integer ;
    Cod,Lib : String ;
BEGIN
Fillchar(LaRupture,SizeOf(LaRupture),#0) ; Cod:='' ; Lib:='' ; j:=0 ;
For i:=0 to L.Count-1 do
   BEGIN
   if Copy(St,1,Length(L[i])-1)=Copy(L[i],1,Length(L[i])-1) then
      BEGIN
      LOR:=TOR(L.Objects[i]) ;
      If j<Length(L[i]) Then BEGIN Cod:=L[i] ; Lib:=LOR.Libelle ; j:=Length(L[i]) ; END ;
      END ;
   END ;
If Cod<>'' Then BEGIN LaRupture.CodRupt:=Cod ; LaRupture.LibRupt:=Lib ; END ;
END ;

Procedure DansQuelleRuptDev(L : TStringList ; St : String ; Var LaDevise : tLDev ) ;
Var LOR : TOR ;
    i,k : integer ;
    Cod,Lib,Sti : String ;
BEGIN
Fillchar(LaDevise,SizeOf(LaDevise),#0) ; Cod:='' ; Lib:='' ;
For i:=0 to L.Count-1 do
   BEGIN
   if Copy(St,1,Length(L[i])-1)=Copy(L[i],1,Length(L[i])-1) then
      BEGIN
      LOR:=TOR(L.Objects[i]) ; Cod:=L[i] ; Lib:=LOR.Libelle ; LaDevise.Code:=Cod ; k:=0 ;
      While Lib<>'' do
        BEGIN
        Sti:=readTokenSt(Lib) ;
        Inc(k);
        if Sti<>'' then
           BEGIN
           Case k Of 2 : LaDevise.Libelle:=Sti ; 3 : LaDevise.Symbole:=Sti ; 4 : LaDevise.Decimale:=StrToInt(Sti) ; END ;
           END ;
        END ;
      Break ;
      END ;
   END ;
END ;


procedure AddRuptCorres (L : TStringList ; St : String ; Tot : Array of double) ;
Var LOR : TOR ;
    i,j : integer ;
BEGIN
For i:=0 to L.Count-1 do
   BEGIN
   //if Copy(St,1,Length(L[i])-1)=Copy(L[i],1,Length(L[i])-1) then
   //if St=L[i]+Copy(St,Length(L[i])+1,V_PGI.Cpta[fbGene].Lg) then
//   if St=Copy(St,Length(L[i])+1,V_PGI.Cpta[fbGene].Lg)+L[i] then
   LOR:=TOR(L.Objects[i]) ;
   If LOR<>NIL Then
      BEGIN
      If Copy(St,1,LOR.NbChar-1)=Copy(L[i],1,Length(L[i])-1)  Then
         BEGIN
         for j:=0 to High(Tot) do LOR.Tot[j]:=LOR.Tot[j]+Tot[j] ;
         LOR.Filled:=TRUE ;
         END ;
      END ;
   END ;
END ;

Function PrintRupt (L : TStringList ; St : String ; Var Cod,Lib : string ; Var PourTotalGen : Boolean ; EnRupture : Boolean ; Var Tot : array of Double) : boolean ;
Var LOR  : TOR ;
    i,j  : integer ;
    Find : boolean ;
    //Passe : String ;
BEGIN
if EnRupture then St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' ;
for j:=0 to High(Tot) do Tot[j]:=0 ;
Find:=FALSE ;  PourTotalGen:=FALSE ;
If L=NIL Then BEGIN Result:=Find ; Exit ; END ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   if ((L[i]<St) and (Not LOR.Printed) and (LOR.Filled) ) then
      BEGIN
      Find:=TRUE ;
      for j:=0 to High(Tot) do Tot[j]:=LOR.Tot[j];
      Lib:=LOR.Libelle ;
      PourTotalGen:=LOR.Principal ;
      Cod:=L[i] ;
      LOR.Printed:=TRUE ;
      Break ;
      END ;
   END ;
PrintRupt:=Find ;
END ;
(*
Function PrintRupt2 (L : TStringList ; St : String ; Var Cod,Lib : string ; Var PourTotalGen : Boolean ; EnRupture : Boolean ; Var Tot : array of Double ;
                     E : TQuery ; F : Array of TField) : boolean ;
Var LOR  : TOR ;
    i,j  : integer ;
    Find : boolean ;
    //Passe : String ;
BEGIN
if EnRupture then St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' ;
If E<>NIL Then
  BEGIN
  If Not E.Eof Then
    BEGIN
    E.Next ;
    If Not E.Eof Then
      BEGIN
      If St>F[0].AsString Then St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' Else St:=F[0].AsString ;
      E.Prior ;
      END Else St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' ;
    END Else St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' ;
  END ;
for j:=0 to High(Tot) do Tot[j]:=0 ;
Find:=FALSE ;  PourTotalGen:=FALSE ;
If L=NIL Then BEGIN Result:=Find ; Exit ; END ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   if ((L[i]<St) and (Not LOR.Printed) and (LOR.Filled) ) then
      BEGIN
      Find:=TRUE ;
      for j:=0 to High(Tot) do Tot[j]:=LOR.Tot[j];
      Lib:=LOR.Libelle ;
      PourTotalGen:=LOR.Principal ;
      Cod:=L[i] ;
      LOR.Printed:=TRUE ;
      Break ;
      END ;
   END ;
PrintRupt2:=Find ;
END ;

*)
Function PrintRupt2 (L : TStringList ; St : String ; Var Cod,Lib : string ; Var PourTotalGen : Boolean ; EnRupture : Boolean ; Var Tot : array of Double ;
                     F : Array of TField ; Q : TQuery = NIL; STTS : String = '') : boolean ;
Var LOR  : TOR ;
    i,j  : integer ;
    Find : boolean ;
    OkSTTS,RuptSTTS : Boolean ;
    QSTTS : tVariantField ;
    VCS1,VCS2 : Variant ;
    OldSt : String ;
    //Passe : String ;
BEGIN
if EnRupture then St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' ;
OkSTTS:=FALSE ; RuptSTTS:=FALSE ; OldSt:=St ;
If (STTS<>'') And (Q<>NIL) Then
  BEGIN
  QSTTS:=TVariantField(Q.FindField(STTS)) ;
  If QSTTS<>NIL Then
    BEGIN
    VCS1:=QSTTS.AsVariant ;
    If (VarType(VCS1) In [varEmpty,varNull])=FALSE Then OkSTTS:=TRUE ;
    END ;
  END ;
If (Q<>NIL) Then
  BEGIN
  If Not Q.Eof Then
    BEGIN
    Q.Next ;
    If Not Q.Eof Then
      BEGIN
      If OkSTTS Then
        BEGIN
        QSTTS:=TVariantField(Q.FindField(STTS)) ;
        VCS2:=QSTTS.AsVariant ;
        If VCS2<>VCS1 Then RuptSTTS:=TRUE ;
        END ;
      If RuptSTTS Then St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' Else St:=F[0].AsString ;
      Q.Prior ;
      END Else St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' ;
    END Else St:='zzzzzzzzzzzzzzzzzzzzzzzzzz' ;
  END ;
for j:=0 to High(Tot) do Tot[j]:=0 ;
Find:=FALSE ;  PourTotalGen:=FALSE ;
If L=NIL Then BEGIN Result:=Find ; Exit ; END ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   if ((L[i]<St) and (Not LOR.Printed) and (LOR.Filled) ) then
      BEGIN
      Find:=TRUE ;
      for j:=0 to High(Tot) do Tot[j]:=LOR.Tot[j];
      Lib:=LOR.Libelle ;
      PourTotalGen:=LOR.Principal ;
      Cod:=L[i] ;
      LOR.Printed:=TRUE ;
      Break ;
      END ;
   END ;
PrintRupt2:=Find ;
END ;

procedure VideRupt (Var L : TStringList ) ;
Var LOR : TOR ;
    i : integer ;
BEGIN
If L=Nil Then Exit ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   LOR.Free ;
   END ;
L.Clear ;
L.Free ; L:=NIL ;
END ;

{============================== GESTION DES GROUPES ===========================}
procedure ChargeGroup (Var L : TStringList ; Lib : Array of String) ;
Var LOR : TOR ;
    j : integer ;
BEGIN
L:=TStringList.Create ;
L.Clear ;
For j:=0 to High(Lib) do
   BEGIN
   LOR:=TOR.Create ;
   LOR.Libelle:=Lib[j] ; LOR.Filled:=FALSE ;
   FillChar(LOR.Tot,SizeOf(LOR.Tot),0) ;
   L.AddObject('',LOR) ;
   END ;
END ;

procedure AddGroupLibre (L : TStringList ; Q : TQuery ; fb : TfichierBase ; Ftri : String ; Tot : Array of double) ;
Var F : Array [0..9] of TField ;
    StF,St,Sti,Pref : String;
    i : integer ;
BEGIN
St:=FTri ;
for i:=0 to 9 do F[i]:=Nil ;
Case Fb of
    fbGene               :  BEGIN Pref:='G'  ; END ;
    fbAux                :  BEGIN Pref:='T'  ; END ;
    fbBudgen             :  BEGIN Pref:='BG' ; END ;
    fbBudSec1..fbBudSec5 :  BEGIN Pref:='BS' ; END ;
    fbAxe1..fbAxe5       :  BEGIN Pref:='S'  ; END ;
    fbImmo               :  BEGIN Pref:='I'  ; END ;
    End ;
i:=0 ;
While St<>'' do
   BEGIN
   Sti:=readTokenSt(St) ;
//Simon gaffe a ASC et DSC
   if Sti<>'' then
      BEGIN
      StF:=Pref+'_TABLE'+Copy(Sti,3,1) ;
      F[i]:=Q.FindField(StF) ;
      Inc(i);
      END ;
   END ;
AddGroup (L,F,Tot) ;
END ;

procedure AddGroup (L : TStringList ; F : Array of TField ; Tot : Array of double) ;
Var LOR : TOR ;
    i,j : integer ;
    St : Array [0..12] of String ;
BEGIN
for j:=0 to High(F) do
   BEGIN
//   Rony
//   if F[j].IsNull then St[j]:='ere\r' else
   if (F[j]=NIL) or (F[j].AsString='') then St[j]:='ere\r' else
      BEGIN
      Case F[j].Tag of
        0 : St[j]:=F[j].AsString ;
        1 : St[j]:=FormatDatetime('mmmm yyyy',F[j].AsDateTime) ;
        2 : St[j]:=FormatDatetime('yyyy',F[j].AsDateTime) ;
        100 : St[j]:=Copy(F[j].AsString,1,LgSpecif1) ; //Spécif CERIC
        END ;
      END ;
   END ;
For i:=0 to L.Count-1 do
  if St[i]<>'ere\r' then
   BEGIN
   L[i]:=St[i] ;
   LOR:=TOR(L.Objects[i]) ;
   for j:=0 to High(Tot) do LOR.Tot[j]:=LOR.Tot[j]+Tot[j] ;
   LOR.Filled:=TRUE ;
   END ;
END ;

Function PrintGroupLibre (L : TStringList ; Q : TQuery ; fb : TfichierBase ; Ftri : String ;
                          Var Cod,Lib,Lib1 : string ; Var Tot : array of Double ; Var NumR : Integer ;
                          Var Col : TColor ; Var LibRuptInf : Array Of TRuptInf  ; STTS : String = '') : boolean ;
Var F : Array [0..9] of TField ;
    TSt : Array[0..9] Of String ;
    StF,St,Sti,Pref,STypeCpte,Etoiles,SQL,STyp : String;
    i,j : integer ;
    Q1 : TQuery ;
BEGIN
St:=FTri ; STypeCpte:='' ; Lib1:='' ; Fillchar(Tst,SizeOf(Tst),#0) ;
for i:=0 to High(LibRuptInf) do BEGIN LibRuptInf[i].CodRupt:='' ; LibRuptInf[i].LibRupt:='' ;  END ;
for i:=0 to 9 do F[i]:=Nil ;
Case Fb of
    fbGene               :  BEGIN Pref:='G'  ; STypeCpte:='G' ; END ;
    fbAux                :  BEGIN Pref:='T'  ; STypeCpte:='T' ; END ;
    fbBudgen             :  BEGIN Pref:='BG' ; STypeCpte:='B' ; END ;
    fbBudSec1..fbBudSec5 :  BEGIN Pref:='BS' ; STypeCpte:='D' ; END ;
    fbAxe1..fbAxe5       :  BEGIN Pref:='S'  ; STypeCpte:='S' ; END ;
    fbImmo               :  BEGIN Pref:='I'  ; STypeCpte:='I' ; END ;
    End ;
i:=0 ;
While St<>'' do
   BEGIN
   Sti:=readTokenSt(St) ;
//Simon gaffe a ASC et DSC
   if Sti<>'' then
      BEGIN
      TSt[i]:=Sti ; StF:=Pref+'_TABLE'+Copy(Sti,3,1) ; F[i]:=Q.FindField(StF) ; Inc(i);
      END ;
   END ;
Result:=PrintGroup(L,Q,F,Cod,Lib,Tot,NumR,STTS) ;
If Result And (STypeCpte<>'') And (Cod<>'') Then
   BEGIN
   SQL:='Select NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE=:TYP AND NT_NATURE=:COD ' ;
   Q1:=TQuery.Create(Application) ;
   Q1.DatabaseName:=DBSOC.DataBaseName ;
   Q1.SessionName:=DBSOC.SessionName ;
//   Q1:=OpenSQL('Select NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE="'+STypeCpte+'" AND NT_NATURE="'+Cod+'" ',TRUE) ;
   Q1.SQL.Clear ; Q1.SQL.Add(SQL) ; ChangeSQL(Q1) ; //Q1.Prepare ;
   PrepareSQLODBC(Q1) ;
//   STyp:=STypeCpte+FormatFloat('00',NumR) ;
   STyp:=TSt[NumR] ;
   Q1.Close ; Q1.ParamByName('TYP').AsString:=STyp ; Q1.ParamByName('COD').AsString:=Cod ; Q1.Open ;
   If Not Q1.Eof Then Lib1:=Q1.Fields[1].AsString ;
   If NumR>0 Then
      BEGIN
      For i:=0 To NumR-1 Do If i<=L.Count Then
        BEGIN
//        STyp:=STypeCpte+FormatFloat('00',i) ;
        STyp:=TSt[i] ;
        Q1.Close ; Q1.ParamByName('TYP').AsString:=STyp ; Q1.ParamByName('COD').AsString:=L[i] ; Q1.Open ;
        Etoiles:='' ;  For j:=0 to i do Etoiles:=Etoiles+'*' ;
        If Not Q1.Eof Then
           BEGIN
           LibRuptInf[i].CodRupt:=Etoiles+' '+Q1.Fields[0].AsString+' ' ;
           LibRuptInf[i].LibRupt:=Q1.Fields[1].AsString ;
           END ;
        END ;
      END ;
   Ferme(Q1) ;
   END ;
St:=Cod ;
Etoiles:='' ;
For i:=0 to NumR do Etoiles:=Etoiles+'*' ;
Cod:=Etoiles+' '+St ;

Case NumR of
  0 : Col:=clPurple ;
  1 : Col:=clMaroon ;
  2 : Col:=clGreen ;
  3 : Col:=clNavy ;
  4 : Col:=clOlive ;
  5 : Col:=clTeal ;
  6 : Col:=clGray ;
  7 : Col:=clRed ;
  8 : Col:=clBlue ;
  9 : Col:=clBlack ;
  End ;

END ;



Function PrintGroup (L : TStringList ; E : TQuery ; F : Array of TField ; Var Cod,Lib : string ; Var Tot : array of Double ; Var NumR : Integer ; STTS : String = '') : boolean ;
Var LOR  : TOR ;
    i,j  : integer ;
    Find : boolean ;
    St   : Array [0..12] of String ;
    ARupte : boolean ;
    OkSTTS,RuptSTTS : Boolean ;
    QSTTS : tVariantField ;
    VCS1,VCS2 : Variant ;
BEGIN
for j:=0 to High(Tot) do Tot[j]:=0 ;
NumR:=0 ;
for j:=0 to High(F) do St[j]:='ere\r' ;
ARupte:=FALSE ;
OkSTTS:=FALSE ; RuptSTTS:=FALSE ;
If STTS<>'' Then
  BEGIN
  QSTTS:=TVariantField(E.FindField(STTS)) ;
  If QSTTS<>NIL Then
    BEGIN
    VCS1:=QSTTS.AsVariant ;
    If (VarType(VCS1) In [varEmpty,varNull])=FALSE Then OkSTTS:=TRUE ;
    END ;
  END ;
if Not E.EOF then
   BEGIN
   E.Next ;
   if Not E.EOF then
      BEGIN
      If OkSTTS Then
        BEGIN
        QSTTS:=TVariantField(E.FindField(STTS)) ;
        VCS2:=QSTTS.AsVariant ;
        If VCS2<>VCS1 Then RuptSTTS:=TRUE ;
        END ;
      for j:=0 to High(F) do
         BEGIN
         LOR:=TOR(L.Objects[j]) ;
         if ARupte then St[j]:='ere\r' else
            BEGIN
            if F[j]=Nil then St[j]:='' else
             Case F[j].Tag of
              0 : St[j]:=F[j].AsString ;
              1 : St[j]:=FormatDatetime('mmmm yyyy',F[j].AsDateTime) ;
              2 : St[j]:=FormatDatetime('yyyy',F[j].AsDateTime) ;
              100 : St[j]:=Copy(F[j].AsString,1,LgSpecif1) ; //Spécif CERIC
              END ;
            END ;
         if ((L[j]<>'') and (L[j]<>St[j]) and (LOR.Filled)) then ARupte:=TRUE ;
         END ;
      E.Prior ;
      END ;
   END ;
Find:=FALSE ;
For i:=L.Count-1 downto 0 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   if ((L[i]<>'') and ((L[i]<>St[i]) Or RuptSTTS) and (LOR.Filled)) then
      BEGIN
      Find:=TRUE ;
      for j:=0 to High(Tot) do Tot[j]:=LOR.Tot[j];
      Lib:=LOR.Libelle ;
      Cod:=L[i] ; NumR:=i ;
      LOR.Filled:=FALSE ;
      Break ;
      END ;
   if Not LOR.Filled then
      BEGIN
      for j:=0 to High(Tot) do LOR.Tot[j]:=0 ;
      END ;
   END ;
PrintGroup:=Find ;
END ;

procedure VideGroup (Var L : TStringList ) ;
BEGIN
VideRupt(L) ;
END ;

{============================== GESTION DES RECAPS ===========================}
procedure ChargeRecap (Var L : TStringList ) ;
BEGIN
L:=TStringList.Create ; L.Sorted:=TRUE ;
L.Clear ;
END ;

(*
procedure AddRecap (L : TStringList ; Cod,Lib : String ; Tot : Array of double) ;
Var LOR : TOR ;
    i,j : integer ;
BEGIN
if Cod='' then exit ;
i:=L.IndexOf(Cod) ;
if i<0 then
   BEGIN
   LOR:=TOR.Create ;
   LOR.Libelle:=Lib ;
   for j:=0 to High(Tot) do LOR.Tot[j]:=Tot[j] ;
   LOR.Filled:=TRUE ;
   L.AddObject(Cod,LOR) ;
   END else
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   for j:=0 to High(Tot) do LOR.Tot[j]:=LOR.Tot[j]+Tot[j] ;
   LOR.Filled:=TRUE ;
   END ;
END ;
*)

procedure AddRecap (L : TStringList ; F : Array Of String ; FL : Array Of String ; Tot : Array of double) ;
Var LOR : TOR ;
    i,j,k : integer ;
    Cod,Lib : String ;
BEGIN
Cod:='' ; for j:=0 to High(F) do BEGIN if F[j]<>'' Then Cod:=Cod+F[j] ; END ; if Cod='' then exit ;
i:=L.IndexOf(Cod) ;
if i<0 then
   BEGIN
   LOR:=TOR.Create ;
   Lib:='' ; for k:=0 to High(FL) do BEGIN if FL[k]<>'' Then Lib:=Lib+' '+FL[k] ; END ;
   LOR.Libelle:=Lib ;
   for j:=0 to High(Tot) do LOR.Tot[j]:=Tot[j] ;
   LOR.Filled:=TRUE ;
   L.AddObject(Cod,LOR) ;
   END else
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   for j:=0 to High(Tot) do LOR.Tot[j]:=LOR.Tot[j]+Tot[j] ;
   LOR.Filled:=TRUE ;
   END ;
END ;

Function PrintRecap (L : TStringList ; Var Cod,Lib : string ; Var Tot : array of Double) : boolean ;
Var LOR  : TOR ;
    i,j  : integer ;
    Find : boolean ;
BEGIN
for j:=0 to High(Tot) do Tot[j]:=0 ;
Find:=FALSE ;
If L=NIL Then BEGIN PrintRecap:=FALSE ; Exit ; END ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   if LOR.Filled then
      BEGIN
      Find:=TRUE ;
      for j:=0 to High(Tot) do Tot[j]:=LOR.Tot[j];
      Lib:=LOR.Libelle ;
      Cod:=L[i] ;
      LOR.Filled:=FALSE ;
      Break ;
      END ;
   END ;
PrintRecap:=Find ;
if Find=FALSE then
   BEGIN
   For i:=0 to L.Count-1 do
      BEGIN
      LOR:=TOR(L.Objects[i]) ;
      LOR.Free ;
      END ;
   L.Clear ;
   END ;
END ;

procedure VideRecap (Var L : TStringList ) ;
BEGIN
VideRupt(L) ;
END ;


Function DansRupt (L : TStringList ; St : String) : Boolean ;
Var i : integer ;
    OkOk : Boolean ;
BEGIN
OkOk:=False ;
For i:=0 to L.Count-1 do
   BEGIN
   if Copy(St,1,Length(L[i])-1)=Copy(L[i],1,Length(L[i])-1) then
      BEGIN
      OkOk:=True ; Break ;
      END ;
   END ;
Result:=OkOk ;
END ;

Function DansRuptCorresp(L : TStringList ; St : String) : Boolean ;
Var i : integer ;
    OkOk : Boolean ;
    LOR : TOR ;
BEGIN
OkOk:=False ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   If LOR<>NIL Then
      BEGIN
      If Copy(St,1,LOR.NbChar-1)=Copy(L[i],1,Length(L[i])-1)  Then
         BEGIN
         OkOk:=True ; Break ;
         END ;
      END ;
   END ;
Result:=OkOk ;
END ;

Function DansLibre (F : Array of TField ; L1, L2, Trie : String) : Boolean ;
Var i,J  : integer ;
    OkOk : Boolean ;
    St,LC1, LC2   : Array [0..9] of String ;
    StTemp1, stTemp2, StTrie, StQuel : String ;
BEGIN
OkOk:=False ;
for j:=0 to High(F) do
   BEGIN
   if (F[j]=NIL) or (F[j].AsString='') then St[j]:='ere\r' else St[j]:=F[j].AsString ;
   END ;
StTemp1:=L1 ; StTemp2:=L2 ; i:=0 ; StTrie:=Trie ;
While StTemp1<>'' do
      BEGIN
      LC1[i]:=ReadtokenSt(StTemp1) ; LC2[i]:=ReadtokenSt(StTemp2) ;
      Inc(i) ;
      END;
While Trie<>'' do
      BEGIN
      StTrie:=ReadTokenSt(Trie) ;
      I:=StrToInt(Copy(StTrie,3,1)) ; StQuel:=LC1[i] ; If F[i]=Nil then continue ;
      if StQuel='' then Break ;
      If (Stquel[1]<>'#') and (Stquel[1]<>'-') then
         BEGIN
         If (Stquel[1]='*') and (F[i].AsString<>'') then OkOk:=True else
//CP 03/02/98 If ((LC1[i]>=F[i].AsString) and (LC2[i]<=F[i].AsString)) then
            If ((F[i].AsString>=LC1[i]) and (F[i].AsString<=LC2[i])) then OkOk:=True else OkOk:=False ;
         END Else okok:=False ;
      If Not OkOk then Break ;
      END;

(*
For i:=0 to High(St) do
    BEGIN
    if St[i]<>'ere\r' then
       BEGIN
       OkOk:=True ; Break ;
       END ;
    END ;
*)
Result:=OkOk ;
END ;

Function DansRuptLibre(QL : TQuery ; fb : TfichierBase ; L1, L2, Ftri : String) : Boolean ;
Var F    : Array [0..9] of TField ;
    i    : integer ;
    StF,St, Sti,Pref : String;
BEGIN
St:=FTri ;
for i:=0 to 9 do F[i]:=Nil ;
Case Fb of
    fbGene               :  Pref:='G'  ;
    fbAux                :  Pref:='T'  ;
    fbBudgen             :  Pref:='BG' ;
    fbBudSec1..fbBudSec5 :  Pref:='BS' ;
    fbAxe1..fbAxe5       :  Pref:='S'  ;
    fbImmo               :  Pref:='I'  ;
    End ;
//i:=0 ;
While St<>'' do
   BEGIN
   Sti:=readTokenSt(St) ;
   if Sti<>'' then
      BEGIN
      I:=StrToInt(Copy(Sti,3,1)) ;
//      StF:=Pref+'_TABLE'+Copy(Sti,3,1) ;
      StF:=Pref+'_TABLE'+IntToStr(i) ;
      F[i]:=QL.FindField(StF) ;
//      Inc(i);
      END ;
   END ;
Result:=DansLibre(F,L1, L2, FTri) ;
END ;


Function DansChoixCodeLibre(Co : String ; Q : TQuery ; fb : TfichierBase ; L1, L2, Trie : String) : Boolean ;
Var StCode, StListes1, StListes2, StC1, StC2, StTrie, Pref, St : String ;
    LISTE1, LISTE2 : array[0..9] of String ; i : Byte ;
BEGIN
Result:=True ;
Case Fb of
    fbGene               :  Pref:='G'  ;
    fbAux                :  Pref:='T'  ;
    fbBudgen             :  Pref:='BG' ;
    fbBudSec1..fbBudSec5 :  Pref:='BS' ;
    fbAxe1..fbAxe5       :  Pref:='S'  ;
    fbImmo               :  Pref:='I'  ;
    End ;
StListes1:=L1 ; StListes2:=L2 ; i:=0 ; StCode:=Co ;
While StListes1<>'' do
      BEGIN
      LISTE1[i]:=readTokenSt(StListes1) ;
      LISTE2[i]:=readTokenSt(StListes2) ;
      Inc(i) ;
      END ;
While Trie<>'' do
      BEGIN
      StTrie:=readTokenSt(Trie) ; If StTrie='' then Exit ;
      I:=StrToInt(Copy(StTrie,3,1)) ;
      StC1:=LISTE1[i] ; StC2:=LISTE2[i] ;
      St:=Q.FindField(''+Pref+'_TABLE'+Copy(StTrie,3,1)+'').AsString ;
      If StCode<>St then begin Result:=False ; continue ; end ;
      if ((StC1[1]<>'#') AND  (StC1[1]<>'-'))
         or ((StC2[1]<>'#') AND (StC2[1]<>'-')) then
         BEGIN
         if ((StCode>=StC1) and (StCode<=StC2)) or (StC1[1]='*') then
            BEGIN
            Result:=True ; Break ;
            END Else
            BEGIN
            Result:=False ; Continue ;
            END ;
         END Else Result:=False ;
      END ;
END ;

Function SimulePrintGroup (L : TStringList ; E : TQuery ; F : Array of TField ; Var Cod,Lib : string ; Var Tot : array of Double ; Var NumR : Integer) : boolean ;
Var LOR  : TOR ;
    i,j  : integer ;
    Find : boolean ;
    St   : Array [0..12] of String ;
    ARupte : boolean ;
BEGIN
for j:=0 to High(Tot) do Tot[j]:=0 ;
NumR:=0 ;
for j:=0 to High(F) do St[j]:='ere\r' ;
ARupte:=FALSE ;
if Not E.EOF then
   BEGIN
   E.Next ;
   if Not E.EOF then
      BEGIN
      for j:=0 to High(F) do
         BEGIN
         LOR:=TOR(L.Objects[j]) ;
         if ARupte then St[j]:='ere\r' else
            BEGIN
            if F[j]=Nil then St[j]:='' else
             Case F[j].Tag of
              0 : St[j]:=F[j].AsString ;
              1 : St[j]:=FormatDatetime('mmmm yyyy',F[j].AsDateTime) ;
              2 : St[j]:=FormatDatetime('yyyy',F[j].AsDateTime) ;
              100 : St[j]:=Copy(F[j].AsString,1,LgSpecif1) ; //Spécif CERIC
              END ;
            END ;
         if ((L[j]<>'') and (L[j]<>St[j]) and (LOR.Filled)) then ARupte:=TRUE ;
         END ;
      E.Prior ;
      END ;
   END ;
Find:=FALSE ;
For i:=L.Count-1 downto 0 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   if ((L[i]<>'') and (L[i]<>St[i]) and (LOR.Filled)) then
      BEGIN
      Find:=TRUE ;
      for j:=0 to High(Tot) do Tot[j]:=LOR.Tot[j];
      Lib:=LOR.Libelle ;
      Cod:=L[i] ; NumR:=i ;
      (*
      LOR.Filled:=FALSE ;
      *)
      Break ;
      END ;
   (*
   if Not LOR.Filled then
      BEGIN
      for j:=0 to High(Tot) do LOR.Tot[j]:=0 ;
      END ;
   *)
   END ;
SimulePrintGroup:=Find ;
END ;

Function WhatRegroup(St : String) : String ;
Var St1,S : String ;
BEGIN
S:='' ;
While St<>'' Do
  BEGIN
  St1:=ReadTokenSt(St) ; S:=S+' OR CR_CORRESP="'+St1+'" ' ;
  END ;
If S<>'' Then
  BEGIN
  Delete(S,1,3) ; S:='AND ('+S+') ' ;
  END ;
Result:=S ;
END ;

procedure ChargeRegroup (Var L : TStringList ; Qui : String) ;
Var LOR : TOR ;
    Q : TQuery ;
    St : String ;
BEGIN
L:=TStringList.Create ; L.Sorted:=TRUE ;
St:='Select CR_CORRESP,CR_LIBELLE,CR_ABREGE from CORRESP where CR_TYPE="ZG1" ' ;
If Qui<>'' Then St:=St+WhatRegroup(Qui) ;
Q:=OpenSQL(St,TRUE) ;
L.Clear ;
While Not Q.EOF do
   BEGIN
   LOR:=TOR.Create ;
   LOR.Libelle:=Q.Fields[0].AsString+';'+Q.Fields[1].AsString+';'+Q.Fields[2].AsString+';' ;
   LOR.Principal:=FALSE ;
   FillChar(LOR.Tot,SizeOf(LOR.Tot),0) ;
   L.AddObject(Q.Fields[0].AsString,LOR) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

Function Decante(St : String ; Var St1,St2,St3 : String) : Boolean ;
BEGIN
St1:='' ; St2:='' ; St3:='' ; Result:=FALSE ;
If St='' Then Exit ;
St1:=ReadTokenSt(St) ; If St1='' Then Exit ;
St2:=ReadTokenSt(St) ; If St2='' Then Exit ;
St3:=ReadTokenSt(St) ; If St3='' Then Exit ;
Result:=TRUE ;
END ;

Function DansRegroup(L : TStringList ; Cpt : String ; Var Ind : Integer) : Boolean ;
Var LOR : TOR ;
    i : integer ;
    St,St1,St2,St3 : String ;
BEGIN
Result:=FALSE ; Ind:=-1 ;
If L=Nil Then Exit ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   St:=LOR.Libelle ;
   If Decante(St,St1,St2,St3) Then
     BEGIN
     If (Copy(Cpt,1,Length(St2))>=St2) And (Copy(Cpt,1,Length(St3))<=St3) Then BEGIN Result:=TRUE ; Ind:=i ; Exit ; END ;
     END ;
   END ;
END ;

Function EstRegroup(L : TStringList ; Cpt : String ; Var Ind : Integer) : Boolean ;
Var LOR : TOR ;
    i : integer ;
    St,St1,St2,St3 : String ;
BEGIN
Result:=FALSE ; Ind:=-1 ;
If L=Nil Then Exit ;
For i:=0 to L.Count-1 do
   BEGIN
   LOR:=TOR(L.Objects[i]) ;
   St:=LOR.Libelle ;
   If Decante(St,St1,St2,St3) Then
     BEGIN
     If St1=Cpt Then BEGIN Result:=TRUE ; Ind:=i ; Exit ; END ;
     END ;
   END ;
END ;

Function TypeRegroup (L : TStringList ; St : String) : Byte ;
Var ind : integer ;
BEGIN
Result:=0 ;
if DansRegroup(L,St,Ind) then Result:=1 Else if EstRegroup(L,St,Ind) then Result:=2 ;
END ;


Procedure AddRegroup (L : TStringList ; St : String ; Tot : Array of double) ;
Var LOR : TOR ;
    ind,j : integer ;
BEGIN
if DansRegroup(L,St,Ind) then
    BEGIN
    LOR:=TOR(L.Objects[Ind]) ;
    for j:=0 to High(Tot) do LOR.Tot[j]:=LOR.Tot[j]+Tot[j] ;
    LOR.Filled:=TRUE ;
    END ;
END ;

Procedure PrintRegroup (L : TStringList ; St : String ; Var Tot1 : TabTot) ;
Var LOR : TOR ;
    ind,j : integer ;
    Tot : Array[0..4] of double ;
BEGIN
for j:=0 to High(Tot) do Tot[j]:=0 ;
Tot[1]:=Tot1[1].TotDebit ; Tot[2]:=Tot1[1].TotCredit ;
Tot[3]:=Tot1[2].TotDebit ; Tot[4]:=Tot1[2].TotCredit ;
if EstRegroup(L,St,Ind) then
    BEGIN
    LOR:=TOR(L.Objects[ind]) ;
    if (Not LOR.Printed) and (LOR.Filled)  then
       BEGIN
       for j:=0 to High(Tot) do Tot[j]:=LOR.Tot[j];
       LOR.Printed:=TRUE ;
       END ;
    END ;
//for j:=0 to High(Tot1) do Tot1[j]:=Tot[j] ;
Tot1[1].TotDebit  := Tot[1] ; Tot1[1].TotCredit := Tot[2] ;
Tot1[2].TotDebit  := Tot[3] ; Tot1[2].TotCredit := Tot[4] ;
END ;


end.
