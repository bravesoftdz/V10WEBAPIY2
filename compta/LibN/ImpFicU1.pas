unit ImpFicU1 ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, Buttons, ExtCtrls, Hctrls, TImpFic,HMsgBox, HStatus,
{$ifdef eAGLClient}
  UTOB;
{$ELSE}
  PrintDBG,
  HSysMenu,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$IFDEF VER150}
   Variants,
 {$ENDIF}
  ColMemo, ComCtrls, HTB97,
  Menus, FileCtrl ;
{$ENDIF}

Procedure FaitFichierEnCours(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Procedure FaitFichierEnCoursDetail(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Procedure InitFEC(Var FEC : tFicEnCours) ;
Procedure RAZFec(Var FEC : tFicEnCours) ;
Procedure FaitEnCoursComptable(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Procedure FaitEnCoursTraiteNonEchu(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Procedure FaitEnCoursTraiteEchuNonRegle(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Function  EnCoursRegleNonEchu(Aux : String ; DateButoir : tDateTime) : Double ;
procedure VideStringList ( L : HTStringList ) ;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  HEnt1, Ent1;


Function FiltreTL(Var Crit : tCritExpECC) : String ;
Var I : Integer ;
    St : String ;
BEGIN
St:='' ;
For i:=0 To 9 Do
  BEGIN
  If Crit.TL[i]<>'' Then St:=St+' AND T_TABLE'+IntToStr(i)+'="'+Crit.TL[i]+'" ' ;
  END ;
Result:=St ;
END ;

Procedure FaitEnCoursComptable(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Var St,Cpt,StTL : String ;
    Q : TQuery ;
    i : Integer ;
    EnCoursP : TEnCoursP ;
    Existe : Boolean ;
    Where411 : String ;
BEGIN
(*
st:='SELECT * FROM TIERS WHERE (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD") ' ;
StTL:=FiltreTL(Crit) ; If StTL<>'' Then St:=St+StTL ;
*)
st:='SELECT E_AUXILIAIRE, T_CODEIMPORT, SUM(E_DEBIT) AS DP, SUM(E_CREDIT) AS CP, SUM(E_DEBITDEV) AS DD, SUM(E_CREDITDEV) AS CD,0 AS DE, 0 AS CE FROM ECRITURE ' ;
St:=St+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
St:=St+' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
St:=St+' WHERE E_AUXILIAIRE<>"" AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD") ' ;
If Crit.Aux1<>'' Then St:=St+' AND E_AUXILIAIRE>="'+Crit.Aux1+'" ' ;
If Crit.Aux2<>'' Then St:=St+' AND E_AUXILIAIRE<="'+Crit.Aux2+'" ' ;
If Crit.St411SOLDE<>'' Then
  BEGIN
  Where411:=AnalyseCompte(Crit.St411SOLDE,fbGene,FALSE,FALSE) ;
  If Where411<>'' Then St:=St+' AND '+Where411  ;
  END ;
StTL:=FiltreTL(Crit) ; If StTL<>'' Then St:=St+StTL ;
St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" ' ;
St:=St+' AND ((E_EXERCICE="'+VH^.EnCours.Code+'") OR (E_EXERCICE="'+VH^.SUIVANT.Code+'" AND E_ECRANOUVEAU="N")) ' ;
St:=St+' GROUP BY E_AUXILIAIRE, T_CODEIMPORT ' ;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q),'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  Cpt:=Q.FindField('E_AUXILIAIRE').AsString ;
  i:=FEC.LR.IndexOf(Cpt) ;
  If i<0 Then
    BEGIN
    Existe:=FALSE ;
    EnCoursP:=TEnCoursP.Create ;
    EnCoursP.T.Cpt:=Q.FindField('E_AUXILIAIRE').AsString ;
    EnCoursP.T.Cpt1:=Q.FindField('T_CODEIMPORT').AsString ;
    END Else
    BEGIN
    Existe:=TRUE ;
    EnCoursP:=TEnCoursP(FEC.LR.Objects[i]) ;
    END ;
  (*
  EnCoursP.T.Montant[eccComptable].D:=Q.FindField('T_TOTDEBE').AsFloat+Q.FindField('T_TOTDEBS').AsFloat+Q.FindField('T_TOTDEBANO').AsFloat ;
  EnCoursP.T.Montant[eccComptable].C:=Q.FindField('T_TOTCREE').AsFloat+Q.FindField('T_TOTCRES').AsFloat+Q.FindField('T_TOTCREANO').AsFloat ;
  *)
  EnCoursP.T.Montant[eccComptable].D:=Q.FindField('DP').AsFloat ;
  EnCoursP.T.Montant[eccComptable].C:=Q.FindField('CP').AsFloat ;
  EnCoursP.T.Montant[eccComptable].DE:=Q.FindField('DE').AsFloat ;
  EnCoursP.T.Montant[eccComptable].CE:=Q.FindField('CE').AsFloat ;
  If Not Existe Then FEC.LR.AddObject(Cpt,EnCoursP) ;
  Q.Next ;
  END ;
FiniMove ;
Ferme(Q) ;
END ;

Procedure FaitEnCoursTraiteNonEchu(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Var St,StTL,Cpt,Where413,St1,St2,St3 : String ;
    Q : TQuery ;
    i : Integer ;
    EnCoursP : TEnCoursP ;
    Existe : Boolean ;
BEGIN
st:='SELECT E_AUXILIAIRE, SUM(E_DEBIT) AS DP, SUM(E_CREDIT) AS CP, SUM(E_DEBITDEV) AS DD, SUM(E_CREDITDEV) AS CD, 0 AS DE, 0 AS CE FROM ECRITURE ' ;
St:=St+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
St:=St+' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
St:=St+' LEFT JOIN MODEPAIE ON E_MODEPAIE=MP_MODEPAIE ' ;
St:=St+' WHERE E_AUXILIAIRE<>"" AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD") ' ;
If Crit.Aux1<>'' Then St:=St+' AND E_AUXILIAIRE>="'+Crit.Aux1+'" ' ;
If Crit.Aux2<>'' Then St:=St+' AND E_AUXILIAIRE<="'+Crit.Aux2+'" ' ;
StTL:=FiltreTL(Crit) ; If StTL<>'' Then St:=St+StTL ;
(*
If Crit.Gen1<>'' Then St:=St+' AND E_GENERAL>="'+Crit.Gen1+'" ' ;
If Crit.Gen2<>'' Then St:=St+' AND E_GENERAL<="'+Crit.Gen2+'" ' ;
*)
//Where413:=AnalyseCompte(Crit.St413,fbGene,FALSE,FALSE) ;
St1:='' ; St2:='' ; St3:='' ;
If Crit.St413EEP<>'' Then
  BEGIN
  Where413:=AnalyseCompte(Crit.St413EEP,fbGene,FALSE,FALSE) ;
  If Where413<>'' Then St1:=' OR '+Where413  ;
  END ;
If (Crit.St413EAR<>'') And (Crit.MethodeECTNE=TraiteEtFacture) Then
  BEGIN
  Where413:=AnalyseCompte(Crit.St413EAR,fbGene,FALSE,FALSE) ;
  If Where413<>'' Then St2:=' OR '+Where413 ;
  END ;
If (Crit.St411<>'') And (Crit.MethodeECTNE=TraiteEtFacture)  Then
  BEGIN
  Where413:=AnalyseCompte(Crit.St411,fbGene,FALSE,FALSE) ;
  If Where413<>'' Then St3:=' OR ('+Where413+' AND E_LETTRAGE="" AND MP_CATEGORIE="LCR")'  ;
  END ;
Where413:=St1+St2+St3 ; delete(Where413,1,3) ; St:=St+' AND ( '+Where413+') ' ;
//If Crit.DateCpta1<>0 Then St:=st+' AND E_DATECOMPTABLE>="'+USDATETIME(Crit.DateCpta1)+'" ' ;
//If Crit.DateCpta2<>0 Then St:=st+' AND E_DATECOMPTABLE<="'+USDATETIME(Crit.DateCpta2)+'" ' ;
If (Crit.MethodeECTNE<>Global) Then
  BEGIN
  If Crit.DateButoir<>0 Then St:=st+' AND E_DATEECHEANCE>"'+USDATETIME(Crit.DateButoir)+'" ' ;
//  St:=St+' AND E_DEBIT<>0 ' ; // 13459
  END ;
St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" AND (E_ECRANOUVEAU="N" OR (E_ECRANOUVEAU="H" AND E_EXERCICE="'+VH^.ExoV8.Code+'"))' ;
St:=St+' GROUP BY E_AUXILIAIRE ' ;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q),'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  Cpt:=Q.Fields[0].AsString ; i:=FEC.LR.IndexOf(Cpt) ;
  If i<0 Then
    BEGIN
    Existe:=FALSE ;
    EnCoursP:=TEnCoursP.Create ;
    EnCoursP.T.Cpt:=Q.FindField('E_AUXILIAIRE').AsString ;
    EnCoursP.T.Cpt1:=Q.FindField('E_AUXILIAIRE').AsString ;
    END Else
    BEGIN
    Existe:=TRUE ;
    EnCoursP:=TEnCoursP(FEC.LR.Objects[i]) ;
    END ;
  EnCoursP.T.Montant[eccPortefeuilleNonEchu].D:=Q.FindField('DP').AsFloat ;
  EnCoursP.T.Montant[eccPortefeuilleNonEchu].C:=Q.FindField('CP').AsFloat ;
  EnCoursP.T.Montant[eccPortefeuilleNonEchu].DE:=Q.FindField('DE').AsFloat ;
  EnCoursP.T.Montant[eccPortefeuilleNonEchu].CE:=Q.FindField('CE').AsFloat ;
  If Not Existe Then FEC.LR.AddObject(Cpt,EnCoursP) ;
  Q.Next ;
  END ;
FiniMove ;
Ferme(Q) ;
END ;

Procedure FaitEnCoursTraiteEchuNonRegle(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Var St,Cpt,StTL,Where413,St1,St2,St3 : String ;
    Q : TQuery ;
    i : Integer ;
    EnCoursP : TEnCoursP ;
    Existe : Boolean ;
BEGIN
st:='SELECT E_AUXILIAIRE, SUM(E_DEBIT) AS DP, SUM(E_CREDIT) AS CP, SUM(E_DEBITDEV) AS DD, SUM(E_CREDITDEV) AS CD, 0 AS DE, 0 AS CE FROM ECRITURE ' ;
St:=St+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
St:=St+' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
St:=St+' LEFT JOIN MODEPAIE ON E_MODEPAIE=MP_MODEPAIE ' ;
St:=St+' WHERE E_AUXILIAIRE<>"" AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD") ' ;
If Crit.Aux1<>'' Then St:=St+' AND E_AUXILIAIRE>="'+Crit.Aux1+'" ' ;
If Crit.Aux2<>'' Then St:=St+' AND E_AUXILIAIRE<="'+Crit.Aux2+'" ' ;
StTL:=FiltreTL(Crit) ; If StTL<>'' Then St:=St+StTL ;

Where413:=AnalyseCompte(Crit.St413EAR,fbGene,FALSE,FALSE) ;
If Where413<>'' Then St:=St+' AND '+Where413 ;

St1:='' ; St2:='' ; St3:='' ;
If Crit.St413EEP<>'' Then
  BEGIN
  Where413:=AnalyseCompte(Crit.St413EEP,fbGene,FALSE,FALSE) ;
  If Where413<>'' Then St1:=' OR '+Where413  ;
  END ;
If (Crit.St413EAR<>'') And (Crit.MethodeECTENR=TraiteEtFacture) Then
  BEGIN
  Where413:=AnalyseCompte(Crit.St413EAR,fbGene,FALSE,FALSE) ;
  If Where413<>'' Then St2:=' OR '+Where413 ;
  END ;
If (Crit.St411<>'') And (Crit.MethodeECTENR=TraiteEtFacture)  Then
  BEGIN
  Where413:=AnalyseCompte(Crit.St411,fbGene,FALSE,FALSE) ;
  If Where413<>'' Then
    BEGIN
    If Crit.OkTousMP Then St3:=' OR ('+Where413+' AND E_LETTRAGE="")'  
                     Else St3:=' OR ('+Where413+' AND E_LETTRAGE="" AND MP_CATEGORIE="LCR")'  ;
    END ;
  END ;
Where413:=St1+St2+St3 ; delete(Where413,1,3) ; St:=St+' AND ( '+Where413+') ' ;

//If Crit.DateCpta1<>0 Then St:=st+' AND E_DATECOMPTABLE>="'+USDATETIME(Crit.DateCpta1)+'" ' ;
//If Crit.DateCpta2<>0 Then St:=st+' AND E_DATECOMPTABLE<="'+USDATETIME(Crit.DateCpta2)+'" ' ;
If Crit.DateButoir<>0 Then St:=st+' AND E_DATEECHEANCE<="'+USDATETIME(Crit.DateButoir)+'" ' ;
St:=St+' AND E_LETTRAGE="" ' ;
// St:=St+' AND E_DEBIT<>0 ' ;  // 13459
//St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" ' ;
St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" AND (E_ECRANOUVEAU="N" OR (E_ECRANOUVEAU="H" AND E_EXERCICE="'+VH^.ExoV8.Code+'"))' ;
St:=St+' GROUP BY E_AUXILIAIRE ' ;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q),'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  Cpt:=Q.Fields[0].AsString ; i:=FEC.LR.IndexOf(Cpt) ;
  If i<0 Then
    BEGIN
    Existe:=FALSE ;
    EnCoursP:=TEnCoursP.Create ;
    EnCoursP.T.Cpt:=Q.FindField('E_AUXILIAIRE').AsString ;
    EnCoursP.T.Cpt1:=Q.FindField('E_AUXILIAIRE').AsString ;
    END Else
    BEGIN
    Existe:=TRUE ;
    EnCoursP:=TEnCoursP(FEC.LR.Objects[i]) ;
    END ;
  EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].D:=Q.FindField('DP').AsFloat ;
  EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].C:=Q.FindField('CP').AsFloat ;
  EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].DE:=Q.FindField('DE').AsFloat ;
  EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].CE:=Q.FindField('CE').AsFloat ;
  If Not Existe Then FEC.LR.AddObject(Cpt,EnCoursP) ;
  Q.Next ;
  END ;
FiniMove ;
Ferme(Q) ;
END ;


Procedure InitFEC(Var FEC : tFicEnCours) ;
BEGIN
FEC.LR:=HtStringList.Create ; FEC.LR.Sorted:=TRUE ; FEC.LR.Duplicates:=DupIgnore ;
END ;

procedure VideStringList ( L : HTStringList ) ;
Var i : integer ;
BEGIN
if L=Nil then Exit ; if L.Count<=0 then Exit ;
for i:=0 to L.Count-1 do If L.Objects[i]<>NIL Then L.Objects[i].Free ;
L.Clear ;
END ;


Procedure RAZFec(Var FEC : tFicEnCours) ;
BEGIN
VideStringList(FEC.LR) ; FEC.LR:=Nil ;
END ;

Function FaitCpt(Cpt,Cpt1,Masque : String) : String ;
Var i,l : Integer ;
BEGIN
Result:='' ; l:=Length(Cpt) ;
If Masque<>'' Then
  BEGIN
  For i:=1 To Length(Masque) Do
    BEGIN
    If (Masque[i]<>'-')  And (i<=l) Then
      BEGIN
      If (Masque[i]='X') Then Result:=Result+Cpt[i] Else Result:=Result+Masque[i] ;
      END ;
    END ;
  END Else Result:=Cpt1 ;
END ;

Function RecupFloat(X : Double ; Dec : Integer ; Sep : String) : String ;
Var St1,StF : String ;
    i : integer ;
BEGIN
StF:='' ; for i:=1 to Dec do StF:=StF+'0' ;
St1:=FormatFloat('#,##0.'+StF,X) ;
St1:=FindEtReplace(St1,V_PGI.SepMillier,'',TRUE) ;
St1:=FindEtReplace(St1,V_PGI.SepDecimal,sep,TRUE) ;
Result:=St1;
END ;

Function FormatSt(St : String ; L : Integer ; D : Boolean) : string ;
Var St1 : String ;
BEGIN
if Not D then Result:=Format_String(St,L) else
   BEGIN
   St1:=Trim(Copy(St,1,L)) ;
   While Length(St1)<L do St1:=' '+St1 ;
   Result:=St1 ;
   END ;
END ;

Function RecupDate(D: TDateTime ; FormatDate,SepDate : String ; Annee4 : Boolean) : String ;
Var J,M,A : Word ;
    St1,StS : String ;
BEGIN
DecodeDate(D,A,M,J) ;
if SEPDATE='' then StS:='' else StS:=SEPDATE ;
if FormatDate='JMA' then
   BEGIN
   St1:=FormatFloat('00',J)+StS+FormatFloat('00',M) ;
   if ANNEE4 then St1:=St1+StS+IntToStr(A) else St1:=St1+StS+Copy(IntToStr(A),3,2) ;
   END else
if FormatDate='MJA' then
   BEGIN
   St1:=FormatFloat('00',M)+StS+FormatFloat('00',J) ;
   if ANNEE4 then St1:=St1+StS+IntToStr(A) else St1:=St1+StS+Copy(IntToStr(A),3,2) ;
   END else
if FormatDate='AMJ' then
   BEGIN
   St1:=FormatFloat('00',M)+StS+FormatFloat('00',J) ;
   if ANNEE4 then St1:=IntToStr(A)+StS+St1 else St1:=Copy(IntToStr(A),3,2)+StS+St1 ;
   END ;
Result:=St1 ;
END ;



Function FaitEnCoursORLI(EnCoursP : TEnCoursP ; Var Crit : tCritExpECC) : String ;
Var St : String ;
    St1,Cpt,Sens : String ;
    Solde,D,C : Double ;
BEGIN
St:='' ;
//Cpt:=EnCoursP.T.Cpt1 ;
Cpt:=FaitCpt(EnCoursP.T.Cpt,EnCoursP.T.Cpt1,Crit.Masque) ;
St:=FormatSt(Cpt,17,FALSE) ;
D:=EnCoursP.T.Montant[eccComptable].D ;  C:=EnCoursP.T.Montant[eccComptable].C ;
If D<C Then Sens:='C' Else Sens:='D' ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'') ;
//Solde:=D-C ; St1:=RecupFloat(Solde,2,'') ;
St:=St+FormatSt(St1,10,TRUE)+Sens ;
D:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].D ; C:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].C ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'') ;
//Solde:=D-C ; St1:=RecupFloat(Solde,2,'') ;
St:=St+FormatSt(St1,10,TRUE) ;
D:=EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].D ; C:=EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].C ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'') ;
//Solde:=D-C ; St1:=RecupFloat(Solde,2,'') ;
St:=St+FormatSt(St1,10,TRUE) ;
St:=St+FormatSt(V_PGI.DevisePivot,3,FALSE) ;
Result:=St ;
END;

Function FaitEnCoursProduFlex(EnCoursP : TEnCoursP ; Var Crit : tCritExpECC) : String ;
Var St : String ;
    St1,Cpt,Sens : String ;
    Solde,D,C : Double ;
BEGIN
St:='' ;
//Cpt:=EnCoursP.T.Cpt1 ;
Cpt:=FaitCpt(EnCoursP.T.Cpt,EnCoursP.T.Cpt1,Crit.Masque) ;
St:=FormatSt(Cpt,17,FALSE) ;
D:=EnCoursP.T.Montant[eccComptable].D ;  C:=EnCoursP.T.Montant[eccComptable].C ;
If D<C Then Sens:='C' Else Sens:='D' ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'') ;
//Solde:=D-C ; St1:=RecupFloat(Solde,2,'') ;
St:=St+FormatSt(St1,10,TRUE)+Sens ;
D:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].D ; C:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].C ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'') ;
//Solde:=D-C ; St1:=RecupFloat(Solde,2,'') ;
St:=St+FormatSt(St1,10,TRUE) ;
D:=EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].D ; C:=EnCoursP.T.Montant[eccPortefeuilleEchuNonregle].C ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'') ;
//Solde:=D-C ; St1:=RecupFloat(Solde,2,'') ;
St:=St+FormatSt(St1,10,TRUE) ;
Result:=St ;
END;
(*
Function FaitEnCoursNegoce(EnCoursP : TEnCoursP) : String ;
Var St : String ;
    Debut : Boolean ;
    St1,Cpt : String ;
    Solde,D,C : Double ;
BEGIN
St:='' ;
Cpt:=EnCoursP.T.Cpt1 ; St:=FormatSt(Cpt,10,FALSE) ;
C:=EnCoursP.T.Montant[eccComptable].C ; St1:=RecupFloat(C,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;
C:=EnCoursP.T.Montant[eccComptable].CE ; St1:=RecupFloat(C,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;
D:=EnCoursP.T.Montant[eccComptable].D ; St1:=RecupFloat(D,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;
D:=EnCoursP.T.Montant[eccComptable].DE ; St1:=RecupFloat(D,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;

D:=EnCoursP.T.Montant[eccPortefeuille].D ; C:=EnCoursP.T.Montant[eccPortefeuille].C ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;

D:=EnCoursP.T.Montant[eccPortefeuille].DE ; C:=EnCoursP.T.Montant[eccPortefeuille].CE ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;

D:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].D ; C:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].C ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;

D:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].DE ; C:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].CE ;
Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,',') ; St:=St+FormatSt(St1,14,TRUE) ;

Result:=St ;
END;
*)

Function FaitEnCoursNegoce(EnCoursP : TEnCoursP ; Var Crit : tCritExpECC ; Var StSup : String) : String ;
Var St : String ;
    St1,Cpt : String ;
    Solde,D,C : Double ;
    SD1,SD2 : String ;
BEGIN
//Cpt:=EnCoursP.T.Cpt1 ;
Cpt:=FaitCpt(EnCoursP.T.Cpt,EnCoursP.T.Cpt1,Crit.Masque) ;
St:=St+FormatSt(Cpt,10,FALSE)+';' ;
SD1:=RecupDate(Crit.DateNeg1,'AMJ','',TRUE) ; SD2:=RecupDate(Crit.DateNeg2,'AMJ','',TRUE) ;
St:=St+SD1+';'+SD2+';' ;
StSup:='EFF;'+St ; St:='CPT;'+St ;
D:=EnCoursP.T.Montant[eccComptable].D ; C:=EnCoursP.T.Montant[eccComptable].C ;
//Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'.') ; St:=St+FormatSt(St1,13,TRUE)+';' ;
Solde:=D-C ; St1:=RecupFloat(Solde,2,'.') ; St:=St+FormatSt(St1,13,TRUE)+';' ;

D:=EnCoursP.T.Montant[eccComptable].DE ; C:=EnCoursP.T.Montant[eccComptable].CE ;
//Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'.') ; St:=St+FormatSt(St1,13,TRUE) ;
Solde:=D-C ; St1:=RecupFloat(Solde,2,'.') ; St:=St+FormatSt(St1,13,TRUE)+';' ;

D:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].D ; C:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].C ;
//Solde:=(Abs(D-C)) ;
Solde:=D-C ;
If Solde<>0 Then
  BEGIN
  St1:=RecupFloat(Solde,2,'.') ; StSup:=StSup+FormatSt(St1,13,TRUE)+';' ;

  D:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].DE ; C:=EnCoursP.T.Montant[eccPortefeuilleNonEchu].CE ;
//  Solde:=(Abs(D-C)) ; St1:=RecupFloat(Solde,2,'.') ; StSup:=StSup+FormatSt(St1,13,TRUE) ;
  Solde:=D-C ; St1:=RecupFloat(Solde,2,'.') ; StSup:=StSup+FormatSt(St1,13,TRUE) ;
  END else StSup:='' ;

Result:=St ;
END;

Procedure FaitFichierEnCours(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Var F : TextFile ;
    i : Integer ;
    EnCoursP : TEnCoursP ;
    St,St1 : String ;
BEGIN
AssignFile(F,FEC.NomFic) ;
{$I-} ReWrite (F) ; {$I+}
if IoResult<>0 then BEGIN HShowMessage('4;Exportation des en-cours;Le chemin d''accès du fichier n''est pas valide.;W;O;O;O;','','') ; Exit ; END ;
CloseFile(F) ;
{$I-} ReWrite (F) ; {$I+}
InitMove(FEC.LR.Count-1,'') ;
For i:=0 To FEC.LR.Count-1 Do
  BEGIN
  EnCoursP:=TEnCoursP(FEC.LR.Objects[i]) ;
  St1:='' ;
  If EnCoursP<>NIL Then
    BEGIN
    If FEC.Format='ORL' Then St:=FaitEnCoursOrli(EnCoursP,Crit) Else
     If FEC.Format='NEG' Then St:=FaitEnCoursNegoce(EnCoursP,Crit,St1) Else
      If FEC.Format='PRO' Then St:=FaitEnCoursProduflex(EnCoursP,Crit) ;
    If St<>'' Then Writeln(F,St) ; If St1<>'' Then Writeln(F,St1) ;
    END ;
  MoveCur(False) ;
  END ;
FiniMove ;
CloseFile(F) ;
END ;

Function FaitRequeteDetail(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) : String ;
Var St,StTL,StCpt : String ;
BEGIN
st:='SELECT T_CODEIMPORT, E_AUXILIAIRE, E_GENERAL, E_JOURNAL, E_NUMEROPIECE, E_DATECOMPTABLE, E_NATUREPIECE, ' ;
St:=St+' E_REFINTERNE, E_LIBELLE, E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV, E_COUVERTURE, E_COUVERTUREDEV, ';
St:=St+' E_DEVISE, E_COTATION, E_DATEECHEANCE, E_MODEPAIE FROM ECRITURE ' ;
St:=St+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
St:=St+' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
St:=St+' WHERE E_AUXILIAIRE<>"" AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD") ' ;
If Crit.Aux1<>'' Then St:=St+' AND E_AUXILIAIRE>="'+Crit.Aux1+'" ' ;
If Crit.Aux2<>'' Then St:=St+' AND E_AUXILIAIRE<="'+Crit.Aux2+'" ' ;
If Crit.DateCpta1<>0 Then St:=st+' AND E_DATECOMPTABLE>="'+USDATETIME(Crit.DateCpta1)+'" ' ;
If Crit.DateCpta2<>0 Then St:=st+' AND E_DATECOMPTABLE<="'+USDATETIME(Crit.DateCpta2)+'" ' ;
If Crit.DateModif1<>0 Then St:=st+' AND E_DATEMODIF>="'+USDATETIME(Crit.DateModif1)+'" ' ;
If Crit.DateModif2<>0 Then St:=st+' AND E_DATEMODIF<="'+USDATETIME(Crit.DateModif2)+'" ' ;
St:=St+' AND (E_ETATLETTRAGE="AL" ' ;
If Crit.OkPL Then St:=St+' OR E_ETATLETTRAGE="PL" ' ;
If Crit.OkTL Then St:=St+' OR E_ETATLETTRAGE="TL" ' ;
St:=St+') ' ;
StTL:=FiltreTL(Crit) ; If StTL<>'' Then St:=St+StTL ;
StCpt:='' ;
If (Crit.St411<>'') Then
  BEGIN
  StCpt:=AnalyseCompte(Crit.St411,fbGene,FALSE,FALSE) ;
  If StCpt<>'' Then St:=St+' AND '+StCpt+' ' ;
  END ;
St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" AND (E_ECRANOUVEAU="N" OR (E_ECRANOUVEAU="H" AND E_EXERCICE="'+VH^.ExoV8.Code+'")) ' ;
St:=St+' ORDER BY E_AUXILIAIRE, E_DATECOMPTABLE, E_NUMEROPIECE ' ;
Result:=St ;
END ;

Function EstCptPortefeuille(CptRub,ComptGene : String) : Boolean ;
Var St,StTemp : String ;
BEGIN
Result:=False ;
While CptRub<>'' do
  BEGIN
  St:=ReadTokenSt(CptRub) ; StTemp:='' ;
  if St='' then Continue ;
  if Pos('(',St)>0 then St:=Copy(St,1,Pos('(',St)-1) ;
  if Pos(':',St)>0 then BEGIN StTemp:=Copy(St,Pos(':',St)+1,200) ; Delete(St,Pos(':',St),200) ; END ;
  if StTemp<>'' then
     BEGIN
     St:=BourrelaDonc(St,fbGene) ; StTemp:=BourreLaDonc(StTemp,fbGene) ;
     if(ComptGene>=St) And (ComptGene<=StTemp) then
        BEGIN Result:=True ; Exit ; END ;
     END else
     BEGIN
     if Length(St)<=VH^.Cpta[fbGene].Lg then
        if(Pos(St,ComptGene)=1)Or(St=ComptGene) then
           BEGIN Result:=True ; Exit ; END ;
     END ;
  END ;
END ;

Procedure FaitFichierEnCoursDetail(Var FEC : tFicEnCours ; Var Crit : tCritExpECC) ;
Var F : TextFile ;
    St,Cpt,Sens : String ;
    Q : TQuery ;
    STM : String ;
    XX : Double ;
BEGIN
AssignFile(F,FEC.NomFic) ;
{$I-} ReWrite (F) ; {$I+}
if IoResult<>0 then BEGIN HShowMessage('4;Exportation des en-cours;Le chemin d''accès du fichier n''est pas valide.;W;O;O;O;','','') ; Exit ; END ;
CloseFile(F) ;
{$I-} ReWrite (F) ; {$I+}
St:=FaitRequeteDetail(FEC,Crit) ;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q)-1,'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ; St:='' ;
  XX:=Q.FindField('E_DEBIT').AsFloat+Q.FindField('E_CREDIT').AsFloat-Q.FindField('E_COUVERTURE').AsFloat ;
  XX:=Arrondi(XX,4) ;
  If XX<>0 Then
    BEGIN
    Cpt:=FaitCpt(Q.FindField('E_AUXILIAIRE').AsString,Q.FindField('T_CODEIMPORT').AsString,Crit.Masque) ;
    St:=FormatSt(Cpt,10,FALSE)+FormatSt(Q.FindField('E_JOURNAL').AsString,3,FALSE)+
        FormatSt(IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger),10,TRUE)+
        RecupDate(Q.FindField('E_DATECOMPTABLE').AsDateTime,'JMA','',TRUE)+
        FormatSt(Q.FindField('E_NATUREPIECE').AsString,2,FALSE)+
        FormatSt(Q.FindField('E_REFINTERNE').AsString,35,FALSE)+
        FormatSt(Q.FindField('E_LIBELLE').AsString,35,FALSE) ;
    Sens:='D' ; If Arrondi(Q.FindField('E_DEBIT').AsFloat,V_PGI.OkDecV)=0 Then Sens:='C' ;
    STM:=FormatSt(RecupFloat(XX,4,'.'),13,TRUE) ;
    St:=St+Sens+STM+FormatSt(Q.FindField('E_DEVISE').AsString,3,FALSE)+
        FormatSt(RecupFloat(Q.FindField('E_COTATION').AsFloat,5,'.'),13,TRUE) ;
    If Q.FindField('E_DEVISE').AsString<>V_PGI.DevisePivot Then
      BEGIN
      XX:=Q.FindField('E_DEBITDEV').AsFloat+Q.FindField('E_CREDITDEV').AsFloat-Q.FindField('E_COUVERTUREDEV').AsFloat ;
      XX:=Arrondi(XX,4) ; STM:=FormatSt(RecupFloat(XX,4,'.'),13,TRUE) ;
      END ;
    St:=St+StM+RecupDate(Q.FindField('E_DATEECHEANCE').AsDateTime,'JMA','',TRUE)+
        FormatSt(Q.FindField('E_MODEPAIE').AsString,3,FALSE) ;
    If EstCptPortefeuille(Crit.St413EEP,Q.FindField('E_GENERAL').AsString) Then St:=St+'O' Else St:=St+'N' ;
    END ;
  If St<>'' Then Writeln(F,St) ;
  Q.Next ;
  END ;
FiniMove ;
CloseFile(F) ;
END ;

Function EnCoursRegleNonEchu(Aux : String ; DateButoir : tDateTime) : Double ;
Var St,St1 : String ;
    Q : TQuery ;
BEGIN
Result:=0 ;
st:='SELECT SUM(E_DEBIT-E_CREDIT) AS SOLDE FROM ECRITURE ' ;
St:=St+' LEFT JOIN MODEPAIE ON E_MODEPAIE=MP_MODEPAIE ' ;
St:=St+' LEFT JOIN GENERAUX ON E_CONTREPARTIEGEN=G_GENERAL ' ;
St:=St+' WHERE E_AUXILIAIRE="'+Aux+'" ' ;
St:=St+' AND (MP_CATEGORIE="LCR") AND (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI") '  ;
If DateButoir<>0 Then St:=st+' AND E_DATEECHEANCE>"'+USDATETIME(DateButoir)+'" ' ;
St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" AND (E_ECRANOUVEAU="N" OR (E_ECRANOUVEAU="H" AND E_EXERCICE="'+VH^.ExoV8.Code+'"))' ;
St1:=LWhereV8 ; If St1<>'' Then St:=St+' AND '+St1 ;
Q:=OpenSQL(St,TRUE) ;
If Not Q.Eof Then Result:=Q.FindField('SOLDE').AsFloat ;
Ferme(Q) ;
END ;

end.
