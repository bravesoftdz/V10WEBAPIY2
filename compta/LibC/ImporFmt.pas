unit ImporFmt;

interface

uses DB,{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} HCtrls, TImpFic,Ent1,SysUtils;


Procedure EcrireFormat(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TQuery) ;
Function ChargeFormat(Var Fichier : Text ; FileName : String ; Nature,Import,Code : String ; Var Entete : TFmtEntete ; Var Detail : TTabFmtDetail ; Var Debut : Boolean ) : boolean ;
Procedure LireFormat(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TDataSet) ;


implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  HEnt1;

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

Function ChargeFormat(Var Fichier : Text ; FileName : String ; Nature,Import,Code : String ; Var Entete : TFmtEntete ; Var Detail : TTabFmtDetail ; Var Debut : Boolean ) : boolean ;
Var QEntete,QDetail : TQuery ;
    L : Integer ;
BEGIN
Result:=FALSE ; Debut:=TRUE ;
QEntete:=OpenSQL('SELECT * FROM FMTIMPOR WHERE IM_NATURE="'+Nature+'" AND IM_IMPORTATION="'+Import+'" AND IM_CODE="'+Code+'"',TRUE) ;
if QEntete.EOF then BEGIN Ferme(QEntete) ; Exit ; END ;
QDetail:=OpenSQL('SELECT * FROM FMTIMPDE WHERE ID_NATURE="'+Nature+'" AND ID_IMPORTATION="'+Import+'" AND ID_CODE="'+Code+'" ORDER BY ID_NUMCHAMP',TRUE) ;
if QDetail.EOF then BEGIN Ferme(QDetail) ; Exit ; END ;
Entete.IMPORTATION:=(QEntete.FindField('IM_IMPORTATION').AsString='X') ;
Entete.FORMATDATE:=QEntete.FindField('IM_FORMATDATE').AsString ;
Entete.ANNEE4:=(QEntete.FindField('IM_ANNEE4').AsString='X') ;
Entete.CRLF:=(QEntete.FindField('IM_CRLF').AsString='X') ;
Entete.ASCII:=(QEntete.FindField('IM_ASCII').AsString='X') ;
Entete.MOINSDEVANT:=(QEntete.FindField('IM_MOINSDEVANT').AsString='X') ;
Entete.SEPDATE:=QEntete.FindField('IM_SEPDATE').AsString[1] ;
Entete.SEPDECIMAL:=QEntete.FindField('IM_SEPDECIMAL').AsString[1] ;
Entete.SEPMILLIER:=QEntete.FindField('IM_SEPMILLIER').AsString[1] ;
Entete.DELIMITEUR:=QEntete.FindField('IM_DELIMITEUR').AsString[1] ;
if Entete.DELIMITEUR='9' then Entete.DELIMITEUR:=#9 ;
Entete.IGNORELIGNE:=QEntete.FindField('IM_IGNORELIGNE').AsInteger ;
Entete.TAILLEREC:=QEntete.FindField('IM_TAILLEREC').AsInteger ;
Entete.FORMATSENS:=QEntete.FindField('IM_FORMATSENS').AsInteger ;
Entete.NbLigne:=0 ; Entete.NbChamps:=0 ; Entete.Longueur:=0 ;
While Not QDetail.EOF do
   BEGIN
   L:=QDetail.FindField('ID_NUMCHAMP').AsInteger ;
   Detail[L].NumLigne:=QDetail.FindField('ID_NUMLIGNE').AsInteger ;
   if Detail[L].NumLigne>Entete.NbLigne then Entete.NbLigne:=Detail[L].NumLigne ;
   if L>Entete.NbChamps then Entete.NbChamps:=L ;
   Detail[L].Champ:=QDetail.FindField('ID_CHAMP').AsString ;
   Detail[L].Debut:=QDetail.FindField('ID_DEBUT').AsInteger ;
   Detail[L].Longueur:=QDetail.FindField('ID_LONGUEUR').AsInteger ;
   Detail[L].Decimal:=QDetail.FindField('ID_DECIMAL').AsInteger ;
   Detail[L].Corresp:=QDetail.FindField('ID_CORRESP').AsString ;
   Detail[L].Droite:=(QDetail.FindField('ID_ALIGNEDROITE').AsString='X') ;
   if Detail[L].Longueur+Detail[L].Debut>Entete.Longueur then Entete.Longueur:=Detail[L].Longueur+Detail[L].Debut ;
   QDetail.Next ;
   END ;
if (Entete.Importation) AND (Not FileExists(FileName)) then
    BEGIN
    Ferme(QEntete) ; Ferme(QDetail) ; Exit ;
    END ;
Result:=TRUE ;
AssignFile(Fichier,FileName) ;
if (Entete.Importation) then reset(Fichier) else rewrite(Fichier) ;
Ferme(QEntete) ; Ferme(QDetail) ;
END ;

Function TransSt(St1 : String ; Entete : TFmtEntete ; Detail : TFmtDetail ) : String ;
Var St,Sta,Stb : String ;
BEGIN
Result:=Trim(St1) ;
if Detail.Corresp<>'' then
   BEGIN
   St:=Detail.Corresp ;
   if Pos('#',St)<>0 then BEGIN Result:=ASCII2ANSI(Result) ; Exit ; END ;
   While St<>'' do
      BEGIN
      Sta:=ReadTokenSt(St) ; Stb:=ReadTokenSt(St) ;
      if St1=Sta then BEGIN Result:=Stb ; exit ; END ;
      END ;
   END else
   BEGIN
   if Entete.ASCII then  Result:=ASCII2ANSI(Result) ;
   END ;
END ;

Function TransInt(St1 : String ; Entete : TFmtEntete ; Detail : TFmtDetail ) : Integer ;
BEGIN
if Entete.SepMillier<>'B' then
  St1:=FindEtReplace(St1,Entete.SepMillier,'',TRUE) ;
if (St1<>'') and (Not Entete.MoinsDevant) AND (St1[Length(St1)]='-') then
   St1:='-'+Copy(St1,1,Length(St1)-1) ;
Result:=Round(Valeur(St1)) ;
END ;

Function TransFloat(St1 : String ; Entete : TFmtEntete ; Detail : TFmtDetail ) : Double ;
Var X : Double ;
    i : integer ;
BEGIN
if Entete.SepDecimal<>'B' then
  St1:=FindEtReplace(St1,Entete.SepDecimal,V_PGI.SepDecimal,TRUE) ;
if Entete.SepMillier<>'B' then
  St1:=FindEtReplace(St1,Entete.SepMillier,'',TRUE) ;
if (St1<>'') and (Not Entete.MoinsDevant) AND (St1[Length(St1)]='-') then
   St1:='-'+Copy(St1,1,Length(St1)-1) ;
X:=Valeur(St1) ;
if Entete.SepDecimal='B' then
  for i:=1 to Detail.Decimal do X:=X/10 ;
Result:=Arrondi(X,Detail.Decimal) ;
END ;

Function TransDate(St1 : String ; Entete : TFmtEntete ; Detail : TFmtDetail ) : TDateTime ;
Var J,M,A : String ;
    i,An : Integer ;
BEGIN
if Entete.SEPDATE<>'B' then i:=1 else i:=0 ;
if Entete.FormatDate='JMA' then
   BEGIN
   J:=Copy(St1,1,2) ;
   M:=Copy(St1,3+i,2) ;
   if Entete.ANNEE4 then A:=Copy(St1,5+2*i,4) else A:=Copy(St1,5+2*i,2) ;
   END else
if Entete.FormatDate='MJA' then
   BEGIN
   M:=Copy(St1,1,2) ;
   J:=Copy(St1,3+i,2) ;
   if Entete.ANNEE4 then A:=Copy(St1,5+2*i,4) else A:=Copy(St1,5+2*i,2) ;
   END else
if Entete.FormatDate='AMJ' then
   BEGIN
   if Entete.ANNEE4 then
      BEGIN
      A:=Copy(St1,1,4) ; M:=Copy(St1,5+i,2) ; J:=Copy(St1,7+2*i,2) ;
      END else
      BEGIN
      A:=Copy(St1,1,2) ; M:=Copy(St1,3+i,2) ; J:=Copy(St1,5+2*i,2) ;
      END ;
   END ;
An:=StrToInt(A) ;
// Année sur 2 caractères
// Pour les initialisations de dates
// et l'an 2000 ????
if An=0 then An:=1900 else
 if An<80 then An:=2000+An else
  if ((An>79) and (An<100)) then An:=1900+An ;
A:=IntToStr(An) ;
Result:=EncodeDate(StrToInt(A),StrToInt(M),StrToInt(J)) ;
END ;


Procedure ExtractChamp(St1 : String ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; L : Integer ; Q : TDataSet) ;
Var C,i : Integer ;
    St,StC : String ;
    T,T2 : TField ;
    OkDebit,OkSens : Boolean ;

BEGIN
OkDebit:=TRUE ;
For C:=1 to Entete.NbChamps do
  if Detail[C].NumLigne=L then
    BEGIN
    OkSens:=(Detail[C].Champ='XX_SENS') ;
    if Entete.DELIMITEUR='B' then
       BEGIN
       T:=Q.FindField(Detail[C].Champ) ;
       if (Not OkSens) And (T=Nil) then Continue ;
       St:=Copy(St1,Detail[C].Debut,Detail[C].Longueur) ;
       END else
       BEGIN
       i:=Pos(Entete.DELIMITEUR,St1) ; St:=Copy(St1,1,i-1) ; Delete(St1,1,i) ;
       T:=Q.FindField(Detail[C].Champ) ;
       if (Not OkSens) And (T=Nil) then Continue ;
       END ;
    if Trim(St)<>'' then St:=Trim(St) else continue ;
    if OkSens then OkDebit:=(TransSt(St,Entete,Detail[C])='D') else
    Case T.DataType of
       ftString : if St<>'' then T.AsString:=TransSt(St,Entete,Detail[C]) ;
       ftSmallint, ftInteger, ftWord : if St<>'' then T.AsInteger:=TransInt(St,Entete,Detail[C]) ;
       ftFloat,ftCurrency, ftBCD : if St<>'' then T.AsFloat:=TransFloat(St,Entete,Detail[C]) ;
       ftDate, ftTime, ftDateTime : if St<>'' then T.AsDateTime:=TransDate(St,Entete,Detail[C]) ;
// Pour les longchar de budjal....
       ftMemo : if St<>'' then T.AsString:=T.AsString+TransSt(St,Entete,Detail[C]) ;
       END ;
    END ;
if Entete.FormatSens=2 then
For C:=1 to Entete.NbChamps do
  if Detail[C].NumLigne=L then
    if (Pos('DEBIT',Detail[C].Champ)>0) AND (Not OKDebit) then
       BEGIN
       StC:=FindEtReplace(Detail[C].Champ,'DEBIT','CREDIT',FALSE) ;
       T:=Q.FindField(Detail[C].Champ) ; T2:=Q.FindField(StC) ; if (T=NIL) or (T2=NIL) then continue ;
       T2.AsFloat:=T.AsFloat ; T.AsFloat:=0 ;
       END else
    if (Pos('CREDIT',Detail[C].Champ)>0) AND (OkDebit) then
       BEGIN
       StC:=FindEtReplace(Detail[C].Champ,'CREDIT','DEBIT',FALSE) ;
       T:=Q.FindField(Detail[C].Champ) ; T2:=Q.FindField(StC) ; if (T=NIL) or (T2=NIL) then continue ;
       T2.AsFloat:=T.AsFloat ; T.AsFloat:=0 ;
       END ;
END ;


Procedure LireFormatCRLF(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TDataSet) ;
Var St : String ;
    i,L : integer ;
BEGIN
if Debut then
   BEGIN
   Debut:=FALSE ;
   For i:=1 to Entete.IGNORELIGNE do
     if Not EOF(Fichier) then readln(fichier,st) ;
   END ;
For l:=1 to Entete.NbLigne do if Not EOF(Fichier) then
   BEGIN
   Readln(Fichier,St) ;
   ExtractChamp(St,Entete,Detail,L,Q) ;
   END ;
END ;

Procedure LireFormatFixe(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TDataSet) ;
Var St : String ;
    i,L,j : integer ;
    Ch : Char ;
BEGIN
if Debut then
   BEGIN
   Debut:=FALSE ;
   For i:=1 to Entete.IGNORELIGNE do
      BEGIN
      St:='' ;
      For j:=1 to Entete.TailleRec do
         if Not EOF(Fichier) then BEGIN Read(fichier,Ch) ; St:=St+Ch ; END ;
      END ;
   END ;
For l:=1 to Entete.NbLigne do if Not EOF(Fichier) then
   BEGIN
   St:='' ;
   For j:=1 to Entete.TailleRec do
      if Not EOF(Fichier) then BEGIN Read(fichier,Ch) ; St:=St+Ch ; END ;
   ExtractChamp(St,Entete,Detail,L,Q) ;
   END ;
END ;

Procedure LireFormat(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TDataSet) ;
BEGIN
if Entete.CRLF then LireFormatCRLF(Fichier,Entete,Detail,Debut,Q)
               else LireFormatFixe(Fichier,Entete,Detail,Debut,Q) ;
END ;

(*===================================== EXPORT ===================================*)
Function RecupCorresp(St1 : String ; Detail : TFmtDetail ;Q : TDataset) : String ;
Var St,Sta,Stb : String ;
    StElse : String ;
    p,p1 : integer ;
    Deb,Lon : Integer ;
BEGIN
Result:=St1 ;
if Detail.Corresp<>'' then
   BEGIN
   St:=Detail.Corresp ;
   StElse:='' ;
   p:=Pos('#',St) ;
   if p<>0 then
     BEGIN
     p1:=Pos(':',St) ;
     Deb:=StrToInt(Copy(St,p+1,p1-(p+1))) ;
     Lon:=StrToInt(Copy(St,p1+1,Length(St)-p1)) ;
     Result:=Copy(St1,Deb,Lon) ;
     Exit ;
     END ;
   p:=Pos('[',St) ;
   if p<>0 then
     BEGIN
     if Q.FindField(copy(St,p+1,Length(St)-2))<>nil then Result:=Q.FindField(copy(St,p+1,Length(St)-2)).AsString ;
     Exit ;
     END ;
   While St<>'' do
     BEGIN
     Sta:=ReadTokenSt(St) ; Stb:=ReadTokenSt(St) ;
     if Sta='ELSE' then StElse:=Stb ;
     if TRIM(St1)=TRIM(Sta) then BEGIN Result:=Stb ; exit ; END ;
     END ;
   if StElse<>'' then Result:=StElse ;
   END ;
END ;

Function RecupSt(St1 : String ; Entete : TFmtEntete ; Detail : TFmtDetail ) : String ;
BEGIN
Result:=Trim(St1) ;
END ;

Function RecupInt(I : Integer ; Entete : TFmtEntete ; Detail : TFmtDetail ) : String ;
Var St1 : String ;
BEGIN
St1:=FormatFloat('#,##0',i) ;
if Entete.SepMillier='B' then St1:=FindEtReplace(St1,V_PGI.SepMillier,'',TRUE)
                         else St1:=FindEtReplace(St1,V_PGI.SepMillier,Entete.SepMillier,TRUE) ;
if (St1<>'') and (Not Entete.MoinsDevant) AND (St1[1]='-') then
   St1:=Copy(St1,2,Length(St1)-1)+'-' ;
Result:=St1 ;
END ;

Function RecupFloat(X : Double ; Entete : TFmtEntete ; Detail : TFmtDetail ) : String ;
Var St1,StF : String ;
    i : integer ;
BEGIN
StF:='' ; for i:=1 to Detail.Decimal do StF:=StF+'0' ;
St1:=FormatFloat('#,##0.'+StF,X) ;
if Entete.SepMillier='B' then St1:=FindEtReplace(St1,V_PGI.SepMillier,'',TRUE)
                         else St1:=FindEtReplace(St1,V_PGI.SepMillier,Entete.SepMillier,TRUE) ;
if Entete.SepDecimal='B' then St1:=FindEtReplace(St1,V_PGI.SepDecimal,'',TRUE)
                         else St1:=FindEtReplace(St1,V_PGI.SepDecimal,Entete.SepDecimal,TRUE) ;
if (St1<>'') and (Not Entete.MoinsDevant) AND (St1[1]='-') then
   St1:=Copy(St1,2,Length(St1)-1)+'-' ;
Result:=St1;
END ;

Function RecupDate(D: TDateTime ; Entete : TFmtEntete ; Detail : TFmtDetail ) : String ;
Var J,M,A : Word ;
    St1,StS : String ;
BEGIN
DecodeDate(D,A,M,J) ;

if Entete.SEPDATE='B' then StS:='' else StS:=Entete.SEPDATE ;
if Entete.FormatDate='JMA' then
   BEGIN
   St1:=FormatFloat('00',J)+StS+FormatFloat('00',M) ;
   if Entete.ANNEE4 then St1:=St1+StS+IntToStr(A) else St1:=St1+StS+Copy(IntToStr(A),3,2) ;
   END else
if Entete.FormatDate='MJA' then
   BEGIN
   St1:=FormatFloat('00',M)+StS+FormatFloat('00',J) ;
   if Entete.ANNEE4 then St1:=St1+StS+IntToStr(A) else St1:=St1+StS+Copy(IntToStr(A),3,2) ;
   END else
if Entete.FormatDate='AMJ' then
   BEGIN
   St1:=FormatFloat('00',M)+StS+FormatFloat('00',J) ;
   if Entete.ANNEE4 then St1:=IntToStr(A)+StS+St1 else St1:=Copy(IntToStr(A),3,2)+StS+St1 ;
   END ;
Result:=St1 ;
END ;


Function RecupChamp(Entete : TFmtEntete ; Detail : TTabFmtDetail ; L : Integer ; Q : TQuery) : string ;
Var C,iSens,id,ic : Integer ;
    St,StC,St1,StD,St2,Sta,StB : String ;
    T,T2 : TField ;
    X : Double ;

BEGIN
iSens:=0 ; St1:='' ;
if Entete.DELIMITEUR='B' then St1:=Format_String(' ',entete.Longueur) ;
For C:=1 to Entete.NbChamps do
  if Detail[C].NumLigne=L then
    BEGIN
    if (Detail[C].Champ='XX_SENS') then
       BEGIN
       iSens:=C ; St:='@@@' ;
       END else
       BEGIN
       T:=Q.FindField(Detail[C].Champ) ;
       if T<>NIL then
         Case T.DataType of
           ftString : St:=RecupSt(T.AsString,Entete,Detail[C]) ;
           ftSmallint, ftInteger, ftWord : St:=RecupInt(T.AsInteger,Entete,Detail[C]) ;
           ftFloat,ftCurrency, ftBCD : St:=RecupFloat(T.AsFloat,Entete,Detail[C]) ;
           ftDate, ftTime, ftDateTime : St:=RecupDate(T.AsDateTime,Entete,Detail[C]) ;
           ftMemo : St:=RecupSt(T.AsString,Entete,Detail[C]) ;
           END else St:='' ;
       St:=RecupCorresp(St,Detail[C],Q) ;
       END ;
    if Entete.ASCII then St:=ANSI2ASCII(St) ;
    if Entete.DELIMITEUR='B' then St1:=Insere(St1,FormatSt(St,Detail[C].Longueur,Detail[C].Droite),Detail[C].Debut,Detail[C].Longueur)
                             else St1:=St1+St+Entete.DELIMITEUR ;
    END ;
if Entete.FormatSens>=2 then
For C:=1 to Entete.NbChamps do
  if Detail[C].NumLigne=L then
     BEGIN
     id:=Pos('DEBIT',Detail[C].Champ) ;
     ic:=Pos('CREDIT',Detail[C].Champ) ;
     if (Ic>0) or (id>0) then
       BEGIN
       StD:=Detail[C].Champ ; StC:=Detail[C].Champ ;
       if ic>0 then StD:=FindEtReplace(Detail[C].Champ,'CREDIT','DEBIT',FALSE)
               else StC:=FindEtReplace(Detail[C].Champ,'DEBIT','CREDIT',FALSE) ;
       T:=Q.FindField(StD) ; T2:=Q.FindField(StC) ;
       if (T=NIL) or (T2=NIL) then continue ;
       if (iSens>0) And (Entete.FormatSens=2) then
          BEGIN
//          if T.AsFloat>0 then X:=T.AsFloat else X:=T2.AsFloat ;
          if Arrondi(T.AsFloat,5)<>0 then X:=T.AsFloat else X:=T2.AsFloat ;
          St:=RecupFloat(X,Entete,Detail[C]) ; St:=RecupCorresp(St,Detail[C],Q) ;
          StC:='C' ; StD:='D' ;
          if Detail[iSens].Corresp<>'' then
             BEGIN
             St2:=Detail[iSens].Corresp ;
             while St2<>'' do
                BEGIN
                Sta:=ReadTokenSt(St2) ; Stb:=ReadTokenSt(St2) ;
                if Stb='D' then StD:=Sta ;
                if Stb='C' then StC:=Sta ;
                END ;
             END ;
          if Entete.ASCII then St:=ANSI2ASCII(St) ;
          if Entete.DELIMITEUR='B' then St1:=Insere(St1,FormatSt(St,Detail[C].Longueur,Detail[C].Droite),Detail[C].Debut,Detail[C].Longueur)
                                   else St1:=St1+St+Entete.DELIMITEUR ;
//          if T.AsFloat>0 then St:=StD else St:=StC ;
          if Arrondi(T.AsFloat,5)<>0 then St:=StD else St:=StC ;
          St1:=FindEtReplace(St1,Copy('@@@',1,Detail[isens].Longueur),St,TRUE) ;
          END else
          BEGIN
          if (Entete.FormatSens=3) then
             BEGIN
             if T.AsFloat>0 then X:=-T.AsFloat else X:=T2.AsFloat ;
             END else
             BEGIN
             if T.AsFloat>0 then X:=T.AsFloat else X:=-T2.AsFloat ;
             END ;
          St:=RecupFloat(X,Entete,Detail[C]) ; St:=RecupCorresp(St,Detail[C],Q) ;
          if Entete.ASCII then St:=ANSI2ASCII(St) ;
          if Entete.DELIMITEUR='B' then St1:=Insere(St1,FormatSt(St,Detail[C].Longueur,Detail[C].Droite),Detail[C].Debut,Detail[C].Longueur)
                                   else St1:=St1+St+Entete.DELIMITEUR ;
          END ;
       END ;
     END ;
Result:=St1 ;
END ;



Procedure EcrireFormatCRLF(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TQuery) ;
Var St : String ;
    i,L : integer ;
BEGIN
if Debut then
   BEGIN
   Debut:=FALSE ; St:='ENTETE' ;
   For i:=1 to Entete.IGNORELIGNE do Writeln(fichier,st) ;
   END ;
For l:=1 to Entete.NbLigne do
   BEGIN
   St:=RecupChamp(Entete,Detail,L,Q) ;
   Writeln(Fichier,St) ;
   END ;
END ;

Procedure EcrireFormatFixe(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TQuery) ;
Var St : String ;
    i,L : integer ;
BEGIN
if Debut then
   BEGIN
   Debut:=FALSE ; St:=Format_String('ENTETE',Entete.TailleRec) ;
   For i:=1 to Entete.IGNORELIGNE do Write(fichier,st) ;
   END ;
For l:=1 to Entete.NbLigne do
   BEGIN
   St:=RecupChamp(Entete,Detail,L,Q) ;
   St:=Format_String(St,Entete.TailleRec) ;
   Write(fichier,st) ;
   END ;
END ;



Procedure EcrireFormat(Var Fichier : Text ; Entete : TFmtEntete ; Detail : TTabFmtDetail ; Var Debut : Boolean ; Q : TQuery) ;
BEGIN
if Entete.CRLF then EcrireFormatCRLF(Fichier,Entete,Detail,Debut,Q)
               else EcrireFormatFixe(Fichier,Entete,Detail,Debut,Q) ;
END ;


end.
