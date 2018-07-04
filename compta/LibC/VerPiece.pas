{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 01/08/2003
Modifié le ... :   /  /    
Description .. : - 01/08/2003 - Contrôle lignes/pièces : tri des requêtes pour 
Suite ........ :  contrôle correct
Mots clefs ... : 
*****************************************************************}
unit VerPiece;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Mask, Hctrls, Buttons, ExtCtrls, Ent1, HEnt1, CRITEDT, hmsgbox, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  HStatus, Rapsuppr, RappType, HSysMenu;

procedure VerifPiece ;

type  TINFOPIECE = Record
      NumeroPiece  : Integer ;
      NumLigne     : Integer ;
      NumVentil    : Integer ;
      NumEche    : Integer ;
      END ;

type
  TFVerPiece = class(TForm)
    Panel1: TPanel;
    HPB: TPanel;
    TTravail: TLabel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BStop: THBitBtn;
    TFNumPiece2: TLabel;
    TFNumPiece1: THLabel;
    FNumPiece1: TMaskEdit;
    FNumPiece2: TMaskEdit;
    QE: TQuery;
    QA: TQuery;
    MsgRien: THMsgBox;
    MsgBar: THMsgBox;
    MsgLibel: THMsgBox;
    MsgLibel2: THMsgBox;
    Panel2: TPanel;
    TNBError1: TLabel;
    TNBError2: TLabel;
    QAnaIPur: TQuery;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BValiderClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
  private
    { Déclarations privées }
    OkVerif, StopVerif,
    DejaLance, LanceDirect,
    PremierPassage, ParLaForme : Boolean ;
    LaListe                    : TList ;
    NbEnreg, NbError,
    ErrEcr, ErrAnal,
    NumP1, NumP2,
    NumP11, NumP12,
    AvantPiece, AvantLigne,
    AvantVentil, AvantEche  : Integer ;
    Function GoListe1(P,L : Integer ; Quel, Rem : String ; I : Byte) : DelInfo ;
    Function GoListe2(P,L : Integer ; Quel, Rem : String ; F : TField) : DelInfo ;
    Function NumeroOk : Boolean ;
    Function SiEnreg : Boolean ;
    Procedure InitLabel  ;
    Procedure Bouton(Etat : Boolean) ;
    Procedure LanceVerif  ;
    Procedure SqlEcr  ;
    Procedure SqlAnalNonPur  ;
    Procedure SqlAnalPur  ;
    function  TestBreak : Boolean ;
    function  PieceOk(M : TINFOPIECE) : Boolean ;
    function  PieceAnalNonPurOk(M : TINFOPIECE) : Boolean ;
    function  PieceAnalPurOk(M : TINFOPIECE) : Boolean ;
    function  DonneInfo(Pi,Li,DeQui : Integer) : TInfoMvt ;
    //Procedure PrepareAnal  ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

procedure VerifPiece ;
var VerPiece : TFVerPiece ; Maj : Boolean ;
BEGIN
VerPiece:=TFVerPiece.Create(Application) ;
try
 With VerPiece do
      BEGIN
      If Not (V_PGI.SAV) then
         BEGIN
         ParLaForme:=False ; OkVerif:=True ; NbError:=0 ;
         LaListe.Clear ; StopVerif:=False ;  Maj:=false ; LanceDirect:=False ;
         NumP1:=0 ;
         NumP2:=999999999 ;
         If SiEnreg then
            BEGIN
            LanceVerif ;
            if Not OkVerif then RapportdErreurMvt(Laliste,3,Maj,False) ;
            END ;
         END else ShowModal ;
      END ;
 finally
 VerPiece.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;


procedure TFVerPiece.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
DejaLance:=False ; InitLabel ; ParLaForme:=True ;
end;

procedure TFVerPiece.FormCreate(Sender: TObject);
begin
LaListe:=TList.Create ;
end;

procedure TFVerPiece.FormClose(Sender: TObject; var Action: TCloseAction);
begin
LaListe.Free ;
end;

procedure TFVerPiece.BValiderClick(Sender: TObject);
Var MAJ : Boolean ;
begin
OkVerif:=True ; StopVerif:=False ;
ErrEcr:=0 ; ErrAnal:=0 ; MAJ:=False ;
if NumeroOk then
   BEGIN
   LaListe.Clear ;
   TNBError1.Caption:=MsgBar.Mess[14] ;
   Bouton(False) ; Application.ProcessMessages ;
   if SiEnreg then
      BEGIN
      DejaLance:=True ;
      LanceVerif ;
      if Not OkVerif then
         BEGIN
         RapportdErreurMvt(Laliste,3,MAJ,False) ;
         END Else if not StopVerif then MsgRien.Execute(2,'','') ; // sinon message tout est ok
      END Else MsgRien.Execute(0,'','') ; // message : aucun enreg
   InitLabel ;
   Bouton(True) ;
   END Else MsgRien.Execute(1,'','') ;
end;

Function TFVerPiece.NumeroOk : Boolean ;
BEGIN
Result:=(StrToInt(FNumPiece1.Text)<=StrToInt(FNumPiece2.Text)) ;
if Result then
   BEGIN
   NumP1:=StrToInt(FNumPiece1.Text) ;
   NumP2:=StrToInt(FNumPiece2.Text) ;
   If Not DejaLance then begin NumP11:=NumP1 ; NumP12:=NumP2 ; end
                    Else LanceDirect:=((NumP11=NumP1)and(NumP12=NumP2)) ;

   END ;
END ;

procedure TFVerPiece.BStopClick(Sender: TObject);
begin
StopVerif:=True ;
end;

function TFVerPiece.TestBreak : Boolean ;
BEGIN
Application.ProcessMessages ;
if StopVerif then if MsgRien.Execute(3,'','')<>mryes then StopVerif:=False ;
Result:=StopVerif ;
END ;


Function TFVerPiece.SiEnreg : Boolean ;
BEGIN
If LanceDirect then begin Result:=True ;InitMove(NbEnreg,MsgBar.Mess[0]) ; Exit ; end ;
NbEnreg:=0 ;
QE.Close ; QE.SQL.Clear ;
(*
QE.SQL.Add(' Select E_NUMEROPIECE, E_NUMLIGNE ') ;
QE.SQL.Add(' From ECRITURE Where E_QUALIFPIECE="N" ') ;
QE.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(NumP1)+' and E_NUMEROPIECE<='+IntToStr(NumP2)+' ') ;
QE.SQL.Add(' Group by E_NUMEROPIECE, E_NUMLIGNE ') ;
*)
QE.SQL.Add(' Select E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
QE.SQL.Add(' From ECRITURE Where E_QUALIFPIECE="N" ') ;
QE.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(NumP1)+' and E_NUMEROPIECE<='+IntToStr(NumP2)+' ') ;
QE.SQL.Add(' Group by E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
QE.SQL.Add(' order by E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
ChangeSql(QE) ; //QE.Prepare ;
PrepareSQLODBC(QE) ;
QE.Open ; NbEnreg:=NbEnreg+QE.RecordCount ;
QA.Close ; QA.SQL.Clear ;
QA.SQL.Add(' Select  Y_NUMEROPIECE, Y_NUMVENTIL ') ;
QA.SQL.Add(' From ANALYTIQ Where Y_TYPEANALYTIQUE<>"X" and Y_QUALIFPIECE="N"') ;
QA.SQL.Add(' and Y_NUMEROPIECE>='+IntToStr(NumP1)+' and Y_NUMEROPIECE<='+IntToStr(NumP2)+' ') ;
QA.SQL.Add(' Group by Y_NUMEROPIECE, Y_NUMVENTIL ') ;
QA.SQL.Add(' order by Y_NUMEROPIECE, Y_NUMVENTIL ') ;
ChangeSql(QA) ; //QA.Prepare ;
PrepareSQLODBC(QA) ;
QA.Open ; NbEnreg:=NbEnreg+QA.RecordCount ;
QAnaIPur.Close ; QAnaIPur.SQL.Clear ;
QAnaIPur.SQL.Add(' Select  Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL ') ;
QAnaIPur.SQL.Add(' From ANALYTIQ Where Y_TYPEANALYTIQUE="X" and Y_QUALIFPIECE="N"') ;
QAnaIPur.SQL.Add(' And Y_NUMEROPIECE>='+IntToStr(NumP1)+' and Y_NUMEROPIECE<='+IntToStr(NumP2)+' ') ;
QAnaIPur.SQL.Add(' Group by  Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL ') ;
QAnaIPur.SQL.Add(' Order by  Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL ') ;
ChangeSql(QAnaIPur) ; //QAnaIPur.Prepare ;
PrepareSQLODBC(QAnaIPur) ;
QAnaIPur.Open ; NbEnreg:=NbEnreg+QAnaIPur.RecordCount ;
Application.ProcessMessages ;
InitMove(NbEnreg,MsgBar.Mess[0]) ;
SiEnreg:=Not (NbEnreg=0) ;
END ;

Procedure TFVerPiece.InitLabel ;
BEGIN
TTravail.Caption:='' ;
TNBError1.Caption:='' ; TNBError2.Caption:='' ;
TNBError1.Font.color:=clNavy ; TNBError2.Font.color:=clNavy ;
END ;

Procedure TFVerPiece.Bouton(Etat : Boolean) ;
Var i : Byte ;
BEGIN
BStop.Cancel:=Not Etat ;
for i:=0 to ComponentCount-1 do
    BEGIN
    if Components[i] is TBitBtn then
       BEGIN
       if ((UpperCase(TBitBtn(Components[i]).Name)='BANNULER') or
           (UpperCase(TBitBtn(Components[i]).Name)='BFERME') or
           (UpperCase(TBitBtn(Components[i]).Name)='BAIDE') or
           (UpperCase(TBitBtn(Components[i]).Name)='BVALIDER')) then TBitBtn(Components[i]).Enabled:=Etat ;
       END else
    if Components[i] is TCustomEdit then TCustomEdit(Components[i]).Enabled:=Etat ;
    END ;
END ;

Procedure TFVerPiece.LanceVerif  ;
BEGIN
TNBError1.Caption:='' ; NbError:=0 ; { Init de NbError}
TTravail.Caption:=MsgBar.Mess[18] ;
Application.ProcessMessages ;
if not StopVerif then SqlEcr ;
if not StopVerif then SqlAnalNonPur ;
if not StopVerif then SqlAnalPur ;
FiniMove ;
END ;

function TFVerPiece.DonneInfo(Pi,Li,DeQui : Integer) : TInfoMvt ;
Var Q : TQuery ; Y : TInfoMvt ;
BEGIN                         { Cherche Info pour afficher dans la grille d'erreurs }
Y:=TInfoMvt.Create ;          { 1 = Ecr ; 2 = Anal Non Pur ; 3 = Anal Pur }
Q:=NIL ;
If DeQui=1 then
   BEGIN
   Q:=OpenSql('select E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_REFINTERNE from ECRITURE '
             +'where E_NUMEROPIECE='+intToStr(Pi)+' and E_NUMLIGNE='+intToStr(Li)+'  '
             +' Order by E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE ',True) ;
   If Not Q.Eof then
      BEGIN
      Y.Journal:=Q.FindField('E_JOURNAL').AsString ;
      Y.EXERCICE:=Q.FindField('E_EXERCICE').AsString ;
      Y.DATECOMPTABLE:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
      Y.REFINTERNE:=Q.FindField('E_REFINTERNE').AsString ;
      Y.Numeropiece:=Pi ;
      Y.Numligne:=Li ;
      END ;
   END Else
If DeQui=2 then
   BEGIN
   Q:=OpenSql('select Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_REFINTERNE, Y_AXE from analytiq '
             +'where Y_NUMEROPIECE='+intToStr(Pi)+' and Y_NUMLIGNE='+intToStr(Li)+'  '
             +' Order by Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_REFINTERNE, Y_AXE ',True) ;
   if not Q.Eof then
      BEGIN
      Y.Journal:=Q.FindField('Y_JOURNAL').AsString ;
      Y.EXERCICE:=Q.FindField('Y_EXERCICE').AsString ;
      Y.DATECOMPTABLE:=Q.FindField('Y_DATECOMPTABLE').AsDateTime ;
      Y.REFINTERNE:=Q.FindField('Y_REFINTERNE').AsString ;
      Y.AXE:=Q.FindField('Y_AXE').AsString ;
      Y.Numeropiece:=Pi ;
      Y.Numligne:=Li ;
      END ;
   END Else
If DeQui=3 then   { Ici, Li == NumVentil }
   BEGIN
   Q:=OpenSql('select Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_REFINTERNE, Y_AXE, Y_NUMLIGNE from analytiq '
             +'where Y_NUMEROPIECE='+intToStr(Pi)+' and Y_NUMVENTIL='+intToStr(Li)+'  '
             +' Order by Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL ',True) ;
   if not Q.Eof then
      BEGIN
      Y.Journal:=Q.FindField('Y_JOURNAL').AsString ;
      Y.EXERCICE:=Q.FindField('Y_EXERCICE').AsString ;
      Y.DATECOMPTABLE:=Q.FindField('Y_DATECOMPTABLE').AsDateTime ;
      Y.REFINTERNE:=Q.FindField('Y_REFINTERNE').AsString ;
      Y.AXE:=Q.FindField('Y_AXE').AsString ;
      Y.Numeropiece:=Pi ;
      Y.Numligne:=Q.FindField('Y_NUMLIGNE').AsInteger ;
      END ;
   END ;
If Q<>NIL Then Ferme(Q) ;
Result:=Y ;
END ;

Function TFVerPiece.GoListe1(P,L : Integer ; Quel, Rem : String ; I : Byte ) : DelInfo ;
Var X : DelInfo ;
    Entitee : TInfoMvt ;
BEGIN
Result:=NIL ;
If (Quel='E') then
   BEGIN
   Entitee:=DonneInfo(P,L,1) ;
   END Else
If (Quel='A') then
   BEGIN
   Entitee:=DonneInfo(P,L,2) ;
   END Else
If (Quel='ANP') then
   BEGIN
   Entitee:=DonneInfo(P,L,3) ;
   END Else Exit ;

Inc(NbError) ; X:=DelInfo.Create ;
X.LeCod:=Entitee.Journal ; X.LeLib:=Entitee.REFINTERNE ;
X.LeMess:=IntToStr(Entitee.Numeropiece)+'/'+IntToStr(Entitee.Numligne) ;
X.LeMess2:=DateToStr(Entitee.datecomptable) ; X.LeMess3:=MsgLibel.Mess[i]+' '+Rem ;
if (Quel='A')or(Quel='ANP') then X.LeMess4:=Entitee.exercice+';'+Entitee.Axe else X.LeMess4:=Entitee.exercice ;
Result:=X ;
END ;

Function TFVerPiece.GoListe2(P,L : Integer ; Quel, Rem : String ; F : TField ) : DelInfo ;
Var X : DelInfo ; Entitee : TInfoMvt ;
BEGIN
Result:=NIL ;
If (Quel='E') then
   BEGIN
   Entitee:=DonneInfo(P,L,1) ;
   END Else
If (Quel='A') then
   BEGIN
   Entitee:=DonneInfo(P,L,2) ;
   END Else
If (Quel='ANP') then
   BEGIN
   Entitee:=DonneInfo(P,L,3) ;
   END Else Exit ;
Inc(NbError) ; X:=DelInfo.Create ;
X.LeCod:=Entitee.Journal ; X.LeLib:=Entitee.REFINTERNE ;
X.LeMess:=IntToStr(Entitee.Numeropiece)+'/'+IntToStr(Entitee.Numligne) ;
X.LeMess2:=DateToStr(Entitee.datecomptable) ; X.LeMess3:=MsgLibel2.Mess[0]+' "'+F.FieldName+'" '+MsgLibel2.Mess[1]+' '+Rem ;
if (Quel='A')or(Quel='ANP') then X.LeMess4:=Entitee.exercice+';'+Entitee.Axe else X.LeMess4:=Entitee.exercice ;
Result:=X ;
END ;

Procedure TFVerPiece.SqlEcr  ;
Var Mvt : TINFOPIECE ; i : Byte ;
BEGIN
QE.First ;
PremierPassage:=True ; i:=0 ;
While not QE.Eof do
      BEGIN
      MoveCur(False) ;
      Fillchar(Mvt,SizeOf(Mvt),#0) ;
      Mvt.NumeroPiece:=QE.FindField('E_NUMEROPIECE').AsInteger ;
      Mvt.NumLigne:=QE.FindField('E_NumLigne').AsInteger ;
      Mvt.NUMECHE:=QE.FindField('E_NUMECHE').AsInteger ;
      If i=0 then
         BEGIN
         AvantPiece:=Mvt.NumeroPiece+1 ;
         AvantLigne:=Mvt.NumLigne+1 ;
         AvantEche:=Mvt.NUMECHE+1 ;
         END ;
      if Not PieceOk(Mvt) then OkVerif:=False  ;
      AvantPiece:=Mvt.NumeroPiece ;
      AvantLigne:=Mvt.NumLigne ;
      AvantEche:=Mvt.NumEche ;
      PremierPassage:=False ;
      if TestBreak then Break ;
      i:=1;QE.Next;
      END ;
END ;

function TFVerPiece.PieceOk(M : TINFOPIECE) : Boolean ;
Var NewPiece : Boolean ;
BEGIN
Result:=True ;
NewPiece:=(M.NumeroPiece<>AvantPiece) ;
if M.NumeroPiece<1 then
   BEGIN
   LaListe.Add(GoListe2(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(M.NumeroPiece)+') ',QE.FindField('E_NUMEROPIECE'))) ;
   Result:=False ;
   END ;
if M.NumLigne<1 then
   BEGIN
   LaListe.Add(GoListe2(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(M.NumLigne)+') ',QE.FindField('E_NUMLIGNE'))) ;
   Result:=False ;
   END ;
If NewPiece then
   BEGIN
   if M.NumLigne<>1 then
      BEGIN
      LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(M.NumLigne)+') ',1)) ;
      Result:=False ;
      END ;
   if Not PremierPassage then
      if Not ((AvantPiece+1)=M.NumeroPiece) then
         BEGIN
         LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(AvantPiece)+')'+' ('+IntToStr(M.NumeroPiece)+') ',0));
         Result:=False ;
         END ;
   END Else
   BEGIN
   If (M.NumeroPiece<=NumP2) then
      BEGIN
      if (AvantPiece=M.NumeroPiece) then
         BEGIN
         if (AvantLigne=M.NumLigne) then
            BEGIN
            If Not((AvantEche+1)=M.NumEche) then
               BEGIN
               LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(AvantEche)+')'+' ('+IntToStr(M.NumEche)+') ',7)) ;
               Result:=False ;
               END ;
            END Else
            BEGIN
            if Not((AvantLigne+1)=M.NumLigne) then
               BEGIN
               LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(AvantLigne)+')'+' ('+IntToStr(M.NumLigne)+') ',1)) ;
               Result:=False ;
               END ;
            END ;
         END ;
      END ;
   END ;
END ;

Procedure TFVerPiece.SqlAnalNonPur  ;
Var Mvt : TINFOPIECE ; i : Integer ;
BEGIN
QA.First ;
i:=0 ; PremierPassage:=True ;
While not QA.Eof do
      BEGIN
      MoveCur(False) ;
      Fillchar(Mvt,SizeOf(Mvt),#0) ;
      Mvt.NumeroPiece:=QA.FindField('Y_NUMEROPIECE').AsInteger ;
      Mvt.NumVentil:=QA.FindField('Y_NUMVENTIL').AsInteger ;
      if i=0 then
         begin
         AvantPiece:=Mvt.NumeroPiece+1 ;
         AvantVentil:=Mvt.NumVentil+1 ;
         end ;
      if Not PieceAnalNonPurOk(Mvt) then OkVerif:=False  ;
      AvantPiece:=Mvt.NumeroPiece ;
      AvantVentil:=Mvt.NumVentil ; PremierPassage:=False ;
      if TestBreak then Break ;
      i:=1 ; QA.Next;
      END ;
END ;

function TFVerPiece.PieceAnalNonPurOk(M : TINFOPIECE) : Boolean ;
Var NewPiece : Boolean ;
BEGIN
Result:=True ;
NewPiece:=(M.NumeroPiece<>AvantPiece) ;
If NewPiece then
   BEGIN
   If M.NumVentil<>1 then
      BEGIN
      LaListe.Add(GoListe1(M.NumeroPiece,M.NumVentil,'ANP',' ('+IntToStr(M.NumVentil)+') ',6));
      Result:=False ;
      END ;
   END Else
   BEGIN
   If M.NumeroPiece=AvantPiece then
      BEGIN
      If Not((AvantVentil+1)=M.NumVentil) then
         BEGIN
         LaListe.Add(GoListe1(M.NumeroPiece,M.NumVentil,'ANP',' ('+IntToStr(AvantVentil)+')'+' ('+IntToStr(M.NumVentil)+') ',6)) ;
         Result:=False ;
         END ;
      END ;
   END ;
END ;

Procedure TFVerPiece.SqlAnalPur  ;
Var Mvt : TINFOPIECE ; i : byte ;
BEGIN
QAnaIPur.First ;
i:=0 ; PremierPassage:=True ;
While not QAnaIPur.Eof do
      BEGIN
      MoveCur(False) ;
      Fillchar(Mvt,SizeOf(Mvt),#0) ;
      Mvt.NumeroPiece:=QAnaIPur.FindField('Y_NUMEROPIECE').AsInteger ;
      Mvt.NumLigne:=QAnaIPur.FindField('Y_NumLigne').AsInteger ;
      Mvt.NumVentil:=QAnaIPur.FindField('Y_NUMVENTIL').AsInteger ;
      if i=0 then
         begin
         AvantPiece:=Mvt.NumeroPiece+1 ;
         AvantLigne:=Mvt.NumLigne+1 ;
         AvantVentil:=Mvt.NumVentil+1 ;
         end ;
      if Not PieceAnalPurOk(Mvt) then OkVerif:=False  ;
      AvantPiece:=Mvt.NumeroPiece ;
      AvantLigne:=Mvt.NumLigne ;
      AvantVentil:=Mvt.NumVentil ; PremierPassage:=False ;
      if TestBreak then Break ;
      i:=1 ; QAnaIPur.Next;
      END ;
END ;

function TFVerPiece.PieceAnalPurOk(M : TINFOPIECE) : Boolean ;
Var NewPiece : Boolean ;
BEGIN
Result:=True ;
NewPiece:=(M.NumeroPiece<>AvantPiece) ;
if M.NumeroPiece<1 then
   BEGIN
   LaListe.Add(GoListe2(M.NumeroPiece,M.NumLigne,'A',' ('+IntToStr(M.NumeroPiece)+') ',QA.FindField('Y_NUMEROPIECE'))) ;
   Result:=False ; exit ;
   END ;
(*
Rony 29/04/97
if M.NumLigne<>0 then
   BEGIN
   LaListe.Add(GoListe2(M.NumeroPiece,M.NumLigne,'A',' ('+IntToStr(M.NumLigne)+') ',QA.FindField('Y_NUMLIGNE'))) ;
   Result:=False ;
   END ;
*)
if M.NumVentil<1 then
   BEGIN
   LaListe.Add(GoListe2(M.NumeroPiece,M.NumLigne,'A',' ('+IntToStr(M.NumVentil)+') ',QA.FindField('Y_NUMVENTIL'))) ;
   Result:=False ; exit ;
   END ;

If NewPiece then
   BEGIN
   If M.NumVentil<>1 then
      BEGIN
      LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'A',' ('+IntToStr(M.NumVentil)+') ',5));
      Result:=False ;
      END ;
If Not PremierPassage then
   if Not ((AvantPiece+1)=M.NumeroPiece) then
      BEGIN
      // Rony 23/04/97 ---> A priori, pas de test sur les numeros de pièces logiques....
      //LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'A',' ('+IntToStr(AvantPiece)+')'+' ('+IntToStr(M.NumeroPiece)+') ',3)) ;
      //Result:=False ;
      END ;
   END Else
   BEGIN
   If (M.NumeroPiece<=NumP2) then   {}
      BEGIN
      if (AvantPiece=M.NumeroPiece) then
         BEGIN
         If Not((AvantVentil+1)=M.NumVentil) then
            BEGIN
            LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'A',' ('+IntToStr(AvantVentil)+')'+' ('+IntToStr(M.NumVentil)+') ',5)) ;
            Result:=False ;
            END ;
         END ;
      END ;
   (* Rony 29/04/97
   If (M.NumeroPiece<=NumP2) then   {}
      BEGIN
      if (AvantPiece=M.NumeroPiece) then
         BEGIN
         if (AvantLigne=M.NumLigne) then
            BEGIN
            If Not((AvantVentil+1)=M.NumVentil) then
               BEGIN
               LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'A',' ('+IntToStr(AvantVentil)+')'+' ('+IntToStr(M.NumVentil)+') ',5)) ;
               Result:=False ;
               END ;
            END ;
         END ;
      END ;
   *)
   END ;
END ;

(*
function TFVerPiece.PieceOk(M : TINFOPIECE) : Boolean ;
Var NewPiece : Boolean ;
BEGIN
Result:=True ;
NewPiece:=(M.NumeroPiece<>AvantPiece) ;
if M.NumeroPiece<1 then
   BEGIN
   LaListe.Add(GoListe2(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(M.NumeroPiece)+') ',QE.FindField('E_NUMEROPIECE'))) ;
   Result:=False ;
   END ;
if M.NumLigne<1 then
   BEGIN
   LaListe.Add(GoListe2(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(M.NumLigne)+') ',QE.FindField('E_NUMLIGNE'))) ;
   Result:=False ;
   END ;
If NewPiece then
   BEGIN
   if M.NumLigne<>1 then
      BEGIN
      LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(M.NumLigne)+') ',1)) ;
      Result:=False ;
      END ;
   if Not PremierPassage then
      if Not ((AvantPiece+1)=M.NumeroPiece) then
         BEGIN
         LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(AvantPiece)+')'+' ('+IntToStr(M.NumeroPiece)+') ',0));
         Result:=False ;
         END ;
   END Else
   BEGIN
   If (M.NumeroPiece<=NumP2) then
      BEGIN
      if (AvantPiece=M.NumeroPiece) then
         BEGIN
         if Not((AvantLigne+1)=M.NumLigne) then
            BEGIN
            LaListe.Add(GoListe1(M.NumeroPiece,M.NumLigne,'E',' ('+IntToStr(AvantLigne)+')'+' ('+IntToStr(M.NumLigne)+') ',1)) ;
            Result:=False ;
            END ;
         END ;
      END ;
   END ;
END ;
*)
end.
