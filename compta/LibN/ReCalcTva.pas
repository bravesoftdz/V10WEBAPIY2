unit ReCalcTva;

interface

uses
  Windows, Messages, SysUtils, 
  {$IFNDEF EAGLSERVER}
  Forms,
  Dialogs,
  Graphics,
  Controls,
  Mask,
  HSysMenu,
  Buttons,
  HStatus,
  hmsgbox ,
  {$ENDIF}
  StdCtrls, Hcompte, Ent1, HEnt1,  Hctrls,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   ExtCtrls, TimpFic, ImpFicU,
{$IFDEF EAGLSERVER}
  Paramsoc,
{$ENDIF}
  Classes, UTob
   ;

{$IFNDEF EAGLSERVER}
Procedure CalcLaTva ;
Procedure CalculTvaImport;
{$ENDIF}
procedure CCalcTvaSimplifiee(  vCodeExo, Wheres : string; WhereJournal : string='' ) ;

{$IFNDEF EAGLSERVER}

type

  TFCalcTva = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    TFDateCpta1: THLabel;
    FDateCpta1: TMaskEdit;
    TFDateCpta2: TLabel;
    FDateCpta2: TMaskEdit;
    TFEtab: THLabel;
    FEtab: THValComboBox;
    TFNumPiece1: THLabel;
    FNumPiece1: TMaskEdit;
    FNumPiece2: TMaskEdit;
    TFExercice: THLabel;
    FExercice: THValComboBox;
    HPB: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    Label1: TLabel;
    FCbV: TCheckBox;
    FCbA: TCheckBox;
    FCbHT: TCheckBox;
    HMess: THMsgBox;
    procedure FExerciceChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    InfoImp : TInfoImport ;
    Procedure LanceTraitement ;
  public
    { Déclarations publiques }
  end;
{$ENDIF}

implementation

{$IFNDEF EAGLSERVER}
{$R *.DFM}
{$ENDIF}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPTypeCons,
  {$ENDIF MODENT1}
  SaisUtil, SaisComm, Ed_Tools ;

{$IFNDEF EAGLSERVER}
Procedure CalcLaTva ;
var TT : TFCalcTva;
BEGIN
TT:=TFCalcTva.Create(Application) ;
 Try
  TT.ShowModal ;
 Finally
  TT.Free ;
 End ;
END ;

Procedure CalculTvaImport;
var TT : TFCalcTva;
BEGIN
TT:=TFCalcTva.Create(Application) ;
With TT do
begin
  Try
   BeginTrans ;
   LanceTraitement ;
   CommitTrans ;
  Except
   RollBack ;
  End ;
  Free ;
end;
end;
{$ENDIF}

{$IFDEF A SUPPRIMER}
Procedure TraitementSurEcr ( OkTva : Boolean ; M : RMVT ; ForceFlag : boolean ; Var InfoImp : TInfoImport ; QFiche : TQFiche) ;
Var TPiece : TList ;
    Q : TQuery ;
    O : TOB ;
    NbTiers,Lig,LigTiers,NumBase,i : integer ;
    SorteTva  : TSorteTva ;
    ExigeTva  : TExigeTva ;
    Solde,TTC,Coef : double ;
    Okok,OkTiers,ExisteEnc : boolean ;
    NaturePiece,NatureJal,RegimeTva,RegimeEnc,StE,NatGene,CodeTva : String3 ;
    Auxi,Gene,StUp,SQL,Coll : String ;
    TabTvaEnc : Array[1..5] of Double ;
    CptLu,CptLuJ : TCptLu ;
BEGIN
{#TVAENC}

{$IFDEF NOVH}
if not GetParamSocSecur('SO_OUITVAENC',False) then exit;
{$ELSE}
if Not VH^.OuiTvaEnc then Exit ;
{$ENDIF}

if M.Jal='' then Exit ;
Okok:=True ; ExisteEnc:=False ;
{Identification nature facture}
NaturePiece:=M.Nature ;
if ((NATUREPIECE<>'FC') and (NATUREPIECE<>'AC') and (NATUREPIECE<>'FF') and (NATUREPIECE<>'AF')) then OkTva:=FALSE ;
If (Not OkTva) Then Exit ;
{Identification journal facture}
NatureJal:='' ; TTC:=0 ;
CptLuJ.Cpt:=M.Jal ;
//If ChercheCptLu(InfoImp.LJalLu,CptLuJ) Then NatureJal:=CptLuJ.Nature ;
If AlimLTabCptLu(3,QFiche[3],InfoImp.LJalLu,NIL,CptLuJ) Then NatureJal:=CptLuJ.Nature ;
SorteTva:=stvDivers ;
if NatureJal='VTE' then SorteTva:=stvVente else if NatureJal='ACH' then SorteTva:=stvAchat ;
if SorteTva=stvDivers then OkTva:=FALSE ;
{Constitution d'une liste des lignes de l'écriture}
TPiece:=TList.Create ;
Q:=OpenSQL('Select * from ECRITURE Where '+WhereEcriture(tsGene,M,False),True) ;
While Not Q.EOF do
   BEGIN
   CptLu.Cpt:=Q.FindField('E_GENERAL').AsString ; AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,NIL,CptLu) ;
   CptLu.Cpt:=Q.FindField('E_AUXILIAIRE').AsString ; AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,NIL,CptLu) ;
   O:=TOB.Create('ECRITURE',Nil,-1) ;
   O.SelectDB('',Q,True) ; TPiece.Add(O) ;
   Q.Next ;
   END ;
Ferme(Q) ;
{TVA : Détermination du nombre de tiers}
{CTR : Détermination des contreparties}
NbTiers:=0 ; LigTiers:=-1 ;
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ; OkTiers:=False ;
    if ((O.GetValue('E_ECHE')='X') and (O.GetValue('E_NUMECHE')>=1)) then
       BEGIN
                                          // fiche 10537
             if (O.GetValue('E_NUMECHE')=1) and (O.GetValue('E_TYPEMVT')='TTC') then
             begin
                    OkTiers:=True ;
                    TTC:=Arrondi(TTC+O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT'),V_PGI.OkDecV) ;
             end;
       END ;
    if OkTiers then
       BEGIN
       Inc(NbTiers) ;
       if NbTiers=1 then LigTiers:=Lig else BEGIN Okok:=False ; Break ; END ;
       END ;
    END ;
if ((Not Okok) or (NbTiers<>1) or (LigTiers<0) or (TTC=0)) then OkTva:=FALSE ;
If (Not OkTva) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Lecture paramètres tiers}
O:=TOBM(TPiece[LigTiers]) ;
Coll:=O.GetValue('E_GENERAL') ; if Not EstCollFact(Coll) then OkTva:=FALSE ;
If (Not OkTva) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
RegimeTva:='' ; RegimeEnc:='' ;
if O.GetValue('E_AUXILIAIRE')<>'' then
   BEGIN
   Auxi:=O.GetValue('E_AUXILIAIRE') ;
   CptLu.Cpt:=Auxi ;
   If ChercheCptLu(InfoImp.LAuxLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
   END else
   BEGIN
   Gene:=O.GetValue('E_GENERAL') ;
   CptLu.Cpt:=Gene ;
   If ChercheCptLu(InfoImp.LGenLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
   (*
   Q:=OpenSQL('Select G_REGIMETVA, G_TVAENCAISSEMENT from GENERAUX Where G_GENERAL="'+Gene+'"',True) ;
   if Not Q.EOF then BEGIN RegimeTva:=Q.Fields[0].AsString ; RegimeEnc:=Q.Fields[1].AsString ; END ;
   Ferme(Q) ;
   *)
   CptLu.Cpt:=Gene ;
   If ChercheCptLu(InfoImp.LGenLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
   END ;
if ((RegimeTva='') or ((SorteTva=stvAchat) and (RegimeEnc=''))) then OkTva:=FALSE ;
If (Not OkTva) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Affectation exigibilité Tva}
if SorteTva=stvAchat then ExigeTva:=Code2Exige(RegimeEnc) else ExigeTva:=tvaMixte ;
{Si Fournisseur débit --> débit donc rien à traiter}
if ExigeTva=tvaDebit then OkTva:=FALSE ;
If (Not OkTva) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Affectation des régimes, des tva,tpf et enc O/N sur les lignes}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOB(TPiece[Lig]) ; Gene:=O.GetValue('E_GENERAL') ;
    CptLu.Cpt:=Gene ;
    If ChercheCptLu(InfoImp.LGenLu,CptLu) Then
      BEGIN
      NatGene:=CptLu.Nature ;
      if ((NatGene='CHA') or (NatGene='PRO') or (NatGene='IMO')) then
         BEGIN
         If OkTva And ((O.GetValue('E_TVA')='') Or ForceFlag) Then O.PutValue('E_TVA',CptLu.Tva) ;
         If OkTva And ((O.GetValue('E_TPF')='') Or ForceFlag) Then O.PutValue('E_TPF',CptLu.Tpf) ;
         If (O.GetValue('E_TVAENCAISSEMENT')='') Or ForceFlag Then
           BEGIN
           StE:=FlagEncais(ExigeTva,(CptLu.TvaEncHT='X')) ;
           END else StE:=O.GetValue('E_TVAENCAISSEMENT') ;
         if StE='X' then ExisteEnc:=True ;
          if O.Getvalue('E_REGIMETVA') = '' then
         O.PutValue('E_REGIMETVA',RegimeTva) ;
         If OkTva And (O.GetValue('E_TVAENCAISSEMENT')<>StE) Then O.PutValue('E_TVAENCAISSEMENT',StE) ;
         END else
         BEGIN
         If OkTva Then O.PutValue('E_TVA','') ;
         END ;
      END ;
    END ;
{Si aucune ligne en encaissement --> aucun traitement}
if Not ExisteEnc then OkTva:=FALSE ;
If (Not OkTva) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Calcul du tableau des bases Tva}
FillChar(TabTvaEnc,Sizeof(TabTvaEnc),#0) ;
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    CodeTva:=O.GetValue('E_TVA') ; if CodeTva='' then Continue ;
    if SorteTva=stvVente then Solde:=O.GetValue('E_CREDIT')-O.GetValue('E_DEBIT')
                         else Solde:=O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT') ;
    if O.GetValue('E_TVAENCAISSEMENT')='X' then
       BEGIN
       NumBase:=Tva2NumBase(CodeTva) ;
       if NumBase>0 then TabTvaEnc[NumBase]:=TabTvaEnc[NumBase]+Solde ;
       END else
       BEGIN
       TabTvaEnc[5]:=TabTvaEnc[5]+Solde ;
       END ;
    END ;
{Report du tableau des bases sur les échéances}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    {b FP 31/11/2005 FQ17049: Il faut recalculer pour toutes les écritures tiers
    if ((O.GetMvt('E_ECHE')='X') and (O.GetMvt('E_NUMECHE')>0)) then}
   // ajout me fiche 10537 if O.GetMvt('E_AUXILIAIRE')<>'' then
    if ((O.GetValue('E_ECHE')='X') and (O.GetValue('E_NUMECHE')>0)) and (O.GetValue('E_TYPEMVT')='TTC') then
       BEGIN
       Solde:=O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT') ;
       Coef:=Solde/TTC ;
       for i:=1 to 4 do If OkTva Then O.PutValue('E_ECHEENC'+IntToStr(i),Arrondi(TabTvaEnc[i]*Coef,V_PGI.OkDecV)) ;
       If OkTva Then O.PutValue('E_ECHEDEBIT',Arrondi(TabTvaEnc[5]*Coef,V_PGI.OkDecV)) ;
       If OkTva Then O.PutValue('E_EMETTEURTVA','X') ;
       END ;
    END ;
{Mise à jour fichier}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOB(TPiece[Lig]) ;
    if Not O.IsOneModifie then Continue ;
    O.SetAllModifie(True) ;
    SQL:='UPDATE ECRITURE SET '+StUp+' Where  '+WhereEcriture(tsGene,M,False)
        +' AND E_NUMLIGNE='+IntToStr(O.GetValue('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetValue('E_NUMECHE')) ;
    if ExecuteSQL(SQL)<>1 then BEGIN V_PGI.IoError:=oeUnknown ; Break ; END ;
    END ;
{Dispose mémoire}
VideListe(TPiece) ; TPiece.Free ;
END ;
{$ENDIF A SUPPRIMER}

Function WhereJal(NatJal, Where : String) : String ;
Var Q : TQuery ;
    St : String ;
BEGIN
St:='' ;
Q:=OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL="'+NatJal+'"' + Where,TRUE) ;
While Not Q.Eof Do
  BEGIN
  If ST='' Then St:=' E_JOURNAL="'+Q.Fields[0].AsString+'" ' Else St:=St+' Or E_JOURNAL="'+Q.Fields[0].AsString+'" ' ;
  Q.Next ;
  END ;
Ferme(Q) ;
If St<>'' Then St:=' ('+St+') ' ;
Result:=St ;
END ;



{$IFNDEF EAGLSERVER}
procedure TFCalcTva.FExerciceChange(Sender: TObject);
begin
If FExercice.ItemIndex>0 Then ExoToDates(FExercice.Value,FDateCpta1,FDateCpta2) ;
end;

procedure TFCalcTva.FormShow(Sender: TObject);
begin
FExercice.Value:=VH^.Encours.Code ;
end;


Procedure TFCalcTva.LanceTraitement ;
Var Q,Q1 : TQuery ;
    St,StA,StV : String ;
    QFiche : TQFiche ;
    JalP,ExoP,QualP : String ;
    DateP : TDateTime ;
    NumP : Integer ;
    M : Rmvt ;
BEGIN
If (Not FcbV.Checked) And (Not FcbA.Checked) Then Exit ;
StA:='' ; StV:='' ;
If (FcbA.Checked) Then StA:=WhereJal('ACH', '') ; If (FcbV.Checked) Then StV:=WhereJal('VTE', '') ;
If (StA='') And (StV='') Then Exit ;
St:='SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_QUALIFPIECE FROM ECRITURE ' ;
St:=St+' WHERE E_JOURNAL<>"'+w_w+'" AND E_QUALIFPIECE="N" ' ;
if FExercice.ItemIndex>0 then St:=St+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
St:=St+' And E_DATECOMPTABLE>="'+USDATE(FDateCpta1)+'" ' ;
St:=St+' And E_DATECOMPTABLE<="'+USDATE(FDateCpta2)+'" ' ;
if FEtab.ItemIndex>0 then St:=St+' AND E_ETABLISSEMENT="'+FEtab.Value+'" ' ;
St:=St+' And E_NUMEROPIECE>='+FNumPiece1.Text+' and E_NUMEROPIECE<='+FNumPiece2.Text+' ' ;
St:=St+' AND (' ;
If (StV<>'') And (StA<>'') Then St:=St+StA+' Or '+StV+') '
                           Else St:=St+StA+StV+') ' ;
St:=St+' GROUP BY E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_QUALIFPIECE ' ;
Q:=OpenSQL(St,TRUE) ;
//InitRequete(QFiche[0],0) ; InitRequete(QFiche[1],1) ; InitRequete(QFiche[3],3) ;
If Not Q.Eof Then InitMove(RecordsCount(Q),'') ;
While Not Q.Eof Do
  BEGIN
  Movecur(FALSE);
  JalP:=Q.FindField('E_JOURNAL').AsString ;
  ExoP:=Q.FindField('E_EXERCICE').AsString ;
  DateP:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  NumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  QualP:=Q.FindField('E_QUALIFPIECE').AsString ;
  Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+JalP+'"'
             +' AND E_EXERCICE="'+ExoP+'"'
             +' AND E_DATECOMPTABLE="'+USDateTime(DateP)+'"'
             +' AND E_NUMEROPIECE='+IntToStr(NumP)
             +' AND E_QUALIFPIECE="'+QualP+'" ',True) ;
  if Not Q1.EOF then
    BEGIN
    M:=MvtToIdent(Q1,fbGene,False) ; Ferme(Q1) ;
//    TraitementSurEcr(TRUE,M,FCbHT.Checked,InfoImp,QFiche) ;
    ElementsTvaEncRevise (M, FCbHT.Checked, InfoImp, QFiche) ;
    END Else Ferme(Q1) ;
  Q.Next ;
  END ;
FiniMove ;
Ferme(Q) ;
Ferme(QFiche[0]) ; Ferme(QFiche[1]) ; Ferme(QFiche[3]) ;
END ;

procedure TFCalcTva.BValiderClick(Sender: TObject);
Var i : Integer ;
begin
i:=HMess.Execute(0,'','') ;
If i<>mrYes Then Exit ;
Try
 BeginTrans ;
 LanceTraitement ;
 CommitTrans ;
Except
 RollBack ;
End ;
end;

procedure TFCalcTva.FormCreate(Sender: TObject);
begin
FillChar(InfoImp,SizeOf(InfoImp),#0) ;
InfoImp.LGenLu:=HTStringList.Create ;
InfoImp.LAuxLu:=HTStringList.Create ;
InfoImp.LJalLu:=HTStringList.Create ;
end;

procedure TFCalcTva.FormClose(Sender: TObject; var Action: TCloseAction);
begin
VideListeInfoImp(InfoImp,TRUE) ;
end;
{$ENDIF}

procedure CCalcTvaSimplifiee(  vCodeExo, Wheres : string; WhereJournal : string='' ) ;
Var Q,Q1 : TQuery ;
    St,StA,StV : String ;
    QFiche : TQFiche ;
    JalP,ExoP,QualP : String ;
    DateP : TDateTime ;
    NumP : Integer ;
    M : Rmvt ;
    InfoImp : TInfoImport ;
begin
FillChar(InfoImp,SizeOf(InfoImp),#0) ;
InfoImp.LGenLu:=HTStringList.Create ;
InfoImp.LAuxLu:=HTStringList.Create ;
InfoImp.LJalLu:=HTStringList.Create ;
StA:=WhereJal('ACH', WhereJournal) ; StV:=WhereJal('VTE', WhereJournal) ;
If (StA='') And (StV='') Then begin VideListeInfoImp(InfoImp,TRUE) ; Exit ; end;
St:='SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_QUALIFPIECE FROM ECRITURE ' ;
St:=St+' WHERE E_JOURNAL<>"'+w_w+'" AND E_QUALIFPIECE="N" ' ;
if vCodeExo <>'' then St:=St+' AND E_EXERCICE="'+vCodeExo+'" ' ;
if Wheres <> '' then St := St + 'AND ' + Wheres;
St:=St+' AND (' ;
If (StV<>'') And (StA<>'') Then St:=St+StA+' Or '+StV+') '
                           Else St:=St+StA+StV+') ' ;
St:=St+' GROUP BY E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_QUALIFPIECE ' ;
Q:=OpenSQL(St,TRUE) ;
//InitRequete(QFiche[0],0) ; InitRequete(QFiche[1],1) ; InitRequete(QFiche[3],3) ;
While Not Q.Eof Do
  BEGIN
  JalP:=Q.FindField('E_JOURNAL').AsString ;
  ExoP:=Q.FindField('E_EXERCICE').AsString ;
  DateP:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  NumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  QualP:=Q.FindField('E_QUALIFPIECE').AsString ;
  Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+JalP+'"'
             +' AND E_EXERCICE="'+ExoP+'"'
             +' AND E_DATECOMPTABLE="'+USDateTime(DateP)+'"'
             +' AND E_NUMEROPIECE='+IntToStr(NumP)
             +' AND E_QUALIFPIECE="'+QualP+'" ',True) ;
  if Not Q1.EOF then
    BEGIN
    M:=MvtToIdent(Q1,fbGene,False) ; Ferme(Q1) ;
   // TraitementSurEcr(TRUE,M,true,InfoImp,QFiche) ;
    ElementsTvaEncRevise (M, True, InfoImp, QFiche) ;
    END Else Ferme(Q1) ;
  Q.Next ;
  END ;
Ferme(Q) ;
Ferme(QFiche[0]) ; Ferme(QFiche[1]) ; Ferme(QFiche[3]) ;
VideListeInfoImp(InfoImp,TRUE) ;   
END ;


end.
