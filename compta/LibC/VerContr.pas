unit VerContr;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu,
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls,
  Mask,
  Hctrls,
  Buttons,
  ExtCtrls,
  hmsgbox,
  Ent1,
  HEnt1,
  SaisUtil,
  Saiscomm,
  HCompte,
  HStatus,
  UTOB,
  RecupUtil,
{$IFDEF VER150}
  Variants,
{$ENDIF}
 {$IFDEF EAGLCLIENT}
  uWA,
 {$ELSE}
  CalCulContr ,
 {$ENDIF}
  Menus, ADODB ;

type
  TFVerContr = class(TForm)
    HMTrad: THSystemMenu;
    HPB: TPanel;
    PBoutons: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    E_EXERCICE: THValComboBox;
    TE_EXERCICE: THLabel;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_NATUREPIECE: THValComboBox;
    TE_NATUREPIECE: THLabel;
    Label1: THLabel;
    ENATUREGENE: THLabel;
    E_JOURNAL: THValComboBox;
    E_JOURNAL_: THValComboBox;
    Q: TQuery;
    HM: THMsgBox;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    LPiece,ListeErr  : TList ;
    procedure LanceMajContreparties ;
    procedure FabriqueRequete ;
    procedure TraitePieces ;
    procedure ChargePiece ;
    function  GetGeneTreso(OL : TOBC) : String17 ;
    Procedure MajLigneAna(OL : TOBC ; Gene,Auxi : String17) ;
    Procedure MajLigne(OL : TOBC ; Gene,Auxi : String17) ;
    function  ChargeOBC(Lig : integer ; var O : TOBC) : boolean ;
    Procedure AjouteErrLigne(OL : TOBC) ;
    function  LigneDeMemeSens(i : integer ; OL : TOBC) : boolean ;
    function  TrouveLig(LigneEnCours : integer ; Chp : String) : Integer ;
    function  TraiteLaPiece : boolean ;
    function  TraiteSurMontant : boolean ;
    procedure ClearListe(T : TList) ;
    procedure FreeListe(T : TList) ;
    procedure PrepareShow;
    function InitTobParam : TOB ;

  public
  end;

var LDesc : HTStrings ;

Procedure VerifContreparties ;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  RapSuppr;


// GCO - 25/04/2002
Procedure VerifContreparties ;
var FVerContr : TFVerContr ;
BEGIN
FVerContr:=TFVerContr.Create(Application) ;
try
  if VH^.EnSerie then
  begin
    FVerContr.PrepareShow;
    FVerContr.BValider.Click;
  end
  else
    FVerContr.ShowModal ;
finally
  FVerContr.Free ;
END ;
SourisNormale ;
END ;

procedure TFVerContr.PrepareShow;
begin
  E_DATECOMPTABLE.Text:=StDate1900 ;
  E_DATECOMPTABLE_.Text:=StDate2099 ;
  E_EXERCICE.ItemIndex :=0 ;
  E_NATUREPIECE.ItemIndex :=0 ;
  E_JOURNAL.ItemIndex:=0 ;
  E_JOURNAL_.ItemIndex := E_JOURNAL_.Items.Count-1 ;
end;

procedure TFVerContr.FormShow(Sender: TObject);
begin
  PrepareShow;
end;
// Fin GCO

procedure TFVerContr.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFVerContr.E_EXERCICEChange(Sender: TObject);
begin
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

function TFVerContr.InitTobParam : TOB ;
var
LT, L1   : TOB;
begin
    LT := TOB.Create('$PARAM', nil, -1) ;
    LT.AddChampSupValeur('USERLOGIN' , V_PGI.UserLogin ) ;
    LT.AddChampSupValeur('INIFILE'   , HalSocIni ) ;
    LT.AddChampSupValeur('PASSWORD'  , V_PGI.Password ) ;
    LT.AddChampSupValeur('DOMAINNAME', '' ) ;
    LT.AddChampSupValeur('DATEENTREE', V_PGI.DateEntree ) ;
    LT.AddChampSupValeur('DOSSIER'   , V_PGI.NoDossier ) ;
    LT.AddChampSupValeur('APPLICATION', 'CCS5') ;

    LT.AddChampSupValeur('BaseCommune', EstBaseCommune);

    L1 := TOB.Create('Trans', LT, -1) ;
    L1.AddChampSupValeur('E_EXERCICE'           , E_EXERCICE.value) ;
    L1.AddChampSupValeur('E_DATECOMPTABLE'      , UsDate(E_DATECOMPTABLE)) ;
    L1.AddChampSupValeur('E_DATECOMPTABLE_'     , UsDate(E_DATECOMPTABLE_)) ;
    L1.AddChampSupValeur('E_JOURNAL_'           , E_JOURNAL_.value) ;
    L1.AddChampSupValeur('E_JOURNAL'            , E_JOURNAL.value) ;
    L1.AddChampSupValeur('E_NATUREPIECE'        , E_NATUREPIECE.value) ;

    Result := LT;
end;


procedure TFVerContr.BValiderClick(Sender: TObject);
var SAV,AvecSAV : boolean ;
TB              : TOB;
{$IFDEF EAGLCLIENT}
lTobResult      : TOB;
{$ENDIF}
begin
if not VH^.EnSerie then
  if HM.Execute(0,Caption,'')<>mrYes then Exit ;

//LanceMajContreparties ;
TB := InitTobParam;
{$IFDEF EAGLCLIENT}
      with cWA.create do
      begin
          lTobResult  := Request('CtrCpta.CALCTREPARTIE','', TB,'','');
          free ;
      end ;
      if lTobResult = nil then
      begin
        PgiBox('Erreur, Server CtrCpta introuvable', ' Calcul des contreparties') ;
      end
      else
      if lTobResult.GetValue ('ERROR') <> '' then
           PgiBox(lTobResult.GetValue ('ERROR'), 'Calcul des contreparties')
      else PgiBox ('Calcul terminé');
{$ELSE}
      LanceTraitementContreparties (TB);
{$ENDIF}

TB.Free;
SAV:=False ;
if ListeErr<>NIL Then
  BEGIN
  If ListeErr.Count>0 then
     BEGIN
     if (VH^.EnSerie) or (HM.Execute(2,Caption,'')<>mryes) then Exit ;
     AvecSAV:=V_PGI.SAV ; V_PGI.SAV:=True ;
     RapportdErreurMvt(ListeErr,3,SAV,False) ;
     V_PGI.SAV:=AvecSAV ;
     END ;
  END else if Not VH^.Enserie then HM.Execute(1,Caption,'') ;

FreeListe(ListeErr) ; ListeErr:=NIL ;
Close ;
end;


procedure TFVerContr.LanceMajContreparties ;
BEGIN
FabriqueRequete ;
Try
 BEGINTRANS ;
 TraitePieces ;
 COMMITTRANS ;
Except
 on e:exception do
  begin
   MessageAlerte('Erreur lors du traitement ' + #13#10 + e.message );
   RollBack ;
  end;
End ;
END ;

procedure TFVerContr.FabriqueRequete ;
BEGIN
Q.SQL.Clear ;
(* GP 20/01/98
Q.SQL.Add('SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NATUREPIECE,E_NUMEROPIECE,') ;
Q.SQL.Add('E_NUMLIGNE,E_GENERAL,E_AUXILIAIRE,E_CONTREPARTIEGEN,E_CONTREPARTIEAUX') ;
Q.SQL.Add(',E_DEBIT,E_CREDIT,G_NATUREGENE,J_NATUREJAL,J_CONTREPARTIE FROM ECRITURE,JOURNAL,GENERAUX') ;
Q.SQL.Add(' WHERE J_JOURNAL=E_JOURNAL') ;
Q.SQL.Add(' AND G_GENERAL=E_GENERAL') ;
Q.SQL.Add(' AND E_JOURNAL>="'+E_JOURNAL.Value+'"') ;
Q.SQL.Add(' AND E_JOURNAL<="'+E_JOURNAL_.Value+'"') ;
if (E_EXERCICE.Value<>'') then Q.SQL.Add(' AND E_EXERCICE="'+E_EXERCICE.Value+'"')
                          else Q.SQL.Add(' AND E_EXERCICE<>""') ;
Q.SQL.Add(' AND E_DATECOMPTABLE>="'+UsDate(E_DATECOMPTABLE)+'"') ;
Q.SQL.Add(' AND E_DATECOMPTABLE<="'+UsDate(E_DATECOMPTABLE_)+'"') ;
if (E_NATUREPIECE.Value<>'') then  Q.SQL.Add(' AND E_NATUREPIECE="'+E_NATUREPIECE.Value+'"') ;
Q.SQL.Add(' AND (J_NATUREJAL="BQE" OR J_NATUREJAL="CAI" OR J_NATUREJAL="ACH" OR J_NATUREJAL="VTE"'+
          ' OR J_NATUREJAL="OD")') ;
Q.SQL.Add(' ORDER BY E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE') ;
*)
Q.SQL.Add('SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NATUREPIECE,E_NUMEROPIECE,') ;
Q.SQL.Add('E_NUMLIGNE,E_GENERAL,E_AUXILIAIRE,E_CONTREPARTIEGEN,E_CONTREPARTIEAUX') ;
Q.SQL.Add(',E_DEBIT,E_CREDIT,G_NATUREGENE,J_NATUREJAL,J_CONTREPARTIE,E_QUALIFPIECE,E_ANA FROM ECRITURE') ;
Q.SQL.Add(' LEFT JOIN JOURNAL ON E_JOURNAL=J_JOURNAL') ;
Q.SQL.Add(' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL') ;
Q.SQL.Add(' WHERE J_JOURNAL>="'+E_JOURNAL.Value+'"') ;
Q.SQL.Add(' AND J_JOURNAL<="'+E_JOURNAL_.Value+'"') ;
Q.SQL.Add(' AND (J_NATUREJAL="BQE" OR J_NATUREJAL="CAI" OR J_NATUREJAL="ACH" OR J_NATUREJAL="VTE"'+
          ' OR J_NATUREJAL="OD")') ;
if (E_EXERCICE.Value<>'') then Q.SQL.Add(' AND E_EXERCICE="'+E_EXERCICE.Value+'"')
                          else Q.SQL.Add(' AND E_EXERCICE<>""') ;
Q.SQL.Add(' AND E_DATECOMPTABLE>="'+UsDate(E_DATECOMPTABLE)+'"') ;
Q.SQL.Add(' AND E_DATECOMPTABLE<="'+UsDate(E_DATECOMPTABLE_)+'"') ;
Q.SQL.Add(' AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="S" OR E_QUALIFPIECE="R" OR E_QUALIFPIECE="U") ') ;
Q.SQL.Add(' AND (E_ECRANOUVEAU="N") ') ;
if (E_NATUREPIECE.Value<>'') then  Q.SQL.Add(' AND E_NATUREPIECE="'+E_NATUREPIECE.Value+'"') ;
Q.SQL.Add(' ORDER BY E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE') ;
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
END ;

procedure TFVerContr.TraitePieces ;
var OL,OAux,OHT,OTreso : TOBC ;
    i,Lig,NumL : Integer ;
    OkRupt,MemeSens,OkBQE : boolean ;
    LTreso,LHT,LAux : Integer ;
    GeneTreso : String17 ;
BEGIN
Q.Open ; if Q.Eof then BEGIN Q.Close ; Exit ; END ;
LDesc:=HTStringList.Create ;
Q.GetFieldNames(LDesc) ;
ListeErr:=TList.Create ;
LPiece:=TList.Create ;
InitMove(RecordsCount(Q),'') ;
While not Q.Eof do
  BEGIN
  ChargePiece ;
//  If Q.Eof Then Break ; // 13356
  if not TraiteLaPiece then Continue ;
  //Traitement de la pièce
  i:=0 ; OkRupt:=True ; LAux:=-1 ; LTreso:=-1 ; OkBQE:=False ; LHT:=-1 ;
  OTreso:=nil ; OAux:=nil ; OHT:=nil ;
  While i<=LPiece.Count-1 do
    BEGIN
    OL:=LPiece[i] ;
    NumL:=OL.GetMvt('E_NUMLIGNE') ;
    if OkRupt then
      BEGIN
      OTreso:=nil ; OAux:=nil ; OHT:=nil ;
      OkBQE:=(OL.GetMvt('J_NATUREJAL')='BQE') or (OL.GetMvt('J_NATUREJAL')='CAI') ;
      LAux:=TrouveLig(i,'E_AUXILIAIRE') ;
      LHT:=TrouveLig(i,'E_GENERAL') ; if (LAux<NumL) then LAux:=NumL+1 ;
      if (LHT<0) then BEGIN if LAux<>NumL+1 then LHT:=NumL+1 else LHT:=NumL ; END ;
      if (LHT>LPiece.Count) or (LAux>LPiece.Count) then break ;
      if not ChargeOBC(LAux,OAux) then BEGIN AjouteErrLigne(OL) ; Break ; END ;
      if not ChargeOBC(LHT,OHT) then BEGIN AjouteErrLigne(OL) ; Break ; END ;
      LTreso:=-1 ;
      if OkBQE then
        BEGIN
        GeneTreso:=GetGeneTreso(OL) ;
        for Lig:=i to LPiece.Count-1 do
          BEGIN
          OL:=LPiece[Lig] ;
          if (OL.GetMvt('E_GENERAL')<>GeneTreso) then OTreso:=OL else LTreso:=Lig+1 ;
          END ;
        OL:=LPiece[i] ;
        END ;
      OkRupt:=False ;
      END ;

    MemeSens:=LigneDeMemeSens(i,OL) ;

    // Changement d'Auxiliaire dans la même pièce
    if (OL.GetMvt('E_AUXILIAIRE')<>'') and (NumL>LAux) then
       if not OkBQE then
         BEGIN
         if (NumL>LHT) and not MemeSens then BEGIN OkRupt:=True ; Continue ; END ;
         END else
         BEGIN
         if (NumL>LTreso) and not MemeSens then BEGIN OkRupt:=True ; Continue ; END ;
         END ;
    if OkBQE then
      BEGIN
      if (OL.GetMvt('E_GENERAL')<>GeneTreso) then MajLigne(OL,GeneTreso,'')
                                             else if (NumL=LTreso) then MajLigne(OL,OTreso.GetMvt('E_GENERAL'),OTreso.GetMvt('E_AUXILIAIRE')) ;
      END else
      BEGIN
      if (NumL=LAux) then MajLigne(OL,OHT.GetMvt('E_GENERAL'),OHT.GetMvt('E_AUXILIAIRE'))
                     else MajLigne(OL,OAux.GetMvt('E_GENERAL'),OAux.GetMvt('E_AUXILIAIRE')) ;
      END ;
    Inc(i) ;
    END ;
  END ;
FiniMove ;
Q.Close ;
LDesc.Clear ; LDesc.Free ;
FreeListe(LPiece) ;
END ;

function TFverContr.TraiteLaPiece : boolean ;
var OldL,L,i,j : integer ;
    GeneTreso : String17 ;
    OkBQE : boolean ;
BEGIN
Result:=False ;
if TraiteSurMontant then Exit ;
GeneTreso:='' ;
OkBQE:=(TOBC(LPiece[0]).GetMvt('J_NATUREJAL')='BQE') or (TOBC(LPiece[0]).GetMvt('J_NATUREJAL')='CAI') ;
if OkBQE then GeneTreso:=GetGeneTreso(TOBC(LPiece[0])) ;
OldL:=TrouveLig(0,'E_AUXILIAIRE') ;
if OldL=-1 then Exit ;
For i:=1 to LPiece.Count-1 do
  BEGIN
  // au moins 2 auxiliaires consécutifs
  L:=TrouveLig(i,' E_AUXILIAIRE') ;
  if (L-OldL=1) or ((L>-1) and (TOBC(LPiece[OldL]).GetMvt('E_AUXILIAIRE')<>TOBC(LPiece[L]).GetMvt('E_AUXILIAIRE'))) then
    BEGIN
    if OkBQE then
      BEGIN
      // 2 comptes de tréso
      L:=0 ;
      for j:=0 to LPiece.Count-1 do if TOBC(LPiece[j]).GetMvt('E_GENERAL')=GeneTreso then Inc(L) ;
      if L>1 then BEGIN AjouteErrLigne(TOBC(LPiece[0])) ; Exit ; END ;
      END else
      BEGIN
      // 2 comptes de tréso
      OldL:=TrouveLig(0,'E_GENERAL') ;
      for j:=1 to LPiece.Count-1 do
        BEGIN
        L:=TrouveLig(j,'E_GENERAL') ;
        if L-OldL=1 then BEGIN AjouteErrLigne(TOBC(LPiece[0])) ; Exit ; END ;
        END ;
      END ;
    END ;
  END ;
Result:=True ;
END ;

function TFVerContr.TraiteSurMontant : boolean ;
var j,i,LAux,LTreso  : integer ;
    OL : TOBC ;
    M : Double ;
    LT : TList ;
    UnTour : boolean ;
BEGIN
UnTour:=False ; Result:=FALSE ;
If LPiece=Nil Then Exit ; LT:=TList.Create ;
for i:=0 to LPiece.Count-1 do LT.Add(TOBC(LPiece[i])) ;
While (LT.Count>0) do
  BEGIN
  LAux:=-1 ; LTreso:=-1 ; UnTour:=False ;
  for i:=0 to LT.Count-1 do
    BEGIN
    OL:=LT[i] ;
    M:=Arrondi(Double(OL.GetMvt('E_DEBIT'))-Double(OL.GetMvt('E_CREDIT')),V_PGI.OkDecV) ;
    If M<>0 Then
       BEGIN
       for j:=i to LT.Count-1 do
         BEGIN
         OL:=LT[j] ;
         if M=-(Arrondi(Double(OL.GetMvt('E_DEBIT'))-Double(OL.GetMvt('E_CREDIT')),V_PGI.OkDecV)) then
           BEGIN
           LTreso:=i ;
           LAux:=j ;
           break ;
           END ;
         END ;
       END ;
    if (LAux>-1) then Break ;
    if (i=LT.Count-1) then UnTour:=True ;
    END ;
  if (LAux>-1) and (LTreso>-1) then
    BEGIN
    MajLigne(TOBC(LT[LAux]),TOBC(LT[LTreso]).GetMvt('E_GENERAL'),TOBC(LT[LTreso]).GetMvt('E_AUXILIAIRE')) ;
    MajLigne(TOBC(LT[LTreso]),TOBC(LT[LAux]).GetMvt('E_GENERAL'),TOBC(LT[LAux]).GetMvt('E_AUXILIAIRE')) ;
    if (LAux<LTreso) then Dec(LTreso) ;
    LT.Remove(TOBC(LT[LAux])) ;
    LT.Remove(TOBC(LT[LTreso])) ;
    END else if UnTour then Break ;
  END ;
Result:=(LT.Count=0) ;
if (LT.Count>0) and (LT.Count<>LPiece.Count) and UnTour then AjouteErrLigne(TOBC(LPiece[0])) ;
LT.Clear ; LT.Free ;
END ;


procedure TFVerContr.ChargePiece ;
var NumL : Integer ;
    O    : TOBC ;
BEGIN
ClearListe(LPiece) ;
NumL:=0 ;
While not Q.Eof and (Q.FindField('E_NUMLIGNE').AsInteger>=NumL) do
  BEGIN
  NumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
  O:=TOBC.Create(Q) ;
  LPiece.Add(O) ;
  Q.Next ;
  MoveCur(False) ;
  END ;
END ;

function TFVerContr.LigneDeMemeSens(i : integer ; OL : TOBC) : boolean ;
var Debiteur : boolean ;
    OP : TOBC ;
BEGIN
if i=0 then BEGIN Result:=True ;  Exit ; END ;
OP:=LPiece[i-1] ;
Debiteur:=((Double(OP.GetMvt('E_DEBIT'))-Double(OP.GetMvt('E_CREDIT')))>0) ;
Result:=(((Double(OL.GetMvt('E_DEBIT'))-Double(OL.GetMvt('E_CREDIT')))>0)=Debiteur) ;
END ;

function TFVerContr.GetGeneTreso(OL : TOBC) : String17 ;
var St : String ;
BEGIN
St:=OL.GetMvt('J_CONTREPARTIE') ;
Result:=TrouveAuto(St,1) ;
if (Result='') then
  if ((OL.GetMvt('G_NATUREGENE')='BQE') or (OL.GetMvt('G_NATUREGENE')='CAI')) then Result:=OL.GetMvt('E_GENERAL') ;
END ;

function TFVerContr.TrouveLig(LigneEnCours : integer ; Chp : String) : Integer ;
var i,NumL : Integer ;
    Nat : String3 ;
    OkTva : boolean ;
    OL: TOBC ;
BEGIN
Result:=-1 ; OkTVA:=False ;
For i:=LigneEnCours to LPiece.Count-1 do
  BEGIN
  OL:=LPiece[i] ;
  NumL:=OL.GetMvt('E_NUMLIGNE') ;
  Nat:=OL.GetMvt('G_NATUREGENE') ;
  if (Chp='E_AUXILIAIRE') then
    if (OL.GetMvt('E_AUXILIAIRE')<>'') or ((Nat='TID') or (Nat='TIC')) then Result:=NumL ;
  if (Chp='E_GENERAL') then
    if ((Nat='CHA') or (Nat='PRO')) or ((Nat='IMO') and (OkTVA)) then Result:=NumL ;
  if (Result>-1) then Break ;
  END ;
END ;

function TFVerContr.ChargeOBC(Lig : integer ; var O : TOBC) : boolean ;
BEGIN
Result:=False ;
if (Lig=0) or (Lig>LPiece.Count) then Exit ;
O:=LPiece[Lig-1] ;
Result:=True ;
END ;

Procedure TFVerContr.MajLigneAna(OL : TOBC ; Gene,Auxi : String17) ;
var St : String ;
BEGIN
If OL.GetMvt('E_ANA')<>'X' Then Exit ;
St:='UPDATE ANALYTIQ SET Y_CONTREPARTIEGEN="'+Gene+'",'+'Y_CONTREPARTIEAUX="'+Auxi+'" WHERE'
   +' Y_JOURNAL="'+OL.GetMvt('E_JOURNAL')+'" AND Y_DATECOMPTABLE="'
   +USDateTime(OL.GetMvt('E_DATECOMPTABLE'))+'" AND Y_NATUREPIECE="'
   +OL.GetMvt('E_NATUREPIECE')+'" AND Y_NUMEROPIECE='
   +IntToStr(OL.GetMvt('E_NUMEROPIECE'))+' AND Y_NUMLIGNE='+IntToStr(OL.GetMvt('E_NUMLIGNE'))+' AND '
   +' Y_QUALIFPIECE="'+OL.GetMvt('E_QUALIFPIECE')+'" AND Y_EXERCICE="'+OL.GetMvt('E_EXERCICE')+'" ' ;
ExecuteSQL(St) ;
END ;

Procedure TFVerContr.MajLigne(OL : TOBC ; Gene,Auxi : String17) ;
var St : String ;
BEGIN
if (OL.GetMvt('E_CONTREPARTIEGEN')=Gene) and (OL.GetMvt('E_CONTREPARTIEAUX')=Auxi) then Else
  BEGIN
  St:='UPDATE ECRITURE SET E_CONTREPARTIEGEN="'+Gene+'",'+'E_CONTREPARTIEAUX="'+Auxi+'" WHERE'
     +' E_JOURNAL="'+OL.GetMvt('E_JOURNAL')+'" AND E_DATECOMPTABLE="'
     +USDateTime(OL.GetMvt('E_DATECOMPTABLE'))+'" AND E_NATUREPIECE="'
     +OL.GetMvt('E_NATUREPIECE')+'" AND E_NUMEROPIECE='
     +IntToStr(OL.GetMvt('E_NUMEROPIECE'))+' AND E_NUMLIGNE='+IntToStr(OL.GetMvt('E_NUMLIGNE'))+' AND '
     +' E_QUALIFPIECE="'+OL.GetMvt('E_QUALIFPIECE')+'" AND E_EXERCICE="'+OL.GetMvt('E_EXERCICE')+'" ' ;
  ExecuteSQL(St) ;
  END ;
MajLigneAna(OL,Gene,Auxi) ;
END ;

Procedure TFVerContr.AjouteErrLigne(OL : TOBC) ;
Var X : DelInfo ;
BEGIN
//Inc(NbError) ;
X:=DelInfo.Create ;
X.LeCod:=OL.GetMvt('E_JOURNAL') ;
//X.LeLib:=Entitee.Refinterne ;
X.LeMess:=IntToStr(OL.GetMvt('E_NUMEROPIECE'))+'/'+IntToStr(OL.GetMvt('E_NUMLIGNE')) ;
X.LeMess2:=DateToStr(OL.GetMvt('E_DATECOMPTABLE')) ; X.LeMess3:=OL.GetMvt('E_GENERAL')+' / '+OL.GetMvt('E_AUXILIAIRE') ;
X.LeMess4:=OL.GetMvt('E_EXERCICE') ;
ListeErr.Add(X) ;
END ;


procedure TFVerContr.ClearListe(T : TList) ;
var i : Integer ;
BEGIN
for i:=0 to T.Count-1 do TObject(T[i]).Free ;
T.Clear ;
END;

procedure TFVerContr.FreeListe(T : TList) ;
BEGIN
if T=nil then Exit ;
ClearListe(T) ; T.Free ; 
END;

end.
