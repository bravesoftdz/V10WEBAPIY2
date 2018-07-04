unit ReparMvt;
//Document Intégrateur n° 60017_86.doc (G:\client\0\7)
// Info à garder !! ...
//      Control sur les pièces en Devise vis-à-vis de la Quotité !
//
// Test Sur TotDebit, Totcredit des Cptes Genes ??
//
// Test sur Equilibre bloquant ?
// Test sur TAUXDEV en ECC suppr ?
// Test sur TAUXDEV en ECC suppr ?


interface                         

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mask, StdCtrls, Hctrls, Hcompte, Buttons, ExtCtrls, Ent1, HEnt1, UTILEDT,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$IFDEF VER150}Variants,{$ENDIF}
  DB, HQry, CpteUtil, hmsgbox, SaisUtil, HStatus, RapSuppr,RappType,
  ComCtrls, CRITEDT, HSysMenu,ParamDat, HDB, HRegCpte, UtilSais,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  tCalcCum,
  {$ENDIF MODENT1}
  uLibEcriture,
  ParamSoc,
  uTob; // TOB


procedure RepareMvt;

type  TEnregInfos = Class
      Inf1      : String ;
      Inf2      : String ;
      Inf3      : String ;
      Inf4      : String ;
      Inf5      : String ;
      D1        : TDateTime ;
      D2        : TDateTime ;
      Nb        : Byte ;
      Bol1, Bol2,Bol3 : Boolean ;
      SuperBol  : T5B ;
      Mont      : Double ;
      END ;

type  TINFOSCPTE = Record
      General    : String ;
      NatureGene : String ;
      Collectif, Lettrable, Pointable    : Boolean ;
      Ventilable : T5B ;
      END ;

type  TINFOSAUXI = Record
      Auxi      : String ;
      NatureAux : String ;
      Lettrable : Boolean ;
      END ;

type  TErrLetDev = Class
      Lettres   : String ;
      Auxi      : String ;
      Gene      : String ;
      EtatLet   : String ;
      END ;

type  TINFOSMVTANA = Record
      General   : String ;
      AnalPur   : Boolean ;
      TotEcr, TotDev : Double ;
      TypeANO, EcrANO : String ;
      END ;

type  TINFOSAUTRES = Record
      J_Axe : String ;
      J_COMPTEINTERDIT : String ;
      J_NATUREJAL      : String ;
      J_CONTREPARTIE   : String ;
      EX_DATEEXO       : Array[0..1] Of TDateTime ;
      END ;

Type TERRORIMP = Class
     NumChrono : Integer ;
     NumErreur : String ;
     END ;
type
  TFReparMvt = class(TForm)
    HPB: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    Panel1: TPanel;
    TFJal: THLabel;
    TFaJ: TLabel;
    TFExercice: THLabel;
    FExercice: THValComboBox;
    FDateCpta1: TMaskEdit;
    TFDateCpta1: THLabel;
    TFDateCpta2: TLabel;
    FDateCpta2: TMaskEdit;
    TFEtab: THLabel;
    FEtab: THValComboBox;
    QSum: TQuery;
    QEcr: TQuery;
    QAnal: TQuery;
    MsgRien: THMsgBox;
    MsgLibel: THMsgBox;
    MsgLibel2: THMsgBox;
    TFTypeEcriture: THLabel;
    FTypeEcriture: THValComboBox;
    Shape1: TShape;
    MsgBar: THMsgBox;
    TTravail: TLabel;
    QInfoAnaG: TQuery;
    Panel2: TPanel;
    BStop: THBitBtn;
    TCEtab: THValComboBox;
    QAnaGene: TQuery;
    QSumGeAn: TQuery;
    TCNatPiece: TComboBox;
    HMTrad: THSystemMenu;
    TFNumPiece1: THLabel;
    FNumPiece1: TMaskEdit;
    TFNumPiece2: TLabel;
    FNumPiece2: TMaskEdit;
    TCModP: THValComboBox;
    TErrD: TLabel;
    TErrC: TLabel;
    NbErrD: THNumEdit;
    NbErrC: THNumEdit;
    GeneAttend: THDBCpteEdit;
    HLabel1: THLabel;
    CtrlEcrGen: TCheckBox;
    CtrlEcrAna: TCheckBox;
    CtrlDetGen: TCheckBox;
    CtrlDetAna: TCheckBox;
    CtrlTreso: TCheckBox;
    BVentil: THBitBtn;
    CtrlAnaOff: TCheckBox;
    CtrlSouche: TCheckBox;
    MPAttend: THValComboBox;
    H_MODEPAIE: THLabel;
    btnError: THBitBtn;
    FJal1: THCritMaskEdit;
    FJal2: THCritMaskEdit;
    cbTypeJal: THValComboBox;
    HLabel2: THLabel;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BStopClick(Sender: TObject);
    procedure FDateCpta1KeyPress(Sender: TObject; var Key: Char);
    procedure BVentilClick(Sender: TObject);
    procedure btnErrorClick(Sender: TObject);
    procedure cbTypeJalChange(Sender: TObject);
  private
    { Déclarations privées }
    OkVerif, StopVerif : Boolean ;
    LaListe            : TList ;
    Mvt                : TInfoMvt ;
    FInfoEcr           : TInfoEcriture;
    FMessage           : TMessageCompta;
    InfAutres          : TINFOSAUTRES ;
    NbEnreg, NbError   : Integer ;
    EnBatch            : Boolean ;
    CGen               : TINFOSCPTE ;
    CTiers             : TINFOSAUXI ;
    VCrit              : TCritEdt ;
    ListeDEV           : TStringList ;
    ListeJAL           : TStringList ;
    ListeEXO           : TStringList ;
    ListeSec           : TStringList ;
    ListeGen           : TStringList ;
    ListeBQE           : TStringList ;
    ListeAUX           : TStringList ;
    Function GoListe1(Entitee : TInfoMvt ; Quel, Rem : String ; I : Byte ; CodeIMP : String ) : DelInfo ;
    Function GoListe2(Entitee : TInfoMvt ; Quel, Rem : String ; F : TField ; CodeIMP : String ) : DelInfo ;
    Function OkPourReparation : Integer ;
    Function VerifCptAttente(fb : TFichierBase) : Integer ;
    Function SiEnreg : Boolean ;
    procedure SqlEquilEcr(Ec : Byte) ;
    procedure SQLEcr ;
    Function EquilEcr    : Boolean ;
    Function EquilAnal   : Boolean ;
    Procedure ErrEquilEcr(Var Err : Boolean ; E : Byte ; D,C,DD,CD,DE,CE : Double ; DecE : Integer) ;
    Procedure CorrEquilAna(E : Byte ; M : TabTot ; OnDeb : Boolean ; DecE : Integer ; SensInverse : Boolean) ;
    Procedure VerifAutres(Var IFA : TINFOSMVTANA ) ;
//    procedure EnregMvtAna ;
//    Function VerifInfoExoAnal : Boolean ;
//    Function VerifCptAnal : Boolean ;
//    Function VerifInfoJalAnal : Boolean ;
//    Function VerifInfoPieceAnal : Boolean ;
//    Function VerifInfoDeviseAnal : Boolean ;
    Function Majuscule(St : String) : Boolean ;
    Procedure EnregMvtEcr ;
    Function VerifCompte(Var let, Poin, vent, DeBanque : Boolean) : Boolean ;
    Function VerifInfoAnal(Ventilable : Boolean) : Boolean ;
    Function VerifInfoEche(Lettrable, pointable : Boolean) : Boolean ;
//    Function VerifEtabAnal : Boolean ;
//    Function EcrAnoOk(Journal,Ecran : String) : Boolean ;
//    Function OkExo(E : String3 ; D : TDateTime) : Boolean ;
    procedure Reparation ;
    procedure RecupCrit;
    function  TestBreak : Boolean ;
    function  Existe(V : Byte ; St : String ) : Boolean ;
//    function  ExisteSec(Ax,Cod : String ) : Boolean ;
    Procedure EnregEquilA ;
    Procedure EnregEquil(Q : TQuery ; Quelle : byte) ;
    procedure Corrige(Q : TQuery ; Champ, Valeur : String) ;
    procedure CorrigeInt(Q : TQuery ; Champ : String ; Valeur : Integer) ;
    procedure ChargeInfos ;
    procedure VideInfos ;
    procedure TrouveDecDev(St : String3 ; Var Dec : Byte) ;
    procedure InfosCptGen(Ge : String) ;
    procedure InfosCptAux(Au : String) ;
    Procedure PrepareAnal ;
    Procedure PrepareAutres ;
//    function  VerifTauxDev(MvtEcr : Boolean) : Boolean ;
//    function  DonneQuotite(Dev : String) : Double ;
    Procedure lalisteAdd(G : DelInfo) ;

    procedure PrepareShow;

  public
    { Déclarations publiques }
    {JP 02/06/05 : j'ai sorti le traitement proprement dit de BValiderClick}
    procedure LanceTraitement;
  end;

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  CPTESAV, CPAXE_TOM ; // Axe;

Const MaxNbErreur=500 ;

{$R *.DFM}
(*
Function LeSolde(D,C : Double) : Double ;
BEGIN
Result:=0 ;
if D=C then result:=abs(D)-abs(C) else
if Abs(D)<Abs(C) then result:=abs(D)-abs(C) else
if Abs(D)>Abs(C)  then  result:=(abs(C)-abs(D))*-1 ;
END ;
*)
Function LeSolde(D,C : Double ; Ana,OnDeb : Boolean ; Dec : Integer ; Var D1,C1 : Double) : Boolean ;
Var S : Double ;
    Signe : Integer ;
BEGIN
Result:=FALSE ; D1:=0 ; C1:=0 ;
S:=Abs(Arrondi(D-C,Dec)) ; If S=0 Then Exit ;
Result:=TRUE ;
If D>C Then C1:=S Else D1:=S ;
If Ana Then
   BEGIN
   Signe:=1 ; D1:=0 ; C1:=0 ; If Abs(D)<Abs(C) Then Signe:=-1 ;
   If OnDeb Then D1:=S*Signe Else C1:=S*Signe ;
   END ;
END ;


procedure RepareMvt;
var FReparMvt : TFReparMvt ;
BEGIN
FReparMvt:=TFReparMvt.Create(Application) ;
try
 FReparMvt.EnBatch:=False ;
 if VH^.EnSerie then
 begin
   FReparMvt.PrepareShow;
   FReparMvt.BValider.Click;
 end
 else
   FReparMvt.ShowModal ;

finally
FReparMvt.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
END ;

// GCO - 14/04/2002
procedure TFReparMvt.PrepareShow;
begin
  PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
  FJal1.Text:='' ; FJal2.Text:='' ;
  FEtab.ItemIndex:=0 ;
  FExercice.Value := EXRF(VH^.Entree.Code) ;
  FTypeEcriture.ItemIndex:=0 ;
  TTravail.Caption:='' ;
  GeneAttend.Text:=VH^.Cpta[fbGene].Attente ;
  If OkToutSeul Then CtrlSouche.Enabled:=TRUE ;
  MPAttend.ItemIndex:=0 ;
  {$IFDEF REPARANA}
  CtrlAnaOff.Checked:=TRUE ;
  {$ENDIF}

  if VH^.EnSerie then
  begin
    CtrlTreso.Checked := True;
    CtrlAnaOff.Checked := True;
    CtrlDetGen.Checked := True;
    CtrlDetAna.Checked := True;
  end;
end;

procedure TFReparMvt.FormShow(Sender: TObject);
begin
  PrepareShow;
end;
// FIN GCO

procedure TFReparMvt.FormCreate(Sender: TObject);
begin
LaListe:=TList.Create;
FInfoEcr := TInfoEcriture.Create;
FMessage := TMessageCompta.Create( Self.Caption );
end;

procedure TFReparMvt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
LaListe.Free;
FInfoEcr.Free;
FMessage.Free;
end;

procedure TFReparMvt.BValiderClick(Sender: TObject);
var
  OkRepar : Integer;
begin
  OkVerif:=True ; NbError:=0 ; NbErrD.Value:=0 ; NbErrC.Value:=0 ;
  StopVerif:=False ;
  {Lancement des contrôles avant le lancement du traitement}
  OkRepar := OkPourReparation;

  {Si tout est Ok, on commence la réparation ...}
  if OkRepar = 0 then begin
    LaListe.Clear ;
    EnableControls(Self, False) ;
    TTravail.Visible := True ;
    TTravail.Caption := MsgBar.Mess[16] ;
    BStop.Visible    := True ;
    Application.ProcessMessages ;

    // FQ 10394
    if CtrlTreso.Checked then
      NbError := NbError + DetectEcritureNonValide(VCrit, LaListe, NbErrD, NbErrC);

    if CtrlAnaOff.Checked then begin
      NbError := NbError + ChercheAnaSansGene(VCrit, LaListe, True);
      If NbError > 0 then
        OkVerif := FALSE;
    end;

    {JP 02/06/05 : FQ 13353 : Le traitement proprement dit est maintenant}
    LanceTraitement();

    EnableControls(Self,True) ;
    BStop.Visible:=FALSE ;
  end
  {Sinon affichage du message justificatif}
  else if (not VH^.EnSerie) then
    MsgRien.Execute(OkRepar, Caption, '') ;

If CtrlSouche.Checked Then
  BEGIN
  EnableControls(Self,False) ; TTravail.Visible:=True ; TTravail.Caption:=MsgBar.Mess[16] ; Application.ProcessMessages ;
  RecalculSouche ;
  EnableControls(Self,True) ;
  END ;
// FQ 10394
TTravail.Visible:=FALSE ; TTravail.Caption:='' ; Application.ProcessMessages ;
end;

procedure TFReparMvt.LanceTraitement();
var
  HH : Boolean;
begin
   if SiEnreg then
   BEGIN
      ChargeInfos ;
      Reparation ;
      TTravail.Visible:=FALSE ; TTravail.Caption:='' ;
      FiniMove ;
      if Not OkVerif then
      BEGIN
         HH:=TRUE ; If (LaListe<>NIL) And (LaListe.Count>0) Then RapportdErreurMvt(Laliste,3,HH,False) ;
      END
      Else
      if (not StopVerif) and (not VH^.EnSerie) then MsgRien.Execute(3,Caption,'') ; // sinon message tout est ok
      VideInfos ;
   END
   Else BEGIN
      If CtrlDetGen.Checked Or CtrlDetAna.Checked Or CtrlEcrGen.Checked Or CtrlEcrAna.Checked Then MsgRien.Execute(0,Caption,'') ; // message : aucun enreg
      If CtrlAnaOff.Checked And Not OkVerif Then
      BEGIN
        TTravail.Visible:=FALSE ; TTravail.Caption:='' ;
        FiniMove ;
        if Not OkVerif then
        BEGIN
           NbErrD.Value:=NbError ; NbErrC.Value:=NbError ;
           HH:=TRUE ; If (LaListe<>NIL) And (LaListe.Count>0) Then RapportdErreurMvt(Laliste,3,HH,False) ;
        END
        Else
        if not StopVerif then MsgRien.Execute(3,Caption,'') ; // sinon message tout est ok
      END ;
   END ;

end;

procedure TFReparMvt.ChargeInfos ;
Var Q : TQuery ; infos : TEnregInfos ; NbAux : Integer ;
BEGIN
ListeDev:=TStringList.Create ;
Q:=OpenSql(' Select D_DEVISE, D_DECIMALE, D_QUOTITE from DEVISE ',True) ;
While Not Q.Eof do
      BEGIN
      Infos:=TEnregInfos.Create ;
      Infos.Inf1:=Q.Fields[0].AsString ;
      Infos.Nb:=Q.Fields[1].AsInteger ;
      Infos.Mont:=Q.Fields[2].AsFloat ;
      ListeDev.AddObject('',Infos) ; Q.Next ;
      END ;
Ferme(Q) ;
ListeJal:=TStringList.Create ;
ListeExo:=TStringList.Create ;
ListeSec:=TStringList.Create ;
ListeGen:=TStringList.Create ;
ListeBQE:=TStringList.Create ;
ListeAUX:=TStringList.Create ;
Q:=OpenSql(' Select J_JOURNAL, J_NATUREJAL, J_AXE, J_COMPTEINTERDIT,J_CONTREPARTIE from JOURNAL order by J_JOURNAL ',True) ;
While Not Q.Eof do
      BEGIN
      Infos:=TEnregInfos.Create ;
      Infos.Inf1:=Q.FindField('J_JOURNAL').AsString ;
      Infos.Inf2:=Q.FindField('J_NATUREJAL').AsString ;
      Infos.Inf3:=Q.FindField('J_AXE').AsString ;
      Infos.Inf4:=Q.FindField('J_COMPTEINTERDIT').AsString ;
      Infos.Inf5:=Q.FindField('J_CONTREPARTIE').AsString ;
      ListeJal.AddObject('',Infos) ; Q.Next ;
      END ;
Ferme(Q) ;
Q:=OpenSql(' select EX_EXERCICE, EX_DATEDEBUT, EX_DATEFIN from EXERCICE ',True) ;
While Not Q.Eof do
      BEGIN
      Infos:=TEnregInfos.Create ;
      Infos.Inf1:=Q.FindField('EX_EXERCICE').AsString ;
      Infos.D1:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
      Infos.D2:=Q.FindField('EX_DATEFIN').AsDateTime ;
      ListeExo.AddObject('',Infos) ; Q.Next ;
      END ;
Ferme(Q) ;

Q:=OpenSql(' select S_AXE, S_SECTION from SECTION order by S_AXE, S_SECTION ',True) ;
While Not Q.Eof do
      BEGIN
      Infos:=TEnregInfos.Create ;
      Infos.Inf1:=Q.FindField('S_AXE').AsString ;
      Infos.Inf2:=Q.FindField('S_SECTION').AsString ;
      ListeSec.AddObject('',Infos) ; Q.Next ;
      END ;
Ferme(Q) ;

Q:=OpenSql('select G_GENERAL, G_NATUREGENE, G_COLLECTIF, G_LETTRABLE, G_POINTABLE, '
          +'G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 '
          +'from GENERAUX order by G_GENERAL',True) ;
While Not Q.Eof do
      BEGIN
      Infos:=TEnregInfos.Create ;
      Infos.Inf1:=Q.FindField('G_GENERAL').AsString ;
      Infos.Inf2:=Q.FindField('G_NATUREGENE').AsString ;
      Infos.Bol1:=(Q.FindField('G_COLLECTIF').AsString='X') ;
      Infos.Bol2:=(Q.FindField('G_LETTRABLE').AsString='X') ;
      Infos.Bol3:=(Q.FindField('G_POINTABLE').AsString='X') ;
      Infos.SuperBol[1]:=(Q.FindField('G_VENTILABLE1').AsString='X') ;
      Infos.SuperBol[2]:=(Q.FindField('G_VENTILABLE2').AsString='X') ;
      Infos.SuperBol[3]:=(Q.FindField('G_VENTILABLE3').AsString='X') ;
      Infos.SuperBol[4]:=(Q.FindField('G_VENTILABLE4').AsString='X') ;
      Infos.SuperBol[5]:=(Q.FindField('G_VENTILABLE5').AsString='X') ;
      ListeGen.AddObject('',Infos) ; Q.Next ;
      END ;
Ferme(Q) ;

Q:=OpenSql('select BQ_GENERAL, BQ_DEVISE from BANQUECP WHERE BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
While Not Q.Eof do
      BEGIN
      Infos:=TEnregInfos.Create ;
      Infos.Inf1:=Q.FindField('BQ_GENERAL').AsString ;
      Infos.Inf2:=Q.FindField('BQ_DEVISE').AsString ;
      ListeBQE.AddObject('',Infos) ; Q.Next ;
      END ;
Ferme(Q) ;

NbAux:=0 ;
Q:=OpenSql('select T_AUXILIAIRE, T_NATUREAUXI, T_LETTRABLE from TIERS',True) ;
While (Not Q.Eof)and(NbAux<=2000) do
      BEGIN
      Infos:=TEnregInfos.Create ;
      Infos.Inf1:=Q.FindField('T_AUXILIAIRE').AsString ;
      Infos.Inf2:=Q.FindField('T_NATUREAUXI').AsString ;
      Infos.Bol1:=(Q.FindField('T_LETTRABLE').AsString='X') ;
      ListeAUX.AddObject('',Infos) ; Inc(NbAux) ; Q.Next ;
      END ;
Ferme(Q) ;

TCNatPiece.Clear ;

Q:=OpenSql(' select CO_CODE from COMMUN where CO_TYPE="NTP" ',True) ;
While Not Q.Eof do
      BEGIN
      TCNatPiece.Items.Add(Q.FindField('CO_CODE').AsString) ;
      Q.Next ;
      END ;
Ferme(Q) ;
END ;

procedure TFReparMvt.VideInfos ;
Var i : integer ; infos : TEnregInfos ;
BEGIN
For i:=0 to ListeDev.Count-1 do
   BEGIN
   infos:=TEnregInfos(ListeDev.Objects[i]) ;
   infos.Free ;
   END ;
If ListeDev<>Nil then begin ListeDev.Clear ; ListeDev.Free ; end ;

For i:=0 to ListeJal.Count-1 do
   BEGIN
   infos:=TEnregInfos(ListeJal.Objects[i]) ;
   infos.Free ;
   END ;
If ListeJal<>Nil then begin ListeJal.Clear ; ListeJal.Free ; end ;

For i:=0 to ListeExo.Count-1 do
   BEGIN
   infos:=TEnregInfos(ListeExo.Objects[i]) ;
   infos.Free ;
   END ;
If ListeExo<>Nil then begin ListeExo.Clear ; ListeExo.Free ; end ;

For i:=0 to ListeSec.Count-1 do
   BEGIN
   infos:=TEnregInfos(ListeSec.Objects[i]) ;
   infos.Free ;
   END ;
If ListeSec<>Nil then begin ListeSec.Clear ; ListeSec.Free ; end ;

For i:=0 to ListeGen.Count-1 do
   BEGIN
   infos:=TEnregInfos(ListeGen.Objects[i]) ;
   infos.Free ;
   END ;
If ListeGen<>Nil then begin ListeGen.Clear ; ListeGen.Free ; end ;

For i:=0 to ListeBQE.Count-1 do
   BEGIN
   infos:=TEnregInfos(ListeBQE.Objects[i]) ;
   infos.Free ;
   END ;
If ListeBQE<>Nil then begin ListeBQE.Clear ; ListeBQE.Free ; end ;

For i:=0 to ListeAUX.Count-1 do
   BEGIN
   infos:=TEnregInfos(ListeAUX.Objects[i]) ;
   infos.Free ;
   END ;
If ListeAUX<>Nil then begin ListeAUX.Clear ; ListeAUX.Free ; end ;
TCNatPiece.Clear ;
END ;

procedure TFReparMvt.FExerciceChange(Sender: TObject);
begin
ExoToDates(FExercice.Value,FDateCpta1,FDateCpta2) ;
end;

Function TFReparMvt.VerifCptAttente(fb : TFichierBase) : Integer ;
Var QQ : Tquery ;
    OkAxe : Array[1..5] Of Boolean ;
    i : Integer ;
BEGIN
Result:=0 ; Fillchar(OkAxe,SizeOf(OkAxe),#0) ;
Case fb of
  fbGene : BEGIN
           QQ:=OpenSql(' Select * from generaux Where g_general="'+GeneAttend.Text+'"',True) ;
           If Not QQ.Eof Then
              BEGIN
              If (QQ.FindField('G_VENTILABLE').AsString='X') Or
                 (QQ.FindField('G_COLLECTIF').AsString='X') Or
                 (QQ.FindField('G_LETTRABLE').AsString='X') Then Result:=10 ;
              END Else Result:=9 ;
           Ferme(QQ) ;
           END ;
  fbSect : BEGIN
           QQ:=OpenSql(' Select Distinct(S_AXE) from section',TRUE) ;
           While Not QQ.Eof Do BEGIN i:=StrToInt(Copy(QQ.Fields[0].AsString,2,1)) ; OkAxe[i]:=TRUE ; QQ.Next ; END ;
           Ferme(QQ) ;
           For i:=1 To 5 Do If OkAxe[i] Then
              BEGIN
              QQ:=OpenSql(' Select * from section Where s_Section="'+VH^.Cpta[AxeToFb('A'+IntToStr(i))].Attente+'"',True) ;
              If QQ.Eof Then Result:=10+i ;
              Ferme(QQ) ;
              END ;
           END ;
  END ;
END ;

Function TFReparMvt.OkPourReparation : Integer ;
begin
  Result := 0 ;
  {JP 06/06/05 : FQ 13353 : Gestion de la réparation des écritures bordereau et libres}
  if cbTypeJal.Value = '' then
    Result := 16;
  if (cbTypeJal.Value = 'LIB') then begin
    if (Trim(FDateCpta1.Text) = '/  /') or (Trim(FDateCpta1.Text) = '/  /') or
       (StrToDate(FDateCpta1.Text) <> DebutDeMois(StrToDate(FDateCpta1.Text))) or
       (StrToDate(FDateCpta2.Text) <> FinDeMois(StrToDate(FDateCpta2.Text))) then
      Result := 17;
  end;

  if Result = 0 then begin
    if not CtrlPerExo(StrToDate(FDateCpta1.Text),StrToDate(FDateCpta2.Text)) then
      Result := 2;
  end;

  if Result = 0 then begin
     if CtrlEcrGen.Checked then Result := VerifCptAttente(fbGene) ;
     if CtrlEcrAna.Checked and (Result = 0) then Result := VerifCptAttente(fbSect) ;
   end ;
  if Result = 0 then RecupCrit ;
end;

procedure TFReparMvt.RecupCrit;
BEGIN
Fillchar(VCrit,SizeOf(VCrit),#0) ;
With VCrit Do
     BEGIN
     Etab:='' ; Cpt1:=FJal1.Text ; Cpt2:=FJal2.Text ;
     Date1:=StrToDate(FDateCpta1.Text) ;
     Date2:=StrToDate(FDateCpta2.Text) ;
     if FEtab.ItemIndex<>0 then Etab:=FEtab.Value ;
     Exo.Code:=FExercice.Value ;
     if FTypeEcriture.ItemIndex<>0 then QualifPiece:=FTypeEcriture.Value ;
     If FNumPiece1.Text<>'' then GL.NumPiece1:=StrToInt(FNumPiece1.Text) else GL.NumPiece1:=0 ;
     If FNumPiece2.Text<>'' then GL.NumPiece2:=StrToInt(FNumPiece2.Text) else GL.NumPiece2:=999999999 ;
     END ;
END ;


procedure TFReparMvt.Reparation ;
Var OnContinue : Boolean ;
BEGIN
StopVerif:= FALSE; OnContinue:=TRUE ; NbError:=0 ;
Application.ProcessMessages ; OkVerif:=True ;
Try
  BeginTrans ; If CtrlDetGen.Checked Then SQLEcr ; CommitTrans ;
Except
  MsgRien.Execute(8,Caption,'') ; Rollback ; OnContinue:=FALSE ;
End ;

If OnContinue And CtrlEcrGen.Checked And (Not StopVerif) Then
   BEGIN
   Try
     BeginTrans ; SqlEquilEcr(1) ; CommitTrans ;
   Except
     MsgRien.Execute(8,Caption,'') ; Rollback ; OnContinue:=FALSE ;
   End ;
   END ;
If OnContinue And CtrlEcrAna.Checked And (Not StopVerif) Then
   BEGIN
   Try
     BeginTrans ; SqlEquilEcr(2) ; CommitTrans ;
   Except
     MsgRien.Execute(8,Caption,'') ; Rollback ; //OnContinue:=FALSE ;
   End ;
   END ;
(*
if not StopVerif then SqlAnal ;
*)
END ;

procedure TFReparMvt.BStopClick(Sender: TObject);
begin
StopVerif:=True ;
end;

procedure TFReparMvt.Corrige(Q : TQuery ; Champ, Valeur : String) ;
BEGIN
Q.Edit ; Q.FindField(Champ).AsString:=Valeur ; Q.Post ;
END ;

procedure TFReparMvt.CorrigeInt(Q : TQuery ; Champ : String ; Valeur : Integer) ;
BEGIN
Q.Edit ; Q.FindField(Champ).AsInteger:=Valeur ; Q.Post ;
END ;

function TFReparMvt.TestBreak : Boolean ;
BEGIN
Application.ProcessMessages ;
if StopVerif then
   BEGIN
   if MsgRien.Execute(4,Caption,'')<>mryes then StopVerif:=False ;
   END ;
Result:=StopVerif ;
END ;

Function TFReparMvt.SiEnreg : Boolean ;
Var AuMoinsUn : Boolean ;
BEGIN
Result:=False ; NbEnreg:=0 ; AuMoinsUn:=FALSE ;
If CtrlEcrGen.Checked Then
   BEGIN
   QSum.Close ;    { Somme D et C (Devise pivot et devise) de chaque Pièce }
   QSum.SQL.Clear ;

   {JP 02/06/05 : FQ 13353 : On travail par mois sur les saisies Libre et par N° de groupe en Bordereau et par pièce en ...
                  Par ailleurs, je récupère le numéro de ligne maximal sur le regroupement, pour calculer le numéro de
                  ligne de l'écriture de réparation}
   if cbTypeJal.Value = 'LIB' then
     QSum.SQL.Add(' Select E_JOURNAL, E_EXERCICE, MONTH(E_DATECOMPTABLE) AS MOIS, MAX(E_DATECOMPTABLE) AS E_DATECOMPTABLE, E_NUMEROPIECE, E_DEVISE, ')
   else if cbTypeJal.Value = 'BOR' then
     QSum.SQL.Add(' Select E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMGROUPEECR, E_DEVISE, ')
   else
     QSum.SQL.Add(' Select E_JOURNAL, E_EXERCICE, E_NUMEROPIECE, E_DATECOMPTABLE, E_DEVISE, ');

   QSum.SQL.Add(' Sum(E_DEBIT) SUMD, Sum(E_CREDIT) as SUMC, ') ;
   QSum.SQL.Add(' Sum(E_DEBITDEV) SUMDDEV, Sum(E_CREDITDEV) SUMCDEV, ') ;
   QSum.SQL.Add(' 0 SUMDEUR, 0 as SUMCEUR, ') ;
   QSum.SQL.Add(' MIN(E_DEBIT) MIND, MAX(E_DEBIT) MAXD, ') ;

   QSum.SQL.Add(' MIN(E_CREDIT) MINC, MAX(E_CREDIT) MAXC, E_QUALIFPIECE ');

   QSum.SQL.Add(' FROM ECRITURE ') ;
   QSum.SQL.Add(' LEFT JOIN JOURNAL ON E_JOURNAL=J_JOURNAL ') ;
   QSum.SQL.Add(' Where J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ANA" ') ;
   QSum.SQL.Add(' And J_JOURNAL<>"'+w_w+'" ') ;

   {JP 02/06/05 : FQ 13353 : On modifie la requête en fonction du type de journaux}
   if cbTypeJal.Value = 'LIB' then
     QSum.SQL.Add(' AND E_MODESAISIE = "LIB" ')
   else if cbTypeJal.Value = 'BOR' then
     QSum.SQL.Add(' AND E_MODESAISIE = "BOR" ')
   else
     QSum.SQL.Add(' AND (E_MODESAISIE = "-" OR E_MODESAISIE = "") ') ;

   QSum.SQL.Add(' And E_EXERCICE="'+VCrit.Exo.Code+'" ') ;
   if VCrit.QualifPiece<>'' then QSum.SQL.Add(' And E_QUALIFPIECE="'+VCrit.QualifPiece+'" ') ;
   if VCrit.Cpt1<>'' then QSum.SQL.Add(' And E_JOURNAL>="'+VCrit.Cpt1+'"') ;
   if VCrit.Cpt2<>'' then QSum.SQL.Add(' And E_JOURNAL<="'+VCrit.Cpt2+'"') ;
   QSum.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(VCrit.GL.NumPiece1)+' and E_NUMEROPIECE<='+IntTostr(VCrit.GL.NumPiece2)+' ') ;
   if VCrit.Etab<>'' then QSum.SQL.Add(' And E_ETABLISSEMENT="'+VCrit.Etab+'" ') ;
   QSum.SQL.Add(' AND E_QUALIFPIECE<>"C" ') ;
   QSum.SQL.Add(' AND E_DATECOMPTABLE>="'+USDATETIME(VCrit.Date1)+'"') ;
   QSum.SQL.Add(' And E_DATECOMPTABLE<="'+USDATETIME(VCrit.Date2)+'"') ;

   {JP 02/06/05 : FQ 13353 : On modifie la clause GROUP By de la requête en fonction du type de journaux}
   if cbTypeJal.Value = 'LIB' then
     QSum.SQL.Add(' GROUP BY E_JOURNAL, E_EXERCICE, MONTH(E_DATECOMPTABLE), E_NUMEROPIECE, E_DEVISE , E_QUALIFPIECE')
   else if cbTypeJal.Value = 'BOR' then
     QSum.SQL.Add(' GROUP BY E_JOURNAL, E_EXERCICE, E_NUMEROPIECE, E_DATECOMPTABLE, E_NUMGROUPEECR, E_DEVISE , E_QUALIFPIECE')
   else
     QSum.SQL.Add(' GROUP BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_DEVISE, E_QUALIFPIECE ');

   ChangeSql(QSum) ; //QSum.Prepare ;
   PrepareSQLODBC(QSum) ;
   QSum.Open ; AuMoinsUn:=Not QSum.EOF ;
   END ;
Application.ProcessMessages ;
{--------------------------------------------------------------------------------------------}
If CtrlEcrAna.Checked Then
   BEGIN
   QSumGeAn.Close ;   { Somme D et C (Devise pivot et devise) de chaque Pièce/Ligne }
   QSumGeAn.SQL.Clear ;
   QSumGeAn.SQL.Add(' Select E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_DEVISE, ') ;
   QSumGeAn.SQL.Add(' E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV, ') ;
   QSumGeAn.SQL.Add(' 0, 0, E_NUMLIGNE, ') ;
   QSumGeAn.Sql.Add(' E_REFINTERNE, E_GENERAL, E_ETABLISSEMENT, ') ;
   QSumGeAn.Sql.Add(' E_NATUREPIECE, E_ECRANOUVEAU,  ') ;
   QSumGeAn.Sql.Add(' G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, ') ;
   QSumGeAn.Sql.Add(' E_TYPEANOUVEAU, E_QUALIFPIECE ') ;
   QSumGeAn.SQL.Add(' From ECRITURE ') ;
   QSumGeAn.SQL.Add(' LEFT JOIN JOURNAL ON E_JOURNAL=J_JOURNAL ') ;
   QSumGeAn.Sql.Add(' Left Join GENERAUX on E_GENERAL=G_GENERAL where E_ANA="X" ') ;
   QSumGeAn.SQL.Add(' AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANA" ') ;
   QSumGeAn.SQL.Add(' And E_EXERCICE="'+VCrit.Exo.Code+'" ') ;

   if VCrit.QualifPiece<>'' then QSumGeAn.SQL.Add(' And E_QUALIFPIECE="'+VCrit.QualifPiece+'" ') ;
   if VCrit.Cpt1<>'' then QSumGeAn.SQL.Add(' And E_JOURNAL>="'+VCrit.Cpt1+'"') ;
   if VCrit.Cpt2<>'' then QSumGeAn.SQL.Add(' And E_JOURNAL<="'+VCrit.Cpt2+'"') ;
   QSumGeAn.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(VCrit.GL.NumPiece1)+' and E_NUMEROPIECE<='+IntTostr(VCrit.GL.NumPiece2)+' ') ;
   QSumGeAn.SQL.Add(' And E_DATECOMPTABLE>="'+USDATETIME(VCrit.Date1)+'" ') ;
   QSumGeAn.SQL.Add(' And E_DATECOMPTABLE<="'+USDATETIME(VCrit.Date2)+'" ') ;
   if VCrit.Etab<>'' then QSumGeAn.SQL.Add(' And E_ETABLISSEMENT="'+VCrit.Etab+'" ') ;
   QSumGeAn.SQL.Add(' And E_QUALIFPIECE<>"C" ') ;

   {JP 02/06/05 : FQ 13353 : On modifie la requête en fonction du type de journaux}
   if cbTypeJal.Value = 'LIB' then
     QSumGeAn.SQL.Add(' AND E_MODESAISIE = "LIB" ')
   else if cbTypeJal.Value = 'BOR' then
     QSumGeAn.SQL.Add(' AND E_MODESAISIE = "BOR" ')
   else
     QSumGeAn.SQL.Add(' AND (E_MODESAISIE = "-" OR E_MODESAISIE = "") ') ;

   { Construction de la clause Order By de la SQL }
   QSumGeAn.SQL.Add(' Order BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_DEVISE, E_QUALIFPIECE ') ;
   ChangeSql(QSumGeAn) ; //QSumGeAn.Prepare ;
   PrepareSQLODBC(QSumGeAn) ;
   QSumGeAn.Open ; AuMoinsUn:=AuMoinsUn Or (Not QSumGeAn.EOF) ;
   END ;
Application.ProcessMessages ;
{--------------------------------------------------------------------------------------------}
{ Verif des Mvts Analytique }

If CtrlDetAna.Checked Then
   BEGIN
   QAnal.Close ; QAnal.Sql.Clear ;
   QAnal.SQL.Add(' Select Y_JOURNAL, Y_DATECOMPTABLE, Y_NUMLIGNE,  ') ;
   QAnal.SQL.Add(' Y_AXE, Y_DEVISE, Y_REFINTERNE, Y_GENERAL, Y_NUMVENTIL, ') ;
   QAnal.SQL.Add(' Y_SECTION, Y_ECRANOUVEAU,  Y_NATUREPIECE, Y_ETABLISSEMENT, ') ;
   QAnal.SQL.Add(' Y_TYPEANALYTIQUE, ') ;
   QAnal.SQL.Add(' Y_DEBIT, Y_CREDIT, Y_DEBITDEV, Y_CREDITDEV, ') ;
   QAnal.SQL.Add(' Y_NUMEROPIECE, Y_CONFIDENTIEL, Y_EXERCICE, ') ;
   QAnal.SQL.Add(' Y_TAUXDEV, Y_QUALIFPIECE ') ;
   QAnal.SQL.Add(' From ANALYTIQ ') ;
   QAnal.SQL.Add(' LEFT JOIN JOURNAL ON Y_JOURNAL=J_JOURNAL ') ;
   QAnal.SQL.Add(' Where ') ;
   QAnal.SQL.Add(' J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ANA" ') ;
   QAnal.SQL.Add(' AND J_JOURNAL<>"'+w_w+'" ') ;
   QAnal.SQL.Add(' And Y_EXERCICE="'+VCrit.Exo.Code+'" ') ;
   if VCrit.QualifPiece<>'' then QAnal.SQL.Add(' And Y_QUALIFPIECE="'+VCrit.QualifPiece+'" ') ;
   if VCrit.Cpt1<>'' then QAnal.SQL.Add(' And Y_JOURNAL>="'+VCrit.Cpt1+'" ') ;
   if VCrit.Cpt2<>'' then QAnal.SQL.Add(' And Y_JOURNAL<="'+VCrit.Cpt2+'" ') ;
   QAnal.SQL.Add(' And Y_NUMEROPIECE>='+IntToStr(VCrit.GL.NumPiece1)+' and Y_NUMEROPIECE<='+IntTostr(VCrit.GL.NumPiece2)+' ') ;
   QAnal.SQL.Add(' And Y_DATECOMPTABLE>="'+USDATETIME(VCrit.Date1)+'" ') ;
   QAnal.SQL.Add(' And Y_DATECOMPTABLE<="'+USDATETIME(VCrit.Date2)+'" ') ;
   if VCrit.Etab<>'' then QAnal.SQL.Add(' And Y_ETABLISSEMENT="'+VCrit.Etab+'" ') ;
   QAnal.SQL.Add(' and Y_QUALIFPIECE<>"C" ') ;
   { Construction de la clause Order By de la SQL }
   QAnal.SQL.Add(' Order By Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE ') ;
   ChangeSql(QAnal) ; //QAnal.Prepare ;
   PrepareSQLODBC(QAnal) ;
   QAnal.Open ; AuMoinsUn:=AuMoinsUn Or (Not QAnal.EOF) ;
   END ;

Application.ProcessMessages ;
{--------------------------------------------------------------------------------------------}
If CtrlDetGen.Checked Then
   BEGIN
   QEcr.Close ; QEcr.SQL.Clear ;
   QEcr.SQL.Add(' Select E_JOURNAL, E_DATECOMPTABLE, E_NUMLIGNE,  ') ;
   QEcr.SQL.Add(' E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_MODEPAIE, ') ;
   QEcr.SQL.Add(' E_DEBIT,  E_CREDIT , E_DEBITDEV, E_CREDITDEV, ') ;
   QEcr.SQL.Add(' E_ANA, E_REFINTERNE, E_DATEECHEANCE, ') ;
   QEcr.SQL.Add(' E_NUMECHE, E_ETABLISSEMENT, E_ETATLETTRAGE, E_ECHE, ') ;
   QEcr.SQL.Add(' E_ECRANOUVEAU, E_COUVERTURE, E_COUVERTUREDEV, E_NATUREPIECE, ') ;
   QEcr.SQL.Add(' E_NUMEROPIECE, E_CONFIDENTIEL, E_EXERCICE,  ') ;
   QEcr.SQL.Add(' E_TAUXDEV, E_QUALIFPIECE ') ;
   QEcr.SQL.Add(' From ECRITURE ') ;
   QEcr.SQL.Add(' Where E_JOURNAL<>"'+w_w+'" ') ;
   QEcr.SQL.Add(' And E_EXERCICE="'+VCrit.Exo.Code+'" ') ;
   if VCrit.QualifPiece<>'' then QEcr.SQL.Add(' And E_QUALIFPIECE="'+VCrit.QualifPiece+'" ') ;
   if VCrit.Cpt1<>'' then QEcr.SQL.Add(' And E_JOURNAL>="'+VCrit.Cpt1+'"') ;
   if VCrit.Cpt2<>'' then QEcr.SQL.Add(' And E_JOURNAL<="'+VCrit.Cpt2+'"') ;
   QEcr.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(VCrit.GL.NumPiece1)+' and E_NUMEROPIECE<='+IntTostr(VCrit.GL.NumPiece2)+' ') ;
   QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+USDATETIME(VCrit.Date1)+'"') ;
   QEcr.SQL.Add(' And E_DATECOMPTABLE<="'+USDATETIME(VCrit.Date2)+'"') ;
   if VCrit.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+VCrit.Etab+'" ') ;
   { Construction de la clause Order By de la SQL }
   QEcr.SQL.Add(' And E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU="N"') ;

   {JP 02/06/05 : FQ 13353 : On modifie la requête en fonction du type de journaux}
   if cbTypeJal.Value = 'LIB' then
     QEcr.SQL.Add(' AND E_MODESAISIE = "LIB" ')
   else if cbTypeJal.Value = 'BOR' then
     QEcr.SQL.Add(' AND E_MODESAISIE = "BOR" ')
   else
     QEcr.SQL.Add(' AND (E_MODESAISIE = "-" OR E_MODESAISIE = "") ') ;

   QEcr.SQL.Add(' order BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ') ;
   ChangeSql(QEcr) ; //QEcr.Prepare ;
   PrepareSQLODBC(QEcr) ;
   QEcr.Open ; AuMoinsUn:=AuMoinsUn Or (Not QEcr.EOF) ;
   END ;
Application.ProcessMessages ;
if AuMoinsUn then
  BEGIN
  Result:=True ;
  InitMove(10000,MsgBar.Mess[0]) ;
  END ;
END ;


procedure TFReparMvt.SqlEquilEcr(Ec : Byte) ; {EC=1 : Verif équilibre Ecriture ; EC=2 : Verif équilibre Ecr-Ana}
BEGIN
if Ec=1 then
   BEGIN
   { Somme D et C (Devise pivot et devise) de chaque Pièce }
   QSum.First ;
   While not QSum.Eof do
         BEGIN
         MoveCur(False) ;
         If Mvt<>NIL Then BEGIN Mvt.FREE ; Mvt:=NIL ; END ;
         EnregEquil(QSum, Ec) ;
         if Not EquilEcr then OkVerif:=False  ; { Test si la pièce est équilibrée : ECRITURE seule }
         if TestBreak then Break ;
         QSum.Next;
         END ;
   END Else
   BEGIN
   { Somme D et C (Devise pivot et devise) de chaque Pièce/Ligne pour vérif ECR-ANA}
   QSumGeAn.First ; PrepareAnal ;
   While not QSumGeAn.Eof do
         BEGIN
         MoveCur(False) ;
         If Mvt<>NIL Then BEGIN Mvt.FREE ; Mvt:=NIL ; END ;
         EnregEquilA ;
         if Not EquilAnal then OkVerif:=False  ; {Test de l' équilibre la pièce/Ligne ECRITURE Avec la pièce/Ligne de L'ANALYTIQUE }
         if TestBreak then Break ;
         QSumGeAn.Next;
         END ;
   END ;
END ;

Procedure TFReparMvt.PrepareAnal ; { Requete paramatrée sur la table ANALYTIQE}
BEGIN
QAnaGene.Close ; QAnaGene.Sql.Clear ;
QAnaGene.SQL.Add(' Select') ;
QAnaGene.SQL.Add(' Sum(Y_DEBIT) SUMD, Sum(Y_CREDIT) SUMC, ') ;
QAnaGene.SQL.Add(' Sum(Y_DEBITDEV) SUMDDEV, Sum(Y_CREDITDEV) SUMCDEV, ') ;
QAnaGene.SQL.Add(' 0 SUMDEUR, 0 SUMCEUR, Y_AXE AXET, ') ;
QAnaGene.SQL.Add(' Sum(Y_POURCENTAGE) POURCENT ') ;
//SG6 09.03.05 mode analytique croisaxe
if VH^.AnaCroisaxe then
  QAnaGene.SQL.Add(', Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5') ;
//Fin SG6
QAnaGene.SQL.Add(' From ANALYTIQ ') ;
QAnaGene.SQL.Add(' Where ') ;
QAnaGene.SQL.Add(' Y_JOURNAL=:JAL ') ;
QAnaGene.SQL.Add(' And Y_EXERCICE=:EXO ') ;
QAnaGene.SQL.Add(' And Y_DATECOMPTABLE=:DAT ') ;
QAnaGene.SQL.Add(' And Y_NUMEROPIECE=:PIE ') ;
QAnaGene.SQL.Add(' And Y_NUMLIGNE=:LIG ') ;
QAnaGene.SQL.Add(' And Y_DEVISE=:DEV ') ;
//QAnaGene.SQL.Add(' And Y_ETABLISSEMENT=:ETA ') ;
QAnaGene.SQL.Add(' And Y_ETABLISSEMENT<>:ETA ') ;
//QAnaGene.SQL.Add(' And Y_AXE=:AXE ') ;
QAnaGene.SQL.Add(' And (Y_AXE=:AXE1 or Y_AXE=:AXE2 or Y_AXE=:AXE3 or Y_AXE=:AXE4 or Y_AXE=:AXE5) ') ;
QAnaGene.SQL.Add(' And Y_QUALIFPIECE=:QUALIF ') ;
QAnaGene.SQL.Add(' And Y_QUALIFPIECE<>"C" ') ;
{ Construction de la clause Group By de la SQL }
//SG6 09.03.05 Gestion ana croisaxe
if not VH^.AnaCroisaxe then
  QAnaGene.SQL.Add(' Group By Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE , Y_AXE, Y_DEVISE ')
else
  QAnaGene.SQL.Add(' Group By Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE , Y_AXE, Y_DEVISE, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 ') ;
//Fin SG6
ChangeSql(QAnaGene) ; //QAnaGene.Prepare ;
PrepareSQLODBC(QAnaGene) ;
END ;


Procedure TFReparMvt.EnregEquilA ; { Enregistrement des valeurs de la 1ère requete sur la table ECRITURE... }
Var J : Byte ;                 { ...Pour la vérif de l'equilibre de ECR-ANA }
BEGIN
Mvt:=TInfoMvt.Create ;
Mvt.JOURNAL:=QSumGeAn.Fields[0].AsString ; // journal
Mvt.EXERCICE:=QSumGeAn.FieldS[1].AsString ;
Mvt.DATECOMPTABLE:=QSumGeAn.Fields[2].AsDateTime ;
Mvt.NUMEROPIECE:=QSumGeAn.Fields[3].AsInteger ;
Mvt.DEVISE:=QSumGeAn.Fields[4].AsString ;
Mvt.DEBIT:=QSumGeAn.Fields[5].AsFloat ;
Mvt.CREDIT:=QSumGeAn.FieldS[6].AsFloat ;
Mvt.DEBITDEV:=QSumGeAn.FieldS[7].AsFloat ;
Mvt.CREDITDEV:=QSumGeAn.FieldS[8].AsFloat ;
Mvt.NUMLIGNE:=QSumGeAn.FieldS[11].AsInteger ;
Mvt.REFINTERNE:=QSumGeAn.Fields[12].AsString ;
Mvt.GENERAL:=QSumGeAn.Fields[13].AsString ;
Mvt.ETABLISSEMENT:=QSumGeAn.Fields[14].AsString ;
Mvt.NATUREPIECE:=QSumGeAn.Fields[15].AsString ;
Mvt.ECRANOUVEAU:=QSumGeAn.Fields[16].AsString ;
(**)
for j:=1 to high(Mvt.AxePlus) do Mvt.AxePlus[j]:='A0' ;
If QSumGeAn.Fields[17].AsString='X' then Mvt.AxePlus[1]:='A1' ;
If QSumGeAn.Fields[18].AsString='X' then Mvt.AxePlus[2]:='A2' ;
If QSumGeAn.Fields[19].AsString='X' then Mvt.AxePlus[3]:='A3' ;
If QSumGeAn.Fields[20].AsString='X' then Mvt.AxePlus[4]:='A4' ;
If QSumGeAn.Fields[21].AsString='X' then Mvt.AxePlus[5]:='A5' ;
(**)
Mvt.QUALIFPIECE:=QSumGeAn.Fields[23].AsString ;
END ;

Procedure TFReparMvt.EnregEquil(Q : TQuery ; Quelle : byte) ; { Enregistrement des valeurs de la requete...}
                                         { ...Pour la vérif de l'équilibre sur la table ECRITURE seul !!!}
BEGIN
  Mvt:=TInfoMvt.Create ;

  Mvt.DATECOMPTABLE := Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  Mvt.JOURNAL       := Q.FindField('E_JOURNAL').AsString ; // journal
  Mvt.EXERCICE      := Q.FindField('E_EXERCICE').AsString ;
  Mvt.NUMEROPIECE   := Q.FindField('E_NUMEROPIECE').AsInteger ;
  Mvt.DEVISE        := Q.FindField('E_DEVISE').AsString ;
  Mvt.DEBIT         := Q.FindField('SUMD').AsFloat ;
  Mvt.CREDIT        := Q.FindField('SUMC').AsFloat ;
  Mvt.DEBITDEV      := Q.FindField('SUMDDEV').AsFloat ;
  Mvt.CREDITDEV     := Q.FindField('SUMCDEV').AsFloat ;
  Mvt.DEBITMin      := Q.FindField('MIND').AsFloat ;
  Mvt.DEBITMax      := Q.FindField('MAXD').AsFloat ;
  Mvt.CREDITMin     := Q.FindField('MINC').AsFloat ;
  Mvt.CREDITMax     := Q.FindField('MAXD').AsFloat ;
  {JP 06/06/05 : FQ 13353 : Pour les écritures de bordereau, on a besoin d'une information supplémentaire
                 qui est le numéro de regroupement des écritures : on le stocke dans le numéro de ventilation
                 qui n'est pas utilisé ici}
  if cbTypeJal.Value = 'BOR' then
    Mvt.NUMVENTIL := Q.FindField('E_NUMGROUPEECR').AsInteger;

  if cbTypeJal.Value = '-' then
    Mvt.QUALIFPIECE := Q.FindField('E_QUALIFPIECE').AsString
  else
    {JP 02/06/05 : en mode libre ou bordereau, QualifPièce est à N}
    Mvt.QUALIFPIECE := 'N';

  {JP 02/06/05 : FQ 13353 : On récupère le mode de saisie depuis la combo de filtre}
  Mvt.MODESAISIE := cbTypeJal.Value;
end;

Function TFReparMvt.EquilEcr : Boolean ; { Test de l'équilibre sur la Table ECRITURE seul !!!}
Var CBonP, CBonD, {CBonE,} Resultat, OkOK,CBonPAbs  : Boolean ;
    MD, MC, MDD, MCD, MDE, MCE : Double ;
    DecDev           : Byte ;
BEGIN
Resultat:=True ; DecDev:=V_PGI.OkDecV ; OkOK:=True ;
If Mvt.DEVISE='' then OkOk:=False  ;
If not OkOk then  begin Result:=Resultat ; Exit ; end ;
{ Devise Pivot : Pièce équilibrée ? }
MD:=Arrondi(Mvt.DEBIT,V_PGI.OkdecV) ; MC:=Arrondi(Mvt.CREDIT,V_PGI.OkdecV) ;
CBonP:=(MD=MC) ;
{ Devise pivot : Pièce à 0 ? }
CBonPAbs:=True ;
if (Arrondi(Mvt.DEBITMIN,V_PGI.OkdecV)=0) and (Arrondi(Mvt.DEBITMAX,V_PGI.OkdecV)=0) and
   (Arrondi(Mvt.CREDITMIN,V_PGI.OkdecV)=0) and (Arrondi(Mvt.CREDITMAX,V_PGI.OkdecV)=0) then CBonPAbs:=False ;
{ Devise : Pièce équilibrée ? }
if Mvt.DEVISE<>V_PGI.DevisePivot then
   BEGIN
   TrouveDecDev(Mvt.DEVISE,DecDev) ;
   MDD:=Arrondi(Mvt.DEBITDEV,DecDev)   ; MCD:=Arrondi(Mvt.CREDITDEV,DecDev) ;
// CriCri : 02/07/97
// DecDev -> DecEuro
   END Else
   BEGIN
   MDD:=Arrondi(Mvt.DEBITDEV,V_PGI.OkdecV)  ; MCD:=Arrondi(Mvt.CREDITDEV,V_PGI.OkdecV) ;
// CriCri : 02/07/97
// DecDev -> DecEuro
   //CBonD:=(MDD=MCD) ;
   END ;
CBonD:=(MDD=MCD) ;
//If V_PGI.SAV then CBonE:=(MDE=MCE) else
//CBonE:=True ;
{ Pièce à zéro en Devise Pivot ; Affichage dans la grille }
MDE := 0;
MCE := 0;
if not CBonPAbs then                     // A 1?
   BEGIN
   ErrEquilEcr(Resultat,6,MD,MC,MDD,MCD,MDE,MCE,DecDev) ;
   END Else
{ Pièce non équilibré en Devise Pivot et en Devise ; Affichage dans la grille }
if (not CBonP) and (not CBonD) then    // A 13
   BEGIN
   ErrEquilEcr(Resultat,1,MD,MC,MDD,MCD,MDE,MCE,DecDev) ;
   END Else
{ Pièce non équilibré en Devise Pivot ; Affichage dans la grille }
if not CBonP then                     // A 11
   BEGIN
   ErrEquilEcr(Resultat,2,MD,MC,MDD,MCD,MDE,MCE,DecDev) ;
   END Else
{ Pièce non équilibré en Devise ; Affichage dans la grille }
if not CBonD then                    // A 12
   BEGIN
   ErrEquilEcr(Resultat,3,MD,MC,MDD,MCD,MDE,MCE,DecDev) ;
   END ;
{ Si la devise est renseigné les montants en DEBITDEV et CREDITDEV doivent etre <> 0 }
If Mvt.DEVISE<>V_PGI.DevisePivot then
   If ((Mvt.DEBITDEV=0)and(Mvt.CREDITDEV=0)) then
      BEGIN
      ErrEquilEcr(Resultat, 4,MD,MC,0,0,0,0,DecDev) ;
      END ;
{ Pièce non équilibré en Euro ; Affichage dans la grille }
(* GP
if not CBonE then                    // A 12
   BEGIN
   ErrEquilEcr(Resultat,5,MD,MC,MDD,MCD,MDE,MCE) ;
   END ;
*)
Result:=Resultat ;

END ;

Procedure TFReparMvt.LaListeAdd (G : DelInfo ) ;
BEGIN
if NbError=MaxNbErreur then MsgRien.Execute(7,Caption,'') ;
if NbError>MaxNbErreur then Exit ;
LaListe.Add(G);
END ;

////////////////////////////////////////////////////////////////////////////////
Procedure TFReparMvt.ErrEquilEcr(var Err : Boolean ; E : Byte ; D,C,DD,CD,DE,CE : Double ; DecE : Integer) ;
var Q : TQuery ;      {Recherche RefInterne sur Piece Non Equil}
    //T : TQUERY ;      {Corrections si possible }
    St, Cpt : String ;
    OkP, OkD, OkBanque : Boolean ;
    D1, C1, DD1, CD1 : Double ;
    lTobEcr : Tob; // la grosse TOB
    lRecError : TRecError;

begin
  St := 'SELECT E_REFINTERNE, E_NUMLIGNE, E_NATUREPIECE, E_QUALIFPIECE, E_ETABLISSEMENT, ' ;
  St := St + ' E_DEBITDEV, E_CREDITDEV, E_DEVISE, E_DATECOMPTABLE, E_TAUXDEV, E_ECRANOUVEAU, E_MODEPAIE, E_DATEECHEANCE, E_DATEVALEUR, ' ; // SBO 31/10/2006 manque virgule, fait planter le findfield plus bas sur E_DATEVALEUR en mode journal de banque
  St := St + ' E_ENCAISSEMENT, E_CONTREPARTIEGEN';
  St := St + ' FROM ECRITURE WHERE E_JOURNAL = "' + Mvt.Journal + '" ' ;
  St := St + ' AND E_EXERCICE = "' + Mvt.Exercice + '" ' ;

  {JP 02/06/05 : FQ 13353 : On travail par mois sur les saisies Libre et par date et regroupement en Bordereau
                 et par pièce en saisie pièce}
  if cbTypeJal.Value = 'LIB' then
    St := St + ' AND E_DATECOMPTABLE BETWEEN "' + USDATETIME(DebutDeMois(Mvt.DATECOMPTABLE)) + '" ' +
               'AND "' + USDATETIME(FinDeMois(Mvt.DATECOMPTABLE)) + '" '
//  else if cbTypeJal.Value = 'BOR' then
  //  St := St + ' AND E_DATECOMPTABLE = "' + USDATETIME(Mvt.DATECOMPTABLE) + '" AND E_NUMGROUPEECR = ' + IntToStr(Mvt.NUMVENTIL)
  else
    St := St + ' AND E_DATECOMPTABLE = "' + USDATETIME(Mvt.DATECOMPTABLE) + '" ';

  St := St + ' AND E_NUMEROPIECE =' + IntToStr(Mvt.NUMEROPIECE) + ' ' ;
  St := St + ' AND E_QUALIFPIECE = "' + Mvt.QUALIFPIECE + '" ' ;
  {JP 06/06/05 : Suppression de la date comptable du order by car elle n'etait pas compatible avec les traitements
                 autres que la saisie pièce et la date ne servait à rien}
  St := St + ' ORDER BY E_JOURNAL, E_EXERCICE, E_NUMEROPIECE, E_NUMLIGNE DESC, E_DEVISE ';

  Q := OpenSql(St,True) ;

  if Q.EOF then
  begin
    Ferme(Q);
    Exit;
  end;

  // Calcul du numéro de ligne
  Mvt.NUMLIGNE   := Q.FindField('E_NUMLIGNE').AsInteger;
  Mvt.REFINTERNE := Q.FindField('E_REFINTERNE').AsString ;

  // Ecart de change : ne pas faire T.Insert
  if (E = 4) And
     ((Q.FindField('E_DEBITDEV').AsFloat=0) and (Q.FindField('E_CREDITDEV').AsFloat=0)) And
     (Q.FindField('E_NATUREPIECE').AsSTring='ECC') then
  begin
    Ferme(Q);
    Exit;
  end;

  // Création de la Tob pour l'engistrement
  lTobEcr := Tob.Create('ECRITURE', nil , -1);
  CPutDefautEcr( lTobEcr );

  // Remplissage des champs de la tob par les valeurs dans le MVT
  lTobEcr.PutValue('E_JOURNAL',       Mvt.Journal );
  lTobEcr.PutValue('E_MODESAISIE',    Mvt.ModeSaisie);
  lTobEcr.PutValue('E_LIBELLE',       MsgLibel2.Mess[2]);

  CRemplirDateComptable( lTobEcr,     Mvt.DATECOMPTABLE );

  lTobEcr.PutValue('E_NUMEROPIECE',   Mvt.NUMEROPIECE);
  {JP 06/06/05 : Le numéro de ligne était calculé par rapport à un compte et non une "pièce". On récupère
                 maintenant le numéro de ligne maximal de la pièce}
  lTobEcr.PutValue('E_NUMLIGNE',      Mvt.NUMLIGNE + 1);

  lTobEcr.PutValue('E_DEVISE',        Mvt.DEVISE);
  lTobEcr.PutValue('E_TAUXDEV',       Q.FindField('E_TAUXDEV').AsFloat);
  lTobEcr.PutValue('E_REFINTERNE',    Mvt.REFINTERNE);

  Cpt := GeneAttend.Text ;
  OkBanque := FALSE ;
  if Existe(2,Mvt.JOURNAL) then
  begin
    if (InfAutres.J_NATUREJAL = 'BQE') And (InfAutres.J_CONTREPARTIE <> '') then
    begin
      Cpt := InfAutres.J_CONTREPARTIE ;
      OkBanque := TRUE ;
    end;
  end;

  lTobEcr.PutValue('E_GENERAL',         Cpt);
  lTobEcr.PutValue('E_NATUREPIECE',     Q.FindField('E_NATUREPIECE').AsString);
  lTobEcr.PutValue('E_QUALIFPIECE',     Q.FindField('E_QUALIFPIECE').AsString);
  lTobEcr.PutValue('E_ETABLISSEMENT',   Q.FindField('E_ETABLISSEMENT').AsString);
  lTobEcr.PutValue('E_ECRANOUVEAU',     Q.FindField('E_ECRANOUVEAU').AsString);
  lTobEcr.PutValue('E_CONTREPARTIEGEN', Q.FindField('E_CONTREPARTIEGEN').AsString);
  {JP 02/06/05 : FQ 15971 : permet d'éviter un message du genre "Champs E_ENCAISSEMENT obligatoires"}
  lTobEcr.PutValue('E_ENCAISSEMENT'     , Q.FindField('E_ENCAISSEMENT').AsString);
  lTobEcr.PutValue('E_IO',              'X');
  {JP 06/06/05 : FQ 13353 : En bordereau il faut ajouter le regroupement}
  if cbTypeJal.Value = 'BOR' then
    lTobEcr.SetInteger('E_NUMGROUPEECR',  Mvt.NUMVENTIL);

  CSupprimerInfoLettrage( lTobEcr );

  if OkBanque Then
  begin
    if Q.FindField('E_MODEPAIE').AsString <> '' then
    begin
      lTobEcr.PutValue('E_DATEECHEANCE', Q.FindField('E_DATEECHEANCE').AsDateTime);
      lTobEcr.PutValue('E_MODEPAIE',     Q.FindField('E_MODEPAIE').AsString);
      lTobEcr.PutValue('E_DATEVALEUR',   Q.FindField('E_DATEVALEUR').AsDateTime);
    end
    else
    begin
      lTobEcr.PutValue('E_DATEECHEANCE', Q.FindField('E_DATECOMPTABLE').AsDateTime);
      lTobEcr.PutValue('E_DATEVALEUR',   Q.FindField('E_DATECOMPTABLE').AsDateTime);
      lTobEcr.PutValue('E_MODEPAIE',     MpAttend.Value);
    end;
    lTobEcr.PutValue('E_NUMECHE', 1);
    lTobEcr.PutValue('E_ECHE', 'X');
  end;

  Err := False;

  case E of // 1 : Pivot + Dev ; 2 : Pivot ; 3 : Devise ; 4 : Piece Devise avec Monants devise=0 ; 5 : Euro ; 6 : Pièce à 0
   // Pivot + Dev
   1 : if (lTobEcr <> nil) then
       begin
         OkP := LeSolde(D,C,FALSE,FALSE,V_PGI.OkDecV,D1,C1) ;
         OkD := LeSolde(DD,CD,FALSE,FALSE,DecE,DD1,CD1) ; //MontantEuro:=LeSolde(DE,CE) ;
         if OkP Then
         begin
           lTobEcr.PutValue('E_DEBIT',  D1);
           lTobEcr.PutValue('E_CREDIT', C1);
         end;

         if OkD then
         begin
           lTobEcr.PutValue('E_DEBITDEV',  DD1);
           lTobEcr.PutValue('E_CREDITDEV', CD1);
         end;

         if OkP or OkD then LaListeAdd(GoListe1(Mvt,'C',MsgLibel.Mess[15]+FloatToStr(D)+'/'+FloatToStr(DD)+' ; '+MsgLibel.Mess[16]+FloatToStr(C)+'/'+FloatToStr(CD),2,'G'));
       end
       else
       begin
         LaListeAdd(GoListe1(Mvt,'C',MsgLibel.Mess[15]+FloatToStr(D)+'/'+FloatToStr(DD)+' ; '+MsgLibel.Mess[16]+FloatToStr(C)+'/'+FloatToStr(CD),2,'G'));
       end;

   // Pivot
   2 : begin
         OkP := LeSolde(D,C,FALSE,FALSE,V_PGI.OkDecV,D1,C1) ;
         if OkP then
         begin
           lTobEcr.PutValue('E_DEBIT',  D1);
           lTobEcr.PutValue('E_CREDIT', C1);
         end;
         LaListeAdd(GoListe1(Mvt,'C',MsgLibel.Mess[15]+FloatToStr(D)+' ; '+MsgLibel.Mess[16]+FloatToStr(C),0,'G'));
       end;

   // Devise
   3 : begin
         OkD := LeSolde(DD,CD,FALSE,FALSE,DecE,DD1,CD1) ; //MontantEuro:=LeSolde(DE,CE) ;
         lTobEcr.PutValue('E_DEBITDEV',  DD1);
         lTobEcr.PutValue('E_CREDITDEV', CD1);
         if OkD then
           LaListeAdd(GoListe1(Mvt,'C',MsgLibel.Mess[15]+FloatToStr(DD)+' ; '+MsgLibel.Mess[16]+FloatToStr(CD),1,'G'));
       end;

   // Piece Devise avec Montants devise = 0
   4 : begin // GP A REVOIR
         if ((Q.FindField('E_DEBITDEV').AsFloat=0) and (Q.FindField('E_CREDITDEV').AsFloat=0)) then
         begin
           if Q.FindField('E_NATUREPIECE').AsString <> 'ECC' then
           begin
             OkD := LeSolde(DD,CD,FALSE,FALSE,DecE,DD1,CD1) ; //MontantEuro:=LeSolde(DE,CE) ;
             if OkD then
             begin
               lTobEcr.PutValue('E_DEBITDEV' , DD1);
               lTobEcr.PutValue('E_CREDITDEV', CD1);
             end;

             if OkD then
               LaListeAdd(GoListe1(Mvt,'C',Mvt.DEVISE+' '+MsgLibel.Mess[15]+FloatToStr(DD)+' ; '+MsgLibel.Mess[16]+FloatToStr(CD),19,'G'));
           end
           else
             Err:=True ;
         end;
       end;
  end ;

  Ferme(Q) ;

  FInfoEcr.Load( Cpt, '', Mvt.JOURNAL);
  lRecError := CIsValidLigneSaisie( lTobEcr, FInfoEcr);

  if lRecError.RC_Error <> RC_PASERREUR then
  begin
  if lRecError.RC_Message <> '' then
   PGIInfo(lRecError.RC_Message)
    else
     FMessage.Execute( lRecError.RC_Error );
  end
  else
    lTobEcr.InsertDb( nil, False);

  FreeAndNil( lTobEcr );
End ;

////////////////////////////////////////////////////////////////////////////////

Procedure TFReparMvt.CorrEquilAna(E : Byte ; M : TabTot ; OnDeb : Boolean ; DecE : Integer ; SensInverse : Boolean) ;
{ Pour Info ! }
{ ToT[1] --> Mvt Gene , Tot[3] --> Anal En Pivot , Tot}
{ ToT[2] --> Mvt Gene , Tot[4] --> Anal En Devise }
{ ToT[5] --> Mvt Gene , Tot[6] --> Anal En Euro }

Var T,Q : TQuery ; St : String ;
    NumVent,Signe : Integer ;
    SolEP, SolED, SolAP, SolAD,TauxDev,D1,C1,DD1,CD1 : Double ;
    OkP,OkD : Boolean ;
    lIntFirstAxe, lIntI : integer;
BEGIN
NumVent:=1 ; TauxDev:=1 ; SensInverse:=SensInverse And (E in [1,2,3]) ; lIntFirstAxe := 0;
St:=' Select Y_NUMVENTIL,Y_TAUXDEV,Y_DEBIT,Y_CREDIT,Y_DEBITDEV,Y_CREDITDEV From ANALYTIQ';
St:=St+' where Y_JOURNAL="'+Mvt.Journal+'" and Y_DATECOMPTABLE="'+USDATETIME(Mvt.DATECOMPTABLE)+'" ' ;
St:=St+' and Y_EXERCICE="'+Mvt.Exercice+'" and Y_NUMEROPIECE='+IntToStr(Mvt.NUMEROPIECE)+' ' ;
St:=St+' and Y_NUMLIGNE='+IntToStr(Mvt.NUMLIGNE)+' and Y_DEVISE="'+Mvt.DEVISE+'" ' ;
St:=St+' and Y_ETABLISSEMENT="'+Mvt.ETABLISSEMENT+'" and Y_AXE="'+Mvt.Axe+'" AND Y_QUALIFPIECE="'+Mvt.QUALIFPIECE+'" ' ;
St:=St+' order by Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL desc ' ;
Q:=OpenSql(St, Not SensInverse) ;
If Not Q.EOF Then
   BEGIN
   NumVent:=Q.FindField('Y_NUMVENTIL').AsInteger+1 ;
   TauxDev:=Q.FindField('Y_TAUXDEV').AsFloat ;
   END ;
If SensInverse Then
   BEGIN
   Q.First ;
   While Not Q.Eof Do
     BEGIN
     Q.Edit ;
     If E in [1,2] Then
        BEGIN
        D1:=Q.FindField('Y_DEBIT').AsFloat ; C1:=Q.FindField('Y_CREDIT').AsFloat ;
        Q.FindField('Y_DEBIT').AsFloat:=C1 ; Q.FindField('Y_CREDIT').AsFloat:=D1 ;
        END ;
     If E in [1,3] Then
        BEGIN
        D1:=Q.FindField('Y_DEBITDEV').AsFloat ; C1:=Q.FindField('Y_CREDITDEV').AsFloat ;
        Q.FindField('Y_DEBITDEV').AsFloat:=C1 ; Q.FindField('Y_CREDITDEV').AsFloat:=D1 ;
        END ;

     Q.Post ;
     Q.Next ;
     END ;
   END ;
Ferme(Q) ;
If SensInverse Then Exit ;
T:=OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_SECTION="Ed#"',False) ;
T.Insert ; InitNew(T) ;
T.FindField('Y_JOURNAL').AsString:=Mvt.JOURNAL ;
T.FindField('Y_EXERCICE').AsString:=Mvt.EXERCICE ;
T.FindField('Y_DATECOMPTABLE').AsDateTime:=Mvt.DATECOMPTABLE;
T.FindField('Y_NUMEROPIECE').AsInteger:=Mvt.NUMEROPIECE ;
T.FindField('Y_NUMLIGNE').AsInteger:=Mvt.NUMLIGNE ;
T.FindField('Y_NUMVENTIL').AsInteger:=NumVent ;
T.FindField('Y_DEVISE').AsString:=Mvt.DEVISE ;
T.FindField('Y_REFINTERNE').AsString:=Mvt.REFINTERNE ;
T.FindField('Y_GENERAL').AsString:=Mvt.GENERAL ;
T.FindField('Y_NATUREPIECE').AsString:=Mvt.NATUREPIECE ;
T.FindField('Y_QUALIFPIECE').AsString:=Mvt.QUALIFPIECE ;
T.FindField('Y_ETABLISSEMENT').AsString:=Mvt.ETABLISSEMENT ;
//SG6 10.03.05 Gestion de mode croisaxe
if not VH^.AnaCroisaxe then
begin
  T.FindField('Y_AXE').AsString:=Mvt.AXE ;
  T.FindField('Y_SECTION').AsString:=VH^.Cpta[AxeToFb(Mvt.Axe)].Attente ;
end
else
begin
  for lIntI := 1 to 5 do
  begin
    if GetParamSoc('SO_VENTILA' + IntToStr(lIntI)) then
    begin
      T.FindField('Y_SOUSPLAN' + IntToStr(lIntI)).AsString :=VH^.Cpta[AxeToFb('A' + IntToStr(lIntI))].Attente ;
      if lIntFirstAxe = 0 then lIntFirstAxe := lIntI;
    end;
  end;
  T.FindField('Y_AXE').AsString:='A' +  IntToStr(lIntFirstAxe);
  T.FindField('Y_SECTION').AsString:=VH^.Cpta[AxeToFb('A' + IntToStr(lIntFirstAxe))].Attente ;


end;



T.FindField('Y_LIBELLE').AsString:=MsgLibel2.Mess[2] ;
T.FindField('Y_ECRANOUVEAU').AsString:=Mvt.ECRANOUVEAU ;
T.FindField('Y_TAUXDEV').AsFloat:=TauxDev ;
T.FindField('Y_POURCENTAGE').AsFloat:=100-M[0].TotDebit ;
If OnDeb Then
   BEGIN
   T.FindField('Y_TOTALECRITURE').AsFloat:=M[1].TotDebit ;
   T.FindField('Y_TOTALDEVISE').AsFloat:=M[2].TotDebit ;
   END Else
   BEGIN
   T.FindField('Y_TOTALECRITURE').AsFloat:=M[1].TotCredit ;
   T.FindField('Y_TOTALDEVISE').AsFloat:=M[2].TotCredit ;
   END ;
{ ToT[1] --> Mvt Gene , Tot[3] --> Anal En Pivot , Tot}
{ ToT[2] --> Mvt Gene , Tot[4] --> Anal En Devise }
{ ToT[5] --> Mvt Gene , Tot[6] --> Anal En Euro }

Signe:=1 ; If Not OnDeb Then Signe:=-1 ;
SolEP:=Arrondi(M[1].TotDebit-M[1].TotCredit,V_PGI.OkDecV)*Signe ;
SolED:=Arrondi(M[2].TotDebit-M[2].TotCredit,DecE)*Signe ;
SolAP:=Arrondi(M[3].TotDebit-M[3].TotCredit,V_PGI.OkDecV)*Signe ;
SolAD:=Arrondi(M[4].TotDebit-M[4].TotCredit,DecE)*Signe ;
Case E of
  1 : BEGIN
      OkP:=LeSolde(SolEP,SolAP,TRUE,OnDeb,V_PGI.OkDecV,D1,C1) ;
      OkD:=LeSolde(SolED,SolAD,TRUE,OnDeb,DecE,DD1,CD1) ; //MontantEuro:=LeSolde(DE,CE) ;
      If OkP Then
         BEGIN
         T.FindField('Y_DEBIT').AsFloat:=D1 ; T.FindField('Y_CREDIT').AsFloat:=C1 ;
         END ;
      If OkD Then
         BEGIN
         T.FindField('Y_DEBITDEV').AsFloat:=DD1 ; T.FindField('Y_CREDITDEV').AsFloat:=CD1 ;
         END ;
      END ;
  2 : BEGIN
      OkP:=LeSolde(SolEP,SolAP,TRUE,OnDeb,V_PGI.OkDecV,D1,C1) ;
      If OkP Then
         BEGIN
         T.FindField('Y_DEBIT').AsFloat:=D1 ; T.FindField('Y_CREDIT').AsFloat:=C1 ;
         END ;
      END ;
  3 : BEGIN
      OkD:=LeSolde(SolED,SolAD,TRUE,OnDeb,DecE,DD1,CD1) ; //MontantEuro:=LeSolde(DE,CE) ;
      If OkD Then
         BEGIN
         T.FindField('Y_DEBITDEV').AsFloat:=DD1 ; T.FindField('Y_CREDITDEV').AsFloat:=CD1 ;
         END ;
      END ;
  4 : BEGIN
      OkP:=LeSolde(SolEP,SolAP,TRUE,OnDeb,V_PGI.OkDecV,D1,C1) ;
      If OkP Then
         BEGIN
         T.FindField('Y_DEBIT').AsFloat:=D1 ; T.FindField('Y_CREDIT').AsFloat:=C1 ;
         END ;
      END ;
   End ;
T.Post ; Ferme(T) ;
END ;

Type tSetAxe = Set Of Byte ;

Procedure RecupAxe(St : String ; Var SetAxe : tSetAxe) ;
Var IAxeEnCours : Byte ;
BEGIN
IAxeEnCours:=StrToInt(Copy(St,2,1)) ; If IAxeEnCours In [1..5] Then SetAxe:=SetAxe+[IAxeEnCours] ;
END ;

Function TFReparMvt.EquilAnal : Boolean ;
Var CBonP, CBonD, {CBonE,} OnDeb   : Boolean ;
    InfAna         : TINFOSMVTANA ;
    Tot            : TabTot ;
    DecDev, I      : Byte ;
    AxeDeTravail   : String ;
    SensInverse : Boolean ;
    AxeOk          : Array[1..5] Of Boolean ;
    IAxeEnCours    : Byte ;
    SetAxe         : tSetAxe ;
    szSql : String;
    //
    lIntI : integer;
    lVarField : variant;
BEGIN
Fillchar(InfAna,SizeOf(InfAna),#0) ; Result:=True ;  // B 10
DecDev:=V_PGI.OkDecV ;
// Rony 15/05/97
If Mvt.DEVISE='' then Exit ;

{ Recherche les Piéce/Lignes de type Analytique }
      { Somme des D & C des Pièces Analytique attachés }
QAnaGene.Close ; // QAnaGene.RequestLive:=TRUE ;
QAnaGene.Params[0].AsString:=Mvt.JOURNAL ;
QAnaGene.Params[1].AsString:=Mvt.EXERCICE ;
QAnaGene.Params[2].AsDateTime:=Mvt.DATECOMPTABLE ;
QAnaGene.Params[3].AsInteger:=Mvt.NUMEROPIECE ;
QAnaGene.Params[4].AsInteger:=Mvt.NUMLIGNE ;
QAnaGene.Params[5].AsString:=Mvt.DEVISE ;
//QAnaGene.Params[6].AsString:=Mvt.Etablissement ;
QAnaGene.Params[6].AsString:=w_w ;
//QAnaGene.Params[7].AsString:=Mvt.AXE ;
SetAxe:=[] ;
QAnaGene.Params[7].AsString:=Mvt.AXEPLUS[1] ; RecupAxe(Mvt.AXEPLUS[1],SetAxe) ;
QAnaGene.Params[8].AsString:=Mvt.AXEPLUS[2] ; RecupAxe(Mvt.AXEPLUS[2],SetAxe) ;
QAnaGene.Params[9].AsString:=Mvt.AXEPLUS[3] ; RecupAxe(Mvt.AXEPLUS[3],SetAxe) ;
QAnaGene.Params[10].AsString:=Mvt.AXEPLUS[4] ; RecupAxe(Mvt.AXEPLUS[4],SetAxe) ;
QAnaGene.Params[11].AsString:=Mvt.AXEPLUS[5] ; RecupAxe(Mvt.AXEPLUS[5],SetAxe) ;
QAnaGene.Params[12].AsString:=Mvt.QUALIFPIECE ;
QAnaGene.Open ;
{ Pour Info ! }
{ Tot[0] --> Total % }
{ ToT[1] --> Mvt Gene , Tot[3] --> Anal En Pivot , Tot}
{ ToT[2] --> Mvt Gene , Tot[4] --> Anal En Devise }
{ ToT[5] --> Mvt Gene , Tot[6] --> Anal En Euro }
While Not QAnaGene.Eof Do { L'écriture a t-elle des analytiques rattachées ? }
   BEGIN
   AxeDeTravail:=QAnaGene.FindField('AXET').AsString ;
   Mvt.Axe:=AxeDeTravail ;
   //SG6 09.03.05 Analytique croisaxe
   if not VH^.AnaCroisaxe then
   begin
     If Length(AxeDeTravail)>=2 Then IAxeEnCours:=StrToInt(Copy(AxeDeTravail,2,1)) Else IAxeEnCours:=0 ;
     If IAxeEnCours In [1..5] Then AxeOk[IAxeEnCours]:=TRUE ;
   end
   else
   begin
     for lIntI := 1 to 5 do
     begin
       lVarField := QAnaGene.FindField('Y_SOUSPLAN' + IntToStr(lIntI)).AsVariant;
       if VarIsNull(lVarField) then continue;
       if QAnaGene.FindField('Y_SOUSPLAN' + IntToStr(lIntI)).AsString <> '' then AxeOk[lIntI] := True;
     end;
   end;
   //Fin SG6

   Fillchar(Tot,SizeOf(Tot),#0) ;
   { Total % }
   Tot[0].TotDebit:=QAnaGene.FindField('PourCent').AsFloat ;
   { Les Montants des Pièces/Lignes Ecr == aux montants Pièces/Lignes analytque ? ... }
   { ...En Devise Pivot }
   Tot[1].TotDebit:=Arrondi(Mvt.DEBIT,V_PGI.OkdecV)  ; Tot[1].TotCredit:=Arrondi(Mvt.CREDIT,V_PGI.OkdecV) ;
   OnDeb:=Tot[1].TotDebit<>0 ;
   Tot[3].TotDebit:=Arrondi(QAnaGene.FindField('SumD').AsFloat,V_PGI.OkdecV) ; Tot[3].TotCredit:=Arrondi(QAnaGene.FindField('SumC').AsFloat,V_PGI.OkdecV) ;
   { ...En Devise }
   if Mvt.DEVISE<>V_PGI.DevisePivot then
      BEGIN
      DecDev:=V_PGI.OkDecV ; TrouveDecDev(Mvt.DEVISE,DecDev) ;
      Tot[2].TotDebit:=Arrondi(Mvt.DEBITDEV,DecDev)  ; Tot[2].TotCredit:=Arrondi(Mvt.CREDITDEV,DecDev) ;
      Tot[4].TotDebit:=Arrondi(QAnaGene.FindField('SumDDev').AsFloat,DecDev) ; Tot[4].TotCredit:=Arrondi(QAnaGene.FindField('SumCDev').AsFloat,DecDev) ;
      END Else
      BEGIN
      Tot[2].TotDebit:=Arrondi(Mvt.DEBITDEV,V_PGI.OkdecV)  ; Tot[2].TotCredit:=Arrondi(Mvt.CREDITDEV,V_PGI.OkdecV) ;
      Tot[4].TotDebit:=Arrondi(QAnaGene.FindField('SumDDev').AsFloat,V_PGI.OkdecV) ; Tot[4].TotCredit:=Arrondi(QAnaGene.FindField('SumCDev').AsFloat,V_PGI.OkdecV) ;
      END ;
   Tot[5].TotDebit:=0  ; Tot[5].TotCredit:=0 ;
   Tot[6].TotDebit:=0 ; Tot[6].TotCredit:=0 ;
   CBonP:=(Tot[1].TotDebit=Tot[3].TotDebit) and (Tot[1].totCredit=Tot[3].totCredit) ;
   CBonD:=(Tot[2].TotDebit=Tot[4].TotDebit) and (Tot[2].totCredit=Tot[4].TotCredit) ;
   //If V_PGI.SAV then CBonE:=(Tot[5].TotDebit=Tot[6].TotDebit) and (Tot[5].totCredit=Tot[6].TotCredit)
   //          Else
   //CBonE:=True ;
   { Dev. Pivot et Dev. : affichage dans la liste des Montants des Pièces/Lignes Ecr <> aux montants Pièces/Lignes analytque ? ... }
   If (not CBonP) and (not CBonD) then // B 14
      BEGIN
      LaListeAdd(GoListe1(Mvt,'C',AxeDeTravail+' : '+MsgLibel.Mess[15]+FloatToStr(Tot[3].TotDebit)+' ('+FloatToStr(Tot[1].TotDebit)+') '
                                  +MsgLibel.Mess[16]+FloatToStr(Tot[3].TotCredit)+' ('+FloatToStr(Tot[1].TotCredit)+') '
                                  +'& '+MsgLibel.Mess[15]+FloatToStr(Tot[4].TotDebit)+' ('+FloatToStr(Tot[2].TotDebit)+') '
                                      +MsgLibel.Mess[16]+FloatToStr(Tot[4].TotCredit)+' ('+FloatToStr(Tot[2].TotCredit)+') '
                                  ,5,'A')) ;
      Result:=False ;
      SensInverse:=(Tot[1].TotDebit=Tot[3].TotCredit) And (Tot[3].TotDebit=Tot[1].TotCredit) And
                   (Tot[2].TotDebit=Tot[4].TotCredit) And (Tot[4].TotDebit=Tot[2].TotCredit) ;
      CorrEquilAna(1, Tot,OnDeb,DecDev,SensInverse) ;
      END Else
   { Dev. Pivot : affichage dans la liste des Montants des Pièces/Lignes Ecr <> aux montants Pièces/Lignes analytque ? ... }
   If not CBonP then                   // B 12
      BEGIN
      LaListeAdd(GoListe1(Mvt,'C',AxeDeTravail+' : '+MsgLibel.Mess[15]+FloatToStr(Tot[3].TotDebit)+' ('+FloatToStr(Tot[1].TotDebit)+') ; '
                           +MsgLibel.Mess[16]+FloatToStr(Tot[3].TotCredit)+' ('+FloatToStr(Tot[1].TotCredit)+')'
                           ,3,'A'));
      Result:=False ;
      SensInverse:=(Tot[1].TotDebit=Tot[3].TotCredit) And (Tot[3].TotDebit=Tot[1].TotCredit) ;
      CorrEquilAna(2, Tot,OnDeb,DecDev,SensInverse) ;
      END Else
   { Devise : affichage dans la liste des Montants des Pièces/Lignes Ecr <> aux montants Pièces/Lignes analytque ? ... }
   If not CBonD then                   // B 13
      BEGIN
      LaListeAdd(GoListe1(Mvt,'C'
                           ,AxeDeTravail+' : '+MsgLibel.Mess[15]+FloatToStr(QAnaGene.FindField('SumDDev').AsFloat)+' ('+FloatToStr(Tot[2].TotDebit)+') ; '
                           +MsgLibel.Mess[16]+FloatToStr(QAnaGene.FindField('SumCDev').AsFloat)+' ('+FloatToStr(Tot[2].TotCredit)+')'
                           ,4,'A'));
      Result:=False ;
      SensInverse:=(Tot[2].TotDebit=Tot[4].TotCredit) And (Tot[4].TotDebit=Tot[2].TotCredit) ;
      CorrEquilAna(3, Tot,OnDeb,DecDev,SensInverse) ;
      END ;
   { Si la devise est renseigné les montants doivent etre <> 0 }
   If (Mvt.DEVISE<>V_PGI.DevisePivot) then
      if (Mvt.NATUREPIECE<>'ECC') then   // B 17
         If (QAnaGene.FindField('SumDDev').AsFloat=0) and (QAnaGene.FindField('SumCDev').AsFloat=0) then
            BEGIN
            LaListeAdd(GoListe1(Mvt,'C',AxeDeTravail+' : '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[15]+FloatToStr(QAnaGene.FindField('SumDDev').AsFloat)
                                 +' ; '+MsgLibel.Mess[16]+FloatToStr(QAnaGene.FindField('SumCDev').AsFloat)
                                 ,19,'A'));
            Result:=False ;
            CorrEquilAna(4, Tot,OnDeb,DecDev,FALSE) ;
            END ;
   { Verifs autres infos d'Ana en rapport avec Ecr}
   PrepareAutres ;
   VerifAutres(InfAna) ;
   if (InfAna.General<>Mvt.General) then {Meme Cpt Gene ?}
      BEGIN                              // B 15
      { Cpt Gene Ecr Anal <> Cpt Gene Ecr , Affichage dans la lise }
      LaListeAdd(GoListe1(Mvt,'C',AxeDeTravail+' : '+'('+InfAna.General+') '+' ('+Mvt.General+')',13,'A')) ;
      Result:=False ;
      (*
      Corrige(QAnaGene,'Y_GENERAL',Mvt.General) ; //MAJMVT('Y_GENERAL',Mvt.General) ;
      *)
      // 10395
      szSql := 'UPDATE ANALYTIQ SET Y_GENERAL="'+ Mvt.General +'" WHERE Y_AXE="'+Mvt.AXE+'" AND Y_DATECOMPTABLE="'+USDateTime(Mvt.DATECOMPTABLE)+'" AND Y_NUMEROPIECE='+IntToStr(Mvt.NUMEROPIECE)+' AND Y_NUMLIGNE='+IntToStr(MVT.NUMLIGNE)+ ' AND Y_EXERCICE="'+Mvt.EXERCICE+'" AND Y_QUALIFPIECE="'+ Mvt.QUALIFPIECE+'" AND Y_JOURNAL="'+Mvt.JOURNAL+'"';
      ExecuteSQL(szSQL);
      END ;

   If InfAna.AnalPur then  // B 16
      BEGIN
      { Les Analytiques ne doivent pas être pur }
      LaListeAdd(GoListe1(Mvt,'C','',17,'A'));
      Result:=False ;
      (*
      Corrige(QAnaGene,'Y_TYPEANALYTIQUE','-') ;
      *)
      END ;
   QAnaGene.Next ;
   END ; (*Else
   BEGIN                     // B 11
   { L'écriture n'a pas d' analytiques rattachées, Affichage dans la lise }
{GP}
   LaListeAdd(GoListe1(Mvt,'C',MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.Credit)
                        ,6,'A'));
   Result:=False ;
   { Total % }
   For i:=1 To 5 Do If Mvt.AxePlus[i]='A'+IntToStr(i) Then
      BEGIN
      Mvt.Axe:=Mvt.AxePlus[i] ;
      Tot[0].TotDebit:=QAnaGene.FindField('PourCent').AsFloat ;
      Fillchar(Tot,SizeOf(Tot),#0) ;
      Tot[1].TotDebit:=Arrondi(Mvt.DEBIT,V_PGI.OkdecV)  ; Tot[1].TotCredit:=Arrondi(Mvt.CREDIT,V_PGI.OkdecV) ;
      OnDeb:=Tot[1].TotDebit<>0 ;
      if Mvt.Devise<>V_PGI.DevisePivot then
         BEGIN
         DecDev:=V_PGI.okDecV ; TrouveDecDev(Mvt.DEVISE,DecDev) ;
         Tot[2].TotDebit:=Arrondi(Mvt.DEBITDEV,DecDev)  ; Tot[2].TotCredit:=Arrondi(Mvt.CREDITDEV,DecDev) ;
         END Else Tot[2].TotDebit:=Arrondi(Mvt.DEBITDEV,V_PGI.OkdecV) ; Tot[2].TotCredit:=Arrondi(Mvt.CREDITDEV,V_PGI.OkdecV) ;
      Tot[3].TotDebit:=0  ; Tot[3].TotCredit:=0 ;
      Tot[4].TotDebit:=0  ; Tot[4].TotCredit:=0 ;
      CorrEquilAna(1, Tot,OnDeb,DecDev,FALSE) ;
      END ;
   END ;
*)
For i:=1 To 5 Do
  BEGIN
  If (i In SetAxe) And (AxeOk[i]=FALSE) Then
    BEGIN
   { L'écriture n'a pas d' analytiques rattachées, Affichage dans la lise }
{GP}
    AxeDeTravail:='A'+IntToStr(i) ;
    LaListeAdd(GoListe1(Mvt,'C',AxeDeTravail+' : '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.Credit)
                         ,6,'A'));
    Result:=False ;
    { Total % }
    Fillchar(Tot,SizeOf(Tot),#0) ;
    Tot[1].TotDebit:=Arrondi(Mvt.DEBIT,V_PGI.OkdecV)  ; Tot[1].TotCredit:=Arrondi(Mvt.CREDIT,V_PGI.OkdecV) ;
    OnDeb:=Tot[1].TotDebit<>0 ;
    if Mvt.Devise<>V_PGI.DevisePivot then
       BEGIN
       DecDev:=V_PGI.okDecV ; TrouveDecDev(Mvt.DEVISE,DecDev) ;
       Tot[2].TotDebit:=Arrondi(Mvt.DEBITDEV,DecDev)  ; Tot[2].TotCredit:=Arrondi(Mvt.CREDITDEV,DecDev) ;
       END Else Tot[2].TotDebit:=Arrondi(Mvt.DEBITDEV,V_PGI.OkdecV) ; Tot[2].TotCredit:=Arrondi(Mvt.CREDITDEV,V_PGI.OkdecV) ;
    Tot[3].TotDebit:=0  ; Tot[3].TotCredit:=0 ;
    Tot[4].TotDebit:=0  ; Tot[4].TotCredit:=0 ;
    Mvt.Axe:=AxeDeTravail ;
    CorrEquilAna(1, Tot,OnDeb,DecDev,FALSE) ;
    {JP 04/07/05 : FQ 13353 : à mon avis, je ne pense pas qu'il est utile de boucler sur tous les axes en mode
                   Croisaxe : je pense que c'est ce qui est à l'origine du duplicate Key}
    if VH^.AnaCroisaxe then Break;
    END ;
  END ;

END ;

Procedure TFReparMvt.PrepareAutres ;
BEGIN
QInfoAnaG.Close ;  QInfoAnaG.Sql.Clear ;
QInfoAnaG.SQL.Add(' Select Y_GENERAL, Y_TYPEANALYTIQUE, Y_TYPEANOUVEAU, Y_ECRANOUVEAU, ');
QInfoAnaG.SQL.Add(' Y_JOURNAL, Y_DATECOMPTABLE, Y_NUMLIGNE, ');
QInfoAnaG.SQL.Add(' Y_EXERCICE,  Y_NUMEROPIECE  ') ;
QInfoAnaG.SQL.Add(' From ANALYTIQ ');
QInfoAnaG.SQL.Add(' Where ');
QInfoAnaG.Sql.Add(' Y_JOURNAL=:JAL ') ;
QInfoAnaG.Sql.Add(' And Y_EXERCICE=:EXO ') ;
QInfoAnaG.Sql.Add(' And Y_DATECOMPTABLE=:DAT ') ;
QInfoAnaG.Sql.Add(' And Y_NUMEROPIECE=:PIE ') ;
QInfoAnaG.Sql.Add(' And Y_NUMLIGNE=:LIG ') ;
QInfoAnaG.Sql.Add(' And Y_DEVISE=:DEV ') ;
QInfoAnaG.Sql.Add(' And (Y_AXE=:AXE1 OR Y_AXE=:AXE2 OR Y_AXE=:AXE3 OR Y_AXE=:AXE4 OR Y_AXE=:AXE5) ') ;
QInfoAnaG.SQL.Add(' And Y_ETABLISSEMENT=:ETA ') ;
QInfoAnaG.SQL.Add(' And Y_QUALIFPIECE=:QUALIF ') ;
QInfoAnaG.SQL.Add(' And Y_QUALIFPIECE<>"C" ') ;
QInfoAnaG.SQL.Add(' order by Y_JOURNAL, Y_EXERCICE,  Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE ') ;
ChangeSql(QInfoAnaG) ;//QInfoAnaG.Prepare ;
PrepareSQLODBC(QInfoAnaG) ;
END ;

Procedure TFReparMvt.VerifAutres(Var IFA : TINFOSMVTANA ) ;
BEGIN
Fillchar(IFA,SizeOf(IFA),#0) ;
QInfoAnaG.Close ;
QInfoAnaG.Params[0].AsString:=Mvt.JOURNAL ;
QInfoAnaG.Params[1].AsString:=Mvt.EXERCICE ;
QInfoAnaG.Params[2].AsDateTime:=Mvt.DATECOMPTABLE ;
QInfoAnaG.Params[3].AsInteger:=Mvt.NUMEROPIECE ;
QInfoAnaG.Params[4].AsInteger:=Mvt.NUMLIGNE ;
QInfoAnaG.Params[5].AsString:=Mvt.DEVISE ;
QInfoAnaG.Params[6].AsString:=Mvt.AXEPLUS[1] ;
QInfoAnaG.Params[7].AsString:=Mvt.AXEPLUS[2] ;
QInfoAnaG.Params[8].AsString:=Mvt.AXEPLUS[3] ;
QInfoAnaG.Params[9].AsString:=Mvt.AXEPLUS[4] ;
QInfoAnaG.Params[10].AsString:=Mvt.AXEPLUS[5] ;
QInfoAnaG.Params[11].AsString:=Mvt.ETABLISSEMENT ;
QInfoAnaG.Params[12].AsString:=Mvt.QUALIFPIECE ;
QInfoAnaG.Open ;
With IFA Do
     BEGIN
     General:=QInfoAnaG.Fields[0].AsString ;
     AnalPur:=(QInfoAnaG.Fields[1].AsString='X') ;
     TypeANO:=QInfoAnaG.Fields[2].AsString ;
     EcrANO:=QInfoAnaG.Fields[3].AsString ;
     END ;
END ;


{procedure TFReparMvt.EnregMvtAna ;  // A cause de PourVérif (Compt. ou Impor) CAD Prefixe de table
BEGIN
Mvt:=TInfoMvt.Create ;
Mvt.JOURNAL:=QAnal.Fields[0].AsString ;
Mvt.DATECOMPTABLE:=QAnal.Fields[1].AsDateTime ;
Mvt.NUMLIGNE:=QAnal.Fields[2].AsInteger ;
Mvt.Axe:=QAnal.Fields[3].AsString ;
Mvt.DEVISE:=QAnal.Fields[4].AsString ;
Mvt.REFINTERNE:=QAnal.Fields[5].AsString ;
Mvt.GENERAL:=QAnal.Fields[6].AsString ;
Mvt.NUMVENTIL:=QAnal.Fields[7].AsInteger ;
Mvt.SECTION:=QAnal.Fields[8].AsString ;
Mvt.ECRANOUVEAU:=QAnal.Fields[9].AsString ;
Mvt.NATUREPIECE:=QAnal.Fields[10].AsString ;
Mvt.ETABLISSEMENT:=QAnal.Fields[11].AsString ;
Mvt.TYPEANALYTIQUE:=QAnal.Fields[12].AsString ;
Mvt.DEBIT:=QAnal.Fields[13].AsFloat ;
Mvt.CREDIT:=QAnal.Fields[14].AsFloat ;
Mvt.DEBITDEV:=QAnal.Fields[15].AsFloat ;
Mvt.CREDITDEV:=QAnal.Fields[16].AsFloat ;
Mvt.NUMEROPIECE:=QAnal.Fields[17].AsInteger ;
Mvt.CONFIDENTIEL:=QAnal.Fields[18].AsString ;
Mvt.EXERCICE:=QAnal.Fields[19].AsString ;
Mvt.TAUXDEV:=QAnal.Findfield('Y_TAUXDEV').AsFloat ;
END ; }

{Function TFReparMvt.VerifCptAnal : Boolean ;
Var OkAx : Boolean ;
BEGIN
Result:=True ; OkAx:=True ;
If Mvt.Axe='' then       // D 11
   BEGIN
   //if MAJ then Corrige()
   LaListeAdd(GoListe1(Mvt,'C','',60,'Y'));
   Result:=False ; Exit ;
   END Else if (Char(Mvt.Axe[2])<>'1') and (Char(Mvt.Axe[2])<>'2') and (Char(Mvt.Axe[2])<>'3')
               And (Char(Mvt.Axe[2])<>'4') and (Char(Mvt.Axe[2])<>'5')then
               BEGIN
               LaListeAdd(GoListe1(Mvt,'C','('+Mvt.Axe+')',61,'Y'));
               Result:=False ; Exit ;
               END ;
If Mvt.Section='' then    // D 12
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C','',59,'Y'));
   Result:=False ; Exit ;
   END ;
// Y a t-il un code section existant sur un axe ?
if Not ExisteSec(Mvt.Axe, Mvt.Section) then
   BEGIN
   // Section n'existe pas, Affichage dans la lise
   LaListeAdd(GoListe1(Mvt,'C','',11,'Y'));
   Result:=False ;
   END ;

If Mvt.TYPEANALYTIQUE='X' then
   If Mvt.GENERAL='' then Exit ; //En Anal Pur le gene peut ne pas etre renseigné
If Mvt.General='' then          // D 13
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C',' "'+Mvt.General+'" ',58,'Y'));
   Result:=False ; Exit ;
   END ;
Fillchar(CGen,SizeOf(CGen),#0) ;
InfosCptGen(Mvt.General) ;
if Not Cgen.Ventilable[StrToInt(Char(Mvt.Axe[2]))] then OkAx:=False ;
   // Verif Gene ventil
If Not OkAx then
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.GENERAL+') ('+Mvt.Axe+')',28,'Y'));
   Result:=False ;
   END ;
END ;}

{Function TFReparMvt.VerifInfoJalAnal : Boolean ;
BEGIN
Result:=True ; Fillchar(InfAutres,SizeOf(InfAutres),#0) ;
if Mvt.Journal='' then                            // D 21
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C',' "'+Mvt.JOURNAL+'" ',47,'Y'));
   Result:=False ; Exit ;
   END Else
   BEGIN
   If Not Existe(2,Mvt.JOURNAL) then // Verif si Journal ok           // D 22
      BEGIN
      LaListeAdd(GoListe1(Mvt,'C',' "'+Mvt.JOURNAL+'" ',36,'Y'));
      Result:=False ; Exit ;
      END ;
   END ;

If Mvt.TYPEANALYTIQUE='X' then // Journal D' OD Analytique
   BEGIN
   if (InfAutres.J_NATUREJAL<>'ODA') and (InfAutres.J_NATUREJAL<>'ANA') then    // D 23
      BEGIN
      // Journal doit être de nature ODA et ODA
      LaListeAdd(GoListe1(Mvt,'C','',26,'Y'));
      Result:=False ;
      END Else
      BEGIN
      If InfAutres.J_Axe<>Mvt.AXE then                  // D 24
         BEGIN
         // Journal n'ayant pas le meme Axe ou inexistant
         LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.AXE+') ',27,'Y'));
         Result:=False ;
         END ;
      END ;
   END Else         // Analytique Comptable
   BEGIN
   if (InfAutres.J_NATUREJAL='ODA') or (InfAutres.J_NATUREJAL='ANA') then                  // D 25
      BEGIN
      // Journal est de nature ODA ou ANA
      LaListeAdd(GoListe1(Mvt,'C','',40,'Y'));
      Result:=False ;
      END ;
   END ;
END ;}

{Function TFReparMvt.VerifInfoPieceAnal : Boolean ;
BEGIN
Result:=True ;                        // D 40
If Mvt.TYPEANALYTIQUE='X' then // Anal Pur
   BEGIN
   If Mvt.NUMLIGNE<>0 then
      BEGIN
      // Numero de ligne ne doit pas etre >0, Affichage dans la lise
      LaListeAdd(GoListe2(Mvt,'C','',QAnal.FindField('Y_NUMLIGNE'),'Y'));
      Result:=False ;
      END ;
   END ;
IF Mvt.NUMVENTIL<1 then
   BEGIN
   LaListeAdd(GoListe2(Mvt,'C','',QAnal.FindField('Y_NUMVENTIL'),'Y'));
   Result:=False ;
   END ;
if Mvt.NATUREPIECE<> '' then
   BEGIN
   If Not Existe(3,Mvt.NATUREPIECE) then
      BEGIN
      LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.NATUREPIECE+') ',51,'Y'));
      Result:=False ;
      END Else
      BEGIN   (*Prevoir aussi Proc. NaturePieceJal si Ok*)
      if Not NATUREPIECECOMPTEOK(Mvt.NATUREPIECE,Mvt.GENERAL) then
         BEGIN
         LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.NATUREPIECE+') '+' ('+Mvt.GENERAL+') ',32,'Y'));
         Result:=False ;
         END ;
      If Mvt.NATUREPIECE='ECC' then // Verif Montant en Ecart de change
         BEGIN
         if (Mvt.DEBIT=0) and (Mvt.CREDIT=0) then
            BEGIN
            LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.NATUREPIECE+') ',43,'Y'));
            Result:=False ;
            END ;
         if (Mvt.DEBITDEV<>0) and (Mvt.CREDITDEV<>0) then
            BEGIN
            LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.NATUREPIECE+') ',44,'Y'));
            Result:=False ;
            END ;
         If Mvt.Devise=V_PGI.DevisePivot then
            BEGIN
            LaListe.Add(GoListe1(Mvt,'C',' ('+Mvt.NATUREPIECE+') '+' ('+Mvt.DEVISE+') ',63,'Y'));
            Result:=False ;
            END ;
         END ;
      END ;
   END Else
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C',' "'+Mvt.NATUREPIECE+'" ',50,'Y'));
   Result:=False ;
   END ;
If Mvt.NUMEROPIECE<=0 then
   BEGIN
   LaListeAdd(GoListe2(Mvt,'C','',QAnal.FindField('Y_NUMEROPIECE'),'Y'));
   Result:=False ;
   END ;
END ; }

{Function TFReparMvt.VerifInfoDeviseAnal : Boolean ;
BEGIN
Result:=True ;                           // D 60
if Mvt.DEVISE<>'' then
   BEGIN
   If Not Existe(1,Mvt.DEVISE) then  // Verif si Devise ok
      BEGIN
      LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.EXERCICE+') ',55,'Y'));
      Result:=False ;
      END ;
   END Else
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C',' "'+Mvt.DEVISE+'" ',54,'Y'));
   Result:=False ;
   END ;
END ;}

{Function TFReparMvt.VerifInfoExoAnal : Boolean ;
BEGIN
Result:=True ;
if Not OkExo(Mvt.Exercice, Mvt.DateComptable) then     // D 50
   BEGIN
   LaListeAdd(GoListe2(Mvt,'C','',QAnal.FindField('Y_EXERCICE'),'Y'));
   Result:=False ;
   END ;
END ;}

{Function TFReparMvt.VerifEtabAnal : Boolean ;
BEGIN
Result:=True ;               // xdoc
If Mvt.Etablissement='' then
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C','""',71,'Y'));
   Result:=False ;
   END Else
If Not (TCEtab.values.IndexOf(Mvt.Etablissement)>-1) then
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.Etablissement+') ',72,'Y'));
   Result:=False ;
   END ;
END ;}

Function TFReparMvt.GoListe1(Entitee : TInfoMvt ; Quel, Rem : String ; I : Byte ; CodeIMP : String ) : DelInfo ;
Var X : DelInfo ;
    StAxe : String ;
BEGIN
Inc(NbError) ; NbErrD.Value:=NbError ; NbErrC.Value:=NbError ; X:=DelInfo.Create ;
If Quel='C' then
   BEGIN
   X.LeCod:=Entitee.Journal ; X.LeLib:=Entitee.Refinterne ;
   X.LeMess:=IntToStr(Entitee.Numeropiece)+'/'+IntToStr(Entitee.Numligne) ;
   X.LeMess2:=DateToStr(Entitee.datecomptable) ;
   if i<255 then X.LeMess3:=MsgLibel.Mess[i]+' '+Rem
            else X.LeMess3:=Rem ;
   X.LeMess4:=Entitee.exercice ;
   if (CodeIMP='A') then
      BEGIN
      StAxe:='' ;
      For i:=1 to 5 do If Entitee.AXEPLUS[i]<>'A0' Then StAxe:=StAxe+Entitee.AXEPLUS[i]+';' Else StAxe:=StAxe+'A1'+';' ;
      X.LeMess4:=X.LeMess4+';'+StAxe ;
      END Else if (CodeIMP='Y') then X.LeMess4:=X.LeMess4+';'+Entitee.Axe ;
  // if (CodeIMP='A') or (CodeIMP='Y') then X.LeMess4:=Entitee.exercice+';'+Entitee.Axe else X.LeMess4:=Entitee.exercice ;
   END ;
Result:=X ;
END ;

Function TFReparMvt.GoListe2(Entitee : TInfoMvt ; Quel, Rem : String ; F : TField ; CodeIMP : String ) : DelInfo ;
Var X : DelInfo ; 
BEGIN
Inc(NbError) ; NbErrD.Value:=NbError ; NbErrC.Value:=NbError ;X:=DelInfo.Create ;
If Quel='C' then
   BEGIN
   X.LeCod:=Entitee.Journal ; X.LeLib:=Entitee.Refinterne ;
   X.LeMess:=IntToStr(Entitee.Numeropiece)+'/'+IntToStr(Entitee.Numligne) ;
   X.LeMess2:=DateToStr(Entitee.datecomptable) ; X.LeMess3:=MsgLibel2.Mess[0]+' "'+F.FieldName+'" '+MsgLibel2.Mess[1]+' '+Rem ;
   if (CodeIMP='A') or (CodeIMP='Y') then X.LeMess4:=Entitee.exercice+';'+Entitee.Axe else X.LeMess4:=Entitee.exercice ;
   END Else
If Quel='L' then
   BEGIN
   X.LeCod:=Entitee.lettrage ; X.LeLib:=Entitee.auxiliaire ;
   X.LeMess:=Entitee.general ;
   X.LeMess2:=Entitee.ETATLETTRAGE ; X.LeMess3:=MsgLibel2.Mess[0]+' "'+F.FieldName+'" '+MsgLibel2.Mess[1]+' '+Rem ;
   X.LeMess4:=Quel ;
   END ;
Result:=X ;
END ;

{Function TFReparMvt.EcrAnoOk(Journal,Ecran : String) : Boolean ;
BEGIN
Result:=True ;
if (Journal<>'ANO') and (Journal<>'CLO') then
   BEGIN
   If Ecran<>'N' then Result:=False ;
   END Else
if Journal='ANO' then
   BEGIN
   if (Ecran<>'H') and (ECRAN<>'OAN') then Result:=False ;
   END Else
if Journal='CLO' then
   BEGIN
   if Ecran<>'C' then Result:=False ;
   END ;
END;}

{Function TFReparMvt.OkExo(E : String3 ; D : TDateTime) : Boolean ;
Var i : Integer ; Inf : TEnregInfos ; Ok : Boolean ;
BEGIN
Ok:=False ;
For i:=0 to ListeEXO.Count-1 do
    BEGIN
    if ListeEXO.Objects[i]<>Nil then
       BEGIN
       Inf:=TEnregInfos(ListeEXO.Objects[i]) ;
       IF (Inf.Inf1=E) then
          BEGIN
          if (D>=Inf.D1) then
             If (D<=Inf.D2) then
                BEGIN
                InfAutres.EX_DATEEXO[0]:=Inf.D1 ;
                InfAutres.EX_DATEEXO[1]:=Inf.D2 ; Ok:=True ;
                END ;
          Break ;
          END ;
       END Else Break ;
    END ;
Result:=Ok ;
END ;}

function TFReparMvt.Existe(V : Byte ; St : String ) : Boolean ;
Var i : Integer ; Inf : TEnregInfos ; trouve : Boolean ;
BEGIN
Trouve:=False ;
Case V of
  1 : BEGIN   // Devise
      For i:=0 to ListeDEV.Count-1 do
          BEGIN
          if ListeDEV.Objects[i]<>Nil then
             BEGIN
             Inf:=TEnregInfos(ListeDEV.Objects[i]) ;
             Trouve:=(Inf.Inf1=St) ;
             IF Trouve then Begin Break ; End ;
             END Else begin Break ; end ;
          END ;
      END ;
  2 : BEGIN  // Journal
      For i:=0 to ListeJAL.Count-1 do
          BEGIN
          if ListeJAL.Objects[i]<>Nil then
             BEGIN
             Inf:=TEnregInfos(ListeJAL.Objects[i]) ;
             Trouve:=(Inf.Inf1=St) ;
             IF Trouve then
                BEGIN
                InfAutres.J_NATUREJAL:=Inf.Inf2 ;
                InfAutres.J_Axe:=Inf.Inf3 ;
                InfAutres.J_COMPTEINTERDIT:=Inf.Inf4 ;
                InfAutres.J_CONTREPARTIE:=Inf.Inf5 ;
                Break ;
                END ;
             END Else begin Break ; end ;
          END ;
      END ;
  3 : BEGIN // Nature de Piece
      Trouve:=(TCNatPiece.Items[TCNatPiece.Items.IndexOf(St)]=St) ;
      END ;
 end ;
Result:=Trouve ;
END ;

procedure TFReparMvt.TrouveDecDev(St : String3 ; Var Dec : Byte) ;
Var i : Integer ; Inf : TEnregInfos ;
BEGIN
For i:=0 to ListeDEV.Count-1 do
    BEGIN
    if ListeDEV.Objects[i]<>Nil then
       BEGIN
       Inf:=TEnregInfos(ListeDEV.Objects[i]) ;
       if (Inf.Inf1=St) then Begin Dec:=Inf.Nb ; Break ; End ;
       END Else Break ;
    END ;
END ;

{function TFReparMvt.ExisteSec(Ax, Cod : String ) : Boolean ;
Var i : Integer ; Inf : TEnregInfos ; Trouve : Boolean ;
BEGIN
Trouve:=False ;
For i:=0 to ListeSec.Count-1 do
    BEGIN
    if ListeSec.Objects[i]<>Nil then
       BEGIN
       Inf:=TEnregInfos(ListeSec.Objects[i]) ;
       Trouve:=((Inf.Inf1=Ax)and(Inf.Inf2=Cod)) ;
       IF Trouve then Begin Break ; End ;
       END Else begin Break ; end ;
    END ;
Result:=Trouve ;
END ;}

procedure TFReparMvt.InfosCptGen(Ge : String) ;
Var i : Integer ; Inf : TEnregInfos ;
BEGIN
For i:=0 to ListeGen.Count-1 do
    BEGIN
    if ListeGen.Objects[i]<>Nil then
       BEGIN
       Inf:=TEnregInfos(ListeGen.Objects[i]) ;
       IF (Inf.Inf1=Ge) then
          BEGIN
          CGen.General:=Inf.Inf1 ;
          CGen.NatureGene:=Inf.Inf2 ;
          CGen.Collectif:=Inf.Bol1 ;
          CGen.Lettrable:=Inf.Bol2 ;
          CGen.Pointable:=Inf.Bol3 ;
          CGen.Ventilable:=Inf.SuperBol ;
          Break ;
          END ;
       END Else Break ;
    END ;
END ;

procedure TFReparMvt.InfosCptAux(Au : String) ;
Var i : Integer ; Inf : TEnregInfos ; CAux : TGTiers ; OkOk : Boolean ;
BEGIN
Caux:=NIL ; OkOk:=FALSe ;
For i:=0 to ListeAux.Count-1 do
    BEGIN
    if ListeAux.Objects[i]<>Nil then
       BEGIN
       Inf:=TEnregInfos(ListeAux.Objects[i]) ;
       IF (Inf.Inf1=Au) then
          BEGIN
          CTiers.Auxi:=Inf.Inf1 ;
          CTiers.NatureAux:=Inf.Inf2 ;
          CTiers.Lettrable:=Inf.Bol1 ; OkOk:=TRUE ;
          Break ;
          END ;
       END Else begin Break ; end ;
    END ;
If Not OkOk then
   BEGIN
   If CAux=Nil then CAux:=TGTiers.Create(Au) ;
   if CAux<>Nil then
      BEGIN
      CTiers.Auxi:=CAux.Auxi ;
      CTiers.NatureAux:=CAux.NatureAux ;
      CTiers.Lettrable:=CAux.Lettrable ;
      END ;
   If CAux<>NIL Then BEGIN CAux.FREE ; END ;
   END ;
END ;
(*
procedure TFReparMvt.SQLEcrl ;
Var SiLet, SiPoint, SiVent, MEcr : Boolean ;
BEGIN
SiLet:=False ; SiPoint:=False ; SiVent:=False ; MEcr:=True ; QEcr.First ;
While not QEcr.Eof do
      BEGIN
      If Mvt<>NIL Then BEGIN Mvt.FREE ; Mvt:=NIL ; END ;
      MoveCur(False) ;
      EnregMvtEcr ;
      If VerifCompte(SiLet, SiPoint, SiVent) Then
         BEGIN
         if VerifInfoJal then
            BEGIN
            if VerifInfoExo then
               BEGIN
               IF VerifInfoEcrAno then
                  BEGIN
                  If Mvt.ECRANOUVEAU<>'H' then
                     BEGIN
                     if Not VerifInfoEche(SiLet,SiPoint) then OkVerif:=False ;
                     If Not VerifInfoDevise then OkVerif:=False ;
                     if Not VerifInfoPieces then OkVerif:=False ;
                     If Not VerifTauxDev(MEcr) then OkVerif:=False ;
                     END ;
                  if Not VerifInfoAnal(SiVent) then OkVerif:=False ;
                  //if not VerifInfoConf then OkVerif:=False ;
                  if not VerifEtab then OkVerif:=False ;
                  END Else OkVerif:=False ;
               END else OkVerif:=False ;
            END else OkVerif:=False ;
         END Else OkVerif:=False ;
      if TestBreak then Break ;
      QEcr.Next ;
      END ;
END ;
*)

procedure TFReparMvt.SQLEcr ;
Var SiLet, SiPoint, SiVent, {MEcr,} DeBanque : Boolean ;
BEGIN
SiLet:=False ; SiPoint:=False ; SiVent:=False ; {MEcr:=True ;} QEcr.First ;
While not QEcr.Eof do
      BEGIN
      If Mvt<>NIL Then BEGIN Mvt.FREE ; Mvt:=NIL ; END ;
      MoveCur(False) ;
      EnregMvtEcr ;
      If VerifCompte(SiLet, SiPoint, SiVent, DeBanque) Then
         BEGIN
         If Not VerifInfoAnal(SiVent) then OkVerif:=False ;
         if Not VerifInfoEche(SiLet,SiPoint) then OkVerif:=False ;
         (*
         if VerifInfoJal then
            BEGIN
            if VerifInfoExo then
               BEGIN
               IF VerifInfoEcrAno then
                  BEGIN
                  If (Mvt.ECRANOUVEAU<>'H') And (Mvt.ECRANOUVEAU<>'C') then
                     BEGIN
                     if Not VerifInfoEche(SiLet,SiPoint) then OkVerif:=False ;
                     If Not VerifInfoDevise then OkVerif:=False ;
                     if Not VerifInfoPieces then OkVerif:=False ;
                     If Not VerifTauxDev(MEcr) then OkVerif:=False ;
                     END ;
                  if Not VerifInfoAnal(SiVent) then OkVerif:=False ;
                  //if not VerifInfoConf then OkVerif:=False ;
                  if not VerifEtab then OkVerif:=False ;
                  END Else OkVerif:=False ;
               END else OkVerif:=False ;
            END else OkVerif:=False ;
         *)
         END Else OkVerif:=False ;
      if TestBreak then Break ;
      QEcr.Next ;
      END ;
END ;

Procedure TFReparMvt.EnregMvtEcr ;
BEGIN               {Pour Verif aussi avant imports}
Mvt:=TInfoMvt.Create ;
Mvt.JOURNAL:=QEcr.Fields[0].AsString ;
Mvt.DATECOMPTABLE:=QEcr.Fields[1].AsDateTime ;
Mvt.NUMLIGNE:=QEcr.Fields[2].AsInteger ;
Mvt.GENERAL:=QEcr.Fields[3].AsString ;
Mvt.AUXILIAIRE:=QEcr.Fields[4].AsString ;
Mvt.DEVISE:=QEcr.Fields[5].AsString ;
Mvt.MODEPAIE:=QEcr.Fields[6].AsString ;
Mvt.DEBIT:=QEcr.Fields[7].AsFloat ;
Mvt.CREDIT:=QEcr.Fields[8].AsFloat ;
Mvt.DEBITDEV:=QEcr.Fields[9].AsFloat ;
Mvt.CREDITDEV:=QEcr.Fields[10].AsFloat ;
Mvt.ANA:=QEcr.Fields[11].AsString ;
Mvt.REFINTERNE:=QEcr.Fields[12].AsString ;
Mvt.DATEECHEANCE:=QEcr.Fields[13].AsDateTime ;
Mvt.NUMECHE:=QEcr.Fields[14].AsInteger ;
Mvt.ETABLISSEMENT:=QEcr.Fields[15].AsString ;
Mvt.ETATLETTRAGE:=QEcr.Fields[16].AsString ;
Mvt.ECHE:=QEcr.Fields[17].AsString ;
Mvt.ECRANOUVEAU:=QEcr.Fields[18].AsString ;
Mvt.COUVERTURE:=QEcr.Fields[19].AsFloat ;
Mvt.COUVERTUREDEV:=QEcr.Fields[20].AsFloat ;
Mvt.NATUREPIECE:=QEcr.Fields[21].AsString ;
Mvt.NUMEROPIECE:=QEcr.Fields[22].AsInteger ;
Mvt.CONFIDENTIEL:=QEcr.Fields[23].AsString ;
Mvt.EXERCICE:=QEcr.Fields[24].AsString ;
Mvt.TAUXDEV:=QEcr.Findfield('E_TAUXDEV').AsFloat ;
END ;

Function TFReparMvt.VerifCompte(Var let, Poin, vent, DeBanque : Boolean ) : Boolean ; // C 10
{ Verif si le Cpt existe ou s'il doit etre renseigné }
BEGIN
Result:=True ; Let:=False ; Poin:=False ; Vent:=False ; DeBanque:=FALSE ;
{ Verif si Cpt Géné et Tiers est renseigné ou existe }
If Mvt.General='' then   // C 11
   BEGIN
   LaListeAdd(GoListe1(Mvt,'C',' "'+Mvt.General+'" ',57,'E'));
   VerifCompte:=False ;
   Corrige(QEcr,'E_GENERAL',GeneAttend.Text){ A Supp, bien sur ...} ;
   Mvt.general:=GeneAttend.Text ;
   END ;
Fillchar(CGen,SizeOf(CGen),#0) ;
InfosCptGen(Mvt.General) ;
Fillchar(CTiers,SizeOf(CTiers),#0) ;
InfosCptAux(Mvt.Auxiliaire) ;
If Cgen.General<>'' then
   BEGIN
   If CGen.Collectif then  { Si Gene est collectif }
      BEGIN
      If Mvt.AUXILIAIRE<>'' then { Si Tiers Renseigné }
         BEGIN
         (*
         if CTiers.Auxi='' then  { Si le Tiers n'existe pas } // C 16
            BEGIN     {  }
            LaListeAdd(GoListe1(Mvt,'C','('+Mvt.AUXILIAIRE+') ',10,'E'));
            Result:=False ;
            END ;
         *)
         Let:=CTiers.Lettrable ; {Modif : 24/12}
         END Else                             { Si Tiers Non Renseigné }
         BEGIN
         { Sinon, Le Tiers doit etre renseigné }  // C 14
{GP}     (*
         Corrige(QEcr,'E_GENERAL',GeneAttend.Text){ A Supp, bien sur ...} ;
         LaListeAdd(GoListe1(Mvt,'C','',7,'E'));
         Result:=False ;
         *)
         END ;
      END Else                                  { Si Gene est Non - collectif }
      BEGIN
      { Peut-etre parce que Le Géné est TID ou TIC }
      If (CGen.NatureGene='TID') or (CGen.NatureGene='TIC') or
         {JP 17/10/05 : FQ 16871 : maintenant, les comptes divers peuvent être lettrables}
         (CGen.NatureGene='DIV') then
         BEGIN
         Let:=CGen.Lettrable; {Modif : 24/12}
         END ;
      if Mvt.AUXILIAIRE<>'' then
         BEGIN        { Le Tiers doit etre vide }  // C 15
         (*
         LaListeAdd(GoListe1(Mvt,'C','',8,'E'));
         Result:=False ;
         *)
         END ;
      Poin:=CGen.Pointable ;  {Modif : 24/12}
      END ;
   (*
   If CGen.NatureGene='BQE' then   // C 13
      if DevBqe(CGen.General)<>Mvt.DEVISE then
         BEGIN
         LaListeAdd(GoListe1(Mvt,'C','('+CGen.General+') '+'('+Mvt.DEVISE+') ',33,'E'));
         Result:=False ;
         END ;
   *)
   vent:=CGen.Ventilable[1] or CGen.Ventilable[2] or CGen.Ventilable[3] or CGen.Ventilable[4] or CGen.Ventilable[5];
   DeBanque:=Cgen.NatureGene='BQE' ;
   END Else               // C 12
   BEGIN
   (*
   LaListeAdd(GoListe1(Mvt,'C','('+Mvt.GENERAL+') ',9,'E'));
   Result:=False ;
   *)
   END ;
END ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 02/12/2004
Description .. : Ré-écriture complete de la fonction en utilisant le tableau
Suite ........ : Excel crée par Olivier Lambert
Suite ........ : GCO - 02/12/2004 - FQ 15004
Mots clefs ... :
*****************************************************************}
function TFReparMvt.VerifInfoEche(Lettrable, pointable : Boolean) : Boolean ;
var OkPourLet, OkPourPoint : Boolean ;
begin
  Result      := True;
  OkPourLet   := (Mvt.EtatLettrage='AL') or (Mvt.EtatLettrage='PL') or (Mvt.EtatLettrage='TL');
  OkPourPoint := (Mvt.EtatLettrage='RI');

  if (Lettrable) or (Pointable and ((CGen.NatureGene = 'BQE') or (CGen.NatureGene = 'CAI'))) then
  begin
    if (((Mvt.MODEPAIE <> '') and (TCModP.values.IndexOf(Mvt.MODEPAIE)>-1))=FALSE) or (Mvt.ModePaie = '') then
    begin
      if Mvt.MODEPAIE = '' then
      begin
         LaListeAdd(GoListe1(Mvt,'C',' "'+Mvt.MODEPAIE+'" ',35,'E'));
         Result:=False ;
      end
      else
      begin
        LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.MODEPAIE+') ',39,'E'));
        Result:=False ;
      end;
      Corrige(QEcr,'E_MODEPAIE',MpAttend.Value) ;
    end;

    if Mvt.ECHE <> 'X' then
    begin
      LaListeAdd(GoListe2(Mvt,'C','',QEcr.FindField('E_ECHE'),'E'));
      Corrige(QEcr,'E_ECHE','X') ;
      Result := False ;
    end;

    if Mvt.NUMECHE < 1 then
    begin
      LaListeAdd(GoListe2(Mvt,'C','',QEcr.FindField('E_NUMECHE'),'E'));
      CorrigeInt(QEcr,'E_NUMECHE',1) ;
      Result := False ;
    end;

    if (Lettrable) and (not OkPourLet) then
    begin
      LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.EtatLettrage+') ',91,'E'));
      Corrige(QEcr,'E_ETATLETTRAGE','AL') ;
      Result:=False ;
    end;

    if (Pointable) and (not OkPourPoint) then
    begin
      LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.EtatLettrage+') ',91,'E'));
      Corrige(QEcr,'E_ETATLETTRAGE','RI') ;
      Result := False ;
    end;

  end // if (Lettrable) or (Pointable.....
  else
  begin
    // Pas Lettrable, Pas Pointable
    if Mvt.EtatLettrage <> 'RI' Then
    begin
      LaListeAdd(GoListe1(Mvt,'C',' ('+Mvt.EtatLettrage+') ',91,'E'));
      Result:=False ;
      Corrige(QEcr,'E_ETATLETTRAGE','RI') ;
    end;

    if Mvt.NumEche <> 0 then
    begin
      LaListeAdd(GoListe2(Mvt,'C','',QEcr.FindField('E_NUMECHE'),'E'));
      CorrigeInt(QEcr,'E_NUMECHE',0) ;
      Result := False ;
    end;

    if Mvt.Eche <> '-' Then
    begin
      LaListeAdd(GoListe2(Mvt,'C','',QEcr.FindField('E_ECHE'),'E'));
      Corrige(QEcr,'E_ECHE','-') ;
      Result:=False ;
    end;

    if Mvt.ModePaie <> '' Then
    begin
      Corrige(QEcr,'E_MODEPAIE','') ;
    end;

  end; // else

end;
////////////////////////////////////////////////////////////////////////////////

Function TFReparMvt.VerifInfoAnal(Ventilable : Boolean) : Boolean ;
Var Ok : Boolean ;
BEGIN
Ok:=True ;
If Ventilable then    { Gene ventil ? } // C 50
   BEGIN
   If Mvt.ANA='X' then
      BEGIN
      (*
      if Mvt.NUMECHE=0 then Ok:=True Else Ok:=False ;
      *)
      END Else ok:=False;
                            { E_ANA Error }
   If Not Ok then
      BEGIN
      LaListeAdd(GoListe2(Mvt,'C','',QEcr.FindField('E_ANA'),'E'));
      Corrige(QEcr,'E_ANA','X') ;
      END ;
   END Else
   BEGIN
   If Mvt.ANA='-' then
      BEGIN
      (*
      if Mvt.NUMECHE=0 then Ok:=True Else Ok:=False ;
      *)
      END Else ok:=False;
                            { E_ANA Error }
   If Not Ok then
      BEGIN
      LaListeAdd(GoListe2(Mvt,'C','',QEcr.FindField('E_ANA'),'E'));
      Corrige(QEcr,'E_ANA','-') ;
      END ;
   END ;
VerifInfoAnal:=Ok ;
END ;

{function TFReparMvt.VerifTauxDev(MvtEcr : Boolean) : Boolean ;
Var Quot : Double ;
BEGIN
Result:=True ; Exit ;
If Mvt.DEVISE='' then Exit ;

IF Mvt.Devise=V_PGI.DevisePivot then
   BEGIN
   If Mvt.TAUXDEV<>1 then
      BEGIN
      if MvtEcr then LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),77,'E'))
                else LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),79,'Y')) ;
      Result:=False ;
      END ;
   END Else
   BEGIN
   if Mvt.NATUREPIECE='ECC' then exit ;
   if (Mvt.DebitDev=0) and (Mvt.CreditDev=0) then Exit ;
   Quot:=DonneQuotite(Mvt.Devise) ;
   IF Abs(Mvt.Debit)>Abs(Mvt.Credit) then
      BEGIN
      if Arrondi(Mvt.TAUXDEV,4)<>Arrondi((Quot*(Mvt.Debit/Mvt.DebitDev)),4) then
         BEGIN
         If MvtEcr then LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),78,'E'))
                   Else LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),80,'Y'));
         Result:=False ;
         END ;
      END Else
   IF Abs(Mvt.Credit)>Abs(Mvt.Debit) then
      BEGIN
      if Arrondi(Mvt.TAUXDEV,4)<>Arrondi((Quot*(Mvt.Credit/Mvt.CreditDev)),4) then
         BEGIN
         If MvtEcr then LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),78,'E'))
                   Else LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),80,'Y'));
         Result:=False ;
         END ;
      END ;
   END ;
END ;}

{function TFReparMvt.DonneQuotite(Dev : String) : Double ;
Var i : Integer ; Inf : TEnregInfos ;
BEGIN
Result:=0 ;
For i:=0 to ListeDEV.Count-1 do
    BEGIN
    if ListeDEV.Objects[i]<>Nil then
       BEGIN
       Inf:=TEnregInfos(ListeDEV.Objects[i]) ;
       if (Inf.Inf1=Dev) then Begin Result:=Inf.Mont ; Break ; End ;
       END Else Break ;
    END ;
END ;}

Function TFReparMvt.Majuscule(St : String) : Boolean ;
BEGIN
Majuscule:=(St=UpperCase(St)) ;
END ;

procedure TFReparMvt.FDateCpta1KeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFReparMvt.BVentilClick(Sender: TObject);
begin
FicheAxeAna ('') ;
end;

// FQ 10394
procedure TFReparMvt.btnErrorClick(Sender: TObject);
var
  HH : Boolean;
begin
  HH:=TRUE;
  If (LaListe<>NIL) And (LaListe.Count>0) Then
    RapportdErreurMvt(Laliste,3,HH,False);
end;

{JP 06/06/05 : Changement du libellé en fonction du mode de saisie et filtre sur les journaux
{---------------------------------------------------------------------------------------}
procedure TFReparMvt.cbTypeJalChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if (cbTypeJal.Value = 'LIB') or (cbTypeJal.Value = 'BOR') then begin
    TFNumPiece1.Caption := '&N° de folios de';
    {Pas de gestion de qualif pièce en bordereau et libre}
    FTypeEcriture.ItemIndex := 0;
    FTypeEcriture.Enabled := False;
  end
  else begin
    TFNumPiece1.Caption := '&N° de pièces de';
    FTypeEcriture.Enabled := True;
  end;

  if (cbTypeJal.Value = 'LIB') then begin
    FJal1.Plus := 'J_MODESAISIE = "LIB"';
    FJal2.Plus := 'J_MODESAISIE = "LIB"';
  end
  else if (cbTypeJal.Value = 'BOR') then begin
    FJal1.Plus := 'J_MODESAISIE = "BOR"';
    FJal2.Plus := 'J_MODESAISIE = "BOR"';
  end
  else begin
    FJal1.Plus := 'J_MODESAISIE = "-"';
    FJal2.Plus := 'J_MODESAISIE = "-"';
  end;
  FJal1.Text := '';
  FJal2.Text := '';
end;

end.

