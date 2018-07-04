{$A-,H-}
unit IMPCPTA ;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Forms, Dialogs, StdCtrls, Spin, DB, Hctrls, FileCtrl,GACCESS,
     ExtCtrls,Ent1, HStatus, ComCtrls, HEnt1, GDeclaDf, SaisUtil, Buttons, HCompte,
     hmsgbox, Cpteutil, GBilan,ImpCptaU, CpteSav, HSysMenu, SoldeCpt, Hqry, Paramsoc,
     UtilPGI, ADODB, udbxDataset;

{ Point à traiter :
?? Non decrit dans la structure ??  Alimenter la zone Y_ECRANOUVEAU de ANALYTIQ
}

Procedure RecupSCM ;

Type TabTva = Array[1..19] Of String13 ;

Type TTAuxMP = Record
               MPa : String3 ;
               Taux : Double ;
               End ;

Type TEdtLegalV8 = Array[1..3] Of TPerLegalV8 ; // 1 = Jal Div, 2 = GLGen, 3 = GL Aux

Type TabExoLegalV8 = Array[1..2] Of TEdtLegalV8 ; // 1 = En Cours, 2 = Suivant

Type TMP = Record
           Code     : String3 ;
           Abrege   : String17 ;
           Libelle : String70 ;
           APartirde : String3 ;
           PlusJour  : SmallInt ;
           ArrondiJour : String3 ;
           NombreEcheance : SmallInt ;
           SeparePar      : String3 ;
           MontantMin     : Double ;
           RemplaceMin    : String3 ;
           TauxMP         : Array[1..12] Of TTAuxMP ;
           End ;

Type TabMP  = Array[1..50] Of TMP ;

Type TRegV8 =    Record
                 Lig : SmallInt ;
                 TauxAch    : Double ;
                 CptTAch    : String17 ;
                 CptEAch    : String17 ;
                 TauxVte    : Double ;
                 CptTVte    : String17 ;
                 CptEVte    : String17 ;
                 CodeTaux   : String3 ;
                 CodeRegime : String3 ;
                 End ;

Type TRegimeV8 = Array[1..5] Of TRegV8 ;

Type tabRegimeV8 = Array[1..2,1..19] Of TRegimeV8 ;

Type TDevise = Record
               Intit : String ;
               Deci  : Byte ;
               Perte : String ;
               Ben   : String ;
               DeciT : Byte ;
               Iso   : String  ;
               end ;

Type TabDevise = Array[1..20] Of TDevise ;

Type TabRupt = Array[1..99] Of String13 ;

Type TGuideV8 = Record
                Guide : String3 ;
                Libelle : String35 ;
                Abrege  : String35 ;
                NaturePiece : String3 ;
                Journal : String3 ;
                end ;

Type TabGuideV8 = Array[1..99] Of TGuideV8 ;

type
  TFImpCpt = class(TForm)
    QDelete      : TQuery;
    TRIB: THTable;
    TTiers: THTable;
    TGeneraux: THTable;
    TSection: THTable;
    Taxe: THTable;
    TJournal: THTable;
    TBanque: THTable;
    TBanqueCP: THTable;
    TVentil: THTable;
    TECRGEN: THTable;
    TEXO: THTable;
    TECRANAL: THTable;
    TSoc: THTable;
    TSouche: THTable;
    TDEV: THTable;
    TChancell: THTable;
    TTVA: THTable;
    TCommun: THTable;
    TMODR: THTable;
    TChoix: THTable;
    TRupt: THTable;
    TModP: THTable;
    TstrAn: THTable;
    TStrTra: THTable;
    TRef: THTable;
    TRefAuto: THTable;
    TGuides: THTable;
    TEcrGui: THTable;
    TAnaGui: THTable;
    TContAbo: THTable;
    TEDT: THTable;
    TRUB: THTable;
    TEXBQ: THTable;
    TEtab: THTable;
    TRelance: THTable;
    HMsgReprise  : THMsgBox;
    Panel1: TPanel;
    GroupBox3: TGroupBox;
    Label1: TLabel;
    FRep: TEdit;
    BDir: TButton;
    Label2: TLabel;
    Listesoc: TComboBox;
    GroupBox4: TGroupBox;
    FVider: TCheckBox;
    FRapide: TCheckBox;
    GroupBox1: TGroupBox;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    Image4: TImage;
    Image5: TImage;
    Image6: TImage;
    Image7: TImage;
    Image8: TImage;
    Image9: TImage;
    Image10: TImage;
    Image11: TImage;
    Image12: TImage;
    Image13: TImage;
    Image14: TImage;
    Image15: TImage;
    Image16: TImage;
    AImage1: TImage;
    AImage2: TImage;
    AImage3: TImage;
    AImage4: TImage;
    AImage5: TImage;
    AImage6: TImage;
    AImage7: TImage;
    AImage8: TImage;
    AImage9: TImage;
    AImage10: TImage;
    AImage11: TImage;
    AImage12: TImage;
    AImage13: TImage;
    AImage14: TImage;
    AImage15: TImage;
    AImage16: TImage;
    OkGen: TCheckBox;
    OkTiers: TCheckBox;
    OkSect: TCheckBox;
    OkJal: TCheckBox;
    OkVentilanal: TCheckBox;
    OkEcrGen: TCheckBox;
    OkSoc: TCheckBox;
    OkDev: TCheckBox;
    OkChancell: TCheckBox;
    OkTva: TCheckBox;
    OkModR: TCheckBox;
    OkRupt: TCheckBox;
    OkRef: TCheckBox;
    OKOdA: TCheckBox;
    GroupBox2: TGroupBox;
    OkCalcSolde: TCheckBox;
    OkEtatFi: TCheckBox;
    OkGuide: TCheckBox;
    Panel2: TPanel;
    BValider: TBitBtn;
    BAnnule: TBitBtn;
    BAide: TBitBtn;
    BSel: TBitBtn;
    BStop: TBitBtn;
    TMessage: TLabel;
    RV8: TEdit;
    LetHal: TCheckBox;
    ANHal: TCheckBox;
    HMTrad: THSystemMenu;
    AnaHAL: TCheckBox;
    procedure RECUP(Sender: TObject);
    procedure BAnnuleClick(Sender: TObject);
    procedure BDirClick(Sender: TObject);
    procedure ALIMLISTESOC(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OkEcrGenClick(Sender: TObject);
    procedure BSelClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure RV8DblClick(Sender: TObject);
    procedure AnaHALClick(Sender: TObject);
  private { Private-déclarations }
    EnCours,Stopper   : boolean ;
    TbExoLegalV8      : TabExoLegalV8 ;
    TbReg             : TabTva ;
    TbLibReg          : TabTva ;
    TbRelV8           : TabTva ;
    TbRupt            : TabRupt ;
    TbMP              : TabMP ;
    tbEtatFi          : TabEtatFi ;
    TbRegV8           : TabRegimeV8 ;
    P_CptGen,P_CptAux : String17 ;
    P_ExerciceMvt     : String3 ;
    Dc                : Byte ;
    TbDev             : TabDevise ;
    LCodeTva          : TStringList ;
    LCptTva           : TStringList ;
    LesRefJ           : TStringList ;
    procedure VideTable(NomTable,StSup : String) ;
    procedure recupTva ;
    procedure recupgeneraux ;
    procedure RecupPlanRef ;
    procedure recupAuxiliaire ;
    procedure recupSect ;
    procedure recupjal ;
    procedure recupVentilAnal;
    procedure recupVentilType;
    procedure recupEcrGen(OldPos6 : LongInt) ;
    procedure recupAnalGeneAbo(Var Ecr,Fiche6 : EnregCPTA ; Var NoLigne : Integer ;
                               MaxGuide : SmallInt ; GU_TYPE : String3) ;
    procedure recupEcrAboOuTyp(Var MaxGuide : SmallInt ; LeType : Byte) ;
    procedure AjusteLettrage ;
    procedure recupECHE(Var Fiche4,Fiche6 : EnregCPTA ; Var NoLigne : Integer) ;
    procedure recupMonoECHE(Var Fiche4,Fiche6 : EnregCPTA ; Var NoLigne : Integer) ;
    procedure MajEEXBQE(Cpt,Ref : String ; DD : TDateTime ) ;
    procedure recupTroncEcriture(Var Fiche,Fiche6 : EnregCPTA ; Var NoLigne : Integer ; Var TypeMvt : String3) ;
    procedure recupGuide(Var MaxGuide : SmallInt) ;
    procedure recupAnalGene(Var Ecr,Fiche6 : EnregCPTA ; Var NoLigne : Integer ; TypeMvt : String3) ;
    procedure recupODA ;
    procedure RecupEtatFi ;
    procedure recupDevises ;
    procedure recupChancell ;
    procedure recupRuptures ;
    procedure recupSoc ;
    procedure RecupLibAuto ;
    procedure recupModR ;
    procedure AlimRef(Fiche : EnregCpta) ;
    procedure RecupRefAuto ;
    procedure AlimTva ;
    procedure AlimRel ;
    procedure TransfertVente2Achat(Regime,i : SmallInt) ;
    procedure RetoucheFinal ;
    procedure RetoucheRegime ;
    procedure MajCommun ;
    procedure MajTableTVA ;
    procedure MajTableModR ;
    procedure MajTableRelance ;
    procedure MajTableModP ;
    procedure MajRub(TARB : TabRB ; i : Integer) ;
    procedure RecupFamilleRub(i : Integer) ;
    procedure FreeTbEtatFi(i : Integer) ;
    procedure addexo(i : Integer ; Deb,Fin : Word ; LeEtat : String ; ExoEnCours : Boolean) ;
    Procedure RecupExo ;
    Procedure RecupEdtLegal ;
    procedure addSouche(C : Char) ;
    Function  TrouveExo(dd : Word ; Var D2 : TDateTime) : String ;
    procedure LMess (st : String);
    procedure LituneDevise(i:Integer) ;
    procedure LitModRV8 ;
    procedure LitLesDevisesV8 ;
    Function  ISTVAV8(Cpt : String17) : Boolean ;
    procedure LitTvaV8 ;
    Function  QuelTypeMvtV8(Cpt : String17 ; NatCpt : String3 ; NatPiece : Byte) : String3 ;
    procedure LitEtatFiV8 ;
    procedure RecupCodeRupt(St : String3) ;
    procedure RecupCodeRegime ;
    procedure RecupRupt(C : Char) ;
    procedure RecupCodeTranche ;
    procedure RecupTranche ;
    procedure LitRuptV8(C : Char) ;
    procedure RecupLesVentilType ;
    procedure SauveLaPos6(LaPos : LongInt) ;
    Function  LitLaPos6 : LongInt ;
    Function  RepriseEcr : LongInt ;
    procedure LitEdtLegal ;
    procedure MajSO_RECUPCPTA(Stt : String) ;
    function  LitSO_RECUPCPTA : String ;
    procedure InitImage ;
    Function RecupCptJal(PosJal : LongInt ; TypeF : Byte ; Var Jal : EnregCpta) : String200 ;
  public  { Public-déclarations }
  end;

// rem simon
PROCEDURE LITFICHIER ( OkOuvre,Bloc,OkFerme : boolean ) ;

implementation

{$R *.DFM}

{=============================================================================}
procedure TFImpCpt.LitEdtLegal ;
Var Exo,i : SmallInt ;
    LegalV8 : TPerLegalV8 ;
begin
Fillchar(TbexoLegalV8,SizeOf(TbexoLegalV8),#0) ;
OuvreNum ;
For Exo:=VSAA^.CurExercice To VSAA^.CurExercice+1 Do
  BEGIN
  For i:=1 To 3 Do
    BEGIN
    LitLegalPeriodeV8(i-1,LegalV8,Exo,False) ;
    TbExoLegalV8[Exo+1-VSAA^.CurExercice][i]:=LegalV8 ;
    END ;
  END ;
FermeNum ;
end ;

{=============================================================================}
function NatPiece(LaNat : Byte) : String ;
begin
Case LaNat Of
  1 : Result:='FC' ; 2 : Result:='AC' ; 3 : Result:='RC' ;
  4 : Result:='FF' ; 5 : Result:='AF' ; 6 : Result:='RF' ;
  7 : Result:='OD' ;
  Else Result:='OD' ;
  END ;
end ;

{=============================================================================}
Function Bool(b:byte) : String3 ;
begin
If b=1 Then Bool:='X' Else Bool:='-' ;
end ;

{============================================================================}
FUNCTION CODELETTRE ( NumLettre : Longint ; Total : Boolean) : String4 ;

Var St : String4 ;
    Niveau : Array[1..4] of Longint ;
    i : integer ;
BEGIN
St:='    ' ; CodeLettre:='' ;
NumLettre:=((NumLettre-1) mod 456977)+1 ;
If NumLettre<=0 then exit ;
Niveau[4]:=NumLettre-1 ; For i:=3 downto 1 do Niveau[i]:=Niveau[i+1] div 26 ;
For i:=1 to 4 do
    BEGIN
    Niveau[i]:=(Niveau[i] mod 26)+1 ;
    St[i]:=Chr(64+Niveau[i]) ;
    END ;
If Not Total then For i:=1 to Length(St) do St[i]:=Chr(Ord(St[i])+32) ;
CodeLettre:=St ;
END ;

{=============================================================================}
procedure TFImpCpt.VideTable (NomTable,StSup : String) ;
BEGIN
(*
emptytable
*)
if FVider.checked then
   BEGIN
   BeginTrans ;
   LMEss('Suppression en cours de '+NomTable) ;
   QDelete.Close ;
   QDelete.SQL.Clear ;
   QDelete.SQL.Add('DELETE FROM '+NomTable) ;
   If STSup<>'' Then QDelete.SQL.Add(stSup) ;
   ChangeSQL(QDelete) ; QDelete.ExecSQL ;
   CommitTrans ;
   END ;
END ;

{=============================================================================}
PROCEDURE TFImpCpt.LMess (st : String);

BEGIN
If not FRapide.Checked Then Exit ;
TMessage.Caption:=st ;
Application.ProcessMessages ;
END ;

{=============================================================================}
procedure TFImpCpt.LituneDevise(i:Integer) ;
Var Fiche16 : EnregCpta ;
    Pos16 : Longint ;
    Cle16 : String ;
    OkOk : Boolean ;
begin
InitNewSAA(16,Fiche16) ; Fiche16.NoTableTA:=191 ; Fiche16.TypBordTA:=Byte(i-1) ;
Cle16:=CreCle16(4,Fiche16.NoTableTA,Fiche16.CodeTrancheTA,Fiche16.IntituleTA,Fiche16.Compte1TA,Fiche16.TypBordTA) ;
SearchKey(VSAA^.IdxFF[16],Pos16,Cle16) ;
OkOk:=VSAA^.Ok ;
Fillchar(TBDev[i],SizeOf(TBDev[i]),#0) ;
If OkOk Then
   BEGIN
   GetRec(VSAA^.DatFF[16],Pos16,Fiche16) ;
   If (Fiche16.NoTableTA=191) And (Fiche16.TypBordTA=byte(i-1)) Then
      BEGIN
      TBDev[i].Intit :=Trim(Copy(Fiche16.IntituleTA,1,10)) ;
      TBDev[i].Deci  :=Fiche16.DernBorReelTA ;
      TBDev[i].Perte :=Trim(Fiche16.Compte1TA) ;
      TBDev[i].Ben   :=Trim(Fiche16.Compte2TA) ;
      TBDev[i].DeciT :=Fiche16.DernBorAboTA ;
      TBDev[i].Iso   :=Trim(Copy(Fiche16.CodeTrancheTA,11,3)) ;
      END ;
   END ;
end ;

{=============================================================================}
procedure TFImpCpt.LitLesDevisesV8 ;
Var i : Integer ;
begin
Ouvre(16,2,FALSE,spCPTA) ;
TBDev[1].Intit :='FRANCS' ;
TBDev[1].Deci  :=Dc ;
TBDev[1].Iso   :=Trim(VSAA^.Parasoc.Canabis) ;
If TBDev[1].Iso='' Then TBDev[1].Iso:='FRF' ;
For i:=2 To 20 Do LitUneDevise(i) ;
GAccess.Ferme(16,2,FALSE) ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupDevises ;
Var i : Integer ;
begin
VideTable('DEVISE','') ;
BeginTrans ;
TDev.open ;
For i:=1 To 20 Do
  If TbDev[i].ISO<>'' Then
     BEGIN
     TDev.Insert ;
     InitNew(TDev) ;
     TDev.FindField('D_DEVISE').AsString:=TbDev[i].Iso ;
     TDev.FindField('D_LIBELLE').AsString:=TbDev[i].Intit ;
     TDev.FindField('D_FERME').AsString:='-' ;
     TDev.FindField('D_SYMBOLE').AsString:='' ;
     TDev.FindField('D_DECIMALE').AsInteger:=TbDev[i].Deci ;
     TDev.FindField('D_QUOTITE').AsFloat:=1 ;
     TDev.FindField('D_CPTLETTRDEBIT').AsString:=TbDev[i].Perte ;
     TDev.FindField('D_CPTLETTRCREDIT').AsString:=TbDev[i].Ben ;
     TDev.FindField('D_CPTPROVDEBIT').AsString:='' ;
     TDev.FindField('D_CPTPROVCREDIT').AsString:='' ;
     TDev.FindField('D_SOCIETE').AsString:='' ;
     TDev.FindField('D_MAXDEBIT').AsFloat:=0 ;
     TDev.FindField('D_MAXCREDIT').AsFloat:=0 ;
     TDev.Post ;
     END ;
TDev.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.TransfertVente2Achat(Regime,i : SmallInt) ;
Var j : SmallInt ;
begin
for j:=1 To 5 Do
  If (TbRegV8[2,Regime][J].TauxVte=TbRegV8[1,Regime][i].TauxAch) And
     (TbRegV8[2,Regime][J].Lig<>0) And (TbRegV8[2,Regime][J].TauxVte<>0) Then
    BEGIN
    TbRegV8[1,Regime][i].CptTVte:=TbRegV8[2,Regime][j].CptTVte ;
    TbRegV8[1,Regime][i].CptEVte:=TbRegV8[2,Regime][j].CptEVte ;
    TbRegV8[1,Regime][i].TauxVte:=TbRegV8[2,Regime][j].TauxVte ;
    TbRegV8[2,Regime][j].TauxVte:=0 ;
    TbRegV8[2,Regime][j].CptTVte:='' ;
    TbRegV8[2,Regime][j].CptEVte:='' ;
    TbRegV8[2,Regime][j].Lig:=0 ;
    END ;
end ;

{=============================================================================}
procedure TFImpCpt.RetoucheFinal ;
Var IVente,Regime,i : smallInt ;
begin
For iVente:=1 To 2 Do For regime:=1 To 19 Do For i:=1 To 5 Do
  If TbRegV8[Ivente,Regime][i].Lig>0 Then
    BEGIN
    If Abs(Arrondi(TbRegV8[Ivente,Regime][i].TauxAch-18.60,2))<=0.01 Then TbRegV8[Ivente,Regime][i].CodeTaux:='ANC' Else
       If Abs(Arrondi(TbRegV8[Ivente,Regime][i].TauxAch-20.60,2))<=0.01 Then TbRegV8[Ivente,Regime][i].CodeTaux:='NOR' Else
          If Abs(Arrondi(TbRegV8[Ivente,Regime][i].TauxAch-5.50,2))<=0.01 Then TbRegV8[Ivente,Regime][i].CodeTaux:='RED' Else
             If TbRegV8[Ivente,Regime][i].CodeRegime='EXO' Then TbRegV8[Ivente,Regime][i].CodeTaux:='EXO'
    END ;
For iVente:=1 To 2 Do For regime:=1 To 19 Do For i:=1 To 5 Do
  If TbRegV8[Ivente,Regime][i].Lig>0 Then LCodeTVA.Add(TbRegV8[Ivente,Regime][i].CodeTaux);
end ;

{=============================================================================}
procedure TFImpCpt.RetoucheRegime ;
Var Regime,i : SmallInt ;
begin
For Regime:=1 To 19 Do
   For i:=1 To 5 Do
     With TbRegV8[1,Regime][i] Do If (Lig<>0) And (TauxAch<>0) Then TransfertVente2Achat(Regime,i) ;
RetoucheFinal ;
end ;

{=============================================================================}
Function TFImpCpt.ISTVAV8(Cpt : String17) : Boolean ;
Var i : Integer ;
begin
Result:=FALSE ;
For i:=0 To LCptTva.Count-1 Do If Cpt=LCptTva.Strings[i] Then BEGIN Result:=TRUE ; Exit ; END ;
end ;

{=============================================================================}
procedure TFImpCpt.LitTvaV8 ;
Var ivente,Regime : SmallInt ;
    Cle,Test : String40 ;
    NoTable : Byte ;
    StCode : String40 ;
    kk,ll : SmallInt ;
    OkOk : Boolean ;
    Fiche16 : EnregCpta ;
    Pos16 : LongInt ;
    LeTaux : Double ;

begin
AlimTVA ; Fillchar(TbRegV8,SizeOf(TbRegV8),#0) ;
Ouvre(16,2,FALSE,spCPTA) ;
For ivente:=1 To 2 Do
  BEGIN
  For Regime:=1 To 19 Do
    If TbReg[Regime]<>'' Then
       BEGIN
       NoTable:=115 ; StCode:=Chr(Regime)+Chr(ivente) ; kk:=0 ;
       Cle:=CreCle16(1,NoTable,StCode,'','',0) ; Test:=Cle ;
       Repeat
        Inc(kk) ; if kk=1 then SearchKey(VSAA^.Idxff[16],Pos16,Test) else NextKey(VSAA^.Idxff[16],Pos16,Test) ;
        OkOk:=((VSAA^.Ok) and (Copy(Test,1,Length(Cle))=Cle)) ;
        if OkOk then
           BEGIN
           GetRec(VSAA^.Datff[16],Pos16,Fiche16) ;
           LeTaux:=Arrondi(Fiche16.TauxTvaTA,2) ;
           ll:=Fiche16.LongCodeTA ;
           If ((LeTaux<>0) Or (TbReg[Regime]='EXO')) And (ll>0) And (ll<6) Then
              With TbRegV8[IVente,Regime][ll] Do
                BEGIN
                Lig     :=ll ;
                TauxAch :=0  ; CptTAch :='' ; CptEAch :='' ;
                TauxVte :=0  ; CptTVte :='' ; CptEVte :='' ;
                CodeTaux:='00'+InttoStr(ll) ;
                CodeRegime:=TbReg[Regime] ;
                If Ivente=1 Then
                   BEGIN
                   TauxAch :=LeTaux ;
                   CptTAch :=Trim(Fiche16.IntituleTA) ;
                   CptEAch :=Trim(Fiche16.Compte1TA) ;
                   If CptTACh<>'' Then LCptTVA.Add(CptTAch);
                   If CptEACh<>'' Then LCptTVA.Add(CptEAch);
                   END Else
                   BEGIN
                   TauxVte :=LeTaux ;
                   CptTVte :=Trim(Fiche16.IntituleTA) ;
                   CptEVte :=Trim(Fiche16.Compte1TA) ;
                   If CptTVte<>'' Then LCptTVA.Add(CptTVte);
                   If CptEVte<>'' Then LCptTVA.Add(CptEVte);
                   END ;
                END ;
           END ;
       Until ((Not OkOk)) ;
       END ;
  END ;
GAccess.Ferme(16,2,FALSE) ;
RetoucheRegime ;
end ;

{=============================================================================}
procedure TFImpCpt.MajCommun ;
Var i : Integer ;
    St : String ;
begin
VideTable('CHOIXCOD','WHERE CC_TYPE="'+VH^.DefCatTVA+'"') ;
VideTable('COMMUN','WHERE CO_TYPE="'+VH^.DefCatTVA+'"') ;
BeginTrans ;
TChoix.open ; //CHOIX
For i:=0 To LCodeTva.Count-1 Do
  BEGIN
  TChoix.Insert ;
  TChoix.FindField('CC_TYPE').ASString:=VH^.DefCatTVA ;
  TChoix.FindField('CC_CODE').ASString:=LCodeTva.Strings[i] ;
  If LCodeTva.Strings[i]='NOR' Then St:='Taux Normal' Else
     If LCodeTva.Strings[i]='RED' Then St:='Taux Réduit' Else
        If LCodeTva.Strings[i]='ANC' Then St:='Ancien Taux Normal' Else
           If LCodeTva.Strings[i]='EXO' Then St:='Exonére' Else St:='Taux '+LCodeTva.Strings[i] ;
  TChoix.FindField('CC_LIBELLE').ASString:=St ;
  TChoix.FindField('CC_ABREGE').ASString:=St ;
  TChoix.Post ;
  END ;
TChoix.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.MajTableTVA ;
Var Ivente,Regime,i : SmallInt ;
begin
VideTable('TXCPTTVA','WHERE TV_TVAOUTPF="'+VH^.DefCatTVA+'"') ;
BeginTrans ;
TTVA.open ;
For iVente:=1 To 2 Do For regime:=1 To 19 Do For i:=1 To 5 Do
  If (TbRegV8[Ivente,Regime][i].Lig>0) Or (TbReg[Regime]='EXO') Then
     If Not (TTva.FindKey([VH^.DefCatTVA,TbRegV8[Ivente,Regime][i].CodeTaux,TbRegV8[Ivente,Regime][i].CodeRegime]))Then
    BEGIN
    TTVA.Insert ;
    TTVA.FindField('TV_TVAOUTPF').AsString:=VH^.DefCatTVA ;
    TTVA.FindField('TV_CODETAUX').AsString:=TbRegV8[Ivente,Regime][i].CodeTaux ;
    TTVA.FindField('TV_REGIME').AsString:=TbRegV8[Ivente,Regime][i].CodeRegime ;
    TTVA.FindField('TV_TAUXACH').AsFloat:=TbRegV8[Ivente,Regime][i].TauxAch ;
    If TbReg[Regime]='EXO' Then TTVA.FindField('TV_TAUXVTE').AsFloat:=0.00
                           Else TTVA.FindField('TV_TAUXVTE').AsFloat:=TbRegV8[Ivente,Regime][i].TauxVte ;
    TTVA.FindField('TV_CPTEACH').AsString:=TbRegV8[Ivente,Regime][i].CptTAch ;
    TTVA.FindField('TV_CPTEVTE').AsString:=TbRegV8[Ivente,Regime][i].CptTVte ;
    TTVA.FindField('TV_SOCIETE').AsString:='' ;
    TTVA.FindField('TV_ENCAISACH').AsString:=TbRegV8[Ivente,Regime][i].CptEAch ;
    TTVA.FindField('TV_ENCAISVTE').AsString:=TbRegV8[Ivente,Regime][i].CptEVte ;
    TTVA.Post ;
    END ;
TTVA.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupTVA ;
begin
LitTvaV8 ; RecupCodeRegime ; // Régime de TVA
MajCommun ; MajTableTva ; // Paramétrage TVA
end ;

{=============================================================================}
procedure TFImpCpt.RecupChancell ;
Var i,j,k : Integer ;
    yy,mm,jj : Word ;
    Date1 : TDAteTime ;
    PerDev : TPerDev ;
    OkOk : Boolean ;
begin
VideTable('CHANCELL','') ;
BeginTrans ;
TChancell.open ;
OuvreDev ;
For i:=1 to 19 do
   BEGIN
   If TbDev[i+1].ISO<>'' Then
      BEGIN
      For j:=1 to VSAA^.Parasoc.OffsetDev Do
         BEGIN
         Read(VSAA^.FichierDev,PerDev) ;
         For k:=1 To 31 Do
           BEGIN
           If Arrondi(Perdev[k],TbDev[i+1].DeciT)<>0 Then
              BEGIN
              Date1:=Int2Date(VSAA^.ParaSoc.HistoDevise.Deb) ;
              Date1:=PlusMois(Date1,j-1) ;
              DecodeDate(Date1,yy,mm,jj) ;
              OkOk:=(((mm=1) Or (mm=3) Or (mm=5) Or (mm=7) or (mm=8) Or (mm=10) or (mm=12)) And (k<=31)) Or
                    (((mm=4) Or (mm=6) Or (mm=9) Or (mm=11)) And (k<=30)) Or
                    ((mm=2) And (k<=28)) ;
              If OkOk Then
                 BEGIN
                 Date1:=Encodedate(yy,mm,k) ;
                 TChancell.Insert ;
                 TChancell.FindField('H_DEVISE').ASString:=TbDev[i+1].ISO ;
                 TChancell.FindField('H_DATECOURS').ASFloat:=Date1 ;
                 TChancell.FindField('H_TAUXREEL').ASFloat:=Arrondi(PerDev[k],TbDev[i+1].DeciT) ;
                 TChancell.FindField('H_TAUXCLOTURE').ASFloat:=0 ;
                 TChancell.FindField('H_COMMENTAIRE').ASString:='' ;
                 TChancell.FindField('H_SOCIETE').ASString:='' ;
                 TChancell.Post ;
                 END ;
              END ;
           END ;
         END ;
      END ;
   END ;
FermeDev ;
TChancell.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.LitRuptV8(C : Char) ;
Var i : Integer ;
    St13 : String13 ;
    LaPos : LongInt ;
    IMax : Integer ;
begin
IMax:=19 ; LaPos:=0 ;
Case C Of
  'G' : LaPos:=4010 ; 'T' : LaPos:=12610 ; 'S' : LaPos:=12660 ; 'A' : LaPos:=2710 ;
  'V' : BEGIN LaPos:=13660 ; IMax:=99 ; END ;
  End ;
Fillchar(TbRupt,SizeOf(TbRupt),#0) ;
OuvreChoixCod ;
Seek(VSAA^.FichierChCh,LaPos) ; i:=1 ;
repeat
  read(VSAA^.FichierChCh,St13) ;
  If St13<>'' then BEGIN TbRupt[i]:=ASCII2ANSI(St13) ; Inc(i) ; END ;
Until ((St13='') or (i>IMax)) ;
FermeChoixCod ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupCodeRupt(St : String3)  ;
var i : Integer ;
    St1 : String ;
begin
VideTable('CHOIXCOD','WHERE CC_TYPE="'+St+'"') ;
BeginTrans ;
TChoix.open ;
For i:=1 To 19 Do
  BEGIN
  If TbRupt[i]<>'' Then
     BEGIN
     TChoix.Insert ;
     InitNew(TChoix) ;
     TChoix.FindField('CC_TYPE').AsString:=St ;
     TChoix.FindField('CC_CODE').AsString:=FormatFloat('000',i) ;
     If St='VTY' Then St1:='Ventil N° ' Else St1:='Plan de rupture N° ' ;
     TChoix.FindField('CC_LIBELLE').AsString:=St1+IntToStr(i)+' '+TbRupt[i] ;
     TChoix.FindField('CC_ABREGE').AsString:=TbRupt[i] ;
     TChoix.FindField('CC_LIBRE').AsString:=TrouveIndexRupt(2,i) ;
     TChoix.Post ;
     END ;
  END ;
TChoix.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupCodeRegime ;
var i : Integer ;
begin
VideTable('CHOIXCOD','WHERE CC_TYPE="RTV"') ;
BeginTrans ;
TChoix.open ;
For i:=1 To 19 Do
  BEGIN
  If TbReg[i]<>'' Then
     BEGIN
     TChoix.Insert ;
     InitNew(TChoix) ;
     TChoix.FindField('CC_TYPE').AsString:='RTV' ;
     TChoix.FindField('CC_CODE').AsString:=TbReg[i] ;
     TChoix.FindField('CC_LIBELLE').AsString:=TbLibReg[i] ;
     TChoix.FindField('CC_ABREGE').AsString:=TbLibReg[i] ;
     TChoix.Post ;
     END ;
  END ;
TChoix.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupRupt(C : Char) ;
Var i,j : SmallInt ;
    NoTable,LaTable : Byte ;
    Fiche16 : EnregCpta ;
    Pos16,kk : LongInt ;
    Cle,Test : String40 ;
    OkOk : Boolean ;
    St : String17 ;
    CC : String3 ;
begin
LaTable:=0 ;
Case C Of
  'G' : BEGIN CC:='RUG' ; LaTable:=19 ; END ;
  'T' : BEGIN CC:='RUT' ; LaTable:=140 ; END ;
  'S' : BEGIN CC:='RU1' ; LaTable:=120 ; END ;
  END ;
VideTable('RUPTURE','WHERE RU_NATURERUPT="'+CC+'"') ;
BeginTrans ;
TRupt.open ;
Ouvre(16,2,FALSE,spCPTA) ;
For i:=1 To 19 Do
  BEGIN
  NoTable:=LaTable+i ;
  If TbRupt[i]<>'' Then
     BEGIN
     Cle:=CreCle16(1,NoTable,'','','',0) ;  ; Test:=Cle ; kk:=0 ;
     Repeat
       Inc(kk) ;
       If kk=1 then SearchKey(VSAA^.Idxff[16],Pos16,Cle) else NextKey(VSAA^.Idxff[16],Pos16,Cle) ;
       OkOk:=VSAA^.Ok and (Copy(Cle,1,Length(Test))=Test) ;
       if OkOk then
          BEGIN
          GetRec(VSAA^.Datff[16],Pos16,Fiche16) ;
          TRupt.Insert ;
          InitNew(TRupt) ;
          TRupt.FindField('RU_NATURERUPT').AsString:=CC ;
          TRupt.FindField('RU_PLANRUPT').AsString:=FormatFloat('000',i) ;
          St:=Fiche16.CodetrancheTA ;
          For j:=1 To Length(St) Do If St[j]='°' Then St[j]:=' ' ; St:=Trim(St)+'x' ;
          TRupt.FindField('RU_CLASSE').AsString:=st ;
          TRupt.FindField('RU_LIBELLECLASSE').AsString:=ASCII2ANSI(Fiche16.IntituleTA) ;
          TRupt.Post ;
          END ;
     Until Not OkOk ;
     END ;
  END ;
GAccess.Ferme(16,2,FALSE) ;
TRupt.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupTranche ;
Var i : SmallInt ;
    NoTable : Byte ;
    Fiche16 : EnregCpta ;
    Pos16,kk : LongInt ;
    Cle,Test : String40 ;
    OkOk : Boolean ;
begin
VideTable('SSSTRUCR','WHERE PS_AXE="A1"') ;
BeginTrans ;
TstrAn.open ;
Ouvre(16,2,FALSE,spCPTA) ;
For i:=1 To 19 Do
  BEGIN
  NoTable:=i ;
  If TbRupt[i]<>'' Then
     BEGIN
     Cle:=CreCle16(1,NoTable,'','','',0) ;  ; Test:=Cle ; kk:=0 ;
     Repeat
       Inc(kk) ;
       If kk=1 then SearchKey(VSAA^.Idxff[16],Pos16,Cle) else NextKey(VSAA^.Idxff[16],Pos16,Cle) ;
       OkOk:=VSAA^.Ok and (Copy(Cle,1,Length(Test))=Test) ;
       if OkOk then
          BEGIN
          GetRec(VSAA^.Datff[16],Pos16,Fiche16) ;
          TstrAn.Insert ;
          InitNew(TstrAn) ;
          TstrAn.FindField('PS_AXE').AsString:='A1' ;
          TstrAn.FindField('PS_SOUSSECTION').AsString:=FormatFloat('000',i) ;
          TstrAn.FindField('PS_CODE').AsString:=ASCII2ANSI(Fiche16.CodeTrancheTA) ;
          TstrAn.FindField('PS_LIBELLE').AsString:=ASCII2ANSI(Fiche16.IntituleTA) ;
          TstrAn.FindField('PS_ABREGE').AsString:=ASCII2ANSI(Fiche16.IntituleTA) ;
          TstrAn.Post ;
          END ;
     Until Not OkOk ;
     END ;
  END ;
GAccess.Ferme(16,2,FALSE) ;
TstrAn.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupCodeTranche ;
var i : Integer ;
    Deb,Lg : SmallInt ;
begin
VideTable('STRUCRSE','WHERE SS_AXE="A1"') ;
BeginTrans ;
TStrTra.open ;
For i:=1 To 19 Do
  BEGIN
  If TbRupt[i]<>'' Then
     BEGIN
     TStrTra.Insert ;
     InitNew(TStrTra) ;
     TStrTra.FindField('SS_AXE').AsString:='A1' ;
     TStrTra.FindField('SS_SOUSSECTION').AsString:=FormatFloat('000',i) ;
     TStrTra.FindField('SS_LIBELLE').AsString:=TbRupt[i] ;
     TStrTra.FindField('SS_CONTROLE').AsString:='-' ;
     LitStructureTrancheV8(i,deb,lg) ;
     TStrTra.FindField('SS_DEBUT').AsInteger:=Deb ;
     TStrTra.FindField('SS_LONGUEUR').AsInteger:=Lg ;
     TStrTra.Post ;
     END ;
  END ;
TStrTra.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupRuptures ;
begin
LitRuptV8('G') ; RecupCodeRupt('RUG') ; RecupRupt('G') ;
LitRuptV8('T') ; RecupCodeRupt('RUT') ; RecupRupt('T') ;
If AnaHal.Checked Then
   BEGIN
   LitRuptV8('S') ; RecupCodeRupt('RU1') ; RecupRupt('S') ;
   LitRuptV8('A') ; RecupCodeTranche ; RecupTranche ;
   END ;

end ;

{=============================================================================}
procedure TFImpCpt.recupSoc ;
Var laSoc : String3 ;
    QSoc : TQuery ;
    VSoc : Integer ;
begin
VSoc:=47 ;
QSoc := OpenSQL('SELECT SO_VERSIONBASE FROM SOCIETE',True) ;
If Not QSoc.Eof Then VSoc := QSoc.FindField('SO_VERSIONBASE').AsInteger ;
HCtrls.Ferme(QSoc) ;
VideTable('ETABLISS','') ;
VideTable('SOCIETE','') ;
BeginTrans ;
TSoc.open ;
TSoc.Insert ;
InitNew(TSoc) ;
{$IFDEF SPEC302}
With VSAA^.ParaSoc Do
  BEGIN
  LaSoc:=Trim(Copy(CodeSociete,1,3)) ;
  If LaSoc='' Then LaSoc:='001' ;
  TSoc.findfield('SO_SOCIETE').asstring:=LaSoc ;
  TSoc.findfield('SO_ETABLISDEFAUT').asstring:=LaSoc ;
  TSoc.findfield('SO_LIBELLE').asstring:=Trim(Copy(RaisonSoc,1,35)) ;
  TSoc.findfield('SO_ADRESSE1').asstring:=Trim(Adresse1) ;
  TSoc.findfield('SO_ADRESSE2').asstring:=Trim(Adresse2) ;
  TSoc.findfield('SO_ADRESSE3').asstring:=Trim(Adresse3) ;
  TSoc.findfield('SO_CODEPOSTAL').asstring:=Trim(Copy(CodePostal,1,9));
  TSoc.findfield('SO_VILLE').asstring:=Trim(Ville) ;
  TSoc.findfield('SO_DIVTERRIT').asstring:='' ;
  TSoc.findfield('SO_TELEPHONE').asstring:='' ;
  TSoc.findfield('SO_TELEX').asstring:='' ;
  TSoc.findfield('SO_NIF').asstring:='' ;
  TSoc.findfield('SO_APE').asstring:=Trim(Ape) ;
  TSoc.findfield('SO_SIRET').asstring:='' ;
  TSoc.findfield('SO_CONTACT').asstring:='' ;
  TSoc.findfield('SO_LGCPTEGEN').asInteger:=FormatGen.Lg ;
  TSoc.findfield('SO_LGCPTEAUX').asInteger:=FormatAux.Lg ;
  TSoc.findfield('SO_BOURREGEN').AsString:='0' ;
  TSoc.findfield('SO_BOURREAUX').AsString:='0' ;
  TSoc.findfield('SO_VERIFRIB').asstring:='-' ;
  If VerifRIB=1 Then TSoc.findfield('SO_VERIFRIB').asstring:='X' ;
  TSoc.findfield('SO_MONTANTNEGATIF').asstring:='-' ;
  If MontantNegatif=1 Then  TSoc.findfield('SO_MONTANTNEGATIF').asstring:='X' ;
  TSoc.findfield('SO_SOLDEGEN').asstring:='X' ;
  TSoc.findfield('SO_NBJECRAVANT').asInteger:=PlageCpta.Deb ;
  TSoc.findfield('SO_NBJECRAPRES').asInteger:=PlageCpta.Fin ;
  TSoc.findfield('SO_NBJECHAVANT').asInteger:=PlageEche.Deb ;
  TSoc.findfield('SO_NBJECHAPRES').asInteger:=PlageEche.Fin ;
  Case TvaEncVente Of
    2 : TSoc.findfield('SO_TVAENCAISSEMENT').asstring:='TE' ;
    3 : TSoc.findfield('SO_TVAENCAISSEMENT').asstring:='TM' ;
    Else TSoc.findfield('SO_TVAENCAISSEMENT').asstring:='TD' ;
    END ;
  TSoc.findfield('SO_BILDEB1').asstring:=Trim(CptesBilan.Deb) ;
  TSoc.findfield('SO_BILFIN1').asstring:=Trim(CptesBilan.Fin) ;
  TSoc.findfield('SO_CHADEB1').asstring:=Trim(Cha.Deb) ;
  TSoc.findfield('SO_CHAFIN1').asstring:=Trim(Cha.Fin) ;
  TSoc.findfield('SO_PRODEB1').asstring:=Trim(Pro.Deb) ;
  TSoc.findfield('SO_PROFIN1').asstring:=Trim(Pro.Fin) ;
  TSoc.findfield('SO_GENATTEND').asstring:=Trim(GeneAttend) ;
  TSoc.findfield('SO_CLIATTEND').asstring:=Trim(ClientAttend) ;
  TSoc.findfield('SO_FOUATTEND').asstring:=Trim(FournAttend) ;
  TSoc.findfield('SO_RESULTAT').asstring:=Trim(Resultat) ;
  TSoc.findfield('SO_OUVREBIL').asstring:=Trim(OuvreBilan) ;
  TSoc.findfield('SO_FERMEBIL').asstring:=Trim(OuvreBilan) ;
  TSoc.findfield('SO_OUVREPERTE').asstring:=Trim(Perte) ;
  TSoc.findfield('SO_FERMEPERTE').asstring:=Trim(Perte) ;
  TSoc.findfield('SO_OUVREBEN').asstring:=Trim(Benefice) ;
  TSoc.findfield('SO_FERMEBEN').asstring:=Trim(Benefice) ;
  TSoc.findfield('SO_JALFERME').asstring:=Trim(JalFerme) ;
  TSoc.findfield('SO_JALOUVRE').asstring:=Trim(JalOuvre) ;
  TSoc.findfield('SO_DATECLOTUREPER').asDateTime:=Int2Date(VSAA^.CloturePer) ;
  TSoc.findfield('SO_DATECLOTUREPRO').asDateTime:=Int2Date(VSAA^.CloturePro) ;
  TSoc.findfield('SO_DATEDERNENTREE').asDateTime:=Int2Date(VSAA^.LastDate) ;
  TSoc.findfield('SO_CORSGE1').asstring:='X' ;
  TSoc.findfield('SO_CORSGE2').asstring:='-' ;
  TSoc.findfield('SO_CORSAU1').asstring:='X' ;
  TSoc.findfield('SO_CORSAU2').asstring:='-' ;
  TSoc.findfield('SO_CORSA11').asstring:='X' ;
  TSoc.findfield('SO_CORSA12').asstring:='-' ;
  TSoc.findfield('SO_CORSA21').asstring:='-' ;
  TSoc.findfield('SO_CORSA22').asstring:='-' ;
  TSoc.findfield('SO_CORSA31').asstring:='-' ;
  TSoc.findfield('SO_CORSA32').asstring:='-' ;
  TSoc.findfield('SO_CORSA41').asstring:='-' ;
  TSoc.findfield('SO_CORSA42').asstring:='-' ;
  TSoc.findfield('SO_CORSA51').asstring:='-' ;
  TSoc.findfield('SO_CORSA52').asstring:='-' ;
  TSoc.findfield('SO_CORSBU1').asstring:='-' ;
  TSoc.findfield('SO_CORSBU2').asstring:='-' ;
  TSoc.findfield('SO_DEVISEPRINC').asstring:=Trim(Canabis) ;
  TSoc.findfield('SO_FURTIVITE').AsInteger:=1000 ;
  If Trim(Canabis)='' Then TSoc.findfield('SO_DEVISEPRINC').asstring:='FRF' ;
  TSoc.FindField('SO_RECUPCPTA').ASString:='100000000000000000000000000' ;
  TSoc.findfield('SO_DECVALEUR').AsInteger:=Dc ;
  TSoc.findfield('SO_DECQTE').AsInteger:=2 ;
  TSoc.findfield('SO_DECPRIX').AsInteger:=2 ;
  TSoc.findfield('SO_VERSIONBASE').AsInteger:=VSoc ;
  END ;
{$ELSE}
With VSAA^.ParaSoc Do
  BEGIN
  LaSoc:=Trim(Copy(CodeSociete,1,3)) ;
  If LaSoc='' Then LaSoc:='001' ;
  TSoc.FindField('SO_SOCIETE').AsString:=LaSoc ;
  TSoc.FindField('SO_VERSIONBASE').AsInteger:=VSoc ;

  SetParamSoc('SO_ETABLISDEFAUT',LaSoc) ;
  SetParamSoc('SO_LIBELLE',Trim(Copy(RaisonSoc,1,35))) ;
  SetParamSoc('SO_ADRESSE1',Trim(Adresse1)) ;
  SetParamSoc('SO_ADRESSE2',Trim(Adresse2)) ;
  SetParamSoc('SO_ADRESSE3',Trim(Adresse3)) ;
  SetParamSoc('SO_CODEPOSTAL',Trim(Copy(CodePostal,1,9))) ;
  SetParamSoc('SO_VILLE',Trim(Ville)) ;
  SetParamSoc('SO_DIVTERRIT','') ;
  SetParamSoc('SO_TELEPHONE','') ;
  SetParamSoc('SO_TELEX','') ;
  SetParamSoc('SO_NIF','') ;
  SetParamSoc('SO_APE',Trim(Ape)) ;
  SetParamSoc('SO_SIRET','') ;
  SetParamSoc('SO_CONTACT','') ;
  SetParamSoc('SO_LGCPTEGEN',FormatGen.Lg) ;
  SetParamSoc('SO_LGCPTEAUX',FormatAux.Lg) ;
  SetParamSoc('SO_BOURREGEN','0') ;
  SetParamSoc('SO_BOURREAUX','0') ;
  SetParamSoc('SO_VERIFRIB','-') ; If VerifRIB=1 Then SetParamSoc('SO_VERIFRIB','X') ;
  SetParamSoc('SO_MONTANTNEGATIF','-') ;
  If MontantNegatif=1 Then  SetParamSoc('SO_MONTANTNEGATIF','X') ;
  SetParamSoc('SO_SOLDEGEN','X') ;
  SetParamSoc('SO_NBJECRAVANT',PlageCpta.Deb) ;
  SetParamSoc('SO_NBJECRAPRES',PlageCpta.Fin) ;
  SetParamSoc('SO_NBJECHAVANT',PlageEche.Deb) ;
  SetParamSoc('SO_NBJECHAPRES',PlageEche.Fin) ;
  Case TvaEncVente Of
    2 : SetParamSoc('SO_TVAENCAISSEMENT','TE') ;
    3 : SetParamSoc('SO_TVAENCAISSEMENT','TM') ;
    Else SetParamSoc('SO_TVAENCAISSEMENT','TD') ;
    END ;
  SetParamSoc('SO_BILDEB1',Trim(CptesBilan.Deb)) ;
  SetParamSoc('SO_BILFIN1',Trim(CptesBilan.Fin)) ;
  SetParamSoc('SO_CHADEB1',Trim(Cha.Deb)) ;
  SetParamSoc('SO_CHAFIN1',Trim(Cha.Fin)) ;
  SetParamSoc('SO_PRODEB1',Trim(Pro.Deb)) ;
  SetParamSoc('SO_PROFIN1',Trim(Pro.Fin)) ;
  SetParamSoc('SO_GENATTEND',Trim(GeneAttend)) ;
  SetParamSoc('SO_CLIATTEND',Trim(ClientAttend)) ;
  SetParamSoc('SO_FOUATTEND',Trim(FournAttend)) ;
  SetParamSoc('SO_RESULTAT',Trim(Resultat)) ;
  SetParamSoc('SO_OUVREBIL',Trim(OuvreBilan)) ;
  SetParamSoc('SO_FERMEBIL',Trim(OuvreBilan)) ;
  SetParamSoc('SO_OUVREPERTE',Trim(Perte)) ;
  SetParamSoc('SO_FERMEPERTE',Trim(Perte)) ;
  SetParamSoc('SO_OUVREBEN',Trim(Benefice)) ;
  SetParamSoc('SO_FERMEBEN',Trim(Benefice)) ;
  SetParamSoc('SO_JALFERME',Trim(JalFerme)) ;
  SetParamSoc('SO_JALOUVRE',Trim(JalOuvre)) ;
  SetParamSoc('SO_DATECLOTUREPER',Int2Date(VSAA^.CloturePer)) ;
  SetParamSoc('SO_DATECLOTUREPRO',Int2Date(VSAA^.CloturePro)) ;
  SetParamSoc('SO_DATEDERNENTREE',Int2Date(VSAA^.LastDate)) ;
  SetParamSoc('SO_CORSGE1','X') ;
  SetParamSoc('SO_CORSGE2','-') ;
  SetParamSoc('SO_CORSAU1','X') ;
  SetParamSoc('SO_CORSAU2','-') ;
  SetParamSoc('SO_CORSA11','X') ;
  SetParamSoc('SO_CORSA12','-') ;
  SetParamSoc('SO_CORSA21','-') ;
  SetParamSoc('SO_CORSA22','-') ;
  SetParamSoc('SO_CORSA31','-') ;
  SetParamSoc('SO_CORSA32','-') ;
  SetParamSoc('SO_CORSA41','-') ;
  SetParamSoc('SO_CORSA42','-') ;
  SetParamSoc('SO_CORSA51','-') ;
  SetParamSoc('SO_CORSA52','-') ;
  SetParamSoc('SO_CORSBU1','-') ;
  SetParamSoc('SO_CORSBU2','-') ;
  SetParamSoc('SO_DEVISEPRINC',Trim(Canabis)) ;
  SetParamSoc('SO_FURTIVITE',1000) ;
  If Trim(Canabis)='' Then SetParamSoc('SO_DEVISEPRINC','FRF') ;
  SetParamSoc('SO_RECUPCPTA','100000000000000000000000000') ;
  SetParamSoc('SO_DECVALEUR',Dc) ;
  SetParamSoc('SO_DECQTE',2) ;
  SetParamSoc('SO_DECPRIX',2) ;
  END ;
{$ENDIF}
TSoc.Post ;
TSoc.Close ;

TEtab.Open ;
TEtab.Insert ;
InitNew(TEtab) ;
With VSAA^.ParaSoc Do
  BEGIN
  TEtab.findfield('ET_ETABLISSEMENT').asstring:=LaSoc;
  TEtab.findfield('ET_LIBELLE').asstring:=Trim(Copy(RaisonSoc,1,35));
  TEtab.findfield('ET_ABREGE').asstring:=Trim(CodeSociete) ;
  TEtab.findfield('ET_ADRESSE1').asstring:=Trim(Adresse1);
  TEtab.findfield('ET_ADRESSE2').asstring:=Trim(Adresse2);
  TEtab.findfield('ET_ADRESSE3').asstring:=Trim(Adresse3);
  TEtab.findfield('ET_CODEPOSTAL').asstring:=Trim(Copy(CodePostal,1,9));
  TEtab.findfield('ET_VILLE').asstring:=Trim(Ville);
  END ;
Tetab.Post ;
TEtab.Close ;
VH^.EtablisDefaut:=LaSoc ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupLibAuto ;
Var NoTable : Byte ;
    kk,Pos16 : LongInt ;
    Cle,Test : String40 ;
    OkOk : Boolean ;
    Fiche16 : EnregCPTA ;
begin
VideTable('REFAUTO','WHERE RA_JOURNAL=""') ;
BeginTrans ;
TRefAuto.Open ;
Ouvre(16,2,FALSE,spCPTA) ;
NoTable:=101 ;
Cle:=CreCle16(1,NoTable,'','','',0) ;  ; Test:=Cle ; kk:=0 ;
Repeat
  Inc(kk) ;
  If kk=1 then SearchKey(VSAA^.Idxff[16],Pos16,Cle) else NextKey(VSAA^.Idxff[16],Pos16,Cle) ;
  OkOk:=VSAA^.Ok and (Copy(Cle,1,Length(Test))=Test) ;
  if OkOk then
     BEGIN
     GetRec(VSAA^.Datff[16],Pos16,Fiche16) ;
     TRefAuto.Insert ;
     TRefAuto.FindField('RA_CODE').ASString:=Trim(Fiche16.CodeTrancheTA) ;
     TRefAuto.FindField('RA_LIBELLE').ASString:='Libellé automatique' ;
     TRefAuto.FindField('RA_JOURNAL').ASString:='' ;
     TRefAuto.FindField('RA_NATUREPIECE').ASString:='' ;
     TRefAuto.FindField('RA_FORMULEREF').ASString:='' ;
     TRefAuto.FindField('RA_FORMULELIB').ASString:=Trim(Fiche16.IntituleTA) ;
     TRefAuto.Post ;
     END ;
Until Not OkOk ;
TRefAuto.Close ;
GAccess.Ferme(16,2,FALSE) ; 
CommitTrans ;
end ;

{=============================================================================}
Function trouvesens(b:byte) : String3 ;
begin
Result:='D' ;
Case b Of
  1 : Result:='D' ;
  2 : Result:='C' ;
  3 : Result:='M' ;
  END ;
end ;

{=============================================================================}
procedure TFImpCpt.recupgeneraux ;
Var LaPos,Lg : LongInt ;
    Fiche : EnregCPTA ;
  Label 1 ;
begin
LitModRV8 ; AlimTVA ;
VideTable('GENERAUX','') ; VideTable('BANQUECP','') ; 
BeginTrans ;
Ouvre(0,0,FALSE,spCPTA) ; Lg:=FileLen(VSAA^.DAtFF[0])-1 ;
InitMove(Lg,'recup des generaux') ;
tGeneraux.open ;
For LaPos:=1 To Lg Do
  BEGIN
  if MoveCur(FALSE) or Stopper then goto 1;
  GetRec(VSAA^.DatFF[0],LaPos,Fiche) ;
  If Fiche.Status=0 Then
    BEGIN
    LMess('Création du compte '+Fiche.CompteG) ;
    TGeneraux.Insert ;
    InitNew(TGeneraux) ;
    TGeneraux.findfield('G_GENERAL').asstring:=Trim(Fiche.compteG) ;
    TGeneraux.findfield('G_LIBELLE').asstring:=ASCII2ANSI(Trim(Fiche.IntituleG)) ;
    TGeneraux.findfield('G_ABREGE').asstring:=ASCII2ANSI(Trim(Fiche.AbregeG)) ;
    TGeneraux.findfield('G_NATUREGENE').asstring:='DIV' ;
    TGeneraux.findfield('G_SUIVITRESO').asstring:='RIE' ;
    Case Fiche.NatureG Of
      1 : BEGIN
          If Fiche.LettrableG=1 Then TGeneraux.findfield('G_NATUREGENE').asstring:='TID' Else
             If Fiche.CollectifG=1 Then TGeneraux.findfield('G_NATUREGENE').asstring:='COC' ;
          TGeneraux.findfield('G_SUIVITRESO').asstring:='ENC' ;
          END ;
      2 : BEGIN
          If Fiche.LettrableG=1 Then TGeneraux.findfield('G_NATUREGENE').asstring:='TIC' Else
             If Fiche.CollectifG=1 Then TGeneraux.findfield('G_NATUREGENE').asstring:='COF' ;
          TGeneraux.findfield('G_SUIVITRESO').asstring:='DEC' ;
          END ;
      3 : TGeneraux.findfield('G_NATUREGENE').asstring:='IMO' ;
      4 : If Fiche.CollectifG=1 Then TGeneraux.findfield('G_NATUREGENE').asstring:='COS' ;
      5 : TGeneraux.findfield('G_NATUREGENE').asstring:='BQE' ;
      6 : TGeneraux.findfield('G_NATUREGENE').asstring:='CAI' ;
      7 : TGeneraux.findfield('G_NATUREGENE').asstring:='CHA' ;
      8 : TGeneraux.findfield('G_NATUREGENE').asstring:='PRO' ;
      Else If Fiche.CollectifG=1 Then TGeneraux.findfield('G_NATUREGENE').asstring:='COD' ;
      END ;
(*
QLoc:=OpenSql('Select PQ_BANQUE From BANQUES',True) ;
St:='' ; if Not QLoc.Eof then St:=QLoc.Fields[0].AsString ; Ferme(QLoc) ;
ExecuteSql('INSERT INTO BANQUECP (BQ_CODE, BQ_GENERAL, BQ_BANQUE, BQ_LIBELLE, BQ_DEVISE) '+
           'VALUES ("'+G_GENERAL.text+'", "'+G_GENERAL.text+'", "'+St+'", "'+MsgBox.Mess[35]+'", '+
           '"'+V_PGI.DevisePivot+'")') ;
*)
    If Fiche.NatureG=5 Then
       BEGIN
       TbanqueCP.Open ; TbanqueCP.Insert ; InitNew(TbanqueCP) ;
       TbanqueCP.findfield('BQ_GENERAL').asString:=Trim(Fiche.CompteG) ;
       TbanqueCP.findfield('BQ_CODE').asString:=Trim(Fiche.CompteG) ;
       TbanqueCP.findfield('BQ_LIBELLE').asString:=Trim(Fiche.IntituleG) ;
       TbanqueCP.findfield('BQ_DEVISE').asString:=TbDev[1].ISO ;
       TbanqueCP.post ;
       TbanqueCP.close ;
       END ;
    If Fiche.LettrableG=1 Then
       BEGIN
       TGeneraux.findfield('G_MODEREGLE').asstring:=TbMP[Fiche.ModePaiementG].Code ;
       TGeneraux.findfield('G_REGIMETVA').asstring:=TbReg[1] ;
       END ;
    If fiche.CentraliseG=1 Then TGeneraux.findfield('G_CENTRALISABLE').asstring:='X'
                           Else TGeneraux.findfield('G_CENTRALISABLE').asstring:='-' ;
    TGeneraux.findfield('G_PURGEABLE').asstring:='X' ;
    If fiche.SoldeProgrG=1 Then TGeneraux.findfield('G_SOLDEPROGRESSIF').asstring:='X'
                           Else TGeneraux.findfield('G_SOLDEPROGRESSIF').asstring:='-' ;
    If Fiche.SautPageG=1 Then TGeneraux.findfield('G_SAUTPAGE').asstring:='X'
                         Else TGeneraux.findfield('G_SAUTPAGE').asstring:='-' ;
    If Fiche.TotMensG=1 Then TGeneraux.findfield('G_TOTAUXMENSUELS').asstring:='X'
                        Else TGeneraux.findfield('G_TOTAUXMENSUELS').asstring:='-' ;
    If Fiche.collectifG=1 Then TGeneraux.findfield('G_COLLECTIF').asstring:='X'
                          Else TGeneraux.findfield('G_COLLECTIF').asstring:='-' ;
    If Fiche.NatureG In [5,6] Then
       BEGIN
       TGeneraux.findfield('G_COLLECTIF').asstring:='-' ;
       END ;
    TGeneraux.findfield('G_DATECREATION').asdatetime:=Int2Date(Fiche.DatecreatG) ;
    TGeneraux.findfield('G_DATEMODIF').asdatetime:=Int2Date(Fiche.DateModifG) ;
    TGeneraux.findfield('G_DATEOUVERTURE').asdatetime:=Int2Date(Fiche.DatecreatG) ;
    TGeneraux.findfield('G_DATEFERMETURE').asDateTime:=Int2Date(Fiche.DateFermeG) ;
    If Fiche.DateFermeG>10 Then TGeneraux.findfield('G_FERME').asstring:='X'
                           Else TGeneraux.findfield('G_FERME').asstring:='-' ;
    TGeneraux.findfield('G_DATEDERNMVT').asDatetime:=Int2Date(Fiche.DateMvtG) ;
    TGeneraux.findfield('G_DEBITDERNMVT').asFloat:=0 ;
    TGeneraux.findfield('G_CREDITDERNMVT').asFloat:=0 ;
    If Fiche.SensMvtG=1 Then TGeneraux.findfield('G_DEBITDERNMVT').asFloat:=Arrondi(Fiche.TotMvtG,Dc)
                        Else TGeneraux.findfield('G_CREDITDERNMVT').asFloat:=Arrondi(Fiche.TotMvtG,Dc) ;
    TGeneraux.findfield('G_NUMDERNMVT').asInteger:=Fiche.DerNumPieceG ;
    TGeneraux.findfield('G_LIGNEDERNMVT').asInteger:=Fiche.DerNumLigneG ;
    TGeneraux.findfield('G_SENS').asstring:=TrouveSens(Fiche.SensG) ;
    TGeneraux.findfield('G_TOTALDEBIT').asFloat:=0 ;
    TGeneraux.findfield('G_TOTALCREDIT').asFloat:=0 ;
    If Fiche.LettrableG=1 Then TGeneraux.findfield('G_LETTRABLE').asstring:='X'
                          Else TGeneraux.findfield('G_LETTRABLE').asstring:='-' ;
    If Fiche.POintableG=1 Then TGeneraux.findfield('G_POINTABLE').asstring:='X'
                          Else TGeneraux.findfield('G_POINTABLE').asstring:='-' ;
    Case Fiche.NatureG Of { Charge/produit/immo non pointable }
      3,7,8 : TGeneraux.findfield('G_POINTABLE').asstring:='-' ;
      END ;
    TGeneraux.findfield('G_VENTILABLE1').asstring:='-' ;
    TGeneraux.findfield('G_VENTILABLE2').asstring:='-' ;
    TGeneraux.findfield('G_VENTILABLE3').asstring:='-' ;
    TGeneraux.findfield('G_VENTILABLE4').asstring:='-' ;
    TGeneraux.findfield('G_VENTILABLE5').asstring:='-' ;
    If (Fiche.VentilableG=1) And
       (AnaHal.Checked)    Then BEGIN
                                TGeneraux.findfield('G_VENTILABLE').asstring:='X' ;
                                TGeneraux.findfield('G_VENTILABLE1').asstring:='X' ;
                                END
                           Else TGeneraux.findfield('G_VENTILABLE').asstring:='-' ;
    If Fiche.NatureG In [7,8] Then
       BEGIN
       TGeneraux.findfield('G_LETTRABLE').asstring:='-' ;
       TGeneraux.findfield('G_POINTABLE').asstring:='-' ;
       END ;
    TGeneraux.findfield('G_QUALIFQTE1').asstring:='AUC' ;
    TGeneraux.findfield('G_QUALIFQTE2').asstring:='AUC' ;
    TGeneraux.findfield('G_CORRESP1').asstring:='' ;
    TGeneraux.findfield('G_CORRESP2').asstring:='' ;
    TGeneraux.findfield('G_TVAENCAISSEMENT').asstring:='-' ;
    {
    TGeneraux.findfield('G_TVA').asstring:= ;
    }
    TGeneraux.findfield('G_BUDGENE').asstring:='' ;
    TGeneraux.findfield('G_RISQUETIERS').asstring:='-' ;
    TGeneraux.findfield('G_DERNLETTRAGE').asstring:=codelettre(Fiche.DerNumLettreG,TRUE) ;
    TGeneraux.findfield('G_DEBNONPOINTE').asFloat:=Arrondi(Fiche.TotPointeNG[1],Dc) ;
    TGeneraux.findfield('G_CREDNONPOINTE').asFloat:=Arrondi(Fiche.TotPointeNG[2],Dc) ;
    LitBloc(Fiche.PosBlocG,'0',spCPTA,TGeneraux.findfield('G_BLOCNOTE')) ;
    TGeneraux.post ;
    END ;
  END ;
1:TGeneraux.Close ;
GAccess.Ferme(0,0,FALSE) ;
FiniMove ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.recupAuxiliaire ;
Var LaPos,Lg : LongInt ;
    Pos13,kk,NoRib : LongInt ;
    OkOk : Boolean ;
    Fiche,Fiche12,Fiche13 : EnregCPTA ;
    Cle,Test : String40 ;
    NatGen : String3 ;
  Label 1 ;
begin
LitModRV8 ; AlimTva ; AlimRel ; MajTableRelance ;
VideTable('RIB','') ; VideTable('TIERS','') ;
BeginTrans ;
Ouvre(1,0,FALSE,spCPTA) ; Ouvre(12,0,FALSE,spCPTA) ; Ouvre(13,2,FALSE,spCPTA) ;
Lg:=FileLen(VSAA^.DAtFF[1])-1 ;
InitMove(Lg,'recup des Auxiliaires') ;
TRIB.Open ; TTiers.open ; TGeneraux.Open ;
For LaPos:=1 To Lg Do
  BEGIN
  if MoveCur(FALSE) or Stopper then goto 1;
  GetRec(VSAA^.DatFF[1],LaPos,Fiche) ;
  If Fiche.Status=0 Then
    BEGIN
    LMess('Création du tiers '+Fiche.AuxiliaireT) ;
    TTiers.Insert ;
    InitNew(TTiers) ;
    TTiers.findfield('T_AUXILIAIRE').asstring:=Trim(Fiche.AuxiliaireT) ;
    TTiers.findfield('T_LIBELLE').asstring:=ASCII2ANSI(Trim(Fiche.IntituleT)) ;
    TTiers.findfield('T_ABREGE').asstring:=ASCII2ANSI(Trim(Fiche.AbregeT)) ;
    TTiers.findfield('T_COLLECTIF').asstring:=Trim(Fiche.CollectifT) ;
    TTiers.findfield('T_NATUREAUXI').asstring:='DIV' ;
    Case Fiche.NatureT Of
      1 : TTiers.findfield('T_NATUREAUXI').asstring:='CLI' ;
      2 : TTiers.findfield('T_NATUREAUXI').asstring:='FOU' ;
      4 : TTiers.findfield('T_NATUREAUXI').asstring:='SAL' ;
      Else BEGIN
           If TGeneraux.FindKey([Trim(Fiche.CollectifT)]) Then
              BEGIN
              NatGen:=TGeneraux.FindField('G_NATUREGENE').AsString ;
              If NatGen='COC' Then TTiers.findfield('T_NATUREAUXI').asstring:='AUD' Else
                 If NatGen='COF' Then TTiers.findfield('T_NATUREAUXI').asstring:='AUC' Else
                    If NatGen='COD' Then TTiers.findfield('T_NATUREAUXI').asstring:='DIV' ;
              END ;
           END ;
      END ;
    If fiche.SoldeProgrT=1 Then TTiers.findfield('T_SOLDEPROGRESSIF').asstring:='X'
                           Else TTiers.findfield('T_SOLDEPROGRESSIF').asstring:='-' ;
    If Fiche.SautPageT=1 Then TTiers.findfield('T_SAUTPAGE').asstring:='X'
                         Else TTiers.findfield('T_SAUTPAGE').asstring:='-' ;
    If Fiche.TotMensT=1 Then TTiers.findfield('T_TOTAUXMENSUELS').asstring:='X'
                        Else TTiers.findfield('T_TOTAUXMENSUELS').asstring:='-' ;
    TTiers.findfield('T_DATECREATION').asdatetime:=Int2Date(Fiche.DatecreatT) ;
    TTiers.findfield('T_DATEMODIF').asdatetime:=Int2Date(Fiche.DateModifT);
    TTiers.findfield('T_DATEOUVERTURE').asdatetime:=Int2Date(Fiche.DatecreatT) ;
    TTiers.findfield('T_DATEFERMETURE').asDateTime:=Int2Date(Fiche.DateFermeT) ;
    If Fiche.DateFermeT>10 Then TTiers.findfield('T_FERME').asstring:='X'
                           Else TTiers.findfield('T_FERME').asstring:='-' ;
    TTiers.findfield('T_DATEDERNMVT').asDatetime:=Int2Date(Fiche.DateMvtT) ;
    TTiers.findfield('T_DEBITDERNMVT').asFloat:=0 ;
    TTiers.findfield('T_CREDITDERNMVT').asFloat:=0 ;
    If Fiche.SensMvtT=1 Then TTiers.findfield('T_DEBITDERNMVT').asFloat:=Arrondi(Fiche.TotMvtT,Dc)
                        Else TTiers.findfield('T_CREDITDERNMVT').asFloat:=Arrondi(Fiche.TotMvtT,Dc) ;
    TTiers.findfield('T_NUMDERNMVT').asInteger:=Fiche.DerNumPieceT ;
    TTiers.findfield('T_LIGNEDERNMVT').asInteger:=Fiche.DerNumLigneT ;
    TTiers.findfield('T_TOTALDEBIT').asFloat:=0 ;
    TTiers.findfield('T_TOTALCREDIT').asFloat:=0 ;
    If Fiche.LettrableT=1 Then TTiers.findfield('T_LETTRABLE').asstring:='X'
                          Else TTiers.findfield('T_LETTRABLE').asstring:='-' ;
    If Fiche.LettrableT=1 Then
       BEGIN
       TTiers.findfield('T_MODEREGLE').asstring:=TbMP[Fiche.PaiementT].Code ;
       If Fiche.RelanceT>1 Then TTiers.findfield('T_RELANCEREGLEMENT').asstring:=FormatFloat('000',Fiche.RelanceT+2) ;
       END ;
    TTiers.findfield('T_CORRESP1').asstring:='' ;
    TTiers.findfield('T_CORRESP2').asstring:='' ;
    TTiers.findfield('T_DERNLETTRAGE').asstring:=codelettre(Fiche.DerNumLettreT,TRUE) ;
    If (Fiche.RegimeTVAT>0) And (Fiche.RegimeTVAT<11)
       Then TTiers.findfield('T_REGIMETVA').asstring:=TbReg[Fiche.RegimeTVAT]
       Else TTiers.findfield('T_REGIMETVA').asstring:='FRA' ;
    TTiers.findfield('T_SECTEUR').asstring:='...' ;
    If Fiche.PosAdresseT>0 Then
       BEGIN
       GetRec(VSAA^.DAtFF[12],Fiche.PosAdresseT,Fiche12) ;
       If Fiche12.Status=0 Then
          BEGIN
          TTiers.findfield('T_ADRESSE1').asstring:=ASCII2ANSI(Fiche12.Adresse1GT) ;
          TTiers.findfield('T_ADRESSE2').asstring:=ASCII2ANSI(Fiche12.Adresse2GT) ;
          TTiers.findfield('T_ADRESSE3').asstring:=ASCII2ANSI(Fiche12.Adresse3GT) ;
          TTiers.findfield('T_CODEPOSTAL').asstring:=Fiche12.CodePostalGT ;
          TTiers.findfield('T_VILLE').asstring:=ASCII2ANSI(Fiche12.VilleGT) ;
          TTiers.findfield('T_NIF').asstring:=Fiche12.NoCommunGT ;
          TTiers.findfield('T_SIRET').asstring:=Fiche12.APEGT ;
          TTiers.findfield('T_TELEPHONE').asstring:=Fiche12.TelephoneGT ;
          TTiers.findfield('T_FAX').asstring:=Fiche12.FaxGT ;
          TTiers.findfield('T_COMMENTAIRE').asstring:=ASCII2ANSI(Fiche12.ContactGT) ;
          Case Fiche12.PaysGT Of
            1 : TTiers.findfield('T_PAYS').asstring:='FRA' ;
            Else TTiers.findfield('T_PAYS').asstring:='FRA' ;
            END ;
          { Voir Forme juridique }
          END ;
       END ;

    Cle:=CreCle13(1,LaPos,0) ;
    Test:=Cle ; kk:=0 ; NoRib:=0 ;
    Repeat
       Inc(kk) ;
       If kk=1 Then SearchKey(VSAA^.Idxff[13],Pos13,Cle) else Nextkey(VSAA^.Idxff[13],Pos13,Cle) ;
       OkOk:=VSAA^.Ok And (Copy(Cle,1,Length(Test))=Test) ;
       If OkOK Then
         BEGIN
         GetRec(VSAA^.Datff[13],Pos13,Fiche13) ;
         If FIche13.Status=0 Then
            BEGIN
            TRIB.Insert ;
            Inc(NoRib) ;
            TRIB.findfield('R_AUXILIAIRE').asstring:=Trim(Fiche.AuxiliaireT) ;
            TRIB.findfield('R_NUMERORIB').asinteger:=NoRib ;
            TRIB.findfield('R_PRINCIPAL').asstring:='-' ;
            If NoRib=1 Then TRIB.findfield('R_PRINCIPAL').asstring:='X' ;
            TRIB.findfield('R_ETABBQ').asstring:=Copy(Fiche13.CodeBanqueB,1,5) ;
            TRIB.findfield('R_GUICHET').asstring:=Copy(Fiche13.CodeGuichetB,1,5) ;
            TRIB.findfield('R_NUMEROCOMPTE').asstring:=Copy(Fiche13.CompteBanqueB,1,11) ;
            TRIB.findfield('R_CLERIB').asstring:=Copy(Fiche13.CleRibB,1,2) ;
            TRIB.findfield('R_DOMICILIATION').asstring:=Copy(Fiche13.BanqueB,1,24) ;
            TRib.Post ;
            END ;
         END ;
    Until Not OkOk ;
    LitBloc(Fiche.PosBlocT,'1',spCPTA,TTiers.findfield('T_BLOCNOTE')) ;
    TTiers.post ;
    END ;
  END ;
1:TGeneraux.Close ; TRIB.CLose ; TTiers.Close ;
GAccess.Ferme(1,0,FALSE) ; GAccess.Ferme(12,0,FALSE) ; GAccess.Ferme(13,2,FALSE) ;
FiniMove ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.recupSect ;
Var LaPos,Lg : LongInt ;
    Fiche : EnregCPTA ;
  Label 1 ;
begin

(*VideTable('AXE') ;*) VideTable('SECTION','') ; If Not AnaHal.Checked Then Exit ;
BeginTrans ;
Ouvre(2,0,FALSE,spCPTA) ; Lg:=FileLen(VSAA^.DAtFF[2])-1 ;
InitMove(Lg,'recup des Sections') ;

TAxe.Open ;
If TAxe.FindKey(['A1']) Then
   BEGIN
   TAxe.Edit ;
   TAxe.findfield('X_LONGSECTION').asInteger:=VSAA^.ParaSoc.FormatAna.Lg ;
   TAxe.findfield('X_BOURREANA').asString:='0' ;
   TAxe.findfield('X_SECTIONATTENTE').asString:=Trim(VSAA^.ParaSoc.SectionAttend) ;
   TAxe.Post ;
   END ;
TAxe.Close ;
TSection.open ;
For LaPos:=1 To Lg Do
  BEGIN
  if MoveCur(FALSE) or Stopper then goto 1;
  GetRec(VSAA^.DatFF[2],LaPos,Fiche) ;
  If Fiche.Status=0 Then
    BEGIN
    LMess('Création de la section '+Fiche.CompteA) ;
    TSection.Insert ;
    InitNew(TSection) ;
    TSection.findfield('S_SECTION').asstring:=Trim(Fiche.compteA) ;
    TSection.findfield('S_LIBELLE').asstring:=ASCII2ANSI(Trim(Fiche.IntituleA)) ;
    TSection.findfield('S_ABREGE').asstring:=ASCII2ANSI(Trim(Fiche.AbregeA)) ;
    If fiche.SoldeProgrA=1 Then TSection.findfield('S_SOLDEPROGRESSIF').asstring:='X'
                           Else TSection.findfield('S_SOLDEPROGRESSIF').asstring:='-' ;
    If Fiche.SautPageA=1 Then TSection.findfield('S_SAUTPAGE').asstring:='X'
                         Else TSection.findfield('S_SAUTPAGE').asstring:='-' ;
    If Fiche.TotMensA=1 Then TSection.findfield('S_TOTAUXMENSUELS').asstring:='X'
                        Else TSection.findfield('S_TOTAUXMENSUELS').asstring:='-' ;
    TSection.findfield('S_DATECREATION').asdatetime:=Int2Date(Fiche.DatecreatA) ;
    TSection.findfield('S_DATEMODIF').asdatetime:=Int2Date(Fiche.DateModifA) ;
    TSection.findfield('S_DATEOUVERTURE').asdatetime:=Int2Date(Fiche.DatecreatA) ;
    TSection.findfield('S_DATEFERMETURE').asDateTime:=Int2Date(Fiche.DateFermeA) ;
    If Fiche.DateFermeA>10 Then TSection.findfield('S_FERME').asstring:='X'
                           Else TSection.findfield('S_FERME').asstring:='-' ;
    TSection.findfield('S_DATEDERNMVT').asDatetime:=Int2Date(Fiche.DateMvtA) ;
    TSection.findfield('S_DEBITDERNMVT').asFloat:=0 ;
    TSection.findfield('S_CREDITDERNMVT').asFloat:=0 ;
    If Fiche.SensMvtA=1 Then TSection.findfield('S_DEBITDERNMVT').asFloat:=Arrondi(Fiche.TotMvtA,Dc)
                        Else TSection.findfield('S_CREDITDERNMVT').asFloat:=Arrondi(Fiche.TotMvtA,Dc) ;
    TSection.findfield('S_NUMDERNMVT').asInteger:=Fiche.DerNumPieceA ;
    TSection.findfield('S_LIGNEDERNMVT').asInteger:=Fiche.DerNumLigneA ;
    TSection.findfield('S_SENS').asstring:='M' ;
    TSection.findfield('S_TOTALDEBIT').asFloat:=0 ;
    TSection.findfield('S_TOTALCREDIT').asFloat:=0 ;
    TSection.findfield('S_CLEREPARTITION').asstring:='' ;
    TSection.findfield('S_CORRESP1').asstring:='' ;
    TSection.findfield('S_CORRESP2').asstring:='' ;
    TSection.findfield('S_AXE').asstring:='A1' ;
    TSection.findfield('S_MAITREOEUVRE').asstring:='' ;
    TSection.findfield('S_CHANTIER').asstring:='' ;
    TSection.findfield('S_MODELE').asstring:='-' ;
    LitBloc(Fiche.PosBlocA,'2',spCPTA,TSection.findfield('S_BLOCNOTE')) ;
    TSection.post ;
    END ;
  END ;
1:TSection.Close ;
GAccess.Ferme(2,0,FALSE) ;
FiniMove ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.addSouche(C : Char) ;

BEGIN
TSouche.Open ;
TSouche.Insert ;
TSouche.findfield('SH_TYPE').asstring:='CPT' ;
Case C Of
  'S' : BEGIN
        TSouche.findfield('SH_SOUCHE').asstring:='SIM' ;
        TSouche.findfield('SH_NUMDEPART').asInteger:=VSAA^.ParaSoc.DernSim+1 ;
        TSouche.findfield('SH_SIMULATION').asstring:='X' ;
        TSouche.findfield('SH_LIBELLE').asstring:='Simulation' ;
        TSouche.findfield('SH_ANALYTIQUE').asstring:='-' ;
        END ;
  'N' : BEGIN
        TSouche.findfield('SH_SOUCHE').asstring:='CPT' ;
        TSouche.findfield('SH_NUMDEPART').asInteger:=VSAA^.ParaSoc.DernPiece+1 ;
        TSouche.findfield('SH_SIMULATION').asstring:='-' ;
        TSouche.findfield('SH_LIBELLE').asstring:='Souche comptable' ;
        TSouche.findfield('SH_ANALYTIQUE').asstring:='-' ;
        END ;
  'A' : BEGIN
        TSouche.findfield('SH_SOUCHE').asstring:='ANA' ;
        TSouche.findfield('SH_NUMDEPART').asInteger:=VSAA^.ParaSoc.DernODAnal+1 ;
        TSouche.findfield('SH_SIMULATION').asstring:='-' ;
        TSouche.findfield('SH_LIBELLE').asstring:='OD analytique' ;
        TSouche.findfield('SH_ANALYTIQUE').asstring:='X' ;
        END ;
  END ;
TSouche.findfield('SH_NUMDEPARTS').asInteger:=1 ;
TSouche.findfield('SH_NUMDEPARTP').asInteger:=1 ;
TSouche.findfield('SH_SOUCHEEXO').asstring:='-' ;
TSouche.findfield('SH_ABREGE').asstring:='Souche multi journal' ;
TSouche.findfield('SH_NATUREPIECE').asstring:='' ;
TSouche.findfield('SH_JOURNAL').asstring:='' ;
TSouche.findfield('SH_MASQUENUM').asstring:='' ;
TSouche.findfield('SH_SOCIETE').asstring:='' ;
TSouche.findfield('SH_DATEDEBUT').asDateTime:=IDate1900 ;
TSouche.findfield('SH_DATEFIN').asDateTime:=iDate2099 ;
TSouche.findfield('SH_FERME').asstring:='-' ;
TSouche.Post ;
TSouche.Close ;
END ;

{=============================================================================}
Function AnalyseCptInt(Cpt1,Cpt2 : String) : String200 ;
Var i : SmallInt ;
begin
Result:='' ; If Cpt1=CPt2 Then BEGIN Result:=Cpt1 ; Exit ; END ;
i:=1 ; While (i<=VSAA^.ParaSoc.FormatGen.Lg) And (Cpt1[i]=Cpt2[i]) Do Inc(i) ;
Dec(i) ;
If i>0 Then Result:=Copy(Cpt1,1,i) Else Result:=Cpt1+':'+Cpt2 ;
end ;

{=============================================================================}
Function TFImpCpt.RecupCptJal(PosJal : LongInt ; TypeF : Byte ; Var Jal : EnregCpta) : String200 ;
Var Cle : String40 ;
    Fiche : enregCpta ;
    PosVentil,kk : longint ;
    Depart : SmallInt ;
    OkOk : Boolean ;
    St : String200 ;
begin
If Jal.NatureJ In [3,4] Then Exit ;
Result:='' ; St:='' ;
Depart:=1 ;
Cle:=CreCle15(TypeF,PosJal,Depart) ; kk:=0 ;
Repeat
  Inc(kk) ;
  If kk=1 then SearchKey(VSAA^.Idxff[15],PosVentil,Cle) else NextKey(VSAA^.Idxff[15],PosVentil,Cle) ;
  OkOk:=((VSAA^.Ok) and (Cle[1]=#1)) ;
  If OkOk then
     BEGIN
     GetRec(VSAA^.Datff[15],PosVentil,Fiche) ;
     OkOk:=(Fiche.TypeEnrVJ=TypeF) and (Fiche.PosCodeJalVJ=PosJal) ;
     If OkOk  then
        BEGIN
        Case TypeF Of
          CpteInt : BEGIN
                    If kk=1 Then St:=AnalyseCptInt(Trim(Fiche.Compte1VJ),Trim(Fiche.Compte2VJ))+';'
                            Else St:=St+AnalyseCptInt(Trim(Fiche.Compte1VJ),Trim(Fiche.Compte2VJ))+';'
                    END ;
          Ventil  : If ANAHal.Checked Then BEGIN
                    If kk=1 Then St:=AnalyseCptInt(Trim(Fiche.Compte1VJ),Trim(Fiche.Compte1VJ))+';'
                            Else St:=St+AnalyseCptInt(Trim(Fiche.Compte1VJ),Trim(Fiche.Compte1VJ))+';'
                    END ;
          END ;
        END ;
     END ;
Until Not OkOk ;
Result:=St ;
end ;

{=============================================================================}
procedure TFImpCpt.AlimRef(Fiche : EnregCpta) ;
Var i : Integer ;
    St : String ;
begin
For i:=1 To 7 Do
  BEGIN
  If Fiche.OptionsPieceJ[i].Actif=1 Then
     BEGIN
     St:=Trim(Fiche.OptionsPieceJ[i].RefAuto) ;
     If St<>'' Then LesRefJ.Add(Format_String(Fiche.CodeJ,3)+IntToStr(i)+St);
     END ;
  END ;

end ;

{=============================================================================}
procedure TFImpCpt.RecupRefAuto ;
Var i : Integer ;
    St : String ;
begin
VideTable('REFAUTO','WHERE RA_JOURNAL<>""') ;
BeginTrans ;
TRefAuto.Open ;
For i:=0 To LesRefJ.Count-1 Do
  BEGIN
  TRefAuto.Insert ;
  St:=LesRefJ.Strings[i] ;
  TRefAuto.FindField('RA_CODE').ASString:='R'+FormatFloat('00',i) ;
  TRefAuto.FindField('RA_LIBELLE').ASString:='Référence automatique' ;
  TRefAuto.FindField('RA_JOURNAL').ASString:=Trim(Copy(St,1,3)) ;
  TRefAuto.FindField('RA_NATUREPIECE').ASString:=NatPiece(StrToInt(Copy(St,4,1))) ;
  TRefAuto.FindField('RA_FORMULEREF').ASString:=Copy(St,5,Length(St)-4) ;
  TRefAuto.FindField('RA_FORMULELIB').ASString:='' ;
  TRefAuto.Post ;
  END ;
TRefAuto.Close ;
CommitTrans ;
end ;

{=============================================================================}
function TrouveTypeContrepartie(Var Fiche : EnregCpta) : String3 ;
Var i,ii : Integer ;
begin
Result:='MAN' ; ii:=0 ;
For i:=1 To 19 Do
  BEGIN
  If Fiche.OptionsPieceJ[i].actif=1 Then
     BEGIN
     ii:=Fiche.OptionsTresoJ[i].EcrContrep ; Break ;
     END ;
  END ;
If ii=1 Then Result:='LIG' Else If ii=2 Then Result:='PID' Else If ii=3 Then Result:='PIS' ;
end ;

{=============================================================================}
procedure TFImpCpt.recupjal ;
Var LaPos,Lg : LongInt ;
    Fiche,Fiche12,Fiche13 : EnregCPTA ;
    NoBanque : Integer ;
    LaBanque,St : String ;
    i : SmallInt ;
    OkOk : Boolean ;
  Label 1 ;
begin
VideTable('BANQUES','') ; VideTable('JOURNAL','') ; VideTable('SOUCHE','') ;
BeginTrans ;
Ouvre(3,0,FALSE,spCPTA) ; Lg:=FileLen(VSAA^.DAtFF[3])-1 ;
Ouvre(15,2,FALSE,spCPTA) ;
InitMove(Lg,'recup des Journaux') ;
AddSouche('S') ; AddSouche('N') ; AddSouche('A') ;
TJournal.open ; NoBanque:=0 ;
For LaPos:=1 To Lg Do
  BEGIN
  if MoveCur(FALSE) or Stopper then goto 1;
  GetRec(VSAA^.DatFF[3],LaPos,Fiche) ;
  OkOk:=(Fiche.Status=0)  ;
  If OkOk Then OkOk:=OkOk And (AnaHal.Checked Or ((Not AnaHal.Checked) And (Fiche.NatureJ<>9))) ;
  If OkOk Then
    BEGIN
    LMess('Création du Journal '+Fiche.CodeJ) ;
    TJournal.Insert ;
    InitNew(TJournal) ;
    TJournal.findfield('J_JOURNAL').asstring:=Trim(Fiche.codeJ) ;
    TJournal.findfield('J_LIBELLE').asstring:=ASCII2ANSI(Trim(Fiche.IntituleJ)) ;
    TJournal.findfield('J_ABREGE').asstring:=ASCII2ANSI(Trim(Fiche.AbregeJ)) ;
    TJournal.findfield('J_NATUREJAL').asstring:='OD' ;
    TJournal.findfield('J_AXE').asstring:='' ;
    Case Fiche.NatureJ Of
      1 : TJournal.findfield('J_NATUREJAL').asstring:='ACH' ;
      2 : TJournal.findfield('J_NATUREJAL').asstring:='VTE' ;
      3 : TJournal.findfield('J_NATUREJAL').asstring:='BQE' ;
      4 : TJournal.findfield('J_NATUREJAL').asstring:='CAI' ;
      5 : TJournal.findfield('J_NATUREJAL').asstring:='OD' ;
      6 : TJournal.findfield('J_NATUREJAL').asstring:='ANO' ;
      7 : TJournal.findfield('J_NATUREJAL').asstring:='CLO' ;
      8 : TJournal.findfield('J_NATUREJAL').asstring:='REG' ;
      9 : BEGIN
          TJournal.findfield('J_NATUREJAL').asstring:='ODA' ;
          TJournal.findfield('J_AXE').asstring:='A1' ;
          END ;
      END ;
    If fiche.CentraliseJ=1 Then TJournal.findfield('J_CENTRALISABLE').asstring:='X'
                           Else TJournal.findfield('J_CENTRALISABLE').asstring:='-' ;
    If fiche.DeviseJ=1 Then TJournal.findfield('J_MULTIDEVISE').asstring:='X'
                       Else TJournal.findfield('J_MULTIDEVISE').asstring:='-' ;
    TJournal.findfield('J_FERME').asstring:='-' ;
    TJournal.findfield('J_DATECREATION').asdatetime:=Sysutils.Date ;
    TJournal.findfield('J_DATEMODIF').asdatetime:=Sysutils.Date ;
    TJournal.findfield('J_DATEOUVERTURE').asdatetime:=Sysutils.Date ;
    TJournal.findfield('J_DATEFERMETURE').asDateTime:=IDate1900 ;
    TJournal.findfield('J_DATEDERNMVT').asDatetime:=IDate1900 ;
    TJournal.findfield('J_DEBITDERNMVT').asFloat:=0 ;
    TJournal.findfield('J_CREDITDERNMVT').asFloat:=0 ;
    TJournal.findfield('J_NUMDERNMVT').asInteger:=0 ;
    AlimRef(Fiche) ;
    If Fiche.NatureJ=9 Then
       BEGIN
       TJournal.findfield('J_COMPTEURNORMAL').asstring:='ANA' ;
       TJournal.findfield('J_COMPTEURSIMUL').asstring:='' ;
       END Else
       BEGIN
       TJournal.findfield('J_COMPTEURNORMAL').asstring:='CPT' ;
       TJournal.findfield('J_COMPTEURSIMUL').asstring:='SIM' ;
       END ;
    TJournal.findfield('J_COMPTEINTERDIT').asstring:=RecupCptJal(LaPos,CpteInt,Fiche) ;
    TJournal.findfield('J_COMPTEAUTOMAT').asstring:=RecupCptJal(LaPos,Ventil,Fiche) ;
    LitBloc(Fiche.PosBlocJ,'3',spCPTA,TJournal.findfield('J_BLOCNOTE')) ;
    St:='------------------------' ;
    For i:=1 To 24 Do If TbExoLegalV8[1][1][i].LValide='O' Then St[i]:='X' ;
    TJournal.findfield('J_VALIDEEN').asstring:=St ;
    St:='------------------------' ;
    For i:=1 To 24 Do If TbExoLegalV8[2][1][i].LValide='O' Then St[i]:='X' ;
    TJournal.findfield('J_VALIDEEN1').asstring:=St ;
    If (Fiche.NatureJ IN [3,4]) Then
       BEGIN
       TJournal.findfield('J_CONTREPARTIE').asstring:=Trim(Fiche.CpteContrepGJ) ;
       TJournal.findfield('J_TYPECONTREPARTIE').asstring:=TrouveTypeContrepartie(Fiche) ;
       If ((Fiche.PosAdresseJ>0) Or (Fiche.PosBAnqueJ>0)) And (Fiche.NatureJ=3) Then
          BEGIN
          Inc(NoBanque) ;
          Tbanque.Open ; Tbanque.Insert ; InitNew(Tbanque) ;
          Tbanque.findfield('PQ_DE_MODE').asString:='TX' ;
          Tbanque.findfield('PQ_PD_MODE').asString:='TX' ;
          Tbanque.findfield('PQ_CR_MODE').asString:='TX' ;
          Tbanque.findfield('PQ_CO_MODE').asString:='TX' ;
          Tbanque.findfield('PQ_DE_TXREF').asString:='TBB' ;
          Tbanque.findfield('PQ_PD_TXREF').asString:='TBB' ;
          Tbanque.findfield('PQ_CR_TXREF').asString:='TBB' ;
          Tbanque.findfield('PQ_CO_TXREF').asString:='TBB' ;
          Tbanque.findfield('PQ_BB_DATETAUX').asDateTime:=Date ;
          Tbanque.findfield('PQ_DE_DATETAUX').asDateTime:=Date ;
          Tbanque.findfield('PQ_PD_DATETAUX').asDateTime:=Date ;
          Tbanque.findfield('PQ_CR_DATETAUX').asDateTime:=Date ;
          Tbanque.findfield('PQ_CO_DATETAUX').asDateTime:=Date ;
          Tbanque.findfield('PQ_REMCB').asString:='MAG' ;
          Tbanque.findfield('PQ_REMCHQ').asString:='MAG' ;
          Tbanque.findfield('PQ_REMESP').asString:='MAG' ;
          Tbanque.findfield('PQ_REMLCR').asString:='MAG' ;
          Tbanque.findfield('PQ_REMPRE').asString:='MAG' ;
          Tbanque.findfield('PQ_REMTRI').asString:='MAG' ;
          Tbanque.findfield('PQ_REMVIR').asString:='MAG' ;
          Tbanque.findfield('PQ_REMVIT').asString:='MAG' ;
          Tbanque.findfield('PQ_REMTIP').asString:='MAG' ;
          Tbanque.findfield('PQ_REMTEP').asString:='MAG' ;
          Tbanque.findfield('PQ_REMVIC').asString:='MAG' ;
          Tbanque.findfield('PQ_RESTCB').asString:='DET' ;
          Tbanque.findfield('PQ_RESTCHQ').asString:='DET' ;
          Tbanque.findfield('PQ_RESTESP').asString:='DET' ;
          Tbanque.findfield('PQ_RESTLCR').asString:='DET' ;
          Tbanque.findfield('PQ_RESTPRE').asString:='DET' ;
          Tbanque.findfield('PQ_RESTTRI').asString:='DET' ;
          Tbanque.findfield('PQ_RESTVIR').asString:='DET' ;
          Tbanque.findfield('PQ_RESTVIT').asString:='DET' ;
          Tbanque.findfield('PQ_RESTTIP').asString:='DET' ;
          Tbanque.findfield('PQ_RESTTEP').asString:='DET' ;
          Tbanque.findfield('PQ_RESTVIC').asString:='DET' ;
          Tbanque.findfield('PQ_BANQUE').asString:=FormatFloat('000',NoBanque) ;
          (*
          TGeneraux.Open ;
          If TGeneraux.FindKey([Trim(Fiche.CpteContrepGJ)]) Then
             BEGIN
             TGeneraux.Edit ;
             TGeneraux.findfield('').asString:= ;
             TGeneraux.Post ;
             END ;
          TGeneraux.Close ;
          *)
          LaBanque:=Trim(Fiche.CpteContrepGJ) ;
          TbanqueCP.Open ;
          If TBanqueCP.Findkey([LaBanque]) Then TBanqueCP.Edit Else
             BEGIN
             TbanqueCP.Insert ; InitNew(TbanqueCP) ;
             TbanqueCP.findfield('BQ_GENERAL').asString:=Trim(Fiche.CpteContrepGJ) ;
             TbanqueCP.findfield('BQ_CODE').asString:=Trim(Fiche.CpteContrepGJ) ;
             TbanqueCP.findfield('BQ_DEVISE').asString:=TbDev[1].ISO ;
             END ;
          If Fiche.PosAdresseJ>0 then
             BEGIN
             Ouvre(12,0,FALSE,spCPTA) ;
             GetRec(VSAA^.Datff[12],Fiche.PosAdresseJ,Fiche12) ;
             GAccess.Ferme(12,0,FALSE) ;
             TbanqueCP.findfield('BQ_ADRESSE1').asString:=Fiche12.Adresse1GT;
             TbanqueCP.findfield('BQ_ADRESSE2').asString:=Fiche12.Adresse2GT;
             TbanqueCP.findfield('BQ_ADRESSE3').asString:=Fiche12.Adresse3GT;
             TbanqueCP.findfield('BQ_CODEPOSTAL').asString:=Fiche12.CodePostalGT;
             TbanqueCP.findfield('BQ_VILLE').asString:=Fiche12.VilleGT; ;
             TbanqueCP.findfield('BQ_PAYS').asString:='FRA' ;
             END ;
          if Fiche.PosBanqueJ>0 then
             BEGIN
             Ouvre(13,0,FALSE,spCPTA) ;
             GetRec(VSAA^.Datff[13],Fiche.PosBanqueJ,Fiche13) ;
             GAccess.Ferme(13,0,FALSE) ;
             LaBanque:=Fiche13.BanqueB ;
             TbanqueCP.findfield('BQ_LIBELLE').asString:=Fiche13.BanqueB ;
             TbanqueCP.findfield('BQ_ETABBQ').asString:=Copy(Fiche13.CodeBAnqueB,1,5) ;
             TbanqueCP.findfield('BQ_GUICHET').asString:=Copy(Fiche13.CodeGuichetB,1,5) ;
             TbanqueCP.findfield('BQ_NUMEROCOMPTE').asString:=Copy(Fiche13.CompteBanqueB,1,11) ;
             TbanqueCP.findfield('BQ_CLERIB').asString:=Copy(Fiche13.CleRibB,1,2) ;
             END ;
          TbanqueCP.findfield('BQ_BANQUE').asString:=FormatFloat('000',NoBanque) ;
          Tbanque.findfield('PQ_LIBELLE').asString:=LaBanque ;
          Tbanque.post ;
          Tbanque.close ;
          TbanqueCP.post ;
          TbanqueCP.close ;
          END ;
       END ;
    TJournal.post ;
    END ;
  END ;
1:TRefAuto.Close ; TJournal.Close ;
GAccess.Ferme(15,2,FALSE) ;
GAccess.Ferme(3,0,FALSE) ;
FiniMove ;
CommitTrans ;
RecupRefAuto ;
end ;

{=============================================================================}
procedure TFImpCpt.recupVentilAnal;
Var LaPos,Lg : LongInt ;
    Fiche : EnregCPTA ;
  Label 1 ;
begin
VideTable('VENTIL','') ;
BeginTrans ;
Ouvre(14,0,FALSE,spCPTA) ; Lg:=FileLen(VSAA^.DAtFF[14])-1 ;
InitMove(Lg,'recup des ventilations analytiques') ;
TVentil.open ;
For LaPos:=1 To Lg Do
  BEGIN
  if MoveCur(FALSE) or Stopper then goto 1;
  GetRec(VSAA^.DatFF[14],LaPos,Fiche) ;
  If Fiche.Status=0 Then
    BEGIN
   If Trim(Fiche.CompteGeneVA)<>'Ventil Type' Then
       BEGIN
       LMess('Création de la ventilation '+Fiche.CompteGeneVA+'/'+Fiche.CompteAnalVA) ;
       TVentil.Insert ;
       TVentil.findfield('V_COMPTE').asstring:=Trim(Fiche.compteGeneVA) ;
       TVentil.findfield('V_NUMEROVENTIL').asInteger:=Fiche.NoOrdreVA ;
       TVentil.findfield('V_NATURE').asstring:='GE1' ;
       TVentil.findfield('V_SECTION').asstring:=Trim(Fiche.compteAnalVA) ;
       TVentil.findfield('V_TAUXMONTANT').asFloat:=Arrondi(Fiche.PourCentVA,ADecimP) ;
       TVentil.findfield('V_TAUXQTE1').asFloat:=0 ;
       TVentil.findfield('V_TAUXQTE2').asFloat:=0 ;
       TVentil.post ;
       END ;
    END ;
  END ;
1:TVentil.Close ;
GAccess.Ferme(14,0,FALSE) ;
FiniMove ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupLesVentilType ;
Var i : SmallInt ;
    Fiche14 : EnregCpta ;
    Pos14,kk,LaPos : LongInt ;
    Cle,Test : String40 ;
    OkOk : Boolean ;
  Label 1 ;
begin
BeginTrans ;
TVentil.open ;
Ouvre(14,2,FALSE,spCPTA) ; Ouvre(2,2,FALSE,spCPTA) ;
InitMove(99,'recup des ventilations Types') ;
For i:=1 To 99 Do
  BEGIN
  LaPos:=1000000+i ;
  if MoveCur(FALSE) or Stopper then goto 1;
  If TbRupt[i]<>'' Then
     BEGIN
     Cle:=CreCle14(LaPos,0) ; kk:=1 ; Test:=CreCle14(LaPos,0) ;
     Repeat
       If kk=1 then SearchKey(VSAA^.Idxff[14],Pos14,Cle) else NextKey(VSAA^.Idxff[14],Pos14,Cle) ;
       Inc(kk) ; OkOk:=((VSAA^.Ok) and (Copy(Cle,1,Length(Test))=Test)) ;
       if OkOk then
          BEGIN
          GetRec(VSAA^.Datff[14],Pos14,Fiche14) ;
          TVentil.Insert ;
          LMess('Création de la Grille N° '+IntToStr(i)+'/'+Fiche14.CompteAnalVA) ;
          TVentil.findfield('V_COMPTE').asstring:=FormatFloat('000',i) ; ;
          TVentil.findfield('V_NUMEROVENTIL').asInteger:=Fiche14.NoOrdreVA ;
          TVentil.findfield('V_NATURE').asstring:='TY1' ;
          TVentil.findfield('V_SECTION').asstring:=Trim(Fiche14.compteAnalVA) ;
          TVentil.findfield('V_TAUXMONTANT').asFloat:=Arrondi(Fiche14.PourCentVA,ADecimP) ;
          TVentil.findfield('V_TAUXQTE1').asFloat:=0 ;
          TVentil.findfield('V_TAUXQTE2').asFloat:=0 ;
          TVentil.post ;
          END ;
     Until Not OkOk ;
     END ;
  END ;
1:TVentil.Close ;
GAccess.Ferme(14,2,FALSE) ; GAccess.Ferme(2,2,FALSE) ;
FiniMove ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.recupVentilType;
begin
LitRuptV8('V') ; RecupCodeRupt('VTY') ; RecupLesVentilType ;
end ;

{============================================================================}
PROCEDURE LITFICHIER ( OkOuvre,Bloc,OkFerme : boolean ) ;

Const OffsetLecture = 1000 ;
      ReelMax = 279 ; { Penser … reporter cette valeur dans RecupLeParasoc }
Type RecTab = Array[1..ReelMax] of real48 ;

Var i : integer ;
    Tab : RecTab ;
    Xp : Real48 ;
BEGIN
FillChar(VSAA^.ParaSoc,SizeOf(VSAA^.ParaSoc),#0) ;
if OkOuvre then OuvreNum ;
Seek(VSAA^.fichierNum,OffsetLecture) ;
For i:=1 to ReelMax do Read(VSAA^.FichierNum,Tab[i]) ;
Move(Tab,VSAA^.ParaSoc,SizeOf(VSAA^.ParaSoc)) ;
Seek(VSAA^.FichierNum,12); read(VSAA^.FichierNum,xp) ;
if ((Xp<=0) or (Xp>20)) then Xp:=1 ; VSAA^.CurExercice:=Round(xp) ;
seek(VSAA^.FichierNum,1) ; read(VSAA^.FichierNum,xp) ; VSAA^.LastDate:=round(xp) ;
Seek(VSAA^.FichierNum,24) ; read(VSAA^.FichierNum,xp) ; VSAA^.OkDecV:=round(Xp) ;
seek(VSAA^.FichierNum,28) ; read(VSAA^.FichierNum,xp) ; VSAA^.CloturePer   :=round(xp) ;
seek(VSAA^.FichierNum,32) ; read(VSAA^.FichierNum,xp) ; VSAA^.CloturePro   :=round(xp) ;
seek(VSAA^.FichierNum,46) ; read(VSAA^.FichierNum,xp) ; VSAA^.ClotureTva   :=round(xp) ;
if OkFerme then FermeNum ;
END ;


(*
{=============================================================================}
Function trouveexo(DD : Word ; Var NbExoCree : Integer) : String ;

Var TabDate : Array[1..20] Of TFourDate ;

BEGIN
If DD
While ParaSoc.DAteExo[i].Fin>40000 Do
   BEGIN
   TabDate[i].Deb:=ParaSoc.DateExo[i].Deb ;
   TabDate[i].Fin:=ParaSoc.DateExo[i].Fin ;
   Inc(i) ;
If Not PourPurge then
   BEGIN
   TabDate[i].Deb:=ParaSoc.EnCours.Deb ;
   TabDate[i].Fin:=ParaSoc.EnCours.Fin ;
   AlimPointeurZoom(ParaSoc.EnCours.Deb,ParaSoc.EnCours.Fin,First,Old,Courant,i) ;
   TabDate[i].Deb:=ParaSoc.Suivant.Deb ;
   TabDate[i].Fin:=ParaSoc.Suivant.Fin ;
   AlimPointeurZoom(ParaSoc.Suivant.Deb,ParaSoc.Suivant.Fin,First,Old,Courant,i) ;
   END ;
END ;
*)
{=============================================================================}
procedure TFImpCpt.addexo(i : Integer ; Deb,Fin : Word ; LeEtat : String ; ExoEnCours : Boolean) ;
Var ii,j : SmallInt ;
    St : ShortString ;
BEGIN
TExo.Insert ;
InitNew(TExo) ;
TExo.findfield('EX_EXERCICE').asstring:=FormatFloat('000',i) ;
TExo.findfield('EX_LIBELLE').asstring:='Exo de '+DateToStr(Int2Date(Deb))+
                                      ' au '+DateToStr(Int2Date(Fin)) ;
TExo.findfield('EX_ABREGE').asstring:='Exercice de '+DateToStr(Int2Date(Deb))+
                                      ' au '+DateToStr(Int2Date(Fin)) ;
TExo.findfield('EX_DATEDEBUT').asDateTime:=Int2Date(Deb)  ;
TExo.findfield('EX_DATEFIN').asDateTime:=Int2Date(Fin)  ;
TExo.findfield('EX_ETATCPTA').asstring:=LeEtat ;
If (LeEtat='OUV') Or (LeEtat='CPR') Then
   BEGIN
   ii:=2 ; If ExoEnCours Then ii:=1 ;
   St:='------------------------' ;
   For j:=1 To 24 Do If TbExoLegalV8[ii][1][j].LValide='O' Then St[j]:='X' ;
   TExo.findfield('EX_VALIDEE').asstring:=St ;
   END ;
Texo.Post ;

END ;

{=============================================================================}
Procedure TFImpCpt.RecupEdtLegal ;
Var i,j,k : SmallInt ;
    Deb : TDateTime ;
    St1 : ShortString ;
begin
VideTable('EDTLEGAL','WHERE ED_OBLIGATOIRE="X"') ;
TEdt.Open ; i:=1 ;
For i:=1 TO 2 Do // N N+1
  BEGIN
  If i=1 Then Deb:=Int2Date(VSAA^.ParaSoc.EnCours.Deb) Else Deb:=Int2Date(VSAA^.ParaSoc.Suivant.Deb) ;
  For j:=1 To 3 Do // Jal, GL Gen, GL aux
    For k:=1 To 24 Do
      BEGIN
      If TbExoLegalV8[i][j][k].LFait='O' Then
         BEGIN
         TEdt.Insert ;
         InitNew(TEdt) ;
         TEdt.findfield('ED_OBLIGATOIRE').asstring:='X';
         If j=1 Then St1:='JLD' Else If j=2 Then ST1:='GLG' Else St1:='GLT' ;
         TEdt.findfield('ED_TYPEEDITION').asstring:=St1;
         TEdt.findfield('ED_EXERCICE').asstring:=FormatFloat('000',VSAA^.CurExercice+i-1) ;
         TEdt.findfield('ED_PERIODE').asDateTime:=PlusMois(Deb,k-1);
         (*
         TEdt.findfield('ED_DATE1').asDateTime:=;
         TEdt.findfield('ED_DATE2').asDateTime:=;
         *)
         TEdt.findfield('ED_DATEEDITION').asDateTime:=Int2Date(TbExoLegalV8[i][j][k].LDate) ;
         TEdt.findfield('ED_NUMEROEDITION').asInteger:=TbExoLegalV8[i][j][k].LNoEdition;
         (*
         TEdt.findfield('ED_CUMULDEBIT').asFloat:=;
         TEdt.findfield('ED_CUMULCREDIT').asFloat:=;
         TEdt.findfield('ED_CUMULSOLDED').asFloat:=;
         TEdt.findfield('ED_CUMULSOLDEC').asFloat:=;
         TEdt.findfield('ED_UTILISATEUR').asstring:=;
         TEdt.findfield('ED_DESTINATION').asstring:=;
         *)
         TEdt.Post ;
         END ;
      END ;
  END ;
TEdt.Close ;
end ;

{=============================================================================}
Procedure TFImpCpt.RecupExo ;
Var i : SmallInt ;
    EtatExo : String3 ;
    LaSoc : String ;
begin
VideTable('EXERCICE','') ;
TExo.Open ; i:=1 ;
While VSAA^.ParaSoc.DAteExo[i].Fin>40000 Do
   BEGIN
   AddExo(i,VSAA^.ParaSoc.DateExo[i].Deb,VSAA^.ParaSoc.DateExo[i].Fin,'CDE',FALSE) ; Inc(i) ;
   END ;
EtatExo:='OUV' ;
If VSAA^.CloturePro=VSAA^.ParaSoc.EnCours.Fin Then EtatExo:='CPR' ;
AddExo(i,VSAA^.ParaSoc.EnCours.Deb,VSAA^.ParaSoc.EnCours.Fin,EtatExo,TRUE) ;
LaSoc:=Trim(Copy(VSAA^.ParaSoc.CodeSociete,1,3)) ;
If LaSoc='' Then LaSoc:='001' ;
TSoc.open ;
If TSoc.FindKey([LaSoc]) Then
   BEGIN
{$IFDEF SPEC302}
   TSoc.Edit ;
   TSoc.findfield('SO_EXOV8').asString:=FormatFloat('000',i) ;
   TSoc.Post ;
{$ELSE}
   SetparamSoc('SO_EXOV8',FormatFloat('000',i)) ;
{$ENDIF}
   END ;
TSoc.Close ;
Inc(i) ; AddExo(i,VSAA^.ParaSoc.Suivant.Deb,VSAA^.ParaSoc.Suivant.Fin,'OUV',FALSE) ;
TExo.Close ;
end ;

{=============================================================================}
Function TFImpCpt.TrouveExo(dd : Word ; Var D2 : TDateTime) : String ;
Var i : Integer ;
Begin
D2:=0 ;
i:=1 ; TrouveExo:=FormatFloat('000',VSAA^.CurExercice) ;
If (dd>=VSAA^.parasoc.EnCours.Deb) And
   (dd<=VSAA^.parasoc.EnCours.Fin) Then
   BEGIN
   TrouveExo:=FormatFloat('000',VSAA^.CurExercice) ; D2:=Int2Date(VSAA^.parasoc.EnCours.Fin) ;
   END Else
  If (dd>=VSAA^.parasoc.Suivant.Deb) And
     (dd<=VSAA^.parasoc.Suivant.Fin) Then
     BEGIN
     TrouveExo:=FormatFloat('000',VSAA^.CurExercice+1) ; D2:=Int2Date(VSAA^.parasoc.Suivant.Fin) ;
     END Else
     BEGIN
     While VSAA^.ParaSoc.DAteExo[i].Fin>40000 Do
        BEGIN
        If (dd>=VSAA^.parasoc.DAteExo[i].Deb) And
           (dd<=VSAA^.parasoc.DAteExo[i].Fin) Then
           BEGIN
           TrouveExo:=FormatFloat('000',i) ; D2:=Int2Date(VSAA^.parasoc.DAteExo[i].Fin) ;
           Exit ;
           END Else Inc(i) ;
        END ;
     END ;

End ;

{=============================================================================}
procedure TFImpCpt.AlimRel ;
Var i : Integer ;
    St13 : String13 ;
BEGIN
Fillchar(TbRelV8,SizeOf(TbRelV8),#0) ;
OuvreChoixCod ;
Seek(VSAA^.FichierChCh,7530) ; i:=1 ;
repeat
  read(VSAA^.FichierChCh,St13) ;
  If St13<>'' then BEGIN TbRelV8[i]:=uppercase(ASCII2ANSI(St13)) ; Inc(i) ; END ;
Until ((St13='') or (i>10)) ;
FermeChoixCod ;

END ;

{=============================================================================}
procedure TFImpCpt.AlimTva ;
Var i,ii : Integer ;
    St13 : String13 ;
BEGIN
Fillchar(TbReg,SizeOf(TbReg),#0) ; Fillchar(TbLibReg,SizeOf(TbLibReg),#0) ;
OuvreChoixCod ;
Seek(VSAA^.FichierChCh,1910) ; i:=1 ;
repeat
  read(VSAA^.FichierChCh,St13) ;
  If St13<>'' then BEGIN TbReg[i]:=uppercase(ASCII2ANSI(St13)) ; TbLibReg[i]:=TbReg[i] ; Inc(i) ; END ;
Until ((St13='') or (i>19)) ;

FermeChoixCod ; ii:=0 ;
For i:=1 To 19 Do
  BEGIN
  If TbReg[i]='FRANCE' Then  TbReg[i]:='FRA' Else
     If TbReg[i]='EXPORT' Then  TbReg[i]:='EXP' Else
        If (TbReg[i]='EXONERE') Or (TbReg[i]='EXONéRé') Then TbReg[i]:='EXO' Else
           If TbReg[i]='DOM-TOM' Then  TbReg[i]:='DOM' Else
              If TbReg[i]='CORSE' Then  TbReg[i]:='COR' Else
                 BEGIN
                 If TbReg[i]<>'' Then
                    BEGIN
                    Inc(ii) ; TbReg[i]:=FormatFloat('000',ii) ;
                    END ;
                 END ;
  END ;
END ;

{=============================================================================}
Function FApartirDe(b : Byte) : String3 ;
begin
Result:='ECR' ;
Case b Of 2 : Result:='DEB' ; 3 : Result:='FIN' ; END ;
end ;

{=============================================================================}
Function FArrondiJour(b : Byte) : String3 ;
begin
Result:='PAS' ;
Case b Of 5 : Result:='05M' ; 10 : Result:='10M' ; 15 : Result:='15M' ;
          20 : Result:='20M' ; 25 : Result:='25M' ; 30,31 : Result:='FIN' ;
  END ;
end ;

{=============================================================================}
Function FSeparePar(b : Byte) : String3 ;
begin
Result:='QUI' ;
Case b Of  2 : Result:='1M' ; 3 : Result:='2M' ; 4 : Result:='3M' ; END ;
end ;

{=============================================================================}
Function LeModePaie(b : Byte) : String ;
BEGIN
Result:='DIV' ;
Case b Of
  1 : Result:='ESP' ; 2 : Result:='CHQ' ; 3 : Result:='TRD' ; 4 : Result:='TRA' ;
  5 : Result:='CTB' ; 6 : Result:='LCR' ; 7 : Result:='BOR' ; 8 : Result:='VIR' ;
  9 : Result:='PRL' ;
  END ;
END ;

{=============================================================================}
procedure TFImpCpt.LitModRV8 ;
Var i,j : Integer ;
    St13 : String13 ;
    NbMp : SmallInt ;
    CMP  : CalcModePaie ;
BEGIN
Fillchar(TbMP,SizeOf(TbMP),#0) ;
OuvreChoixCod ;
Seek(VSAA^.FichierChCh,3510) ; i:=1 ; NbMp:=0 ;
repeat
  read(VSAA^.FichierChCh,St13) ;
  If St13<>'' then
    BEGIN
    TbMP[i].Abrege:=uppercase(St13) ; TbMP[i].Code:=FormatFloat('000',i) ; Inc(i) ; Inc(NbMP) ;
    END ;
Until ((St13='') or (i>50)) ;
FermeChoixCod ;
Ouvre_Regle ;
For i:=1 To NbMP Do
  BEGIN
  ChargeCalcMode(i,CMP) ;
  TbMP[i].Libelle        :=CMP.Titre ;
  TbMP[i].APartirde      :=FAPartirDe(CMP.Origine) ;
  TbMP[i].PlusJour       :=CMP.NbJours ;
  TbMP[i].ArrondiJour    :=FArrondiJour(CMP.Arrivee) ;
  TbMP[i].NombreEcheance :=CMP.NbEche ;
  TbMP[i].SeparePar      :=FSeparePar(CMP.Decal) ;
  TbMP[i].MontantMin     :=CMP.Montant1 ;
  TbMP[i].RemplaceMin    :=TbMP[CMP.Remp1].Code ;
  For j:=1 To TbMP[i].NombreEcheance Do
    BEGIN
    TbMP[i].TauxMP[j].MPA:=LeModePaie(CMP.Paiement[j]) ;
    TbMP[i].TauxMP[j].Taux:=CMP.Pourcent[j] ;
    END ;
  END ;
Ferme_Regle ;
END ;

{=============================================================================}
procedure TFImpCpt.MajTableRelance ;
Var i : Integer ;
begin
VideTable('RELANCE','WHERE (RR_TYPERELANCE="RRG") AND (RR_FAMILLERELANCE<>"001") AND (RR_FAMILLERELANCE<>"002")') ;
BeginTrans ;
TRelance.open ;
For i:=1 To 10 Do If TbRelV8[i]<>'' Then
    BEGIN
    TRelance.Insert ;
    InitNew(TRelance) ;
    TRelance.Findfield('RR_TYPERELANCE').AsString:='RRG' ;
    TRelance.Findfield('RR_FAMILLERELANCE').AsString:=FormatFloat('000',i+2) ;
    TRelance.Findfield('RR_LIBELLE').AsString:=TbRelV8[i] ;
    TRelance.Post ;
    END ;
TRelance.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.MajTableModR ;
Var i,j : Integer ;
begin
VideTable('MODEREGL','') ;
BeginTrans ;
TMODR.open ;
For i:=1 To 50 Do If TbMP[i].Code<>'' Then
    BEGIN
    TMODR.Insert ;
    InitNew(TMODR) ;
    TMODR.Findfield('MR_MODEREGLE').AsString      :=TbMP[i].Code ;
    TMODR.Findfield('MR_LIBELLE').AsString        :=TbMP[i].Libelle ;
    TMODR.Findfield('MR_ABREGE').AsString         :=TbMP[i].Abrege ;
    TMODR.Findfield('MR_APARTIRDE').AsString      :=TbMP[i].APArtirDe ;
    TMODR.Findfield('MR_PLUSJOUR').AsInteger      :=TbMP[i].PlusJour ;
    TMODR.Findfield('MR_ARRONDIJOUR').AsString    :=TbMP[i].ArrondiJour ;
    TMODR.Findfield('MR_NOMBREECHEANCE').AsInteger:=TbMP[i].NombreEcheance ;
    TMODR.Findfield('MR_SEPAREPAR').AsString      :=TbMP[i].SeparePar ;
    TMODR.Findfield('MR_MONTANTMIN').AsFloat      :=TbMP[i].MontantMin ;
    TMODR.Findfield('MR_REMPLACEMIN').AsString    :=TbMP[i].RemplaceMin ;
    TMODR.Findfield('MR_MODEGUIDE').AsString      :='-' ;
    TMODR.Findfield('MR_REPARTECHE').AsString     :='' ;
    For j:=1 To TbMP[i].NombreEcheance Do
      BEGIN
      TMODR.Findfield('MR_MP'+IntToStr(j)).AsString :=TbMP[i].TauxMP[j].MPA ;
      TMODR.Findfield('MR_TAUX'+IntToStr(j)).AsFloat:=TbMP[i].TauxMP[j].Taux ;
      END ;
    TMODR.Findfield('MR_SOCIETE').AsString:='' ;
    TMODR.Post ;
    END ;
TMODR.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.MajTableModP ;
Var i : Integer ;
begin
Exit ;
VideTable('MODEPAIE','') ;
BeginTrans ;
TMODP.open ;
For i:=1 To 50 Do If TbMP[i].Code<>'' Then
    BEGIN
    TMODP.Insert ;
    InitNew(TMODP) ;
    TMODP.Findfield('MP_MODEPAIE').AsString      :=TbMP[i].Code ;
    TMODP.Findfield('MP_LIBELLE').AsString        :=TbMP[i].Libelle ;
    TMODP.Findfield('MP_ABREGE').AsString         :=TbMP[i].Abrege ;
    TMODP.Findfield('MR_SOCIETE').AsString:='' ;
    TMODP.Post ;
    END ;
TMODP.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupModR ;

begin
LitModRV8 ; MajTableModR ; MajTableModP ;
end ;

{=============================================================================}
Function TFImpCpt.QuelTypeMvtV8(Cpt : String17 ; NatCpt : String3 ; NatPiece : Byte) : String3 ;
BEGIN
Result:='DIV' ;
If NatPiece In [1,2,4,5]=FALSE Then Exit ;
If (NatCpt='CHA') Or (NatCpt='PRO') Then BEGIN Result:='HT' ; Exit ; END ;
If (NatCpt='CLI') Or (NatCpt='FOU') Or (NatCpt='AUD') Or (NatCpt='AUC') Or
   (NatCpt='TID') Or (NatCpt='TIC') Then BEGIN Result:='TTC' ; Exit ; END ;
If IsTvaV8(Cpt) Then Result:='TVA' ;
END ;

{=============================================================================}
procedure TFImpCpt.MajEEXBQE(Cpt,Ref : String ; DD : TDateTime ) ;
BEGIN
If Not TEXBQ.FindKey([Cpt,DD,Ref]) Then
   BEGIN
   TEXBQ.Insert ;
   InitNew(TEXBQ) ;
   TEXBQ.Findfield('EE_GENERAL').AsString:=Cpt ;
   TEXBQ.Findfield('EE_REFPOINTAGE').AsString:=Ref ;
   TEXBQ.Findfield('EE_DATEPOINTAGE').AsDateTime:=DD ;
   TEXBQ.Post ;
   END ;
END ;

{=============================================================================}
procedure TFImpCpt.recupTroncEcriture(Var Fiche,Fiche6 : EnregCPTA ; Var NoLigne : Integer ; Var TypeMvt : String3) ;
Var Compte : EnregCPTA ;
    Letat : String10 ;
    Pointable,Treso,JalTreso : Boolean ;
    LaDevise : Byte ;
    NatCpt : String3 ;
    CptTreso : String17 ;
    NatCtrTreso : Byte ;
    DateFinExo : TDateTime ;
    XX : Double ;
BEGIN
LaDevise:=Fiche6.DeviseET ;
P_ExerciceMvt:=TrouveExo(Fiche.DateComptaE,DateFinExo) ;
TEcrGen.findfield('E_EXERCICE').asstring:=P_ExerciceMvt ;
TEcrGen.findfield('E_JOURNAL').asstring:=Trim(Fiche.CodeJournalE) ;
TEcrGen.findfield('E_DATECOMPTABLE').asDateTime:=Int2date(Fiche.DateComptaE) ;
TEcrGen.findfield('E_NUMEROPIECE').asInteger:=Fiche.NoPieceE ;
TEcrGen.findfield('E_NUMLIGNE').asInteger:=NoLigne ;
TEcrGen.findfield('E_NUMECHE').asInteger:=0 ;
TEcrGen.findfield('E_ECHE').asstring:='-' ;
TEcrGen.findfield('E_ANA').asstring:='-' ;
If (Fiche.AnalytiqueONE=1) And ANAHal.Checked Then TEcrGen.findfield('E_ANA').asstring:='X' ;
JalTreso:=FALSE ; CptTreso:='' ;
NatCtrTreso:=0 ;
If Fiche.PosJalE>0 Then
   BEGIN
   Getrec(VSAA^.DatFF[3],Fiche.PosJalE,Compte) ;
   JalTreso:=(Compte.NatureJ in [3,4]) And (Fiche6.PosCtrPtrGET>0) ;
   If JalTreso Then
      BEGIN
      If Fiche6.PosCtrPtrGET>0 Then
         BEGIN
         Getrec(VSAA^.DatFF[0],Fiche6.PosCtrPtrGET,Compte) ;
         CptTreso:=Trim(Compte.CompteG) ;
         END ;
      NatCtrTreso:=Fiche6.TypCtrET ;
      END ;
   END ;
P_CptGen:='' ; P_CptAux:='' ; Pointable:=False ; Treso:=FALSE ;  NatCpt:='' ;
If Fiche.PosCpteGeneE>0 Then
   BEGIN
   Getrec(VSAA^.DatFF[0],Fiche.PosCpteGeneE,Compte) ;
   Pointable:=Compte.PointableG=1 ; Treso:=Compte.NatureG in [5,6] ;
   If (Compte.NatureG=1) And  (Compte.LettrableG=1) Then NatCpt:='TID' ;
   If (Compte.NatureG=2) And  (Compte.LettrableG=1) Then NatCpt:='TIC' ;
   If (Compte.NatureG=7) Then NatCpt:='CHA' ;
   If (Compte.NatureG=8) Then NatCpt:='PRO' ;
   END ;
P_CptGen:=Trim(Compte.CompteG) ;
TEcrGen.findfield('E_GENERAL').asstring:=P_CptGen ;
TEcrGen.findfield('E_AUXILIAIRE').asstring:='';
If Fiche.PosCpteAuxE>0 Then
   BEGIN
   Getrec(VSAA^.DatFF[1],Fiche.PosCpteAuxE,Compte) ;
   TEcrGen.findfield('E_AUXILIAIRE').asstring:=Trim(Compte.AuxiliaireT);
   P_CptAux:=Trim(Compte.AuxiliaireT) ;
   If Compte.NatureT=1 Then NatCpt:='CLI' ;
   If Compte.NatureT=2 Then NatCpt:='FOU' ;
   END ;
TEcrGen.findfield('E_DEBIT').asFloat:=0 ;
TEcrGen.findfield('E_CREDIT').asFloat:=0 ;
TEcrGen.findfield('E_DEBITDEV').asFloat:=0 ;
TEcrGen.findfield('E_CREDITDEV').asFloat:=0 ;
XX:=Arrondi(Fiche.MontantFrcsE,Dc) ;
If Fiche.SensE=1 Then
   BEGIN
   TEcrGen.findfield('E_DEBIT').asFloat:=XX ;
   END Else
   BEGIN
   TEcrGen.findfield('E_CREDIT').asFloat:=XX ;
   END ;
TEcrGen.findfield('E_DATETAUXDEV').asDateTime:=Int2Date(Fiche.DateComptaE) ;
If Fiche6.DeviseET<=1 Then
   BEGIN
   TEcrGen.findfield('E_DEVISE').asstring:=TbDev[1].ISO ;
   TEcrGen.findfield('E_TAUXDEV').asFloat:=1 ;
   If FIche.SensE=1 Then TEcrGen.findfield('E_DEBITDEV').asFloat:=XX
                    Else TEcrGen.findfield('E_CREDITDEV').asFloat:=XX ;
   END Else
   BEGIN
   TEcrGen.findfield('E_DEVISE').asstring:=TbDev[LaDevise].ISO ;
   TEcrGen.findfield('E_TAUXDEV').asFloat:=Arrondi(Fiche6.TauxDeviseET,TbDev[LaDevise].DeciT) ;
   If FIche.SensE=1 Then TEcrGen.findfield('E_DEBITDEV').asFloat:=Arrondi(Fiche.MtantDeviseE,TBDev[LaDevise].Deci)
                    Else TEcrGen.findfield('E_CREDITDEV').asFloat:=Arrondi(Fiche.MTantDeviseE,TBDev[LaDevise].Deci) ;
   END ;
TEcrGen.findfield('E_REFINTERNE').asstring:=ASCII2ANSI(Trim(Fiche.referenceE)) ;
TEcrGen.findfield('E_LIBELLE').asstring:=ASCII2ANSI(Trim(Fiche.LibelleE)) ;
TEcrGen.findfield('E_NATUREPIECE').asstring:=NatPiece(Fiche.TypePieceE) ;
TEcrGen.findfield('E_QUALIFPIECE').asstring:='N' ;
If Fiche6.TypeET=EcrSim Then TEcrGen.findfield('E_QUALIFPIECE').asstring:='S' ;
TypeMvt:=QuelTypeMvtV8(P_CptGen,NatCpt,Fiche6.TypePieceET) ;
TEcrGen.findfield('E_TYPEMVT').asstring:=TypeMvt ;
TEcrGen.findfield('E_VALIDE').asstring:='-' ;
If GetEtat(Fiche.EtatE,Validee) Then TEcrGen.findfield('E_VALIDE').asstring:='X' ;
Letat:='0000000000' ;
TEcrGen.findfield('E_ECRANOUVEAU').asstring:='N' ;
If GetEtat(Fiche.EtatE,ANouveau) Then
   begin
   TEcrGen.findfield('E_ECRANOUVEAU').asstring:='H' ;
   end ;
(*
If GetEtat(Fiche.EtatE,Exportee) Then Letat[3]:='X' ;
*)
TEcrGen.findfield('E_ETAT').asstring:=Letat ;
TEcrGen.findfield('E_REFEXTERNE').asstring:='' ;
TEcrGen.findfield('E_DATEREFEXTERNE').asDateTime:=IDate1900 ;
TEcrGen.findfield('E_CONTROLEUR').asstring:='' ;
TEcrGen.findfield('E_DATECREATION').asdateTime:=Int2DAte(Fiche.DateComptaE) ;
TEcrGen.findfield('E_DATEMODIF').asdateTime:=Int2DAte(Fiche.DateComptaE) ;
TEcrGen.findfield('E_QUALIFQTE1').asstring:='AUC' ;
TEcrGen.findfield('E_QUALIFQTE2').asstring:='AUC' ;
{ A FAIRE ???
TEcrGen.findfield('E_SOCIETE').asstring:='' ;
TEcrGen.findfield('E_BLOCNOTE').asstring:='' ;
TEcrGen.findfield('E_VISION').asstring:='' ;
TEcrGen.findfield('E_REFLIBRE').asstring:='' ;
TEcrGen.findfield('E_AFFAIRE').asstring:='' ;
TEcrGen.findfield('E_ENCAISSEMENT').asstring:='' ;
TEcrGen.findfield('E_').asstring:='' ;
}
TEcrGen.findfield('E_ETABLISSEMENT').asstring:=VH^.EtablisDefaut ;
If (Fiche.RegimeE>0) And (Fiche.RegimeE<11)
   Then TEcrGen.findfield('E_REGIMETVA').asstring:=TbReg[Fiche.RegimeE]
   Else TEcrGen.findfield('E_REGIMETVA').asstring:='FRA' ;
TEcrGen.findfield('E_ETATLETTRAGE').asstring:='RI' ;
TEcrGen.findfield('E_REFPOINTAGE').asstring:='' ;
If GetEtat(Fiche.EtatE,Pointee) then
   begin
   If Trim(Fiche.RefPointageE)='' Then Fiche.RefPointageE:='Pointee' ;
   TEcrGen.findfield('E_REFPOINTAGE').asstring:=ASCII2ANSI(Trim(Fiche.RefPointageE)) ;
   TEcrGen.findfield('E_DATEPOINTAGE').asDateTime:=DateFinExo ;
   MajEEXBQE(P_CptGen,ASCII2ANSI(Trim(Fiche.RefPointageE)),DateFinExo) ;
   end ;
If Pointable And Treso Then
   BEGIN
   TEcrGen.findfield('E_DATEECHEANCE').asDateTime:=Int2Date(Fiche.DateComptaE) ;
   TEcrGen.findfield('E_MODEPAIE').asstring:=LeModePaie(Byte(Fiche.PremZoneE[1])) ;
   TEcrGen.findfield('E_NUMECHE').asInteger:=1 ;
   TEcrGen.findfield('E_ECHE').asstring:='X' ;
   END ;
If JalTreso Then
   BEGIN
   If (P_CptGen<>CptTreso) Then TEcrGen.findfield('E_CONTREPARTIEGEN').asstring:=CptTreso Else
      BEGIN
      Case natCtrTreso Of
        1 : ;
        2,3 : ;
        END ;
      END ;
   END ;
{
TEcrGen.findfield('E_TVA').asstring:='' ;
TEcrGen.findfield('E_TPF').asstring:='' ;
TEcrGen.findfield('E_NUMEROIMMO').asstring:='' ;
TEcrGen.findfield('E_BUDGET').asstring:='' ;
TEcrGen.findfield('E_CONTREPARTIEGEN').asstring:='' ;
TEcrGen.findfield('E_CONTREPARTIEAUX').asstring:='' ; asvariant
}
END ;

{=============================================================================}
procedure TFImpCpt.AjusteLettrage ;
BEGIN
If LetHal.Checked Then Exit ;
TEcrGen.findfield('E_COUVERTURE').asFloat:=0 ;
TEcrGen.findfield('E_COUVERTUREDEV').asFloat:=0 ;
TEcrGen.findfield('E_ETATLETTRAGE').asstring:='AL' ;
TEcrGen.findfield('E_LETTRAGE').asstring:='' ;
TEcrGen.findfield('E_LETTRAGEDEV').asstring:='-' ;
TEcrGen.findfield('E_DATEPAQUETMAX').asDateTime:=TEcrGen.findfield('E_DATECOMPTABLE').asDateTime ;
TEcrGen.findfield('E_DATEPAQUETMIN').asDateTime:=TEcrGen.findfield('E_DATECOMPTABLE').asDateTime;
TEcrGen.findfield('E_DATERELANCE').asDateTime:=IDate1900 ;
END ;

{=============================================================================}
procedure TFImpCpt.recupECHE(Var Fiche4,Fiche6 : EnregCPTA ; Var NoLigne : Integer) ;

Var Cle,Test : String40 ;
    Ech : EnregCPTA ;
    PosEch,kk : Longint ;
    OkOk : Boolean ;
    Ecr2 : EnregCPTA ;
    Letat : String10 ;
    PremFois : Boolean ;
    NoEche : Integer ;
    LaDevise : Byte ;
    TypeMvt : String3 ;
    XX,XX2 : Double ;
BEGIN
LaDevise:=Fiche6.DeviseET ;
InitNewSAA(9,Ech) ; Ecr2:=Fiche4 ; kk:=0 ;
Ech.TypeH:=Ecr2.TypeE ; Ech.NoPieceH:=Ecr2.NoPieceE ; Ech.NoEcritureH:=Ecr2.NoOrdreE ;
Ech.NoEcheanceH:=0 ; Cle:=CreCle9(1,Ech.TypeH,Ech.NoPieceH,Ech.NoEcritureH,Ech.NoEcheanceH,0,'','',0,0,0,0,0) ;
Test:=Cle ; NoEche:=0 ; PremFois:=TRUE ;
Repeat
  Inc(kk) ;
  If kk=1 then SearchKey(VSAA^.Idxff[9],PosEch,Test) else NextKey(VSAA^.Idxff[9],PosEch,Test) ;
  OkOk:=((VSAA^.Ok) and (Copy(Test,1,Length(Cle))=Cle)) ;
  If OkOk then
     BEGIN
     GetRec(VSAA^.Datff[9],PosEch,Ech) ;
     TEcrGen.Insert ;
     InitNew(TEcrGen) ;
     If PremFois Then begin Inc(NoLigne) ; PremFois:=FALSE ; end ;
     Inc(NoEche) ; RecupTroncEcriture(FIche4,Fiche6,NoLigne,TypeMvt) ;
     TEcrGen.findfield('E_NUMECHE').asInteger:=NoEche ;
     TEcrGen.findfield('E_MODEPAIE').asstring:=LeModePaie(Ech.PaiementH) ;
     TEcrGen.findfield('E_DATEECHEANCE').asDateTime:=Int2Date(Ech.DateEcheanceH);
     TEcrGen.findfield('E_DEBIT').asFloat:=0 ;
     TEcrGen.findfield('E_CREDIT').asFloat:=0 ;
     XX:=Arrondi(Ech.MontantFrcsH,Dc) ;
     If Ech.SensH=1 Then
        BEGIN
        TEcrGen.findfield('E_DEBIT').asFloat:=XX ;
        END Else
        BEGIN
        TEcrGen.findfield('E_CREDIT').asFloat:=XX ;
        END ;
     TEcrGen.findfield('E_DEBITDEV').asFloat:=0 ;
     TEcrGen.findfield('E_CREDITDEV').asFloat:=0 ;
     XX2:=Arrondi(Ech.MtantLettreH,Dc) ;
     If Fiche6.DeviseET<=1 Then
        BEGIN
        TEcrGen.findfield('E_DEVISE').asstring:=TbDev[1].ISO ;
        TEcrGen.findfield('E_TAUXDEV').asFloat:=1 ;
        TEcrGen.findfield('E_COUVERTURE').asFloat:=XX2 ;
        TEcrGen.findfield('E_COUVERTUREDEV').asFloat:=XX2 ;
        If Ech.SensH=1 Then TEcrGen.findfield('E_DEBITDEV').asFloat:=XX
                       Else TEcrGen.findfield('E_CREDITDEV').asFloat:=XX ;
        END Else
        BEGIN
        TEcrGen.findfield('E_COUVERTURE').asFloat:=XX2 ;
        TEcrGen.findfield('E_COUVERTUREDEV').asFloat:=0 ;
        TEcrGen.findfield('E_DEVISE').asstring:=TbDev[LaDevise].ISO ;
        TEcrGen.findfield('E_TAUXDEV').asFloat:=Arrondi(Fiche6.TauxDeviseET,TbDev[LaDevise].DeciT) ;
        If Ech.SensH=1 Then TEcrGen.findfield('E_DEBITDEV').asFloat:=Arrondi(Ech.MtantDeviseH,TBDev[LaDevise].Deci)
                       Else TEcrGen.findfield('E_CREDITDEV').asFloat:=Arrondi(Ech.MTantDeviseH,TBDev[LaDevise].Deci) ;
        END ;
     Letat:='0000000000' ;
     TEcrGen.findfield('E_ETAT').asstring:=Letat ;
     TEcrGen.findfield('E_ETATLETTRAGE').asstring:='AL' ;
     If GetEtat(Ech.EtatGlobalH,LettrePartiel)
        Then TEcrGen.findfield('E_ETATLETTRAGE').asstring:='PL'
        Else If GetEtat(Ech.EtatGlobalH,LettreTotal)
                Then TEcrGen.findfield('E_ETATLETTRAGE').asstring:='TL' ;
     If Ech.CodeLettreH>0 Then
        TEcrGen.findfield('E_LETTRAGE').asstring:=CodeLettre(Ech.CodeLettreH,GetEtat(Ech.EtatGlobalH,LettreTotal)) ;
     TEcrGen.findfield('E_LETTRAGEDEV').asstring:='-' ;
     TEcrGen.findfield('E_DATEPAQUETMAX').asDateTime:=Int2Date(Ech.DateGlobalH);
     TEcrGen.findfield('E_DATEPAQUETMIN').asDateTime:=Int2Date(Ech.DateGlobalH);
     TEcrGen.findfield('E_DATERELANCE').asDateTime:=Int2Date(Ech.DerniereRelanceH);
     TEcrGen.findfield('E_ECHE').asstring:='X' ;
     AjusteLettrage ;
     TEcrGen.post ;
     END ;
Until (Not OkOk) ;
END ;

{=============================================================================}
procedure TFImpCpt.recupMonoECHE(Var Fiche4,Fiche6 : EnregCPTA ; Var NoLigne : Integer) ;

Var Cle,Test : String40 ;
    Ech : EnregCPTA ;
    PosEch,kk : Longint ;
    OkOk : Boolean ;
    Ecr2 : EnregCPTA ;
    Letat : String10 ;
    NoEche : Integer ;
    LaDevise : Byte ;
    TypeMvt : String3 ;
    XX,XX2 : Double ;
BEGIN
LaDevise:=Fiche6.DeviseET ;
InitNewSAA(9,Ech) ; Ecr2:=Fiche4 ; kk:=0 ;
Ech.TypeH:=Ecr2.TypeE ; Ech.NoPieceH:=Ecr2.NoPieceE ; Ech.NoEcritureH:=Ecr2.NoOrdreE ;
Ech.NoEcheanceH:=0 ; Cle:=CreCle9(1,Ech.TypeH,Ech.NoPieceH,Ech.NoEcritureH,Ech.NoEcheanceH,0,'','',0,0,0,0,0) ;
Test:=Cle ; NoEche:=0 ;
SearchKey(VSAA^.Idxff[9],PosEch,Test) ;
OkOk:=((VSAA^.Ok) and (Copy(Test,1,Length(Cle))=Cle)) ;
If OkOk then
   BEGIN
   GetRec(VSAA^.Datff[9],PosEch,Ech) ;
   Inc(NoLigne) ;
   Inc(NoEche) ; RecupTroncEcriture(FIche4,Fiche6,NoLigne,TypeMvt) ;
   TEcrGen.findfield('E_NUMECHE').asInteger:=NoEche ;
   TEcrGen.findfield('E_MODEPAIE').asstring:=LeModePaie(Ech.PaiementH) ;
   TEcrGen.findfield('E_DATEECHEANCE').asDateTime:=Int2Date(Ech.DateEcheanceH);
   TEcrGen.findfield('E_DEBIT').asFloat:=0 ;
   TEcrGen.findfield('E_CREDIT').asFloat:=0 ;
   XX:=Arrondi(Ech.MontantFrcsH,Dc) ;
   If Ech.SensH=1 Then
      BEGIN
      TEcrGen.findfield('E_DEBIT').asFloat:=XX ;
      END Else
      BEGIN
      TEcrGen.findfield('E_CREDIT').asFloat:=XX ;
      END ;
   TEcrGen.findfield('E_DEBITDEV').asFloat:=0 ;
   TEcrGen.findfield('E_CREDITDEV').asFloat:=0 ;
   XX2:=Arrondi(Ech.MtantLettreH,Dc) ;
   If Fiche6.DeviseET<=1 Then
      BEGIN
      TEcrGen.findfield('E_DEVISE').asstring:=TbDev[1].ISO ;
      TEcrGen.findfield('E_TAUXDEV').asFloat:=1 ;
      TEcrGen.findfield('E_COUVERTURE').asFloat:=XX2 ;
      TEcrGen.findfield('E_COUVERTUREDEV').asFloat:=XX2 ;
      If Ech.SensH=1 Then TEcrGen.findfield('E_DEBITDEV').asFloat:=XX
                     Else TEcrGen.findfield('E_CREDITDEV').asFloat:=XX ;
      END Else
      BEGIN
      TEcrGen.findfield('E_COUVERTURE').asFloat:=XX2 ;
      TEcrGen.findfield('E_COUVERTUREDEV').asFloat:=0 ;
      TEcrGen.findfield('E_DEVISE').asstring:=TbDev[LaDevise].ISO ;
      TEcrGen.findfield('E_TAUXDEV').asFloat:=Arrondi(Fiche6.TauxDeviseET,TbDev[LaDevise].DeciT) ;
      If Ech.SensH=1 Then TEcrGen.findfield('E_DEBITDEV').asFloat:=Arrondi(Ech.MtantDeviseH,TBDev[LaDevise].Deci)
                     Else TEcrGen.findfield('E_CREDITDEV').asFloat:=Arrondi(Ech.MTantDeviseH,TBDev[LaDevise].Deci) ;
      END ;
   Letat:='0000000000' ;
   TEcrGen.findfield('E_ETAT').asstring:=Letat ;
   TEcrGen.findfield('E_ETATLETTRAGE').asstring:='AL' ;
   If GetEtat(Ech.EtatGlobalH,LettrePartiel)
      Then TEcrGen.findfield('E_ETATLETTRAGE').asstring:='PL'
      Else If GetEtat(Ech.EtatGlobalH,LettreTotal)
              Then TEcrGen.findfield('E_ETATLETTRAGE').asstring:='TL' ;
   If Ech.CodeLettreH>0 Then
      TEcrGen.findfield('E_LETTRAGE').asstring:=CodeLettre(Ech.CodeLettreH,GetEtat(Ech.EtatGlobalH,LettreTotal)) ;
   TEcrGen.findfield('E_LETTRAGEDEV').asstring:='-' ;
   TEcrGen.findfield('E_DATEPAQUETMAX').asDateTime:=Int2Date(Ech.DateGlobalH);
   TEcrGen.findfield('E_DATEPAQUETMIN').asDateTime:=Int2Date(Ech.DateGlobalH);
   TEcrGen.findfield('E_DATERELANCE').asDateTime:=Int2Date(Ech.DerniereRelanceH);
   TEcrGen.findfield('E_ECHE').asstring:='X' ;
   AjusteLettrage ;
   END ;
END ;

{=============================================================================}
procedure TFImpCpt.recupAnalGene(Var Ecr,Fiche6 : EnregCPTA ; Var NoLigne : Integer ; TypeMvt : String3) ;

Var Cle,Test : String40 ;
    Ana : EnregCPTA ;
    PosAna,kk,NoLigneAna : Longint ;
    OkOk : Boolean ;
    LaDevise : Byte ;
    TotalDevise,MtantDevise,TotalPivot,TotalEuro : Double ;
    XX,XX2,XX3 : Double ;

BEGIN
If Not AnaHal.Checked Then Exit ;
LaDevise:=Fiche6.DeviseET ;
NoLigneAna:=0 ; kk:=0 ; TotalDevise:=0 ; TotalPivot:=0 ; TotalEuro:=0 ;
With Ecr do
     BEGIN
     Ana.TypeEA:=TypeE ;
     Ana.DAteComptaEA:=DateComptaE ;
     Ana.NoPieceEA:=NoPieceE ;
     Ana.NoEcrGenEA:=NoOrdreE ;
     Ana.NoLiGneEA:=0 ;
     Ana.CodeJournalEA:=CodeJournalE ;
     END ;
with Ana do Cle:=CreCle8(1,TypeEA,DateComptaEA,NoPieceEA,NoEcrGenEA,NoLigneEA,'','',0,CodeJournalEA) ;
Test:=Cle ;
Repeat
  Inc(kk) ; If kk=1 Then SearchKey(VSAA^.Idxff[8],PosAna,Cle) Else NextKey(VSAA^.Idxff[8],PosAna,cle) ;
  okok:=((VSAA^.ok) and (Copy(Cle,1,Length(Test))=Test)) ;
  If OkOk Then
     BEGIN
     GetRec(VSAA^.Datff[8],PosAna,Ana) ;
     LMess('Création de l''écriture N° '+IntToStr(Ecr.NoPieceE)+' Ligne '+
            IntToStr(Ecr.NoOrdreE)+' Analytique '+IntToStr(Ana.NoLigneEA)) ;
     TEcrAnal.Insert ;
     InitNew(TEcrAnal) ; Inc(NoLigneAna) ;
     TEcrAnal.findfield('Y_GENERAL').asstring:=P_CptGen ;
     TEcrAnal.findfield('Y_AXE').asstring:='A1' ;
     TEcrAnal.findfield('Y_DATECOMPTABLE').asdatetime:=Int2date(Ana.DateComptaEA) ;
     TEcrAnal.findfield('Y_NUMEROPIECE').asInteger:=Ana.NoPieceEA ;
     TEcrAnal.findfield('Y_NUMLIGNE').asInteger:=NoLigne ;
     TEcrAnal.findfield('Y_SECTION').asstring:=Trim(Ana.SectionEA) ;
     TEcrAnal.findfield('Y_EXERCICE').asstring:=P_ExerciceMvt ;
     TEcrAnal.findfield('Y_DEBIT').asFloat:=0 ;
     TEcrAnal.findfield('Y_CREDIT').asFloat:=0 ;
     TEcrAnal.findfield('Y_DEBITDEV').asFloat:=0 ;
     TEcrAnal.findfield('Y_CREDITDEV').asFloat:=0 ;
     XX:=Arrondi(ana.MontantEA,Dc) ;
     If Ana.SensEA=1 Then
        BEGIN
        TEcrAnal.findfield('Y_DEBIT').asFloat:=XX ;
        END Else
        BEGIN
        TEcrAnal.findfield('Y_CREDIT').asFloat:=XX ;
        END ;
     XX2:=Arrondi(Ecr.MontantFrcsE,Dc) ;
     TEcrAnal.findfield('Y_TOTALECRITURE').asFloat:=XX2 ;
     XX3:=PivotToEuro(XX2) ;
     TotalEuro:=TotalEuro+XX3 ;
     If LaDevise<=1 Then
        BEGIN
        TEcrAnal.findfield('Y_DEVISE').asstring:=TbDev[1].ISO ; ;
        If Ana.SensEA=1 Then TEcrAnal.findfield('Y_DEBITDEV').asFloat:=Arrondi(ana.MontantEA,Dc)
                        Else TEcrAnal.findfield('Y_CREDITDEV').asFloat:=Arrondi(ana.MontantEA,Dc) ;
        TEcrAnal.findfield('Y_TOTALDEVISE').asFloat:=Arrondi(Ecr.MontantFrcsE,Dc) ;
        TEcrAnal.findfield('Y_TAUXDEV').asFloat:=1 ;
        END Else
        BEGIN
        TEcrAnal.findfield('Y_DEVISE').asstring:=TbDev[LaDevise].ISO ; ;
        TEcrAnal.findfield('Y_TAUXDEV').asFloat:=Arrondi(Fiche6.TauxDeviseET,TbDev[LaDevise].DeciT) ;
        MtantDevise:=Arrondi(ana.MontantEA/Fiche6.TAuxDeviseET,TbDev[LaDevise].Deci) ;
        TEcrAnal.findfield('Y_TOTALDEVISE').asFloat:=Arrondi(Ecr.MtantDeviseE,TbDev[LaDevise].Deci) ;
        TotalDevise:=TotalDevise+MtantDevise ;
        TotalPivot:=TotalPivot+Arrondi(Ana.MontantEA,Dc) ;
        If (Arrondi(TotalPivot,Dc)=Arrondi(Ecr.MontantFrcsE,Dc)) And
           (Arrondi(TotalDevise,TbDev[LaDevise].Deci)<>Arrondi(Ecr.MtantDeviseE,TbDev[LaDevise].Deci)) Then
           BEGIN
           MtantDevise:=Arrondi(Ecr.MtantDeviseE-TotalDevise,TbDev[LaDevise].Deci) ;
           END ;
        If Ana.SensEA=1 Then TEcrAnal.findfield('Y_DEBITDEV').asFloat:=MtantDevise
                        Else TEcrAnal.findfield('Y_CREDITDEV').asFloat:=MtantDevise ;
        END ;
     TEcrAnal.findfield('Y_DATETAUXDEV').asDateTime:=Int2Date(Ana.DateComptaEA) ; ;
     TEcrAnal.findfield('Y_REFINTERNE').asstring:=ASCII2ANSI(Trim(ana.referenceEA)) ;
     TEcrAnal.findfield('Y_LIBELLE').asstring:=ASCII2ANSI(Trim(Ana.LibelleEA)) ;
     TEcrAnal.findfield('Y_NATUREPIECE').asstring:=NatPiece(Ecr.TypePieceE) ;
     TEcrAnal.findfield('Y_QUALIFPIECE').asstring:='N' ;
     If Fiche6.TypeET=EcrSim Then TEcrAnal.findfield('Y_QUALIFPIECE').asstring:='S' ;

     TEcrAnal.findfield('Y_TYPEANALYTIQUE').asstring:='-' ;
     TEcrAnal.findfield('Y_VALIDE').asstring:='-' ;
     If GetEtat(Ecr.EtatE,Validee) Then TEcrAnal.findfield('Y_VALIDE').asstring:='X' ;
     TEcrAnal.findfield('Y_ETAT').asstring:='0000000000' ;
     TEcrAnal.findfield('Y_REFEXTERNE').asstring:='' ;
     TEcrAnal.findfield('Y_DATEREFEXTERNE').asDateTime:=Int2date(Ana.DateComptaEA) ;
     TEcrAnal.findfield('Y_CONTROLEUR').asstring:='   ' ;
     TEcrAnal.findfield('Y_DATECREATION').asDateTime:=Int2Date(Ana.DateComptaEA) ;
     TEcrAnal.findfield('Y_ETABLISSEMENT').asstring:=VH^.EtablisDefaut ;
     {
     TEcrAnal.findfield('Y_DATEMODIF').asstring:=
     TEcrAnal.findfield('Y_SOCIETE').asstring:=
     TEcrAnal.findfield('Y_BLOCNOTE').asstring:=
     TEcrAnal.findfield('Y_VISION').asstring:=
     TEcrAnal.findfield('Y_AFFAIRE').asstring:=
     }
     TEcrAnal.findfield('Y_POURCENTAGE').asFloat:=Arrondi(Ana.TauxRepartEA,ADecimP) ;
     TEcrAnal.findfield('Y_CONTROLE').asString:='' ;
     {
     TEcrAnal.findfield('Y_DATETAUXECU').asFloat:=0 ;
     }
     TEcrAnal.findfield('Y_QTE1').asFloat:=0 ;
     TEcrAnal.findfield('Y_QTE2').asFloat:=0 ;
     TEcrAnal.findfield('Y_QUALIFQTE1').asString:='AUC' ;
     TEcrAnal.findfield('Y_QUALIFQTE2').asString:='AUC' ;
     TEcrAnal.findfield('Y_QUALIFECRQTE1').asString:='AUC' ;
     TEcrAnal.findfield('Y_QUALIFECRQTE2').asString:='AUC' ;
     TEcrAnal.findfield('Y_JOURNAL').asString:=Ana.CodeJournalEA ;
     TEcrAnal.findfield('Y_NUMVENTIL').asInteger:=NoLigneAna ;
     TEcrAnal.findfield('Y_POURCENTQTE1').asFloat:=Arrondi(Ana.TauxRepartEA,Dc) ;
     TEcrAnal.findfield('Y_POURCENTQTE2').asFloat:=Arrondi(Ana.TauxRepartEA,Dc) ;
     TEcrAnal.findfield('Y_TYPEMVT').asString:=TypeMvt ;
     TEcrAnal.findfield('Y_ECRANOUVEAU').asstring:='N' ;
     If GetEtat(Ecr.EtatE,ANouveau) Then
        begin
        TEcrAnal.findfield('Y_ECRANOUVEAU').asstring:='H' ;
        end ;
     TEcrAnal.Post ;
     END ;
Until Not OkOk ;
(*
*)
END ;

{=============================================================================}
procedure TFImpCpt.SauveLaPos6(LaPos : LongInt) ;
Var LaSoc : String3 ;
    St1 : String ;
begin
LaSoc:=Trim(Copy(VSAA^.ParaSoc.CodeSociete,1,3)) ;
If LaSoc='' Then LaSoc:='001' ;
BeginTrans ;
TSoc.Open ;
If Tsoc.FindKey([LaSoc]) Then
   BEGIN
   St1:=TSoc.FindField('SO_RECUPCPTA').AsString ;
   St1:=Copy(St1,1,20)+FormatFloat('0000000',LaPos) ;
{$IFDEF SPEC302}
   TSoc.Edit ;
   TSoc.FindField('SO_RECUPCPTA').AsString:=St1 ;
   TSoc.Post ;
{$ELSE}
   SetParamSoc('SO_RECUPCPTA',St1) ;
{$ENDIF}
   END ;
TSoc.Close ;
CommitTrans ;
end ;

{=============================================================================}
Function TFImpCpt.LitLaPos6 : LongInt ;
Var LaSoc : String3 ;
    St    : String ;
begin
Result:=0 ;
LaSoc:=Trim(Copy(VSAA^.ParaSoc.CodeSociete,1,3)) ;
If LaSoc='' Then LaSoc:='001' ;
BeginTrans ;
TSoc.Open ;
If Tsoc.FindKey([LaSoc]) Then
   BEGIN
{$IFDEF SPEC302}
   St:=Copy(TSoc.FindField('SO_RECUPCPTA').AsString,21,7) ;
{$ELSE}
   St:=GetparamSoc('SO_RECUPCPTA') ;
{$ENDIF}
   If St<>'' Then Result:=StrToInt(St) Else Result:=0 ;
   END ;
TSoc.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.recupEcrGen(OldPos6 : LongInt);
Var LaPos,LaVraiPos,Lg,PosEcr,kk,Pos6Dep : LongInt ;
    Ecr,Fiche6,Jal : EnregCPTA ;
    Cle,Test : String40 ;
    OkOk,OkMvt : Boolean ;
    NoLigne : Integer ;
    TypeMvt : String3 ;
 label 1 ;
begin
Pos6Dep:=OldPos6 ; LaVraiPos:=0 ; 
AlimTva ;
BeginTrans ;
If Pos6Dep=1 Then
   BEGIN
   VideTable('ECRITURE','') ; VideTable('ANALYTIQ','') ;
   VideTable('EEXBQ','') ; RecupExo ;
   END ;
Ouvre(4,2,FALSE,spCPTA) ; Ouvre(0,0,FALSE,spCPTA) ; Ouvre(1,0,FALSE,spCPTA) ;
Ouvre(9,2,FALSE,spCPTA) ; Ouvre(8,2,FALSE,spCPTA) ;Ouvre(6,0,FALSE,spCPTA) ;
Ouvre(3,0,FALSE,spCPTA) ;
Lg:=FileLen(VSAA^.DAtFF[6])-1 ;
InitMove(Lg,'recup des écritures') ;
TEcrGen.open ; TEcrAnal.open ; TEXBQ.Open ;
For LaPos:=Pos6Dep To Lg Do
  BEGIN
  GetRec(VSAA^.DatFF[6],LaPos,Fiche6) ;
  LaVraiPos:=LaPos ;
  if MoveCur(FALSE) or Stopper then goto 1;
  OkMvt:=TRUE ;
  If (Not ANHal.Checked) And (Fiche6.PosJournalET>0) Then
     BEGIN
     Getrec(VSAA^.DatFF[3],Fiche6.PosJournalET,Jal) ;
     OkMvt:=(Jal.NatureJ<>6) ; { AN }
     If Not OkMvt Then OkMvt:=(Fiche6.DateComptaET<VSAA^.parasoc.EnCours.Deb) ;
     END ;
  If OkMvt And (Fiche6.Status=0) And ((Fiche6.TypeET=EcrReel) Or (Fiche6.TypeET=EcrSim)) Then
    BEGIN
    NoLigne:=0 ;
    With Fiche6 Do
         BEGIN
         Ecr.TypeE:=TypeET ;             Ecr.CodeJournalE:=CodeJournalET ;
         Ecr.DateComptaE:=DateComptaET ; Ecr.NoPieceE:=NoPieceET ;
         Ecr.NoOrdreE:=0 ;
         END ;
    With Ecr do Cle:=CreCle4(1,TypeE,CodeJournalE,DateComptaE,NoPieceE,NoOrdreE,'','',0) ;
    Test:=Cle ;  kk:=0 ;
    Repeat
      Inc(kk) ; If kk=1 Then SearchKey(VSAA^.Idxff[4],PosEcr,Cle) Else NextKey(VSAA^.Idxff[4],PosEcr,Cle) ;
      okok:=((VSAA^.ok) and (Copy(Cle,1,Length(Test))=test)) ;
      If OkOk Then
         BEGIN
         GetRec(VSAA^.Datff[4],PosEcr,Ecr) ;
         If Ecr.PosEnteteE=LaPos Then
            BEGIN
            LMess('Création de l''écriture N° '+IntToStr(Ecr.NoPieceE)+' Ligne '+IntToStr(Ecr.NoOrdreE)) ;
            If (Ecr.AnalytiqueONE=1) And (AnaHAL.Checked) And (Ecr.EcheanceONE=1) Then
               BEGIN
               TEcrGen.Insert ;
               InitNew(TEcrGen) ;
               RecupMonoEche(Ecr,Fiche6,NoLigne) ;
               If (Ecr.AnalytiqueONE=1) And (AnaHAL.Checked) Then
                  BEGIN
                  RecupAnalGene(Ecr,Fiche6,NoLigne,TypeMvt) ;
                  END ;
               TEcrGen.post ;
               END Else
               BEGIN
               If Ecr.EcheanceONE=1 Then
                  BEGIN
                  RecupEche(Ecr,Fiche6,NoLigne) ;
                  END Else
                  BEGIN
                  TEcrGen.Insert ;
                  Inc(NoLigne) ; InitNew(TEcrGen) ;
                  RecupTroncEcriture(Ecr,Fiche6,NoLigne,TypeMvt) ;
                  If (Ecr.AnalytiqueONE=1) And (AnaHAL.Checked) Then
                     BEGIN
                     RecupAnalGene(Ecr,Fiche6,NoLigne,TypeMvt) ;
                     END ;
                  TEcrGen.post ;
                  END ;
               END ;
            END ;
         END ;
    Until Not OkOK ;
    END ;
  END ;
1:TEcrGen.Close ; TEcrAnal.Close ; TEXBQ.Close ;
GAccess.Ferme(4,2,FALSE) ; GAccess.Ferme(0,0,FALSE) ; GAccess.Ferme(1,0,FALSE) ; GAccess.Ferme(9,2,FALSE) ;
GAccess.Ferme(8,2,FALSE) ;GAccess.Ferme(6,0,FALSE) ; GAccess.Ferme(3,0,FALSE) ;
FiniMove ;
CommitTrans ;
If Stopper Then SauveLaPos6(LaVraiPos) Else SauveLaPos6(0) ;
end ;

{=============================================================================}
procedure TFImpCpt.recupODA ;

Var Fiche8,Fiche3,General            : EnregCPTA ;
    Pos3,Pos8,kk,kk8,Lg,OldNoPiece   : Longint ;
    Cle3,Cle8,Test8                  : String40 ;
    Okok3,Okok8,OkPourStop           : Boolean ;
    NoLigneAna                       : Integer ;
    DateFinExo                       : TDateTime ;
    XX                               : Double ;
  Label 1 ;
BEGIN
VideTable('ANALYTIQ','WHERE Y_TYPEANALYTIQUE="X"') ;
If Not ANAHal.Checked Then Exit ;
BeginTrans ;
TEcrAnal.Open ;
Pos3:=0 ; kk:=0 ; Cle3:=#1 ; OkPourStop:=FALSE ;
Ouvre(3,2,FALSE,spCPTA) ; Ouvre(8,2,FALSE,spCPTA) ; Ouvre(0,0,FALSE,spCPTA) ;
Lg:=FileLen(VSAA^.Datff[8])-1 ;
InitMove(Lg,'Récup des OD analytiques') ;
OldNoPiece:=0 ;
Repeat
  Inc(kk) ; if kk=1 then SearchKey(VSAA^.Idxff[3],Pos3,Cle3) else NextKey(VSAA^.Idxff[3],Pos3,Cle3) ;
  OkOk3:=((VSAA^.Ok) and (Cle3[1]=#1)) ;
  If OkOk3 then
     BEGIN
     GetRec(VSAA^.Datff[3],Pos3,Fiche3) ;
     If Fiche3.NatureJ=9 then
        BEGIN
        Pos8:=0 ; kk8:=0 ;
        Cle8:=CreCle8(1,EcrReel,0,0,0,0,'','',0,Fiche3.CodeJ) ;
        Test8:=CreCle8(1,EcrReel,0,0,0,0,'','',0,Fiche3.CodeJ) ; NoLigneAna:=0 ;
        Repeat
          if MoveCur(FALSE) then ; If Stopper Then OkPourStop:=TRUE ;
          Inc(kk8) ; If kk8=1 then SearchKey(VSAA^.Idxff[8],Pos8,Cle8) else NextKey(VSAA^.Idxff[8],Pos8,Cle8) ;
          OkOk8:=((VSAA^.Ok) and (Copy(Cle8,1,Length(Test8))<=Test8)) ;
          If OkOk8 then
             BEGIN
             GetRec(VSAA^.Datff[8],Pos8,Fiche8) ;
             If OkPourStop Then Goto 1 ;
             If (kk8>1) And (Fiche8.NoPieceEA<>OldNoPiece) Then NoLigneAna:=0 ;
             OldNoPiece:=Fiche8.NoPieceEA ;
             General.CompteG:='' ;
             if Fiche8.PosCpteGeneEA>0 then GetRec(VSAA^.Datff[0],Fiche8.PosCpteGeneEA,General) ;
             LMess('Création de l''écriture N° '+IntToStr(Fiche8.NoPieceEA)+' Ligne '+
                    IntToStr(Fiche8.NoLigneEA)) ;
             TEcrAnal.Insert ;
             InitNew(TEcrAnal) ; Inc(NoLigneAna) ;
             TEcrAnal.findfield('Y_GENERAL').asstring:=General.CompteG ;
             TEcrAnal.findfield('Y_AXE').asstring:='A1' ;
             TEcrAnal.findfield('Y_DATECOMPTABLE').asdatetime:=Int2date(Fiche8.DateComptaEA) ;
             TEcrAnal.findfield('Y_NUMEROPIECE').asInteger:=Fiche8.NoPieceEA ;
             TEcrAnal.findfield('Y_NUMLIGNE').asInteger:=0 ;
             TEcrAnal.findfield('Y_SECTION').asstring:=Trim(Fiche8.SectionEA) ;
             TEcrAnal.findfield('Y_EXERCICE').asstring:=TrouveExo(Fiche8.DateComptaEA,DateFinExo) ;
             TEcrAnal.findfield('Y_DEBIT').asFloat:=0 ;
             TEcrAnal.findfield('Y_CREDIT').asFloat:=0 ;
             TEcrAnal.findfield('Y_DEBITDEV').asFloat:=0 ;
             TEcrAnal.findfield('Y_CREDITDEV').asFloat:=0 ;
             XX:=Arrondi(Fiche8.MontantEA,Dc) ;
             If Fiche8.SensEA=1 Then
                BEGIN
                TEcrAnal.findfield('Y_DEBIT').asFloat:=XX ;
                END Else
                BEGIN
                TEcrAnal.findfield('Y_CREDIT').asFloat:=XX ;
                END ;
             TEcrAnal.findfield('Y_TOTALECRITURE').asFloat:=0 ;
             TEcrAnal.findfield('Y_DEVISE').asstring:=TbDev[1].ISO ; ;
             If Fiche8.SensEA=1 Then TEcrAnal.findfield('Y_DEBITDEV').asFloat:=Arrondi(Fiche8.MontantEA,Dc)
                                Else TEcrAnal.findfield('Y_CREDITDEV').asFloat:=Arrondi(Fiche8.MontantEA,Dc) ;
             TEcrAnal.findfield('Y_TOTALDEVISE').asFloat:=0;
             TEcrAnal.findfield('Y_TAUXDEV').asFloat:=1 ;
             TEcrAnal.findfield('Y_DATETAUXDEV').asDateTime:=Int2Date(Fiche8.DateComptaEA) ; ;
             TEcrAnal.findfield('Y_REFINTERNE').asstring:=ASCII2ANSI(Trim(Fiche8.referenceEA)) ;
             TEcrAnal.findfield('Y_LIBELLE').asstring:=ASCII2ANSI(Trim(Fiche8.LibelleEA)) ;
             TEcrAnal.findfield('Y_NATUREPIECE').asstring:='OD' ;
             TEcrAnal.findfield('Y_QUALIFPIECE').asstring:='N' ;
             TEcrAnal.findfield('Y_TYPEANALYTIQUE').asstring:='X' ;
             TEcrAnal.findfield('Y_VALIDE').asstring:='-' ;
             TEcrAnal.findfield('Y_ETAT').asstring:='0000000000' ;
             TEcrAnal.findfield('Y_REFEXTERNE').asstring:='' ;
             TEcrAnal.findfield('Y_DATEREFEXTERNE').asDateTime:=Int2date(Fiche8.DateComptaEA) ;
             TEcrAnal.findfield('Y_CONTROLEUR').asstring:='   ' ;
             TEcrAnal.findfield('Y_DATECREATION').asDateTime:=Int2Date(Fiche8.DateComptaEA) ;
             TEcrAnal.findfield('Y_ETABLISSEMENT').asstring:=VH^.EtablisDefaut ;
             {
             TEcrAnal.findfield('Y_DATEMODIF').asstring:=
             TEcrAnal.findfield('Y_SOCIETE').asstring:=
             TEcrAnal.findfield('Y_ETABLISSEMENT').asstring:=V_PGI.EtablisDefaut ;
             TEcrAnal.findfield('Y_BLOCNOTE').asstring:=
             TEcrAnal.findfield('Y_VISION').asstring:=
             TEcrAnal.findfield('Y_AFFAIRE').asstring:=
             }
             TEcrAnal.findfield('Y_POURCENTAGE').asFloat:=Arrondi(Fiche8.TauxRepartEA,ADecimP) ;
             TEcrAnal.findfield('Y_CONTROLE').asString:='' ;
             {
             TEcrAnal.findfield('Y_DATETAUXECU').asFloat:=0 ;
             }
             TEcrAnal.findfield('Y_QTE1').asFloat:=0 ;
             TEcrAnal.findfield('Y_QTE2').asFloat:=0 ;
             TEcrAnal.findfield('Y_QUALIFQTE1').asString:='AUC' ;
             TEcrAnal.findfield('Y_QUALIFQTE2').asString:='AUC' ;
             TEcrAnal.findfield('Y_QUALIFECRQTE1').asString:='AUC' ;
             TEcrAnal.findfield('Y_QUALIFECRQTE2').asString:='AUC' ;
             TEcrAnal.findfield('Y_JOURNAL').asString:=Fiche8.CodeJournalEA ;
             TEcrAnal.findfield('Y_NUMVENTIL').asInteger:=NoLigneAna ;
             TEcrAnal.findfield('Y_POURCENTQTE1').asFloat:=Arrondi(Fiche8.TauxRepartEA,Dc) ;
             TEcrAnal.findfield('Y_POURCENTQTE2').asFloat:=Arrondi(Fiche8.TauxRepartEA,Dc) ;
             TEcrAnal.findfield('Y_TYPEMVT').asString:='DIV' ;
             TEcrAnal.findfield('Y_ECRANOUVEAU').asstring:='N' ;
             TEcrAnal.Post ;
             END Else If OkPourStop Then Goto 1 ;
        Until (Not OkOk8) ;
        END ;
     END ;
Until Not OkOk3 ;
1:GAccess.Ferme(3,2,FALSE) ; GAccess.Ferme(8,2,FALSE) ; GAccess.Ferme(0,0,FALSE) ;
TEcrAnal.Close ;
FiniMove ;
CommitTrans ;
END ;

{=============================================================================}
procedure litGuideV8 (Var guideV8 : TabGuideV8 ; Var MaxGuide : SmallInt) ;
Var Cle16,Test16 : String40 ;
    Pos16,kk : LongInt ;
    Fiche16 : EnregCpta ;
    OkOk : Boolean ;
begin
Cle16:=CreCle16(1,110,'','','',0) ; Test16:=Cle16 ; Pos16:=0 ; kk:=0 ;
Ouvre(16,2,FALSE,spCPTA) ; MaxGuide:=0 ;
Repeat
  Inc(kk) ;
  If kk=1 then SearchKey(VSAA^.Idxff[16],Pos16,Cle16) else NextKey(VSAA^.Idxff[16],Pos16,Cle16) ;
  OkOk:=((VSAA^.Ok) and (Copy(Cle16,1,Length(Test16))=Test16)) ;
  If OkOk then
     BEGIN
     GetRec(VSAA^.Datff[16],Pos16,Fiche16) ; MaxGuide:=kk ;
     guideV8[kk].Guide:=FormatFloat('000',kk) ;
     guideV8[kk].Libelle:=ASCII2ANSI(Fiche16.IntituleTA) ;
     guideV8[kk].Abrege:=Copy(Fiche16.CodeTrancheTA,1,8) ;
     guideV8[kk].NaturePiece:=NatPiece(Fiche16.TypBordTA) ;
     guideV8[kk].Journal:=Trim(Fiche16.Compte1TA) ;
     END ;
Until (Not OkOk) ;
GAccess.Ferme(16,2,FALSE) ;
end ;


{=============================================================================}
procedure TFImpCpt.recupGuide(Var MaxGuide : SmallInt) ;
Var guideV8 : TabGuideV8 ;
    i : SmallInt ;
    OkOk : Boolean ;
    CleG : String40 ;
    PosG,kk : LongInt ;
    FicheG : TGuide ;
    Arret : String ;
begin
VideTable('GUIDE','WHERE GU_TYPE="NOR"') ;
VideTable('ECRGUI','WHERE EG_TYPE="NOR"') ;
VideTable('ANAGUI','WHERE AG_TYPE="NOR"') ;
BeginTrans ;
TGUIDES.open ; TECRGUI.open ;
FillChar(guideV8,SizeOf(guideV8),#0) ; LitGuideV8(guideV8,MaxGuide) ;
i:=1 ; OuvreGuide(2) ; Arret:='-----------' ;
While GuideV8[i].Guide<>'' Do
  BEGIN
  kk:=0 ; CleG:=CreCleGuide(GuideV8[i].Abrege,0) ; OkOk:=FALSE ;
  TGUIDES.Insert ;
  InitNew(TGUIDES) ;
  TGUIDES.FindField('GU_TYPE').AsString:='NOR' ;
  TGUIDES.FindField('GU_GUIDE').AsString:=guideV8[i].Guide ;
  TGUIDES.FindField('GU_LIBELLE').AsString:=Ascii2Ansi(guideV8[i].Libelle) ;
  TGUIDES.FindField('GU_ABREGE').AsString:=Ascii2Ansi(guideV8[i].Abrege) ;
  TGUIDES.FindField('GU_JOURNAL').AsString:=guideV8[i].Journal;
  TGUIDES.FindField('GU_NATUREPIECE').AsString:=guideV8[i].NaturePiece ;
  TGUIDES.FindField('GU_DATECREATION').AsDateTime:=V_PGI.DateEntree ;
  TGUIDES.FindField('GU_DATEMODIF').AsDateTime:=V_PGI.DateEntree ;
  TGUIDES.FindField('GU_DEVISE').AsString:=tbDev[1].Iso ;
  TGUIDES.Post ;
  Repeat
    Inc(kk) ;
    If kk=1 then SearchKey(VSAA^.IdxfGuide,PosG,CleG) else NextKey(VSAA^.IdxfGuide,PosG,CleG) ;
    OkOk:=VSAA^.Ok And (Trim(Copy(CleG,2,8))=Trim(GuideV8[i].Abrege)) ;
    If OkOk Then
       BEGIN
       GetRec(VSAA^.DatFGuide,PosG,FicheG) ;
       FicheG.CodeLG   :=Ascii2Ansi(FicheG.CodeLG)  ; FicheG.Cpt1LG   :=Ascii2Ansi(FicheG.Cpt1LG)  ;
       FicheG.Cpt2LG   :=Ascii2Ansi(FicheG.Cpt2LG)  ; FicheG.RefLG    :=Ascii2Ansi(FicheG.RefLG)   ;
       FicheG.LibLG    :=Ascii2Ansi(FicheG.LibLG)   ; FicheG.DebitLg  :=Ascii2Ansi(FicheG.DebitLg) ;
       FicheG.CreditLg :=Ascii2Ansi(FicheG.CreditLg);
       TECRGUI.Insert ;
       InitNew(TECRGUI) ;
       TECRGUI.FindField('EG_TYPE').AsString:='NOR' ;
       TECRGUI.FindField('EG_GUIDE').AsString:=guideV8[i].Guide ;
       TECRGUI.FindField('EG_NUMLIGNE').AsInteger:=kk ;
       TECRGUI.FindField('EG_GENERAL').AsString:=DecodeGuide(1,FicheG.Cpt1LG,Arret) ;
       TECRGUI.FindField('EG_AUXILIAIRE').AsString:=DecodeGuide(2,FicheG.Cpt2LG,Arret) ;
       TECRGUI.FindField('EG_REFINTERNE').AsString:=DecodeGuide(3,FicheG.RefLG,Arret) ;
       TECRGUI.FindField('EG_LIBELLE').AsString:=DecodeGuide(4,FicheG.LibLG,Arret) ;
       TECRGUI.FindField('EG_DEBITDEV').AsString:=DecodeGuide(5,FicheG.DebitLg,Arret) ;
       TECRGUI.FindField('EG_CREDITDEV').AsString:=DecodeGuide(6,FicheG.CreditLg,Arret) ;
       TECRGUI.FindField('EG_ARRET').AsString:=Arret ;
       TECRGUI.Post ;
       END ;
  Until Not OkOk ;
  Inc(i) ;
  END ;
FermeGuide(2) ;
TGUIDES.Close ; TECRGUI.Close ;
CommitTrans ;
end ;

{=============================================================================}
Function AboSeparePar(b : Byte) : String3 ;
begin
Result:='1M' ;
Case b Of  1 : Result:='SEM' ; 3 : Result:='QUI' ; 4..9 : Result:=IntToStr(b-3)+'M' ; END ;
end ;

{=============================================================================}
Function AboReconduction(b : Byte) : String3 ;
begin
Result:='DEM' ;
Case b Of  1 : Result:='DEM' ; 2 : Result:='SUP' ; 3 : Result:='TAC' ; END ;
end ;

{=============================================================================}
procedure TFImpCpt.recupAnalGeneAbo(Var Ecr,Fiche6 : EnregCPTA ; Var NoLigne : Integer ;
                                    MaxGuide : SmallInt ; GU_TYPE : String3) ;

Var Cle,Test : String40 ;
    Ana : EnregCPTA ;
    PosAna,kk,NoLigneAna : Longint ;
    OkOk : Boolean ;
    TotalDevise,TotalPivot : Double ;
BEGIN
NoLigneAna:=0 ; kk:=0 ; TotalDevise:=0 ; TotalPivot:=0 ;
With Ecr do
     BEGIN
     Ana.TypeEA:=TypeE ;
     Ana.DAteComptaEA:=DateComptaE ;
     Ana.NoPieceEA:=NoPieceE ;
     Ana.NoEcrGenEA:=NoOrdreE ;
     Ana.NoLiGneEA:=0 ;
     Ana.CodeJournalEA:=CodeJournalE ;
     END ;
with Ana do Cle:=CreCle8(1,TypeEA,DateComptaEA,NoPieceEA,NoEcrGenEA,NoLigneEA,'','',0,CodeJournalEA) ;
Test:=Cle ;
Repeat
  Inc(kk) ; If kk=1 Then SearchKey(VSAA^.Idxff[8],PosAna,Cle) Else NextKey(VSAA^.Idxff[8],PosAna,cle) ;
  okok:=((VSAA^.ok) and (Copy(Cle,1,Length(Test))=Test)) ;
  If OkOk Then
     BEGIN
     GetRec(VSAA^.Datff[8],PosAna,Ana) ;
     LMess('Création de l''écriture N° '+IntToStr(Ecr.NoPieceE)+' Ligne '+
            IntToStr(Ecr.NoOrdreE)+' Analytique '+IntToStr(Ana.NoLigneEA)) ;
     TAnaGui.Insert ;
     InitNew(TAnaGui) ; Inc(NoLigneAna) ;
     TAnaGui.findfield('AG_TYPE').asstring:=GU_TYPE ;
     TAnaGui.findfield('AG_GUIDE').asstring:=FormatFloat('000',MaxGuide) ;
     TAnaGui.findfield('AG_NUMLIGNE').asInteger:=NoLigne ;
     TAnaGui.findfield('AG_NUMVENTIL').asInteger:=Ana.NoLigneEA ;
     TAnaGui.findfield('AG_AXE').asstring:='A1' ;
     TAnaGui.findfield('AG_SECTION').asstring:=Trim(Ana.SectionEA) ;
     TAnaGui.findfield('AG_POURCENTAGE').AsFloat:=Arrondi(Ana.TauxRepartEA,ADecimP) ; ;
     TAnaGui.findfield('AG_ARRET').asstring:='XX--' ;
     TAnaGui.Post ;
     END ;
Until Not OkOk ;
END ;

{=============================================================================}
procedure TFImpCpt.recupEcrAboOuTyp(Var MaxGuide : SmallInt ; LeType : Byte) ;
Var LaPos,Lg,PosEcr,kk : LongInt ;
    Ecr,Fiche6,Compte : EnregCPTA ;
    Cle,Test : String40 ;
    OkOk : Boolean ;
    NoLigne,NumAbo : Integer ;
    Gen,Aux,Arret : String17 ;
    Abo : Boolean ;
    GU_TYPE : String3 ;
    Montant : Real48 ;
 label 1 ;
begin
Abo:=LeType=GDECLADF.EcrAbo ;
If Abo Then
   BEGIN
   VideTable('GUIDE','WHERE GU_TYPE="ABO"') ;
   VideTable('ECRGUI','WHERE EG_TYPE="ABO"') ;
   VideTable('ANAGUI','WHERE AG_TYPE="ABO"') ;
   VideTable('CONTABON','WHERE CB_COMPTABLE="X"') ;
   GU_TYPE:='ABO' ;
   END Else GU_TYPE:='NOR' ;
BeginTrans ;
Ouvre(4,2,FALSE,spCPTA) ; Ouvre(0,0,FALSE,spCPTA) ; Ouvre(1,0,FALSE,spCPTA) ;
Ouvre(9,2,FALSE,spCPTA) ; Ouvre(8,2,FALSE,spCPTA) ;Ouvre(6,0,FALSE,spCPTA) ;
If Abo Then TCONTABO.OPEN ;
TGUIDES.open ; TECRGUI.open ; TANAGUI.Open ;
Lg:=FileLen(VSAA^.DAtFF[6])-1 ;
InitMove(Lg,'recup des écritures d''abonnement') ; NumAbo:=0 ;
For LaPos:=1 To Lg Do
  BEGIN
  GetRec(VSAA^.DatFF[6],LaPos,Fiche6) ;
  if MoveCur(FALSE) or Stopper then goto 1;
  If (Fiche6.Status=0) And (Fiche6.TypeET=LeType) Then
    BEGIN
    Inc(MaxGuide) ;
    If Abo Then
       BEGIN
       Inc(NumAbo) ;
       TCONTABO.Insert ;
       InitNew(TCONTABO) ;
       TCONTABO.FindField('CB_CONTRAT').AsString:=FormatFloat('000',NumAbo) ;
       TCONTABO.FindField('CB_COMPTABLE').AsString:='X' ;
       TCONTABO.FindField('CB_LIBELLE').AsString:=Fiche6.referenceET ;
       TCONTABO.FindField('CB_DATECONTRAT').AsDateTime:=Int2Date(Fiche6.DateDepartET) ;
       TCONTABO.FindField('CB_SEPAREPAR').AsString:=AboSeparePar(Fiche6.PeriodiciteET) ;
       TCONTABO.FindField('CB_ARRONDI').AsString:='PAS' ;
       If Fiche6.FinDeMoisET=1 Then TCONTABO.FindField('CB_ARRONDI').AsString:='FIN' ;
       TCONTABO.FindField('CB_RECONDUCTION').AsString:=AboReconduction(Fiche6.ReconductionET) ;
       TCONTABO.FindField('CB_NBREPETITION').AsInteger:=Fiche6.NbRepeteET ;
       TCONTABO.FindField('CB_DEJAGENERE').AsInteger:=Fiche6.NbGenereET ;
       TCONTABO.FindField('CB_DATEDERNGENERE').AsDateTime:=Int2Date(Fiche6.LastDateET) ;
       TCONTABO.FindField('CB_GUIDE').AsString:=FormatFloat('000',MaxGuide) ; ;
       TCONTABO.FindField('CB_DATECREATION').AsDateTime:=Int2DAte(Fiche6.dateComptaET) ;
       TCONTABO.FindField('CB_DATEMODIF').AsDateTime:=Int2DAte(Fiche6.dateComptaET) ;
       TCONTABO.Post ;
       END ;
    TGUIDES.Insert ;
    InitNew(TGUIDES) ;
    TGUIDES.FindField('GU_TYPE').AsString:=GU_TYPE ;
    TGUIDES.FindField('GU_GUIDE').AsString:=FormatFloat('000',MaxGuide) ;
    If Abo Then TGUIDES.FindField('GU_LIBELLE').AsString:=Ascii2Ansi(Fiche6.ReferenceET)
           Else TGUIDES.FindField('GU_LIBELLE').AsString:='Ecr. Type '+Ascii2Ansi(Fiche6.ReferenceET) ;
    TGUIDES.FindField('GU_ABREGE').AsString:=Ascii2Ansi(Fiche6.LibelleET) ;
    TGUIDES.FindField('GU_JOURNAL').AsString:=Fiche6.CodeJournalET;
    TGUIDES.FindField('GU_NATUREPIECE').AsString:=NatPiece(Fiche6.TypePieceET) ; ;
    TGUIDES.FindField('GU_DATECREATION').AsDateTime:=V_PGI.DateEntree ;
    TGUIDES.FindField('GU_DATEMODIF').AsDateTime:=V_PGI.DateEntree ;
    TGUIDES.FindField('GU_DEVISE').AsString:=tbDev[1].Iso ;
    TGUIDES.Post ;

    NoLigne:=0 ;
    With Fiche6 Do
         BEGIN
         Ecr.TypeE:=TypeET ;             Ecr.CodeJournalE:=CodeJournalET ;
         Ecr.DateComptaE:=DateComptaET ; Ecr.NoPieceE:=NoPieceET ;
         Ecr.NoOrdreE:=0 ;
         END ;
    With Ecr do Cle:=CreCle4(1,TypeE,CodeJournalE,DateComptaE,NoPieceE,NoOrdreE,'','',0) ;
    Test:=Cle ;  kk:=0 ;
    Repeat
      Inc(kk) ; If kk=1 Then SearchKey(VSAA^.Idxff[4],PosEcr,Cle) Else NextKey(VSAA^.Idxff[4],PosEcr,Cle) ;
      okok:=((VSAA^.ok) and (Copy(Cle,1,Length(Test))=test)) ;
      If OkOk Then
         BEGIN
         GetRec(VSAA^.Datff[4],PosEcr,Ecr) ;
         Gen:='' ; Aux:='' ; Arret:='XXXX--' ;
         If Ecr.PosEnteteE=LaPos Then
            BEGIN
            If Ecr.PosCpteGeneE>0 Then
               BEGIN
               Getrec(VSAA^.DatFF[0],Ecr.PosCpteGeneE,Compte) ;
               Gen:=Trim(Compte.CompteG) ;
               END ;
            If Ecr.PosCpteAuxE>0 Then
               BEGIN
               Getrec(VSAA^.DatFF[1],Ecr.PosCpteAuxE,Compte) ;
               Aux:=Trim(Compte.AuxiliaireT) ;
               END ;
            LMess('Création de l''écriture N° '+IntToStr(Ecr.NoPieceE)+' Ligne '+IntToStr(Ecr.NoOrdreE)) ;
            Inc(NoLigne) ;
            TECRGUI.Insert ;
            InitNew(TECRGUI) ;
            TECRGUI.FindField('EG_TYPE').AsString:=GU_TYPE ;
            TECRGUI.FindField('EG_GUIDE').AsString:=FormatFloat('000',MaxGuide) ;
            TECRGUI.FindField('EG_NUMLIGNE').AsInteger:=NoLigne ;
            TECRGUI.FindField('EG_GENERAL').AsString:=Gen ;
            TECRGUI.FindField('EG_AUXILIAIRE').AsString:=Aux ;
            If Aux='' Then Arret[2]:='-' ;
            TECRGUI.FindField('EG_REFINTERNE').AsString:=Ecr.ReferenceE ;
            TECRGUI.FindField('EG_LIBELLE').AsString:=Ecr.LibelleE ;
            Montant:=Arrondi(Ecr.MontantFrcsE,DC) ;
            If Ecr.SensE=1 Then
               BEGIN
               TECRGUI.FindField('EG_DEBITDEV').AsString:=FloatToStr(Arrondi(Ecr.MontantFrcsE,DC)) ;
               If Montant=0 Then Arret[5]:='X' ;
               END Else
               BEGIN
               TECRGUI.FindField('EG_CREDITDEV').AsString:=FloatToStr(Arrondi(Ecr.MontantFrcsE,DC)) ;
               If Montant=0 Then Arret[6]:='X' ;
               END ;
            TECRGUI.FindField('EG_ARRET').AsString:=arret ;
            TECRGUI.Post ;
            If (Ecr.AnalytiqueONE=1) And (ANAHal.Checked) Then
               BEGIN
               RecupAnalGeneAbo(Ecr,Fiche6,NoLigne,MaxGuide,GU_TYPE) ;
               END ;
            END ;
         END ;
    Until Not OkOK ;
    END ;
  END ;
1:
TCONTABO.Close ; TGUIDES.Close ; TECRGUI.Close ; TANAGUI.Close ;
GAccess.Ferme(4,2,FALSE) ; GAccess.Ferme(0,0,FALSE) ; GAccess.Ferme(1,0,FALSE) ; GAccess.Ferme(9,2,FALSE) ;
GAccess.Ferme(8,2,FALSE) ;GAccess.Ferme(6,0,FALSE) ;
FiniMove ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.LitEtatFiV8 ;
Var i : Integer ;
    St1,St2 : String13 ;
    OffSet : LongInt ;
    IMax,kk : Integer ;
begin
IMax:=19 ; FillChar(tbEtatFi,SizeOf(tbEtatFi),#0) ;
OuvreChoixCod ;
OffSet:=9250 ;
kk:=38 ;
For i:=1 to (kk div 2) do
    BEGIN
    Seek(VSAA^.FichierChCh,OffSet+i-1)    ; Read(VSAA^.FichierChCh,St1) ;
    Seek(VSAA^.FichierChCh,OffSet+20+i-1) ; Read(VSAA^.FichierChCh,St2) ;
    St1:=Trim(St1) ;
    If St1<>'' Then
      BEGIN
      New(tbEtatFi[i]) ; FillChar(tbEtatFi[i]^,SizeOf(tbEtatFi[i]^),#0) ;
      tbEtatFi[i]^.Code:=ASCII2ANSI(St1) ;
      If St2[1]='>' Then tbEtatFi[i]^.Signe:=1 Else tbEtatFi[i]^.Signe:=2 ;
      END ;
    END ;
FermeChoixCod ;
end ;

{=============================================================================}
procedure TFImpCpt.MajRub(TARB : TabRB ; i : Integer) ;
Var l,c : Integer ;
begin
If tbEtatFi[i]=Nil Then Exit ;
BeginTrans ;
InitMove(tbEtatFi[i]^.Nbl*tbEtatFi[i]^.NbC,'recup des rubriques') ;
TRUB.open ;
For l:=1 To tbEtatFi[i]^.Nbl Do
   For c:=1 To tbEtatFi[i]^.NbC Do
      BEGIN
      If MoveCur(FALSE) Then ;
      If TARB[l,c]<>NIL Then
        BEGIN
        TRUB.Insert ;
        InitNew(TRUB) ;
        // JLD + GP Pour XX_PREDEFINI
        TRUB.FindField('RB_PREDEFINI').AsString:='DOS' ; TRUB.FindField('RB_NODOSSIER').AsString:='000000' ;
        TRUB.FindField('RB_FAMILLES').AsString:=TARB[l,c]^.RB_FAMILLES ;
        TRUB.FindField('RB_RUBRIQUE').AsString:=TARB[l,c]^.RB_RUBRIQUE ;
        TRUB.FindField('RB_LIBELLE').AsString:=TARB[l,c]^.RB_LIBELLE ;
        TRUB.FindField('RB_SIGNERUB').AsString:=TARB[l,c]^.RB_SIGNERUB ;
        TRUB.FindField('RB_TYPERUB').AsString:=TARB[l,c]^.RB_TYPERUB ;
        TRUB.FindField('RB_COMPTE1').AsString:=TARB[l,c]^.RB_COMPTE1 ;
        TRUB.FindField('RB_EXCLUSION1').AsString:=TARB[l,c]^.RB_EXCLUSION1 ;
        TRUB.FindField('RB_COMPTE2').AsString:=TARB[l,c]^.RB_COMPTE2 ;
        TRUB.FindField('RB_EXCLUSION2').AsString:=TARB[l,c]^.RB_EXCLUSION2 ;
        TRUB.FindField('RB_AXE').AsString:=TARB[l,c]^.RB_AXE ;
        TRUB.Post ;
        END ;
      END ;
TRUB.Close ;
FiniMove ;
CommitTrans ;
end ;

procedure TFImpCpt.FreeTbEtatFi(i : Integer) ;
Var l,c : integer ;
begin
If tbEtatFi[i]<>NIL Then
  BEGIN
  For l:=1 To MaxLigneH Do For c:=1 To MaxColH Do
     BEGIN
     If tbEtatFi[i]^.Ligne[l,c]<>NIL Then Dispose(tbEtatFi[i]^.Ligne[l,c]) ;
     tbEtatFi[i]^.Ligne[l,c]:=NIL ;
     END ;
  Dispose(tbEtatFi[i]) ; tbEtatFi[i]:=NIL ;
  END ;
end ;

{=============================================================================}
procedure FreeTARB(Var TARB : TabRB) ;
Var l,c : integer ;
begin
For l:=1 To MaxLigneH Do For c:=1 To MaxColH Do
  BEGIN
  If TARB[l,c]<>NIL Then Dispose(TARB[l,c]) ; TARB[l,c]:=NIL ;
  END ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupFamilleRub(i : Integer) ;
begin
BeginTrans ;
TChoix.open ;
If (tbEtatFi[i]<>NIL) And (tbEtatFi[i]^.Code<>'') Then
  BEGIN
  TChoix.Insert ;
  InitNew(TChoix) ;
  TChoix.FindField('CC_TYPE').AsString:='RBF' ;
  TChoix.FindField('CC_CODE').AsString:=tbEtatFi[i]^.Famille ;
  TChoix.FindField('CC_LIBELLE').AsString:=tbEtatFi[i]^.Libelle ;
  TChoix.FindField('CC_ABREGE').AsString:=tbetatfi[i]^.Code ;
  TChoix.Post ;
  END ;
TChoix.Close ;
CommitTrans ;
end ;

{=============================================================================}
procedure TFImpCpt.recupEtatFi ;
Var i : Integer ;
    TARB : TabRB ;
begin
VideTable('RUBRIQUE','') ;
VideTable('CHOIXCOD','WHERE CC_TYPE="RBF"') ;
LitEtatFiV8 ;
For i:=1 To 19 Do
  If (tbEtatFi[i]<>NIL) And (tbEtatFi[i].Code<>'') Then
     BEGIN
     LMess('Création de l''état financier '+tbEtatFi[i]^.Code) ;
     ChercheBilan(i,tbEtatFi[i],tbEtatFi[i]^.Signe) ;
     TransfertEtatFi(tbEtatFi[i],TARB) ;
     MajRub(TARB,i) ; RecupFamilleRub(i) ;
     FreeTbEtatFi(i) ; FreeTARB(TARB) ;
     END ;
end ;

{=============================================================================}
procedure TFImpCpt.RecupPlanRef ;
Var i : Byte ;
    kk,PosFiche,Lg : LongInt ;
    Cle : String40 ;
    Fiche : EnregCpta ;
    OkOk : Boolean ;
    CC : Char ;
  Label 0 ;
begin
VideTable('PLANREF','') ;
For i:=0 To 9 Do
  BEGIN
  FichiersCPTA[7].Nom:=FRep.Text+'\UTIL\PLANREF'+IntToStr(i) ;
  Ouvre(7,2,FALSE,spCPTA) ;
  kk:=0 ; Lg:=FileLen(VSAA^.DatFF[7])-1 ; If Lg<=0 Then Goto 0 ;
  InitMove(Lg,'Recup du plan de référence N° '+IntToStr(i+1)) ;
  BeginTrans ; TRef.open ; Cle:=#1 ;
  Repeat
    Inc(kk) ;
    If kk=1 then SearchKey(VSAA^.Idxff[7],PosFiche,Cle) else NextKey(VSAA^.Idxff[7],PosFiche,Cle) ;
    OkOk:=VSAA^.Ok and (Cle[1]=#1) ;
    If MoveCur(FALSE) Then ;
    If OkOk then
       BEGIN
       GetRec(VSAA^.Datff[7],PosFiche,Fiche) ;
       If Not TRef.FindKey([i+1,BourreLess(Trim(Fiche.CompteRF),FbGene)]) Then
          BEGIN
          LMess('Plan N°'+IntToStr(i+1)+' : '+'Création du compte '+Fiche.CompteRF) ;
          TRef.Insert ;
          InitNew(TRef) ;
          TRef.findfield('PR_NUMPLAN').asInteger:=i+1 ;
          TRef.findfield('PR_COMPTE').asstring:=BourreLess(Trim(Fiche.CompteRF),FbGene) ;
          TRef.findfield('PR_LIBELLE').asstring:=ASCII2ANSI(Fiche.IntituleRF) ;
          TRef.findfield('PR_ABREGE').asstring:=ASCII2ANSI(Fiche.IntituleRF) ;
          TRef.findfield('PR_NATUREGENE').asstring:='DIV' ;
          TRef.findfield('PR_CENTRALISABLE').asstring:='-' ;
          TRef.findfield('PR_SOLDEPROGRESSIF').asstring:='-' ;
          TRef.findfield('PR_SAUTPAGE').asstring:='-' ;
          TRef.findfield('PR_TOTAUXMENSUELS').asstring:='-' ;
          TRef.findfield('PR_COLLECTIF').asstring:=Bool(Fiche.CollectifRF) ;
          CC:=Chr(i+65) ;
          LitBloc(Fiche.PosBlocRF,CC,spCPTA,TRef.findfield('PR_BLOCNOTE')) ;
          TRef.findfield('PR_SENS').asstring:=TrouveSens(Fiche.SensRF) ;
          TRef.findfield('PR_LETTRABLE').asstring:='-' ;
          TRef.findfield('PR_POINTABLE').asstring:='-' ;
          TRef.findfield('PR_VENTILABLE1').asstring:='-' ;
          TRef.findfield('PR_VENTILABLE2').asstring:='-' ;
          TRef.findfield('PR_VENTILABLE3').asstring:='-' ;
          TRef.findfield('PR_VENTILABLE4').asstring:='-' ;
          TRef.findfield('PR_VENTILABLE5').asstring:='-' ;
          TRef.findfield('PR_REPORTDETAIL').asstring:='-' ;
          TRef.Post ;
          END ;
       END ;
  Until (Not OkOk) ;
  TRef.Close ;
  CommitTrans ;
  FiniMove ;
  0:GAccess.Ferme(7,2,FALSE) ;
  END ;
end ;

{=============================================================================}
procedure TFImpCpt.MajSO_RECUPCPTA(Stt : String) ;
Var LaSoc : String3 ;
    St1,St2 : String ;
begin
LaSoc:=Trim(Copy(VSAA^.ParaSoc.CodeSociete,1,3)) ;
If LaSoc='' Then LaSoc:='001' ;
BeginTrans ;
TSoc.Open ;
If Tsoc.FindKey([LaSoc]) Then
   BEGIN
{$IFDEF SPEC302}
   St1:=TSoc.FindField('SO_RECUPCPTA').AsString ;
   St2:=Stt+Copy(St1,21,7) ;
   TSoc.Edit ;
   TSoc.FindField('SO_RECUPCPTA').AsString:=St2 ;
   TSoc.Post ;
{$ELSE}
   St1:=GetParamSoc('SO_RECUPCPTA') ; St2:=Stt+Copy(St1,21,7) ;
   SetparamSoc('SO_RECUPCPTA',St2) ;
{$ENDIF}
   END ;
TSoc.Close ;
CommitTrans ;
end ;

PROCEDURE CHARGERLESTABLES1 ;
Var i : integer ;
BEGIN
(*
InitMove(MaxCombos,' ') ;
BeginTrans ;
for i:=1 to MaxCombos do
  if ((V_PGI.DECombos[i].TT<>'') And (Copy(V_PGI.DECombos[i].TT,1,2)='TT') and (V_PGI.DeCombos[i].MemLoad=ltMem)) then
    BEGIN
    MoveCur(False) ;
    RemplirListe(V_PGI.DECombos[i].TT,'') ;
    END ;
FiniMove ;
CommitTrans ;
*)
END ;

{=============================================================================}
procedure TFImpCpt.RECUP(Sender: TObject);
Var OldPos6 : LongInt ;
    MaxGuide : SmallInt ;
    Stt : String ;
begin
If (VH^.etablisDefaut='') And (Not OkSoc.Checked) Then OkSoc.Checked:=TRUE ;
EnCours:=TRUE ;
SocPath:=TLabel(ListeSoc.Items.Objects[ListeSoc.ItemIndex]).Caption ;
If Pos(':',SocPath)>0 Then SocPath:=Trim(SocPath) Else SocPath:=Trim(FRep.text)+'\'+Trim(SocPath) ;
LitFichier(TRUE,FALSE,TRUE) ;
Stt:=LitSO_RECUPCPTA ;
Dc:=VSAA^.OkDecV ; LitLesDevisesV8 ; LitEdtLegal ;
TMessage.Visible:=TRUE ;
BStop.Enabled:=TRUE ; BAnnule.Enabled:=FALSE ; BSel.Enabled:=FALSE ;
OldPos6:=RepriseEcr ;
If OkSoc.Checked Then
   BEGIN
   Stt[1]:='1' ; RecupSoc ; RecupLibAuto ; Image1.Visible:=TRUE ;  AImage1.Visible:=FALSE ;
   END ;
If OkDev.Checked Then
   BEGIN
   Stt[2]:='1' ; RecupDevises ; AImage2.Visible:=FALSE ; Image2.Visible:=TRUE ;  
   END ;
If OkCHancell.Checked Then
   BEGIN
   Stt[3]:='1' ; RecupChancell ; Image3.Visible:=TRUE ;  AImage3.Visible:=FALSE ;
   END ;
If OkModR.Checked Then
   BEGIN
   Stt[4]:='1' ; RecupModR ; Image4.Visible:=TRUE ; AImage4.Visible:=FALSE ;
   END ;
If OkTva.Checked Then
   BEGIN
   Stt[5]:='1' ; RecupTva ; Image5.Visible:=TRUE ; AImage5.Visible:=FALSE ;
   END ;
If OkGen.Checked Then
   BEGIN
   Stt[6]:='1' ; RecupGeneraux ; Image6.Visible:=TRUE ;  AImage6.Visible:=FALSE ;
   END ;
If OkRef.Checked Then
   BEGIN
   Stt[7]:='1' ; RecupPlanRef ; Image7.Visible:=TRUE ; AImage7.Visible:=FALSE ;
   END ;
If OkRupt.Checked Then
   BEGIN
   Stt[8]:='1' ; RecupRuptures ; Image8.Visible:=TRUE ;  AImage8.Visible:=FALSE ;
   END ;
If OkTiers.Checked Then
   BEGIN
   Stt[9]:='1' ; RecupAuxiliaire ; Image9.Visible:=TRUE ;  AImage9.Visible:=FALSE ;
   END ;
If OkSect.Checked Then
   BEGIN
   Stt[10]:='1' ; RecupSect ; Image10.Visible:=TRUE ;  AImage10.Visible:=FALSE ;
   END ;
If OkJal.Checked Then
   BEGIN
   Stt[11]:='1' ; RecupJal ; RecupEdtLegal ; Image11.Visible:=TRUE ;  AImage11.Visible:=FALSE ;
   END ;
If OkVentilAnal.Checked Then
   BEGIN
   Stt[12]:='1' ; RecupVentilAnal ; RecupVentilType ; Image12.Visible:=TRUE ;  AImage12.Visible:=FALSE ;
   END ;
If OkEcrGen.Checked Then
   BEGIN
   Stt[13]:='1' ; LitTvaV8 ;
   RecupEcrGen(OldPos6) ;
   SAVDATEMINDATEMAXPAQUET(Tous) ;
   Image13.Visible:=TRUE ;  AImage13.Visible:=FALSE ;
   END ;
If OkODA.Checked Then
   BEGIN
   Stt[14]:='1' ; RecupODA ; Image14.Visible:=TRUE ;  AImage14.Visible:=FALSE ;
   END ;
If OkEtatFi.Checked Then
   BEGIN
   Stt[15]:='1' ; RecupEtatFi ; Image15.Visible:=TRUE ;  AImage15.Visible:=FALSE ;
   END ;
If OkGuide.Checked Then
   BEGIN
   Stt[16]:='1' ; RecupGuide(MaxGuide) ;
   RecupEcrAboOuTyp(MaxGuide,GDECLADF.EcrAbo) ;
   RecupEcrAboOuTyp(MaxGuide,GDECLADF.EcrTyp) ;
   Image16.Visible:=TRUE ;  AImage16.Visible:=FALSE ;
   END ;
If OkCalcSolde.Checked Then
   BEGIN
   LMEss('MAJ Solde des comptes') ;
   MajTotTousComptes(TRUE,'') ;
   END ;
TMessage.Visible:=FALSE ;
EnCours:=FALSE ; Stopper:=FALSE ;
ChargerLesTables1 ;
MajSO_RECUPCPTA(Stt) ;
BStop.Enabled:=FALSE ; BAnnule.Enabled:=TRUE ; BSel.Enabled:=TRUE ;
LCodeTva.Clear ; LCptTva.Clear ; LesRefJ.Clear ;
end;

{=============================================================================}
procedure TFImpCpt.BAnnuleClick(Sender: TObject);
begin
//if Not Encours then Close else
//   BEGIN
//   Stopper:=TRUE ;
//   END ;
Close;
end;

{=============================================================================}
procedure TFImpCpt.BDirClick(Sender: TObject);
var st : AnsiString ;
    i : integer ;
    Options: TSelectDirOpts ;
begin
If directoryExists(FRep.Text) then St:=FRep.text else st:='c:\' ;
i:=1 ; Options:=[] ;
if selectdirectory(st,Options,i) then
    BEGIN
    FRep.text:=St ;
    ALIMLISTESOC(nil);
    END ;
end;

{=============================================================================}
procedure TFImpCpt.ALIMLISTESOC(Sender: TObject);
Var Fichier : File of NomSoc ;
    Fiche : NomSoc ;
    T : TLabel ;
    OkOk : Boolean ;
begin
{ free de listesoc à faite }
If directoryExists(FRep.Text) then
   begin
   OkOk:=TRUE ; { FileExists(FRep.Text+'\SCM.EXE') }
   If OkOk Then
      BEGIN
      AssignFile(fichier,FRep.Text+'\UTIL\SOCIETE') ; reset(fichier) ;
      While (Not EOF(fichier)) do
         BEGIN
         Read(fichier,Fiche) ;
         if Fiche.Nom<>'' then
            BEGIN
            T:=TLabel.Create(Nil) ; T.Caption:=Fiche.Path ;
            ListeSoc.Items.addobject(Fiche.Nom,T) ;
            END ;
         END ;
      System.Close(Fichier) ;
      ListeSoc.ItemIndex:=0 ;
      END Else ShowMessage(HMSGReprise.Mess[2]) ;
   end else ShowMessage(HMSGReprise.Mess[2]) ;
end;

{=============================================================================}
function TFImpCpt.LitSO_RECUPCPTA : String ;
Var LaSoc : String3 ;
begin
TSoc.Open ; Result:='00000000000000000000' ;
LaSoc:=Trim(Copy(VSAA^.ParaSoc.CodeSociete,1,3)) ;
If LaSoc='' Then LaSoc:='001' ;
TSoc.Open ;
If Tsoc.FindKey([LaSoc]) Then
   BEGIN
{$IFDEF SPEC302}
   Result:=Trim(Copy(TSoc.FindField('SO_RECUPCPTA').AsString,1,20)) ;
{$ELSE}
   Result:=Trim(Copy(GetParamSoc('SO_RECUPCPTA'),1,20)) ;
{$ENDIF}
   If Result='' Then Result:='00000000000000000000' ;
   END ;
TSoc.Close ;
end ;

{=============================================================================}
procedure TFImpCpt.InitImage ;
Var LaSoc : String3 ;
    St : String ;
begin
TSoc.Open ;
LaSoc:=V_PGI.CodeSociete ;
If LaSoc='' Then LaSoc:='001' ;
St:='00000000000000000000' ;
TSoc.Open ;
If Tsoc.FindKey([LaSoc]) Then
   BEGIN
{$IFDEF SPEC302}
   St:=Trim(Copy(TSoc.FindField('SO_RECUPCPTA').AsString,1,20)) ;
{$ELSE}
   St:=Trim(Copy(GetParamSoc('SO_RECUPCPTA'),1,20)) ;
{$ENDIF}
   If St='' Then St:='00000000000000000000' ;
   END ;
TSoc.Close ;
Image1.Visible:=St[1]='1' ;  AImage1.Visible:=St[1]='0' ;
Image2.Visible:=St[2]='1' ;  AImage2.Visible:=St[2]='0' ;
Image3.Visible:=St[3]='1' ;  AImage3.Visible:=St[3]='0' ;
Image4.Visible:=St[4]='1' ;  AImage4.Visible:=St[4]='0' ;
Image5.Visible:=St[5]='1' ;  AImage5.Visible:=St[5]='0' ;
Image6.Visible:=St[6]='1' ;  AImage6.Visible:=St[6]='0' ;
Image7.Visible:=St[7]='1' ;  AImage7.Visible:=St[7]='0' ;
Image8.Visible:=St[8]='1' ;  AImage8.Visible:=St[8]='0' ;
Image9.Visible:=St[9]='1' ;  AImage9.Visible:=St[9]='0' ;
Image10.Visible:=St[10]='1' ; AImage10.Visible:=St[10]='0' ;
Image11.Visible:=St[11]='1' ; AImage11.Visible:=St[11]='0' ;
Image12.Visible:=St[12]='1' ; AImage12.Visible:=St[12]='0' ;
Image13.Visible:=St[13]='1' ; AImage13.Visible:=St[13]='0' ;
Image14.Visible:=St[14]='1' ; AImage14.Visible:=St[14]='0' ;
Image15.Visible:=St[15]='1' ; AImage15.Visible:=St[15]='0' ;
Image16.Visible:=St[16]='1' ; AImage16.Visible:=St[16]='0' ;
OkSoc.Checked        :=AImage1.Visible ;
OkDev.Checked        :=AImage2.Visible;
OkChancell.Checked   :=AImage3.Visible;
OkModR.Checked       :=AImage4.Visible;
OkTva.Checked        :=AImage5.Visible;
OkGen.Checked        :=AImage6.Visible;
OkRef.Checked        :=AImage7.Visible;
OkRupt.Checked       :=AImage8.Visible;
OkTiers.Checked      :=AImage9.Visible;
OkSect.Checked       :=AImage10.Visible;
OkJal.Checked        :=AImage11.Visible;
OkVentilanal.Checked :=AImage12.Visible;
OkEcrGen.Checked     :=AImage13.Visible;
OKOdA.Checked        :=AImage14.Visible;
OKEtatFi.Checked     :=AImage15.Visible;
OKGuide.Checked      :=AImage16.Visible;
end ;
{=============================================================================}
procedure TFImpCpt.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Stopper:=FALSE ; EnCours:=FALSE ;
LCodeTva:=TStringList.Create ; LCodeTVA.Sorted:=TRUE ; LCodeTVA.Duplicates:=DupIgnore ;
LCptTva:=TStringList.Create ; LCptTVA.Sorted:=TRUE ; LCptTVA.Duplicates:=DupIgnore ;
LesRefJ:=TStringList.Create ; LesRefJ.Sorted:=TRUE ; LesRefJ.Duplicates:=DupIgnore ;
InitImage ;
BStop.Enabled:=FALSE ;

end;

Procedure RecupSCM ;
Var FImpCpt : TFImpCpt ;
begin
if Not BlocageMonoPoste(False) then Exit ;
FImpCpt:=TFImpCpt.Create(Application) ;
  Try
   FImpCpt.ShowModal ;
  Finally
   FImpCpt.Free ;
   DeblocageMonoPoste(False) ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;


Function TFImpCpt.RepriseEcr : LongInt ;
Var Rep : Integer ;
    OldPos : LongInt ;
begin
OldPos:=LitLaPos6 ;
If (FVider.Checked) And (OkEcrGen.Checked) And (OldPos<>0)  Then
   BEGIN
   Rep:=HMsgReprise.Execute(0,'','') ;
   Case rep of
     mrYes : ;
     mrNo : OldPos:=1 ;
     mrCancel : OldPos:=1 ;
     end ;
   END Else OldPos:=1 ;
Result:=OldPos ;
end;

procedure TFImpCpt.OkEcrGenClick(Sender: TObject);
begin
If TCheckBox(Sender).Checked Then OkCAlcSolde.Checked:=TRUE ;
end;

procedure TFImpCpt.BSelClick(Sender: TObject);
Var OkB : Boolean ;
begin
OkB:=Not OkSoc.Checked ;
OkSoc.Checked        :=OkB;
OkDev.Checked        :=OkB;
OkChancell.Checked   :=OkB;
OkModR.Checked       :=OkB;
OkTva.Checked        :=OkB;
OkGen.Checked        :=OkB;
OkRef.Checked        :=OkB;
OkRupt.Checked       :=OkB;
OkTiers.Checked      :=OkB;
OkSect.Checked       :=OkB;
OkJal.Checked        :=OkB;
OkVentilanal.Checked :=OkB;
OkEcrGen.Checked     :=OkB;
OKOdA.Checked        :=OkB;
OKEtatFi.Checked     :=OkB;
OKGuide.Checked      :=OkB;
end;

procedure TFImpCpt.BStopClick(Sender: TObject);
begin
Stopper:=TRUE ;
BStop.Enabled:=FALSE ; BAnnule.Enabled:=TRUE ; BSel.Enabled:=TRUE ;
end;

procedure TFImpCpt.RV8DblClick(Sender: TObject);
Var i : SmallInt ;
begin
SocPath:=TLabel(ListeSoc.Items.Objects[ListeSoc.ItemIndex]).Caption ;
SocPath:=Trim(FRep.text)+'\'+Trim(SocPath) ;
LitFichier(TRUE,FALSE,TRUE) ;
I:=StrToInt(RV8.Text) ;
TrouveIndexRupt(2,i) ;
end;

procedure TFImpCpt.AnaHALClick(Sender: TObject);
begin
OkSect.enabled:=AnaHal.Checked ;
OkVentilAnal.enabled:=AnaHal.Checked ;
OkODA.enabled:=AnaHal.Checked ;
If Not AnaHal.Checked Then
   BEGIN
   OkSect.Checked:=FALSE ;
   OkVentilAnal.Checked:=FALSE ;
   OkODA.Checked:=FALSE ;
   END ;
end;

end.
{$A+,H+}

